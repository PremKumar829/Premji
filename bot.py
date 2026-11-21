"""
Telegram Group Report Tracker -> Excel-Only (.xlsx) + Auto Assistant Discovery + Penalty Countdown
Deployable directly via GitHub + Render Free Tier (no Google Sheets / external DB needed).

Core Behavior:
  1. Excel-Only Delivery: Sends formatted .xlsx workbook (team_report_YYYY-MM-DD.xlsx)
     to your Excel/Backup Group on every recorded slot submission AND at 8:35 PM IST to Admin DM.
  2. Auto-Recognize Assistants (Case-Insensitive): Deduplicates username variants like
     @PREMGUPTA2M and @Premgupta2m into one canonical assistant; auto-excludes Group Admins.
  3. Slot Timing, Countdown & Penalties:
     - Regular Slots (10 AM, 1 PM, 3 PM, 5 PM): On Time (15m) = ₹0 | Late (after 15m) = ₹5 | Missed = ₹15
     - Final Report (7:15 PM – 8:00 PM): On Time = ₹0 | 7:45 PM Final Warning + 15m Countdown |
       8:00 PM Last Warning | Late after 8:00 PM = ₹10 | Missed = ₹15
  4. Safe HTML formatting + user_id mention fallback + full edited-caption & album support.

Requires: python-telegram-bot[job-queue] >= 21.0, openpyxl >= 3.1.5, python-dotenv, tzdata
"""

import asyncio
import html
import json
import logging
import os
import re
import sqlite3
import tempfile
import threading
from dataclasses import dataclass
from datetime import datetime, time, timedelta
from http.server import BaseHTTPRequestHandler, HTTPServer
from io import BytesIO
from typing import Optional
from zoneinfo import ZoneInfo

from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from telegram import (
    BotCommand,
    BotCommandScopeAllGroupChats,
    BotCommandScopeChat,
    BotCommandScopeDefault,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Update,
)
try:
    from telegram import ReactionTypeEmoji
except ImportError:
    ReactionTypeEmoji = None  # type: ignore[assignment]
from telegram.ext import (
    Application,
    ApplicationBuilder,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

# ---------------------------------------------------------------------------
# Configuration & Logging
# ---------------------------------------------------------------------------
load_dotenv()

logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("excel_report_bot")


def _parse_optional_int(raw: str, env_name: str) -> Optional[int]:
    cleaned = raw.strip().strip('"').strip("'").strip()
    if not cleaned:
        return None
    try:
        return int(cleaned)
    except ValueError:
        logger.warning("Invalid integer for %s=%r; ignoring.", env_name, raw)
        return None


BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip().strip('"').strip("'")

DEFAULT_ADMIN_CHAT_ID: Optional[int] = _parse_optional_int(
    os.getenv("ADMIN_CHAT_ID", ""), "ADMIN_CHAT_ID"
)
DEFAULT_WORK_CHAT_ID: Optional[int] = _parse_optional_int(
    os.getenv("ALLOWED_CHAT_ID", ""), "ALLOWED_CHAT_ID"
)
DEFAULT_EXCEL_GROUP_ID: Optional[int] = _parse_optional_int(
    os.getenv("EXCEL_GROUP_ID") or os.getenv("DB_BACKUP_CHAT_ID") or "",
    "EXCEL_GROUP_ID",
)

ASSISTANT_TAGS_RAW = os.getenv("ASSISTANT_TAGS", "").strip()
DEFAULT_ASSISTANT_TAGS: list[str] = [
    tag.strip() for tag in ASSISTANT_TAGS_RAW.split(",") if tag.strip()
]

DEFAULT_ENABLE_REMINDERS = os.getenv("ENABLE_REMINDERS", "true").strip().lower() == "true"
PREVENT_DUPLICATE_SLOTS = os.getenv("PREVENT_DUPLICATE_SLOTS", "true").strip().lower() == "true"

DB_PATH = os.getenv("DB_PATH", "reports.db").strip()
IST = ZoneInfo("Asia/Kolkata")
MORNING_KICKOFF_TIME_IST = time(hour=9, minute=45, second=0, tzinfo=IST)   # 9:45 AM IST
DAILY_PENALTY_TIME_IST = time(hour=20, minute=30, second=0, tzinfo=IST)    # 8:30 PM IST
DAILY_SEND_TIME_IST = time(hour=20, minute=35, second=0, tzinfo=IST)       # 8:35 PM IST
DAILY_LEADERBOARD_TIME_IST = time(hour=20, minute=45, second=0, tzinfo=IST)  # 8:45 PM IST
NIGHTLY_BACKUP_TIME_IST = time(hour=21, minute=15, second=0, tzinfo=IST)   # 9:15 PM IST
COUNTDOWN_DURATION_SECS = 15 * 60
COUNTDOWN_REFRESH_SECS = 15
PENALTY_LATE_REGULAR_RS = 5
PENALTY_LATE_FINAL_RS = 10
PENALTY_MISSED_RS = 15

_REPORT_LOCK = asyncio.Lock()
_BOT_USERNAME_CACHE: Optional[str] = None
_SEEN_MEDIA_GROUPS: dict[str, datetime] = {}
_MISMATCH_WARN_CACHE: dict[int, datetime] = {}
_ADMIN_SYNC_CACHE: dict[int, datetime] = {}
_ACTIVE_COUNTDOWNS: dict[tuple[int, str], dict[str, object]] = {}
_COUNTDOWN_TASKS: dict[tuple[int, str], asyncio.Task] = {}
_CELEBRATION_EMOJIS: tuple[str, ...] = ("🔥", "🎉", "💯", "🏆", "⚡", "🤩", "👏")


# ---------------------------------------------------------------------------
# Reporting Slot Definitions (IST)
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class ReportingSlot:
    label: str
    start: time
    end: time  # Inclusive end of On-Time window
    reminder_time: time  # Slot-open reminder + tag warning
    final_warning_time: time  # Warning before On-Time closes (7:45 PM with countdown for Final)


REPORTING_SLOTS: tuple[ReportingSlot, ...] = (
    ReportingSlot("10:00 AM Slot", time(10, 0, 0), time(10, 15, 59), time(10, 0, 0), time(10, 10, 0)),
    ReportingSlot("1:00 PM Slot", time(13, 0, 0), time(13, 15, 59), time(13, 0, 0), time(13, 10, 0)),
    ReportingSlot("3:00 PM Slot", time(15, 0, 0), time(15, 15, 59), time(15, 0, 0), time(15, 10, 0)),
    ReportingSlot("5:00 PM Slot", time(17, 0, 0), time(17, 15, 59), time(17, 0, 0), time(17, 10, 0)),
    ReportingSlot(
        "Final Report (Before 8:00 PM)",
        time(19, 15, 0),
        time(20, 0, 59),
        time(19, 15, 0),
        time(19, 45, 0),
    ),
)
FINAL_SLOT_LABEL = REPORTING_SLOTS[4].label
SLOT_ORDER: dict[str, int] = {slot.label: idx for idx, slot in enumerate(REPORTING_SLOTS)}
SHIFT_SLOTS: dict[str, set[str]] = {
    "full": {s.label for s in REPORTING_SLOTS},
    "morning": {REPORTING_SLOTS[0].label, REPORTING_SLOTS[1].label, REPORTING_SLOTS[2].label},
    "evening": {REPORTING_SLOTS[2].label, REPORTING_SLOTS[3].label, REPORTING_SLOTS[4].label},
}


def detect_active_slot(ist_dt: datetime) -> Optional[str]:
    current_t = ist_dt.time().replace(tzinfo=None)
    for slot in REPORTING_SLOTS:
        if slot.start <= current_t <= slot.end:
            return slot.label
    return None


def parse_slot_alias(raw: str) -> Optional[str]:
    token = raw.strip().lower().replace(" ", "").replace(":00", "")
    if not token:
        return None
    if token in ("10", "10am", "1000", "slot1", "1st", "morning"):
        return REPORTING_SLOTS[0].label
    if token in ("1", "1pm", "13", "1300", "slot2", "2nd", "afternoon"):
        return REPORTING_SLOTS[1].label
    if token in ("3", "3pm", "15", "1500", "slot3", "3rd"):
        return REPORTING_SLOTS[2].label
    if token in ("5", "5pm", "17", "1700", "slot4", "4th", "evening"):
        return REPORTING_SLOTS[3].label
    if token in ("8", "8pm", "20", "1930", "1915", "715", "final", "slot5", "5th", "night"):
        return REPORTING_SLOTS[4].label
    for slot in REPORTING_SLOTS:
        if token in slot.label.lower().replace(" ", ""):
            return slot.label
    return None


def infer_slot_from_time(ist_dt: datetime) -> str:
    active = detect_active_slot(ist_dt)
    if active:
        return active
    current_mins = ist_dt.hour * 60 + ist_dt.minute
    best_slot = REPORTING_SLOTS[0]
    best_diff = 10_000
    for slot in REPORTING_SLOTS:
        slot_mins = slot.start.hour * 60 + slot.start.minute
        diff = abs(current_mins - slot_mins)
        if diff < best_diff:
            best_diff = diff
            best_slot = slot
    return best_slot.label


def _get_slot_by_label(slot_label: str) -> Optional[ReportingSlot]:
    for slot in REPORTING_SLOTS:
        if slot.label == slot_label:
            return slot
    return None


def _parse_ist_time_str(sub_time: str) -> Optional[time]:
    cleaned = re.sub(r"\s*IST\s*$", "", (sub_time or "").strip(), flags=re.IGNORECASE).strip()
    for fmt in ("%I:%M:%S %p", "%I:%M %p", "%I:%M:%S%p", "%I:%M%p", "%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(cleaned, fmt).time()
        except ValueError:
            continue
    return None


def _time_to_secs(t: time) -> int:
    return t.hour * 3600 + t.minute * 60 + t.second


def _is_final_slot(slot_label: str) -> bool:
    return slot_label == FINAL_SLOT_LABEL


def get_penalty_config() -> tuple[int, int, int]:
    """Return (regular_late_rs, final_late_rs, missed_rs) from bot_settings with defaults 5, 10, 15."""
    try:
        with sqlite3.connect(DB_PATH) as conn:
            rows = dict(
                conn.execute(
                    "SELECT key, value FROM bot_settings WHERE key IN ('penalty_late_regular', 'penalty_late_final', 'penalty_missed')"
                ).fetchall()
            )
        reg = int(rows.get("penalty_late_regular", PENALTY_LATE_REGULAR_RS))
        fin = int(rows.get("penalty_late_final", PENALTY_LATE_FINAL_RS))
        mis = int(rows.get("penalty_missed", PENALTY_MISSED_RS))
        return max(0, reg), max(0, fin), max(0, mis)
    except Exception:
        return PENALTY_LATE_REGULAR_RS, PENALTY_LATE_FINAL_RS, PENALTY_MISSED_RS


def _late_penalty_for_slot(slot_label: str) -> int:
    reg, fin, _ = get_penalty_config()
    return fin if _is_final_slot(slot_label) else reg


def get_missed_penalty_rs() -> int:
    return get_penalty_config()[2]


def is_date_holiday_or_off(target_date: str) -> tuple[bool, str]:
    """Return (True, reason) if holiday_mode is ON or target_date falls on configured off-days (e.g. sun)."""
    try:
        with sqlite3.connect(DB_PATH) as conn:
            rows = dict(
                conn.execute(
                    "SELECT key, value FROM bot_settings WHERE key IN ('holiday_mode', 'off_days')"
                ).fetchall()
            )
    except Exception:
        return False, ""
    if str(rows.get("holiday_mode", "false")).lower() == "true":
        return True, "Holiday Mode ON"
    off_raw = str(rows.get("off_days", "")).strip().lower()
    if off_raw and off_raw != "none":
        off_tokens = {t.strip()[:3] for t in off_raw.split(",") if t.strip()}
        try:
            wk = datetime.strptime(target_date, "%Y-%m-%d").strftime("%a").lower()[:3]
            if wk in off_tokens:
                return True, f"Weekly Off ({wk.upper()})"
        except ValueError:
            pass
    return False, ""


def get_closed_slots_for_date(
    target_date: str, now_ist: Optional[datetime] = None
) -> list[ReportingSlot]:
    """Return slots whose On-Time deadline has already closed on target_date (empty on Holidays/Off-days)."""
    is_off, _reason = is_date_holiday_or_off(target_date)
    if is_off:
        return []
    ref = now_ist or datetime.now(IST)
    today_str = ref.strftime("%Y-%m-%d")
    if target_date < today_str:
        return list(REPORTING_SLOTS)
    if target_date > today_str:
        return []
    cur_sec = _time_to_secs(ref.time().replace(tzinfo=None))
    return [s for s in REPORTING_SLOTS if cur_sec > _time_to_secs(s.end)]


def evaluate_slot_status_and_penalty(ist_dt: datetime, slot_label: str) -> tuple[str, int]:
    """
    Evaluate Status and Penalty (₹0 / ₹5 / ₹10) for a submitted report:
      - Within [slot.start, slot.end] (or <= 15m early): On Time, ₹0
      - Late on regular slots (10am, 1pm, 3pm, 5pm after :15:59): Late (₹5 Penalty), ₹5
      - Late on Final Report (after 8:00 PM / 20:00:59): Late (₹10 Penalty), ₹10
    (Completely unsubmitted closed slots incur ₹15 Missed Penalty in summary/Excel.)
    """
    slot = _get_slot_by_label(slot_label)
    if not slot:
        return "On Time", 0
    cur_sec = _time_to_secs(ist_dt.time().replace(tzinfo=None))
    start_sec = _time_to_secs(slot.start)
    end_sec = _time_to_secs(slot.end)

    if start_sec <= cur_sec <= end_sec:
        return "On Time", 0
    if 0 <= (start_sec - cur_sec) <= 900:
        return "On Time", 0
    pen = _late_penalty_for_slot(slot_label)
    return f"Late (₹{pen} Penalty)", pen


def evaluate_slot_status(ist_dt: datetime, slot_label: str) -> str:
    return evaluate_slot_status_and_penalty(ist_dt, slot_label)[0]


def evaluate_slot_status_from_strings_with_penalty(
    submission_time: str, slot_label: str
) -> tuple[str, int]:
    parsed_t = _parse_ist_time_str(submission_time)
    slot = _get_slot_by_label(slot_label)
    if not parsed_t or not slot:
        return "On Time", 0
    cur_sec = _time_to_secs(parsed_t)
    start_sec = _time_to_secs(slot.start)
    end_sec = _time_to_secs(slot.end)

    if start_sec <= cur_sec <= end_sec:
        return "On Time", 0
    if 0 <= (start_sec - cur_sec) <= 900:
        return "On Time", 0
    pen = _late_penalty_for_slot(slot_label)
    return f"Late (₹{pen} Penalty)", pen


def evaluate_slot_status_from_strings(submission_time: str, slot_label: str) -> str:
    return evaluate_slot_status_from_strings_with_penalty(submission_time, slot_label)[0]


def resolve_submission_slot_and_penalty(ist_dt: datetime) -> tuple[Optional[str], str, int]:
    """
    Map ANY report submission time (>= 10:00 AM IST) to the active or most-recent slot
    with exact ₹0 / ₹5 (regular late) / ₹10 (final after 8 PM) penalty status.
    """
    cur_sec = _time_to_secs(ist_dt.time().replace(tzinfo=None))

    for slot in REPORTING_SLOTS:
        if _time_to_secs(slot.start) <= cur_sec <= _time_to_secs(slot.end):
            return slot.label, "On Time", 0

    latest_started: Optional[ReportingSlot] = None
    for slot in REPORTING_SLOTS:
        if cur_sec >= _time_to_secs(slot.start):
            latest_started = slot

    if latest_started is None:
        return None, "Outside Slot", 0

    status, penalty_rs = evaluate_slot_status_and_penalty(ist_dt, latest_started.label)
    return latest_started.label, status, penalty_rs


def recalculate_report_statuses(conn: sqlite3.Connection) -> int:
    """Auto-sync status and late penalty on existing/restored rows (preserving admin-waived rows)."""
    fixed = 0
    cols = {row[1] for row in conn.execute("PRAGMA table_info(reports)").fetchall()}
    has_pen = "penalty_rs" in cols
    has_waived = "waived_on_time" in cols
    sel_pen = "COALESCE(penalty_rs, 0)" if has_pen else "0"
    sel_waived = "COALESCE(waived_on_time, 0)" if has_waived else "0"
    rows = conn.execute(
        f"SELECT id, submission_time, reporting_slot, status, {sel_pen}, {sel_waived} FROM reports"
    ).fetchall()
    for row_id, sub_time, slot_label, status, cur_pen, is_waived in rows:
        if int(is_waived or 0) == 1:
            continue
        expected_status, expected_pen = evaluate_slot_status_from_strings_with_penalty(
            str(sub_time or ""), str(slot_label or "")
        )
        if expected_status != status or int(cur_pen or 0) != expected_pen:
            if has_pen:
                conn.execute(
                    "UPDATE reports SET status = ?, penalty_rs = ? WHERE id = ?",
                    (expected_status, expected_pen, row_id),
                )
            else:
                conn.execute(
                    "UPDATE reports SET status = ? WHERE id = ?",
                    (expected_status, row_id),
                )
            fixed += 1
    return fixed


def _safe_int_joining(val: str) -> Optional[int]:
    cleaned = str(val or "").replace(",", "").strip()
    if not cleaned:
        return None
    try:
        return int(cleaned)
    except ValueError:
        return None


def compute_daily_joining_from_rows(
    day_rows: list[tuple[str, str, str, str, str, str, int, str]]
) -> int:
    """
    Joining within a single day is cumulative (e.g. 10 AM = 30, 1 PM = 45 -> today's valid joining is 45, NOT 75).
    Returns the latest chronological slot's valid numeric joining value for that day.
    """
    best_rank = (-1, -1)
    best_val = 0
    for idx, r in enumerate(day_rows):
        parsed = _safe_int_joining(r[1])
        if parsed is None:
            continue
        slot_rank = SLOT_ORDER.get(str(r[3]), 99)
        rank = (slot_rank, idx)
        if rank >= best_rank:
            best_rank = rank
            best_val = parsed
    return best_val


def compute_period_joining_from_rows(
    member_rows: list[tuple[str, str, str, str, str, str, int, str]]
) -> int:
    """Sum valid cumulative daily joining across distinct report_date values."""
    by_date: dict[str, list[tuple[str, str, str, str, str, str, int, str]]] = {}
    for r in member_rows:
        by_date.setdefault(str(r[4]), []).append(r)
    return sum(compute_daily_joining_from_rows(d_rows) for d_rows in by_date.values())


_USERNAME_PATTERN = re.compile(
    r"(?:telegram[ \t]*)?(?:username|user[ \t]*name|user|tg[ \t]*id|tg[ \t]*username|id)[ \t]*(?:[:=\-–—]+[ \t]*@?([A-Za-z0-9_]{3,})|[ \t]+@([A-Za-z0-9_]{3,}))",
    re.IGNORECASE,
)
_JOINING_PATTERN = re.compile(
    r"(?:(?:total[ \t]*)?(?:joinings?|jonings?|joins?)[ \t]*[:=\-–—]*[ \t]*(\d[\d,]*))|(?:(?<![A-Za-z0-9_@])(\d[\d,]*)[ \t]+(?:total[ \t]+)?(?:joinings?|jonings?|joins?)\b)",
    re.IGNORECASE,
)


def parse_report_fields(raw_text: str, fallback_username: str) -> tuple[str, str]:
    text = raw_text or ""
    username = fallback_username
    user_match = _USERNAME_PATTERN.search(text)
    if user_match:
        extracted = (user_match.group(1) or user_match.group(2) or "").strip()
        if extracted:
            username = extracted if extracted.startswith("@") else f"@{extracted}"

    joining = ""
    join_match = _JOINING_PATTERN.search(text)
    if join_match:
        raw_num = (join_match.group(1) or join_match.group(2) or "").replace(",", "").strip()
        if raw_num:
            joining = raw_num

    return username, joining


def is_text_report_message(raw_text: str) -> bool:
    if not raw_text:
        return False
    return bool(_JOINING_PATTERN.search(raw_text))


def is_report_caption(raw_text: str) -> bool:
    """Only treat messages/photos as reports when caption explicitly contains Joining or Username/TG ID fields."""
    if not raw_text:
        return False
    return bool(_JOINING_PATTERN.search(raw_text) or _USERNAME_PATTERN.search(raw_text))


def _normalize_tag(raw: str) -> str:
    clean = raw.strip().lstrip(",")
    if not clean:
        return ""
    return clean if clean.startswith("@") or " " in clean else f"@{clean}"


def _format_mention_html(username: str, user_id: Optional[int]) -> str:
    if username.startswith("@"):
        return html.escape(username)
    if user_id is not None:
        return f'<a href="tg://user?id={user_id}">{html.escape(username)}</a>'
    return html.escape(username)


# ---------------------------------------------------------------------------
# SQLite Persistence (Reports + Team Roster + Assistant Tags + Bot Settings)
# ---------------------------------------------------------------------------
def init_db() -> None:
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS reports (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                username TEXT NOT NULL,
                joining TEXT DEFAULT '',
                submission_time TEXT NOT NULL,
                reporting_slot TEXT NOT NULL,
                report_date TEXT NOT NULL,
                status TEXT NOT NULL,
                penalty_rs INTEGER NOT NULL DEFAULT 0,
                caption TEXT DEFAULT '',
                created_at_utc TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS team_members (
                username TEXT PRIMARY KEY,
                user_id INTEGER,
                added_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS assistant_tags (
                username TEXT PRIMARY KEY,
                user_id INTEGER,
                is_active INTEGER NOT NULL DEFAULT 1,
                added_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS bot_settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS summary_excluded (
                username TEXT PRIMARY KEY,
                added_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS known_group_admins (
                admin_key TEXT PRIMARY KEY,
                user_id INTEGER,
                username TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS leaves (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username_key TEXT NOT NULL,
                username TEXT NOT NULL,
                leave_date TEXT NOT NULL,
                slot_label TEXT NOT NULL DEFAULT '',
                reason TEXT DEFAULT '',
                created_at TEXT NOT NULL,
                UNIQUE(username_key, leave_date, slot_label)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS penalty_ledger (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username_key TEXT NOT NULL,
                username TEXT NOT NULL,
                entry_type TEXT NOT NULL,
                amount_rs INTEGER NOT NULL,
                note TEXT DEFAULT '',
                entry_date TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS username_aliases (
                alias_key TEXT PRIMARY KEY,
                alias_display TEXT NOT NULL,
                target_key TEXT NOT NULL,
                target_display TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS assistant_shifts (
                username_key TEXT PRIMARY KEY,
                username TEXT NOT NULL,
                shift_name TEXT NOT NULL DEFAULT 'full',
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS report_photos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_unique_id TEXT NOT NULL,
                username_key TEXT NOT NULL,
                username TEXT NOT NULL,
                report_date TEXT NOT NULL,
                reporting_slot TEXT NOT NULL,
                message_id INTEGER,
                created_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS tracked_chats (
                chat_id INTEGER PRIMARY KEY,
                title TEXT DEFAULT '',
                role TEXT NOT NULL DEFAULT 'work',
                added_at TEXT NOT NULL
            )
            """
        )

        # Lightweight schema migrations
        tag_cols = {
            row[1] for row in conn.execute("PRAGMA table_info(assistant_tags)").fetchall()
        }
        if "is_active" not in tag_cols:
            conn.execute("ALTER TABLE assistant_tags ADD COLUMN is_active INTEGER NOT NULL DEFAULT 1")
        if "user_id" not in tag_cols:
            conn.execute("ALTER TABLE assistant_tags ADD COLUMN user_id INTEGER")

        team_cols = {
            row[1] for row in conn.execute("PRAGMA table_info(team_members)").fetchall()
        }
        if "user_id" not in team_cols:
            conn.execute("ALTER TABLE team_members ADD COLUMN user_id INTEGER")

        report_cols = {
            row[1] for row in conn.execute("PRAGMA table_info(reports)").fetchall()
        }
        if "user_id" not in report_cols:
            conn.execute("ALTER TABLE reports ADD COLUMN user_id INTEGER")
        if "joining" not in report_cols:
            conn.execute("ALTER TABLE reports ADD COLUMN joining TEXT DEFAULT ''")
        if "penalty_rs" not in report_cols:
            conn.execute("ALTER TABLE reports ADD COLUMN penalty_rs INTEGER NOT NULL DEFAULT 0")
        if "caption" not in report_cols:
            conn.execute("ALTER TABLE reports ADD COLUMN caption TEXT DEFAULT ''")
        if "waived_on_time" not in report_cols:
            conn.execute("ALTER TABLE reports ADD COLUMN waived_on_time INTEGER NOT NULL DEFAULT 0")
        if "chat_id" not in report_cols:
            conn.execute("ALTER TABLE reports ADD COLUMN chat_id INTEGER")
        if "source_message_id" not in report_cols:
            conn.execute("ALTER TABLE reports ADD COLUMN source_message_id INTEGER")

        conn.execute(
            "INSERT OR IGNORE INTO bot_settings (key, value) VALUES (?, ?)",
            ("reminders_enabled", "true" if DEFAULT_ENABLE_REMINDERS else "false"),
        )
        conn.execute("INSERT OR IGNORE INTO bot_settings (key, value) VALUES ('clean_chat_mode', 'true')")
        conn.execute("INSERT OR IGNORE INTO bot_settings (key, value) VALUES ('holiday_mode', 'false')")
        conn.execute("INSERT OR IGNORE INTO bot_settings (key, value) VALUES ('off_days', '')")
        conn.execute("INSERT OR IGNORE INTO bot_settings (key, value) VALUES ('daily_joining_target', '0')")
        conn.execute("INSERT OR IGNORE INTO bot_settings (key, value) VALUES ('per_joining_rate', '0')")
        if DEFAULT_WORK_CHAT_ID is not None:
            conn.execute(
                "INSERT OR IGNORE INTO bot_settings (key, value) VALUES (?, ?)",
                ("work_group_id", str(DEFAULT_WORK_CHAT_ID)),
            )
            conn.execute(
                "INSERT OR IGNORE INTO tracked_chats (chat_id, title, role, added_at) VALUES (?, 'Primary Work Group', 'work', ?)",
                (DEFAULT_WORK_CHAT_ID, datetime.now(IST).isoformat()),
            )
        if DEFAULT_EXCEL_GROUP_ID is not None:
            conn.execute(
                "INSERT OR IGNORE INTO bot_settings (key, value) VALUES (?, ?)",
                ("excel_group_id", str(DEFAULT_EXCEL_GROUP_ID)),
            )

        existing_tags = conn.execute("SELECT COUNT(*) FROM assistant_tags").fetchone()[0]
        if existing_tags == 0 and DEFAULT_ASSISTANT_TAGS:
            now_iso = datetime.now(IST).isoformat()
            for tag in DEFAULT_ASSISTANT_TAGS:
                norm = _normalize_tag(tag)
                if norm:
                    conn.execute(
                        "INSERT OR IGNORE INTO assistant_tags (username, user_id, is_active, added_at) VALUES (?, NULL, 1, ?)",
                        (norm, now_iso),
                    )
                    conn.execute(
                        "INSERT OR IGNORE INTO team_members (username, user_id, added_at) VALUES (?, NULL, ?)",
                        (norm, now_iso),
                    )

        deduplicate_case_insensitive_records(conn)
        recalculate_report_statuses(conn)
        conn.commit()


def _pick_canonical_username(names: list[str]) -> str:
    """Choose the best display username among case variants (e.g. @Premgupta2m over PREMGUPTA2M)."""
    cleaned = [n.strip() for n in names if n and n.strip()]
    if not cleaned:
        return ""

    def _score(name: str) -> tuple[int, int, int]:
        has_at = 1 if name.startswith("@") else 0
        is_mixed = 1 if (name != name.upper() and name != name.lower()) else (0 if name != name.upper() else -1)
        return (has_at, is_mixed, len(name))

    return max(cleaned, key=_score)


def deduplicate_case_insensitive_records(conn: sqlite3.Connection) -> int:
    """
    Merge case-variant and alias duplicates across team_members, assistant_tags, summary_excluded, and reports
    (e.g. @PREMGUPTA2M + @Premgupta2m -> single canonical row) so assistants are never marked Missed twice.
    """
    removed = 0
    alias_rows: dict[str, tuple[str, str]] = {}
    has_alias_tbl = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name='username_aliases'"
    ).fetchone()
    if has_alias_tbl:
        for ak, tk, td in conn.execute(
            "SELECT alias_key, target_key, target_display FROM username_aliases"
        ).fetchall():
            alias_rows[str(ak)] = (str(tk), str(td))

    def _canon_key(raw: str) -> str:
        k = (raw or "").strip().lstrip("@").lower()
        return alias_rows[k][0] if k in alias_rows else k

    # 1. Deduplicate team_members
    tm_rows = conn.execute("SELECT rowid, username, user_id, added_at FROM team_members").fetchall()
    tm_groups: dict[str, list[tuple[int, str, Optional[int], str]]] = {}
    for rowid, uname, uid, added_at in tm_rows:
        k = _canon_key(uname or "")
        if k:
            tm_groups.setdefault(k, []).append((rowid, uname, uid, added_at))
    for k, grp in tm_groups.items():
        if len(grp) > 1 or any(g[1] != grp[0][1] for g in grp):
            canon = _pick_canonical_username([g[1] for g in grp])
            for ak, (tk, td) in alias_rows.items():
                if tk == k and td:
                    canon = td
                    break
            best_uid = next((g[2] for g in grp if g[2] is not None), None)
            best_added = min(str(g[3] or "") for g in grp)
            keep_id = grp[0][0]
            for rid, _, _, _ in grp[1:]:
                conn.execute("DELETE FROM team_members WHERE rowid = ?", (rid,))
                removed += 1
            conn.execute(
                "UPDATE team_members SET username = ?, user_id = ?, added_at = ? WHERE rowid = ?",
                (canon, best_uid, best_added, keep_id),
            )

    # 2. Deduplicate assistant_tags
    at_rows = conn.execute("SELECT rowid, username, user_id, is_active, added_at FROM assistant_tags").fetchall()
    at_groups: dict[str, list[tuple[int, str, Optional[int], int, str]]] = {}
    for rowid, uname, uid, is_act, added_at in at_rows:
        k = _canon_key(uname or "")
        if k:
            at_groups.setdefault(k, []).append((rowid, uname, uid, int(is_act or 0), added_at))
    for k, grp in at_groups.items():
        if len(grp) > 1:
            canon = _pick_canonical_username([g[1] for g in grp])
            for ak, (tk, td) in alias_rows.items():
                if tk == k and td:
                    canon = td
                    break
            best_uid = next((g[2] for g in grp if g[2] is not None), None)
            best_act = 1 if any(g[3] == 1 for g in grp) else 0
            best_added = min(str(g[4] or "") for g in grp)
            keep_id = grp[0][0]
            for rid, _, _, _, _ in grp[1:]:
                conn.execute("DELETE FROM assistant_tags WHERE rowid = ?", (rid,))
                removed += 1
            conn.execute(
                "UPDATE assistant_tags SET username = ?, user_id = ?, is_active = ?, added_at = ? WHERE rowid = ?",
                (canon, best_uid, best_act, best_added, keep_id),
            )

    # 3. Deduplicate summary_excluded
    se_rows = conn.execute("SELECT rowid, username FROM summary_excluded").fetchall()
    seen_se: set[str] = set()
    for rowid, uname in se_rows:
        k = _canon_key(uname or "")
        if k in seen_se:
            conn.execute("DELETE FROM summary_excluded WHERE rowid = ?", (rowid,))
            removed += 1
        elif k:
            seen_se.add(k)

    # 4. Canonicalize and deduplicate reports on (key, report_date, reporting_slot)
    canon_map: dict[str, str] = {}
    for ak, (tk, td) in alias_rows.items():
        if td:
            canon_map[tk] = td
    for (uname,) in conn.execute("SELECT username FROM team_members").fetchall():
        k = (uname or "").strip().lstrip("@").lower()
        if k and k not in canon_map:
            canon_map[k] = uname
    for (uname,) in conn.execute("SELECT username FROM assistant_tags").fetchall():
        k = (uname or "").strip().lstrip("@").lower()
        if k and k not in canon_map:
            canon_map[k] = uname

    rep_rows = conn.execute(
        "SELECT id, username, COALESCE(joining, ''), user_id, report_date, reporting_slot FROM reports ORDER BY id ASC"
    ).fetchall()
    rep_groups: dict[tuple[str, str, str], list[tuple[int, str, str, Optional[int]]]] = {}
    for rid, uname, joining, uid, rdate, rslot in rep_rows:
        k = _canon_key(uname or "")
        if not k:
            continue
        rep_groups.setdefault((k, str(rdate), str(rslot)), []).append((rid, uname, joining, uid))

    for (k, _d, _s), grp in rep_groups.items():
        canon = canon_map.get(k) or _pick_canonical_username([g[1] for g in grp])
        if len(grp) > 1:
            grp_sorted = sorted(grp, key=lambda g: (1 if g[2] else 0, 1 if g[3] is not None else 0, g[0]), reverse=True)
            keep_rid = grp_sorted[0][0]
            best_uid = next((g[3] for g in grp_sorted if g[3] is not None), None)
            for rid, _, _, _ in grp_sorted[1:]:
                conn.execute("DELETE FROM reports WHERE id = ?", (rid,))
                removed += 1
            conn.execute(
                "UPDATE reports SET username = ?, user_id = COALESCE(user_id, ?) WHERE id = ?",
                (canon, best_uid, keep_rid),
            )
        elif grp[0][1] != canon:
            conn.execute("UPDATE reports SET username = ? WHERE id = ?", (canon, grp[0][0]))

    return removed


def get_setting(key: str, default: Optional[str] = None) -> Optional[str]:
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute("SELECT value FROM bot_settings WHERE key = ?", (key,)).fetchone()
        return row[0] if row else default


def set_setting(key: str, value: str) -> None:
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            INSERT INTO bot_settings (key, value) VALUES (?, ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """,
            (key, value),
        )
        conn.commit()


def is_reminders_enabled() -> bool:
    val = get_setting("reminders_enabled", "true" if DEFAULT_ENABLE_REMINDERS else "false")
    return (val or "true").lower() == "true"


def is_clean_chat_enabled() -> bool:
    return (get_setting("clean_chat_mode", "true") or "true").lower() == "true"


def get_daily_joining_target() -> int:
    return _safe_int(get_setting("daily_joining_target", "0") or "0")


def get_per_joining_rate() -> int:
    return _safe_int(get_setting("per_joining_rate", "0") or "0")


def get_work_group_id() -> Optional[int]:
    val = get_setting("work_group_id")
    return int(val) if val else DEFAULT_WORK_CHAT_ID


def get_all_work_group_ids() -> list[int]:
    ids: list[int] = []
    primary = get_work_group_id()
    if primary is not None:
        ids.append(primary)
    try:
        with sqlite3.connect(DB_PATH) as conn:
            for (cid,) in conn.execute(
                "SELECT chat_id FROM tracked_chats WHERE role = 'work' ORDER BY chat_id ASC"
            ).fetchall():
                if int(cid) not in ids:
                    ids.append(int(cid))
    except Exception:
        pass
    return ids


def is_allowed_work_group(chat_id: int) -> bool:
    allowed = get_all_work_group_ids()
    if not allowed:
        return True
    return chat_id in allowed


def add_work_group(chat_id: int, title: str = "") -> None:
    now_iso = datetime.now(IST).isoformat()
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            INSERT INTO tracked_chats (chat_id, title, role, added_at) VALUES (?, ?, 'work', ?)
            ON CONFLICT(chat_id) DO UPDATE SET title = COALESCE(NULLIF(excluded.title, ''), tracked_chats.title), role = 'work'
            """,
            (chat_id, title or f"Group {chat_id}", now_iso),
        )
        if not get_setting("work_group_id"):
            conn.execute(
                "INSERT INTO bot_settings (key, value) VALUES ('work_group_id', ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                (str(chat_id),),
            )
        conn.commit()


def remove_work_group(chat_id: int) -> bool:
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute("DELETE FROM tracked_chats WHERE chat_id = ? AND role = 'work'", (chat_id,))
        conn.commit()
        return cur.rowcount > 0


def get_tracked_work_groups() -> list[tuple[int, str]]:
    primary = get_work_group_id()
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            "SELECT chat_id, COALESCE(title, '') FROM tracked_chats WHERE role = 'work' ORDER BY added_at ASC"
        ).fetchall()
    out: dict[int, str] = {int(cid): (ttl or f"Group {cid}") for cid, ttl in rows}
    if primary is not None and primary not in out:
        out[primary] = "Primary Work Group"
    return list(out.items())


def get_excel_group_id() -> Optional[int]:
    val = get_setting("excel_group_id") or get_setting("db_group_id")
    return int(val) if val else DEFAULT_EXCEL_GROUP_ID


def _match_key(username: str) -> str:
    return username.strip().lstrip("@").lower()


def get_summary_excluded() -> list[str]:
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            "SELECT username FROM summary_excluded ORDER BY username ASC"
        ).fetchall()
        return [r[0] for r in rows]


def is_summary_excluded(username: str) -> bool:
    key = _match_key(username)
    if not key:
        return False
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            "SELECT 1 FROM summary_excluded WHERE LOWER(LTRIM(username, '@')) = ? LIMIT 1",
            (key,),
        ).fetchone()
        return row is not None


def _resolve_canonical_name(conn: sqlite3.Connection, raw_username: str) -> str:
    clean = raw_username.strip()
    key = _match_key(clean)
    if not key:
        return ""
    has_alias = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name='username_aliases'"
    ).fetchone()
    if has_alias:
        al = conn.execute(
            "SELECT target_key, target_display FROM username_aliases WHERE alias_key = ? LIMIT 1",
            (key,),
        ).fetchone()
        if al and al[0]:
            key = str(al[0])
            clean = str(al[1] or clean)
    for table in ("team_members", "assistant_tags", "reports"):
        row = conn.execute(
            f"SELECT username FROM {table} WHERE LOWER(LTRIM(username, '@')) = ? LIMIT 1",
            (key,),
        ).fetchone()
        if row and row[0]:
            return str(row[0]).strip()
    return _normalize_tag(clean) if clean.startswith("@") else clean


def add_username_alias(old_username: str, new_username: str) -> tuple[str, str, int]:
    old_disp = _normalize_tag(old_username) or old_username.strip()
    new_disp = _normalize_tag(new_username) or new_username.strip()
    old_key = _match_key(old_disp)
    new_key = _match_key(new_disp)
    now_iso = datetime.now(IST).isoformat()
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            INSERT INTO username_aliases (alias_key, alias_display, target_key, target_display, created_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(alias_key) DO UPDATE SET
                alias_display = excluded.alias_display,
                target_key = excluded.target_key,
                target_display = excluded.target_display
            """,
            (old_key, old_disp, new_key, new_disp, now_iso),
        )
        conn.execute("UPDATE leaves SET username_key = ?, username = ? WHERE username_key = ?", (new_key, new_disp, old_key))
        conn.execute("UPDATE penalty_ledger SET username_key = ?, username = ? WHERE username_key = ?", (new_key, new_disp, old_key))
        conn.execute("UPDATE assistant_shifts SET username_key = ?, username = ? WHERE username_key = ?", (new_key, new_disp, old_key))
        merged = deduplicate_case_insensitive_records(conn)
        conn.commit()
    return old_disp, new_disp, merged


def get_username_aliases() -> list[tuple[str, str]]:
    with sqlite3.connect(DB_PATH) as conn:
        return conn.execute(
            "SELECT alias_display, target_display FROM username_aliases ORDER BY alias_display ASC"
        ).fetchall()


def add_leave(username: str, leave_date: str, slot_label: str = "", reason: str = "") -> str:
    now_iso = datetime.now(IST).isoformat()
    with sqlite3.connect(DB_PATH) as conn:
        canon = _resolve_canonical_name(conn, username) or _normalize_tag(username) or username.strip()
        key = _match_key(canon)
        conn.execute(
            """
            INSERT INTO leaves (username_key, username, leave_date, slot_label, reason, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(username_key, leave_date, slot_label) DO UPDATE SET
                username = excluded.username,
                reason = excluded.reason
            """,
            (key, canon, leave_date, slot_label, reason, now_iso),
        )
        conn.commit()
    return canon


def remove_leave(username: str, leave_date: str, slot_label: str = "") -> int:
    key = _match_key(username)
    with sqlite3.connect(DB_PATH) as conn:
        if slot_label:
            cur = conn.execute(
                "DELETE FROM leaves WHERE username_key = ? AND leave_date = ? AND slot_label = ?",
                (key, leave_date, slot_label),
            )
        else:
            cur = conn.execute(
                "DELETE FROM leaves WHERE username_key = ? AND leave_date = ?",
                (key, leave_date),
            )
        conn.commit()
        return cur.rowcount


def get_leaves_for_date(leave_date: str) -> dict[str, set[str]]:
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            "SELECT username_key, slot_label FROM leaves WHERE leave_date = ?",
            (leave_date,),
        ).fetchall()
    out: dict[str, set[str]] = {}
    for k, sl in rows:
        out.setdefault(str(k), set()).add(str(sl or ""))
    return out


def is_on_leave(username: str, leave_date: str, slot_label: str = "") -> bool:
    key = _match_key(username)
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            "SELECT slot_label FROM leaves WHERE username_key = ? AND leave_date = ?",
            (key, leave_date),
        ).fetchall()
    slots = {str(r[0] or "") for r in rows}
    if "" in slots:
        return True
    return bool(slot_label and slot_label in slots)


def set_assistant_shift(username: str, shift_name: str) -> str:
    norm_shift = shift_name.strip().lower()
    if norm_shift not in SHIFT_SLOTS:
        norm_shift = "full"
    now_iso = datetime.now(IST).isoformat()
    with sqlite3.connect(DB_PATH) as conn:
        canon = _resolve_canonical_name(conn, username) or _normalize_tag(username) or username.strip()
        key = _match_key(canon)
        conn.execute(
            """
            INSERT INTO assistant_shifts (username_key, username, shift_name, updated_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(username_key) DO UPDATE SET
                username = excluded.username,
                shift_name = excluded.shift_name,
                updated_at = excluded.updated_at
            """,
            (key, canon, norm_shift, now_iso),
        )
        conn.commit()
    return canon


def get_assistant_shifts() -> dict[str, str]:
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute("SELECT username_key, shift_name FROM assistant_shifts").fetchall()
    return {str(k): str(s) for k, s in rows}


def get_slots_for_member(username: str, shifts_map: Optional[dict[str, str]] = None) -> set[str]:
    m_key = _match_key(username)
    if shifts_map is None:
        shifts_map = get_assistant_shifts()
    shift_name = shifts_map.get(m_key, "full")
    return SHIFT_SLOTS.get(shift_name, SHIFT_SLOTS["full"])


def add_penalty_ledger_entry(
    username: str, entry_type: str, amount_rs: int, note: str = "", entry_date: Optional[str] = None
) -> str:
    now_ist = datetime.now(IST)
    d_str = entry_date or now_ist.strftime("%Y-%m-%d")
    with sqlite3.connect(DB_PATH) as conn:
        canon = _resolve_canonical_name(conn, username) or _normalize_tag(username) or username.strip()
        key = _match_key(canon)
        conn.execute(
            """
            INSERT INTO penalty_ledger (username_key, username, entry_type, amount_rs, note, entry_date, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (key, canon, entry_type, max(0, int(amount_rs)), note, d_str, now_ist.isoformat()),
        )
        conn.commit()
    return canon


def get_penalty_ledger_summary(
    start_date: Optional[str] = None, end_date: Optional[str] = None
) -> dict[str, tuple[int, int]]:
    """Return {username_key: (paid_rs, waived_rs)} for the given date filter."""
    query = "SELECT username_key, entry_type, SUM(amount_rs) FROM penalty_ledger WHERE 1=1"
    params: list[str] = []
    if start_date:
        query += " AND entry_date >= ?"
        params.append(start_date)
    if end_date:
        query += " AND entry_date <= ?"
        params.append(end_date)
    query += " GROUP BY username_key, entry_type"
    out: dict[str, list[int]] = {}
    with sqlite3.connect(DB_PATH) as conn:
        for uk, etype, total in conn.execute(query, params).fetchall():
            bucket = out.setdefault(str(uk), [0, 0])
            if etype == "paid":
                bucket[0] += int(total or 0)
            elif etype == "waive":
                bucket[1] += int(total or 0)
    return {k: (v[0], v[1]) for k, v in out.items()}


def get_member_ledger_entries(username: str, limit: int = 10) -> list[tuple[str, int, str, str]]:
    key = _match_key(username)
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            """
            SELECT entry_type, amount_rs, COALESCE(note, ''), entry_date
            FROM penalty_ledger WHERE username_key = ?
            ORDER BY id DESC LIMIT ?
            """,
            (key, limit),
        ).fetchall()
    return [(str(t), int(a or 0), str(n), str(d)) for t, a, n, d in rows]


def check_and_record_photo(
    file_unique_id: str, username: str, report_date: str, reporting_slot: str, message_id: Optional[int]
) -> Optional[tuple[str, str, str]]:
    """
    If file_unique_id was previously uploaded on a DIFFERENT message_id, return (prev_user, prev_date, prev_slot).
    Otherwise record it and return None.
    """
    if not file_unique_id:
        return None
    now_iso = datetime.now(IST).isoformat()
    key = _match_key(username)
    with sqlite3.connect(DB_PATH) as conn:
        prev = conn.execute(
            """
            SELECT username, report_date, reporting_slot, message_id
            FROM report_photos WHERE file_unique_id = ? ORDER BY id ASC LIMIT 1
            """,
            (file_unique_id,),
        ).fetchone()
        if prev:
            prev_uid_msg = prev[3]
            if message_id is not None and prev_uid_msg is not None and int(prev_uid_msg) == int(message_id):
                return None
            return str(prev[0]), str(prev[1]), str(prev[2])
        conn.execute(
            """
            INSERT INTO report_photos (file_unique_id, username_key, username, report_date, reporting_slot, message_id, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (file_unique_id, key, username, report_date, reporting_slot, message_id, now_iso),
        )
        conn.commit()
    return None


def get_previous_slot_joining_today(
    username: str, report_date: str, current_slot_label: str
) -> Optional[tuple[str, int]]:
    """Return (prev_slot_label, prev_joining) for an earlier slot submitted today by username."""
    key = _match_key(username)
    cur_rank = SLOT_ORDER.get(current_slot_label, 99)
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            """
            SELECT reporting_slot, COALESCE(joining, '')
            FROM reports WHERE LOWER(LTRIM(username, '@')) = ? AND report_date = ?
            """,
            (key, report_date),
        ).fetchall()
    candidates: list[tuple[int, str, int]] = []
    for sl, j_raw in rows:
        rank = SLOT_ORDER.get(str(sl), 99)
        parsed = _safe_int_joining(j_raw)
        if rank < cur_rank and parsed is not None:
            candidates.append((rank, str(sl), parsed))
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1], candidates[0][2]


def mark_report_waived_on_time(report_id: int) -> Optional[tuple[str, str, str]]:
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            "SELECT username, reporting_slot, report_date FROM reports WHERE id = ?",
            (report_id,),
        ).fetchone()
        if not row:
            return None
        conn.execute(
            "UPDATE reports SET status = 'On Time (Admin OK)', penalty_rs = 0, waived_on_time = 1 WHERE id = ?",
            (report_id,),
        )
        conn.commit()
        return str(row[0]), str(row[1]), str(row[2])


def delete_report_by_id(report_id: int) -> Optional[tuple[str, str, str]]:
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            "SELECT username, reporting_slot, report_date FROM reports WHERE id = ?",
            (report_id,),
        ).fetchone()
        if not row:
            return None
        conn.execute("DELETE FROM reports WHERE id = ?", (report_id,))
        conn.commit()
        return str(row[0]), str(row[1]), str(row[2])


def exclude_from_summary(username: str) -> Optional[str]:
    clean = username.strip()
    key = _match_key(clean)
    if not key:
        return None
    now_iso = datetime.now(IST).isoformat()
    with sqlite3.connect(DB_PATH) as conn:
        canonical = _resolve_canonical_name(conn, clean)
        if not canonical:
            return None
        existing = conn.execute(
            "SELECT username FROM summary_excluded WHERE LOWER(LTRIM(username, '@')) = ? LIMIT 1",
            (key,),
        ).fetchone()
        if existing:
            return str(existing[0])
        conn.execute(
            "INSERT OR IGNORE INTO summary_excluded (username, added_at) VALUES (?, ?)",
            (canonical, now_iso),
        )
        conn.commit()
        return canonical


def include_in_summary(username: str) -> bool:
    key = _match_key(username)
    if not key:
        return False
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute(
            "DELETE FROM summary_excluded WHERE LOWER(LTRIM(username, '@')) = ?",
            (key,),
        )
        conn.commit()
        return cur.rowcount > 0


def toggle_summary_exclusion(username: str) -> bool:
    """Returns True if now REMOVED/hidden from summary, False if now SHOWN."""
    if is_summary_excluded(username):
        include_in_summary(username)
        return False
    exclude_from_summary(username)
    return True


def get_known_admin_sets() -> tuple[set[int], set[str]]:
    admin_ids: set[int] = set()
    admin_keys: set[str] = set()
    if DEFAULT_ADMIN_CHAT_ID is not None:
        admin_ids.add(DEFAULT_ADMIN_CHAT_ID)
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            "SELECT user_id, username FROM known_group_admins"
        ).fetchall()
    for uid, uname in rows:
        if uid is not None:
            admin_ids.add(int(uid))
        k = _match_key(str(uname or ""))
        if k:
            admin_keys.add(k)
    return admin_ids, admin_keys


def is_known_group_admin(username: str = "", user_id: Optional[int] = None) -> bool:
    admin_ids, admin_keys = get_known_admin_sets()
    if user_id is not None and user_id in admin_ids:
        return True
    key = _match_key(username)
    return bool(key and key in admin_keys)


def sync_group_admins(admin_pairs: list[tuple[Optional[int], str]]) -> int:
    """Persist live Telegram group admins and purge them from assistant_tags & team_members."""
    now_iso = datetime.now(IST).isoformat()
    removed = 0
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("DELETE FROM known_group_admins")
        for uid, uname in admin_pairs:
            clean_u = (uname or "").strip()
            key = _match_key(clean_u)
            admin_key = f"id:{uid}" if uid is not None else f"u:{key}"
            if not admin_key or (uid is None and not key):
                continue
            conn.execute(
                """
                INSERT OR REPLACE INTO known_group_admins (admin_key, user_id, username, updated_at)
                VALUES (?, ?, ?, ?)
                """,
                (admin_key, uid, clean_u or f"Admin({uid})", now_iso),
            )
            if uid is not None:
                c1 = conn.execute("DELETE FROM team_members WHERE user_id = ?", (uid,))
                c2 = conn.execute("DELETE FROM assistant_tags WHERE user_id = ?", (uid,))
                removed += c1.rowcount + c2.rowcount
            if key:
                c3 = conn.execute(
                    "DELETE FROM team_members WHERE LOWER(LTRIM(username, '@')) = ?",
                    (key,),
                )
                c4 = conn.execute(
                    "DELETE FROM assistant_tags WHERE LOWER(LTRIM(username, '@')) = ?",
                    (key,),
                )
                removed += c3.rowcount + c4.rowcount
        conn.commit()
    return removed


def _upsert_team_member_case_insensitive(
    conn: sqlite3.Connection, username: str, user_id: Optional[int], now_iso: str
) -> tuple[bool, str]:
    clean = username.strip()
    key = _match_key(clean)
    if not clean or not key:
        return False, clean
    existing = conn.execute(
        "SELECT rowid, username, user_id FROM team_members WHERE LOWER(LTRIM(username, '@')) = ? LIMIT 1",
        (key,),
    ).fetchone()
    if existing:
        canon = _pick_canonical_username([existing[1], clean])
        new_uid = user_id if user_id is not None else existing[2]
        if canon != existing[1] or new_uid != existing[2]:
            conn.execute(
                "UPDATE team_members SET username = ?, user_id = COALESCE(?, user_id) WHERE rowid = ?",
                (canon, new_uid, existing[0]),
            )
        return False, canon
    conn.execute(
        "INSERT OR IGNORE INTO team_members (username, user_id, added_at) VALUES (?, ?, ?)",
        (clean, user_id, now_iso),
    )
    return True, clean


def _upsert_assistant_tag_case_insensitive(
    conn: sqlite3.Connection,
    username: str,
    user_id: Optional[int],
    now_iso: str,
    force_active: Optional[int] = None,
) -> tuple[bool, str]:
    clean = username.strip()
    key = _match_key(clean)
    if not clean or not key:
        return False, clean
    existing = conn.execute(
        "SELECT rowid, username, user_id, is_active FROM assistant_tags WHERE LOWER(LTRIM(username, '@')) = ? LIMIT 1",
        (key,),
    ).fetchone()
    if existing:
        canon = _pick_canonical_username([existing[1], clean])
        new_uid = user_id if user_id is not None else existing[2]
        new_act = force_active if force_active is not None else int(existing[3])
        if canon != existing[1] or new_uid != existing[2] or new_act != int(existing[3]):
            conn.execute(
                "UPDATE assistant_tags SET username = ?, user_id = COALESCE(?, user_id), is_active = ? WHERE rowid = ?",
                (canon, new_uid, new_act, existing[0]),
            )
        return False, canon
    act_val = 1 if force_active is None else force_active
    conn.execute(
        "INSERT OR IGNORE INTO assistant_tags (username, user_id, is_active, added_at) VALUES (?, ?, ?, ?)",
        (clean, user_id, act_val, now_iso),
    )
    return True, clean


def auto_discover_assistant(username: str, user_id: Optional[int] = None) -> bool:
    """
    Automatically register any non-admin group member as an assistant on their first message.
    Case-insensitive deduplication prevents @PREMGUPTA2M + @Premgupta2m duplicates.
    """
    clean = username.strip()
    key = _match_key(clean)
    if not clean or not key:
        return False
    if is_known_group_admin(clean, user_id):
        return False
    now_iso = datetime.now(IST).isoformat()
    with sqlite3.connect(DB_PATH) as conn:
        if conn.execute(
            "SELECT 1 FROM summary_excluded WHERE LOWER(LTRIM(username, '@')) = ? LIMIT 1",
            (key,),
        ).fetchone():
            return False
        _, canon = _upsert_team_member_case_insensitive(conn, clean, user_id, now_iso)
        inserted, _ = _upsert_assistant_tag_case_insensitive(conn, canon, user_id, now_iso)
        conn.commit()
        return inserted


def add_team_member(username: str) -> bool:
    clean = username.strip()
    key = _match_key(clean)
    if not clean or not key:
        return False
    now_iso = datetime.now(IST).isoformat()
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            "DELETE FROM summary_excluded WHERE LOWER(LTRIM(username, '@')) = ?",
            (key,),
        )
        inserted, _ = _upsert_team_member_case_insensitive(conn, clean, None, now_iso)
        conn.commit()
        return inserted


def remove_team_member(username: str) -> bool:
    clean = username.strip()
    key = _match_key(clean)
    if not key:
        return False
    now_iso = datetime.now(IST).isoformat()
    with sqlite3.connect(DB_PATH) as conn:
        canonical = _resolve_canonical_name(conn, clean)
        cur = conn.execute(
            "DELETE FROM team_members WHERE LOWER(LTRIM(username, '@')) = ?",
            (key,),
        )
        cur_tag = conn.execute(
            "DELETE FROM assistant_tags WHERE LOWER(LTRIM(username, '@')) = ?",
            (key,),
        )
        already_excl = conn.execute(
            "SELECT 1 FROM summary_excluded WHERE LOWER(LTRIM(username, '@')) = ? LIMIT 1",
            (key,),
        ).fetchone()
        if canonical and not already_excl:
            conn.execute(
                "INSERT OR IGNORE INTO summary_excluded (username, added_at) VALUES (?, ?)",
                (canonical, now_iso),
            )
        conn.commit()
        return cur.rowcount > 0 or cur_tag.rowcount > 0 or not bool(already_excl)


def get_team_members() -> list[str]:
    admin_ids, admin_keys = get_known_admin_sets()
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            "SELECT username, user_id FROM team_members ORDER BY username ASC"
        ).fetchall()
    filtered = [
        u
        for u, uid in rows
        if (uid is None or uid not in admin_ids) and _match_key(u) not in admin_keys
    ]
    return _dedup_case_insensitive(filtered)


def get_tagged_assistants_with_ids() -> list[tuple[str, Optional[int]]]:
    admin_ids, admin_keys = get_known_admin_sets()
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            "SELECT username, user_id FROM assistant_tags WHERE is_active = 1 ORDER BY username ASC"
        ).fetchall()
    by_key: dict[str, tuple[str, Optional[int]]] = {}
    for u, uid in rows:
        key = _match_key(u)
        if not key or key in admin_keys or (uid is not None and uid in admin_ids):
            continue
        if key not in by_key:
            by_key[key] = (u, uid)
        else:
            prev_u, prev_uid = by_key[key]
            by_key[key] = (
                _pick_canonical_username([prev_u, u]),
                uid if uid is not None else prev_uid,
            )
    return sorted(by_key.values(), key=lambda x: x[0].lower())


def get_tagged_assistants() -> list[str]:
    return [u for u, _uid in get_tagged_assistants_with_ids()]


def add_assistant_tag(username: str) -> bool:
    norm = _normalize_tag(username)
    if not norm:
        return False
    now_iso = datetime.now(IST).isoformat()
    with sqlite3.connect(DB_PATH) as conn:
        _, canon = _upsert_assistant_tag_case_insensitive(conn, norm, None, now_iso, force_active=1)
        _upsert_team_member_case_insensitive(conn, canon, None, now_iso)
        conn.commit()
        return True


def remove_assistant_tag(username: str) -> bool:
    key = _match_key(username)
    if not key:
        return False
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute(
            "UPDATE assistant_tags SET is_active = 0 WHERE LOWER(LTRIM(username, '@')) = ?",
            (key,),
        )
        conn.commit()
        return cur.rowcount > 0


def replace_assistant_tags(usernames: list[str]) -> list[str]:
    now_iso = datetime.now(IST).isoformat()
    cleaned = _dedup_case_insensitive([_normalize_tag(raw) for raw in usernames if _normalize_tag(raw)])

    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("UPDATE assistant_tags SET is_active = 0")
        for norm in cleaned:
            _, canon = _upsert_assistant_tag_case_insensitive(conn, norm, None, now_iso, force_active=1)
            _upsert_team_member_case_insensitive(conn, canon, None, now_iso)
        conn.commit()
    return cleaned


def clear_assistant_tags() -> None:
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("UPDATE assistant_tags SET is_active = 0")
        conn.commit()


def toggle_assistant_tag(username: str) -> bool:
    tagged_keys = {_match_key(u) for u in get_tagged_assistants()}
    if _match_key(username) in tagged_keys:
        remove_assistant_tag(username)
        return False
    add_assistant_tag(username)
    return True


def has_slot_submission(
    username: str, report_date: str, reporting_slot: str
) -> Optional[tuple[int, str, str]]:
    key = _match_key(username)
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            """
            SELECT id, submission_time, COALESCE(joining, '') FROM reports
            WHERE LOWER(LTRIM(username, '@')) = ? AND report_date = ? AND reporting_slot = ?
            LIMIT 1
            """,
            (key, report_date, reporting_slot),
        ).fetchone()
        return (int(row[0]), str(row[1]), str(row[2])) if row else None


def get_pending_tagged_assistants_for_slot(
    report_date: str, slot_label: str
) -> list[tuple[str, Optional[int]]]:
    is_off, _ = is_date_holiday_or_off(report_date)
    if is_off:
        return []
    tagged = get_tagged_assistants_with_ids()
    if not tagged:
        return []
    leaves_map = get_leaves_for_date(report_date)
    shifts_map = get_assistant_shifts()
    with sqlite3.connect(DB_PATH) as conn:
        submitted_rows = conn.execute(
            """
            SELECT DISTINCT LOWER(LTRIM(username, '@')), user_id FROM reports
            WHERE report_date = ? AND reporting_slot = ?
            """,
            (report_date, slot_label),
        ).fetchall()
    submitted_keys = {str(r[0]) for r in submitted_rows if r[0]}
    submitted_ids = {int(r[1]) for r in submitted_rows if r[1] is not None}
    pending: list[tuple[str, Optional[int]]] = []
    for u, uid in tagged:
        ukey = _match_key(u)
        if ukey in submitted_keys or (uid is not None and uid in submitted_ids):
            continue
        user_leaves = leaves_map.get(ukey, set())
        if "" in user_leaves or slot_label in user_leaves:
            continue
        if slot_label not in get_slots_for_member(u, shifts_map):
            continue
        pending.append((u, uid))
    return pending


def get_report_row_id(username: str, report_date: str, reporting_slot: str) -> Optional[int]:
    key = _match_key(username)
    with sqlite3.connect(DB_PATH) as conn:
        canon = _resolve_canonical_name(conn, username) or username
        ckey = _match_key(canon)
        row = conn.execute(
            """
            SELECT id FROM reports
            WHERE LOWER(LTRIM(username, '@')) IN (?, ?) AND report_date = ? AND reporting_slot = ?
            LIMIT 1
            """,
            (key, ckey, report_date, reporting_slot),
        ).fetchone()
        return int(row[0]) if row else None


def upsert_report_to_db(
    user_id: Optional[int],
    username: str,
    joining: str,
    submission_time: str,
    reporting_slot: str,
    report_date: str,
    status: str,
    caption: str,
    created_at_utc: str,
    overwrite_existing: bool = True,
    penalty_rs: int = 0,
    chat_id: Optional[int] = None,
    source_message_id: Optional[int] = None,
) -> str:
    with sqlite3.connect(DB_PATH) as conn:
        canonical = _resolve_canonical_name(conn, username) or username
        key = _match_key(canonical)
        existing = conn.execute(
            """
            SELECT id, COALESCE(joining, ''), COALESCE(caption, ''), COALESCE(waived_on_time, 0) FROM reports
            WHERE LOWER(LTRIM(username, '@')) = ? AND report_date = ? AND reporting_slot = ?
            LIMIT 1
            """,
            (key, report_date, reporting_slot),
        ).fetchone()

        if existing:
            row_id = existing[0]
            if overwrite_existing:
                new_joining = joining if joining != "" else existing[1]
                new_caption = caption if caption else existing[2]
                is_waived = int(existing[3] or 0) == 1
                final_status = "On Time (Admin OK)" if is_waived else status
                final_pen = 0 if is_waived else penalty_rs
                conn.execute(
                    """
                    UPDATE reports
                    SET user_id = COALESCE(?, user_id),
                        username = ?,
                        joining = ?,
                        submission_time = ?,
                        status = ?,
                        penalty_rs = ?,
                        caption = ?,
                        chat_id = COALESCE(?, chat_id),
                        source_message_id = COALESCE(?, source_message_id)
                    WHERE id = ?
                    """,
                    (
                        user_id,
                        canonical,
                        new_joining,
                        submission_time,
                        final_status,
                        final_pen,
                        new_caption,
                        chat_id,
                        source_message_id,
                        row_id,
                    ),
                )
                conn.commit()
                action = "updated"
            else:
                action = "skipped"
        else:
            conn.execute(
                """
                INSERT INTO reports (
                    user_id, username, joining, submission_time, reporting_slot, report_date, status, penalty_rs, caption, created_at_utc, chat_id, source_message_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    user_id,
                    canonical,
                    joining,
                    submission_time,
                    reporting_slot,
                    report_date,
                    status,
                    penalty_rs,
                    caption,
                    created_at_utc,
                    chat_id,
                    source_message_id,
                ),
            )
            conn.commit()
            action = "inserted"

    if not is_known_group_admin(canonical, user_id):
        auto_discover_assistant(canonical, user_id=user_id)
    return action


def fetch_reports(
    date_filter: Optional[str] = None,
    start_date: Optional[str] = None,
) -> list[tuple[str, str, str, str, str, str, int, str]]:
    with sqlite3.connect(DB_PATH) as conn:
        if date_filter:
            cursor = conn.execute(
                """
                SELECT username, COALESCE(joining, ''), submission_time, reporting_slot, report_date, status, COALESCE(penalty_rs, 0), COALESCE(caption, '')
                FROM reports
                WHERE report_date = ?
                ORDER BY id ASC
                """,
                (date_filter,),
            )
        elif start_date:
            cursor = conn.execute(
                """
                SELECT username, COALESCE(joining, ''), submission_time, reporting_slot, report_date, status, COALESCE(penalty_rs, 0), COALESCE(caption, '')
                FROM reports
                WHERE report_date >= ?
                ORDER BY report_date ASC, id ASC
                """,
                (start_date,),
            )
        else:
            cursor = conn.execute(
                """
                SELECT username, COALESCE(joining, ''), submission_time, reporting_slot, report_date, status, COALESCE(penalty_rs, 0), COALESCE(caption, '')
                FROM reports
                ORDER BY id ASC
                """
            )
        return [
            (r[0], r[1], r[2], r[3], r[4], r[5], int(r[6] or 0), r[7])
            for r in cursor.fetchall()
        ]


def fetch_user_today_slots(username: str, report_date: str) -> dict[str, tuple[str, str, str, int]]:
    key = _match_key(username)
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            """
            SELECT reporting_slot, submission_time, COALESCE(joining, ''), status, COALESCE(penalty_rs, 0)
            FROM reports
            WHERE LOWER(LTRIM(username, '@')) = ? AND report_date = ?
            """,
            (key, report_date),
        ).fetchall()
    return {
        slot: (sub_time, joining, status, int(pen or 0))
        for slot, sub_time, joining, status, pen in rows
    }


def count_total_reports() -> int:
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute("SELECT COUNT(*) FROM reports").fetchone()
        return int(row[0]) if row else 0


# ---------------------------------------------------------------------------
# Styled Excel (.xlsx) Workbook Builder (5 Sheets) & Restore Parser
# ---------------------------------------------------------------------------
def _style_sheet(
    ws, headers: list[str], rows: list[tuple[str, ...]], status_col_idx: Optional[int] = None
) -> None:
    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    on_time_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    on_time_font = Font(color="166534", bold=True)
    warn5_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    warn5_font = Font(color="92400E", bold=True)
    late_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    late_font = Font(color="991B1B", bold=True)

    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="left", vertical="center")

    ws.freeze_panes = "A2"

    for row_idx, row in enumerate(rows, start=2):
        ws.append(list(row))
        if status_col_idx is not None:
            status_val = str(row[status_col_idx - 1])
            status_cell = ws.cell(row=row_idx, column=status_col_idx)
            if status_val == "On Time":
                status_cell.fill = on_time_fill
                status_cell.font = on_time_font
            elif "₹5" in status_val or "₹10" in status_val:
                status_cell.fill = warn5_fill
                status_cell.font = warn5_font
            elif any(w in status_val.lower() for w in ("late", "missed", "outside", "₹15")):
                status_cell.fill = late_fill
                status_cell.font = late_font

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row in rows:
            val = str(row[col_idx - 1])
            if len(val) > max_len:
                max_len = len(val)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 5, 48)


def _safe_int(val: str) -> int:
    try:
        return int(str(val).strip())
    except ValueError:
        return 0


def _dedup_case_insensitive(items: list[str]) -> list[str]:
    seen: dict[str, str] = {}
    for item in items:
        clean = item.strip()
        if not clean:
            continue
        key = _match_key(clean)
        if not key:
            continue
        if key not in seen:
            seen[key] = clean
        else:
            seen[key] = _pick_canonical_username([seen[key], clean])
    return sorted(seen.values(), key=lambda x: x.lower())


def _dates_for_period(start_date: str, end_date: str) -> list[str]:
    cur = datetime.strptime(start_date, "%Y-%m-%d").date()
    end_d = datetime.strptime(end_date, "%Y-%m-%d").date()
    out: list[str] = []
    while cur <= end_d:
        out.append(cur.strftime("%Y-%m-%d"))
        cur += timedelta(days=1)
    return out


def _get_fine_reset_state() -> tuple[str, Optional[int], dict[str, tuple[str, Optional[int]]]]:
    reset_date = (get_setting("fine_reset_date", "") or "").strip()
    reset_secs = _parse_optional_int(get_setting("fine_reset_time_secs", "") or "", "fine_reset_time_secs")
    raw_map = get_setting("member_fine_reset_dates", "{}") or "{}"
    m_map: dict[str, tuple[str, Optional[int]]] = {}
    try:
        parsed = json.loads(raw_map)
        if isinstance(parsed, dict):
            for k, v in parsed.items():
                if isinstance(v, dict) and v.get("date"):
                    m_map[str(k)] = (str(v["date"]), _parse_optional_int(str(v.get("secs", "") or ""), "member_reset_secs"))
                elif isinstance(v, str) and v:
                    m_map[str(k)] = (v, None)
    except Exception:
        pass
    return reset_date, reset_secs, m_map


def _get_member_first_active_dates() -> dict[str, str]:
    out: dict[str, str] = {}
    with sqlite3.connect(DB_PATH) as conn:
        for uname, min_d in conn.execute(
            "SELECT LOWER(LTRIM(username, '@')), MIN(report_date) FROM reports GROUP BY LOWER(LTRIM(username, '@'))"
        ).fetchall():
            if uname and min_d:
                out[str(uname)] = str(min_d)
        for tbl in ("team_members", "assistant_tags"):
            for uname, added_at in conn.execute(
                f"SELECT LOWER(LTRIM(username, '@')), COALESCE(added_at, '') FROM {tbl}"
            ).fetchall():
                k = str(uname or "")
                d_part = str(added_at or "")[:10]
                if k and re.match(r"^\d{4}-\d{2}-\d{2}$", d_part):
                    if k not in out or d_part < out[k]:
                        out[k] = d_part
    return out


def compute_member_stats_for_period(
    members: list[str],
    period_rows: list[tuple[str, str, str, str, str, str, int, str]],
    date_list: list[str],
    now_ist: Optional[datetime] = None,
) -> list[dict[str, object]]:
    """
    Compute accurate member metrics across date_list:
      - Uses cumulative daily joining (latest slot value per day, NOT sum of slots!)
      - Respects dynamic penalty rates, leaves, holidays/off-days, assistant shifts, and /resetfine baselines
      - Never charges phantom Missed fines on empty/inactive historical days before tracking or member join
      - Includes Paid & Waived ledger adjustments to compute Balance Due
    """
    ref = now_ist or datetime.now(IST)
    today_str = ref.strftime("%Y-%m-%d")
    reg_pen, fin_pen, missed_pen_rate = get_penalty_config()
    shifts_map = get_assistant_shifts()
    target_val = get_daily_joining_target()
    leaves_by_date = {d: get_leaves_for_date(d) for d in date_list}
    ledger_map = get_penalty_ledger_summary(
        start_date=date_list[0] if date_list else None,
        end_date=date_list[-1] if date_list else None,
    )
    fine_reset_date, fine_reset_secs, member_reset_map = _get_fine_reset_state()
    member_first_dates = _get_member_first_active_dates()
    active_report_dates = {str(r[4]) for r in period_rows}

    stats: list[dict[str, object]] = []
    for member in members:
        m_key = _match_key(member)
        member_rows = [r for r in period_rows if _match_key(r[0]) == m_key]
        allowed_slots = get_slots_for_member(member, shifts_map)
        shift_name = shifts_map.get(m_key, "full")

        on_time = sum(1 for r in member_rows if r[5].startswith("On Time"))
        late_reg = sum(
            1
            for r in member_rows
            if r[6] > 0 and not _is_final_slot(r[3])
        )
        late_fin = sum(
            1
            for r in member_rows
            if r[6] > 0 and _is_final_slot(r[3])
        )
        other_late = sum(1 for r in member_rows if not r[5].startswith("On Time") and r[6] == 0)
        late_pen = sum(r[6] for r in member_rows)
        total_joining = compute_period_joining_from_rows(member_rows)

        m_reset = member_reset_map.get(m_key)
        m_first_date = member_first_dates.get(m_key)

        missed = 0
        expected_slots = 0
        leave_days = 0
        for d_str in date_list:
            # Skip historical days with zero team activity (prevents fake 30-day missed fines on fresh/empty DBs)
            if d_str < today_str and d_str not in active_report_dates:
                continue
            if fine_reset_date and d_str < fine_reset_date:
                continue
            if m_reset and d_str < m_reset[0]:
                continue
            if m_first_date and d_str < m_first_date:
                continue

            d_closed = get_closed_slots_for_date(d_str, ref)
            d_leaves = leaves_by_date.get(d_str, {}).get(m_key, set())
            if "" in d_leaves:
                leave_days += 1
                continue
            d_submitted = {r[3] for r in member_rows if r[4] == d_str}
            for s in d_closed:
                if s.label not in allowed_slots or s.label in d_leaves:
                    continue
                s_end_secs = _time_to_secs(s.end)
                if fine_reset_date == d_str and fine_reset_secs is not None and s_end_secs <= fine_reset_secs:
                    continue
                if m_reset and m_reset[0] == d_str and m_reset[1] is not None and s_end_secs <= int(m_reset[1]):
                    continue
                expected_slots += 1
                if s.label not in d_submitted:
                    missed += 1

        missed_pen = missed * missed_pen_rate
        gross_pen = late_pen + missed_pen
        paid_rs, waived_rs = ledger_map.get(m_key, (0, 0))
        balance_due = max(0, gross_pen - paid_rs - waived_rs)
        compliance = round((on_time / expected_slots) * 100) if expected_slots else 100
        active_day_cnt = sum(
            1 for d_str in date_list
            if (d_str == today_str or d_str in active_report_dates)
            and not (fine_reset_date and d_str < fine_reset_date)
            and not (m_first_date and d_str < m_first_date)
        )
        target_diff = (total_joining - (target_val * max(1, active_day_cnt - leave_days))) if target_val > 0 else 0

        stats.append(
            {
                "member": member,
                "key": m_key,
                "shift": shift_name,
                "allowed_slots": allowed_slots,
                "submitted_count": len(member_rows),
                "expected_slots": expected_slots,
                "on_time": on_time,
                "late_regular": late_reg,
                "late_final": late_fin + other_late,
                "missed": missed,
                "late_pen": late_pen,
                "missed_pen": missed_pen,
                "gross_pen": gross_pen,
                "paid_rs": paid_rs,
                "waived_rs": waived_rs,
                "balance_due": balance_due,
                "total_joining": total_joining,
                "compliance": compliance,
                "leave_days": leave_days,
                "target_diff": target_diff,
                "reg_pen_rate": reg_pen,
                "fin_pen_rate": fin_pen,
                "missed_pen_rate": missed_pen_rate,
            }
        )
    return stats


def build_excel_bytes(
    target_date: Optional[str] = None,
    days: int = 1,
    month_str: Optional[str] = None,
) -> tuple[BytesIO, int]:
    now_ist = datetime.now(IST)
    today_str = target_date or now_ist.strftime("%Y-%m-%d")
    closed_today_slots = get_closed_slots_for_date(today_str, now_ist)
    closed_today_labels = {s.label for s in closed_today_slots}
    missed_rate = get_missed_penalty_rs()
    reg_rate, fin_rate, _ = get_penalty_config()

    all_rows = fetch_reports()
    if month_str:
        start_str = f"{month_str}-01"
        y, m = map(int, month_str.split("-"))
        next_m = datetime(y + (1 if m == 12 else 0), (m % 12) + 1, 1) - timedelta(days=1)
        end_str = min(next_m.strftime("%Y-%m-%d"), now_ist.strftime("%Y-%m-%d"))
        period_rows = [r for r in all_rows if start_str <= r[4] <= end_str]
        date_list = _dates_for_period(start_str, end_str)
        period_title = f"Month {month_str}"
        ref_label = month_str
    elif days > 1:
        start_dt = datetime.strptime(today_str, "%Y-%m-%d") - timedelta(days=days - 1)
        start_str = start_dt.strftime("%Y-%m-%d")
        period_rows = [r for r in all_rows if start_str <= r[4] <= today_str]
        date_list = _dates_for_period(start_str, today_str)
        period_title = f"Last {days} Days"
        ref_label = today_str
    else:
        period_rows = [r for r in all_rows if r[4] == today_str]
        date_list = [today_str]
        period_title = f"Reports {today_str}"
        ref_label = today_str

    excluded_keys = {_match_key(u) for u in get_summary_excluded()}
    _admin_ids, admin_keys = get_known_admin_sets()
    hidden_keys = excluded_keys | admin_keys
    members = [
        m
        for m in _dedup_case_insensitive(
            get_team_members() + get_tagged_assistants() + [r[0] for r in period_rows]
        )
        if _match_key(m) not in hidden_keys
    ]
    slot_labels = [s.label for s in REPORTING_SLOTS]

    wb = Workbook()

    ws_log = wb.active
    ws_log.title = period_title[:31]
    log_headers = [
        "Username",
        "Joining",
        "Submission Time",
        "Reporting Slot",
        "Date",
        "Status",
        "Penalty (₹)",
        "Caption",
    ]
    period_excel_rows = [
        (r[0], r[1], r[2], r[3], r[4], r[5], f"₹{r[6]}", r[7]) for r in period_rows
    ]
    all_excel_rows = [
        (r[0], r[1], r[2], r[3], r[4], r[5], f"₹{r[6]}", r[7]) for r in all_rows
    ]
    _style_sheet(ws_log, log_headers, period_excel_rows, status_col_idx=6)

    stats_list = compute_member_stats_for_period(members, period_rows, date_list, now_ist)
    target_val = get_daily_joining_target()

    ws_summary = wb.create_sheet(title="Member Summary")
    summary_rows: list[tuple[str, ...]] = []
    for st in stats_list:
        tgt_str = (
            f"{'+' if st['target_diff'] >= 0 else ''}{st['target_diff']}"
            if target_val > 0
            else "—"
        )
        summary_rows.append(
            (
                str(st["member"]),
                str(st["shift"]).upper(),
                str(st["total_joining"]),
                tgt_str,
                str(st["on_time"]),
                str(st["late_regular"]),
                str(st["late_final"]),
                str(st["missed"]),
                f"₹{st['late_pen']}",
                f"₹{st['missed_pen']}",
                f"₹{st['paid_rs']} / ₹{st['waived_rs']}",
                f"₹{st['balance_due']}",
                f"{st['compliance']}%",
                ref_label,
            )
        )

    _style_sheet(
        ws_summary,
        [
            "Username",
            "Shift",
            "Valid Joining",
            "Target Diff",
            "Submitted On Time",
            f"Late (₹{reg_rate})",
            f"Final Late (₹{fin_rate})",
            "Missed Slots",
            "Late Penalty",
            f"Missed Penalty (₹{missed_rate})",
            "Paid / Waived",
            "Balance Due",
            "Compliance %",
            "Reference",
        ],
        summary_rows,
    )

    target_day_rows = fetch_reports(date_filter=today_str)
    leaves_today = get_leaves_for_date(today_str)
    shifts_map = get_assistant_shifts()
    is_off_today, _ = is_date_holiday_or_off(today_str)

    ws_matrix = wb.create_sheet(title="Slot Matrix")
    matrix_rows: list[tuple[str, ...]] = []
    for member in members:
        m_key = _match_key(member)
        member_day_rows = [r for r in target_day_rows if _match_key(r[0]) == m_key]
        member_day_slots = {r[3]: (r[1], r[2], r[5], r[6]) for r in member_day_rows}
        member_leaves = leaves_today.get(m_key, set())
        member_allowed = get_slots_for_member(member, shifts_map)
        row_cells = [member]
        for label in slot_labels:
            if label in member_day_slots:
                j_val, t_val, st_val, pen_val = member_day_slots[label]
                if st_val.startswith("On Time"):
                    row_cells.append(
                        f"Join: {j_val} ({t_val})" if j_val else f"On Time ({t_val})"
                    )
                else:
                    pen_tag = f"Late ₹{pen_val}" if pen_val > 0 else st_val
                    row_cells.append(
                        f"{pen_tag}: {j_val} ({t_val})" if j_val else f"{pen_tag} ({t_val})"
                    )
            elif is_off_today:
                row_cells.append("Holiday/Off")
            elif "" in member_leaves or label in member_leaves:
                row_cells.append("On Leave")
            elif label not in member_allowed:
                row_cells.append("Off Shift")
            elif label in closed_today_labels:
                row_cells.append(f"Missed (₹{missed_rate})")
            else:
                row_cells.append("Pending")
        # Valid daily cumulative joining (latest slot value, NOT sum of snapshots!)
        day_joining = compute_daily_joining_from_rows(member_day_rows)
        row_cells.append(str(day_joining))
        matrix_rows.append(tuple(row_cells))

    _style_sheet(
        ws_matrix,
        ["Username", "10:00 AM", "1:00 PM", "3:00 PM", "5:00 PM", "Final (<8 PM)", "Valid Daily Joining"],
        matrix_rows,
    )

    ws_all = wb.create_sheet(title="All History")
    _style_sheet(ws_all, log_headers, all_excel_rows, status_col_idx=6)

    # 5th Sheet: Bot Config (persists Work Group ID, Excel Group ID, Reminders, Roster & Tags across Render restarts)
    ws_cfg = wb.create_sheet(title="Bot Config")
    cfg_rows: list[tuple[str, str, str]] = []
    work_gid = get_work_group_id()
    excel_gid = get_excel_group_id()
    if work_gid is not None:
        cfg_rows.append(("work_group_id", str(work_gid), ""))
    if excel_gid is not None:
        cfg_rows.append(("excel_group_id", str(excel_gid), ""))
    cfg_rows.append(("reminders_enabled", "true" if is_reminders_enabled() else "false", ""))
    cfg_rows.append(("penalty_config", f"{reg_rate},{fin_rate},{missed_rate}", ""))
    cfg_rows.append(("daily_joining_target", str(target_val), ""))
    cfg_rows.append(("per_joining_rate", str(get_per_joining_rate()), ""))

    with sqlite3.connect(DB_PATH) as conn:
        tag_rows = conn.execute(
            "SELECT username, user_id, is_active FROM assistant_tags ORDER BY username ASC"
        ).fetchall()
        team_rows = conn.execute(
            "SELECT username, user_id FROM team_members ORDER BY username ASC"
        ).fetchall()
    tag_map: dict[str, tuple[str, Optional[int], int]] = {}
    for u, uid, act in tag_rows:
        tag_map[_match_key(u)] = (u, uid, act)
    for u, uid in team_rows:
        k = _match_key(u)
        if k not in tag_map:
            tag_map[k] = (u, uid, 0)
    for u, uid, act in sorted(tag_map.values(), key=lambda x: x[0].lower()):
        cfg_rows.append(("member_tag", u, f"active={act};user_id={uid if uid is not None else ''}"))
    for u in get_summary_excluded():
        cfg_rows.append(("summary_excluded", u, ""))
    with sqlite3.connect(DB_PATH) as conn:
        for uid, uname in conn.execute(
            "SELECT user_id, username FROM known_group_admins ORDER BY username ASC"
        ).fetchall():
            cfg_rows.append(("group_admin", str(uname or ""), str(uid if uid is not None else "")))
    _style_sheet(ws_cfg, ["Key", "Value", "Extra"], cfg_rows)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output, len(period_rows)


def restore_from_excel_bytes(excel_bytes: bytes) -> dict[str, object]:
    wb = load_workbook(filename=BytesIO(excel_bytes), read_only=True)
    sheet_name = "All History" if "All History" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]

    reports_restored = 0
    members_restored = 0
    tags_restored = 0
    restored_work_gid: Optional[int] = None
    restored_excel_gid: Optional[int] = None
    now_iso = datetime.now(IST).isoformat()
    has_joining_col = False
    has_penalty_col = False

    with sqlite3.connect(DB_PATH) as conn:
        for idx, row in enumerate(ws.iter_rows(values_only=True)):
            if idx == 0:
                headers = [str(c or "").strip().lower() for c in (row or [])]
                has_joining_col = len(headers) > 1 and headers[1] == "joining"
                has_penalty_col = any("penalty" in h for h in headers)
                continue
            if not row or not row[0]:
                continue

            username = str(row[0]).strip()
            if has_joining_col:
                joining = str(row[1] or "").strip()
                sub_time = str(row[2] or "").strip()
                slot_label = str(row[3] or "").strip()
                rep_date = str(row[4] or "").strip()
                status = str(row[5] or "On Time").strip()
                if has_penalty_col:
                    pen_raw = str(row[6] or "0").replace("₹", "").strip()
                    penalty_rs = _safe_int(pen_raw)
                    caption = str(row[7] or "").strip() if len(row) > 7 else ""
                else:
                    penalty_rs = 0
                    caption = str(row[6] or "").strip() if len(row) > 6 else ""
            else:
                joining = ""
                sub_time = str(row[1] or "").strip()
                slot_label = str(row[2] or "").strip()
                rep_date = str(row[3] or "").strip()
                status = str(row[4] or "On Time").strip()
                penalty_rs = 0
                caption = str(row[5] or "").strip() if len(row) > 5 else ""
                if not joining and caption:
                    _, joining = parse_report_fields(caption, username)

            if sub_time and slot_label:
                status, penalty_rs = evaluate_slot_status_from_strings_with_penalty(sub_time, slot_label)

            exists = conn.execute(
                """
                SELECT id, COALESCE(joining, '') FROM reports
                WHERE LOWER(LTRIM(username, '@')) = ? AND report_date = ? AND reporting_slot = ?
                LIMIT 1
                """,
                (_match_key(username), rep_date, slot_label),
            ).fetchone()
            if not exists:
                conn.execute(
                    """
                    INSERT INTO reports (
                        user_id, username, joining, submission_time, reporting_slot, report_date, status, penalty_rs, caption, created_at_utc
                    ) VALUES (NULL, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (username, joining, sub_time, slot_label, rep_date, status, penalty_rs, caption, now_iso),
                )
                reports_restored += 1
            elif exists[1] == "" and joining != "":
                conn.execute("UPDATE reports SET joining = ? WHERE id = ?", (joining, exists[0]))

            cur_m = conn.execute(
                "INSERT OR IGNORE INTO team_members (username, user_id, added_at) VALUES (?, NULL, ?)",
                (username, now_iso),
            )
            members_restored += cur_m.rowcount
            cur_t = conn.execute(
                "INSERT OR IGNORE INTO assistant_tags (username, user_id, is_active, added_at) VALUES (?, NULL, 1, ?)",
                (username, now_iso),
            )
            tags_restored += cur_t.rowcount

        if "Bot Config" in wb.sheetnames:
            ws_cfg = wb["Bot Config"]
            for idx, row in enumerate(ws_cfg.iter_rows(values_only=True)):
                if idx == 0 or not row or not row[0]:
                    continue
                k = str(row[0]).strip()
                v = str(row[1] or "").strip()
                extra = str(row[2] or "").strip() if len(row) > 2 else ""

                if k == "work_group_id" and v:
                    parsed = _parse_optional_int(v, "work_group_id")
                    if parsed is not None:
                        restored_work_gid = parsed
                        conn.execute(
                            "INSERT INTO bot_settings (key, value) VALUES ('work_group_id', ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                            (str(parsed),),
                        )
                elif k == "excel_group_id" and v:
                    parsed = _parse_optional_int(v, "excel_group_id")
                    if parsed is not None:
                        restored_excel_gid = parsed
                        conn.execute(
                            "INSERT INTO bot_settings (key, value) VALUES ('excel_group_id', ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                            (str(parsed),),
                        )
                elif k == "reminders_enabled" and v in ("true", "false"):
                    conn.execute(
                        "INSERT INTO bot_settings (key, value) VALUES ('reminders_enabled', ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                        (v,),
                    )
                elif k == "member_tag" and v:
                    active_flag = 1
                    uid_val: Optional[int] = None
                    for part in extra.split(";"):
                        if part.startswith("active="):
                            active_flag = 1 if part.split("=", 1)[1].strip() == "1" else 0
                        elif part.startswith("user_id="):
                            uid_val = _parse_optional_int(part.split("=", 1)[1], "user_id")
                    cur_m = conn.execute(
                        "INSERT OR IGNORE INTO team_members (username, user_id, added_at) VALUES (?, ?, ?)",
                        (v, uid_val, now_iso),
                    )
                    members_restored += cur_m.rowcount
                    if uid_val is not None:
                        conn.execute(
                            "UPDATE team_members SET user_id = ? WHERE LOWER(username) = LOWER(?) AND user_id IS NULL",
                            (uid_val, v),
                        )
                    cur_t = conn.execute(
                        """
                        INSERT INTO assistant_tags (username, user_id, is_active, added_at) VALUES (?, ?, ?, ?)
                        ON CONFLICT(username) DO UPDATE SET
                            user_id = COALESCE(excluded.user_id, assistant_tags.user_id),
                            is_active = excluded.is_active
                        """,
                        (v, uid_val, active_flag, now_iso),
                    )
                    tags_restored += cur_t.rowcount
                elif k == "summary_excluded" and v:
                    conn.execute(
                        "INSERT OR IGNORE INTO summary_excluded (username, added_at) VALUES (?, ?)",
                        (v, now_iso),
                    )
                elif k == "group_admin" and (v or extra):
                    uid_val = _parse_optional_int(extra, "group_admin_id")
                    akey = f"id:{uid_val}" if uid_val is not None else f"u:{_match_key(v)}"
                    conn.execute(
                        "INSERT OR REPLACE INTO known_group_admins (admin_key, user_id, username, updated_at) VALUES (?, ?, ?, ?)",
                        (akey, uid_val, v or f"Admin({uid_val})", now_iso),
                    )

        deduplicate_case_insensitive_records(conn)
        recalculate_report_statuses(conn)
        conn.commit()

    return {
        "source": "excel",
        "reports": reports_restored,
        "members": members_restored,
        "tags": tags_restored,
        "work_group_id": restored_work_gid,
        "excel_group_id": restored_excel_gid,
    }


def restore_from_sqlite_bytes(db_bytes: bytes) -> dict[str, object]:
    """Recover reports, old work_group_id, excel_group_id, team_members & tags from an old reports.db file."""
    tmp_path = ""
    reports_restored = 0
    members_restored = 0
    tags_restored = 0
    restored_work_gid: Optional[int] = None
    restored_excel_gid: Optional[int] = None
    now_iso = datetime.now(IST).isoformat()

    try:
        with tempfile.NamedTemporaryFile(suffix=".db", delete=False) as tmp:
            tmp.write(db_bytes)
            tmp_path = tmp.name

        with sqlite3.connect(tmp_path) as src, sqlite3.connect(DB_PATH) as dst:
            tables = {
                r[0] for r in src.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
            }

            if "bot_settings" in tables:
                for key, val in src.execute("SELECT key, value FROM bot_settings").fetchall():
                    if key == "work_group_id" and val:
                        parsed = _parse_optional_int(str(val), "work_group_id")
                        if parsed is not None:
                            restored_work_gid = parsed
                            dst.execute(
                                "INSERT INTO bot_settings (key, value) VALUES ('work_group_id', ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                                (str(parsed),),
                            )
                    elif key in ("excel_group_id", "db_group_id") and val:
                        parsed = _parse_optional_int(str(val), "excel_group_id")
                        if parsed is not None:
                            restored_excel_gid = parsed
                            dst.execute(
                                "INSERT INTO bot_settings (key, value) VALUES ('excel_group_id', ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                                (str(parsed),),
                            )
                    elif key in (
                        "reminders_enabled",
                        "clean_chat_mode",
                        "holiday_mode",
                        "off_days",
                        "daily_joining_target",
                        "per_joining_rate",
                        "penalty_late_regular",
                        "penalty_late_final",
                        "penalty_missed",
                        "dm_started_users",
                        "dm_started_profiles",
                        "fine_reset_date",
                        "fine_reset_time_secs",
                        "member_fine_reset_dates",
                    ) and val is not None:
                        dst.execute(
                            "INSERT INTO bot_settings (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                            (str(key), str(val)),
                        )

            if "team_members" in tables:
                tm_cols = {r[1] for r in src.execute("PRAGMA table_info(team_members)").fetchall()}
                uid_expr = "user_id" if "user_id" in tm_cols else "NULL"
                at_expr = "added_at" if "added_at" in tm_cols else f"'{now_iso}'"
                for uname, uid, added_at in src.execute(
                    f"SELECT username, {uid_expr}, {at_expr} FROM team_members"
                ).fetchall():
                    clean_u = str(uname or "").strip()
                    if not clean_u:
                        continue
                    cur_m = dst.execute(
                        "INSERT OR IGNORE INTO team_members (username, user_id, added_at) VALUES (?, ?, ?)",
                        (clean_u, uid, str(added_at or now_iso)),
                    )
                    members_restored += cur_m.rowcount

            if "assistant_tags" in tables:
                at_cols = {r[1] for r in src.execute("PRAGMA table_info(assistant_tags)").fetchall()}
                uid_expr = "user_id" if "user_id" in at_cols else "NULL"
                act_expr = "is_active" if "is_active" in at_cols else "1"
                time_expr = "added_at" if "added_at" in at_cols else f"'{now_iso}'"
                for uname, uid, is_act, added_at in src.execute(
                    f"SELECT username, {uid_expr}, {act_expr}, {time_expr} FROM assistant_tags"
                ).fetchall():
                    clean_u = str(uname or "").strip()
                    if not clean_u:
                        continue
                    active_val = int(is_act) if is_act is not None else 1
                    cur_t = dst.execute(
                        """
                        INSERT INTO assistant_tags (username, user_id, is_active, added_at) VALUES (?, ?, ?, ?)
                        ON CONFLICT(username) DO UPDATE SET
                            user_id = COALESCE(excluded.user_id, assistant_tags.user_id),
                            is_active = excluded.is_active
                        """,
                        (clean_u, uid, active_val, str(added_at or now_iso)),
                    )
                    tags_restored += cur_t.rowcount

            if "summary_excluded" in tables:
                for (uname, added_at) in src.execute(
                    "SELECT username, added_at FROM summary_excluded"
                ).fetchall():
                    clean_u = str(uname or "").strip()
                    if clean_u:
                        dst.execute(
                            "INSERT OR IGNORE INTO summary_excluded (username, added_at) VALUES (?, ?)",
                            (clean_u, str(added_at or now_iso)),
                        )

            if "known_group_admins" in tables:
                for akey, uid, uname, upd in src.execute(
                    "SELECT admin_key, user_id, username, updated_at FROM known_group_admins"
                ).fetchall():
                    dst.execute(
                        "INSERT OR REPLACE INTO known_group_admins (admin_key, user_id, username, updated_at) VALUES (?, ?, ?, ?)",
                        (str(akey), uid, str(uname or ""), str(upd or now_iso)),
                    )

            if "reports" in tables:
                rep_cols = {r[1] for r in src.execute("PRAGMA table_info(reports)").fetchall()}
                uid_expr = "user_id" if "user_id" in rep_cols else "NULL"
                join_expr = "COALESCE(joining, '')" if "joining" in rep_cols else "''"
                pen_expr = "COALESCE(penalty_rs, 0)" if "penalty_rs" in rep_cols else "0"
                cap_expr = "COALESCE(caption, '')" if "caption" in rep_cols else "''"
                created_expr = "created_at_utc" if "created_at_utc" in rep_cols else f"'{now_iso}'"
                rows = src.execute(
                    f"""
                    SELECT {uid_expr}, username, {join_expr}, submission_time, reporting_slot, report_date, status, {pen_expr}, {cap_expr}, {created_expr}
                    FROM reports
                    ORDER BY id ASC
                    """
                ).fetchall()
                for uid, uname, joining, sub_time, slot_lbl, rep_date, status, pen_rs, caption, created_utc in rows:
                    u_str = str(uname).strip()
                    j_str = str(joining or "").strip()
                    cap_str = str(caption or "").strip()
                    if not j_str and cap_str:
                        _, j_str = parse_report_fields(cap_str, u_str)
                    exists = dst.execute(
                        """
                        SELECT id, COALESCE(joining, '') FROM reports
                        WHERE LOWER(LTRIM(username, '@')) = ? AND report_date = ? AND reporting_slot = ?
                        LIMIT 1
                        """,
                        (_match_key(u_str), str(rep_date), str(slot_lbl)),
                    ).fetchone()
                    if not exists:
                        dst.execute(
                            """
                            INSERT INTO reports (
                                user_id, username, joining, submission_time, reporting_slot, report_date, status, penalty_rs, caption, created_at_utc
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (
                                uid,
                                u_str,
                                j_str,
                                str(sub_time),
                                str(slot_lbl),
                                str(rep_date),
                                str(status or "On Time"),
                                int(pen_rs or 0),
                                cap_str,
                                str(created_utc or now_iso),
                            ),
                        )
                        reports_restored += 1
                    elif exists[1] == "" and j_str != "":
                        dst.execute("UPDATE reports SET joining = ? WHERE id = ?", (j_str, exists[0]))

                    dst.execute(
                        "INSERT OR IGNORE INTO team_members (username, user_id, added_at) VALUES (?, ?, ?)",
                        (u_str, uid, now_iso),
                    )
                    dst.execute(
                        "INSERT OR IGNORE INTO assistant_tags (username, user_id, is_active, added_at) VALUES (?, ?, 1, ?)",
                        (u_str, uid, now_iso),
                    )

            deduplicate_case_insensitive_records(dst)
            recalculate_report_statuses(dst)
            dst.commit()
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass

    return {
        "source": "sqlite",
        "reports": reports_restored,
        "members": members_restored,
        "tags": tags_restored,
        "work_group_id": restored_work_gid,
        "excel_group_id": restored_excel_gid,
    }


def restore_any_backup_bytes(raw_bytes: bytes, filename: str = "") -> dict[str, object]:
    lower_name = (filename or "").lower()
    if raw_bytes.startswith(b"SQLite format 3") or lower_name.endswith(".db") or lower_name.endswith(".sqlite"):
        return restore_from_sqlite_bytes(raw_bytes)
    return restore_from_excel_bytes(raw_bytes)


# ---------------------------------------------------------------------------
# Excel Group Push Helper (Sends ONLY .xlsx Excel File)
# ---------------------------------------------------------------------------
async def push_excel_to_group(
    context: ContextTypes.DEFAULT_TYPE,
    target_date: str,
    caption: str,
) -> None:
    excel_group_id = get_excel_group_id()
    if excel_group_id is None:
        return

    try:
        excel_stream, _count = await asyncio.to_thread(build_excel_bytes, target_date, 1)
        filename = f"team_report_{target_date}.xlsx"

        sent_msg = await context.bot.send_document(
            chat_id=excel_group_id,
            document=excel_stream,
            filename=filename,
            caption=caption,
            disable_notification=True,
        )
        try:
            await context.bot.pin_chat_message(
                chat_id=excel_group_id,
                message_id=sent_msg.message_id,
                disable_notification=True,
            )
        except Exception:
            pass
    except Exception as exc:
        logger.warning("Could not push Excel file to group %s: %s", excel_group_id, exc)


# ---------------------------------------------------------------------------
# Interactive Assistant Tag Menu Builder (/tags) — Safe HTML
# ---------------------------------------------------------------------------
def _build_tags_panel(current_chat_id: Optional[int] = None) -> tuple[str, InlineKeyboardMarkup]:
    members = _dedup_case_insensitive(get_team_members() + get_tagged_assistants())
    tagged = get_tagged_assistants()
    tagged_keys = {_match_key(u) for u in tagged}
    rem_on = is_reminders_enabled()
    work_gid = get_work_group_id()

    tagged_str = ", ".join(html.escape(t) for t in tagged) if tagged else "None selected"
    mismatch_warn = ""
    if current_chat_id is not None and current_chat_id < 0 and work_gid is not None and current_chat_id != work_gid:
        mismatch_warn = (
            f"\n\n<b>Warning:</b> This group ID (<code>{current_chat_id}</code>) differs from "
            f"configured Work Group (<code>{work_gid}</code>). Run <code>/setworkgroup</code> here!"
        )

    text = (
        f"<b>Assistant Tag Manager</b>\n"
        f"Reminders Status: <b>{'ON' if rem_on else 'OFF'}</b>\n"
        f"Currently Tagged ({len(tagged)}): {tagged_str}\n\n"
        f"Assistants are auto-discovered when they send any message in the work group.\n"
        f"<i>Tip: Make the bot Admin in the work group so Telegram sends text messages like 'Hi' to the bot.</i>\n\n"
        f"Tap below to toggle tagging, or use <code>/settags @user1 @user2</code>:"
        f"{mismatch_warn}"
    )

    keyboard: list[list[InlineKeyboardButton]] = []
    row_buf: list[InlineKeyboardButton] = []
    for member in members[:30]:
        is_tagged = _match_key(member) in tagged_keys
        label = f"[ON] {member}" if is_tagged else f"[OFF] {member}"
        cb_val = f"tag_toggle:{member}"
        if len(cb_val.encode("utf-8")) > 64:
            continue
        row_buf.append(InlineKeyboardButton(label, callback_data=cb_val))
        if len(row_buf) == 2:
            keyboard.append(row_buf)
            row_buf = []
    if row_buf:
        keyboard.append(row_buf)

    keyboard.append(
        [
            InlineKeyboardButton(
                f"Reminders: {'ON' if rem_on else 'OFF'} (Tap to Toggle)",
                callback_data="rem_toggle",
            ),
            InlineKeyboardButton("Clear All Tags", callback_data="tag_clear"),
        ]
    )
    return text, InlineKeyboardMarkup(keyboard)


# ---------------------------------------------------------------------------
# Interactive Summary View & Name Removal Manager (/summary)
# ---------------------------------------------------------------------------
def _get_visible_members_for_rows(
    period_rows: list[tuple[str, str, str, str, str, str, int, str]]
) -> list[str]:
    excluded_keys = {_match_key(u) for u in get_summary_excluded()}
    _admin_ids, admin_keys = get_known_admin_sets()
    hidden_keys = excluded_keys | admin_keys
    return [
        m
        for m in _dedup_case_insensitive(
            get_team_members() + get_tagged_assistants() + [r[0] for r in period_rows]
        )
        if _match_key(m) not in hidden_keys
    ]


def build_payout_excel_bytes(
    target_date: Optional[str] = None,
    days: int = 7,
    rate_per_joining: Optional[int] = None,
) -> tuple[BytesIO, int]:
    now_ist = datetime.now(IST)
    today_str = target_date or now_ist.strftime("%Y-%m-%d")
    start_dt = datetime.strptime(today_str, "%Y-%m-%d") - timedelta(days=days - 1)
    start_str = start_dt.strftime("%Y-%m-%d")
    date_list = _dates_for_period(start_str, today_str)
    all_rows = fetch_reports()
    period_rows = [r for r in all_rows if start_str <= r[4] <= today_str]
    members = _get_visible_members_for_rows(period_rows)
    stats_list = compute_member_stats_for_period(members, period_rows, date_list, now_ist)
    rate = rate_per_joining if rate_per_joining is not None else get_per_joining_rate()

    wb = Workbook()
    ws = wb.active
    ws.title = f"Payout {days}D"
    rows_out: list[tuple[str, ...]] = []
    for st in stats_list:
        gross_earn = int(st["total_joining"]) * rate
        net_due_pen = int(st["balance_due"])
        net_payout = gross_earn - net_due_pen
        rows_out.append(
            (
                str(st["member"]),
                str(st["total_joining"]),
                f"₹{rate}",
                f"₹{gross_earn}",
                f"₹{st['gross_pen']}",
                f"₹{st['paid_rs']}",
                f"₹{st['waived_rs']}",
                f"₹{net_due_pen}",
                f"₹{net_payout}",
                f"{st['compliance']}%",
            )
        )
    _style_sheet(
        ws,
        [
            "Username",
            "Valid Joinings",
            "Rate / Join",
            "Gross Joining Amt",
            "Gross Penalty",
            "Paid",
            "Waived",
            "Balance Fine Due",
            "Net Payout",
            "Compliance %",
        ],
        rows_out,
    )
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out, len(stats_list)


def _build_daily_penalty_text(today_str: str) -> str:
    rows = fetch_reports(today_str)
    is_off, off_reason = is_date_holiday_or_off(today_str)
    if is_off:
        return f"🌴 <b>{html.escape(off_reason)} ({today_str})</b>\nNo missed penalties applied today!"
    members = _get_visible_members_for_rows(rows)
    stats = compute_member_stats_for_period(members, rows, [today_str])
    pen_lines: list[str] = []
    total_due = 0
    for st in stats:
        due = int(st["balance_due"])
        if int(st["gross_pen"]) <= 0:
            continue
        total_due += due
        parts: list[str] = []
        if int(st["missed"]) > 0:
            parts.append(f"{st['missed']} Missed (₹{st['missed_pen']})")
        late_cnt = int(st["late_regular"]) + int(st["late_final"])
        if late_cnt > 0:
            parts.append(f"{late_cnt} Late (₹{st['late_pen']})")
        adj = int(st["paid_rs"]) + int(st["waived_rs"])
        adj_note = f" [Paid/Waived: ₹{adj}]" if adj > 0 else ""
        pen_lines.append(
            f"- {html.escape(str(st['member']))} — {' + '.join(parts)}{adj_note} = <b>₹{due}</b>"
        )
    if not pen_lines:
        return f"🎉 <b>Penalty Report ({today_str} IST)</b>\nZero penalties today! Everyone is On Time! 👏"
    return (
        f"🚨 <b>Daily Penalty &amp; Missed Summary ({today_str} IST)</b>\n"
        f"Total Net Penalty Due: <b>₹{total_due}</b>\n\n"
        + "\n".join(pen_lines)
        + "\n\n<i>Admins can use /paid @user amt or /waive @user amt</i>"
    )


def _build_leaderboard_text(today_str: str) -> str:
    rows = fetch_reports(today_str)
    members = _get_visible_members_for_rows(rows)
    stats = compute_member_stats_for_period(members, rows, [today_str])
    ranked = sorted(
        stats,
        key=lambda s: (int(s["total_joining"]), int(s["on_time"]), -int(s["gross_pen"])),
        reverse=True,
    )
    target_val = get_daily_joining_target()
    medals = ("🥇", "🥈", "🥉")
    lines = [f"🏆 <b>Daily Leaderboard ({today_str} IST)</b>"]
    if not ranked:
        lines.append("No active assistants found today.")
        return "\n".join(lines)
    total_team_join = sum(int(s["total_joining"]) for s in ranked)
    lines.append(f"Team Valid Total Joining: <b>{total_team_join}</b>\n")
    for idx, st in enumerate(ranked[:15]):
        badge = medals[idx] if idx < 3 else f"#{idx + 1}"
        tgt_note = ""
        if target_val > 0 and int(st["leave_days"]) == 0:
            diff = int(st["target_diff"])
            tgt_note = f" | 🎯 {'+' if diff >= 0 else ''}{diff}"
        star = " ⭐100%" if int(st["compliance"]) == 100 and int(st["on_time"]) > 0 else ""
        lines.append(
            f"{badge} {html.escape(str(st['member']))}: <b>{st['total_joining']}</b> Join | {st['on_time']}/{max(len(st['allowed_slots']), 1)} On-Time{tgt_note}{star}"
        )
    return "\n".join(lines)


def _build_summary_view(today_str: str) -> tuple[str, InlineKeyboardMarkup]:
    rows = fetch_reports(today_str)
    _admin_ids, admin_keys = get_known_admin_sets()
    visible_rows = [r for r in rows if _match_key(r[0]) not in admin_keys]
    all_names = _get_visible_members_for_rows(visible_rows)
    stats = compute_member_stats_for_period(all_names, visible_rows, [today_str])
    target_val = get_daily_joining_target()
    is_off, off_reason = is_date_holiday_or_off(today_str)

    total_joining_team = sum(int(st["total_joining"]) for st in stats)
    total_late_pen = sum(int(st["late_pen"]) for st in stats)
    total_missed_pen = sum(int(st["missed_pen"]) for st in stats)
    total_due = sum(int(st["balance_due"]) for st in stats)

    per_user_lines: list[str] = []
    for st in stats:
        late_cnt = int(st["late_regular"]) + int(st["late_final"])
        late_note = f" ({late_cnt} Late)" if late_cnt > 0 else ""
        miss_note = f" | Missed: {st['missed']} (₹{st['missed_pen']})" if int(st["missed"]) > 0 else ""
        leave_note = " | <b>ON LEAVE</b>" if int(st["leave_days"]) > 0 else ""
        tgt_note = ""
        if target_val > 0 and int(st["leave_days"]) == 0:
            diff = int(st["target_diff"])
            tgt_note = f" | 🎯 {'+' if diff >= 0 else ''}{diff}"
        due = int(st["balance_due"])
        gross = int(st["gross_pen"])
        if due > 0:
            pen_note = f" | Due: <b>₹{due}</b>"
        elif gross > 0:
            pen_note = " | <b>Paid/Waived ₹0</b>"
        else:
            pen_note = ""
        slot_cap = len(st["allowed_slots"])
        per_user_lines.append(
            f"- {html.escape(str(st['member']))}: {st['submitted_count']}/{slot_cap} slots{late_note} | Joining: <b>{st['total_joining']}</b>{tgt_note}{leave_note}{miss_note}{pen_note}"
        )

    off_banner = f"\n🌴 <b>{html.escape(off_reason)}</b> (Missed penalties paused)" if is_off else ""
    lines = [
        f"<b>Daily Summary ({today_str} IST)</b>{off_banner}",
        f"Valid Team Joining: <b>{total_joining_team}</b> | Slot Logs: <b>{len(visible_rows)}</b>",
        f"Net Fine Due: <b>₹{total_due}</b> (Gross: ₹{total_late_pen + total_missed_pen} | Late: ₹{total_late_pen} + Missed: ₹{total_missed_pen})",
        "",
    ]
    if not per_user_lines:
        lines.append("No active assistants shown in summary.")
    else:
        lines.extend(per_user_lines)

    markup = InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(
                    "Remove / Manage Assistant Names",
                    callback_data="sum_manage",
                )
            ]
        ]
    )
    return "\n".join(lines), markup


def _build_summary_manage_panel(today_str: str) -> tuple[str, InlineKeyboardMarkup]:
    rows = fetch_reports(today_str)
    _admin_ids, admin_keys = get_known_admin_sets()
    candidates = [
        m
        for m in _dedup_case_insensitive(
            get_team_members()
            + get_tagged_assistants()
            + [r[0] for r in rows]
            + get_summary_excluded()
        )
        if _match_key(m) not in admin_keys
    ]
    excluded_keys = {_match_key(u) for u in get_summary_excluded()}

    lines = [
        f"<b>Summary Format Name Manager ({today_str})</b>",
        "Tap any assistant below to <b>Remove</b> or <b>Restore</b> their name from the Summary &amp; Excel Summary sheets:",
        "",
        "<code>[SHOWN]</code> = Visible in Summary",
        "<code>[REMOVED]</code> = Hidden from Summary",
        "",
        "Direct commands: <code>/removesummary @user</code> | <code>/restoresummary @user</code>",
    ]

    keyboard: list[list[InlineKeyboardButton]] = []
    row_buf: list[InlineKeyboardButton] = []
    for member in candidates[:30]:
        is_hidden = _match_key(member) in excluded_keys
        label = f"[REMOVED] {member}" if is_hidden else f"[SHOWN] {member}"
        cb_val = f"sum_toggle:{member}"
        if len(cb_val.encode("utf-8")) > 64:
            continue
        row_buf.append(InlineKeyboardButton(label, callback_data=cb_val))
        if len(row_buf) == 2:
            keyboard.append(row_buf)
            row_buf = []
    if row_buf:
        keyboard.append(row_buf)

    keyboard.append(
        [InlineKeyboardButton("Back to Summary Format", callback_data="sum_view")]
    )
    return "\n".join(lines), InlineKeyboardMarkup(keyboard)


# ---------------------------------------------------------------------------
# Authorization Helper
# ---------------------------------------------------------------------------
def _is_admin(update: Update) -> bool:
    if not update.effective_user:
        return False
    if DEFAULT_ADMIN_CHAT_ID is not None:
        return update.effective_user.id == DEFAULT_ADMIN_CHAT_ID
    uname = f"@{update.effective_user.username}" if update.effective_user.username else ""
    return is_known_group_admin(uname, update.effective_user.id)


async def _require_admin(update: Update) -> bool:
    if _is_admin(update):
        return True
    if update.message:
        await update.message.reply_text(
            "⛔ <b>Restricted Command</b>\n"
            "Ye command sirf <b>Admin</b> ke liye hai.\n\n"
            "Assistants ke liye alag commands hain: /myhelp",
            parse_mode="HTML",
        )
    return False


# ---------------------------------------------------------------------------
# Telegram Command & Callback Handlers
# ---------------------------------------------------------------------------
def mark_user_dm_started(user_id: Optional[int], username: str = "") -> None:
    if user_id is None:
        return
    uid_int = int(user_id)
    raw = get_setting("dm_started_users", "[]")
    try:
        arr = json.loads(raw or "[]")
        if not isinstance(arr, list):
            arr = []
    except Exception:
        arr = []
    if uid_int not in arr:
        arr.append(uid_int)
        set_setting("dm_started_users", json.dumps(arr))

    if username:
        raw_p = get_setting("dm_started_profiles", "{}")
        try:
            prof = json.loads(raw_p or "{}")
            if not isinstance(prof, dict):
                prof = {}
        except Exception:
            prof = {}
        clean_u = _normalize_tag(username) or username.strip()
        if prof.get(str(uid_int)) != clean_u:
            prof[str(uid_int)] = clean_u
            set_setting("dm_started_profiles", json.dumps(prof))


def get_dm_started_user_ids() -> set[int]:
    raw = get_setting("dm_started_users", "[]")
    try:
        arr = json.loads(raw or "[]")
        return {int(x) for x in arr if str(x).lstrip("-").isdigit()} if isinstance(arr, list) else set()
    except Exception:
        return set()


def get_dm_started_profiles() -> dict[int, str]:
    raw_p = get_setting("dm_started_profiles", "{}")
    out: dict[int, str] = {}
    try:
        prof = json.loads(raw_p or "{}")
        if isinstance(prof, dict):
            for k, v in prof.items():
                if str(k).lstrip("-").isdigit():
                    out[int(k)] = str(v)
    except Exception:
        pass
    for uid in get_dm_started_user_ids():
        out.setdefault(uid, f"User {uid}")
    # Also enrich display usernames from assistant_tags / team_members if available
    with sqlite3.connect(DB_PATH) as conn:
        for uid, uname in conn.execute(
            "SELECT user_id, username FROM assistant_tags WHERE user_id IS NOT NULL UNION SELECT user_id, username FROM team_members WHERE user_id IS NOT NULL"
        ).fetchall():
            if uid is not None and int(uid) in out and uname:
                out[int(uid)] = str(uname)
    return out


def _assistant_commands_text(uname: str, dm_activated: bool = False) -> str:
    header = (
        f"✅ <b>Personal DM Alerts Activated!</b>\n\n"
        f"Hi <b>{html.escape(uname)}</b>, ab aapko slot deadlines aur 15-min countdown ke urgent alerts direct DM mein milenge.\n\n"
        if dm_activated
        else f"👋 <b>Assistant Command Guide ({html.escape(uname)})</b>\n\n"
    )
    return (
        f"{header}"
        f"<b>Assistants Ke Liye Commands:</b>\n"
        f"/start — Personal DM Alerts activate karein\n"
        f"/mystatus — Aaj ki valid joining, slots &amp; aaj ki penalty dekhein\n"
        f"/myfine — Apna fine balance &amp; payment history dekhein\n"
        f"/myrank — Aaj ki apni rank aur on-time score dekhein\n"
        f"/myhelp — Ye assistant guide wapas dekhein\n\n"
        f"<b>Report Bhejne Ka Format (Work Group Mein):</b>\n"
        f"<code>Username : @AapkaUsername</code>\n"
        f"<code>Joining : 45</code> (ya <code>Total Joining 45</code>)"
    )


async def myhelp_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.effective_user:
        return
    user = update.effective_user
    uname = f"@{user.username}" if user.username else (user.first_name or "Assistant")
    if update.effective_chat and update.effective_chat.type == "private":
        await asyncio.to_thread(mark_user_dm_started, user.id, uname)
    await update.message.reply_text(_assistant_commands_text(uname, dm_activated=False), parse_mode="HTML")


async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.effective_chat or not update.effective_user:
        return

    user = update.effective_user
    uname = f"@{user.username}" if user.username else (user.first_name or "Assistant")
    if update.effective_chat.type == "private":
        await asyncio.to_thread(mark_user_dm_started, user.id, uname)

    if not _is_admin(update):
        if not is_known_group_admin(uname, user.id):
            await asyncio.to_thread(auto_discover_assistant, uname, user.id)
        dm_on = update.effective_chat.type == "private"
        await update.message.reply_text(_assistant_commands_text(uname, dm_activated=dm_on), parse_mode="HTML")
        return

    rem_status = "ON" if is_reminders_enabled() else "OFF"
    clean_status = "ON" if is_clean_chat_enabled() else "OFF"
    dm_started_count = len(get_dm_started_user_ids())
    work_gids = get_all_work_group_ids()
    work_gid_str = ", ".join(str(g) for g in work_gids) if work_gids else "Not set"
    excel_gid = get_excel_group_id() or "Not set"
    tagged_count = len(get_tagged_assistants())
    reg_p, fin_p, miss_p = get_penalty_config()

    await update.message.reply_text(
        f"<b>Excel Report Tracker Bot (Admin Panel)</b>\n"
        f"Your User ID (<code>ADMIN_CHAT_ID</code>): <code>{update.effective_user.id}</code>\n"
        f"Current Chat ID: <code>{update.effective_chat.id}</code>\n"
        f"Work Group(s): <code>{work_gid_str}</code> | Excel Group: <code>{excel_gid}</code>\n"
        f"Reminders: <b>{rem_status}</b> | Clean Chat: <b>{clean_status}</b> | Tagged: <b>{tagged_count}</b>\n"
        f"Assistant DM Mode: <b>Bot API (/start Only | {dm_started_count} active, Admins Excluded)</b>\n"
        f"Penalties: Regular ₹{reg_p} | Final ₹{fin_p} | Missed ₹{miss_p}\n\n"
        f"<b>All-Team Reports &amp; Fines (Admin Only):</b>\n"
        f"/allfines [days|today] — Ek saath sabhi assistants ki penalty\n"
        f"/paidall (ya /paid all) — Sabhi ke pending fines Paid mark karke ₹0 karein\n"
        f"/paid @user &lt;amt|all&gt; | /waive all | /waive @user &lt;amt|all&gt;\n"
        f"/resetfine [@user|all] — Sabhi/ek assistant ke fines ₹0 reset karein\n"
        f"/excel [YYYY-MM-DD] | /weekly | /monthly [YYYY-MM] | /payout [days] [rate]\n"
        f"/summary | /leaderboard | /find @user [days] | /day [YYYY-MM-DD]\n\n"
        f"<b>Assistant DM Alerts (/start Only — Admins Excluded):</b>\n"
        f"/dmstatus | /dmtest @assistant | /dmall [message]\n\n"
        f"<b>Admin Controls &amp; Backfill:</b>\n"
        f"/add (reply) | /addreport @user 40 [10am]\n"
        f"/tags | /settags @a @b | /addtag | /removetag | /cleartags\n"
        f"/reminder on|off | /cleanchat on|off | /countdown [slot|stop]\n"
        f"/leave @user [date] [slot] | /unleave @user [date] | /setshift @user morning|evening|full\n"
        f"/settarget N | /setrate Rs | /setpenalty regular 5 final 10 missed 15\n"
        f"/holiday on|off | /offdays sun|sat,sun|none | /alias @old @new | /aliases\n"
        f"/removesummary @a | /restoresummary @a | /restore\n"
        f"/setworkgroup | /addworkgroup | /removeworkgroup | /groups | /setexcelgroup | /check\n\n"
        f"<i>Assistant Commands: /start, /mystatus, /myfine, /myrank, /myhelp</i>",
        parse_mode="HTML",
    )


async def check_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.effective_chat or not update.effective_user:
        return
    if not await _require_admin(update):
        return
    chat = update.effective_chat
    work_gids = get_all_work_group_ids()
    work_gid = get_work_group_id()
    excel_gid = get_excel_group_id()
    ist_now = datetime.now(IST)
    active_slot = detect_active_slot(ist_now)
    resolved_slot, resolved_status, resolved_pen = resolve_submission_slot_and_penalty(ist_now)
    matches_work = is_allowed_work_group(chat.id)
    same_as_excel = excel_gid is not None and chat.id == excel_gid and not is_allowed_work_group(chat.id)
    total_reports = await asyncio.to_thread(count_total_reports)
    total_members = len(await asyncio.to_thread(get_team_members))
    is_off, off_reason = is_date_holiday_or_off(ist_now.strftime("%Y-%m-%d"))

    pinned_backup_status = "No Excel/Database Group configured"
    if excel_gid is not None:
        try:
            db_chat = await context.bot.get_chat(excel_gid)
            pinned = db_chat.pinned_message
            if pinned and pinned.document:
                pinned_backup_status = f"Pinned: {pinned.document.file_name or 'document'}"
            else:
                pinned_backup_status = "No pinned .xlsx/.db file in Database Group"
        except Exception as exc:
            pinned_backup_status = f"Cannot access Database Group ({exc})"

    slot_display = active_slot or (f"{resolved_slot} [{resolved_status}]" if resolved_slot else "None (Before 10:00 AM)")
    work_disp = ", ".join(str(g) for g in work_gids) if work_gids else "Any Group"
    lines = [
        "<b>Bot Diagnostic Check</b>",
        f"Current Chat ID: <code>{chat.id}</code> ({html.escape(chat.type)})",
        f"Configured Work Group(s): <code>{work_disp}</code>",
        f"Database/Excel Group (<code>EXCEL_GROUP_ID</code>): <code>{excel_gid or 'Not set'}</code>",
        f"Saved Data: <b>{total_reports}</b> reports | <b>{total_members}</b> assistants",
        f"Database Backup: {html.escape(pinned_backup_status)}",
        f"Current IST Time: <b>{ist_now.strftime('%Y-%m-%d %I:%M:%S %p IST')}</b>",
        f"Active Slot / Penalty Now: <b>{html.escape(slot_display)}</b> (₹{resolved_pen})",
        f"Holiday / Off Status: <b>{html.escape(off_reason) if is_off else 'Working Day'}</b>",
        f"Report Recording Here: <b>{'YES' if (matches_work and not same_as_excel) else 'NO (Group Mismatch)'}</b>",
    ]
    if not matches_work:
        lines.append(
            f"\n<b>Fix:</b> Send <code>/setworkgroup</code> or <code>/addworkgroup</code> here."
        )
    elif same_as_excel:
        lines.append(
            "\n<b>Fix:</b> This chat is only set as Database/Excel Group. Run <code>/setworkgroup</code> here if this is your Work Group."
        )
    await update.message.reply_text("\n".join(lines), parse_mode="HTML")


async def add_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Reply /add [optional slot e.g. 10am] to ANY already-sent message to record Username, Joining & Timing."""
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can add past reports.")
        return

    replied = update.message.reply_to_message
    if not replied:
        await update.message.reply_text(
            "<b>How to add reports already sent earlier:</b>\n"
            "1. Reply <code>/add</code> directly to any message/caption in the group (keeps original IST timing!).\n"
            "2. Or use <code>/addreport @username 40 10am</code> (supports multiple lines too!).",
            parse_mode="HTML",
        )
        return

    raw_text = (replied.caption or replied.text or "").strip()
    sender = replied.from_user
    fallback_user = (
        f"@{sender.username}"
        if sender and sender.username
        else (sender.first_name if sender and sender.first_name else "Unknown User")
    )
    sender_id = sender.id if sender else None

    username, joining = parse_report_fields(raw_text, fallback_user)
    effective_uid = sender_id if username.lower() == fallback_user.lower() else None

    orig_dt = replied.date
    if getattr(replied, "forward_origin", None) and getattr(replied.forward_origin, "date", None):
        orig_dt = replied.forward_origin.date

    ist_dt = orig_dt.astimezone(IST)
    submission_time = ist_dt.strftime("%I:%M:%S %p IST")
    report_date = ist_dt.strftime("%Y-%m-%d")

    slot_label = None
    if context.args:
        slot_label = parse_slot_alias(context.args[0])
    if slot_label is None:
        slot_label = infer_slot_from_time(ist_dt)

    status, penalty_rs = evaluate_slot_status_and_penalty(ist_dt, slot_label)

    action = await asyncio.to_thread(
        upsert_report_to_db,
        effective_uid,
        username,
        joining,
        submission_time,
        slot_label,
        report_date,
        status,
        raw_text,
        orig_dt.isoformat(),
        True,
        penalty_rs,
    )

    excel_caption = (
        f"Excel Updated ({action}) | {username} | Joining: {joining or 'N/A'} | {slot_label}\n"
        f"Time: {submission_time} | Date: {report_date} | Status: {status} | Penalty: ₹{penalty_rs}"
    )
    await push_excel_to_group(context, target_date=report_date, caption=excel_caption)

    await update.message.reply_text(
        f"<b>Added ({html.escape(action)}) from message!</b>\n"
        f"Username: <code>{html.escape(username)}</code>\n"
        f"Joining: <b>{html.escape(joining or 'N/A')}</b>\n"
        f"Timing: {html.escape(submission_time)} | Date: {report_date}\n"
        f"Slot: {html.escape(slot_label)} | Status: <b>{html.escape(status)}</b> (Penalty: ₹{penalty_rs})",
        parse_mode="HTML",
        reply_to_message_id=replied.message_id,
    )


async def addreport_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Manually add or bulk-paste already-submitted reports: /addreport @user 40 [10am] [HH:MM]."""
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can add reports.")
        return

    full_text = (update.message.text or "").strip()
    lines = [ln.strip() for ln in full_text.splitlines() if ln.strip()]
    if lines:
        first = re.sub(r"^/addreport(?:@\w+)?", "", lines[0], flags=re.IGNORECASE).strip()
        lines = ([first] if first else []) + lines[1:]

    if not lines:
        await update.message.reply_text(
            "<b>Manual / Bulk Add Usage:</b>\n"
            "<code>/addreport @Govind7549 40 10am</code>\n"
            "Or paste multiple lines:\n"
            "<code>/addreport\n@Govind7549 40 10am\n@Rohan 25 10am</code>",
            parse_mode="HTML",
        )
        return

    ist_now = datetime.now(IST)
    today_str = ist_now.strftime("%Y-%m-%d")
    added_items: list[str] = []

    # Support pasting raw "Username- @...\nJoining- 40" directly after /addreport
    joined_block = "\n".join(lines)
    if _USERNAME_PATTERN.search(joined_block) and _JOINING_PATTERN.search(joined_block) and not any(
        ln.startswith("@") for ln in lines
    ):
        u_name, j_val = parse_report_fields(joined_block, "Unknown User")
        slot_arg = parse_slot_alias(lines[0]) if lines else None
        slot_label = slot_arg or infer_slot_from_time(ist_now)
        sub_time = ist_now.strftime("%I:%M:%S %p IST")
        status, penalty_rs = evaluate_slot_status_and_penalty(ist_now, slot_label)
        await asyncio.to_thread(
            upsert_report_to_db,
            None,
            u_name,
            j_val,
            sub_time,
            slot_label,
            today_str,
            status,
            joined_block,
            ist_now.isoformat(),
            True,
            penalty_rs,
        )
        added_items.append(f"{u_name} | Joining: {j_val or 'N/A'} | {slot_label} ({status})")
    else:
        for ln in lines:
            parts = ln.split()
            if not parts:
                continue
            raw_user = parts[0]
            norm_user = _normalize_tag(raw_user) or raw_user
            joining = ""
            slot_label = None
            custom_time = None

            for token in parts[1:]:
                if token.isdigit() and joining == "":
                    joining = token
                    continue
                maybe_slot = parse_slot_alias(token)
                if maybe_slot:
                    slot_label = maybe_slot
                    continue
                if re.match(r"^\d{1,2}:\d{2}(?:am|pm)?$", token, flags=re.IGNORECASE):
                    custom_time = token

            target_dt = ist_now
            if custom_time:
                parsed_t = _parse_ist_time_str(custom_time)
                if parsed_t:
                    target_dt = datetime.combine(ist_now.date(), parsed_t, tzinfo=IST)
                    sub_time = target_dt.strftime("%I:%M:%S %p IST")
                    if slot_label is None:
                        slot_label = infer_slot_from_time(target_dt)
                else:
                    sub_time = ist_now.strftime("%I:%M:%S %p IST")
            else:
                sub_time = ist_now.strftime("%I:%M:%S %p IST")

            if slot_label is None:
                slot_label = infer_slot_from_time(target_dt)

            status, penalty_rs = evaluate_slot_status_and_penalty(target_dt, slot_label)

            await asyncio.to_thread(
                upsert_report_to_db,
                None,
                norm_user,
                joining,
                sub_time,
                slot_label,
                today_str,
                status,
                f"Username- {norm_user}\nJoining- {joining}" if joining else norm_user,
                target_dt.isoformat(),
                True,
                penalty_rs,
            )
            added_items.append(
                f"{norm_user} | Joining: {joining or 'N/A'} | {slot_label} ({sub_time} | {status})"
            )

    if not added_items:
        await update.message.reply_text("No valid rows found. Example: <code>/addreport @Govind7549 40 10am</code>", parse_mode="HTML")
        return

    await push_excel_to_group(
        context,
        target_date=today_str,
        caption=f"Excel Updated (Backfill: {len(added_items)} entries) | {today_str}",
    )
    summary_lines = "\n".join(f"- {html.escape(item)}" for item in added_items)
    await update.message.reply_text(
        f"<b>Added {len(added_items)} report(s) to Excel:</b>\n{summary_lines}",
        parse_mode="HTML",
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not _is_admin(update):
        await myhelp_command(update, context)
        return
    await start_command(update, context)


async def tags_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.effective_chat:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can manage assistant tags.")
        return

    sync_cid = update.effective_chat.id if update.effective_chat.id < 0 else get_work_group_id()
    await refresh_group_admins(context.bot, sync_cid, force=True)
    text, markup = await asyncio.to_thread(_build_tags_panel, update.effective_chat.id)
    await update.message.reply_text(text, reply_markup=markup, parse_mode="HTML")


async def settags_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can modify assistant tags.")
        return
    if not context.args:
        await update.message.reply_text(
            "Usage: <code>/settags @assistant1 @assistant2</code>", parse_mode="HTML"
        )
        return

    raw_items = [part for arg in context.args for part in arg.split(",") if part.strip()]
    saved = await asyncio.to_thread(replace_assistant_tags, raw_items)
    await update.message.reply_text(
        f"Assistant tags updated ({len(saved)}): {', '.join(saved) if saved else 'None'}"
    )


async def addtag_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can modify assistant tags.")
        return
    if not context.args:
        await update.message.reply_text(
            "Usage: <code>/addtag @assistant1 @assistant2</code>", parse_mode="HTML"
        )
        return

    raw_items = [part for arg in context.args for part in arg.split(",") if part.strip()]
    for raw in raw_items:
        norm = _normalize_tag(raw)
        if norm:
            await asyncio.to_thread(add_assistant_tag, norm)

    tagged = await asyncio.to_thread(get_tagged_assistants)
    await update.message.reply_text(
        f"Updated assistant tags ({len(tagged)}): {', '.join(tagged) or 'None'}"
    )


async def removetag_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can modify assistant tags.")
        return
    if not context.args:
        await update.message.reply_text("Usage: <code>/removetag @assistant1</code>", parse_mode="HTML")
        return

    removed = []
    for raw in context.args:
        if await asyncio.to_thread(remove_assistant_tag, raw):
            removed.append(raw)

    tagged = await asyncio.to_thread(get_tagged_assistants)
    if removed:
        await update.message.reply_text(
            f"Removed from assistant tags: {', '.join(removed)}\nActive tags: {', '.join(tagged) or 'None'}"
        )
    else:
        await update.message.reply_text("No matching assistant found in active tags.")


async def cleartags_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can clear assistant tags.")
        return
    await asyncio.to_thread(clear_assistant_tags)
    await update.message.reply_text(
        "All assistant tags cleared. Use <code>/tags</code> or <code>/settags @user</code> to add again.",
        parse_mode="HTML",
    )


async def tags_callback_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    if not query or not query.data:
        return

    if DEFAULT_ADMIN_CHAT_ID is not None and query.from_user.id != DEFAULT_ADMIN_CHAT_ID:
        await query.answer("Only the admin can modify tags, reminders, or summary names.", show_alert=True)
        return

    data = query.data
    today_str = datetime.now(IST).strftime("%Y-%m-%d")

    if data.startswith("late_ok:"):
        rid = _safe_int(data.split(":", 1)[1])
        res = await asyncio.to_thread(mark_report_waived_on_time, rid)
        if not res:
            await query.answer("Report record not found.", show_alert=True)
            return
        uname, slot_lbl, rdate = res
        await query.answer(f"Marked {uname} On Time (₹0)!")
        await push_excel_to_group(
            context,
            target_date=rdate,
            caption=f"Excel Updated (Admin Approved ₹0) | {uname} | {slot_lbl} | {rdate}",
        )
        if query.message:
            orig_text = getattr(query.message, "text_html", None) or getattr(query.message, "text", "") or ""
            updated_text = f"{orig_text}\n\n✅ <b>Admin Override: Approved On Time (₹0 Penalty)</b>"
            try:
                await query.edit_message_text(text=updated_text, reply_markup=None, parse_mode="HTML")
            except Exception:
                pass
        return

    if data.startswith("late_keep:"):
        await query.answer("Late penalty confirmed.")
        if query.message:
            orig_text = getattr(query.message, "text_html", None) or getattr(query.message, "text", "") or ""
            try:
                await query.edit_message_text(
                    text=f"{orig_text}\n\n🔒 <i>Late penalty confirmed by Admin</i>",
                    reply_markup=None,
                    parse_mode="HTML",
                )
            except Exception:
                pass
        return

    if data.startswith("late_del:"):
        rid = _safe_int(data.split(":", 1)[1])
        res = await asyncio.to_thread(delete_report_by_id, rid)
        if not res:
            await query.answer("Report already deleted.", show_alert=True)
            return
        uname, slot_lbl, rdate = res
        await query.answer(f"Deleted {slot_lbl} report for {uname}.")
        await push_excel_to_group(
            context,
            target_date=rdate,
            caption=f"Excel Updated (Admin Deleted) | {uname} | {slot_lbl} | {rdate}",
        )
        if query.message:
            try:
                await query.edit_message_text(
                    text=f"🗑 <b>Report Ignored &amp; Deleted by Admin</b>\n({html.escape(uname)} | {html.escape(slot_lbl)})",
                    reply_markup=None,
                    parse_mode="HTML",
                )
            except Exception:
                pass
        return

    if data == "sum_manage":
        await query.answer()
        text, markup = await asyncio.to_thread(_build_summary_manage_panel, today_str)
        try:
            await query.edit_message_text(text=text, reply_markup=markup, parse_mode="HTML")
        except Exception:
            pass
        return
    elif data == "sum_view":
        await query.answer()
        text, markup = await asyncio.to_thread(_build_summary_view, today_str)
        try:
            await query.edit_message_text(text=text, reply_markup=markup, parse_mode="HTML")
        except Exception:
            pass
        return
    elif data.startswith("sum_toggle:"):
        member = data.split("sum_toggle:", 1)[1]
        is_now_hidden = await asyncio.to_thread(toggle_summary_exclusion, member)
        await query.answer(
            f"{member}: {'Removed from Summary' if is_now_hidden else 'Restored to Summary'}"
        )
        text, markup = await asyncio.to_thread(_build_summary_manage_panel, today_str)
        try:
            await query.edit_message_text(text=text, reply_markup=markup, parse_mode="HTML")
        except Exception:
            pass
        return

    if data.startswith("tag_toggle:"):
        member = data.split("tag_toggle:", 1)[1]
        is_now_tagged = await asyncio.to_thread(toggle_assistant_tag, member)
        await query.answer(f"{member}: {'Tagged' if is_now_tagged else 'Untagged'}")
    elif data == "rem_toggle":
        new_state = not is_reminders_enabled()
        await asyncio.to_thread(set_setting, "reminders_enabled", "true" if new_state else "false")
        await query.answer(f"Reminders {'ON' if new_state else 'OFF'}")
    elif data == "tag_clear":
        await asyncio.to_thread(clear_assistant_tags)
        await query.answer("Cleared all assistant tags")
    elif data.startswith("allfines_days:"):
        days_val = max(1, min(90, _safe_int(data.split(":", 1)[1]) or 30))
        await query.answer(f"Showing {days_val}-day team fines")
        text, markup = await asyncio.to_thread(_build_all_fines_view, days_val)
        try:
            await query.edit_message_text(text=text, reply_markup=markup, parse_mode="HTML")
        except Exception:
            pass
        return
    elif data.startswith("allfines_paid_ask:"):
        days_val = max(1, min(90, _safe_int(data.split(":", 1)[1]) or 30))
        await query.answer("Confirm mark all paid?")
        confirm_markup = InlineKeyboardMarkup(
            [
                [
                    InlineKeyboardButton("✅ Confirm Mark All Paid (₹0)", callback_data=f"allfines_paid_do:{days_val}"),
                    InlineKeyboardButton("❌ Cancel", callback_data=f"allfines_days:{days_val}"),
                ]
            ]
        )
        try:
            await query.edit_message_reply_markup(reply_markup=confirm_markup)
        except Exception:
            pass
        return
    elif data.startswith("allfines_paid_do:"):
        days_val = max(1, min(90, _safe_int(data.split(":", 1)[1]) or 30))
        settled, total_cleared = await asyncio.to_thread(
            _bulk_settle_all_fines, "paid", "Bulk paid via /allfines"
        )
        if settled:
            await query.answer(f"Marked ₹{total_cleared} Paid for {len(settled)} assistant(s)!")
        else:
            await query.answer("Sabhi assistants ka balance pehle se hi ₹0 hai!")
        text, markup = await asyncio.to_thread(_build_all_fines_view, days_val)
        try:
            await query.edit_message_text(text=text, reply_markup=markup, parse_mode="HTML")
        except Exception:
            pass
        return
    elif data.startswith("allfines_reset_ask:"):
        days_val = max(1, min(90, _safe_int(data.split(":", 1)[1]) or 30))
        await query.answer("Confirm reset?")
        confirm_markup = InlineKeyboardMarkup(
            [
                [
                    InlineKeyboardButton("✅ Confirm Zero All Fines (₹0)", callback_data=f"allfines_reset_do:{days_val}"),
                    InlineKeyboardButton("❌ Cancel", callback_data=f"allfines_days:{days_val}"),
                ]
            ]
        )
        try:
            await query.edit_message_reply_markup(reply_markup=confirm_markup)
        except Exception:
            pass
        return
    elif data.startswith("allfines_reset_do:"):
        days_val = max(1, min(90, _safe_int(data.split(":", 1)[1]) or 30))
        now_ist = datetime.now(IST)
        today_r = now_ist.strftime("%Y-%m-%d")
        now_secs = _time_to_secs(now_ist.time().replace(tzinfo=None))

        def _do_reset_all() -> None:
            with sqlite3.connect(DB_PATH) as conn:
                conn.execute("UPDATE reports SET penalty_rs = 0 WHERE penalty_rs != 0")
                conn.execute("DELETE FROM penalty_ledger")
                conn.commit()
            set_setting("fine_reset_date", today_r)
            set_setting("fine_reset_time_secs", str(now_secs))
            set_setting("member_fine_reset_dates", "{}")

        await asyncio.to_thread(_do_reset_all)
        await query.answer("All team fines reset to ₹0!")
        text, markup = await asyncio.to_thread(_build_all_fines_view, days_val)
        try:
            await query.edit_message_text(text=text, reply_markup=markup, parse_mode="HTML")
        except Exception:
            pass
        return
    else:
        await query.answer()
        return

    chat_id = query.message.chat.id if query.message else None
    text, markup = await asyncio.to_thread(_build_tags_panel, chat_id)
    try:
        await query.edit_message_text(text=text, reply_markup=markup, parse_mode="HTML")
    except Exception:
        pass


async def reminder_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not await _require_admin(update):
        return

    if not context.args:
        status = "ON" if is_reminders_enabled() else "OFF"
        tagged = await asyncio.to_thread(get_tagged_assistants)
        tagged_escaped = ", ".join(html.escape(t) for t in tagged) if tagged else "None"
        await update.message.reply_text(
            f"Slot Reminders &amp; Assistant Warnings are currently <b>{status}</b>.\n"
            f"Tagged Assistants: {tagged_escaped}\n"
            f"Use <code>/reminder on</code>, <code>/reminder off</code>, or <code>/tags</code>.",
            parse_mode="HTML",
        )
        return

    arg = context.args[0].strip().lower()
    if arg in ("on", "enable", "start", "true", "1"):
        await asyncio.to_thread(set_setting, "reminders_enabled", "true")
        await update.message.reply_text(
            "Slot Reminders &amp; Assistant Warnings are now <b>ON</b>.", parse_mode="HTML"
        )
    elif arg in ("off", "disable", "stop", "false", "0"):
        await asyncio.to_thread(set_setting, "reminders_enabled", "false")
        await update.message.reply_text(
            "Slot Reminders &amp; Assistant Warnings are now <b>OFF</b>.", parse_mode="HTML"
        )
    else:
        await update.message.reply_text(
            "Invalid option. Use <code>/reminder on</code> or <code>/reminder off</code>.",
            parse_mode="HTML",
        )


async def setworkgroup_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.effective_chat:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can set groups.")
        return
    if update.effective_chat.type == "private":
        await update.message.reply_text(
            "Run <code>/setworkgroup</code> inside your Telegram work group.", parse_mode="HTML"
        )
        return

    await asyncio.to_thread(set_setting, "work_group_id", str(update.effective_chat.id))
    await update.message.reply_text(
        f"This group (<code>{update.effective_chat.id}</code>) is now set as the <b>Work Report Group</b>.",
        parse_mode="HTML",
    )


async def setexcelgroup_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.effective_chat:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can set groups.")
        return
    if update.effective_chat.type == "private":
        await update.message.reply_text(
            "Run <code>/setexcelgroup</code> inside your dedicated Excel Report group.",
            parse_mode="HTML",
        )
        return

    await asyncio.to_thread(set_setting, "excel_group_id", str(update.effective_chat.id))
    await update.message.reply_text(
        f"This group (<code>{update.effective_chat.id}</code>) is now set as the <b>Excel Report Group</b>.\n"
        f"Every slot submission will automatically send <code>team_report_YYYY-MM-DD.xlsx</code> here.",
        parse_mode="HTML",
    )


async def excel_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.effective_chat:
        return
    if not await _require_admin(update):
        return

    target_date = datetime.now(IST).strftime("%Y-%m-%d")
    if context.args:
        candidate = context.args[0].strip()
        try:
            datetime.strptime(candidate, "%Y-%m-%d")
            target_date = candidate
        except ValueError:
            await update.message.reply_text(
                "Invalid date format. Use <code>/excel YYYY-MM-DD</code>.", parse_mode="HTML"
            )
            return

    excel_stream, count = await asyncio.to_thread(build_excel_bytes, target_date, 1)
    filename = f"team_report_{target_date}.xlsx"

    await update.message.reply_document(
        document=excel_stream,
        filename=filename,
        caption=f"Team Photo Report Excel ({target_date})\nLogged submissions: {count}",
    )


async def weekly_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.effective_chat:
        return
    if not await _require_admin(update):
        return

    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    excel_stream, count = await asyncio.to_thread(build_excel_bytes, today_str, 7)
    filename = f"team_weekly_report_{today_str}.xlsx"

    await update.message.reply_document(
        document=excel_stream,
        filename=filename,
        caption=f"7-Day Consolidated Excel Report (ending {today_str})\nTotal submissions: {count}",
    )


async def summary_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not await _require_admin(update):
        return

    sync_cid = (
        update.effective_chat.id
        if update.effective_chat and update.effective_chat.id < 0
        else get_work_group_id()
    )
    await refresh_group_admins(context.bot, sync_cid)
    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    text, markup = await asyncio.to_thread(_build_summary_view, today_str)
    await update.message.reply_text(text, reply_markup=markup, parse_mode="HTML")


async def removesummary_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can remove names from summary.")
        return
    if not context.args:
        excluded = await asyncio.to_thread(get_summary_excluded)
        excl_str = ", ".join(html.escape(x) for x in excluded) if excluded else "None"
        await update.message.reply_text(
            f"Usage: <code>/removesummary @username</code>\n"
            f"Currently removed from Summary: {excl_str}\n\n"
            f"<i>Tip: You can also tap 'Remove / Manage Assistant Names' right below /summary.</i>",
            parse_mode="HTML",
        )
        return

    raw_items = [part for arg in context.args for part in arg.split(",") if part.strip()]
    removed: list[str] = []
    for raw in raw_items:
        res = await asyncio.to_thread(exclude_from_summary, raw)
        if res:
            removed.append(res)

    if removed:
        await update.message.reply_text(
            f"Removed from Summary &amp; Excel Summary: {', '.join(html.escape(r) for r in removed)}\n"
            f"Use <code>/restoresummary @username</code> anytime to bring them back.",
            parse_mode="HTML",
        )
    else:
        await update.message.reply_text("Could not parse any valid username to remove.")


async def restoresummary_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can restore names to summary.")
        return
    if not context.args:
        await update.message.reply_text("Usage: <code>/restoresummary @username</code>", parse_mode="HTML")
        return

    raw_items = [part for arg in context.args for part in arg.split(",") if part.strip()]
    restored: list[str] = []
    for raw in raw_items:
        if await asyncio.to_thread(include_in_summary, raw):
            restored.append(raw)

    if restored:
        await update.message.reply_text(
            f"Restored to Summary &amp; Excel Summary: {', '.join(html.escape(r) for r in restored)}",
            parse_mode="HTML",
        )
    else:
        await update.message.reply_text("No matching hidden assistant found.")


async def mystatus_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.effective_user:
        return

    user = update.effective_user
    username = f"@{user.username}" if user.username else (user.first_name or "Unknown User")
    now_ist = datetime.now(IST)
    today_str = now_ist.strftime("%Y-%m-%d")
    day_rows = await asyncio.to_thread(fetch_reports, today_str)
    stats_list = await asyncio.to_thread(
        compute_member_stats_for_period, [username], day_rows, [today_str], now_ist
    )
    st = stats_list[0]
    submitted_map = await asyncio.to_thread(fetch_user_today_slots, username, today_str)
    closed_labels = {s.label for s in get_closed_slots_for_date(today_str, now_ist)}
    allowed_slots = st["allowed_slots"]
    leaves_map = await asyncio.to_thread(get_leaves_for_date, today_str)
    user_leaves = leaves_map.get(_match_key(username), set())
    missed_pen_rate = get_missed_penalty_rs()
    target_val = get_daily_joining_target()
    is_off, off_reason = is_date_holiday_or_off(today_str)

    tgt_line = ""
    if target_val > 0:
        diff = int(st["target_diff"])
        tgt_line = f" | Target: {target_val} ({'+' if diff >= 0 else ''}{diff})"

    lines = [
        f"<b>Slot Status for {html.escape(username)} ({today_str})</b>",
        f"Shift: <b>{html.escape(str(st['shift']).upper())}</b> | Valid Daily Joining: <b>{st['total_joining']}</b>{tgt_line}",
        f"Fine Due Today: <b>₹{st['balance_due']}</b> (Gross: ₹{st['gross_pen']} | Paid/Waived: ₹{int(st['paid_rs']) + int(st['waived_rs'])})",
    ]
    if is_off:
        lines.append(f"🌴 <b>{html.escape(off_reason)}</b> — Missed penalties paused")

    for slot in REPORTING_SLOTS:
        if slot.label not in allowed_slots:
            lines.append(f"- {html.escape(slot.label)}: Off-Shift")
        elif "" in user_leaves or slot.label in user_leaves:
            lines.append(f"- {html.escape(slot.label)}: <b>ON LEAVE (₹0)</b>")
        elif slot.label in submitted_map:
            sub_t, j_val, st_val, pen_val = submitted_map[slot.label]
            j_note = f" | Join: {html.escape(j_val)}" if j_val else ""
            pen_note = f" | <b>{html.escape(st_val)}</b>" if pen_val > 0 else " | On Time (₹0)"
            lines.append(
                f"- {html.escape(slot.label)}: Done ({html.escape(sub_t)}{j_note}{pen_note})"
            )
        elif slot.label in closed_labels:
            lines.append(f"- {html.escape(slot.label)}: <b>Missed (₹{missed_pen_rate} Penalty)</b>")
        else:
            lines.append(f"- {html.escape(slot.label)}: Pending")

    if update.effective_chat and update.effective_chat.type != "private":
        if not is_allowed_work_group(update.effective_chat.id):
            lines.append(
                f"\n<b>Warning:</b> This group (<code>{update.effective_chat.id}</code>) is not registered as a Work Group. "
                f"Admin can run <code>/addworkgroup</code> or <code>/setworkgroup</code> here."
            )

    await update.message.reply_text("\n".join(lines), parse_mode="HTML")


async def leave_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/leave @user [today|YYYY-MM-DD] [10am|1pm|3pm|5pm|final|full]"""
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can grant leaves.")
        return
    if not context.args:
        await update.message.reply_text(
            "<b>Usage:</b> <code>/leave @username [today|YYYY-MM-DD] [slot|full]</code>\n"
            "Example: <code>/leave @Premgupta2m today 3pm</code> or <code>/leave @Premgupta2m</code>",
            parse_mode="HTML",
        )
        return
    raw_user = context.args[0]
    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    leave_date = today_str
    slot_label = ""
    for tok in context.args[1:]:
        low = tok.strip().lower()
        if low == "today":
            leave_date = today_str
        elif low == "tomorrow":
            leave_date = (datetime.now(IST) + timedelta(days=1)).strftime("%Y-%m-%d")
        elif low == "full":
            slot_label = ""
        elif re.match(r"^\d{4}-\d{2}-\d{2}$", low):
            try:
                datetime.strptime(low, "%Y-%m-%d")
                leave_date = low
            except ValueError:
                pass
        else:
            parsed_s = parse_slot_alias(low)
            if parsed_s:
                slot_label = parsed_s

    canon = await asyncio.to_thread(add_leave, raw_user, leave_date, slot_label)
    scope_txt = slot_label if slot_label else "Full Day (All Slots)"
    await update.message.reply_text(
        f"🌴 <b>Leave Approved</b>\nAssistant: <code>{html.escape(canon)}</code>\nDate: <b>{leave_date}</b> | Scope: <b>{html.escape(scope_txt)}</b>\n<i>Zero missed penalty will be charged for this scope.</i>",
        parse_mode="HTML",
    )


async def unleave_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can remove leaves.")
        return
    if not context.args:
        await update.message.reply_text("Usage: <code>/unleave @username [today|YYYY-MM-DD] [slot]</code>", parse_mode="HTML")
        return
    raw_user = context.args[0]
    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    leave_date = today_str
    slot_label = ""
    for tok in context.args[1:]:
        low = tok.strip().lower()
        if re.match(r"^\d{4}-\d{2}-\d{2}$", low):
            leave_date = low
        else:
            parsed_s = parse_slot_alias(low)
            if parsed_s:
                slot_label = parsed_s
    removed = await asyncio.to_thread(remove_leave, raw_user, leave_date, slot_label)
    await update.message.reply_text(
        f"Removed {removed} leave record(s) for <code>{html.escape(raw_user)}</code> on {leave_date}."
        if removed
        else f"No matching leave found for <code>{html.escape(raw_user)}</code> on {leave_date}.",
        parse_mode="HTML",
    )


def _bulk_settle_all_fines(
    entry_type: str, note: str = "", now_ist: Optional[datetime] = None
) -> tuple[list[tuple[str, int]], int]:
    """Record Paid or Waived entries for all non-admin assistants whose net balance_due > 0."""
    now_ist = now_ist or datetime.now(IST)
    today_str = now_ist.strftime("%Y-%m-%d")
    start_30 = (now_ist - timedelta(days=29)).strftime("%Y-%m-%d")
    rows_30 = fetch_reports(None, start_30)
    dates_30 = _dates_for_period(start_30, today_str)

    _admin_ids, admin_keys = get_known_admin_sets()
    members = [
        m
        for m in _dedup_case_insensitive(
            get_team_members() + get_tagged_assistants() + [r[0] for r in rows_30]
        )
        if _match_key(m) not in admin_keys
    ]
    if not members:
        return [], 0

    stats = compute_member_stats_for_period(members, rows_30, dates_30, now_ist)
    settled: list[tuple[str, int]] = []
    total_cleared = 0
    clean_note = note.strip() or f"Bulk {entry_type} all"
    for st in stats:
        due_amt = int(st["balance_due"])
        if due_amt > 0:
            canon = add_penalty_ledger_entry(
                str(st["member"]), entry_type, due_amt, clean_note, today_str
            )
            settled.append((canon, due_amt))
            total_cleared += due_amt
    return settled, total_cleared


async def _handle_ledger_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    entry_type: str,
    force_all: bool = False,
) -> None:
    if not update.message:
        return
    if not await _require_admin(update):
        return

    is_bulk_all = force_all or (
        bool(context.args)
        and context.args[0].strip().lower() in ("all", "*", "team", "everyone")
    )

    if is_bulk_all:
        note_args = list(context.args or [])
        if not force_all and note_args:
            note_args = note_args[1:]
        if note_args and note_args[0].strip().lower() == "all":
            note_args = note_args[1:]
        note = " ".join(note_args).strip()

        now_ist = datetime.now(IST)
        settled, total_cleared = await asyncio.to_thread(
            _bulk_settle_all_fines, entry_type, note, now_ist
        )
        if not settled:
            await update.message.reply_text(
                "ℹ️ Sabhi assistants ka net fine balance pehle se hi <b>₹0</b> hai.",
                parse_mode="HTML",
            )
            return

        verb = "All Pending Fines Marked PAID" if entry_type == "paid" else "All Pending Fines WAIVED"
        lines = [
            f"💳 <b>{verb}</b>",
            f"Assistants Cleared: <b>{len(settled)}</b> | Total Cleared: <b>₹{total_cleared}</b>",
            "",
        ]
        for idx, (canon, amt_val) in enumerate(settled[:30], start=1):
            lines.append(f"{idx}. <code>{html.escape(canon)}</code> — <b>₹{amt_val}</b>")
        if len(settled) > 30:
            lines.append(f"...and {len(settled) - 30} more")
        lines.append("\n✅ <i>Sabhi assistants ka Net Fine Balance Due ab ₹0 hai!</i>")
        await update.message.reply_text("\n".join(lines), parse_mode="HTML")
        return

    if len(context.args) < 2:
        await update.message.reply_text(
            f"Usage:\n"
            f"• Single assistant: <code>/{entry_type} @username &lt;amount|all&gt; [note]</code>\n"
            f"• Sabhi assistants ek saath: <code>/{entry_type} all [note]</code> ya <code>/{entry_type}all</code>",
            parse_mode="HTML",
        )
        return
    raw_user = context.args[0]
    amt_tok = context.args[1].strip().lstrip("₹")
    if not (amt_tok.isdigit() or amt_tok.lower() == "all"):
        await update.message.reply_text(
            f"Usage:\n"
            f"• Single assistant: <code>/{entry_type} @username &lt;amount|all&gt; [note]</code>\n"
            f"• Sabhi assistants ek saath: <code>/{entry_type} all [note]</code> ya <code>/{entry_type}all</code>",
            parse_mode="HTML",
        )
        return

    now_ist = datetime.now(IST)
    today_str = now_ist.strftime("%Y-%m-%d")
    start_30 = (now_ist - timedelta(days=29)).strftime("%Y-%m-%d")
    rows_30 = await asyncio.to_thread(fetch_reports, None, start_30)
    dates_30 = _dates_for_period(start_30, today_str)
    with sqlite3.connect(DB_PATH) as conn:
        canon_lookup = _resolve_canonical_name(conn, raw_user) or _normalize_tag(raw_user) or raw_user

    if amt_tok.lower() == "all":
        cur_st = (await asyncio.to_thread(compute_member_stats_for_period, [canon_lookup], rows_30, dates_30, now_ist))[0]
        amt = int(cur_st["balance_due"])
        if amt <= 0:
            await update.message.reply_text(
                f"ℹ️ <code>{html.escape(canon_lookup)}</code> ka balance pehle se hi <b>₹0</b> hai.",
                parse_mode="HTML",
            )
            return
    else:
        amt = int(amt_tok)

    note = " ".join(context.args[2:]).strip()
    canon = await asyncio.to_thread(add_penalty_ledger_entry, canon_lookup, entry_type, amt, note, today_str)
    st = (await asyncio.to_thread(compute_member_stats_for_period, [canon], rows_30, dates_30, now_ist))[0]
    verb = "Payment Recorded" if entry_type == "paid" else "Fine Waived"
    await update.message.reply_text(
        f"💳 <b>{verb}</b>\n"
        f"Assistant: <code>{html.escape(canon)}</code> | Amount: <b>₹{amt}</b>\n"
        f"30-Day Gross Fine: ₹{st['gross_pen']} | Paid: ₹{st['paid_rs']} | Waived: ₹{st['waived_rs']}\n"
        f"Remaining Balance Due: <b>₹{st['balance_due']}</b>",
        parse_mode="HTML",
    )


async def paid_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await _handle_ledger_command(update, context, "paid")


async def paidall_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await _handle_ledger_command(update, context, "paid", force_all=True)


async def waive_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await _handle_ledger_command(update, context, "waive")


async def waiveall_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await _handle_ledger_command(update, context, "waive", force_all=True)


async def resetfine_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Admin command /resetfine [@user | all] to reset existing late/missed/balance fines to ₹0."""
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can reset fines.")
        return

    now_ist = datetime.now(IST)
    today_str = now_ist.strftime("%Y-%m-%d")
    now_secs = _time_to_secs(now_ist.time().replace(tzinfo=None))
    target = context.args[0].strip() if context.args else "all"

    if target.lower() in ("all", "*", "team", "everyone"):
        def _reset_all() -> int:
            with sqlite3.connect(DB_PATH) as conn:
                cur = conn.execute("UPDATE reports SET penalty_rs = 0 WHERE penalty_rs != 0")
                conn.execute("DELETE FROM penalty_ledger")
                conn.commit()
                changed = cur.rowcount
            set_setting("fine_reset_date", today_str)
            set_setting("fine_reset_time_secs", str(now_secs))
            set_setting("member_fine_reset_dates", "{}")
            return changed

        cleared_rows = await asyncio.to_thread(_reset_all)
        await update.message.reply_text(
            f"✅ <b>All Team Fines Reset to ₹0!</b>\n"
            f"Reset Date Baseline: <b>{today_str} ({now_ist.strftime('%I:%M %p IST')})</b>\n"
            f"Cleared Late Penalty Rows: <b>{cleared_rows}</b>\n"
            f"<i>Sabhi assistants ka Gross, Missed aur Net Balance ab ₹0 hai.</i>",
            parse_mode="HTML",
        )
        return

    with sqlite3.connect(DB_PATH) as conn:
        canon = _resolve_canonical_name(conn, target) or _normalize_tag(target) or target
        m_key = _match_key(canon)
        cur = conn.execute(
            "UPDATE reports SET penalty_rs = 0 WHERE LOWER(LTRIM(username, '@')) = ? AND penalty_rs != 0",
            (m_key,),
        )
        conn.execute("DELETE FROM penalty_ledger WHERE username_key = ?", (m_key,))
        conn.commit()
        changed = cur.rowcount

    raw_map = get_setting("member_fine_reset_dates", "{}") or "{}"
    try:
        m_map = json.loads(raw_map)
        if not isinstance(m_map, dict):
            m_map = {}
    except Exception:
        m_map = {}
    m_map[m_key] = {"date": today_str, "secs": now_secs}
    await asyncio.to_thread(set_setting, "member_fine_reset_dates", json.dumps(m_map))

    await update.message.reply_text(
        f"✅ <b>Fine Reset to ₹0 for {html.escape(canon)}!</b>\n"
        f"Cleared Late Rows: <b>{changed}</b> | Baseline: <b>{today_str} ({now_ist.strftime('%I:%M %p IST')})</b>",
        parse_mode="HTML",
    )


def _build_all_fines_view(days: int = 30) -> tuple[str, InlineKeyboardMarkup]:
    """Consolidated all-assistants penalty sheet (Today + Period Late/Missed/Paid/Waived/Net Due)."""
    days = max(1, min(90, int(days)))
    now_ist = datetime.now(IST)
    today_str = now_ist.strftime("%Y-%m-%d")
    start_str = (now_ist - timedelta(days=days - 1)).strftime("%Y-%m-%d")
    period_rows = fetch_reports(None, start_str)
    today_rows = [r for r in period_rows if r[4] == today_str]
    date_list = _dates_for_period(start_str, today_str)

    _admin_ids, admin_keys = get_known_admin_sets()
    members = [
        m
        for m in _dedup_case_insensitive(
            get_team_members() + get_tagged_assistants() + [r[0] for r in period_rows]
        )
        if _match_key(m) not in admin_keys
    ]

    reg_p, fin_p, miss_p = get_penalty_config()
    if not members:
        markup = InlineKeyboardMarkup(
            [[InlineKeyboardButton("🔄 Refresh", callback_data=f"allfines_days:{days}")]]
        )
        return (
            f"💳 <b>All Assistants Penalty Report ({'Today' if days == 1 else f'Last {days} Days'})</b>\n"
            f"Abhi koi non-admin assistant record nahi mila.",
            markup,
        )

    period_stats = compute_member_stats_for_period(members, period_rows, date_list, now_ist)
    today_stats = compute_member_stats_for_period(members, today_rows, [today_str], now_ist)
    today_by_key = {str(s["key"]): s for s in today_stats}

    total_due = sum(int(s["balance_due"]) for s in period_stats)
    total_gross = sum(int(s["gross_pen"]) for s in period_stats)
    total_late = sum(int(s["late_pen"]) for s in period_stats)
    total_missed = sum(int(s["missed_pen"]) for s in period_stats)
    total_paid = sum(int(s["paid_rs"]) for s in period_stats)
    total_waived = sum(int(s["waived_rs"]) for s in period_stats)
    total_today_gross = sum(int(today_by_key[str(s["key"])]["gross_pen"]) for s in period_stats if str(s["key"]) in today_by_key)

    sorted_stats = sorted(
        period_stats,
        key=lambda s: (
            int(s["balance_due"]),
            int(today_by_key.get(str(s["key"]), {}).get("gross_pen", 0)),
            int(s["gross_pen"]),
            -int(s["total_joining"]),
        ),
        reverse=True,
    )

    fined_rows = [
        s for s in sorted_stats
        if int(s["balance_due"]) > 0 or int(s["gross_pen"]) > 0 or int(s["paid_rs"]) > 0 or int(s["waived_rs"]) > 0
    ]
    clean_names = [
        str(s["member"]) for s in sorted_stats
        if int(s["balance_due"]) == 0 and int(s["gross_pen"]) == 0 and int(s["paid_rs"]) == 0 and int(s["waived_rs"]) == 0
    ]

    period_lbl = f"Today ({today_str})" if days == 1 else f"Last {days} Days (to {today_str})"
    lines = [
        f"💳 <b>All Assistants Penalty Report — {period_lbl}</b>",
        f"Total Net Balance Due: <b>₹{total_due}</b> | Today Gross: <b>₹{total_today_gross}</b>",
        f"Gross: ₹{total_gross} (Late ₹{total_late} + Missed ₹{total_missed}) | Paid: ₹{total_paid} | Waived: ₹{total_waived}",
        f"<i>Rates: Reg ₹{reg_p} | Final ₹{fin_p} | Missed ₹{miss_p}</i>",
        "",
    ]

    if fined_rows:
        for idx, s in enumerate(fined_rows[:35], start=1):
            t_st = today_by_key.get(str(s["key"]), {})
            t_due = int(t_st.get("balance_due", 0))
            late_cnt = int(s["late_regular"]) + int(s["late_final"])
            miss_cnt = int(s["missed"])
            adj_bits = []
            if int(s["paid_rs"]) > 0:
                adj_bits.append(f"Paid ₹{s['paid_rs']}")
            if int(s["waived_rs"]) > 0:
                adj_bits.append(f"Waived ₹{s['waived_rs']}")
            adj_txt = f" | {', '.join(adj_bits)}" if adj_bits else ""
            lines.append(
                f"{idx}. {html.escape(str(s['member']))} — Due: <b>₹{s['balance_due']}</b> "
                f"(Today ₹{t_due} | {late_cnt}L ₹{s['late_pen']} + {miss_cnt}M ₹{s['missed_pen']}{adj_txt})"
            )
    else:
        lines.append("🎉 <b>Zero Pending Fines!</b> Sabhi assistants ka balance ₹0 hai.")

    if clean_names:
        shown_clean = ", ".join(html.escape(n) for n in clean_names[:15])
        extra_cnt = f" +{len(clean_names) - 15} more" if len(clean_names) > 15 else ""
        lines.append(f"\n✅ <b>Zero-Fine ({len(clean_names)}):</b> {shown_clean}{extra_cnt}")

    lines.append("\n<i>Settle: /paid all (ya /paidall) | /paid @user all | /waive all | /resetfine [@user|all]</i>")

    markup = InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton("• Today •" if days == 1 else "Today", callback_data="allfines_days:1"),
                InlineKeyboardButton("• 7 Days •" if days == 7 else "7 Days", callback_data="allfines_days:7"),
                InlineKeyboardButton("• 30 Days •" if days == 30 else "30 Days", callback_data="allfines_days:30"),
                InlineKeyboardButton("🔄 Refresh", callback_data=f"allfines_days:{days}"),
            ],
            [
                InlineKeyboardButton("💳 Mark All Paid (₹0)", callback_data=f"allfines_paid_ask:{days}"),
                InlineKeyboardButton("🧹 Reset All Fines to ₹0", callback_data=f"allfines_reset_ask:{days}"),
            ],
        ]
    )
    return "\n".join(lines), markup


async def allfines_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Admin command /allfines [today|7|30] — show all assistants' penalties together."""
    if not update.message:
        return
    if not await _require_admin(update):
        return
    days = 30
    if context.args:
        arg = context.args[0].strip().lower()
        if arg in ("today", "1", "1d"):
            days = 1
        elif arg.rstrip("d").isdigit():
            days = max(1, min(90, int(arg.rstrip("d"))))
    text, markup = await asyncio.to_thread(_build_all_fines_view, days)
    await update.message.reply_text(text, reply_markup=markup, parse_mode="HTML")


async def myrank_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Dedicated Assistant command: shows ONLY the caller's own rank, today's valid joining, and on-time score."""
    if not update.message or not update.effective_user:
        return
    user = update.effective_user
    self_raw = f"@{user.username}" if user.username else (user.first_name or "Assistant")
    with sqlite3.connect(DB_PATH) as conn:
        canon = _resolve_canonical_name(conn, self_raw) or _normalize_tag(self_raw) or self_raw
    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    rows = await asyncio.to_thread(fetch_reports, today_str)
    members = await asyncio.to_thread(_get_visible_members_for_rows, rows)
    if _match_key(canon) not in {_match_key(m) for m in members}:
        members.append(canon)
    stats = await asyncio.to_thread(compute_member_stats_for_period, members, rows, [today_str])
    ranked = sorted(
        stats,
        key=lambda s: (int(s["total_joining"]), int(s["on_time"]), -int(s["gross_pen"])),
        reverse=True,
    )
    my_key = _match_key(canon)
    my_idx = next((i + 1 for i, s in enumerate(ranked) if str(s["key"]) == my_key), len(ranked))
    my_st = next((s for s in ranked if str(s["key"]) == my_key), ranked[0])
    target_val = get_daily_joining_target()
    tgt_line = f"\n🎯 Target: {target_val} ({'+' if int(my_st['target_diff']) >= 0 else ''}{my_st['target_diff']})" if target_val > 0 else ""

    await update.message.reply_text(
        f"🏅 <b>My Daily Rank — {html.escape(canon)} ({today_str})</b>\n"
        f"Your Position: <b>#{my_idx}</b> out of <b>{len(ranked)}</b> assistants\n"
        f"Valid Daily Joining: <b>{my_st['total_joining']}</b>\n"
        f"On-Time Slots: <b>{my_st['on_time']}/{max(len(my_st['allowed_slots']), 1)}</b> | Compliance: <b>{my_st['compliance']}%</b>"
        f"{tgt_line}",
        parse_mode="HTML",
    )


async def leaderboard_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not await _require_admin(update):
        return
    target_date = datetime.now(IST).strftime("%Y-%m-%d")
    if context.args and re.match(r"^\d{4}-\d{2}-\d{2}$", context.args[0].strip()):
        target_date = context.args[0].strip()
    text = await asyncio.to_thread(_build_leaderboard_text, target_date)
    await update.message.reply_text(text, parse_mode="HTML")


async def monthly_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.effective_chat:
        return
    if not await _require_admin(update):
        return
    month_str = datetime.now(IST).strftime("%Y-%m")
    if context.args:
        cand = context.args[0].strip()
        if re.match(r"^\d{4}-\d{2}$", cand):
            month_str = cand
        else:
            await update.message.reply_text("Usage: <code>/monthly [YYYY-MM]</code>", parse_mode="HTML")
            return
    stream, count = await asyncio.to_thread(build_excel_bytes, None, 1, month_str)
    await update.message.reply_document(
        document=stream,
        filename=f"team_monthly_report_{month_str}.xlsx",
        caption=f"Monthly Consolidated Excel ({month_str}) | Valid Daily Joining sums across days | Logs: {count}",
    )


async def payout_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.effective_chat:
        return
    if not await _require_admin(update):
        return
    days = 7
    rate: Optional[int] = None
    for tok in context.args:
        clean_t = tok.strip().lstrip("₹").lower()
        if clean_t.endswith("d") and clean_t[:-1].isdigit():
            days = max(1, min(90, int(clean_t[:-1])))
        elif tok.strip().startswith("₹") or (clean_t.isdigit() and days != 7):
            rate = int(clean_t)
        elif clean_t.isdigit():
            if int(clean_t) <= 31 and rate is None and len(context.args) >= 2:
                days = max(1, int(clean_t))
            else:
                rate = int(clean_t)
    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    used_rate = rate if rate is not None else get_per_joining_rate()
    stream, m_count = await asyncio.to_thread(build_payout_excel_bytes, today_str, days, used_rate)
    await update.message.reply_document(
        document=stream,
        filename=f"team_payout_{days}d_{today_str}.xlsx",
        caption=(
            f"💰 <b>Payout Sheet ({days} Days ending {today_str})</b>\n"
            f"Rate: ₹{used_rate}/joining | Assistants: {m_count}\n"
            f"(Tip: Set default rate via <code>/setrate 50</code>)"
        ),
        parse_mode="HTML",
    )


async def setrate_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not await _require_admin(update):
        return
    if not context.args or not context.args[0].lstrip("₹").isdigit():
        await update.message.reply_text(
            f"Current Rate per Valid Joining: <b>₹{get_per_joining_rate()}</b>\nUsage: <code>/setrate 50</code>",
            parse_mode="HTML",
        )
        return
    rate = int(context.args[0].lstrip("₹"))
    await asyncio.to_thread(set_setting, "per_joining_rate", str(rate))
    await update.message.reply_text(f"Payout rate updated to <b>₹{rate}</b> per valid joining.", parse_mode="HTML")


async def settarget_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can set daily joining target.")
        return
    if not context.args or not context.args[0].isdigit():
        await update.message.reply_text(
            f"Current Daily Joining Target: <b>{get_daily_joining_target()}</b>\nUsage: <code>/settarget 50</code> (or <code>/settarget 0</code> to disable)",
            parse_mode="HTML",
        )
        return
    tgt = int(context.args[0])
    await asyncio.to_thread(set_setting, "daily_joining_target", str(tgt))
    await update.message.reply_text(
        f"Daily Joining Target set to <b>{tgt}</b>." if tgt > 0 else "Daily Joining Target disabled.",
        parse_mode="HTML",
    )


async def alias_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can link username aliases.")
        return
    if len(context.args) < 2:
        await update.message.reply_text("Usage: <code>/alias @old_username @main_username</code>", parse_mode="HTML")
        return
    old_u, new_u, merged = await asyncio.to_thread(add_username_alias, context.args[0], context.args[1])
    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    await push_excel_to_group(context, target_date=today_str, caption=f"Excel Merged Alias | {old_u} -> {new_u}")
    await update.message.reply_text(
        f"🔗 <b>Alias Linked &amp; History Merged</b>\n<code>{html.escape(old_u)}</code> ➡️ <code>{html.escape(new_u)}</code>\nMerged duplicate rows: <b>{merged}</b>",
        parse_mode="HTML",
    )


async def aliases_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not await _require_admin(update):
        return
    pairs = await asyncio.to_thread(get_username_aliases)
    if not pairs:
        await update.message.reply_text("No username aliases linked yet. Use <code>/alias @old @new</code>.", parse_mode="HTML")
        return
    lines = ["<b>Linked Username Aliases:</b>"] + [
        f"- <code>{html.escape(a)}</code> ➡️ <code>{html.escape(t)}</code>" for a, t in pairs
    ]
    await update.message.reply_text("\n".join(lines), parse_mode="HTML")


async def holiday_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not await _require_admin(update):
        return
    if not context.args:
        cur = get_setting("holiday_mode", "false")
        await update.message.reply_text(
            f"Holiday Mode is currently <b>{'ON' if cur == 'true' else 'OFF'}</b>.\nUse <code>/holiday on</code> or <code>/holiday off</code>.",
            parse_mode="HTML",
        )
        return
    arg = context.args[0].strip().lower()
    if arg in ("on", "true", "1", "yes"):
        await asyncio.to_thread(set_setting, "holiday_mode", "true")
        await update.message.reply_text("🌴 <b>Holiday Mode is ON</b> — Reminders &amp; missed penalties paused.", parse_mode="HTML")
    elif arg in ("off", "false", "0", "no"):
        await asyncio.to_thread(set_setting, "holiday_mode", "false")
        await update.message.reply_text("✅ <b>Holiday Mode is OFF</b> — Normal reporting active.", parse_mode="HTML")
    else:
        await update.message.reply_text("Usage: <code>/holiday on</code> or <code>/holiday off</code>", parse_mode="HTML")


async def offdays_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not await _require_admin(update):
        return
    if not context.args:
        cur = get_setting("off_days", "") or "none"
        await update.message.reply_text(
            f"Recurring Off-Days: <b>{html.escape(cur.upper())}</b>\nUsage: <code>/offdays sun</code>, <code>/offdays sat,sun</code>, or <code>/offdays none</code>",
            parse_mode="HTML",
        )
        return
    raw = ",".join(context.args).lower()
    if "none" in raw or "off" in raw:
        await asyncio.to_thread(set_setting, "off_days", "")
        await update.message.reply_text("Recurring weekly off-days cleared (all 7 days active).")
        return
    valid_days = {"mon", "tue", "wed", "thu", "fri", "sat", "sun"}
    tokens = [t.strip()[:3] for t in raw.split(",") if t.strip()[:3] in valid_days]
    unique_toks = list(dict.fromkeys(tokens))
    if not unique_toks:
        await update.message.reply_text("Valid days: <code>mon, tue, wed, thu, fri, sat, sun</code> or <code>none</code>", parse_mode="HTML")
        return
    val = ",".join(unique_toks)
    await asyncio.to_thread(set_setting, "off_days", val)
    await update.message.reply_text(f"🌴 Recurring Off-Days set to: <b>{html.escape(val.upper())}</b>", parse_mode="HTML")


async def cleanchat_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not await _require_admin(update):
        return
    if not context.args:
        await update.message.reply_text(
            f"Clean Chat Mode (auto-delete previous reminders): <b>{'ON' if is_clean_chat_enabled() else 'OFF'}</b>\nUse <code>/cleanchat on</code> or <code>/cleanchat off</code>",
            parse_mode="HTML",
        )
        return
    arg = context.args[0].strip().lower()
    if arg in ("on", "true", "1"):
        await asyncio.to_thread(set_setting, "clean_chat_mode", "true")
        await update.message.reply_text("🧹 <b>Clean Chat Mode ON</b> — Previous slot reminders will be auto-deleted when a new slot starts.", parse_mode="HTML")
    elif arg in ("off", "false", "0"):
        await asyncio.to_thread(set_setting, "clean_chat_mode", "false")
        await update.message.reply_text("Clean Chat Mode <b>OFF</b>.", parse_mode="HTML")
    else:
        await update.message.reply_text("Usage: <code>/cleanchat on</code> or <code>/cleanchat off</code>", parse_mode="HTML")


async def setpenalty_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/setpenalty regular 5 final 10 missed 15"""
    if not update.message:
        return
    if not await _require_admin(update):
        return
    reg_p, fin_p, miss_p = get_penalty_config()
    if not context.args:
        await update.message.reply_text(
            f"<b>Current Penalty Rates:</b>\n"
            f"- Regular Late (10am/1pm/3pm/5pm): <b>₹{reg_p}</b>\n"
            f"- Final Late (after 8:00 PM): <b>₹{fin_p}</b>\n"
            f"- Missed Slot: <b>₹{miss_p}</b>\n\n"
            f"Update: <code>/setpenalty regular 5 final 10 missed 15</code>",
            parse_mode="HTML",
        )
        return
    tokens = [t.strip().lower().lstrip("₹") for t in context.args]
    new_reg, new_fin, new_miss = reg_p, fin_p, miss_p
    i = 0
    while i < len(tokens) - 1:
        k, v = tokens[i], tokens[i + 1]
        if v.isdigit():
            if k in ("regular", "reg", "late"):
                new_reg = int(v)
            elif k in ("final", "fin"):
                new_fin = int(v)
            elif k in ("missed", "miss"):
                new_miss = int(v)
            i += 2
            continue
        i += 1
    if len(tokens) == 3 and all(t.isdigit() for t in tokens):
        new_reg, new_fin, new_miss = int(tokens[0]), int(tokens[1]), int(tokens[2])

    await asyncio.to_thread(set_setting, "penalty_late_regular", str(new_reg))
    await asyncio.to_thread(set_setting, "penalty_late_final", str(new_fin))
    await asyncio.to_thread(set_setting, "penalty_missed", str(new_miss))
    with sqlite3.connect(DB_PATH) as conn:
        fixed = recalculate_report_statuses(conn)
        conn.commit()
    await update.message.reply_text(
        f"⚙️ <b>Penalty Rates Updated</b>\nRegular Late: <b>₹{new_reg}</b> | Final Late: <b>₹{new_fin}</b> | Missed: <b>₹{new_miss}</b>\nSynced {fixed} existing row(s).",
        parse_mode="HTML",
    )


async def setshift_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not await _require_admin(update):
        return
    if not context.args:
        shifts = await asyncio.to_thread(get_assistant_shifts)
        lines = [
            "<b>Assistant Shift Assignments:</b>",
            "• <code>morning</code>: 10 AM, 1 PM, 3 PM",
            "• <code>evening</code>: 3 PM, 5 PM, Final (8 PM)",
            "• <code>full</code>: All 5 slots",
            "",
        ]
        if shifts:
            lines.extend(f"- @{html.escape(k)}: <b>{html.escape(v.upper())}</b>" for k, v in sorted(shifts.items()))
        else:
            lines.append("All assistants are on <b>FULL</b> shift by default.")
        lines.append("\nSet: <code>/setshift @user morning|evening|full</code>")
        await update.message.reply_text("\n".join(lines), parse_mode="HTML")
        return
    if len(context.args) < 2 or context.args[1].strip().lower() not in SHIFT_SLOTS:
        await update.message.reply_text("Usage: <code>/setshift @user morning|evening|full</code>", parse_mode="HTML")
        return
    sh = context.args[1].strip().lower()
    canon = await asyncio.to_thread(set_assistant_shift, context.args[0], sh)
    slots_desc = ", ".join(sorted(SHIFT_SLOTS[sh], key=lambda x: SLOT_ORDER.get(x, 99)))
    await update.message.reply_text(
        f"🕒 Shift updated for <code>{html.escape(canon)}</code> ➡️ <b>{sh.upper()}</b>\nSlots: {html.escape(slots_desc)}",
        parse_mode="HTML",
    )


async def find_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not await _require_admin(update):
        return
    if not context.args:
        await update.message.reply_text("Usage: <code>/find @username [days=7]</code>", parse_mode="HTML")
        return
    raw_user = context.args[0]
    days = int(context.args[1]) if len(context.args) > 1 and context.args[1].isdigit() else 7
    days = max(1, min(90, days))
    now_ist = datetime.now(IST)
    today_str = now_ist.strftime("%Y-%m-%d")
    start_str = (now_ist - timedelta(days=days - 1)).strftime("%Y-%m-%d")
    date_list = _dates_for_period(start_str, today_str)
    rows = await asyncio.to_thread(fetch_reports, None, start_str)
    with sqlite3.connect(DB_PATH) as conn:
        canon = _resolve_canonical_name(conn, raw_user) or _normalize_tag(raw_user) or raw_user
    st = (await asyncio.to_thread(compute_member_stats_for_period, [canon], rows, date_list, now_ist))[0]
    m_key = _match_key(canon)
    m_rows = [r for r in rows if _match_key(r[0]) == m_key]
    by_date: dict[str, list[tuple[str, str, str, str, str, str, int, str]]] = {}
    for r in m_rows:
        by_date.setdefault(r[4], []).append(r)

    lines = [
        f"🔍 <b>Assistant Lookup: {html.escape(canon)} (Last {days} Days)</b>",
        f"Valid Joining Total: <b>{st['total_joining']}</b> | Compliance: <b>{st['compliance']}%</b>",
        f"On-Time: {st['on_time']} | Late: {int(st['late_regular']) + int(st['late_final'])} | Missed: {st['missed']}",
        f"Gross Fine: ₹{st['gross_pen']} | Paid: ₹{st['paid_rs']} | Waived: ₹{st['waived_rs']} | <b>Balance Due: ₹{st['balance_due']}</b>",
        "",
    ]
    for d in reversed(date_list):
        d_rows = by_date.get(d, [])
        if not d_rows:
            continue
        day_j = compute_daily_joining_from_rows(d_rows)
        slot_bits = ", ".join(
            f"{r[3].split(' ')[0]}={r[1] or '-'}({'✓' if r[6] == 0 else f'₹{r[6]}'})"
            for r in sorted(d_rows, key=lambda x: SLOT_ORDER.get(x[3], 99))
        )
        lines.append(f"• <b>{d}</b>: Valid Join <b>{day_j}</b> [{html.escape(slot_bits)}]")
    if len(lines) <= 5:
        lines.append("No submissions found in this period.")
    await update.message.reply_text("\n".join(lines[:35]), parse_mode="HTML")


async def day_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not await _require_admin(update):
        return
    target_date = datetime.now(IST).strftime("%Y-%m-%d")
    if context.args:
        cand = context.args[0].strip()
        try:
            datetime.strptime(cand, "%Y-%m-%d")
            target_date = cand
        except ValueError:
            await update.message.reply_text("Usage: <code>/day YYYY-MM-DD</code>", parse_mode="HTML")
            return
    text, _markup = await asyncio.to_thread(_build_summary_view, target_date)
    await update.message.reply_text(text, parse_mode="HTML")


async def myfine_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.effective_user:
        return
    user = update.effective_user
    self_name = f"@{user.username}" if user.username else (user.first_name or "Unknown User")
    target_raw = context.args[0] if (context.args and _is_admin(update)) else self_name
    with sqlite3.connect(DB_PATH) as conn:
        canon = _resolve_canonical_name(conn, target_raw) or _normalize_tag(target_raw) or target_raw
    now_ist = datetime.now(IST)
    today_str = now_ist.strftime("%Y-%m-%d")
    start_30 = (now_ist - timedelta(days=29)).strftime("%Y-%m-%d")
    rows_30 = await asyncio.to_thread(fetch_reports, None, start_30)
    dates_30 = _dates_for_period(start_30, today_str)
    st_30 = (await asyncio.to_thread(compute_member_stats_for_period, [canon], rows_30, dates_30, now_ist))[0]
    today_rows = [r for r in rows_30 if r[4] == today_str]
    st_today = (await asyncio.to_thread(compute_member_stats_for_period, [canon], today_rows, [today_str], now_ist))[0]
    ledger_recent = await asyncio.to_thread(get_member_ledger_entries, canon, 5)

    lines = [
        f"💳 <b>Fine Card — {html.escape(canon)}</b>",
        f"Today ({today_str}): Gross ₹{st_today['gross_pen']} | Balance Due: <b>₹{st_today['balance_due']}</b>",
        f"Last 30 Days: Valid Join <b>{st_30['total_joining']}</b> | Late ₹{st_30['late_pen']} + Missed ₹{st_30['missed_pen']} = Gross <b>₹{st_30['gross_pen']}</b>",
        f"Paid: <b>₹{st_30['paid_rs']}</b> | Waived: <b>₹{st_30['waived_rs']}</b>",
        f"<b>Net Fine Balance Due: ₹{st_30['balance_due']}</b>",
    ]
    if ledger_recent:
        lines.append("\n<b>Recent Payments / Waivers:</b>")
        for etype, amt, note, edate in ledger_recent:
            n_txt = f" ({html.escape(note)})" if note else ""
            lines.append(f"- {edate}: {etype.upper()} ₹{amt}{n_txt}")
    await update.message.reply_text("\n".join(lines), parse_mode="HTML")


async def addworkgroup_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.effective_chat:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can add work groups.")
        return
    if update.effective_chat.type == "private":
        await update.message.reply_text("Run <code>/addworkgroup</code> inside the group you want to track.", parse_mode="HTML")
        return
    title = update.effective_chat.title or f"Group {update.effective_chat.id}"
    await asyncio.to_thread(add_work_group, update.effective_chat.id, title)
    await update.message.reply_text(
        f"✅ Added <b>{html.escape(title)}</b> (<code>{update.effective_chat.id}</code>) to active Work Groups!",
        parse_mode="HTML",
    )


async def removeworkgroup_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.effective_chat:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can remove work groups.")
        return
    cid = int(context.args[0]) if (context.args and context.args[0].lstrip("-").isdigit()) else update.effective_chat.id
    removed = await asyncio.to_thread(remove_work_group, cid)
    await update.message.reply_text(
        f"Removed work group <code>{cid}</code> from extra tracked groups."
        if removed
        else f"Group <code>{cid}</code> was not in tracked_chats.",
        parse_mode="HTML",
    )


async def groups_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can view configured groups.")
        return
    wgroups = await asyncio.to_thread(get_tracked_work_groups)
    excel_gid = get_excel_group_id()
    lines = ["<b>Configured Groups:</b>"]
    if wgroups:
        for cid, ttl in wgroups:
            lines.append(f"• Work: <b>{html.escape(ttl)}</b> (<code>{cid}</code>)")
    else:
        lines.append("• Work: Any group (none locked yet)")
    lines.append(f"• Excel/DB Group: <code>{excel_gid or 'Not set'}</code>")
    await update.message.reply_text("\n".join(lines), parse_mode="HTML")


async def addmember_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can modify the team roster.")
        return
    if not context.args:
        await update.message.reply_text("Usage: <code>/addmember @user1 @user2</code>", parse_mode="HTML")
        return

    added = []
    for raw in context.args:
        norm = _normalize_tag(raw)
        if norm and await asyncio.to_thread(add_team_member, norm):
            added.append(norm)

    if added:
        await update.message.reply_text(f"Added to team roster: {', '.join(added)}")
    else:
        await update.message.reply_text("Member(s) already in roster.")


async def removemember_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can modify the team roster.")
        return
    if not context.args:
        await update.message.reply_text("Usage: <code>/removemember @username</code>", parse_mode="HTML")
        return

    removed = []
    for raw in context.args:
        if await asyncio.to_thread(remove_team_member, raw):
            removed.append(raw)

    if removed:
        await update.message.reply_text(f"Removed from roster & tags: {', '.join(removed)}")
    else:
        await update.message.reply_text("No matching member found in roster.")


async def team_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not await _require_admin(update):
        return
    sync_cid = (
        update.effective_chat.id
        if update.effective_chat and update.effective_chat.id < 0
        else get_work_group_id()
    )
    await refresh_group_admins(context.bot, sync_cid)
    members = await asyncio.to_thread(get_team_members)
    tagged = {m.lower() for m in await asyncio.to_thread(get_tagged_assistants)}
    if not members:
        await update.message.reply_text(
            "No non-admin assistants detected yet. Assistants are auto-detected when they message in the work group.",
            parse_mode="HTML",
        )
        return
    lines = ["Recognized Assistants (Non-Admin Roster):"]
    for m in members:
        badge = " (Tag: ON)" if m.lower() in tagged else " (Tag: OFF)"
        lines.append(f"- {m}{badge}")
    await update.message.reply_text("\n".join(lines))


def _format_restore_summary(res: dict[str, object], filename: str) -> str:
    lines = [
        f"<b>Recovery Complete from {html.escape(filename or 'backup')}</b>",
        f"- Reports Added: <b>{res['reports']}</b>",
        f"- Members Recovered: <b>{res['members']}</b>",
        f"- Assistant Tags Updated: <b>{res['tags']}</b>",
    ]
    if res.get("work_group_id") is not None:
        lines.append(f"- Old Work Group Restored: <code>{res['work_group_id']}</code>")
    if res.get("excel_group_id") is not None:
        lines.append(f"- Excel/DB Group Restored: <code>{res['excel_group_id']}</code>")
    return "\n".join(lines)


async def restore_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Recover old reports, old Work Group ID, members & tags from ANY .xlsx or .db (reports.db) file.
    Works either by replying /restore to a file OR running /restore directly to pull the pinned backup
    from the Database/Excel Group.
    """
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Unauthorized.")
        return

    target_doc = None
    replied = update.message.reply_to_message
    if replied and replied.document:
        target_doc = replied.document
    else:
        # Try pinned message in Database/Excel group, or current chat
        candidate_chat_ids = []
        excel_gid = get_excel_group_id()
        if excel_gid is not None:
            candidate_chat_ids.append(excel_gid)
        if update.effective_chat and update.effective_chat.id not in candidate_chat_ids:
            candidate_chat_ids.append(update.effective_chat.id)

        for cid in candidate_chat_ids:
            try:
                c = await context.bot.get_chat(cid)
                if c.pinned_message and c.pinned_message.document:
                    fname = (c.pinned_message.document.file_name or "").lower()
                    if fname.endswith(".xlsx") or fname.endswith(".db") or fname.endswith(".sqlite"):
                        target_doc = c.pinned_message.document
                        break
            except Exception:
                continue

    if not target_doc:
        await update.message.reply_text(
            "<b>How to recover old data &amp; old group:</b>\n"
            "1. Reply <code>/restore</code> directly to any old <code>reports.db</code> or <code>team_report_*.xlsx</code> file in your Database Group.\n"
            "2. Or forward old <code>reports.db</code> / <code>.xlsx</code> files into the Database Group — the bot auto-recovers them immediately!",
            parse_mode="HTML",
        )
        return

    fname = target_doc.file_name or "backup"
    tg_file = await target_doc.get_file()
    buf = BytesIO()
    await tg_file.download_to_memory(out=buf)
    try:
        res = await asyncio.to_thread(restore_any_backup_bytes, buf.getvalue(), fname)
    except Exception as exc:
        logger.exception("Restore failed for %s: %s", fname, exc)
        await update.message.reply_text(
            f"Could not restore from <code>{html.escape(fname)}</code>: {html.escape(str(exc))}",
            parse_mode="HTML",
        )
        return

    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    await push_excel_to_group(
        context,
        target_date=today_str,
        caption=f"Excel Synced After Recovery ({res['reports']} reports recovered from {fname})",
    )
    await update.message.reply_text(_format_restore_summary(res, fname), parse_mode="HTML")


async def handle_backup_document_upload(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Auto-recover when Admin forwards/uploads an old reports.db or .xlsx file into the Database Group (or DM).
    """
    message = update.message
    user = update.effective_user
    chat = update.effective_chat
    if not message or not message.document or not user or not chat or user.is_bot:
        return

    fname = (message.document.file_name or "").lower()
    if not (fname.endswith(".db") or fname.endswith(".sqlite") or fname.endswith(".xlsx")):
        return

    excel_gid = get_excel_group_id()
    is_allowed_place = (chat.type == "private") or (excel_gid is not None and chat.id == excel_gid)
    if not is_allowed_place or not _is_admin(update):
        return

    tg_file = await message.document.get_file()
    buf = BytesIO()
    await tg_file.download_to_memory(out=buf)
    try:
        res = await asyncio.to_thread(restore_any_backup_bytes, buf.getvalue(), message.document.file_name or fname)
    except Exception as exc:
        logger.warning("Auto backup restore failed for %s: %s", fname, exc)
        await message.reply_text(
            f"Backup read failed for <code>{html.escape(fname)}</code>: {html.escape(str(exc))}",
            parse_mode="HTML",
        )
        return

    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    await push_excel_to_group(
        context,
        target_date=today_str,
        caption=f"Excel Synced After Auto-Recovery ({res['reports']} reports from {message.document.file_name or fname})",
    )
    await message.reply_text(
        _format_restore_summary(res, message.document.file_name or fname),
        parse_mode="HTML",
        reply_to_message_id=message.message_id,
    )


# ---------------------------------------------------------------------------
# Scheduled Jobs (Daily Excel + Tagged Assistant Warnings + 15-Min Animated Countdown)
# ---------------------------------------------------------------------------
_COUNTDOWN_ANIM_FRAMES: tuple[str, ...] = ("⏳🚨", "⌛⏰", "⏱🔔", "🔥⚡")


def _get_slot_countdown_deadline(now_ist: datetime, slot_label: str) -> Optional[datetime]:
    slot = _get_slot_by_label(slot_label)
    if not slot:
        return None
    if _is_final_slot(slot_label):
        # Final Report countdown runs from 7:45 PM (final_warning_time) to 8:00:59 PM (slot.end)
        warn_dt = datetime.combine(now_ist.date(), slot.final_warning_time, tzinfo=IST)
        end_dt = datetime.combine(now_ist.date(), slot.end, tzinfo=IST)
        if warn_dt <= now_ist <= end_dt:
            return end_dt
        return None

    slot_end_dt = datetime.combine(now_ist.date(), slot.end, tzinfo=IST)
    deadline_dt = slot_end_dt + timedelta(seconds=COUNTDOWN_DURATION_SECS)
    if slot_end_dt < now_ist <= deadline_dt:
        return deadline_dt
    return None


def _build_countdown_text(
    slot_label: str,
    today_str: str,
    remaining_secs: int,
    pending: list[tuple[str, Optional[int]]],
    tick: int,
    ended_early: bool = False,
    deadline_dt: Optional[datetime] = None,
) -> str:
    if ended_early or not pending:
        return (
            f"🎉✅ <b>ALL CLEAR — {html.escape(slot_label)} ({today_str})</b>\n"
            f"All tagged assistants have submitted their report before the 15-min countdown ended!\n"
            f"👏 Great discipline!"
        )

    reg_pen, fin_pen, missed_pen = get_penalty_config()
    rem = max(0, remaining_secs)
    mins, secs = divmod(rem, 60)
    total = COUNTDOWN_DURATION_SECS
    filled = max(0, min(12, round((rem / total) * 12)))
    bar = "█" * filled + "░" * (12 - filled)
    pct = round((rem / total) * 100)
    anim = _COUNTDOWN_ANIM_FRAMES[tick % len(_COUNTDOWN_ANIM_FRAMES)]
    tags_line = " ".join(_format_mention_html(u, uid) for u, uid in pending)
    cutoff_str = deadline_dt.strftime("%I:%M %p IST") if deadline_dt else "00:00"
    is_final = _is_final_slot(slot_label)

    if rem > 0:
        if is_final:
            return (
                f"{anim} <b>FINAL WARNING + 15-MIN COUNTDOWN</b> — {html.escape(slot_label)}\n\n"
                f"⏱ Time Left Until 8:00 PM: <code>{mins:02d}:{secs:02d}</code>  <code>[{bar}]</code> <b>{pct}%</b>\n"
                f"🕒 Deadline: <b>{html.escape(cutoff_str)}</b>\n\n"
                f"🟢 Before 8:00 PM: <b>₹0 (On Time)</b>\n"
                f"🟡 After 8:00 PM (Late): <b>₹{fin_pen} Penalty</b>\n"
                f"🔴 Missed Report: <b>₹{missed_pen} Penalty</b>\n\n"
                f"<b>Pending Assistants ({len(pending)}):</b>\n{tags_line}\n\n"
                f"<i>Submit before 8:00 PM IST to avoid ₹{fin_pen} / ₹{missed_pen} penalty!</i>"
            )
        return (
            f"{anim} <b>15-MINUTE LATE COUNTDOWN</b> — {html.escape(slot_label)}\n\n"
            f"⏱ Time Remaining: <code>{mins:02d}:{secs:02d}</code>  <code>[{bar}]</code> <b>{pct}%</b>\n"
            f"🕒 Grace Cutoff: <b>{html.escape(cutoff_str)}</b>\n\n"
            f"🟡 Late Submit (Timing ke baad bhi bhejein): <b>₹{reg_pen} Penalty ONLY</b>\n"
            f"🔴 Missed Slot (Nahi bheja toh): <b>₹{missed_pen} Penalty</b>\n\n"
            f"<b>Pending Assistants ({len(pending)}):</b>\n{tags_line}\n\n"
            f"<i>Abhi submit karein (₹{reg_pen} Late) warna ₹{missed_pen} Missed penalty lagegi!</i>"
        )

    if is_final:
        return (
            f"🚨💥 <b>8:00 PM LAST WARNING — ₹{fin_pen} PENALTY ACTIVE!</b>\n"
            f"Slot: <b>{html.escape(slot_label)}</b> ({today_str})\n\n"
            f"<code>[░░░░░░░░░░░░]</code> <b>0%</b>\n"
            f"🟡 Submit Late Now: <b>₹{fin_pen} Penalty</b>\n"
            f"🔴 If Missed Completely: <b>₹{missed_pen} Penalty</b>\n\n"
            f"<b>Still Pending ({len(pending)}):</b>\n{tags_line}"
        )
    return (
        f"🚨💥 <b>COUNTDOWN ENDED (00:00) — ₹{missed_pen} MISSED PENALTY!</b>\n"
        f"Slot: <b>{html.escape(slot_label)}</b> ({today_str})\n\n"
        f"<code>[░░░░░░░░░░░░]</code> <b>0%</b>\n"
        f"🔴 Unsubmitted reports are now <b>Missed (₹{missed_pen} Penalty)</b>. Late submit still saves you to <b>₹{reg_pen}</b>.\n\n"
        f"<b>Still Pending ({len(pending)}):</b>\n{tags_line}"
    )


async def _cleanup_old_reminder_messages(bot, chat_id: int) -> None:
    if not is_clean_chat_enabled():
        return
    key = f"clean_rem_msgs:{chat_id}"
    raw = await asyncio.to_thread(get_setting, key, "[]")
    try:
        msg_ids = json.loads(raw or "[]")
    except Exception:
        msg_ids = []
    if not isinstance(msg_ids, list) or not msg_ids:
        return
    for mid in msg_ids:
        try:
            await bot.delete_message(chat_id=chat_id, message_id=int(mid))
        except Exception:
            pass
    await asyncio.to_thread(set_setting, key, "[]")


async def _record_reminder_message(chat_id: int, message_id: int) -> None:
    key = f"clean_rem_msgs:{chat_id}"
    raw = await asyncio.to_thread(get_setting, key, "[]")
    try:
        msg_ids = json.loads(raw or "[]")
        if not isinstance(msg_ids, list):
            msg_ids = []
    except Exception:
        msg_ids = []
    msg_ids.append(int(message_id))
    await asyncio.to_thread(set_setting, key, json.dumps(msg_ids[-15:]))


async def send_dm_to_assistant(
    bot, uname: str, uid: Optional[int], html_text: str
) -> tuple[bool, str]:
    """
    Strictly send Bot API DM ONLY to non-admin Assistants who have started the bot (/start).
    Group Admins are always skipped.
    """
    if is_known_group_admin(uname, uid):
        return False, "skipped_admin"

    if uid is None:
        return False, "needs_start_no_user_id"

    try:
        await bot.send_message(chat_id=uid, text=html_text, parse_mode="HTML")
        await asyncio.to_thread(mark_user_dm_started, uid, uname)
        return True, "bot_api"
    except Exception as exc:
        err_str = str(exc)
        low = err_str.lower()
        if "can't initiate conversation" in low or "bot was blocked" in low or "forbidden" in low:
            return False, "needs_start"
        return False, err_str


async def _get_bot_username(bot) -> str:
    global _BOT_USERNAME_CACHE
    if _BOT_USERNAME_CACHE:
        return _BOT_USERNAME_CACHE
    try:
        me = await bot.get_me()
        if me and me.username:
            _BOT_USERNAME_CACHE = me.username
            return me.username
    except Exception:
        pass
    return ""


async def notify_pending_via_dm(
    bot, slot_label: str, pending: list[tuple[str, Optional[int]]], stage: str = "countdown"
) -> None:
    """Send urgent personal DM via Bot API ONLY to pending non-admin assistants who have clicked /start."""
    if not is_reminders_enabled():
        return
    reg_pen, fin_pen, missed_pen = get_penalty_config()
    late_amt = fin_pen if _is_final_slot(slot_label) else reg_pen
    for uname, uid in pending:
        if is_known_group_admin(uname, uid):
            continue
        if stage == "open":
            dm_text = (
                f"📢 <b>{html.escape(slot_label)} Report Open!</b>\n"
                f"Hi {html.escape(uname)}, apni report (Username + Joining) group mein submit karein.\n"
                f"On-Time rehne par Penalty: <b>₹0</b> (Late: ₹{late_amt} | Missed: ₹{missed_pen})."
            )
        elif stage == "5min":
            dm_text = (
                f"⚡ <b>5 Minutes Left — {html.escape(slot_label)}</b>\n"
                f"Hi {html.escape(uname)}, ₹0 penalty ke liye sirf <b>5 minutes</b> bache hain!\n"
                f"Jaldi group mein report send karein."
            )
        elif stage == "8pm_final":
            dm_text = (
                f"🚨 <b>8:00 PM LAST WARNING — Final Report</b>\n"
                f"Hi {html.escape(uname)}, On-Time window close ho gaya hai!\n"
                f"Abhi submit karne par Late Penalty <b>₹{late_amt}</b> hai (Miss hone par <b>₹{missed_pen}</b>)."
            )
        else:
            dm_text = (
                f"⏰ <b>Urgent Report Reminder ({html.escape(slot_label)})</b>\n"
                f"Hi {html.escape(uname)}, aapki report abhi pending hai!\n"
                f"Abhi submit karne par Late Penalty <b>₹{late_amt}</b> hai (Miss hone par <b>₹{missed_pen}</b>)."
            )
        ok, via = await send_dm_to_assistant(bot, uname, uid, dm_text)
        if not ok:
            logger.debug("Assistant DM skipped for %s: %s", uname, via)
        await asyncio.sleep(0.35)


async def handle_any_private_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Passively record user_id & username of anyone who has already started the bot in Private DM."""
    if not update.effective_chat or update.effective_chat.type != "private" or not update.effective_user:
        return
    user = update.effective_user
    if getattr(user, "is_bot", False):
        return
    uname = f"@{user.username}" if user.username else (user.first_name or "Assistant")
    await asyncio.to_thread(mark_user_dm_started, user.id, uname)
    if not _is_admin(update) and not is_known_group_admin(uname, user.id):
        await asyncio.to_thread(auto_discover_assistant, uname, user.id)


async def dmstatus_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only admin can view assistant DM status.")
        return
    work_gid = get_work_group_id()
    await refresh_group_admins(context.bot, work_gid, force=True)
    tagged = await asyncio.to_thread(get_tagged_assistants_with_ids)
    started_profiles = await asyncio.to_thread(get_dm_started_profiles)

    seen_keys: set[str] = set()
    non_admin_assistants: list[tuple[str, Optional[int]]] = []
    for u, uid in tagged:
        if is_known_group_admin(u, uid):
            continue
        seen_keys.add(_match_key(u))
        non_admin_assistants.append((u, uid))
    for uid, u in started_profiles.items():
        if DEFAULT_ADMIN_CHAT_ID is not None and uid == DEFAULT_ADMIN_CHAT_ID:
            continue
        if is_known_group_admin(u, uid):
            continue
        if _match_key(u) not in seen_keys:
            seen_keys.add(_match_key(u))
            non_admin_assistants.append((u, uid))

    lines = [
        "<b>Assistant DM Delivery Status (/start Only | Admins Strictly Excluded)</b>",
        f"Active Non-Admin Assistants: <b>{len(non_admin_assistants)}</b>\n",
    ]
    for uname, uid in non_admin_assistants:
        if uid is not None and uid in started_profiles:
            status = "✅ Bot DM Ready (/start done)"
        else:
            status = "⏳ Needs /start (Tap 1-Tap DM button in group countdown)"
        lines.append(f"- {html.escape(uname)}: {status}")
    if not non_admin_assistants:
        lines.append("No active non-admin assistants yet.")
    lines.append("\nCommands: <code>/dmtest @assistant</code> | <code>/dmall [message]</code>")
    await update.message.reply_text("\n".join(lines), parse_mode="HTML")


async def dmall_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Admin command /dmall [optional message] to immediately DM all non-admin assistants who already clicked /start."""
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only admin can send DMs to started assistants.")
        return

    work_gid = get_work_group_id()
    await refresh_group_admins(context.bot, work_gid, force=True)
    started_profiles = await asyncio.to_thread(get_dm_started_profiles)
    tagged = await asyncio.to_thread(get_tagged_assistants_with_ids)

    candidates: dict[int, str] = {}
    for uid, uname in started_profiles.items():
        if DEFAULT_ADMIN_CHAT_ID is not None and uid == DEFAULT_ADMIN_CHAT_ID:
            continue
        if is_known_group_admin(uname, uid):
            continue
        candidates[uid] = uname
    for uname, uid in tagged:
        if uid is not None and uid not in candidates and not is_known_group_admin(uname, uid):
            candidates[uid] = uname

    if not candidates:
        await update.message.reply_text(
            "Abhi tak kisi bhi non-admin Assistant ka /start user ID record nahi hua hai.\n"
            "Group countdown ke 1-Tap button ya <code>/start</code> se activate hote hi unko DM chala jayega.",
            parse_mode="HTML",
        )
        return

    custom_msg = " ".join(context.args).strip()
    now_ist = datetime.now(IST)
    today_str = now_ist.strftime("%Y-%m-%d")
    active_slot = detect_active_slot(now_ist)

    delivered: list[str] = []
    skipped_unstarted: list[str] = []
    for uid, uname in candidates.items():
        if custom_msg:
            dm_text = f"📢 <b>Admin Reminder for {html.escape(uname)}</b>\n{html.escape(custom_msg)}"
        elif active_slot:
            dm_text = (
                f"⏰ <b>Report Reminder ({html.escape(active_slot)})</b>\n"
                f"Hi {html.escape(uname)}, apni report (Username + Joining) group mein submit kar dein."
            )
        else:
            today_slots = await asyncio.to_thread(fetch_user_today_slots, uname, today_str)
            dm_text = (
                f"🔔 <b>Personal DM Alerts Active ({today_str})</b>\n"
                f"Hi {html.escape(uname)}, aaj aapki <b>{len(today_slots)}/5</b> slot reports recorded hain.\n"
                f"Agle slots par aapko direct DM alert time par milega!"
            )
        ok, _why = await send_dm_to_assistant(context.bot, uname, uid, dm_text)
        if ok:
            delivered.append(uname)
        else:
            skipped_unstarted.append(uname)
        await asyncio.sleep(0.25)

    lines = [
        f"📨 <b>DM Broadcast Report (/start Assistants Only)</b>",
        f"✅ Delivered ({len(delivered)}): {', '.join(html.escape(u) for u in delivered) if delivered else 'None'}",
    ]
    if skipped_unstarted:
        lines.append(f"⏳ Not Started Yet ({len(skipped_unstarted)}): {', '.join(html.escape(u) for u in skipped_unstarted)}")
    await update.message.reply_text("\n".join(lines), parse_mode="HTML")


async def dmtest_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only admin can test assistant DMs.")
        return
    if not context.args:
        await update.message.reply_text("Usage: <code>/dmtest @assistant [optional message]</code>", parse_mode="HTML")
        return
    raw_u = context.args[0]
    custom_msg = " ".join(context.args[1:]).strip() or "🔔 Test Urgent Reminder: Report submission DM alerts active hain!"
    with sqlite3.connect(DB_PATH) as conn:
        canon = _resolve_canonical_name(conn, raw_u) or _normalize_tag(raw_u) or raw_u
        row = conn.execute(
            "SELECT user_id FROM assistant_tags WHERE LOWER(LTRIM(username, '@')) = ? UNION SELECT user_id FROM team_members WHERE LOWER(LTRIM(username, '@')) = ? LIMIT 1",
            (_match_key(canon), _match_key(canon)),
        ).fetchone()
    uid = int(row[0]) if (row and row[0] is not None) else None
    if uid is None:
        for p_uid, p_uname in get_dm_started_profiles().items():
            if _match_key(p_uname) == _match_key(canon):
                uid = p_uid
                break
    if is_known_group_admin(canon, uid):
        await update.message.reply_text("🚫 Rejected: Ye user Group Admin hai — DM sirf non-admin Assistants ko jaata hai!")
        return

    ok, detail = await send_dm_to_assistant(context.bot, canon, uid, f"👋 Hi {html.escape(canon)},\n{html.escape(custom_msg)}")
    if ok:
        await update.message.reply_text(
            f"✅ Test DM delivered to <code>{html.escape(canon)}</code> via <b>Bot API DM (/start)</b>!",
            parse_mode="HTML",
        )
    else:
        await update.message.reply_text(
            f"⚠️ Could not send DM to <code>{html.escape(canon)}</code>.\nReason: <code>{html.escape(detail)}</code>\n\n"
            f"<i>Fix: Ask the assistant to open the bot and press <code>/start</code> (or tap 'Enable Personal DM Alerts' on the group countdown).</i>",
            parse_mode="HTML",
        )


async def _run_countdown_animation(
    bot,
    chat_id: int,
    message_id: int,
    slot_label: str,
    today_str: str,
    end_dt: datetime,
    reply_markup: Optional[InlineKeyboardMarkup] = None,
) -> None:
    key = (chat_id, slot_label)
    tick = 1
    try:
        while True:
            await asyncio.sleep(COUNTDOWN_REFRESH_SECS)
            now_ist = datetime.now(IST)
            remaining = int((end_dt - now_ist).total_seconds())
            pending = await asyncio.to_thread(
                get_pending_tagged_assistants_for_slot, today_str, slot_label
            )

            if not pending:
                text = _build_countdown_text(
                    slot_label, today_str, 0, [], tick, ended_early=True, deadline_dt=end_dt
                )
                try:
                    await bot.edit_message_text(
                        chat_id=chat_id,
                        message_id=message_id,
                        text=text,
                        reply_markup=None,
                        parse_mode="HTML",
                    )
                except Exception:
                    pass
                break

            text = _build_countdown_text(
                slot_label, today_str, max(0, remaining), pending, tick, deadline_dt=end_dt
            )
            try:
                await bot.edit_message_text(
                    chat_id=chat_id,
                    message_id=message_id,
                    text=text,
                    reply_markup=reply_markup,
                    parse_mode="HTML",
                )
            except Exception as exc:
                err_msg = str(exc).lower()
                if "message to edit not found" in err_msg or "chat not found" in err_msg:
                    break
            tick += 1
            if remaining <= 0:
                break
    except asyncio.CancelledError:
        pass
    finally:
        _ACTIVE_COUNTDOWNS.pop(key, None)
        _COUNTDOWN_TASKS.pop(key, None)


async def refresh_active_countdown(bot, chat_id: int, slot_label: str) -> bool:
    """Immediately update an active countdown message when a pending assistant submits a report.
    Returns True if an active countdown was running and updated."""
    key = (chat_id, slot_label)
    state = _ACTIVE_COUNTDOWNS.get(key)
    if not state:
        return False
    msg_id = int(state["message_id"])
    today_str = str(state["today_str"])
    end_dt = state["end_dt"]
    reply_markup = state.get("reply_markup")
    if not isinstance(end_dt, datetime):
        return False
    remaining = max(0, int((end_dt - datetime.now(IST)).total_seconds()))
    pending = await asyncio.to_thread(
        get_pending_tagged_assistants_for_slot, today_str, slot_label
    )
    text = _build_countdown_text(
        slot_label,
        today_str,
        remaining,
        pending,
        tick=remaining,
        ended_early=(len(pending) == 0),
        deadline_dt=end_dt,
    )
    try:
        await bot.edit_message_text(
            chat_id=chat_id,
            message_id=msg_id,
            text=text,
            reply_markup=(None if not pending else reply_markup),
            parse_mode="HTML",
        )
    except Exception:
        pass
    if not pending:
        task = _COUNTDOWN_TASKS.pop(key, None)
        _ACTIVE_COUNTDOWNS.pop(key, None)
        if task and not task.done():
            task.cancel()
    return True


async def launch_penalty_countdown(
    bot, chat_id: int, slot_label: str, force_even_if_empty: bool = False, send_dm: bool = True
) -> bool:
    await refresh_group_admins(bot, chat_id, force=True)
    now_ist = datetime.now(IST)
    today_str = now_ist.strftime("%Y-%m-%d")
    is_off, _ = is_date_holiday_or_off(today_str)
    if is_off and not force_even_if_empty:
        return False
    pending = await asyncio.to_thread(
        get_pending_tagged_assistants_for_slot, today_str, slot_label
    )
    if not pending:
        if force_even_if_empty:
            await bot.send_message(
                chat_id=chat_id,
                text=f"All active tagged assistants have already submitted <b>{html.escape(slot_label)}</b> today!",
                parse_mode="HTML",
            )
        return False

    key = (chat_id, slot_label)
    old_task = _COUNTDOWN_TASKS.pop(key, None)
    if old_task and not old_task.done():
        old_task.cancel()

    anchored_deadline = _get_slot_countdown_deadline(now_ist, slot_label)
    end_dt = anchored_deadline or (now_ist + timedelta(seconds=COUNTDOWN_DURATION_SECS))
    initial_remaining = max(1, int((end_dt - now_ist).total_seconds()))

    initial_text = _build_countdown_text(
        slot_label, today_str, initial_remaining, pending, tick=0, deadline_dt=end_dt
    )

    dm_btn_markup: Optional[InlineKeyboardMarkup] = None
    started_ids = await asyncio.to_thread(get_dm_started_user_ids)
    needs_start = any((uid is None or uid not in started_ids) for _u, uid in pending)
    if needs_start:
        bot_uname = await _get_bot_username(bot)
        if bot_uname:
            dm_btn_markup = InlineKeyboardMarkup(
                [[InlineKeyboardButton("🔔 Enable Personal DM Alerts (1-Tap)", url=f"https://t.me/{bot_uname}?start=dm_enable")]]
            )

    try:
        sent = await bot.send_message(
            chat_id=chat_id, text=initial_text, reply_markup=dm_btn_markup, parse_mode="HTML"
        )
        await _record_reminder_message(chat_id, sent.message_id)
    except Exception as exc:
        logger.warning("Could not start penalty countdown in %s: %s", chat_id, exc)
        return False

    _ACTIVE_COUNTDOWNS[key] = {
        "message_id": sent.message_id,
        "end_dt": end_dt,
        "today_str": today_str,
        "reply_markup": dm_btn_markup,
    }
    _COUNTDOWN_TASKS[key] = asyncio.create_task(
        _run_countdown_animation(
            bot, chat_id, sent.message_id, slot_label, today_str, end_dt, reply_markup=dm_btn_markup
        )
    )
    if send_dm:
        asyncio.create_task(notify_pending_via_dm(bot, slot_label, pending))
    return True


async def send_daily_excel_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    excel_stream, today_count = await asyncio.to_thread(build_excel_bytes, today_str, 1)
    filename = f"team_report_{today_str}.xlsx"

    if DEFAULT_ADMIN_CHAT_ID is not None:
        try:
            await context.bot.send_document(
                chat_id=DEFAULT_ADMIN_CHAT_ID,
                document=excel_stream,
                filename=filename,
                caption=(
                    f"Daily Automated Excel Report ({today_str} | 8:35 PM IST)\n"
                    f"Total slot submissions today: {today_count}"
                ),
            )
        except Exception as exc:
            logger.warning("Could not send daily Excel to Admin DM (%s): %s", DEFAULT_ADMIN_CHAT_ID, exc)

    await push_excel_to_group(
        context,
        target_date=today_str,
        caption=f"Daily Closing Excel Report | {today_str} | 8:35 PM IST | Total today: {today_count}",
    )


async def send_morning_kickoff_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    """9:45 AM IST Morning Kickoff across all work groups."""
    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    is_off, off_reason = is_date_holiday_or_off(today_str)
    work_gids = get_all_work_group_ids()
    if not work_gids:
        return
    reg_p, fin_p, miss_p = get_penalty_config()
    tgt = get_daily_joining_target()
    tgt_line = f"\n🎯 <b>Daily Joining Target:</b> {tgt}" if tgt > 0 else ""

    for cid in work_gids:
        await _cleanup_old_reminder_messages(context.bot, cid)
        if is_off:
            msg = f"🌴 <b>Good Morning Team! ({today_str})</b>\nToday is marked as <b>{html.escape(off_reason)}</b>. Slot reminders &amp; missed penalties are paused!"
        else:
            msg = (
                f"☀️ <b>GOOD MORNING TEAM! ({today_str} IST)</b>\n"
                f"Reporting schedule for today:\n"
                f"• 10:00 AM | 1:00 PM | 3:00 PM | 5:00 PM (15 min grace)\n"
                f"• Final Report: 7:15 PM – 8:00 PM IST (7:45 PM Countdown){tgt_line}\n\n"
                f"⚖️ Penalties: Regular Late <b>₹{reg_p}</b> | Final Late <b>₹{fin_p}</b> | Missed <b>₹{miss_p}</b>\n"
                f"<i>Note: Daily Joining counts your latest cumulative joining number!</i>"
            )
        try:
            sent = await context.bot.send_message(chat_id=cid, text=msg, parse_mode="HTML")
            await _record_reminder_message(cid, sent.message_id)
        except Exception as exc:
            logger.warning("Morning kickoff failed for %s: %s", cid, exc)


async def send_daily_penalty_summary_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    """8:30 PM IST Auto Daily Penalty & Missed Summary."""
    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    text = await asyncio.to_thread(_build_daily_penalty_text, today_str)
    targets = get_all_work_group_ids() or ([DEFAULT_ADMIN_CHAT_ID] if DEFAULT_ADMIN_CHAT_ID else [])
    for cid in targets:
        try:
            await context.bot.send_message(chat_id=cid, text=text, parse_mode="HTML")
        except Exception as exc:
            logger.warning("Daily penalty summary failed for %s: %s", cid, exc)


async def send_daily_leaderboard_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    """8:45 PM IST Auto Daily Leaderboard."""
    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    is_off, _ = is_date_holiday_or_off(today_str)
    if is_off:
        return
    text = await asyncio.to_thread(_build_leaderboard_text, today_str)
    for cid in get_all_work_group_ids():
        try:
            await context.bot.send_message(chat_id=cid, text=text, parse_mode="HTML")
        except Exception as exc:
            logger.warning("Daily leaderboard failed for %s: %s", cid, exc)


async def send_nightly_backup_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    """9:15 PM IST Silent .db + .xlsx Backup to Admin DM & Excel Group."""
    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    if DEFAULT_ADMIN_CHAT_ID is not None and os.path.exists(DB_PATH):
        try:
            with open(DB_PATH, "rb") as f:
                db_bytes = BytesIO(f.read())
            db_bytes.name = "reports.db"
            await context.bot.send_document(
                chat_id=DEFAULT_ADMIN_CHAT_ID,
                document=db_bytes,
                filename=f"reports_{today_str}.db",
                caption=f"Nightly SQLite Database Backup ({today_str} | 9:15 PM IST)",
                disable_notification=True,
            )
        except Exception as exc:
            logger.warning("Nightly DB backup to Admin DM failed: %s", exc)
    await push_excel_to_group(
        context,
        target_date=today_str,
        caption=f"Nightly Archived Excel Backup | {today_str} | 9:15 PM IST",
    )


async def send_slot_open_reminder_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_reminders_enabled() or not context.job:
        return

    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    is_off, _ = is_date_holiday_or_off(today_str)
    if is_off:
        return

    work_gids = get_all_work_group_ids()
    if not work_gids:
        return

    slot_label = str(context.job.data)
    reg_pen, fin_pen, missed_pen = get_penalty_config()
    tagged_pairs = await asyncio.to_thread(get_pending_tagged_assistants_for_slot, today_str, slot_label)
    tags_line = " ".join(_format_mention_html(u, uid) for u, uid in tagged_pairs)
    attention_block = f"Attention: {tags_line}\n" if tags_line else ""

    if _is_final_slot(slot_label):
        msg = (
            f"<b>WARNING &amp; REMINDER:</b> {html.escape(slot_label)} is now OPEN (7:15 PM IST)!\n"
            f"{attention_block}"
            f"Submit your report (Username + Joining) before <b>8:00 PM IST</b> for <b>On Time (₹0)</b>.\n"
            f"7:45 PM: Final Warning + 15-Min Countdown | After 8:00 PM: <b>₹{fin_pen} Late</b> | Missed: <b>₹{missed_pen}</b>."
        )
    else:
        msg = (
            f"<b>WARNING &amp; REMINDER:</b> {html.escape(slot_label)} is now OPEN!\n"
            f"{attention_block}"
            f"Submit within 15 minutes for <b>On Time (₹0)</b>.\n"
            f"After 15 mins (Late submit): <b>₹{reg_pen} Penalty</b> | If Missed: <b>₹{missed_pen} Penalty</b>."
        )

    for work_group_id in work_gids:
        await refresh_group_admins(context.bot, work_group_id)
        await _cleanup_old_reminder_messages(context.bot, work_group_id)
        try:
            sent = await context.bot.send_message(chat_id=work_group_id, text=msg, parse_mode="HTML")
            await _record_reminder_message(work_group_id, sent.message_id)
        except Exception as exc:
            logger.warning("Could not send slot-open reminder to %s: %s", work_group_id, exc)
    if tagged_pairs:
        asyncio.create_task(notify_pending_via_dm(context.bot, slot_label, tagged_pairs, stage="open"))


async def send_slot_final_warning_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_reminders_enabled() or not context.job:
        return

    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    is_off, _ = is_date_holiday_or_off(today_str)
    if is_off:
        return

    work_gids = get_all_work_group_ids()
    if not work_gids:
        return

    slot_label = str(context.job.data)
    if _is_final_slot(slot_label):
        for wg_id in work_gids:
            await launch_penalty_countdown(context.bot, wg_id, slot_label)
        return

    pending = await asyncio.to_thread(
        get_pending_tagged_assistants_for_slot, today_str, slot_label
    )
    if not pending:
        return

    reg_pen, _fin_pen, missed_pen = get_penalty_config()
    tags_line = " ".join(_format_mention_html(u, uid) for u, uid in pending)
    msg = (
        f"<b>5 MINUTES LEFT (ON-TIME WINDOW)</b> — {html.escape(slot_label)}\n"
        f"Pending Assistants: {tags_line}\n"
        f"Submit in 5 mins for ₹0! After 15 mins: Late = <b>₹{reg_pen}</b> | Missed = <b>₹{missed_pen}</b>."
    )
    for work_group_id in work_gids:
        await refresh_group_admins(context.bot, work_group_id)
        try:
            sent = await context.bot.send_message(chat_id=work_group_id, text=msg, parse_mode="HTML")
            await _record_reminder_message(work_group_id, sent.message_id)
        except Exception as exc:
            logger.warning("Could not send 5-min final warning to %s: %s", work_group_id, exc)
    asyncio.create_task(notify_pending_via_dm(context.bot, slot_label, pending, stage="5min"))


async def send_final_8pm_last_warning_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    """Fires at 8:00 PM IST to notify any remaining pending assistants that Final Late Penalty is now active."""
    if not is_reminders_enabled():
        return
    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    is_off, _ = is_date_holiday_or_off(today_str)
    if is_off:
        return
    work_gids = get_all_work_group_ids()
    if not work_gids:
        return
    pending = await asyncio.to_thread(
        get_pending_tagged_assistants_for_slot, today_str, FINAL_SLOT_LABEL
    )
    if not pending:
        return
    _reg_pen, fin_pen, missed_pen = get_penalty_config()
    tags_line = " ".join(_format_mention_html(u, uid) for u, uid in pending)
    msg = (
        f"🚨 <b>8:00 PM LAST WARNING — FINAL REPORT</b>\n"
        f"Pending Assistants: {tags_line}\n"
        f"On-Time window closed! Submit late now for <b>₹{fin_pen} Penalty</b> — "
        f"if missed, <b>₹{missed_pen} Penalty</b> applies."
    )
    for work_group_id in work_gids:
        await refresh_group_admins(context.bot, work_group_id)
        try:
            sent = await context.bot.send_message(chat_id=work_group_id, text=msg, parse_mode="HTML")
            await _record_reminder_message(work_group_id, sent.message_id)
        except Exception as exc:
            logger.warning("Could not send 8:00 PM last warning to %s: %s", work_group_id, exc)
    asyncio.create_task(notify_pending_via_dm(context.bot, FINAL_SLOT_LABEL, pending, stage="8pm_final"))


async def start_penalty_countdown_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_reminders_enabled() or not context.job:
        return
    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    is_off, _ = is_date_holiday_or_off(today_str)
    if is_off:
        return
    slot_label = str(context.job.data)
    for wg_id in get_all_work_group_ids():
        await launch_penalty_countdown(context.bot, wg_id, slot_label)


async def countdown_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Admin command /countdown [10am|1pm|3pm|5pm|final|stop] to trigger or stop the 15-min animated countdown."""
    if not update.message or not update.effective_chat:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can control the penalty countdown.")
        return

    work_gid = get_work_group_id()
    target_cid = (
        update.effective_chat.id
        if update.effective_chat.id < 0
        else (work_gid or update.effective_chat.id)
    )

    arg = context.args[0].strip().lower() if context.args else ""
    if arg == "stop":
        stopped = 0
        for key in list(_COUNTDOWN_TASKS.keys()):
            if key[0] == target_cid:
                t = _COUNTDOWN_TASKS.pop(key, None)
                _ACTIVE_COUNTDOWNS.pop(key, None)
                if t and not t.done():
                    t.cancel()
                stopped += 1
        await update.message.reply_text(
            f"Stopped {stopped} active countdown(s)." if stopped else "No active countdown running."
        )
        return

    ist_now = datetime.now(IST)
    slot_label = parse_slot_alias(arg) if arg else None
    if slot_label is None:
        slot_label = resolve_submission_slot_and_penalty(ist_now)[0] or REPORTING_SLOTS[0].label

    started = await launch_penalty_countdown(
        context.bot, target_cid, slot_label, force_even_if_empty=True
    )
    if started and update.effective_chat.id != target_cid:
        await update.message.reply_text(
            f"Started 15-min animated penalty countdown for <b>{html.escape(slot_label)}</b> in the Work Group.",
            parse_mode="HTML",
        )


# ---------------------------------------------------------------------------
# Group Activity: Auto-Discover Assistants (Runs in group=-1 on ALL updates)
# ---------------------------------------------------------------------------
def _extract_username(update: Update) -> Optional[str]:
    user = update.effective_user
    if not user or user.is_bot:
        return None
    return f"@{user.username}" if user.username else (user.first_name or "Unknown User")


async def refresh_group_admins(
    bot, chat_id: Optional[int] = None, force: bool = False
) -> None:
    """Fetch Telegram Group Admins & purge them from assistant tags/roster so only non-admin Assistants are tagged."""
    target_cid = chat_id if (chat_id is not None and chat_id < 0) else get_work_group_id()
    if target_cid is None or target_cid >= 0:
        return

    now_ist = datetime.now(IST)
    last_sync = _ADMIN_SYNC_CACHE.get(target_cid)
    if not force and last_sync is not None and (now_ist - last_sync).total_seconds() < 120:
        return

    _ADMIN_SYNC_CACHE[target_cid] = now_ist
    try:
        admins = await bot.get_chat_administrators(target_cid)
    except Exception as exc:
        logger.debug("Could not fetch chat administrators for %s: %s", target_cid, exc)
        return

    pairs: list[tuple[Optional[int], str]] = []
    seen_ids: set[int] = set()
    for member in admins:
        u = member.user
        if not u or u.is_bot:
            continue
        uname = f"@{u.username}" if u.username else (u.first_name or f"Admin({u.id})")
        pairs.append((u.id, uname))
        seen_ids.add(u.id)

    if DEFAULT_ADMIN_CHAT_ID is not None and DEFAULT_ADMIN_CHAT_ID not in seen_ids:
        pairs.append((DEFAULT_ADMIN_CHAT_ID, ""))

    purged = await asyncio.to_thread(sync_group_admins, pairs)
    if purged > 0:
        logger.info("Purged %s Group Admin entries from assistant tags/roster.", purged)


async def send_celebration_reaction(
    context: ContextTypes.DEFAULT_TYPE,
    chat_id: int,
    message_id: int,
    is_late: bool = False,
) -> None:
    """
    React to valid report messages with big full-screen popup animation (is_big=True).
    Tries multiple celebratory emojis first and falls back to a single big emoji if the chat caps bot reactions at 1.
    """
    if ReactionTypeEmoji is None:
        return

    if is_late:
        try:
            await context.bot.set_message_reaction(
                chat_id=chat_id,
                message_id=message_id,
                reaction=[ReactionTypeEmoji("👍")],
                is_big=True,
            )
        except Exception:
            pass
        return

    multi = [
        ReactionTypeEmoji("🔥"),
        ReactionTypeEmoji("🎉"),
        ReactionTypeEmoji("💯"),
    ]
    try:
        await context.bot.set_message_reaction(
            chat_id=chat_id,
            message_id=message_id,
            reaction=multi,
            is_big=True,
        )
        return
    except Exception:
        pass

    chosen = _CELEBRATION_EMOJIS[message_id % len(_CELEBRATION_EMOJIS)]
    for candidate in (chosen, "🔥", "👍"):
        try:
            await context.bot.set_message_reaction(
                chat_id=chat_id,
                message_id=message_id,
                reaction=[ReactionTypeEmoji(candidate)],
                is_big=True,
            )
            return
        except Exception:
            continue


async def handle_any_group_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Runs in handler group=-1 on EVERY work-group message (new or edited, text, photo, command, join)
    so non-admin assistants are auto-recognized and Group Admins are always excluded.
    """
    message = update.message or update.edited_message
    user = update.effective_user
    chat = update.effective_chat
    if not message or not user or not chat or user.is_bot:
        return

    excel_group_id = get_excel_group_id()
    if excel_group_id is not None and chat.id == excel_group_id and not is_allowed_work_group(chat.id):
        return

    if not is_allowed_work_group(chat.id):
        return

    await refresh_group_admins(context.bot, chat.id)

    if message.new_chat_members:
        for member in message.new_chat_members:
            if member.is_bot:
                continue
            m_name = f"@{member.username}" if member.username else (member.first_name or "Unknown")
            if is_known_group_admin(m_name, member.id):
                continue
            newly = await asyncio.to_thread(auto_discover_assistant, m_name, member.id)
            if newly:
                logger.info("Auto-recognized new group member as assistant: %s", m_name)

    username = _extract_username(update)
    if is_known_group_admin(username or "", user.id):
        return

    if username:
        newly = await asyncio.to_thread(auto_discover_assistant, username, user.id)
        if newly:
            logger.info("Auto-recognized assistant from group message: %s (id=%s)", username, user.id)


async def handle_slot_report(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Record Username, Joining, Timing, Caption & Penalty for non-admin assistants
    (supports new messages AND edited captions/messages), with duplicate-photo detection,
    zero-growth/drop alerts, cumulative daily joining, and one-tap late-report admin buttons.
    """
    is_edited = update.message is None and update.edited_message is not None
    message = update.message or update.edited_message
    user = update.effective_user
    chat = update.effective_chat

    if not message or not user or not chat or user.is_bot:
        return

    if message.document:
        doc_name = (message.document.file_name or "").lower()
        if doc_name.endswith(".db") or doc_name.endswith(".sqlite") or doc_name.endswith(".xlsx"):
            return

    raw_caption = (message.caption or message.text or "").strip()

    if not is_report_caption(raw_caption):
        return

    work_group_id = get_work_group_id()
    excel_group_id = get_excel_group_id()
    if excel_group_id is not None and chat.id == excel_group_id and not is_allowed_work_group(chat.id):
        return

    if not is_allowed_work_group(chat.id):
        logger.warning("Ignored report: chat_id=%s not in allowed work groups", chat.id)
        now_ist = datetime.now(IST)
        last_warn = _MISMATCH_WARN_CACHE.get(chat.id)
        if last_warn is None or (now_ist - last_warn).total_seconds() >= 300:
            _MISMATCH_WARN_CACHE[chat.id] = now_ist
            await message.reply_text(
                f"<b>Report not counted:</b> This group ID (<code>{chat.id}</code>) is not registered as a Work Group (primary: <code>{work_group_id}</code>).\n"
                f"Admin: send <code>/addworkgroup</code> or <code>/setworkgroup</code> here right now.",
                parse_mode="HTML",
                reply_to_message_id=message.message_id,
            )
        return

    await refresh_group_admins(context.bot, chat.id)
    sender_label = _extract_username(update) or ""
    if is_known_group_admin(sender_label, user.id):
        return

    fallback_user = sender_label or "Unknown User"
    username, joining = parse_report_fields(raw_caption, fallback_user)
    effective_uid = user.id if _match_key(username) == _match_key(fallback_user) else None

    ist_dt = message.date.astimezone(IST)
    submission_time = ist_dt.strftime("%I:%M:%S %p IST")
    report_date = ist_dt.strftime("%Y-%m-%d")
    active_slot, status, penalty_rs = resolve_submission_slot_and_penalty(ist_dt)

    if active_slot is None:
        if not is_edited and message.media_group_id:
            mg_key = f"{chat.id}:{message.media_group_id}"
            if mg_key in _SEEN_MEDIA_GROUPS:
                return
            _SEEN_MEDIA_GROUPS[mg_key] = datetime.now(IST)
        if ReactionTypeEmoji is not None:
            try:
                await context.bot.set_message_reaction(
                    chat_id=chat.id,
                    message_id=message.message_id,
                    reaction=[ReactionTypeEmoji("⏳")],
                    is_big=False,
                )
            except Exception:
                pass
        return

    growth_note = ""
    async with _REPORT_LOCK:
        if not is_edited and message.media_group_id:
            mg_key = f"{chat.id}:{message.media_group_id}"
            now_ist = datetime.now(IST)
            stale = [k for k, ts in _SEEN_MEDIA_GROUPS.items() if (now_ist - ts).total_seconds() > 600]
            for k in stale:
                _SEEN_MEDIA_GROUPS.pop(k, None)
            if mg_key in _SEEN_MEDIA_GROUPS:
                return
            _SEEN_MEDIA_GROUPS[mg_key] = now_ist

        existing = await asyncio.to_thread(
            has_slot_submission, username, report_date, active_slot
        )
        should_update_missing_joining = bool(existing and existing[2] == "" and joining != "")

        if PREVENT_DUPLICATE_SLOTS and existing and not should_update_missing_joining and not is_edited:
            await message.reply_text(
                f"Duplicate skipped for {username}: {active_slot} was already recorded today at {existing[1]}.",
                reply_to_message_id=message.message_id,
            )
            return

        if message.photo:
            photo_uid = message.photo[-1].file_unique_id
            dup_photo = await asyncio.to_thread(
                check_and_record_photo, photo_uid, username, report_date, active_slot, message.message_id
            )
            if dup_photo:
                prev_u, prev_d, prev_s = dup_photo
                await message.reply_text(
                    f"⚠️ <b>Duplicate / Re-used Photo Detected!</b>\n"
                    f"Ye photo pehle hi <code>{html.escape(prev_u)}</code> ne {prev_d} ({html.escape(prev_s)}) mein upload ki thi.\n"
                    f"Please upload a fresh screenshot.",
                    parse_mode="HTML",
                    reply_to_message_id=message.message_id,
                )
                return

        prev_join_info = await asyncio.to_thread(
            get_previous_slot_joining_today, username, report_date, active_slot
        )
        new_join_int = _safe_int_joining(joining)
        if prev_join_info and new_join_int is not None:
            prev_slot_lbl, prev_val = prev_join_info
            if new_join_int < prev_val:
                growth_note = f"\n📉 <i>Alert: Joining dropped from {prev_val} ({html.escape(prev_slot_lbl)}) to {new_join_int} — please verify!</i>"
            elif new_join_int == prev_val:
                growth_note = f"\n⚠️ <i>Zero-Growth Note: Joining same as {html.escape(prev_slot_lbl)} ({new_join_int}).</i>"

        logger.info(
            "Slot report %s | user=%s | joining=%s | chat=%s | time=%s | slot=%s | status=%s | penalty=₹%s",
            "edited" if is_edited else "recorded",
            username,
            joining,
            chat.id,
            submission_time,
            active_slot,
            status,
            penalty_rs,
        )

        try:
            action = await asyncio.to_thread(
                upsert_report_to_db,
                effective_uid,
                username,
                joining,
                submission_time,
                active_slot,
                report_date,
                status,
                raw_caption,
                message.date.isoformat(),
                True,
                penalty_rs,
                chat.id,
                message.message_id,
            )
        except Exception as exc:
            logger.exception("Failed to save report to database: %s", exc)
            await message.reply_text(
                "Report received, but saving failed. Please notify an admin.",
                reply_to_message_id=message.message_id,
            )
            return

    await send_celebration_reaction(
        context, chat.id, message.message_id, is_late=(penalty_rs > 0)
    )
    was_updated = await refresh_active_countdown(context.bot, chat.id, active_slot)
    if (
        not was_updated
        and _get_slot_countdown_deadline(ist_dt, active_slot) is not None
        and is_reminders_enabled()
    ):
        await launch_penalty_countdown(context.bot, chat.id, active_slot)

    excel_caption = (
        f"Excel Updated | {username} | Joining: {joining or 'N/A'} | {active_slot}\n"
        f"Time: {submission_time} | Date: {report_date} | Status: {status} | Penalty: ₹{penalty_rs}"
    )
    if raw_caption:
        excel_caption += f"\nCaption: {raw_caption}"
    await push_excel_to_group(context, target_date=report_date, caption=excel_caption)

    join_line = f"\nValid Joining Today: <b>{html.escape(joining)}</b>" if joining else ""
    label_verb = "Updated" if action == "updated" else "Recorded"
    if penalty_rs == 0:
        header_badge = f"🎉🔥💯 🏆 ✅ <b>Report {label_verb}!</b>"
        penalty_line = "Status: <b>On Time</b> (Penalty: <b>₹0</b>)"
    else:
        header_badge = f"⏰✅ <b>Late Report {label_verb}</b>"
        penalty_line = f"Status: <b>{html.escape(status)}</b> (Penalty: <b>₹{penalty_rs}</b>)"

    confirmation = (
        f"{header_badge}\n"
        f"Assistant: <code>{html.escape(username)}</code>{join_line}\n"
        f"Slot: {html.escape(active_slot)}\n"
        f"Time: {html.escape(submission_time)} | {penalty_line}"
        f"{growth_note}"
    )

    reply_markup = None
    if penalty_rs > 0:
        rid = await asyncio.to_thread(get_report_row_id, username, report_date, active_slot)
        if rid is not None:
            reply_markup = InlineKeyboardMarkup(
                [
                    [
                        InlineKeyboardButton("✅ Mark On Time (₹0)", callback_data=f"late_ok:{rid}"),
                        InlineKeyboardButton(f"🟡 Keep ₹{penalty_rs}", callback_data=f"late_keep:{rid}"),
                        InlineKeyboardButton("🗑 Ignore", callback_data=f"late_del:{rid}"),
                    ]
                ]
            )

    await message.reply_text(
        confirmation,
        parse_mode="HTML",
        reply_markup=reply_markup,
        reply_to_message_id=message.message_id,
    )


# ---------------------------------------------------------------------------
# Lightweight Health-Check Server (Required for Render Free Web Service)
# ---------------------------------------------------------------------------
class _HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        self.send_response(200)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write(b"OK - Telegram Excel Report Bot is running.")

    def log_message(self, format: str, *args: object) -> None:
        return


def start_health_server() -> None:
    port = int(os.getenv("PORT", "10000"))
    server = HTTPServer(("0.0.0.0", port), _HealthHandler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    logger.info("Health check server listening on port %s", port)


# ---------------------------------------------------------------------------
# Startup Auto-Restore From Pinned .xlsx or .db in Database Group & Job Scheduling
# ---------------------------------------------------------------------------
async def auto_restore_from_excel_group(application: Application) -> None:
    excel_group_id = get_excel_group_id()
    if excel_group_id is None:
        return

    if await asyncio.to_thread(count_total_reports) > 0:
        return

    try:
        chat = await application.bot.get_chat(excel_group_id)
        pinned = chat.pinned_message
        if pinned and pinned.document:
            fname = pinned.document.file_name or ""
            lower_fname = fname.lower()
            if lower_fname.endswith(".xlsx") or lower_fname.endswith(".db") or lower_fname.endswith(".sqlite"):
                tg_file = await pinned.document.get_file()
                buf = BytesIO()
                await tg_file.download_to_memory(out=buf)
                res = await asyncio.to_thread(restore_any_backup_bytes, buf.getvalue(), fname)
                logger.info("Auto-restored from pinned %s: %s", fname, res)
    except Exception as exc:
        logger.warning("Database/Excel auto-restore skipped: %s", exc)


async def post_init(application: Application) -> None:
    await auto_restore_from_excel_group(application)
    work_gids = get_all_work_group_ids()
    for wg_id in work_gids:
        await refresh_group_admins(application.bot, wg_id, force=True)

    # Scoped Telegram / command menus: Assistants see ONLY their 5 commands; Admin sees full control suite
    assistant_menu = [
        BotCommand("start", "Personal DM Alerts activate karein"),
        BotCommand("mystatus", "Aaj ki apni valid joining & slots dekhein"),
        BotCommand("myfine", "Apna fine balance & history dekhein"),
        BotCommand("myrank", "Aaj ki apni rank & score dekhein"),
        BotCommand("myhelp", "Assistant commands ki guide dekhein"),
    ]
    admin_menu = [
        BotCommand("allfines", "Ek saath sabhi assistants ki penalty dekhein"),
        BotCommand("resetfine", "Sabhi ya ek assistant ka fine ₹0 karein"),
        BotCommand("summary", "Aaj ki poori team summary dekhein"),
        BotCommand("leaderboard", "Full team leaderboard dekhein"),
        BotCommand("excel", "Daily Excel report download karein"),
        BotCommand("weekly", "7-day weekly Excel report"),
        BotCommand("monthly", "Monthly Excel analytics"),
        BotCommand("payout", "Team joining salary payout sheet"),
        BotCommand("paid", "Fine payment record karein (/paid all ya @user amt)"),
        BotCommand("paidall", "Sabhi assistants ka fine ek saath PAID (₹0) karein"),
        BotCommand("waive", "Fine waive karein (/waive all ya @user amt)"),
        BotCommand("dmstatus", "Started assistants DM status dekhein"),
        BotCommand("dmall", "Sabhi started assistants ko DM bhejein"),
        BotCommand("countdown", "Slot countdown start/stop karein"),
        BotCommand("tags", "Tagged assistants manage karein"),
        BotCommand("leave", "Assistant ki leave mark karein"),
        BotCommand("restore", "Backup .db/.xlsx se data recover karein"),
        BotCommand("help", "Sabhi admin commands dekhein"),
    ]
    try:
        await application.bot.set_my_commands(assistant_menu, scope=BotCommandScopeDefault())
        await application.bot.set_my_commands(assistant_menu, scope=BotCommandScopeAllGroupChats())
        if DEFAULT_ADMIN_CHAT_ID is not None:
            await application.bot.set_my_commands(
                admin_menu, scope=BotCommandScopeChat(chat_id=DEFAULT_ADMIN_CHAT_ID)
            )
    except Exception as exc:
        logger.warning("Could not register scoped BotCommands: %s", exc)

    if application.job_queue is None:
        return

    application.job_queue.run_daily(
        send_morning_kickoff_job,
        time=MORNING_KICKOFF_TIME_IST,
        name="morning_kickoff",
    )
    application.job_queue.run_daily(
        send_daily_penalty_summary_job,
        time=DAILY_PENALTY_TIME_IST,
        name="daily_penalty_summary",
    )
    application.job_queue.run_daily(
        send_daily_excel_job,
        time=DAILY_SEND_TIME_IST,
        name="daily_excel_report",
    )
    application.job_queue.run_daily(
        send_daily_leaderboard_job,
        time=DAILY_LEADERBOARD_TIME_IST,
        name="daily_leaderboard",
    )
    application.job_queue.run_daily(
        send_nightly_backup_job,
        time=NIGHTLY_BACKUP_TIME_IST,
        name="nightly_backup",
    )

    today_ist = datetime.now(IST).date()
    for slot in REPORTING_SLOTS:
        open_dt = datetime.combine(today_ist, slot.reminder_time, tzinfo=IST)
        application.job_queue.run_daily(
            send_slot_open_reminder_job,
            time=open_dt.timetz(),
            data=slot.label,
            name=f"reminder_open_{slot.label}",
        )

        warn_dt = datetime.combine(today_ist, slot.final_warning_time, tzinfo=IST)
        application.job_queue.run_daily(
            send_slot_final_warning_job,
            time=warn_dt.timetz(),
            data=slot.label,
            name=f"reminder_warn_{slot.label}",
        )

        if _is_final_slot(slot.label):
            last_warn_dt = datetime.combine(today_ist, time(20, 0, 0), tzinfo=IST)
            application.job_queue.run_daily(
                send_final_8pm_last_warning_job,
                time=last_warn_dt.timetz(),
                data=slot.label,
                name="final_8pm_last_warning",
            )
        else:
            countdown_start_dt = datetime.combine(today_ist, slot.end, tzinfo=IST) + timedelta(seconds=1)
            application.job_queue.run_daily(
                start_penalty_countdown_job,
                time=countdown_start_dt.timetz(),
                data=slot.label,
                name=f"penalty_countdown_{slot.label}",
            )

    if work_gids and is_reminders_enabled():
        now_ist = datetime.now(IST)
        is_off, _ = is_date_holiday_or_off(now_ist.strftime("%Y-%m-%d"))
        if not is_off:
            for wg_id in work_gids:
                for slot in REPORTING_SLOTS:
                    if _get_slot_countdown_deadline(now_ist, slot.label) is not None:
                        await launch_penalty_countdown(application.bot, wg_id, slot.label)

    logger.info(
        "Scheduled Kickoff (9:45), Penalties (20:30), Excel (20:35), Leaderboard (20:45), Backup (21:15 IST), and 7:45/8:00 Final Countdown."
    )


def main() -> None:
    if not BOT_TOKEN:
        raise ValueError("TELEGRAM_BOT_TOKEN environment variable is missing.")

    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)

    init_db()
    start_health_server()

    application = (
        ApplicationBuilder()
        .token(BOT_TOKEN)
        .concurrent_updates(True)
        .post_init(post_init)
        .build()
    )

    # Passively track non-admin assistants who have already started the bot in Private DM
    application.add_handler(
        MessageHandler(filters.ChatType.PRIVATE, handle_any_private_message),
        group=-2,
    )

    # Run assistant auto-discovery in group=-1 on ALL group messages (new + edited)
    application.add_handler(
        MessageHandler(
            (filters.UpdateType.MESSAGE | filters.UpdateType.EDITED_MESSAGE) & filters.ChatType.GROUPS,
            handle_any_group_message,
        ),
        group=-1,
    )

    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(CommandHandler("myhelp", myhelp_command))
    application.add_handler(CommandHandler("myrank", myrank_command))
    application.add_handler(CommandHandler("check", check_command))
    application.add_handler(CommandHandler(["add", "record"], add_command))
    application.add_handler(CommandHandler("addreport", addreport_command))
    application.add_handler(CommandHandler("countdown", countdown_command))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler(["tags", "assistants"], tags_command))
    application.add_handler(CommandHandler("settags", settags_command))
    application.add_handler(CommandHandler(["addtag", "addassistant"], addtag_command))
    application.add_handler(CommandHandler(["removetag", "removeassistant"], removetag_command))
    application.add_handler(CommandHandler("cleartags", cleartags_command))
    application.add_handler(CallbackQueryHandler(tags_callback_handler))
    application.add_handler(CommandHandler("reminder", reminder_command))
    application.add_handler(CommandHandler("setworkgroup", setworkgroup_command))
    application.add_handler(CommandHandler("addworkgroup", addworkgroup_command))
    application.add_handler(CommandHandler("removeworkgroup", removeworkgroup_command))
    application.add_handler(CommandHandler("groups", groups_command))
    application.add_handler(CommandHandler(["setexcelgroup", "setdbgroup"], setexcelgroup_command))
    application.add_handler(CommandHandler(["excel", "report"], excel_command))
    application.add_handler(CommandHandler("weekly", weekly_command))
    application.add_handler(CommandHandler("monthly", monthly_command))
    application.add_handler(CommandHandler("payout", payout_command))
    application.add_handler(CommandHandler("setrate", setrate_command))
    application.add_handler(CommandHandler("summary", summary_command))
    application.add_handler(CommandHandler("leaderboard", leaderboard_command))
    application.add_handler(CommandHandler(["removesummary", "hidesummary"], removesummary_command))
    application.add_handler(CommandHandler(["restoresummary", "addsummary"], restoresummary_command))
    application.add_handler(CommandHandler(["mystatus", "myreport"], mystatus_command))
    application.add_handler(CommandHandler("myfine", myfine_command))
    application.add_handler(CommandHandler(["allfines", "allfine", "fines", "penalties", "allpenalty"], allfines_command))
    application.add_handler(CommandHandler("leave", leave_command))
    application.add_handler(CommandHandler("unleave", unleave_command))
    application.add_handler(CommandHandler("paid", paid_command))
    application.add_handler(CommandHandler(["paidall", "clearpaid"], paidall_command))
    application.add_handler(CommandHandler("waive", waive_command))
    application.add_handler(CommandHandler("waiveall", waiveall_command))
    application.add_handler(CommandHandler(["resetfine", "resetfines", "clearfine", "clearfines"], resetfine_command))
    application.add_handler(CommandHandler("settarget", settarget_command))
    application.add_handler(CommandHandler("alias", alias_command))
    application.add_handler(CommandHandler("aliases", aliases_command))
    application.add_handler(CommandHandler("holiday", holiday_command))
    application.add_handler(CommandHandler("offdays", offdays_command))
    application.add_handler(CommandHandler("cleanchat", cleanchat_command))
    application.add_handler(CommandHandler("setpenalty", setpenalty_command))
    application.add_handler(CommandHandler("setshift", setshift_command))
    application.add_handler(CommandHandler("find", find_command))
    application.add_handler(CommandHandler(["day", "date"], day_command))
    application.add_handler(CommandHandler("addmember", addmember_command))
    application.add_handler(CommandHandler("removemember", removemember_command))
    application.add_handler(CommandHandler("team", team_command))
    application.add_handler(CommandHandler("dmstatus", dmstatus_command))
    application.add_handler(CommandHandler("dmtest", dmtest_command))
    application.add_handler(CommandHandler(["dmall", "dmremind", "senddm"], dmall_command))
    application.add_handler(CommandHandler(["restore", "recover"], restore_command))

    # Auto-restore when Admin forwards/sends .db or .xlsx in Database Group or Admin DM
    application.add_handler(
        MessageHandler(
            filters.Document.ALL & (filters.ChatType.GROUPS | filters.ChatType.PRIVATE),
            handle_backup_document_upload,
        ),
        group=1,
    )

    group_report_filter = (
        (filters.UpdateType.MESSAGE | filters.UpdateType.EDITED_MESSAGE)
        & filters.ChatType.GROUPS
        & (filters.PHOTO | filters.Document.ALL | filters.VIDEO | (filters.TEXT & ~filters.COMMAND))
    )
    application.add_handler(MessageHandler(group_report_filter, handle_slot_report))

    logger.info("Starting Telegram Excel Report Bot (polling)...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
