"""
Telegram Group Photo Report Tracker -> Excel-Only (.xlsx) + Auto Assistant Discovery
Deployable directly via GitHub + Render Free Tier (no Google Sheets / external DB needed).

Core Behavior:
  1. Excel-Only Delivery: Sends formatted .xlsx workbook (team_report_YYYY-MM-DD.xlsx)
     to your Excel/Backup Group on every recorded slot submission AND at 8:30 PM IST to Admin DM.
  2. Auto-Recognize Assistants: When any non-admin member sends ANY message in the work group
     (or joins), the bot automatically registers them as an assistant and enables tagging
     (unless manually turned OFF via /tags or /removetag).
     Note: Make the bot Admin in your Work Group (or remove & re-add after disabling Group
     Privacy in @BotFather) so Telegram delivers all plain-text messages like "Hi" to the bot.
  3. Strict Slot-Time Recording: Photo reports are ONLY recorded during the 5 active IST slot
     windows (+15 min grace). Uploads outside slot hours are NOT logged to Excel.
  4. Safe HTML formatting + user_id mention fallback for members without @username.

Requires: python-telegram-bot[job-queue] >= 21.0, openpyxl >= 3.1.5, python-dotenv, tzdata
"""

import asyncio
import html
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
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
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
DAILY_SEND_TIME_IST = time(hour=20, minute=30, second=0, tzinfo=IST)  # 8:30 PM IST

_REPORT_LOCK = asyncio.Lock()
_SEEN_MEDIA_GROUPS: dict[str, datetime] = {}
_MISMATCH_WARN_CACHE: dict[int, datetime] = {}
_ADMIN_SYNC_CACHE: dict[int, datetime] = {}
_CELEBRATION_EMOJIS: tuple[str, ...] = ("🔥", "🎉", "💯", "🏆", "⚡", "🤩", "👏")


# ---------------------------------------------------------------------------
# Reporting Slot Definitions (IST) with 15-Minute Grace Period
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class ReportingSlot:
    label: str
    start: time
    end: time  # Inclusive end of window (includes 15-minute grace period)
    reminder_time: time  # Slot-open reminder + tag warning
    final_warning_time: time  # 5-minute warning before grace closes


REPORTING_SLOTS: tuple[ReportingSlot, ...] = (
    ReportingSlot("10:00 AM Slot", time(10, 0, 0), time(10, 15, 59), time(10, 0, 0), time(10, 10, 0)),
    ReportingSlot("1:00 PM Slot", time(13, 0, 0), time(13, 15, 59), time(13, 0, 0), time(13, 10, 0)),
    ReportingSlot("3:00 PM Slot", time(15, 0, 0), time(15, 15, 59), time(15, 0, 0), time(15, 10, 0)),
    ReportingSlot("5:00 PM Slot", time(17, 0, 0), time(17, 15, 59), time(17, 0, 0), time(17, 10, 0)),
    ReportingSlot(
        "Final Report (Before 8:00 PM)",
        time(19, 30, 0),
        time(20, 15, 59),
        time(19, 45, 0),
        time(20, 10, 0),
    ),
)


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
    if token in ("8", "8pm", "20", "1930", "final", "slot5", "5th", "night"):
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


def evaluate_slot_status(ist_dt: datetime, slot_label: str) -> str:
    """Strictly check if an IST timestamp falls within the slot's cutoff window."""
    slot = _get_slot_by_label(slot_label)
    if not slot:
        return "On Time"
    current_t = ist_dt.time().replace(tzinfo=None)
    if slot.start <= current_t <= slot.end:
        return "On Time"
    if current_t > slot.end:
        return "Late"
    start_sec = slot.start.hour * 3600 + slot.start.minute * 60 + slot.start.second
    cur_sec = current_t.hour * 3600 + current_t.minute * 60 + current_t.second
    if 0 <= (start_sec - cur_sec) <= 900:
        return "On Time"
    return "Late"


def evaluate_slot_status_from_strings(submission_time: str, slot_label: str) -> str:
    parsed_t = _parse_ist_time_str(submission_time)
    slot = _get_slot_by_label(slot_label)
    if not parsed_t or not slot:
        return "On Time"
    if slot.start <= parsed_t <= slot.end:
        return "On Time"
    if parsed_t > slot.end:
        return "Late"
    start_sec = slot.start.hour * 3600 + slot.start.minute * 60 + slot.start.second
    cur_sec = parsed_t.hour * 3600 + parsed_t.minute * 60 + parsed_t.second
    if 0 <= (start_sec - cur_sec) <= 900:
        return "On Time"
    return "Late"


def recalculate_report_statuses(conn: sqlite3.Connection) -> int:
    """Auto-fix any previously saved 'On Time' rows whose submission_time is past the slot's 15-min cutoff."""
    fixed = 0
    rows = conn.execute(
        "SELECT id, submission_time, reporting_slot, status FROM reports WHERE status = 'On Time'"
    ).fetchall()
    for row_id, sub_time, slot_label, status in rows:
        expected = evaluate_slot_status_from_strings(str(sub_time or ""), str(slot_label or ""))
        if expected != status:
            conn.execute("UPDATE reports SET status = ? WHERE id = ?", (expected, row_id))
            fixed += 1
    return fixed


_USERNAME_PATTERN = re.compile(
    r"(?:username|user\s*name|user|tg\s*id|id)\s*[:=\-–—]+\s*(@?[A-Za-z0-9_]{3,})",
    re.IGNORECASE,
)
_JOINING_PATTERN = re.compile(
    r"(?:total\s*)?(?:joining|joins|join)\s*[:=\-–—]*\s*(\d+)",
    re.IGNORECASE,
)


def parse_report_fields(raw_text: str, fallback_username: str) -> tuple[str, str]:
    text = raw_text or ""
    username = fallback_username
    user_match = _USERNAME_PATTERN.search(text)
    if user_match:
        extracted = user_match.group(1).strip()
        if extracted:
            username = extracted if extracted.startswith("@") else f"@{extracted}"

    joining = ""
    join_match = _JOINING_PATTERN.search(text)
    if join_match:
        joining = join_match.group(1).strip()

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
        if "caption" not in report_cols:
            conn.execute("ALTER TABLE reports ADD COLUMN caption TEXT DEFAULT ''")

        conn.execute(
            "INSERT OR IGNORE INTO bot_settings (key, value) VALUES (?, ?)",
            ("reminders_enabled", "true" if DEFAULT_ENABLE_REMINDERS else "false"),
        )
        if DEFAULT_WORK_CHAT_ID is not None:
            conn.execute(
                "INSERT OR IGNORE INTO bot_settings (key, value) VALUES (?, ?)",
                ("work_group_id", str(DEFAULT_WORK_CHAT_ID)),
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

        recalculate_report_statuses(conn)
        conn.commit()


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


def get_work_group_id() -> Optional[int]:
    val = get_setting("work_group_id")
    return int(val) if val else DEFAULT_WORK_CHAT_ID


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
    for table in ("team_members", "assistant_tags", "reports"):
        row = conn.execute(
            f"SELECT username FROM {table} WHERE LOWER(LTRIM(username, '@')) = ? LIMIT 1",
            (key,),
        ).fetchone()
        if row and row[0]:
            return str(row[0]).strip()
    return _normalize_tag(clean) if clean.startswith("@") else clean


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


def auto_discover_assistant(username: str, user_id: Optional[int] = None) -> bool:
    """
    Automatically register any non-admin group member as an assistant on their first message.
    Blocks Group Admins and respects summary_excluded so removed members stay excluded.
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
        conn.execute(
            "INSERT OR IGNORE INTO team_members (username, user_id, added_at) VALUES (?, ?, ?)",
            (clean, user_id, now_iso),
        )
        cur = conn.execute(
            "INSERT OR IGNORE INTO assistant_tags (username, user_id, is_active, added_at) VALUES (?, ?, 1, ?)",
            (clean, user_id, now_iso),
        )
        if user_id is not None:
            conn.execute(
                "UPDATE team_members SET user_id = ? WHERE username = ? AND user_id IS NULL",
                (user_id, clean),
            )
            conn.execute(
                "UPDATE assistant_tags SET user_id = ? WHERE username = ? AND user_id IS NULL",
                (user_id, clean),
            )
        conn.commit()
        return cur.rowcount > 0


def add_team_member(username: str) -> bool:
    clean = username.strip()
    key = _match_key(clean)
    if not clean or not key:
        return False
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            "DELETE FROM summary_excluded WHERE LOWER(LTRIM(username, '@')) = ?",
            (key,),
        )
        cur = conn.execute(
            "INSERT OR IGNORE INTO team_members (username, user_id, added_at) VALUES (?, NULL, ?)",
            (clean, datetime.now(IST).isoformat()),
        )
        conn.commit()
        return cur.rowcount > 0


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
    return [
        u
        for u, uid in rows
        if (uid is None or uid not in admin_ids) and _match_key(u) not in admin_keys
    ]


def get_tagged_assistants_with_ids() -> list[tuple[str, Optional[int]]]:
    admin_ids, admin_keys = get_known_admin_sets()
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            "SELECT username, user_id FROM assistant_tags WHERE is_active = 1 ORDER BY username ASC"
        ).fetchall()
    return [
        (u, uid)
        for u, uid in rows
        if (uid is None or uid not in admin_ids) and _match_key(u) not in admin_keys
    ]


def get_tagged_assistants() -> list[str]:
    return [u for u, _uid in get_tagged_assistants_with_ids()]


def add_assistant_tag(username: str) -> bool:
    norm = _normalize_tag(username)
    if not norm:
        return False
    now_iso = datetime.now(IST).isoformat()
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            INSERT INTO assistant_tags (username, user_id, is_active, added_at) VALUES (?, NULL, 1, ?)
            ON CONFLICT(username) DO UPDATE SET is_active = 1
            """,
            (norm, now_iso),
        )
        conn.execute(
            "INSERT OR IGNORE INTO team_members (username, user_id, added_at) VALUES (?, NULL, ?)",
            (norm, now_iso),
        )
        conn.commit()
        return True


def remove_assistant_tag(username: str) -> bool:
    norm = _normalize_tag(username)
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute(
            """
            UPDATE assistant_tags SET is_active = 0
            WHERE LOWER(username) = LOWER(?) OR LOWER(username) = LOWER(?)
            """,
            (username.strip(), norm),
        )
        conn.commit()
        return cur.rowcount > 0


def replace_assistant_tags(usernames: list[str]) -> list[str]:
    now_iso = datetime.now(IST).isoformat()
    cleaned: list[str] = []
    seen_lower: set[str] = set()
    for raw in usernames:
        norm = _normalize_tag(raw)
        if norm and norm.lower() not in seen_lower:
            seen_lower.add(norm.lower())
            cleaned.append(norm)

    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("UPDATE assistant_tags SET is_active = 0")
        for norm in cleaned:
            conn.execute(
                """
                INSERT INTO assistant_tags (username, user_id, is_active, added_at) VALUES (?, NULL, 1, ?)
                ON CONFLICT(username) DO UPDATE SET is_active = 1
                """,
                (norm, now_iso),
            )
            conn.execute(
                "INSERT OR IGNORE INTO team_members (username, user_id, added_at) VALUES (?, NULL, ?)",
                (norm, now_iso),
            )
        conn.commit()
    return cleaned


def clear_assistant_tags() -> None:
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("UPDATE assistant_tags SET is_active = 0")
        conn.commit()


def toggle_assistant_tag(username: str) -> bool:
    tagged_lower = {u.lower() for u in get_tagged_assistants()}
    if username.lower() in tagged_lower:
        remove_assistant_tag(username)
        return False
    add_assistant_tag(username)
    return True


def has_slot_submission(
    username: str, report_date: str, reporting_slot: str
) -> Optional[tuple[int, str, str]]:
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            """
            SELECT id, submission_time, COALESCE(joining, '') FROM reports
            WHERE LOWER(username) = LOWER(?) AND report_date = ? AND reporting_slot = ?
            LIMIT 1
            """,
            (username, report_date, reporting_slot),
        ).fetchone()
        return (int(row[0]), str(row[1]), str(row[2])) if row else None


def get_pending_tagged_assistants_for_slot(
    report_date: str, slot_label: str
) -> list[tuple[str, Optional[int]]]:
    tagged = get_tagged_assistants_with_ids()
    if not tagged:
        return []
    with sqlite3.connect(DB_PATH) as conn:
        submitted_rows = conn.execute(
            """
            SELECT DISTINCT LOWER(username) FROM reports
            WHERE report_date = ? AND reporting_slot = ?
            """,
            (report_date, slot_label),
        ).fetchall()
    submitted_set = {r[0] for r in submitted_rows}
    return [(u, uid) for u, uid in tagged if u.lower() not in submitted_set]


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
) -> str:
    with sqlite3.connect(DB_PATH) as conn:
        existing = conn.execute(
            """
            SELECT id, COALESCE(joining, ''), COALESCE(caption, '') FROM reports
            WHERE LOWER(username) = LOWER(?) AND report_date = ? AND reporting_slot = ?
            LIMIT 1
            """,
            (username, report_date, reporting_slot),
        ).fetchone()

        if existing:
            row_id = existing[0]
            if overwrite_existing:
                new_joining = joining if joining != "" else existing[1]
                new_caption = caption if caption else existing[2]
                conn.execute(
                    """
                    UPDATE reports
                    SET user_id = COALESCE(?, user_id),
                        joining = ?,
                        submission_time = ?,
                        status = ?,
                        caption = ?
                    WHERE id = ?
                    """,
                    (user_id, new_joining, submission_time, status, new_caption, row_id),
                )
                conn.commit()
                action = "updated"
            else:
                action = "skipped"
        else:
            conn.execute(
                """
                INSERT INTO reports (
                    user_id, username, joining, submission_time, reporting_slot, report_date, status, caption, created_at_utc
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    user_id,
                    username,
                    joining,
                    submission_time,
                    reporting_slot,
                    report_date,
                    status,
                    caption,
                    created_at_utc,
                ),
            )
            conn.commit()
            action = "inserted"

    if not is_known_group_admin(username, user_id):
        auto_discover_assistant(username, user_id=user_id)
    return action


def fetch_reports(
    date_filter: Optional[str] = None,
    start_date: Optional[str] = None,
) -> list[tuple[str, str, str, str, str, str, str]]:
    with sqlite3.connect(DB_PATH) as conn:
        if date_filter:
            cursor = conn.execute(
                """
                SELECT username, COALESCE(joining, ''), submission_time, reporting_slot, report_date, status, COALESCE(caption, '')
                FROM reports
                WHERE report_date = ?
                ORDER BY id ASC
                """,
                (date_filter,),
            )
        elif start_date:
            cursor = conn.execute(
                """
                SELECT username, COALESCE(joining, ''), submission_time, reporting_slot, report_date, status, COALESCE(caption, '')
                FROM reports
                WHERE report_date >= ?
                ORDER BY report_date ASC, id ASC
                """,
                (start_date,),
            )
        else:
            cursor = conn.execute(
                """
                SELECT username, COALESCE(joining, ''), submission_time, reporting_slot, report_date, status, COALESCE(caption, '')
                FROM reports
                ORDER BY id ASC
                """
            )
        return cursor.fetchall()


def fetch_user_today_slots(username: str, report_date: str) -> dict[str, tuple[str, str]]:
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            """
            SELECT reporting_slot, submission_time, COALESCE(joining, '')
            FROM reports
            WHERE LOWER(username) = LOWER(?) AND report_date = ?
            """,
            (username, report_date),
        ).fetchall()
    return {slot: (sub_time, joining) for slot, sub_time, joining in rows}


def count_total_reports() -> int:
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute("SELECT COUNT(*) FROM reports").fetchone()
        return int(row[0]) if row else 0


# ---------------------------------------------------------------------------
# Styled Excel (.xlsx) Workbook Builder (4 Sheets) & Restore Parser
# ---------------------------------------------------------------------------
def _style_sheet(
    ws, headers: list[str], rows: list[tuple[str, ...]], status_col_idx: Optional[int] = None
) -> None:
    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    on_time_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    on_time_font = Font(color="166534", bold=True)
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
            elif any(w in status_val.lower() for w in ("late", "missed", "outside")):
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
        key = clean.lower()
        if key not in seen or (clean != clean.lower() and seen[key] == seen[key].lower()):
            seen[key] = clean
    return sorted(seen.values(), key=lambda x: x.lower())


def build_excel_bytes(target_date: Optional[str] = None, days: int = 1) -> tuple[BytesIO, int]:
    today_str = target_date or datetime.now(IST).strftime("%Y-%m-%d")

    if days > 1:
        start_dt = datetime.strptime(today_str, "%Y-%m-%d") - timedelta(days=days - 1)
        period_rows = fetch_reports(start_date=start_dt.strftime("%Y-%m-%d"))
        period_title = f"Last {days} Days"
    else:
        period_rows = fetch_reports(date_filter=today_str)
        period_title = f"Reports {today_str}"

    all_rows = fetch_reports()
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
        "Caption",
    ]
    _style_sheet(ws_log, log_headers, period_rows, status_col_idx=6)

    ws_summary = wb.create_sheet(title="Member Summary")
    summary_rows: list[tuple[str, ...]] = []
    for member in members:
        member_rows = [r for r in period_rows if _match_key(r[0]) == _match_key(member)]
        on_time = sum(1 for r in member_rows if r[5] == "On Time")
        late_count = sum(1 for r in member_rows if r[5] != "On Time")
        total_joining = sum(_safe_int(r[1]) for r in member_rows)
        expected_slots = len(slot_labels) * days
        missed = max(expected_slots - on_time - late_count, 0)
        compliance = round((on_time / expected_slots) * 100) if expected_slots else 0
        summary_rows.append(
            (
                member,
                str(total_joining),
                str(on_time),
                str(late_count),
                str(missed),
                f"{compliance}%",
                today_str,
            )
        )

    _style_sheet(
        ws_summary,
        ["Username", "Total Joining", "Submitted On Time", "Late", "Missed Slots", "Compliance %", "Reference Date"],
        summary_rows,
    )

    target_day_rows = fetch_reports(date_filter=today_str)
    ws_matrix = wb.create_sheet(title="Slot Matrix")
    matrix_rows: list[tuple[str, ...]] = []
    for member in members:
        member_day_slots = {
            r[3]: (r[1], r[2], r[5]) for r in target_day_rows if _match_key(r[0]) == _match_key(member)
        }
        row_cells = [member]
        day_joining = 0
        for label in slot_labels:
            if label in member_day_slots:
                j_val, t_val, st_val = member_day_slots[label]
                day_joining += _safe_int(j_val)
                if st_val == "On Time":
                    row_cells.append(
                        f"Join: {j_val} ({t_val})" if j_val else f"On Time ({t_val})"
                    )
                else:
                    row_cells.append(
                        f"{st_val}: {j_val} ({t_val})" if j_val else f"{st_val} ({t_val})"
                    )
            else:
                row_cells.append("Missed")
        row_cells.append(str(day_joining))
        matrix_rows.append(tuple(row_cells))

    _style_sheet(
        ws_matrix,
        ["Username", "10:00 AM", "1:00 PM", "3:00 PM", "5:00 PM", "Final (<8 PM)", "Total Joining"],
        matrix_rows,
    )

    ws_all = wb.create_sheet(title="All History")
    _style_sheet(ws_all, log_headers, all_rows, status_col_idx=6)

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

    with sqlite3.connect(DB_PATH) as conn:
        tag_rows = conn.execute(
            "SELECT username, user_id, is_active FROM assistant_tags ORDER BY username ASC"
        ).fetchall()
        team_rows = conn.execute(
            "SELECT username, user_id FROM team_members ORDER BY username ASC"
        ).fetchall()
    tag_map = {u.lower(): (u, uid, act) for u, uid, act in tag_rows}
    for u, uid in team_rows:
        if u.lower() not in tag_map:
            tag_map[u.lower()] = (u, uid, 0)
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

    with sqlite3.connect(DB_PATH) as conn:
        for idx, row in enumerate(ws.iter_rows(values_only=True)):
            if idx == 0:
                headers = [str(c or "").strip().lower() for c in (row or [])]
                has_joining_col = len(headers) > 1 and headers[1] == "joining"
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
                caption = str(row[6] or "").strip() if len(row) > 6 else ""
            else:
                joining = ""
                sub_time = str(row[1] or "").strip()
                slot_label = str(row[2] or "").strip()
                rep_date = str(row[3] or "").strip()
                status = str(row[4] or "On Time").strip()
                caption = str(row[5] or "").strip() if len(row) > 5 else ""
                if not joining and caption:
                    _, joining = parse_report_fields(caption, username)

            if status == "On Time" and sub_time and slot_label:
                status = evaluate_slot_status_from_strings(sub_time, slot_label)

            exists = conn.execute(
                """
                SELECT id, COALESCE(joining, '') FROM reports
                WHERE LOWER(username) = LOWER(?) AND report_date = ? AND reporting_slot = ?
                LIMIT 1
                """,
                (username, rep_date, slot_label),
            ).fetchone()
            if not exists:
                conn.execute(
                    """
                    INSERT INTO reports (
                        user_id, username, joining, submission_time, reporting_slot, report_date, status, caption, created_at_utc
                    ) VALUES (NULL, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (username, joining, sub_time, slot_label, rep_date, status, caption, now_iso),
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
                    elif key == "reminders_enabled" and str(val).lower() in ("true", "false"):
                        dst.execute(
                            "INSERT INTO bot_settings (key, value) VALUES ('reminders_enabled', ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                            (str(val).lower(),),
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
                cap_expr = "COALESCE(caption, '')" if "caption" in rep_cols else "''"
                created_expr = "created_at_utc" if "created_at_utc" in rep_cols else f"'{now_iso}'"
                rows = src.execute(
                    f"""
                    SELECT {uid_expr}, username, {join_expr}, submission_time, reporting_slot, report_date, status, {cap_expr}, {created_expr}
                    FROM reports
                    ORDER BY id ASC
                    """
                ).fetchall()
                for uid, uname, joining, sub_time, slot_lbl, rep_date, status, caption, created_utc in rows:
                    u_str = str(uname).strip()
                    j_str = str(joining or "").strip()
                    cap_str = str(caption or "").strip()
                    if not j_str and cap_str:
                        _, j_str = parse_report_fields(cap_str, u_str)
                    exists = dst.execute(
                        """
                        SELECT id, COALESCE(joining, '') FROM reports
                        WHERE LOWER(username) = LOWER(?) AND report_date = ? AND reporting_slot = ?
                        LIMIT 1
                        """,
                        (u_str, str(rep_date), str(slot_lbl)),
                    ).fetchone()
                    if not exists:
                        dst.execute(
                            """
                            INSERT INTO reports (
                                user_id, username, joining, submission_time, reporting_slot, report_date, status, caption, created_at_utc
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (
                                uid,
                                u_str,
                                j_str,
                                str(sub_time),
                                str(slot_lbl),
                                str(rep_date),
                                str(status or "On Time"),
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
    members = sorted(set(get_team_members()) | set(get_tagged_assistants()))
    tagged = get_tagged_assistants()
    tagged_lower = {u.lower() for u in tagged}
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
        is_tagged = member.lower() in tagged_lower
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
def _build_summary_view(today_str: str) -> tuple[str, InlineKeyboardMarkup]:
    rows = fetch_reports(today_str)
    members = get_team_members()
    excluded = get_summary_excluded()
    excluded_keys = {_match_key(u) for u in excluded}
    _admin_ids, admin_keys = get_known_admin_sets()
    hidden_keys = excluded_keys | admin_keys

    visible_rows = [r for r in rows if _match_key(r[0]) not in admin_keys]

    lines = [
        f"<b>Daily Summary ({today_str} IST)</b>",
        f"Total Slot Reports Recorded: {len(visible_rows)}",
        "",
    ]

    all_names = [
        name
        for name in _dedup_case_insensitive(members + [r[0] for r in visible_rows])
        if _match_key(name) not in hidden_keys
    ]
    if not all_names:
        lines.append("No active assistants shown in summary.")
    else:
        for name in all_names:
            u_rows = [r for r in visible_rows if _match_key(r[0]) == _match_key(name)]
            u_late = sum(1 for r in u_rows if r[5] != "On Time")
            u_join = sum(_safe_int(r[1]) for r in u_rows)
            late_note = f" ({u_late} Late)" if u_late > 0 else ""
            lines.append(f"- {html.escape(name)}: {len(u_rows)}/5 slots{late_note} | Joining: {u_join}")

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
    if DEFAULT_ADMIN_CHAT_ID is None or not update.effective_user:
        return True
    return update.effective_user.id == DEFAULT_ADMIN_CHAT_ID


# ---------------------------------------------------------------------------
# Telegram Command & Callback Handlers
# ---------------------------------------------------------------------------
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.effective_chat or not update.effective_user:
        return

    rem_status = "ON" if is_reminders_enabled() else "OFF"
    work_gid = get_work_group_id() or "Not set"
    excel_gid = get_excel_group_id() or "Not set"
    tagged_count = len(get_tagged_assistants())

    await update.message.reply_text(
        f"<b>Excel Report Tracker Bot is active.</b>\n"
        f"Your User ID (<code>ADMIN_CHAT_ID</code>): <code>{update.effective_user.id}</code>\n"
        f"Current Chat ID: <code>{update.effective_chat.id}</code>\n"
        f"Work Group: <code>{work_gid}</code> | Excel Group: <code>{excel_gid}</code>\n"
        f"Slot Reminders: <b>{rem_status}</b> | Tagged Assistants: <b>{tagged_count}</b>\n\n"
        f"<b>Assistant Tag Controls:</b>\n"
        f"/tags - Open interactive button panel\n"
        f"/settags @a @b - Set exact assistants to tag\n"
        f"/addtag @a | /removetag @a | /cleartags\n"
        f"/reminder on | /reminder off - Toggle reminder warnings\n\n"
        f"<b>Backfill &amp; Excel Setup:</b>\n"
        f"/add (reply to any already-sent report) - Add with original timing\n"
        f"/addreport @user 40 [10am] - Manually add or bulk paste past reports\n"
        f"/removesummary @a | /restoresummary @a - Remove/restore name in Summary\n"
        f"/setworkgroup | /setexcelgroup | /check - Link &amp; diagnose groups\n"
        f"/excel [YYYY-MM-DD] | /weekly | /summary | /mystatus",
        parse_mode="HTML",
    )


async def check_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.effective_chat or not update.effective_user:
        return
    chat = update.effective_chat
    work_gid = get_work_group_id()
    excel_gid = get_excel_group_id()
    ist_now = datetime.now(IST)
    active_slot = detect_active_slot(ist_now)
    matches_work = work_gid is None or chat.id == work_gid
    same_as_excel = excel_gid is not None and chat.id == excel_gid and excel_gid != work_gid
    total_reports = await asyncio.to_thread(count_total_reports)
    total_members = len(await asyncio.to_thread(get_team_members))

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

    lines = [
        "<b>Bot Diagnostic Check</b>",
        f"Current Chat ID: <code>{chat.id}</code> ({html.escape(chat.type)})",
        f"Configured Work Group (<code>ALLOWED_CHAT_ID</code>): <code>{work_gid or 'Any Group'}</code>",
        f"Database/Excel Group (<code>EXCEL_GROUP_ID</code>): <code>{excel_gid or 'Not set'}</code>",
        f"Saved Data: <b>{total_reports}</b> reports | <b>{total_members}</b> assistants",
        f"Database Backup: {html.escape(pinned_backup_status)}",
        f"Current IST Time: <b>{ist_now.strftime('%Y-%m-%d %I:%M:%S %p IST')}</b>",
        f"Active Slot Now: <b>{html.escape(active_slot or 'None (Outside slot hours)')}</b>",
        f"Report Recording Here: <b>{'YES' if (matches_work and not same_as_excel) else 'NO (Group Mismatch)'}</b>",
    ]
    if not matches_work:
        lines.append(
            f"\n<b>Fix:</b> Send <code>/setworkgroup</code> here AND update Render "
            f"<code>ALLOWED_CHAT_ID</code> to <code>{chat.id}</code>."
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

    status = evaluate_slot_status(ist_dt, slot_label)

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
    )

    excel_caption = (
        f"Excel Updated ({action}) | {username} | Joining: {joining or 'N/A'} | {slot_label}\n"
        f"Time: {submission_time} | Date: {report_date} | Status: {status}"
    )
    await push_excel_to_group(context, target_date=report_date, caption=excel_caption)

    await update.message.reply_text(
        f"<b>Added ({html.escape(action)}) from message!</b>\n"
        f"Username: <code>{html.escape(username)}</code>\n"
        f"Joining: <b>{html.escape(joining or 'N/A')}</b>\n"
        f"Timing: {html.escape(submission_time)} | Date: {report_date}\n"
        f"Slot: {html.escape(slot_label)} | Status: <b>{html.escape(status)}</b>",
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
        status = evaluate_slot_status(ist_now, slot_label)
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

            status = evaluate_slot_status(target_dt, slot_label)

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
            )
            added_items.append(f"{norm_user} | Joining: {joining or 'N/A'} | {slot_label} ({sub_time} | {status})")

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

    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can turn reminders ON or OFF.")
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

    if update.effective_chat.type == "private" and not _is_admin(update):
        await update.message.reply_text("Unauthorized: Only the configured admin can fetch reports.")
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

    if update.effective_chat.type == "private" and not _is_admin(update):
        await update.message.reply_text("Unauthorized: Only the configured admin can fetch reports.")
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
    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    submitted_map = await asyncio.to_thread(fetch_user_today_slots, username, today_str)

    lines = [f"<b>Slot Status for {html.escape(username)} ({today_str})</b>"]
    for slot in REPORTING_SLOTS:
        if slot.label in submitted_map:
            sub_t, j_val = submitted_map[slot.label]
            j_note = f" | Join: {html.escape(j_val)}" if j_val else ""
            lines.append(f"- {html.escape(slot.label)}: Done ({html.escape(sub_t)}{j_note})")
        else:
            lines.append(f"- {html.escape(slot.label)}: Pending")

    if update.effective_chat and update.effective_chat.type != "private":
        work_gid = get_work_group_id()
        if work_gid is not None and update.effective_chat.id != work_gid:
            lines.append(
                f"\n<b>Warning:</b> This group (<code>{update.effective_chat.id}</code>) does not match "
                f"Work Group (<code>{work_gid}</code>). Send <code>/setworkgroup</code> here and update "
                f"<code>ALLOWED_CHAT_ID={update.effective_chat.id}</code> on Render!"
            )

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
# Scheduled Jobs (Daily Excel + Tagged Assistant Warnings)
# ---------------------------------------------------------------------------
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
                    f"Daily Automated Excel Report ({today_str} | 8:30 PM IST)\n"
                    f"Total slot submissions today: {today_count}"
                ),
            )
        except Exception as exc:
            logger.warning("Could not send daily Excel to Admin DM (%s): %s", DEFAULT_ADMIN_CHAT_ID, exc)

    await push_excel_to_group(
        context,
        target_date=today_str,
        caption=f"Daily Closing Excel Report | {today_str} | 8:30 PM IST | Total today: {today_count}",
    )


async def send_slot_open_reminder_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_reminders_enabled() or not context.job:
        return

    work_group_id = get_work_group_id()
    if work_group_id is None:
        return

    await refresh_group_admins(context.bot, work_group_id)
    slot_label = str(context.job.data)
    tagged_pairs = await asyncio.to_thread(get_tagged_assistants_with_ids)
    tags_line = " ".join(_format_mention_html(u, uid) for u, uid in tagged_pairs)
    attention_block = f"Attention: {tags_line}\n" if tags_line else ""

    if "Final" in slot_label:
        msg = (
            f"<b>WARNING &amp; REMINDER:</b> {html.escape(slot_label)} is now OPEN!\n"
            f"{attention_block}"
            f"Please submit your report (Username + Joining) before 8:00 PM IST (grace ends 8:15 PM IST). "
            f"Reports are ONLY recorded during active slot hours."
        )
    else:
        msg = (
            f"<b>WARNING &amp; REMINDER:</b> {html.escape(slot_label)} is now OPEN!\n"
            f"{attention_block}"
            f"Submit your report (Username + Joining) within 15 minutes. "
            f"Reports outside slot hours will NOT be recorded."
        )

    try:
        await context.bot.send_message(chat_id=work_group_id, text=msg, parse_mode="HTML")
    except Exception as exc:
        logger.warning("Could not send slot-open reminder to %s: %s", work_group_id, exc)


async def send_slot_final_warning_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_reminders_enabled() or not context.job:
        return

    work_group_id = get_work_group_id()
    if work_group_id is None:
        return

    await refresh_group_admins(context.bot, work_group_id)
    slot_label = str(context.job.data)
    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    pending = await asyncio.to_thread(
        get_pending_tagged_assistants_for_slot, today_str, slot_label
    )

    if not pending:
        return

    tags_line = " ".join(_format_mention_html(u, uid) for u, uid in pending)
    msg = (
        f"<b>FINAL WARNING (5 MINUTES LEFT)</b> — {html.escape(slot_label)}\n"
        f"Pending Assistants: {tags_line}\n"
        f"Submit your report immediately! Slot closes in 5 minutes and will NOT record late uploads."
    )
    try:
        await context.bot.send_message(chat_id=work_group_id, text=msg, parse_mode="HTML")
    except Exception as exc:
        logger.warning("Could not send 5-min final warning to %s: %s", work_group_id, exc)


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
    Runs in handler group=-1 on EVERY work-group message (text, photo, command, join)
    so non-admin assistants are auto-recognized and Group Admins are always excluded.
    """
    message = update.message
    user = update.effective_user
    chat = update.effective_chat
    if not message or not user or not chat or user.is_bot:
        return

    work_group_id = get_work_group_id()
    excel_group_id = get_excel_group_id()
    if excel_group_id is not None and chat.id == excel_group_id and excel_group_id != work_group_id:
        return

    if work_group_id is not None and chat.id != work_group_id:
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
    Record Username, Joining, Timing & Caption ONLY for non-admin assistants when caption matches report format.
    Random photos, admin photos, and messages without Username/Joining are silently ignored (zero reply spam).
    """
    message = update.message
    user = update.effective_user
    chat = update.effective_chat

    if not message or not user or not chat or user.is_bot:
        return

    raw_caption = (message.caption or message.text or "").strip()

    # Ignore random photos, blank screenshots, leaderboards, and casual chats that are NOT a report format
    if not is_report_caption(raw_caption):
        return

    work_group_id = get_work_group_id()
    excel_group_id = get_excel_group_id()
    if excel_group_id is not None and chat.id == excel_group_id and excel_group_id != work_group_id:
        return

    if work_group_id is not None and chat.id != work_group_id:
        logger.warning("Ignored report: chat_id=%s != work_group_id=%s", chat.id, work_group_id)
        now_ist = datetime.now(IST)
        last_warn = _MISMATCH_WARN_CACHE.get(chat.id)
        if last_warn is None or (now_ist - last_warn).total_seconds() >= 300:
            _MISMATCH_WARN_CACHE[chat.id] = now_ist
            await message.reply_text(
                f"<b>Report not counted:</b> This group ID (<code>{chat.id}</code>) does not match "
                f"configured Work Group (<code>{work_group_id}</code>).\n"
                f"Admin: send <code>/setworkgroup</code> here right now, and set "
                f"<code>ALLOWED_CHAT_ID={chat.id}</code> in Render.",
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
    effective_uid = user.id if username.lower() == fallback_user.lower() else None

    ist_dt = message.date.astimezone(IST)
    submission_time = ist_dt.strftime("%I:%M:%S %p IST")
    report_date = ist_dt.strftime("%Y-%m-%d")
    active_slot = detect_active_slot(ist_dt)

    if active_slot is None:
        if message.media_group_id:
            mg_key = f"{chat.id}:{message.media_group_id}"
            if mg_key in _SEEN_MEDIA_GROUPS:
                return
            _SEEN_MEDIA_GROUPS[mg_key] = datetime.now(IST)
        # Quiet reaction only (no noisy reply spam outside slot hours)
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

    async with _REPORT_LOCK:
        if message.media_group_id:
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
        # Allow updating if previous slot entry had empty Joining and new message provides Joining
        should_update_missing_joining = bool(existing and existing[2] == "" and joining != "")

        if PREVENT_DUPLICATE_SLOTS and existing and not should_update_missing_joining:
            await message.reply_text(
                f"Duplicate skipped for {username}: {active_slot} was already recorded today at {existing[1]}.",
                reply_to_message_id=message.message_id,
            )
            return

        status = "On Time"
        logger.info(
            "Slot report recorded | user=%s | joining=%s | chat=%s | time=%s | slot=%s",
            username,
            joining,
            chat.id,
            submission_time,
            active_slot,
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
            )
        except Exception as exc:
            logger.exception("Failed to save report to database: %s", exc)
            await message.reply_text(
                "Report received, but saving failed. Please notify an admin.",
                reply_to_message_id=message.message_id,
            )
            return

    await send_celebration_reaction(context, chat.id, message.message_id, is_late=False)

    excel_caption = (
        f"Excel Updated | {username} | Joining: {joining or 'N/A'} | {active_slot}\n"
        f"Time: {submission_time} | Date: {report_date} | Status: {status}"
    )
    if raw_caption:
        excel_caption += f"\nCaption: {raw_caption}"
    await push_excel_to_group(context, target_date=report_date, caption=excel_caption)

    join_line = f"\nJoining: {joining}" if joining else ""
    label_verb = "Updated" if action == "updated" else "Recorded"
    confirmation = (
        f"🎉🔥💯 🏆 ✅ <b>Report {label_verb}!</b>\n"
        f"Assistant: <code>{html.escape(username)}</code>{join_line}\n"
        f"Slot: {html.escape(active_slot)}\n"
        f"Time: {html.escape(submission_time)} | Status: <b>{status}</b>"
    )
    await message.reply_text(
        confirmation,
        parse_mode="HTML",
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
    work_gid = get_work_group_id()
    if work_gid is not None:
        await refresh_group_admins(application.bot, work_gid, force=True)

    if application.job_queue is None:
        return

    application.job_queue.run_daily(
        send_daily_excel_job,
        time=DAILY_SEND_TIME_IST,
        name="daily_excel_report",
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

    logger.info("Scheduled daily Excel (20:30 IST), slot-open warnings, and 5-min final warnings.")


def main() -> None:
    if not BOT_TOKEN:
        raise ValueError("TELEGRAM_BOT_TOKEN environment variable is missing.")

    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)

    init_db()
    start_health_server()

    application = ApplicationBuilder().token(BOT_TOKEN).post_init(post_init).build()

    # Run assistant auto-discovery in group=-1 on ALL group messages before command/report handlers
    application.add_handler(
        MessageHandler(filters.ChatType.GROUPS, handle_any_group_message),
        group=-1,
    )

    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(CommandHandler("check", check_command))
    application.add_handler(CommandHandler(["add", "record"], add_command))
    application.add_handler(CommandHandler("addreport", addreport_command))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler(["tags", "assistants"], tags_command))
    application.add_handler(CommandHandler("settags", settags_command))
    application.add_handler(CommandHandler(["addtag", "addassistant"], addtag_command))
    application.add_handler(CommandHandler(["removetag", "removeassistant"], removetag_command))
    application.add_handler(CommandHandler("cleartags", cleartags_command))
    application.add_handler(CallbackQueryHandler(tags_callback_handler))
    application.add_handler(CommandHandler("reminder", reminder_command))
    application.add_handler(CommandHandler("setworkgroup", setworkgroup_command))
    application.add_handler(CommandHandler(["setexcelgroup", "setdbgroup"], setexcelgroup_command))
    application.add_handler(CommandHandler(["excel", "report"], excel_command))
    application.add_handler(CommandHandler("weekly", weekly_command))
    application.add_handler(CommandHandler("summary", summary_command))
    application.add_handler(CommandHandler(["removesummary", "hidesummary"], removesummary_command))
    application.add_handler(CommandHandler(["restoresummary", "addsummary"], restoresummary_command))
    application.add_handler(CommandHandler("mystatus", mystatus_command))
    application.add_handler(CommandHandler("addmember", addmember_command))
    application.add_handler(CommandHandler("removemember", removemember_command))
    application.add_handler(CommandHandler("team", team_command))
    application.add_handler(CommandHandler(["restore", "recover"], restore_command))

    # Auto-restore when Admin forwards/sends .db or .xlsx in Database Group or Admin DM
    application.add_handler(
        MessageHandler(filters.Document.ALL & (filters.ChatType.GROUPS | filters.ChatType.PRIVATE), handle_backup_document_upload),
        group=1,
    )

    group_report_filter = filters.ChatType.GROUPS & (
        filters.PHOTO | filters.Document.IMAGE | (filters.TEXT & ~filters.COMMAND)
    )
    application.add_handler(MessageHandler(group_report_filter, handle_slot_report))

    logger.info("Starting Telegram Excel Report Bot (polling)...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
