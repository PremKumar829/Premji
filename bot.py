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
import sqlite3
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

BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()

ADMIN_CHAT_ID_RAW = os.getenv("ADMIN_CHAT_ID", "").strip()
DEFAULT_ADMIN_CHAT_ID: Optional[int] = int(ADMIN_CHAT_ID_RAW) if ADMIN_CHAT_ID_RAW else None

ALLOWED_CHAT_ID_RAW = os.getenv("ALLOWED_CHAT_ID", "").strip()
DEFAULT_WORK_CHAT_ID: Optional[int] = int(ALLOWED_CHAT_ID_RAW) if ALLOWED_CHAT_ID_RAW else None

EXCEL_GROUP_ID_RAW = (os.getenv("EXCEL_GROUP_ID") or os.getenv("DB_BACKUP_CHAT_ID") or "").strip()
DEFAULT_EXCEL_GROUP_ID: Optional[int] = int(EXCEL_GROUP_ID_RAW) if EXCEL_GROUP_ID_RAW else None

ASSISTANT_TAGS_RAW = os.getenv("ASSISTANT_TAGS", "").strip()
DEFAULT_ASSISTANT_TAGS: list[str] = [
    tag.strip() for tag in ASSISTANT_TAGS_RAW.split(",") if tag.strip()
]

DEFAULT_ENABLE_REMINDERS = os.getenv("ENABLE_REMINDERS", "true").strip().lower() == "true"
PREVENT_DUPLICATE_SLOTS = os.getenv("PREVENT_DUPLICATE_SLOTS", "true").strip().lower() == "true"

DB_PATH = os.getenv("DB_PATH", "reports.db").strip()
IST = ZoneInfo("Asia/Kolkata")
DAILY_SEND_TIME_IST = time(hour=20, minute=30, second=0, tzinfo=IST)  # 8:30 PM IST


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


def auto_discover_assistant(username: str, user_id: Optional[int] = None) -> bool:
    """
    Automatically register any non-admin group member as an assistant on their first message.
    Uses INSERT OR IGNORE so if admin later toggles them OFF (is_active=0), it is respected.
    Also updates user_id if missing so mentions work even without @username.
    """
    clean = username.strip()
    if not clean:
        return False
    now_iso = datetime.now(IST).isoformat()
    with sqlite3.connect(DB_PATH) as conn:
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
    if not clean:
        return False
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute(
            "INSERT OR IGNORE INTO team_members (username, user_id, added_at) VALUES (?, NULL, ?)",
            (clean, datetime.now(IST).isoformat()),
        )
        conn.commit()
        return cur.rowcount > 0


def remove_team_member(username: str) -> bool:
    clean = username.strip()
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute("DELETE FROM team_members WHERE LOWER(username) = LOWER(?)", (clean,))
        conn.execute("DELETE FROM assistant_tags WHERE LOWER(username) = LOWER(?)", (clean,))
        conn.commit()
        return cur.rowcount > 0


def get_team_members() -> list[str]:
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute("SELECT username FROM team_members ORDER BY username ASC").fetchall()
        return [r[0] for r in rows]


def get_tagged_assistants_with_ids() -> list[tuple[str, Optional[int]]]:
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            "SELECT username, user_id FROM assistant_tags WHERE is_active = 1 ORDER BY username ASC"
        ).fetchall()
        return [(r[0], r[1]) for r in rows]


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


def has_slot_submission(username: str, report_date: str, reporting_slot: str) -> Optional[str]:
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            """
            SELECT submission_time FROM reports
            WHERE LOWER(username) = LOWER(?) AND report_date = ? AND reporting_slot = ?
            LIMIT 1
            """,
            (username, report_date, reporting_slot),
        ).fetchone()
        return row[0] if row else None


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


def save_report_to_db(
    user_id: int,
    username: str,
    submission_time: str,
    reporting_slot: str,
    report_date: str,
    status: str,
    caption: str,
    created_at_utc: str,
) -> None:
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            INSERT INTO reports (
                user_id, username, submission_time, reporting_slot, report_date, status, caption, created_at_utc
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                user_id,
                username,
                submission_time,
                reporting_slot,
                report_date,
                status,
                caption,
                created_at_utc,
            ),
        )
        conn.commit()
    auto_discover_assistant(username, user_id=user_id)


def fetch_reports(
    date_filter: Optional[str] = None,
    start_date: Optional[str] = None,
) -> list[tuple[str, str, str, str, str, str]]:
    with sqlite3.connect(DB_PATH) as conn:
        if date_filter:
            cursor = conn.execute(
                """
                SELECT username, submission_time, reporting_slot, report_date, status, COALESCE(caption, '')
                FROM reports
                WHERE report_date = ?
                ORDER BY id ASC
                """,
                (date_filter,),
            )
        elif start_date:
            cursor = conn.execute(
                """
                SELECT username, submission_time, reporting_slot, report_date, status, COALESCE(caption, '')
                FROM reports
                WHERE report_date >= ?
                ORDER BY report_date ASC, id ASC
                """,
                (start_date,),
            )
        else:
            cursor = conn.execute(
                """
                SELECT username, submission_time, reporting_slot, report_date, status, COALESCE(caption, '')
                FROM reports
                ORDER BY id ASC
                """
            )
        return cursor.fetchall()


def fetch_user_today_slots(username: str, report_date: str) -> dict[str, str]:
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute(
            """
            SELECT reporting_slot, submission_time
            FROM reports
            WHERE LOWER(username) = LOWER(?) AND report_date = ?
            """,
            (username, report_date),
        ).fetchall()
    return {slot: sub_time for slot, sub_time in rows}


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
            elif status_val in ("Late / Penalty", "Missed"):
                status_cell.fill = late_fill
                status_cell.font = late_font

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row in rows:
            val = str(row[col_idx - 1])
            if len(val) > max_len:
                max_len = len(val)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 5, 48)


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
    members = sorted(set(get_team_members()) | set(get_tagged_assistants()) | {r[0] for r in period_rows})
    slot_labels = [s.label for s in REPORTING_SLOTS]

    wb = Workbook()

    ws_log = wb.active
    ws_log.title = period_title[:31]
    log_headers = ["Username", "Submission Time", "Reporting Slot", "Date", "Status", "Caption"]
    _style_sheet(ws_log, log_headers, period_rows, status_col_idx=5)

    ws_summary = wb.create_sheet(title="Member Summary")
    summary_rows: list[tuple[str, ...]] = []
    for member in members:
        member_rows = [r for r in period_rows if r[0] == member]
        on_time = sum(1 for r in member_rows if r[4] == "On Time")
        expected_slots = len(slot_labels) * days
        missed = max(expected_slots - on_time, 0)
        compliance = round((on_time / expected_slots) * 100) if expected_slots else 0
        summary_rows.append((member, str(on_time), str(missed), f"{compliance}%", today_str))

    _style_sheet(
        ws_summary,
        ["Username", "Submitted On Time", "Missed Slots", "Compliance %", "Reference Date"],
        summary_rows,
    )

    target_day_rows = fetch_reports(date_filter=today_str)
    ws_matrix = wb.create_sheet(title="Slot Matrix")
    matrix_rows: list[tuple[str, ...]] = []
    for member in members:
        member_day_slots = {r[2]: r[1] for r in target_day_rows if r[0] == member}
        row_cells = [member]
        for label in slot_labels:
            row_cells.append(
                f"On Time ({member_day_slots[label]})" if label in member_day_slots else "Missed"
            )
        matrix_rows.append(tuple(row_cells))

    _style_sheet(
        ws_matrix,
        ["Username", "10:00 AM", "1:00 PM", "3:00 PM", "5:00 PM", "Final (<8 PM)"],
        matrix_rows,
    )

    ws_all = wb.create_sheet(title="All History")
    _style_sheet(ws_all, log_headers, all_rows, status_col_idx=5)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output, len(period_rows)


def restore_from_excel_bytes(excel_bytes: bytes) -> int:
    wb = load_workbook(filename=BytesIO(excel_bytes), read_only=True)
    sheet_name = "All History" if "All History" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]

    restored = 0
    now_iso = datetime.now(IST).isoformat()
    with sqlite3.connect(DB_PATH) as conn:
        for idx, row in enumerate(ws.iter_rows(values_only=True)):
            if idx == 0 or not row or not row[0]:
                continue
            username = str(row[0]).strip()
            sub_time = str(row[1] or "").strip()
            slot_label = str(row[2] or "").strip()
            rep_date = str(row[3] or "").strip()
            status = str(row[4] or "On Time").strip()
            caption = str(row[5] or "").strip() if len(row) > 5 else ""

            exists = conn.execute(
                """
                SELECT 1 FROM reports
                WHERE LOWER(username) = LOWER(?) AND report_date = ? AND reporting_slot = ?
                LIMIT 1
                """,
                (username, rep_date, slot_label),
            ).fetchone()
            if not exists:
                conn.execute(
                    """
                    INSERT INTO reports (
                        user_id, username, submission_time, reporting_slot, report_date, status, caption, created_at_utc
                    ) VALUES (NULL, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (username, sub_time, slot_label, rep_date, status, caption, now_iso),
                )
                restored += 1

            conn.execute(
                "INSERT OR IGNORE INTO team_members (username, user_id, added_at) VALUES (?, NULL, ?)",
                (username, now_iso),
            )
            conn.execute(
                "INSERT OR IGNORE INTO assistant_tags (username, user_id, is_active, added_at) VALUES (?, NULL, 1, ?)",
                (username, now_iso),
            )
        conn.commit()
    return restored


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
        f"<b>Excel &amp; Group Setup:</b>\n"
        f"/setworkgroup | /setexcelgroup - Link groups directly\n"
        f"/excel [YYYY-MM-DD] | /weekly | /summary | /mystatus",
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
        await query.answer("Only the admin can modify tags or reminders.", show_alert=True)
        return

    data = query.data
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

    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    rows = await asyncio.to_thread(fetch_reports, today_str)
    members = await asyncio.to_thread(get_team_members)

    lines = [
        f"<b>Daily Summary ({today_str} IST)</b>",
        f"Total Slot Reports Recorded: {len(rows)}",
        "",
    ]

    all_names = sorted(set(members) | {r[0] for r in rows})
    if not all_names:
        lines.append("No submissions or assistants detected today.")
    else:
        for name in all_names:
            u_on = sum(1 for r in rows if r[0] == name)
            lines.append(f"- {html.escape(name)}: {u_on}/5 slots submitted")

    await update.message.reply_text("\n".join(lines), parse_mode="HTML")


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
            lines.append(f"- {html.escape(slot.label)}: Done ({html.escape(submitted_map[slot.label])})")
        else:
            lines.append(f"- {html.escape(slot.label)}: Pending")

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
    members = await asyncio.to_thread(get_team_members)
    tagged = {m.lower() for m in await asyncio.to_thread(get_tagged_assistants)}
    if not members:
        await update.message.reply_text(
            "No assistants detected yet. Make the bot Admin in the work group, send any message, or use <code>/tags</code>.",
            parse_mode="HTML",
        )
        return
    lines = ["Recognized Assistants / Team Roster:"]
    for m in members:
        badge = " (Tag: ON)" if m.lower() in tagged else " (Tag: OFF)"
        lines.append(f"- {m}{badge}")
    await update.message.reply_text("\n".join(lines))


async def restore_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Unauthorized.")
        return

    replied = update.message.reply_to_message
    if not replied or not replied.document:
        await update.message.reply_text(
            "Reply <code>/restore</code> directly to a <code>team_report_*.xlsx</code> file.",
            parse_mode="HTML",
        )
        return

    tg_file = await replied.document.get_file()
    buf = BytesIO()
    await tg_file.download_to_memory(out=buf)
    restored = await asyncio.to_thread(restore_from_excel_bytes, buf.getvalue())
    await update.message.reply_text(f"Restored {restored} report rows from Excel file.")


# ---------------------------------------------------------------------------
# Scheduled Jobs (Daily Excel + Tagged Assistant Warnings)
# ---------------------------------------------------------------------------
async def send_daily_excel_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    excel_stream, today_count = await asyncio.to_thread(build_excel_bytes, today_str, 1)
    filename = f"team_report_{today_str}.xlsx"

    if DEFAULT_ADMIN_CHAT_ID is not None:
        await context.bot.send_document(
            chat_id=DEFAULT_ADMIN_CHAT_ID,
            document=excel_stream,
            filename=filename,
            caption=(
                f"Daily Automated Excel Report ({today_str} | 8:30 PM IST)\n"
                f"Total slot submissions today: {today_count}"
            ),
        )

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

    slot_label = str(context.job.data)
    tagged_pairs = await asyncio.to_thread(get_tagged_assistants_with_ids)
    tags_line = " ".join(_format_mention_html(u, uid) for u, uid in tagged_pairs)
    attention_block = f"Attention: {tags_line}\n" if tags_line else ""

    if "Final" in slot_label:
        msg = (
            f"<b>WARNING &amp; REMINDER:</b> {html.escape(slot_label)} is now OPEN!\n"
            f"{attention_block}"
            f"Please upload your photo report before 8:00 PM IST (grace period ends 8:15 PM IST). "
            f"Reports are ONLY recorded during active slot hours."
        )
    else:
        msg = (
            f"<b>WARNING &amp; REMINDER:</b> {html.escape(slot_label)} is now OPEN!\n"
            f"{attention_block}"
            f"Upload your photo report within 15 minutes. "
            f"Reports outside slot hours will NOT be recorded."
        )

    await context.bot.send_message(chat_id=work_group_id, text=msg, parse_mode="HTML")


async def send_slot_final_warning_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_reminders_enabled() or not context.job:
        return

    work_group_id = get_work_group_id()
    if work_group_id is None:
        return

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
        f"Upload your photo report immediately! Slot closes in 5 minutes and will NOT record late uploads."
    )
    await context.bot.send_message(chat_id=work_group_id, text=msg, parse_mode="HTML")


# ---------------------------------------------------------------------------
# Group Activity: Auto-Discover Assistants (Runs in group=-1 on ALL updates)
# ---------------------------------------------------------------------------
def _extract_username(update: Update) -> Optional[str]:
    user = update.effective_user
    if not user or user.is_bot:
        return None
    return f"@{user.username}" if user.username else (user.first_name or "Unknown User")


async def handle_any_group_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Runs in handler group=-1 on EVERY work-group message (text, photo, command, join)
    so the bot automatically recognizes any assistant immediately.
    """
    message = update.message
    user = update.effective_user
    chat = update.effective_chat
    if not message or not user or not chat or user.is_bot:
        return

    excel_group_id = get_excel_group_id()
    if excel_group_id is not None and chat.id == excel_group_id:
        return

    work_group_id = get_work_group_id()
    if work_group_id is not None and chat.id != work_group_id:
        logger.info("Auto-discover skipped: chat_id=%s != work_group_id=%s", chat.id, work_group_id)
        return

    if message.new_chat_members:
        for member in message.new_chat_members:
            if member.is_bot or (DEFAULT_ADMIN_CHAT_ID is not None and member.id == DEFAULT_ADMIN_CHAT_ID):
                continue
            m_name = f"@{member.username}" if member.username else (member.first_name or "Unknown")
            newly = await asyncio.to_thread(auto_discover_assistant, m_name, member.id)
            if newly:
                logger.info("Auto-recognized new group member as assistant: %s", m_name)

    if DEFAULT_ADMIN_CHAT_ID is not None and user.id == DEFAULT_ADMIN_CHAT_ID:
        return

    username = _extract_username(update)
    if username:
        newly = await asyncio.to_thread(auto_discover_assistant, username, user.id)
        if newly:
            logger.info("Auto-recognized assistant from group message: %s (id=%s)", username, user.id)


async def handle_photo_report(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Record photo reports ONLY during active slot hours and push .xlsx to Excel Group."""
    message = update.message
    user = update.effective_user
    chat = update.effective_chat

    if not message or not user or not chat or user.is_bot:
        return

    excel_group_id = get_excel_group_id()
    if excel_group_id is not None and chat.id == excel_group_id:
        return

    work_group_id = get_work_group_id()
    if work_group_id is not None and chat.id != work_group_id:
        logger.debug("Ignored photo from non-allowed chat_id=%s", chat.id)
        return

    username = _extract_username(update) or "Unknown User"
    ist_dt = message.date.astimezone(IST)
    submission_time = ist_dt.strftime("%I:%M:%S %p IST")
    report_date = ist_dt.strftime("%Y-%m-%d")
    active_slot = detect_active_slot(ist_dt)

    if active_slot is None:
        await message.reply_text(
            f"Not recorded ({submission_time}): Reports are ONLY recorded during slot windows.\n"
            f"Active IST Slots: 10:00-10:15 AM, 1:00-1:15 PM, 3:00-3:15 PM, 5:00-5:15 PM, and 7:30-8:15 PM.",
            reply_to_message_id=message.message_id,
        )
        return

    if PREVENT_DUPLICATE_SLOTS:
        existing_time = await asyncio.to_thread(
            has_slot_submission, username, report_date, active_slot
        )
        if existing_time:
            await message.reply_text(
                f"Duplicate skipped for {username}: {active_slot} was already recorded today at {existing_time}.",
                reply_to_message_id=message.message_id,
            )
            return

    caption = (message.caption or "").strip()
    status = "On Time"

    logger.info(
        "Slot photo report recorded | user=%s | chat=%s | time=%s | slot=%s",
        username,
        chat.id,
        submission_time,
        active_slot,
    )

    try:
        await asyncio.to_thread(
            save_report_to_db,
            user.id,
            username,
            submission_time,
            active_slot,
            report_date,
            status,
            caption,
            message.date.isoformat(),
        )
    except Exception as exc:
        logger.exception("Failed to save report to database: %s", exc)
        await message.reply_text(
            "Report received, but saving failed. Please notify an admin.",
            reply_to_message_id=message.message_id,
        )
        return

    excel_caption = (
        f"Excel Updated | {username} | {active_slot}\n"
        f"Time: {submission_time} | Date: {report_date} | Status: {status}"
    )
    if caption:
        excel_caption += f"\nNote: {caption}"
    await push_excel_to_group(context, target_date=report_date, caption=excel_caption)

    display_name = user.first_name or username
    note_line = f"\nNote: {caption}" if caption else ""
    confirmation = (
        f"Report recorded for {display_name} ({username}).\n"
        f"Slot: {active_slot}\n"
        f"Time: {submission_time} | Date: {report_date}\n"
        f"Status: {status}{note_line}"
    )
    await message.reply_text(confirmation, reply_to_message_id=message.message_id)


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
# Startup Auto-Restore From Pinned Excel (.xlsx) & Job Scheduling
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
        if pinned and pinned.document and pinned.document.file_name.endswith(".xlsx"):
            tg_file = await pinned.document.get_file()
            buf = BytesIO()
            await tg_file.download_to_memory(out=buf)
            restored = await asyncio.to_thread(restore_from_excel_bytes, buf.getvalue())
            logger.info("Auto-restored %s rows from pinned Excel %s", restored, pinned.document.file_name)
    except Exception as exc:
        logger.warning("Excel auto-restore skipped: %s", exc)


async def post_init(application: Application) -> None:
    await auto_restore_from_excel_group(application)

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

    # Run assistant auto-discovery in group=-1 on ALL group messages before command/photo handlers
    application.add_handler(
        MessageHandler(filters.ChatType.GROUPS, handle_any_group_message),
        group=-1,
    )

    application.add_handler(CommandHandler("start", start_command))
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
    application.add_handler(CommandHandler("mystatus", mystatus_command))
    application.add_handler(CommandHandler("addmember", addmember_command))
    application.add_handler(CommandHandler("removemember", removemember_command))
    application.add_handler(CommandHandler("team", team_command))
    application.add_handler(CommandHandler("restore", restore_command))

    group_photo_filter = filters.ChatType.GROUPS & (filters.PHOTO | filters.Document.IMAGE)
    application.add_handler(MessageHandler(group_photo_filter, handle_photo_report))

    logger.info("Starting Telegram Excel Report Bot (polling)...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
