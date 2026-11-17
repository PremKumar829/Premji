"""
Telegram Group Photo Report Tracker -> Automatic Excel (.xlsx) Reporter (Pro Edition)
Deployable directly via GitHub + Render Free Tier (no Google Sheets / external DB needed).

Features:
  - Listens to group photo uploads & image documents in IST (Asia/Kolkata)
  - Validates 5 daily slots (10 AM, 1 PM, 3 PM, 5 PM, Final before 8 PM) + 15-min grace
  - Duplicate slot protection (prevents accidental double-logging per slot per day)
  - Captures optional photo captions in the Excel report
  - Team roster management (/addmember, /removemember, /team) + auto-registration
  - Automatic slot reminder alerts in the work group via JobQueue
  - Interactive commands: /start, /help, /summary, /mystatus, /excel [YYYY-MM-DD], /weekly
  - Automatic daily Excel (.xlsx) DM to ADMIN_CHAT_ID at 8:30 PM IST (4 styled sheets)
  - Telegram-based SQLite backup & restore (/backup, /restore) for Render Free Tier

Requires: python-telegram-bot[job-queue] >= 21.0, openpyxl >= 3.1.5, python-dotenv, tzdata
"""

import asyncio
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
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from telegram import Update
from telegram.ext import (
    Application,
    ApplicationBuilder,
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

# Your personal Telegram numeric User ID so the bot DMs you the daily Excel & backups.
ADMIN_CHAT_ID_RAW = os.getenv("ADMIN_CHAT_ID", "").strip()
ADMIN_CHAT_ID: Optional[int] = int(ADMIN_CHAT_ID_RAW) if ADMIN_CHAT_ID_RAW else None

# Optional: Specific Telegram group chat ID for photo tracking & automated reminders.
ALLOWED_CHAT_ID_RAW = os.getenv("ALLOWED_CHAT_ID", "").strip()
ALLOWED_CHAT_ID: Optional[int] = int(ALLOWED_CHAT_ID_RAW) if ALLOWED_CHAT_ID_RAW else None

ENABLE_REMINDERS = os.getenv("ENABLE_REMINDERS", "true").strip().lower() == "true"
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
    reminder_time: time  # When the bot nudges the group


REPORTING_SLOTS: tuple[ReportingSlot, ...] = (
    ReportingSlot("10:00 AM Slot", time(10, 0, 0), time(10, 15, 59), time(10, 0, 0)),
    ReportingSlot("1:00 PM Slot", time(13, 0, 0), time(13, 15, 59), time(13, 0, 0)),
    ReportingSlot("3:00 PM Slot", time(15, 0, 0), time(15, 15, 59), time(15, 0, 0)),
    ReportingSlot("5:00 PM Slot", time(17, 0, 0), time(17, 15, 59), time(17, 0, 0)),
    ReportingSlot(
        "Final Report (Before 8:00 PM)",
        time(19, 30, 0),
        time(20, 15, 59),
        time(19, 45, 0),
    ),
)


def detect_slot(ist_dt: datetime) -> tuple[str, str]:
    """
    Compare an IST-aware datetime against designated reporting slots.
    Returns (reporting_slot_label, status_label).
    """
    current_t = ist_dt.time().replace(tzinfo=None)

    for slot in REPORTING_SLOTS:
        if slot.start <= current_t <= slot.end:
            return slot.label, "On Time"

    return "Late / Outside Slot", "Late / Penalty"


# ---------------------------------------------------------------------------
# SQLite Persistence (Reports + Team Roster)
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
                added_at TEXT NOT NULL
            )
            """
        )
        # Lightweight migration if upgrading from older schema
        existing_cols = {
            row[1] for row in conn.execute("PRAGMA table_info(reports)").fetchall()
        }
        if "user_id" not in existing_cols:
            conn.execute("ALTER TABLE reports ADD COLUMN user_id INTEGER")
        if "caption" not in existing_cols:
            conn.execute("ALTER TABLE reports ADD COLUMN caption TEXT DEFAULT ''")
        conn.commit()


def ensure_member_registered(username: str) -> None:
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            "INSERT OR IGNORE INTO team_members (username, added_at) VALUES (?, ?)",
            (username, datetime.now(IST).isoformat()),
        )
        conn.commit()


def add_team_member(username: str) -> bool:
    clean = username.strip()
    if not clean:
        return False
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute(
            "INSERT OR IGNORE INTO team_members (username, added_at) VALUES (?, ?)",
            (clean, datetime.now(IST).isoformat()),
        )
        conn.commit()
        return cur.rowcount > 0


def remove_team_member(username: str) -> bool:
    clean = username.strip()
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute("DELETE FROM team_members WHERE LOWER(username) = LOWER(?)", (clean,))
        conn.commit()
        return cur.rowcount > 0


def get_team_members() -> list[str]:
    with sqlite3.connect(DB_PATH) as conn:
        rows = conn.execute("SELECT username FROM team_members ORDER BY username ASC").fetchall()
        return [r[0] for r in rows]


def has_slot_submission(username: str, report_date: str, reporting_slot: str) -> Optional[str]:
    """Return existing submission time if user already submitted this slot today."""
    if reporting_slot == "Late / Outside Slot":
        return None
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
    ensure_member_registered(username)


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


# ---------------------------------------------------------------------------
# Styled Excel (.xlsx) Workbook Builder (4 Sheets)
# ---------------------------------------------------------------------------
def _style_sheet(ws, headers: list[str], rows: list[tuple[str, ...]], status_col_idx: Optional[int] = None) -> None:
    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    on_time_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    on_time_font = Font(color="166534", bold=True)
    late_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    late_font = Font(color="991B1B", bold=True)
    warn_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    warn_font = Font(color="92400E", bold=True)

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
            elif status_val.startswith("0%") or status_val.startswith("20%") or status_val.startswith("40%"):
                status_cell.fill = warn_fill
                status_cell.font = warn_font

    # Auto-fit column widths
    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row in rows:
            val = str(row[col_idx - 1])
            if len(val) > max_len:
                max_len = len(val)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 5, 48)


def build_excel_bytes(target_date: Optional[str] = None, days: int = 1) -> tuple[BytesIO, int]:
    """
    Generate a 4-sheet .xlsx workbook:
      1) Target Date (or Last N Days) Submissions
      2) Member Summary (On Time, Late, Missed Slots, Compliance %)
      3) Slot Matrix (Each member vs 5 slots: On Time / Missed)
      4) Full History
    """
    today_str = target_date or datetime.now(IST).strftime("%Y-%m-%d")

    if days > 1:
        start_dt = datetime.strptime(today_str, "%Y-%m-%d") - timedelta(days=days - 1)
        period_rows = fetch_reports(start_date=start_dt.strftime("%Y-%m-%d"))
        period_title = f"Last {days} Days"
    else:
        period_rows = fetch_reports(date_filter=today_str)
        period_title = f"Reports {today_str}"

    all_rows = fetch_reports()
    members = sorted(set(get_team_members()) | {r[0] for r in period_rows})
    slot_labels = [s.label for s in REPORTING_SLOTS]

    wb = Workbook()

    # Sheet 1: Submissions Log
    ws_log = wb.active
    ws_log.title = period_title[:31]
    log_headers = ["Username", "Submission Time", "Reporting Slot", "Date", "Status", "Caption"]
    _style_sheet(ws_log, log_headers, period_rows, status_col_idx=5)

    # Sheet 2: Member Summary & Compliance %
    ws_summary = wb.create_sheet(title="Member Summary")
    summary_rows: list[tuple[str, ...]] = []
    for member in members:
        member_rows = [r for r in period_rows if r[0] == member]
        on_time = sum(1 for r in member_rows if r[4] == "On Time")
        late = sum(1 for r in member_rows if r[4] == "Late / Penalty")
        expected_slots = len(slot_labels) * days
        missed = max(expected_slots - on_time, 0)
        compliance = round((on_time / expected_slots) * 100) if expected_slots else 0
        summary_rows.append(
            (
                member,
                str(on_time),
                str(late),
                str(missed),
                f"{compliance}%",
                today_str,
            )
        )

    _style_sheet(
        ws_summary,
        ["Username", "On Time", "Late / Penalty", "Missed Slots", "Compliance %", "Reference Date"],
        summary_rows,
    )

    # Sheet 3: Daily Slot Matrix (Target Date)
    target_day_rows = fetch_reports(date_filter=today_str)
    ws_matrix = wb.create_sheet(title="Slot Matrix")
    matrix_rows: list[tuple[str, ...]] = []
    for member in members:
        member_day_slots = {r[2]: r[1] for r in target_day_rows if r[0] == member}
        late_count = sum(1 for r in target_day_rows if r[0] == member and r[4] == "Late / Penalty")
        row_cells = [member]
        for label in slot_labels:
            row_cells.append(f"On Time ({member_day_slots[label]})" if label in member_day_slots else "Missed")
        row_cells.append(str(late_count))
        matrix_rows.append(tuple(row_cells))

    _style_sheet(
        ws_matrix,
        ["Username", "10:00 AM", "1:00 PM", "3:00 PM", "5:00 PM", "Final (<8 PM)", "Late Uploads"],
        matrix_rows,
    )

    # Sheet 4: Full History
    ws_all = wb.create_sheet(title="All History")
    _style_sheet(ws_all, log_headers, all_rows, status_col_idx=5)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output, len(period_rows)


# ---------------------------------------------------------------------------
# Authorization Helper
# ---------------------------------------------------------------------------
def _is_admin(update: Update) -> bool:
    if ADMIN_CHAT_ID is None or not update.effective_user:
        return True
    return update.effective_user.id == ADMIN_CHAT_ID


# ---------------------------------------------------------------------------
# Telegram Command Handlers
# ---------------------------------------------------------------------------
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.effective_chat or not update.effective_user:
        return

    await update.message.reply_text(
        f"Excel Report Tracker Bot is active.\n"
        f"Your User ID (`ADMIN_CHAT_ID`): `{update.effective_user.id}`\n"
        f"Current Chat ID (`ALLOWED_CHAT_ID`): `{update.effective_chat.id}`\n\n"
        f"Commands:\n"
        f"/excel [YYYY-MM-DD] - Download Excel (.xlsx) report\n"
        f"/weekly - Download 7-day Excel workbook\n"
        f"/summary - Live today summary in chat\n"
        f"/mystatus - Check your own slots for today\n"
        f"/addmember @user - Register team member for missed-slot tracking\n"
        f"/removemember @user - Remove member from roster\n"
        f"/team - View registered roster\n"
        f"/backup - Download SQLite DB backup (reply /restore to a backup file to restore)\n"
        f"Auto-report runs daily at 8:30 PM IST.",
        parse_mode="Markdown",
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await start_command(update, context)


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
            await update.message.reply_text("Invalid date format. Use `/excel YYYY-MM-DD`.", parse_mode="Markdown")
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

    on_time_total = sum(1 for r in rows if r[4] == "On Time")
    late_total = sum(1 for r in rows if r[4] == "Late / Penalty")

    lines = [f"*Daily Summary ({today_str} IST)*", f"On Time: {on_time_total} | Late: {late_total}", ""]

    all_names = sorted(set(members) | {r[0] for r in rows})
    if not all_names:
        lines.append("No submissions or team members recorded today.")
    else:
        for name in all_names:
            u_rows = [r for r in rows if r[0] == name]
            u_on = sum(1 for r in u_rows if r[4] == "On Time")
            u_late = sum(1 for r in u_rows if r[4] == "Late / Penalty")
            lines.append(f"- {name}: {u_on}/5 slots on time, {u_late} late")

    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")


async def mystatus_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.effective_user:
        return

    user = update.effective_user
    username = f"@{user.username}" if user.username else (user.first_name or "Unknown User")
    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    submitted_map = await asyncio.to_thread(fetch_user_today_slots, username, today_str)

    lines = [f"*Slot Status for {username} ({today_str})*"]
    for slot in REPORTING_SLOTS:
        if slot.label in submitted_map:
            lines.append(f"- {slot.label}: Done ({submitted_map[slot.label]})")
        else:
            lines.append(f"- {slot.label}: Pending")

    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")


async def addmember_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can modify the team roster.")
        return
    if not context.args:
        await update.message.reply_text("Usage: `/addmember @username`", parse_mode="Markdown")
        return

    added = []
    for raw in context.args:
        if await asyncio.to_thread(add_team_member, raw):
            added.append(raw)

    if added:
        await update.message.reply_text(f"Added to roster: {', '.join(added)}")
    else:
        await update.message.reply_text("Member(s) already in roster.")


async def removemember_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Only the configured admin can modify the team roster.")
        return
    if not context.args:
        await update.message.reply_text("Usage: `/removemember @username`", parse_mode="Markdown")
        return

    removed = []
    for raw in context.args:
        if await asyncio.to_thread(remove_team_member, raw):
            removed.append(raw)

    if removed:
        await update.message.reply_text(f"Removed from roster: {', '.join(removed)}")
    else:
        await update.message.reply_text("No matching member found in roster.")


async def team_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    members = await asyncio.to_thread(get_team_members)
    if not members:
        await update.message.reply_text("Team roster is empty. Use `/addmember @username` or upload a photo.", parse_mode="Markdown")
        return
    await update.message.reply_text("Registered Team Members:\n" + "\n".join(f"- {m}" for m in members))


async def backup_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Unauthorized.")
        return
    if not os.path.exists(DB_PATH):
        await update.message.reply_text("No database file found yet.")
        return

    with open(DB_PATH, "rb") as db_file:
        await update.message.reply_document(
            document=db_file,
            filename="reports.db",
            caption="SQLite database backup. Reply `/restore` to this file anytime to restore after a Render restart.",
            parse_mode="Markdown",
        )


async def restore_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return
    if not _is_admin(update):
        await update.message.reply_text("Unauthorized.")
        return

    replied = update.message.reply_to_message
    if not replied or not replied.document:
        await update.message.reply_text("Reply `/restore` directly to a `reports.db` backup file.", parse_mode="Markdown")
        return

    tg_file = await replied.document.get_file()
    await tg_file.download_to_drive(custom_path=DB_PATH)
    await asyncio.to_thread(init_db)
    await update.message.reply_text("Database restored successfully from Telegram backup.")


# ---------------------------------------------------------------------------
# Scheduled Jobs (Daily Excel + Slot Reminders)
# ---------------------------------------------------------------------------
async def send_daily_excel_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    if ADMIN_CHAT_ID is None:
        logger.warning("ADMIN_CHAT_ID not set; skipping scheduled 8:30 PM IST Excel report.")
        return

    today_str = datetime.now(IST).strftime("%Y-%m-%d")
    excel_stream, today_count = await asyncio.to_thread(build_excel_bytes, today_str, 1)
    filename = f"team_report_{today_str}.xlsx"

    await context.bot.send_document(
        chat_id=ADMIN_CHAT_ID,
        document=excel_stream,
        filename=filename,
        caption=(
            f"Daily Automated Excel Report ({today_str} | 8:30 PM IST)\n"
            f"Total submissions today: {today_count}"
        ),
    )

    if os.path.exists(DB_PATH):
        with open(DB_PATH, "rb") as db_file:
            await context.bot.send_document(
                chat_id=ADMIN_CHAT_ID,
                document=db_file,
                filename=f"reports_backup_{today_str}.db",
                caption="Automatic daily database backup.",
            )


async def send_slot_reminder_job(context: ContextTypes.DEFAULT_TYPE) -> None:
    if not ENABLE_REMINDERS or ALLOWED_CHAT_ID is None or not context.job:
        return

    slot_label = str(context.job.data)
    if "Final" in slot_label:
        msg = "Reminder: Final Report window is open. Upload your photo report before 8:00 PM IST (grace up to 8:15 PM)."
    else:
        msg = f"Reminder: {slot_label} is now open. Upload your photo report within 15 minutes."

    await context.bot.send_message(chat_id=ALLOWED_CHAT_ID, text=msg)


# ---------------------------------------------------------------------------
# Photo Report Handler
# ---------------------------------------------------------------------------
async def handle_photo_report(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    message = update.message
    user = update.effective_user
    chat = update.effective_chat

    if not message or not user or not chat:
        return

    if ALLOWED_CHAT_ID is not None and chat.id != ALLOWED_CHAT_ID:
        logger.debug("Ignored message from non-allowed chat_id=%s", chat.id)
        return

    username = f"@{user.username}" if user.username else (user.first_name or "Unknown User")
    ist_dt = message.date.astimezone(IST)
    submission_time = ist_dt.strftime("%I:%M:%S %p IST")
    report_date = ist_dt.strftime("%Y-%m-%d")
    reporting_slot, status = detect_slot(ist_dt)
    caption = (message.caption or "").strip()

    # Prevent duplicate logging in the same valid slot on the same day
    if PREVENT_DUPLICATE_SLOTS:
        existing_time = await asyncio.to_thread(
            has_slot_submission, username, report_date, reporting_slot
        )
        if existing_time:
            await message.reply_text(
                f"Duplicate skipped for {username}: {reporting_slot} was already recorded today at {existing_time}.",
                reply_to_message_id=message.message_id,
            )
            return

    logger.info(
        "Photo report received | user=%s | chat=%s | time=%s | slot=%s | status=%s",
        username,
        chat.id,
        submission_time,
        reporting_slot,
        status,
    )

    try:
        await asyncio.to_thread(
            save_report_to_db,
            user.id,
            username,
            submission_time,
            reporting_slot,
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

    display_name = user.first_name or username
    note_line = f"\nNote: {caption}" if caption else ""
    if status == "On Time":
        confirmation = (
            f"Report recorded for {display_name} ({username}).\n"
            f"Slot: {reporting_slot}\n"
            f"Time: {submission_time} | Date: {report_date}\n"
            f"Status: {status}{note_line}"
        )
    else:
        confirmation = (
            f"Report recorded for {display_name} ({username}), outside scheduled windows.\n"
            f"Slot: {reporting_slot}\n"
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
# Application Entrypoint
# ---------------------------------------------------------------------------
async def post_init(application: Application) -> None:
    if application.job_queue is None:
        return

    application.job_queue.run_daily(
        send_daily_excel_job,
        time=DAILY_SEND_TIME_IST,
        name="daily_excel_report",
    )

    if ENABLE_REMINDERS and ALLOWED_CHAT_ID is not None:
        for slot in REPORTING_SLOTS:
            reminder_dt = datetime.combine(datetime.now(IST).date(), slot.reminder_time, tzinfo=IST)
            application.job_queue.run_daily(
                send_slot_reminder_job,
                time=reminder_dt.timetz(),
                data=slot.label,
                name=f"reminder_{slot.label}",
            )
        logger.info("Scheduled daily Excel report (20:30 IST) and 5 group slot reminders.")
    else:
        logger.info("Scheduled daily Excel report at 20:30 IST (set ALLOWED_CHAT_ID for reminders).")


def main() -> None:
    if not BOT_TOKEN:
        raise ValueError("TELEGRAM_BOT_TOKEN environment variable is missing.")

    init_db()
    start_health_server()

    application = ApplicationBuilder().token(BOT_TOKEN).post_init(post_init).build()

    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler(["excel", "report"], excel_command))
    application.add_handler(CommandHandler("weekly", weekly_command))
    application.add_handler(CommandHandler("summary", summary_command))
    application.add_handler(CommandHandler("mystatus", mystatus_command))
    application.add_handler(CommandHandler("addmember", addmember_command))
    application.add_handler(CommandHandler("removemember", removemember_command))
    application.add_handler(CommandHandler("team", team_command))
    application.add_handler(CommandHandler("backup", backup_command))
    application.add_handler(CommandHandler("restore", restore_command))

    group_photo_filter = filters.ChatType.GROUPS & (filters.PHOTO | filters.Document.IMAGE)
    application.add_handler(MessageHandler(group_photo_filter, handle_photo_report))

    logger.info("Starting Telegram Excel Report Bot (polling)...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
