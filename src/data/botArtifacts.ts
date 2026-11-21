import botPyRaw from '../../bot.py?raw';
import requirementsTxtRaw from '../../requirements.txt?raw';

export const BOT_PY_SOURCE = botPyRaw;
export const REQUIREMENTS_TXT_SOURCE = requirementsTxtRaw;

export interface ReportingSlot {
  id: string;
  label: string;
  shortLabel: string;
  targetDisplay: string;
  windowDisplay: string;
  startSec: number;
  endSec: number;
}

const toSeconds = (h: number, m: number, s: number) => h * 3600 + m * 60 + s;

export const REPORTING_SLOTS: ReportingSlot[] = [
  {
    id: '10am',
    label: '10:00 AM Slot',
    shortLabel: '10:00 AM',
    targetDisplay: '10:00 AM',
    windowDisplay: '10:00:00 AM - 10:15:59 AM IST',
    startSec: toSeconds(10, 0, 0),
    endSec: toSeconds(10, 15, 59),
  },
  {
    id: '1pm',
    label: '1:00 PM Slot',
    shortLabel: '1:00 PM',
    targetDisplay: '1:00 PM',
    windowDisplay: '01:00:00 PM - 01:15:59 PM IST',
    startSec: toSeconds(13, 0, 0),
    endSec: toSeconds(13, 15, 59),
  },
  {
    id: '3pm',
    label: '3:00 PM Slot',
    shortLabel: '3:00 PM',
    targetDisplay: '3:00 PM',
    windowDisplay: '03:00:00 PM - 03:15:59 PM IST',
    startSec: toSeconds(15, 0, 0),
    endSec: toSeconds(15, 15, 59),
  },
  {
    id: '5pm',
    label: '5:00 PM Slot',
    shortLabel: '5:00 PM',
    targetDisplay: '5:00 PM',
    windowDisplay: '05:00:00 PM - 05:15:59 PM IST',
    startSec: toSeconds(17, 0, 0),
    endSec: toSeconds(17, 15, 59),
  },
  {
    id: 'final',
    label: 'Final Report (Before 8:00 PM)',
    shortLabel: 'Final (7:15-8 PM)',
    targetDisplay: 'Before 8:00 PM',
    windowDisplay: '07:15:00 PM - 08:00:59 PM IST',
    startSec: toSeconds(19, 15, 0),
    endSec: toSeconds(20, 0, 59),
  },
];

export interface SlotResult {
  inActiveSlot: boolean;
  reportingSlot: string;
  status: 'On Time' | 'Outside Slot (Not Recorded)';
  formattedTimeIst: string;
}

export function formatTime12Hour(time24: string): string {
  const [hStr, mStr, sStr = '00'] = time24.split(':');
  const h24 = Number(hStr);
  const period = h24 >= 12 ? 'PM' : 'AM';
  const h12 = h24 % 12 === 0 ? 12 : h24 % 12;
  return `${String(h12).padStart(2, '0')}:${mStr.padStart(2, '0')}:${sStr.padStart(2, '0')} ${period} IST`;
}

export function evaluateIstSlot(time24: string): SlotResult {
  const [hStr, mStr, sStr = '0'] = time24.split(':');
  const currentSec = toSeconds(Number(hStr), Number(mStr), Number(sStr));
  const formattedTimeIst = formatTime12Hour(time24);

  for (const slot of REPORTING_SLOTS) {
    if (currentSec >= slot.startSec && currentSec <= slot.endSec) {
      return {
        inActiveSlot: true,
        reportingSlot: slot.label,
        status: 'On Time',
        formattedTimeIst,
      };
    }
  }

  return {
    inActiveSlot: false,
    reportingSlot: 'Outside Slot Window',
    status: 'Outside Slot (Not Recorded)',
    formattedTimeIst,
  };
}

export interface SheetRow {
  id: string;
  username: string;
  joining: string;
  submissionTime: string;
  reportingSlot: string;
  date: string;
  status: 'On Time';
  caption: string;
}

export const INITIAL_SHEET_ROWS: SheetRow[] = [
  {
    id: 'row-1',
    username: '@Govind7549',
    joining: '40',
    submissionTime: '10:02:14 AM IST',
    reportingSlot: '10:00 AM Slot',
    date: '2026-11-18',
    status: 'On Time',
    caption: 'Username- @Govind7549\nJoining- 40',
  },
  {
    id: 'row-2',
    username: '@priya_lead',
    joining: '32',
    submissionTime: '01:08:52 PM IST',
    reportingSlot: '1:00 PM Slot',
    date: '2026-11-18',
    status: 'On Time',
    caption: 'Username- @priya_lead\nJoining- 32',
  },
  {
    id: 'row-3',
    username: '@sneha_qa',
    joining: '55',
    submissionTime: '07:54:31 PM IST',
    reportingSlot: 'Final Report (Before 8:00 PM)',
    date: '2026-11-18',
    status: 'On Time',
    caption: 'Username- @sneha_qa\nJoining- 55',
  },
];

export const INITIAL_TEAM_ROSTER: string[] = [
  '@Govind7549',
  '@priya_lead',
  '@sneha_qa',
  '@vikram_dev',
];

export interface BotCommandInfo {
  command: string;
  scope: string;
  description: string;
}

export const BOT_COMMANDS: BotCommandInfo[] = [
  {
    command: '/add [10am|1pm|3pm|5pm|final]',
    scope: 'Admin (Reply to msg)',
    description: 'Reply to any already-sent report in the group to record Username, Joining & original IST timing (marks Late if past slot cutoff).',
  },
  {
    command: '/addreport @user 40 [10am] [10:05]',
    scope: 'Admin Only',
    description: 'Manually add or bulk-paste past reports into SQLite and Excel (auto-marks On Time vs Late from timestamp).',
  },
  {
    command: '/check',
    scope: 'Group / DM',
    description: 'Shows current Chat ID, configured Work/Excel Group IDs, active IST slot, and whether counting is active.',
  },
  {
    command: '/tags (or /assistants)',
    scope: 'Admin Only',
    description: 'Opens an interactive button panel in Telegram to toggle which auto-detected assistants get tagged.',
  },
  {
    command: '/settags @a @b',
    scope: 'Admin Only',
    description: 'Sets the exact list of assistants to tag in reminders (/addtag, /removetag, /cleartags).',
  },
  {
    command: '/countdown [10am|1pm|3pm|5pm|final|stop]',
    scope: 'Admin Only',
    description: 'Starts/stops the live 15-min animated countdown (7:45 PM Final Warning or manual regular slot; Regular Late = ₹5, Final Late >8 PM = ₹10, Missed = ₹15).',
  },
  {
    command: '/reminder on | off',
    scope: 'Admin Only',
    description: 'Turns scheduled slot open alerts (7:15 PM Final), 7:45 PM animated countdown, and 8:00 PM Last Warning ON or OFF live.',
  },
  {
    command: '/setexcelgroup',
    scope: 'Admin (in Excel Group)',
    description: 'Links your Excel Report Group so team_report_YYYY-MM-DD.xlsx is sent on every slot submission.',
  },
  {
    command: '/setworkgroup',
    scope: 'Admin (in Work Group)',
    description: 'Links your work Telegram group for auto assistant discovery, slot recording, and warnings.',
  },
  {
    command: '/excel [YYYY-MM-DD]',
    scope: 'Admin / Group',
    description: 'Generates and sends the 5-sheet .xlsx workbook (Regular Late ₹5, Final Late ₹10, Missed ₹15 columns) for today or a specific date.',
  },
  {
    command: '/weekly & /monthly [YYYY-MM]',
    scope: 'Admin / Group',
    description: 'Sends 7-day or monthly consolidated Excel workbooks summing Valid Daily Joining across days (no double-counting within the same day) + penalties & balance due.',
  },
  {
    command: '/payout [days] [rate] & /setrate <Rs>',
    scope: 'Admin Only',
    description: 'Generates a net payout Excel sheet: (Valid Joinings × Rate) - Unpaid/Unwaived Fines.',
  },
  {
    command: '/summary, /leaderboard, /mystatus & /myfine',
    scope: 'Group / DM',
    description: 'Shows Valid Daily Joining (latest slot value per day), Daily Target diff, Ranked Leaderboard (auto 8:45 PM), Fine Card, and interactive name manager.',
  },
  {
    command: '/leave & /unleave',
    scope: 'Admin Only',
    description: '/leave @user [today|YYYY-MM-DD] [slot|full] — Exempts assistant from reminders and missed penalties for a slot or full day.',
  },
  {
    command: '/paid @user <amt> & /waive @user <amt>',
    scope: 'Admin Only',
    description: 'Tracks fine payments and waivers in the ledger so Balance Due updates accurately across summaries and Excel.',
  },
  {
    command: '/setshift @user morning|evening|full & /settarget <N>',
    scope: 'Admin Only',
    description: 'Assigns custom reporting shifts per assistant and configures a daily joining goal with + / - variance.',
  },
  {
    command: '/setpenalty regular 5 final 10 missed 15',
    scope: 'Admin Only',
    description: 'Updates regular late, final late (>8 PM), and missed penalty rates live without editing code.',
  },
  {
    command: '/holiday on|off, /offdays sun & /cleanchat on|off',
    scope: 'Admin Only',
    description: 'Pauses alerts on holidays/Sundays and auto-deletes stale slot reminder messages when Clean Chat Mode is ON.',
  },
  {
    command: '/alias @old @new & /aliases',
    scope: 'Admin Only',
    description: 'Merges changed/alternate Telegram usernames into one canonical profile so reports and penalties never split.',
  },
  {
    command: '/find @user [days] & /day [YYYY-MM-DD]',
    scope: 'Group / DM',
    description: 'Searches any assistant history across N days or pulls a full daily snapshot for any past date.',
  },
  {
    command: '/addworkgroup, /removeworkgroup & /groups',
    scope: 'Admin Only',
    description: 'Tracks multiple work groups simultaneously alongside your central Database/Excel group.',
  },
  {
    command: '/removesummary @user | /restoresummary @user',
    scope: 'Admin Only',
    description: 'Removes or restores any assistant name from the Telegram Summary format & Excel Member Summary / Slot Matrix sheets.',
  },
  {
    command: '/dmstatus, /dmtest & /dmall [msg]',
    scope: 'Admin Only',
    description: 'Tracks which non-admin Assistants have clicked /start and lets you send instant or automatic slot DMs to all started assistants (Group Admins strictly excluded). Use /resetfine [@user|all] to zero all fines.',
  },
  {
    command: '/restore (or /recover)',
    scope: 'Admin Only',
    description: 'Recovers all old reports, old Work Group ID, members & tags from pinned or replied .db (reports.db) or .xlsx files in Database Group (auto nightly backup at 9:15 PM IST).',
  },
];

export const RENDER_ENV_BLOCK = `TELEGRAM_BOT_TOKEN=1234567890:AAHxxxxxxxxxxxxxxxxxxxxxxxxxxxx
ADMIN_CHAT_ID=123456789
ALLOWED_CHAT_ID=-1001234567890
EXCEL_GROUP_ID=-1009876543210
ENABLE_REMINDERS=true
PREVENT_DUPLICATE_SLOTS=true
DB_PATH=reports.db`;

export interface SetupStep {
  id: string;
  phase: string;
  title: string;
  summary: string;
  commandsOrNotes: string[];
}

export const SETUP_STEPS: SetupStep[] = [
  {
    id: 'step-1',
    phase: '01',
    title: 'Create Telegram Bot & Turn Off Group Privacy',
    summary:
      'Message @BotFather on Telegram, run /newbot to copy your TELEGRAM_BOT_TOKEN, and turn Group Privacy off so the bot sees all messages and photo uploads in your work group.',
    commandsOrNotes: [
      'In @BotFather: /mybots -> Select Bot -> Bot Settings -> Group Privacy -> Turn off',
      'Add the bot to your Work Group and your Excel Report Group (make it Admin in the Excel Group so it can pin latest .xlsx files).',
    ],
  },
  {
    id: 'step-2',
    phase: '02',
    title: 'Push bot.py & requirements.txt to GitHub',
    summary:
      'Upload bot.py and requirements.txt to your GitHub repository. Keep tokens inside Render Environment Variables.',
    commandsOrNotes: [
      'git init && git add bot.py requirements.txt && git commit -m "Excel-only auto-assistant Telegram bot"',
      'git push -u origin main',
    ],
  },
  {
    id: 'step-3',
    phase: '03',
    title: 'Deploy on Render (Free Web Service) & Paste .env',
    summary:
      'In Render Dashboard, click New -> Web Service, connect your GitHub repo, set Build/Start commands, and paste the Environment block.',
    commandsOrNotes: [
      'Build Command: pip install -r requirements.txt  |  Start Command: python bot.py',
      'Set TELEGRAM_BOT_TOKEN, ADMIN_CHAT_ID, ALLOWED_CHAT_ID, EXCEL_GROUP_ID.',
    ],
  },
  {
    id: 'step-4',
    phase: '04',
    title: 'Auto-Discover Assistants & Manage Tags with /tags',
    summary:
      'Whenever any assistant sends any message in the work group, the bot automatically recognizes them. Use /tags to toggle who gets tagged or /reminder on|off anytime.',
    commandsOrNotes: [
      '/setworkgroup (in Work Group)  |  /setexcelgroup (in Excel Report Group)',
      '/tags -> tap buttons to enable/disable tagging  |  /reminder on | off',
    ],
  },
  {
    id: 'step-5',
    phase: '05',
    title: 'Keep Render Free Tier Awake (10-Minute Ping)',
    summary:
      'Ping your Render URL every 10 minutes via UptimeRobot or cron-job.org. Every slot submission sends an updated .xlsx Excel file and pins it for automatic restart recovery.',
    commandsOrNotes: [
      'Health URL: https://<your-app-name>.onrender.com/ (returns 200 OK)',
    ],
  },
];
