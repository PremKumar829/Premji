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
    shortLabel: 'Final (<8 PM)',
    targetDisplay: 'Before 8:00 PM',
    windowDisplay: '07:30:00 PM - 08:15:59 PM IST',
    startSec: toSeconds(19, 30, 0),
    endSec: toSeconds(20, 15, 59),
  },
];

export interface SlotResult {
  reportingSlot: string;
  status: 'On Time' | 'Late / Penalty';
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
        reportingSlot: slot.label,
        status: 'On Time',
        formattedTimeIst,
      };
    }
  }

  return {
    reportingSlot: 'Late / Outside Slot',
    status: 'Late / Penalty',
    formattedTimeIst,
  };
}

export interface SheetRow {
  id: string;
  username: string;
  submissionTime: string;
  reportingSlot: string;
  date: string;
  status: 'On Time' | 'Late / Penalty';
  caption: string;
}

export const INITIAL_SHEET_ROWS: SheetRow[] = [
  {
    id: 'row-1',
    username: '@arjun_ops',
    submissionTime: '10:04:18 AM IST',
    reportingSlot: '10:00 AM Slot',
    date: '2026-11-17',
    status: 'On Time',
    caption: 'Morning floor dispatch check',
  },
  {
    id: 'row-2',
    username: '@priya_lead',
    submissionTime: '01:14:52 PM IST',
    reportingSlot: '1:00 PM Slot',
    date: '2026-11-17',
    status: 'On Time',
    caption: 'Midday inventory photo',
  },
  {
    id: 'row-3',
    username: 'Rohan Mehta',
    submissionTime: '03:22:09 PM IST',
    reportingSlot: 'Late / Outside Slot',
    date: '2026-11-17',
    status: 'Late / Penalty',
    caption: 'Network issue delay',
  },
  {
    id: 'row-4',
    username: '@sneha_qa',
    submissionTime: '07:54:31 PM IST',
    reportingSlot: 'Final Report (Before 8:00 PM)',
    date: '2026-11-17',
    status: 'On Time',
    caption: 'End-of-day closing log',
  },
];

export const INITIAL_TEAM_ROSTER: string[] = [
  '@arjun_ops',
  '@priya_lead',
  'Rohan Mehta',
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
    command: '/tags (or /assistants)',
    scope: 'Admin Only',
    description: 'Opens an interactive button panel in Telegram to toggle which assistants get tagged.',
  },
  {
    command: '/settags @a @b',
    scope: 'Admin Only',
    description: 'Sets the exact list of assistants to tag in reminders (/addtag, /removetag, /cleartags).',
  },
  {
    command: '/reminder on | off',
    scope: 'Admin Only',
    description: 'Turns scheduled assistant-tagging reminders and 5-min warnings ON or OFF live.',
  },
  {
    command: '/setdbgroup',
    scope: 'Admin (in DB Group)',
    description: 'Links your private Database Group so reports.db syncs automatically on every submission.',
  },
  {
    command: '/setworkgroup',
    scope: 'Admin (in Work Group)',
    description: 'Links your work Telegram group for photo tracking and automated assistant warnings.',
  },
  {
    command: '/excel [YYYY-MM-DD]',
    scope: 'Admin / Group',
    description: 'Generates and sends the 4-sheet .xlsx workbook for today or a specific date.',
  },
  {
    command: '/weekly',
    scope: 'Admin / Group',
    description: 'Sends a 7-day consolidated Excel workbook with Member Compliance %.',
  },
  {
    command: '/summary & /mystatus',
    scope: 'Group / DM',
    description: 'Shows live On Time vs Late breakdown or an individual member’s 5-slot progress.',
  },
  {
    command: '/backup & /restore',
    scope: 'Admin Only',
    description: 'Manual DB export/restore (plus automatic auto-restore from pinned DB group message).',
  },
];

export const RENDER_ENV_BLOCK = `TELEGRAM_BOT_TOKEN=1234567890:AAHxxxxxxxxxxxxxxxxxxxxxxxxxxxx
ADMIN_CHAT_ID=123456789
ALLOWED_CHAT_ID=-1001234567890
DB_BACKUP_CHAT_ID=-1009876543210
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
      'Message @BotFather on Telegram, run /newbot to copy your TELEGRAM_BOT_TOKEN, and turn Group Privacy off so the bot sees photo reports in your work group.',
    commandsOrNotes: [
      'In @BotFather: /mybots -> Select Bot -> Bot Settings -> Group Privacy -> Turn off',
      'Create 2 Telegram groups: (1) Work Report Group, (2) Private Database Backup Group. Add the bot as Admin in the DB Group so it can pin latest backups.',
    ],
  },
  {
    id: 'step-2',
    phase: '02',
    title: 'Push bot.py & requirements.txt to GitHub',
    summary:
      'Upload bot.py and requirements.txt to a GitHub repository. Keep tokens inside Render Environment Variables.',
    commandsOrNotes: [
      'git init && git add bot.py requirements.txt && git commit -m "Add Telegram Excel & DB sync bot"',
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
      'Set TELEGRAM_BOT_TOKEN, ADMIN_CHAT_ID, ALLOWED_CHAT_ID, DB_BACKUP_CHAT_ID, ASSISTANT_TAGS.',
    ],
  },
  {
    id: 'step-4',
    phase: '04',
    title: 'Link Groups & Toggle Reminders Live Inside Telegram',
    summary:
      'You can also bind groups dynamically by typing /setworkgroup inside your work group and /setdbgroup inside your database group, then toggle reminders anytime with /reminder on or /reminder off.',
    commandsOrNotes: [
      '/addassistant @assistant1 @assistant2  -> tags them with warnings at slot open & 5-min cutoff',
      '/reminder on  |  /reminder off  -> instant live switch saved in SQLite',
    ],
  },
  {
    id: 'step-5',
    phase: '05',
    title: 'Keep Render Free Tier Awake (10-Minute Ping)',
    summary:
      'Use UptimeRobot or cron-job.org to ping your Render URL every 10 minutes. Every report submission pushes reports.db to your Database Group and pins it for automatic recovery on Render restarts.',
    commandsOrNotes: [
      'Health URL: https://<your-app-name>.onrender.com/ (returns 200 OK)',
    ],
  },
];
