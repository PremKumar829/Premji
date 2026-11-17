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
    command: '/summary',
    scope: 'Group / DM',
    description: 'Posts a live text breakdown of On Time vs Late counts per team member.',
  },
  {
    command: '/mystatus',
    scope: 'Group / DM',
    description: 'Shows the sender which of the 5 daily slots are Done vs Pending today.',
  },
  {
    command: '/addmember @user',
    scope: 'Admin Only',
    description: 'Adds a team member to the roster so missed slots appear in the Slot Matrix sheet.',
  },
  {
    command: '/removemember @user',
    scope: 'Admin Only',
    description: 'Removes a team member from missed-slot compliance tracking.',
  },
  {
    command: '/backup & /restore',
    scope: 'Admin Only',
    description: 'Exports reports.db to Telegram or restores it by replying /restore to the file.',
  },
];

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
      'Message @BotFather on Telegram, run /newbot to copy your TELEGRAM_BOT_TOKEN, then disable Group Privacy so the bot can detect photo uploads in your team group.',
    commandsOrNotes: [
      'In @BotFather: /mybots -> Select Bot -> Bot Settings -> Group Privacy -> Turn off',
      'Add the bot to your work Telegram group. No Google Cloud or credentials.json needed.',
    ],
  },
  {
    id: 'step-2',
    phase: '02',
    title: 'Push bot.py & requirements.txt to GitHub',
    summary:
      'Create a GitHub repository and upload bot.py and requirements.txt. Keep your bot token strictly inside Render Environment Variables.',
    commandsOrNotes: [
      'git init && git add bot.py requirements.txt && git commit -m "Add Telegram Excel report bot"',
      'git push -u origin main',
    ],
  },
  {
    id: 'step-3',
    phase: '03',
    title: 'Deploy on Render (Free Web Service)',
    summary:
      'In Render Dashboard, click New -> Web Service, connect your GitHub repo, and set Python build and start commands.',
    commandsOrNotes: [
      'Build Command: pip install -r requirements.txt',
      'Start Command: python bot.py',
      'Required Env Var: TELEGRAM_BOT_TOKEN=<your-botfather-token>',
    ],
  },
  {
    id: 'step-4',
    phase: '04',
    title: 'Configure ADMIN_CHAT_ID, Group Reminders & Roster',
    summary:
      'Send /start in the group and in DM to copy your IDs. Set ADMIN_CHAT_ID (for 8:30 PM IST Excel + DB backup DMs) and ALLOWED_CHAT_ID (enables automatic group reminders at each slot).',
    commandsOrNotes: [
      'ADMIN_CHAT_ID=<your-user-id>  |  ALLOWED_CHAT_ID=<your-group-id>',
      'Optional flags: ENABLE_REMINDERS=true  |  PREVENT_DUPLICATE_SLOTS=true',
    ],
  },
  {
    id: 'step-5',
    phase: '05',
    title: 'Keep Render Free Tier Awake & Restore Backup Anytime',
    summary:
      'Ping your Render URL every 10 minutes using UptimeRobot or cron-job.org so polling never sleeps. If Render restarts and resets disk, reply /restore to the latest Telegram backup file.',
    commandsOrNotes: [
      'Health URL: https://<your-app-name>.onrender.com/ (returns 200 OK)',
      'Bot automatically DMs reports_backup_YYYY-MM-DD.db alongside the 8:30 PM Excel report.',
    ],
  },
];
