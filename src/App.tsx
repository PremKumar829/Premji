import { useEffect, useMemo, useState } from 'react';
import {
  AlertTriangle,
  Check,
  CheckCircle2,
  Clock,
  Copy,
  Download,
  FileCode2,
  FileSpreadsheet,
  Plus,
  RotateCcw,
  Search,
  Send,
  Trash2,
  UserMinus,
} from 'lucide-react';
import {
  BOT_COMMANDS,
  BOT_PY_SOURCE,
  INITIAL_SHEET_ROWS,
  INITIAL_TEAM_ROSTER,
  REPORTING_SLOTS,
  REQUIREMENTS_TXT_SOURCE,
  SETUP_STEPS,
  SheetRow,
  evaluateIstSlot,
} from './data/botArtifacts';

type CodeTab = 'bot.py' | 'requirements.txt';
type PreviewTab = 'log' | 'matrix';
type StatusFilter = 'all' | 'on_time' | 'late';

const TIME_PRESETS = [
  { label: '10:08 AM', value: '10:08:24' },
  { label: '10:15 AM (Grace)', value: '10:15:45' },
  { label: '10:16 AM (Late)', value: '10:16:05' },
  { label: '1:05 PM', value: '13:05:12' },
  { label: '3:15 PM (Grace)', value: '15:15:30' },
  { label: '5:10 PM', value: '17:10:00' },
  { label: '7:50 PM (Final)', value: '19:50:18' },
  { label: '8:14 PM (Grace)', value: '20:14:30' },
  { label: '8:18 PM (Late)', value: '20:18:10' },
];

const CHECKLIST_STORAGE_KEY = 'slottrack_excel_render_checklist_v2';

export default function App() {
  const [activeCodeTab, setActiveCodeTab] = useState<CodeTab>('bot.py');
  const [activePreviewTab, setActivePreviewTab] = useState<PreviewTab>('log');
  const [copiedTarget, setCopiedTarget] = useState<string | null>(null);

  // Simulator state
  const [simUsername, setSimUsername] = useState('@vikram_dev');
  const [simFirstName, setSimFirstName] = useState('Vikram');
  const [simTime24, setSimTime24] = useState('10:08:24');
  const [simDate, setSimDate] = useState('2026-11-17');
  const [simCaption, setSimCaption] = useState('Morning site update');
  const [preventDuplicates, setPreventDuplicates] = useState(true);
  const [duplicateNotice, setDuplicateNotice] = useState<string | null>(null);

  // Excel table & roster state
  const [sheetRows, setSheetRows] = useState<SheetRow[]>(INITIAL_SHEET_ROWS);
  const [teamRoster, setTeamRoster] = useState<string[]>(INITIAL_TEAM_ROSTER);
  const [newMemberInput, setNewMemberInput] = useState('');
  const [statusFilter, setStatusFilter] = useState<StatusFilter>('all');
  const [searchQuery, setSearchQuery] = useState('');

  // Setup checklist state persisted locally
  const [completedSteps, setCompletedSteps] = useState<Record<string, boolean>>(() => {
    try {
      const saved = localStorage.getItem(CHECKLIST_STORAGE_KEY);
      return saved ? (JSON.parse(saved) as Record<string, boolean>) : {};
    } catch {
      return {};
    }
  });

  useEffect(() => {
    try {
      localStorage.setItem(CHECKLIST_STORAGE_KEY, JSON.stringify(completedSteps));
    } catch {
      // Ignore storage quota issues
    }
  }, [completedSteps]);

  const liveEvaluation = useMemo(() => evaluateIstSlot(simTime24), [simTime24]);

  const filteredRows = useMemo(() => {
    return sheetRows.filter((row) => {
      if (statusFilter === 'on_time' && row.status !== 'On Time') return false;
      if (statusFilter === 'late' && row.status !== 'Late / Penalty') return false;
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase();
        return (
          row.username.toLowerCase().includes(q) ||
          row.reportingSlot.toLowerCase().includes(q) ||
          row.caption.toLowerCase().includes(q) ||
          row.submissionTime.toLowerCase().includes(q)
        );
      }
      return true;
    });
  }, [sheetRows, statusFilter, searchQuery]);

  const onTimeCount = useMemo(
    () => sheetRows.filter((r) => r.status === 'On Time').length,
    [sheetRows]
  );
  const lateCount = sheetRows.length - onTimeCount;

  // Build live Slot Matrix matching Sheet 3 in bot.py
  const matrixData = useMemo(() => {
    const allMembers = Array.from(
      new Set([...teamRoster, ...sheetRows.map((r) => r.username)])
    ).sort((a, b) => a.localeCompare(b));

    return allMembers.map((member) => {
      const memberRows = sheetRows.filter(
        (r) => r.username.toLowerCase() === member.toLowerCase() && r.date === simDate
      );
      const slotStates = REPORTING_SLOTS.map((slot) => {
        const match = memberRows.find((r) => r.reportingSlot === slot.label);
        return match ? match.submissionTime.replace(' IST', '') : null;
      });
      const onTimeSlots = slotStates.filter(Boolean).length;
      const lateUploads = memberRows.filter((r) => r.status === 'Late / Penalty').length;
      const compliance = Math.round((onTimeSlots / REPORTING_SLOTS.length) * 100);
      return {
        member,
        slotStates,
        onTimeSlots,
        lateUploads,
        compliance,
      };
    });
  }, [teamRoster, sheetRows, simDate]);

  const handleCopy = async (text: string, id: string) => {
    try {
      await navigator.clipboard.writeText(text);
      setCopiedTarget(id);
      window.setTimeout(() => setCopiedTarget((prev) => (prev === id ? null : prev)), 1800);
    } catch {
      // Clipboard fallback ignored
    }
  };

  const handleDownloadFile = (filename: string, content: string) => {
    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
  };

  const handleSimulateUpload = (e: React.FormEvent) => {
    e.preventDefault();
    const cleanUser = simUsername.trim() || simFirstName.trim() || 'Team Member';
    const targetDate = simDate || '2026-11-17';

    if (preventDuplicates && liveEvaluation.reportingSlot !== 'Late / Outside Slot') {
      const existing = sheetRows.find(
        (r) =>
          r.username.toLowerCase() === cleanUser.toLowerCase() &&
          r.date === targetDate &&
          r.reportingSlot === liveEvaluation.reportingSlot
      );
      if (existing) {
        setDuplicateNotice(
          `Duplicate skipped for ${cleanUser}: ${liveEvaluation.reportingSlot} was already recorded today at ${existing.submissionTime}.`
        );
        return;
      }
    }

    setDuplicateNotice(null);
    const newRow: SheetRow = {
      id: `row-${Date.now()}`,
      username: cleanUser,
      submissionTime: liveEvaluation.formattedTimeIst,
      reportingSlot: liveEvaluation.reportingSlot,
      date: targetDate,
      status: liveEvaluation.status,
      caption: simCaption.trim(),
    };
    setSheetRows((prev) => [newRow, ...prev]);

    if (!teamRoster.some((m) => m.toLowerCase() === cleanUser.toLowerCase())) {
      setTeamRoster((prev) => [...prev, cleanUser]);
    }
  };

  const handleAddMember = (e: React.FormEvent) => {
    e.preventDefault();
    const clean = newMemberInput.trim();
    if (!clean) return;
    if (!teamRoster.some((m) => m.toLowerCase() === clean.toLowerCase())) {
      setTeamRoster((prev) => [...prev, clean]);
    }
    setNewMemberInput('');
  };

  const handleRemoveMember = (member: string) => {
    setTeamRoster((prev) => prev.filter((m) => m !== member));
  };

  const handleExportCsv = () => {
    const headers = ['Username', 'Submission Time', 'Reporting Slot', 'Date', 'Status', 'Caption'];
    const lines = [
      headers.join(','),
      ...sheetRows.map((row) =>
        [row.username, row.submissionTime, row.reportingSlot, row.date, row.status, row.caption]
          .map((cell) => `"${cell.replace(/"/g, '""')}"`)
          .join(',')
      ),
    ];
    handleDownloadFile(`team_report_${simDate || '2026-11-17'}.csv`, lines.join('\n'));
  };

  const toggleChecklistStep = (stepId: string) => {
    setCompletedSteps((prev) => ({
      ...prev,
      [stepId]: !prev[stepId],
    }));
  };

  const completedCount = SETUP_STEPS.filter((s) => completedSteps[s.id]).length;
  const activeCodeContent =
    activeCodeTab === 'bot.py' ? BOT_PY_SOURCE : REQUIREMENTS_TXT_SOURCE;

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900">
      {/* Top Bar Contract: 3 zones separated by gap-8 */}
      <header className="sticky top-0 z-30 border-b border-slate-200 bg-white/95 backdrop-blur-sm">
        <div className="mx-auto flex max-w-7xl items-center justify-between gap-8 px-6 py-3.5">
          <a
            href="#top"
            className="text-lg font-bold tracking-tight text-slate-900 whitespace-nowrap shrink-0"
          >
            SlotTrack
          </a>

          <nav className="hidden md:flex items-center gap-7 text-sm font-medium text-slate-600">
            <a
              href="#simulator"
              className="hover:text-slate-900 transition-colors whitespace-nowrap shrink-0"
            >
              Slot Simulator
            </a>
            <a
              href="#excel-preview"
              className="hover:text-slate-900 transition-colors whitespace-nowrap shrink-0"
            >
              Excel &amp; Matrix
            </a>
            <a
              href="#commands"
              className="hover:text-slate-900 transition-colors whitespace-nowrap shrink-0"
            >
              Commands
            </a>
            <a
              href="#bot-code"
              className="hover:text-slate-900 transition-colors whitespace-nowrap shrink-0"
            >
              Bot Script
            </a>
            <a
              href="#render-setup"
              className="hover:text-slate-900 transition-colors whitespace-nowrap shrink-0"
            >
              Render Guide
            </a>
          </nav>

          <div className="flex items-center shrink-0">
            <button
              type="button"
              onClick={() => handleDownloadFile('bot.py', BOT_PY_SOURCE)}
              className="inline-flex items-center gap-2 rounded-lg bg-sky-600 px-4 py-2 text-sm font-semibold text-white hover:bg-sky-700 focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-sky-600 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              <Download className="h-4 w-4" aria-hidden="true" />
              <span>Download bot.py</span>
            </button>
          </div>
        </div>
      </header>

      <main id="top" className="mx-auto max-w-7xl px-6 py-10 space-y-12">
        {/* Hero / Overview Header */}
        <section className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
          <div className="lg:col-span-7 space-y-4">
            <p className="text-xs font-medium text-slate-500">
              GitHub + Render Free Tier · 4-Sheet Excel (.xlsx) · Slot Reminders · Duplicate Guard
            </p>
            <h1 className="text-3xl sm:text-4xl font-bold tracking-tight text-slate-900 [text-wrap:balance]">
              Telegram Team Photo Report Bot with Automated Excel &amp; Compliance Matrix
            </h1>
            <p className="text-base leading-relaxed text-slate-600 max-w-2xl">
              Zero Google Cloud setup. Tracks group photo reports with captions in IST, prevents
              accidental duplicate slot entries, nudges the group at each slot window, and sends a
              4-sheet formatted <code className="font-mono text-sm text-slate-800">.xlsx</code>{' '}
              workbook plus database backup straight to your Telegram DM at 8:30 PM IST.
            </p>
            <div className="pt-2 flex flex-wrap items-center gap-6 text-sm text-slate-700">
              <div>
                <span className="text-slate-500">Deploy:</span>{' '}
                <span className="font-mono font-medium">GitHub + Render Free</span>
              </div>
              <span className="text-slate-300" aria-hidden="true">
                |
              </span>
              <div>
                <span className="text-slate-500">Workbook Sheets:</span>{' '}
                <span className="font-mono text-xs font-medium">
                  Reports · Member Summary · Slot Matrix · All History
                </span>
              </div>
            </div>
          </div>

          {/* Designated IST Reporting Schedule Reference */}
          <div className="lg:col-span-5 rounded-xl border border-slate-200 bg-white p-5">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <h2 className="text-sm font-semibold text-slate-900">
                Designated IST Windows &amp; Reminders
              </h2>
              <span className="text-xs font-mono text-slate-500">+15m grace</span>
            </div>
            <div className="divide-y divide-slate-100">
              {REPORTING_SLOTS.map((slot) => (
                <div
                  key={slot.id}
                  className="flex items-center justify-between py-2.5 text-sm gap-4"
                >
                  <span className="font-medium text-slate-800">{slot.label}</span>
                  <span className="font-mono text-xs text-slate-600 tabular-nums text-right">
                    {slot.windowDisplay}
                  </span>
                </div>
              ))}
            </div>
          </div>
        </section>

        {/* Section 1: Interactive IST Slot Simulator & Duplicate Check */}
        <section id="simulator" className="scroll-mt-20 space-y-4">
          <div>
            <h2 className="text-xl font-bold tracking-tight text-slate-900">
              01. IST Slot, Caption &amp; Duplicate Guard Simulator
            </h2>
            <p className="text-sm text-slate-600">
              Test timestamps, photo captions, and duplicate-slot protection against the exact logic
              in <code className="font-mono text-xs text-slate-800">bot.py</code>.
            </p>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-stretch">
            <form
              onSubmit={handleSimulateUpload}
              className="lg:col-span-7 rounded-xl border border-slate-200 bg-white p-6 space-y-5"
            >
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div>
                  <label
                    htmlFor="sim-username"
                    className="block text-xs font-semibold text-slate-700 mb-1.5"
                  >
                    Telegram Username or Name
                  </label>
                  <input
                    id="sim-username"
                    type="text"
                    value={simUsername}
                    onChange={(e) => setSimUsername(e.target.value)}
                    placeholder="@username"
                    className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-900 focus:border-sky-600 focus:outline-none"
                  />
                </div>
                <div>
                  <label
                    htmlFor="sim-firstname"
                    className="block text-xs font-semibold text-slate-700 mb-1.5"
                  >
                    First Name (for Reply)
                  </label>
                  <input
                    id="sim-firstname"
                    type="text"
                    value={simFirstName}
                    onChange={(e) => setSimFirstName(e.target.value)}
                    placeholder="Vikram"
                    className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-900 focus:border-sky-600 focus:outline-none"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                <div>
                  <label
                    htmlFor="sim-time"
                    className="block text-xs font-semibold text-slate-700 mb-1.5"
                  >
                    Time (IST, 24h)
                  </label>
                  <input
                    id="sim-time"
                    type="time"
                    step="1"
                    value={simTime24}
                    onChange={(e) => setSimTime24(e.target.value || '10:00:00')}
                    className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 font-mono text-sm text-slate-900 focus:border-sky-600 focus:outline-none"
                  />
                </div>
                <div>
                  <label
                    htmlFor="sim-date"
                    className="block text-xs font-semibold text-slate-700 mb-1.5"
                  >
                    Date (IST)
                  </label>
                  <input
                    id="sim-date"
                    type="date"
                    value={simDate}
                    onChange={(e) => setSimDate(e.target.value)}
                    className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 font-mono text-sm text-slate-900 focus:border-sky-600 focus:outline-none"
                  />
                </div>
                <div>
                  <label
                    htmlFor="sim-caption"
                    className="block text-xs font-semibold text-slate-700 mb-1.5"
                  >
                    Photo Caption (Optional)
                  </label>
                  <input
                    id="sim-caption"
                    type="text"
                    value={simCaption}
                    onChange={(e) => setSimCaption(e.target.value)}
                    placeholder="Optional note..."
                    className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-900 focus:border-sky-600 focus:outline-none"
                  />
                </div>
              </div>

              <div>
                <p className="text-xs font-semibold text-slate-700 mb-2">Quick Edge-Case Times</p>
                <div className="flex flex-wrap gap-2">
                  {TIME_PRESETS.map((preset) => {
                    const isActive = simTime24 === preset.value;
                    return (
                      <button
                        key={preset.value}
                        type="button"
                        onClick={() => {
                          setSimTime24(preset.value);
                          setDuplicateNotice(null);
                        }}
                        className={`rounded-lg border px-2.5 py-1.5 text-xs font-medium transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                          isActive
                            ? 'border-sky-600 bg-sky-50 text-sky-700'
                            : 'border-slate-200 bg-slate-50 text-slate-700 hover:bg-slate-100'
                        }`}
                      >
                        {preset.label}
                      </button>
                    );
                  })}
                </div>
              </div>

              <div className="pt-2 flex flex-wrap items-center justify-between gap-4 border-t border-slate-100">
                <label className="inline-flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={preventDuplicates}
                    onChange={(e) => {
                      setPreventDuplicates(e.target.checked);
                      setDuplicateNotice(null);
                    }}
                    className="h-4 w-4 rounded border-slate-300 text-sky-600 focus:ring-sky-600"
                  />
                  <span>Prevent duplicate submission in same slot</span>
                </label>

                <button
                  type="submit"
                  className="inline-flex items-center gap-2 rounded-lg bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-800 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                >
                  <Send className="h-4 w-4" aria-hidden="true" />
                  <span>Simulate Photo Upload</span>
                </button>
              </div>
            </form>

            {/* Telegram Bot Reply Preview */}
            <div className="lg:col-span-5 rounded-xl border border-slate-200 bg-white p-6 flex flex-col justify-between space-y-4">
              <div className="space-y-3">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-semibold text-slate-500">
                    Telegram Group Reply Preview
                  </span>
                  <span
                    className={`inline-flex items-center gap-1.5 text-xs font-semibold ${
                      duplicateNotice
                        ? 'text-amber-700'
                        : liveEvaluation.status === 'On Time'
                          ? 'text-emerald-700'
                          : 'text-red-600'
                    }`}
                  >
                    {duplicateNotice ? (
                      <>
                        <AlertTriangle className="h-3.5 w-3.5" aria-hidden="true" />
                        <span>Duplicate Skipped</span>
                      </>
                    ) : liveEvaluation.status === 'On Time' ? (
                      <>
                        <CheckCircle2 className="h-3.5 w-3.5" aria-hidden="true" />
                        <span>{liveEvaluation.status}</span>
                      </>
                    ) : (
                      <>
                        <AlertTriangle className="h-3.5 w-3.5" aria-hidden="true" />
                        <span>{liveEvaluation.status}</span>
                      </>
                    )}
                  </span>
                </div>

                <div className="rounded-lg border border-slate-200 bg-slate-50 p-4 space-y-3">
                  <div className="text-xs text-slate-500 border-b border-slate-200 pb-2 flex items-center justify-between">
                    <span>
                      Photo from{' '}
                      <span className="font-semibold text-slate-800">
                        {simFirstName.trim() || 'User'}
                      </span>{' '}
                      ({simUsername.trim() || '@user'})
                    </span>
                    <span className="font-mono text-slate-600 tabular-nums">
                      {liveEvaluation.formattedTimeIst}
                    </span>
                  </div>

                  <pre className="whitespace-pre-wrap font-sans text-sm leading-relaxed text-slate-800">
                    {duplicateNotice
                      ? duplicateNotice
                      : liveEvaluation.status === 'On Time'
                        ? `Report recorded for ${simFirstName.trim() || simUsername} (${
                            simUsername.trim() || 'Unknown User'
                          }).\nSlot: ${liveEvaluation.reportingSlot}\nTime: ${
                            liveEvaluation.formattedTimeIst
                          } | Date: ${simDate}\nStatus: ${liveEvaluation.status}${
                            simCaption.trim() ? `\nNote: ${simCaption.trim()}` : ''
                          }`
                        : `Report recorded for ${simFirstName.trim() || simUsername} (${
                            simUsername.trim() || 'Unknown User'
                          }), outside scheduled windows.\nSlot: ${
                            liveEvaluation.reportingSlot
                          }\nTime: ${
                            liveEvaluation.formattedTimeIst
                          } | Date: ${simDate}\nStatus: ${liveEvaluation.status}${
                            simCaption.trim() ? `\nNote: ${simCaption.trim()}` : ''
                          }`}
                  </pre>
                </div>
              </div>

              <div className="text-xs text-slate-500 pt-2 border-t border-slate-100 flex items-center gap-2">
                <Clock className="h-3.5 w-3.5 text-slate-400 shrink-0" aria-hidden="true" />
                <span>
                  Auto-sends 4-sheet <code className="font-mono text-slate-700">.xlsx</code> + DB
                  backup at 8:30 PM IST.
                </span>
              </div>
            </div>
          </div>
        </section>

        {/* Section 2: Excel (.xlsx) Submissions Log & Slot Compliance Matrix */}
        <section id="excel-preview" className="scroll-mt-20 space-y-4">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div>
              <h2 className="text-xl font-bold tracking-tight text-slate-900">
                02. Excel Workbook (.xlsx) &amp; Team Compliance Matrix
              </h2>
              <p className="text-sm text-slate-600">
                Switch between the raw submissions log and the automated{' '}
                <code className="font-mono text-xs text-slate-800">Slot Matrix</code> sheet
                tracking missed slots per team member.
              </p>
            </div>

            <div className="flex flex-wrap items-center gap-3">
              <div className="inline-flex items-center gap-1 rounded-lg bg-slate-200/70 p-1">
                <button
                  type="button"
                  onClick={() => setActivePreviewTab('log')}
                  className={`rounded-md px-3 py-1.5 text-xs font-semibold transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                    activePreviewTab === 'log'
                      ? 'bg-white text-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Submissions Log
                </button>
                <button
                  type="button"
                  onClick={() => setActivePreviewTab('matrix')}
                  className={`rounded-md px-3 py-1.5 text-xs font-semibold transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                    activePreviewTab === 'matrix'
                      ? 'bg-white text-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Slot Matrix &amp; Roster
                </button>
              </div>

              <button
                type="button"
                onClick={handleExportCsv}
                disabled={sheetRows.length === 0}
                className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50 disabled:opacity-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              >
                <FileSpreadsheet className="h-3.5 w-3.5" aria-hidden="true" />
                <span>Export CSV</span>
              </button>
            </div>
          </div>

          {activePreviewTab === 'log' ? (
            <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-200 px-4 py-3 bg-slate-50/70">
                <div className="flex flex-wrap items-center gap-3">
                  <div className="inline-flex items-center gap-1 rounded-lg bg-slate-200/70 p-1">
                    <button
                      type="button"
                      onClick={() => setStatusFilter('all')}
                      className={`rounded-md px-3 py-1.5 text-xs font-semibold transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                        statusFilter === 'all'
                          ? 'bg-white text-slate-900 shadow-xs'
                          : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      All ({sheetRows.length})
                    </button>
                    <button
                      type="button"
                      onClick={() => setStatusFilter('on_time')}
                      className={`rounded-md px-3 py-1.5 text-xs font-semibold transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                        statusFilter === 'on_time'
                          ? 'bg-white text-slate-900 shadow-xs'
                          : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      On Time ({onTimeCount})
                    </button>
                    <button
                      type="button"
                      onClick={() => setStatusFilter('late')}
                      className={`rounded-md px-3 py-1.5 text-xs font-semibold transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                        statusFilter === 'late'
                          ? 'bg-white text-slate-900 shadow-xs'
                          : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      Late ({lateCount})
                    </button>
                  </div>

                  <button
                    type="button"
                    onClick={() => setSheetRows([])}
                    disabled={sheetRows.length === 0}
                    className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50 disabled:opacity-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <Trash2 className="h-3.5 w-3.5" aria-hidden="true" />
                    <span>Clear</span>
                  </button>
                </div>

                <div className="relative w-full sm:w-64">
                  <Search
                    className="pointer-events-none absolute left-3 top-1/2 h-3.5 w-3.5 -translate-y-1/2 text-slate-400"
                    aria-hidden="true"
                  />
                  <input
                    type="text"
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    placeholder="Filter username, slot, caption..."
                    aria-label="Filter report rows"
                    className="w-full rounded-lg border border-slate-300 bg-white pl-8 pr-3 py-1.5 text-xs text-slate-900 focus:border-sky-600 focus:outline-none"
                  />
                </div>
              </div>

              {filteredRows.length === 0 ? (
                <div className="px-6 py-12 text-center space-y-3">
                  <p className="text-sm font-medium text-slate-700">
                    {sheetRows.length === 0
                      ? 'No report rows logged yet.'
                      : 'No rows match the active filter.'}
                  </p>
                  <button
                    type="button"
                    onClick={() => {
                      setSheetRows(INITIAL_SHEET_ROWS);
                      setStatusFilter('all');
                      setSearchQuery('');
                    }}
                    className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50 cursor-pointer"
                  >
                    <RotateCcw className="h-3.5 w-3.5" aria-hidden="true" />
                    <span>Restore Sample Rows</span>
                  </button>
                </div>
              ) : (
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse text-left text-sm">
                    <thead>
                      <tr className="border-b border-slate-200 bg-slate-50 text-xs font-semibold text-slate-600">
                        <th className="px-4 py-2.5">Username</th>
                        <th className="px-4 py-2.5">Submission Time</th>
                        <th className="px-4 py-2.5">Reporting Slot</th>
                        <th className="px-4 py-2.5">Date</th>
                        <th className="px-4 py-2.5">Status</th>
                        <th className="px-4 py-2.5">Caption</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {filteredRows.map((row) => (
                        <tr key={row.id} className="hover:bg-slate-50/80 transition-colors">
                          <td className="px-4 py-2.5 font-medium text-slate-900 whitespace-nowrap">
                            {row.username}
                          </td>
                          <td className="px-4 py-2.5 font-mono text-xs text-slate-700 tabular-nums whitespace-nowrap">
                            {row.submissionTime}
                          </td>
                          <td className="px-4 py-2.5 text-slate-700 whitespace-nowrap">
                            {row.reportingSlot}
                          </td>
                          <td className="px-4 py-2.5 font-mono text-xs text-slate-600 tabular-nums whitespace-nowrap">
                            {row.date}
                          </td>
                          <td className="px-4 py-2.5 whitespace-nowrap">
                            <span
                              className={`inline-flex items-center gap-1.5 text-xs font-semibold ${
                                row.status === 'On Time' ? 'text-emerald-700' : 'text-red-600'
                              }`}
                            >
                              {row.status === 'On Time' ? (
                                <CheckCircle2 className="h-3.5 w-3.5" aria-hidden="true" />
                              ) : (
                                <AlertTriangle className="h-3.5 w-3.5" aria-hidden="true" />
                              )}
                              <span>{row.status}</span>
                            </span>
                          </td>
                          <td className="px-4 py-2.5 text-xs text-slate-600 max-w-xs truncate">
                            {row.caption || '—'}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          ) : (
            <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
              {/* Roster add bar matching /addmember @username */}
              <form
                onSubmit={handleAddMember}
                className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-200 px-4 py-3 bg-slate-50/70"
              >
                <div className="text-xs text-slate-600">
                  Simulated <code className="font-mono text-slate-800">/addmember</code> roster for{' '}
                  <span className="font-mono font-semibold text-slate-800">{simDate}</span>
                </div>
                <div className="flex items-center gap-2">
                  <input
                    type="text"
                    value={newMemberInput}
                    onChange={(e) => setNewMemberInput(e.target.value)}
                    placeholder="@teammate"
                    className="rounded-lg border border-slate-300 bg-white px-3 py-1.5 text-xs text-slate-900 focus:border-sky-600 focus:outline-none"
                  />
                  <button
                    type="submit"
                    className="inline-flex items-center gap-1.5 rounded-lg bg-slate-900 px-3 py-1.5 text-xs font-semibold text-white hover:bg-slate-800 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <Plus className="h-3.5 w-3.5" aria-hidden="true" />
                    <span>Add Member</span>
                  </button>
                </div>
              </form>

              <div className="overflow-x-auto">
                <table className="w-full border-collapse text-left text-xs">
                  <thead>
                    <tr className="border-b border-slate-200 bg-slate-50 font-semibold text-slate-600">
                      <th className="px-4 py-2.5">Username</th>
                      {REPORTING_SLOTS.map((s) => (
                        <th key={s.id} className="px-3 py-2.5 whitespace-nowrap">
                          {s.shortLabel}
                        </th>
                      ))}
                      <th className="px-3 py-2.5 text-right">Late</th>
                      <th className="px-3 py-2.5 text-right">Compliance</th>
                      <th className="px-3 py-2.5 text-right">Roster</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {matrixData.map((row) => (
                      <tr key={row.member} className="hover:bg-slate-50/80 transition-colors">
                        <td className="px-4 py-2.5 font-medium text-sm text-slate-900 whitespace-nowrap">
                          {row.member}
                        </td>
                        {row.slotStates.map((timeStr, idx) => (
                          <td key={idx} className="px-3 py-2.5 font-mono whitespace-nowrap">
                            {timeStr ? (
                              <span className="text-emerald-700 font-semibold">{timeStr}</span>
                            ) : (
                              <span className="text-red-600">Missed</span>
                            )}
                          </td>
                        ))}
                        <td className="px-3 py-2.5 font-mono text-right tabular-nums text-slate-700">
                          {row.lateUploads}
                        </td>
                        <td className="px-3 py-2.5 font-mono text-right tabular-nums font-semibold text-slate-900">
                          {row.compliance}%
                        </td>
                        <td className="px-3 py-2.5 text-right">
                          <button
                            type="button"
                            onClick={() => handleRemoveMember(row.member)}
                            title="Remove from roster"
                            aria-label={`Remove ${row.member}`}
                            className="inline-flex items-center justify-center rounded-md p-1 text-slate-400 hover:text-red-600 hover:bg-red-50 transition-colors cursor-pointer"
                          >
                            <UserMinus className="h-3.5 w-3.5" aria-hidden="true" />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </section>

        {/* Section 3: Telegram Bot Commands Reference */}
        <section id="commands" className="scroll-mt-20 space-y-4">
          <div>
            <h2 className="text-xl font-bold tracking-tight text-slate-900">
              03. Built-In Bot Commands
            </h2>
            <p className="text-sm text-slate-600">
              Manage your team roster, trigger instant Excel exports, and back up your database
              directly inside Telegram.
            </p>
          </div>

          <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
            <div className="divide-y divide-slate-100">
              {BOT_COMMANDS.map((item) => (
                <div
                  key={item.command}
                  className="px-5 py-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-2 sm:gap-6"
                >
                  <div className="flex items-center gap-3 min-w-0">
                    <code className="font-mono text-xs font-semibold text-sky-700 whitespace-nowrap">
                      {item.command}
                    </code>
                    <span className="text-xs text-slate-400" aria-hidden="true">
                      ·
                    </span>
                    <span className="text-xs font-medium text-slate-500 whitespace-nowrap">
                      {item.scope}
                    </span>
                  </div>
                  <p className="text-sm text-slate-700 sm:text-right">{item.description}</p>
                </div>
              ))}
            </div>
          </div>
        </section>

        {/* Section 4: Complete Python Deliverables (bot.py + requirements.txt) */}
        <section id="bot-code" className="scroll-mt-20 space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div>
              <h2 className="text-xl font-bold tracking-tight text-slate-900">
                04. Complete Python Script &amp; Pip Packages
              </h2>
              <p className="text-sm text-slate-600">
                Upload only <code className="font-mono text-xs text-slate-800">bot.py</code> and{' '}
                <code className="font-mono text-xs text-slate-800">requirements.txt</code> to your
                GitHub repository.
              </p>
            </div>

            <div className="flex items-center gap-2">
              <div className="inline-flex items-center gap-1 rounded-lg bg-slate-200/70 p-1">
                <button
                  type="button"
                  onClick={() => setActiveCodeTab('bot.py')}
                  className={`rounded-md px-3 py-1.5 text-xs font-semibold transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                    activeCodeTab === 'bot.py'
                      ? 'bg-white text-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  bot.py
                </button>
                <button
                  type="button"
                  onClick={() => setActiveCodeTab('requirements.txt')}
                  className={`rounded-md px-3 py-1.5 text-xs font-semibold transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                    activeCodeTab === 'requirements.txt'
                      ? 'bg-white text-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  requirements.txt
                </button>
              </div>

              <button
                type="button"
                onClick={() => handleCopy(activeCodeContent, activeCodeTab)}
                className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs font-semibold text-slate-700 hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              >
                {copiedTarget === activeCodeTab ? (
                  <>
                    <Check className="h-3.5 w-3.5 text-emerald-600" aria-hidden="true" />
                    <span>Copied</span>
                  </>
                ) : (
                  <>
                    <Copy className="h-3.5 w-3.5" aria-hidden="true" />
                    <span>Copy</span>
                  </>
                )}
              </button>

              <button
                type="button"
                onClick={() => handleDownloadFile(activeCodeTab, activeCodeContent)}
                className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs font-semibold text-slate-700 hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              >
                <FileCode2 className="h-3.5 w-3.5" aria-hidden="true" />
                <span>Download</span>
              </button>
            </div>
          </div>

          <div className="rounded-xl border border-slate-200 bg-slate-950 text-slate-100 overflow-hidden">
            <div className="flex items-center justify-between border-b border-slate-800 px-4 py-2.5 text-xs text-slate-400">
              <span className="font-mono">{activeCodeTab}</span>
              <span className="font-mono tabular-nums">
                {activeCodeContent.split('\n').length} lines
              </span>
            </div>
            <pre className="max-h-[540px] overflow-auto p-5 font-mono text-xs leading-6 text-slate-200">
              <code>{activeCodeContent}</code>
            </pre>
          </div>
        </section>

        {/* Section 5: GitHub + Render Deployment Checklist */}
        <section id="render-setup" className="scroll-mt-20 space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-end justify-between gap-2">
            <div>
              <h2 className="text-xl font-bold tracking-tight text-slate-900">
                05. GitHub &amp; Render Free Deployment Guide
              </h2>
              <p className="text-sm text-slate-600">
                Step-by-step deployment using only GitHub and Render (no Google Cloud setup).
              </p>
            </div>
            <span className="text-xs font-mono text-slate-600 tabular-nums">
              {completedCount} of {SETUP_STEPS.length} steps checked
            </span>
          </div>

          <div className="rounded-xl border border-slate-200 bg-white divide-y divide-slate-200">
            {SETUP_STEPS.map((step) => {
              const isDone = Boolean(completedSteps[step.id]);
              return (
                <div key={step.id} className="p-5 flex items-start gap-4">
                  <button
                    type="button"
                    onClick={() => toggleChecklistStep(step.id)}
                    aria-pressed={isDone}
                    aria-label={`Mark step ${step.phase} complete`}
                    className={`mt-0.5 flex h-6 w-6 shrink-0 items-center justify-center rounded-md border transition-colors cursor-pointer ${
                      isDone
                        ? 'border-sky-600 bg-sky-600 text-white'
                        : 'border-slate-300 bg-white text-transparent hover:border-slate-400'
                    }`}
                  >
                    <Check className="h-3.5 w-3.5" aria-hidden="true" />
                  </button>

                  <div className="space-y-2 min-w-0 flex-1">
                    <div className="flex flex-wrap items-baseline gap-2">
                      <span className="font-mono text-xs font-semibold text-sky-700">
                        {step.phase}.
                      </span>
                      <h3
                        className={`text-sm font-semibold ${
                          isDone ? 'text-slate-500 line-through' : 'text-slate-900'
                        }`}
                      >
                        {step.title}
                      </h3>
                    </div>
                    <p className="text-sm text-slate-600 leading-relaxed">{step.summary}</p>
                    <div className="space-y-1.5 pt-1">
                      {step.commandsOrNotes.map((line, idx) => (
                        <div
                          key={idx}
                          className="rounded-lg bg-slate-50 border border-slate-200/80 px-3 py-2 font-mono text-xs text-slate-800 overflow-x-auto"
                        >
                          {line}
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </section>
      </main>

      <footer className="border-t border-slate-200 bg-white mt-12">
        <div className="mx-auto max-w-7xl px-6 py-6 flex flex-col sm:flex-row items-center justify-between gap-4 text-xs text-slate-500">
          <span>SlotTrack Bot · Telegram Excel Report Automation (GitHub + Render)</span>
          <div className="flex items-center gap-4">
            <button
              type="button"
              onClick={() => handleDownloadFile('bot.py', BOT_PY_SOURCE)}
              className="hover:text-slate-900 transition-colors cursor-pointer"
            >
              Download bot.py
            </button>
            <button
              type="button"
              onClick={() => handleDownloadFile('requirements.txt', REQUIREMENTS_TXT_SOURCE)}
              className="hover:text-slate-900 transition-colors cursor-pointer"
            >
              Download requirements.txt
            </button>
          </div>
        </div>
      </footer>
    </div>
  );
}
