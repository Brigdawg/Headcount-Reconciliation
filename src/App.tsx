import { useMemo, useState } from 'react';
import {
  Area,
  Bar,
  CartesianGrid,
  ComposedChart,
  Legend,
  Line,
  ReferenceLine,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import {
  AS_OF,
  BOARD_AS_OF,
  CLOSE_AS_OF,
  COMPANY,
  DEPT_LIST,
  OUTLOOK_RANGES,
  VIEWING_AS_OF,
  auditTrail,
  departments,
  getBridgeForDept,
  getMissingSeats,
  getMonthsForDept,
  historySummary,
  roles,
  snapshots,
  summarizeScope,
  tickets,
  typeLabel,
  type AppView,
  type ApprovalStatus,
  type AuditEvent,
  type BridgeStep,
  type Dept,
  type DeptBudget,
  type OutlookRange,
  type Persona,
  type RoleRow,
  type RoleType,
  type TicketType,
} from './data/mockData';

type ExportSheet = 'summary' | 'departments' | 'roles' | 'tickets';
type ManualMode = 'correct' | 'departure' | 'add';
type ManualAdj = { id: string; ts: string; kind: string; summary: string; dept: Dept };
type CorrectField = 'actual' | 'planned' | 'board';
type DepartureDisposition = 'backfill' | 'not_backfilling' | 'pivot' | 'close';
type AddSeatType = 'new' | 'backfill' | 'not_backfilling' | 'pivot' | 'close';

function statusLabel(s: string) {
  return s.replace(/_/g, ' ');
}

function fmtMoney(n: number) {
  if (n >= 1_000_000) return `$${(n / 1_000_000).toFixed(1)}M`;
  return `$${(n / 1_000).toFixed(0)}K`;
}

function Variance({ n, suffix = ' vs board' }: { n: number; suffix?: string }) {
  const cls = n < 0 ? 'neg' : n > 0 ? 'pos' : 'neutral';
  const sign = n > 0 ? '+' : '';
  return (
    <span className={`delta ${cls}`}>
      {sign}
      {n}
      {suffix}
    </span>
  );
}

/** Readable bridge rows: start → deltas with running total → end */
function buildBridgeRows(bridge: BridgeStep[]) {
  let running = 0;
  const maxAbs = Math.max(
    ...bridge.filter((s) => s.kind === 'neg' || s.kind === 'pos').map((s) => Math.abs(s.value)),
    1,
  );
  return bridge.map((step) => {
    if (step.kind === 'start') {
      running = step.value;
      return { ...step, running, barPct: 0 };
    }
    if (step.kind === 'end') {
      running = step.value;
      return { ...step, running, barPct: 0 };
    }
    running += step.value;
    return {
      ...step,
      running,
      barPct: (Math.abs(step.value) / maxAbs) * 100,
    };
  });
}

const NAV: { id: AppView; label: string }[] = [
  { id: 'home', label: 'Home' },
  { id: 'departments', label: 'Departments' },
  { id: 'approvals', label: 'Approvals' },
  { id: 'roles', label: 'Roles' },
  { id: 'audit', label: 'Audit' },
];

export default function App() {
  const [view, setView] = useState<AppView>('home');
  const [persona, setPersona] = useState<Persona>('finance');
  const [dept, setDept] = useState<Dept | 'all'>('all');
  const [chartMode, setChartMode] = useState<'outlook' | 'bridge'>('bridge');
  const [outlookRange, setOutlookRange] = useState<OutlookRange>('h2');
  const [ticketFilter, setTicketFilter] = useState<'all' | ApprovalStatus>('all');
  const [roleTypeFilter, setRoleTypeFilter] = useState<'all' | RoleType>('all');
  const [selectedTicket, setSelectedTicket] = useState(tickets[0].id);
  const [selectedRole, setSelectedRole] = useState(roles[0].id);
  const [snapshotId, setSnapshotId] = useState(snapshots[snapshots.length - 1].id);
  const [toast, setToast] = useState<string | null>(null);
  const [exportOpen, setExportOpen] = useState(false);
  const [exportSheet, setExportSheet] = useState<ExportSheet>('summary');

  const [liveDepts, setLiveDepts] = useState<DeptBudget[]>(() => departments.map((d) => ({ ...d })));
  const [liveRoles, setLiveRoles] = useState<RoleRow[]>(() => roles.map((r) => ({ ...r })));
  const [liveTickets, setLiveTickets] = useState(() => [...tickets]);
  const [liveAudit, setLiveAudit] = useState<AuditEvent[]>(() => [...auditTrail]);
  const [manualAdjs, setManualAdjs] = useState<ManualAdj[]>([]);

  const [manualOpen, setManualOpen] = useState(false);
  const [manualMode, setManualMode] = useState<ManualMode>('correct');

  const [correctDept, setCorrectDept] = useState<Dept>('Engineering');
  const [correctField, setCorrectField] = useState<CorrectField>('actual');
  const [correctValue, setCorrectValue] = useState('');
  const [correctReason, setCorrectReason] = useState('');

  const [depTitle, setDepTitle] = useState('');
  const [depDept, setDepDept] = useState<Dept>('Engineering');
  const [depPerson, setDepPerson] = useState('');
  const [depDisposition, setDepDisposition] = useState<DepartureDisposition>('backfill');
  const [depPivotTo, setDepPivotTo] = useState('');
  const [depDate, setDepDate] = useState(VIEWING_AS_OF);
  const [depNote, setDepNote] = useState('');

  const [addTitle, setAddTitle] = useState('');
  const [addDept, setAddDept] = useState<Dept>('Engineering');
  const [addType, setAddType] = useState<AddSeatType>('new');
  const [addBoard, setAddBoard] = useState(1);
  const [addPlanned, setAddPlanned] = useState(1);
  const [addActual, setAddActual] = useState(0);
  const [addNote, setAddNote] = useState('');

  const scope = useMemo(() => summarizeScope(dept, liveDepts, liveRoles), [dept, liveDepts, liveRoles]);
  const months = useMemo(() => getMonthsForDept(dept, outlookRange), [dept, outlookRange]);
  const hist = useMemo(() => historySummary(months), [months]);
  const bridge = useMemo(() => getBridgeForDept(dept, liveDepts), [dept, liveDepts]);
  const bridgeRows = useMemo(() => buildBridgeRows(bridge), [bridge]);
  const missingSeats = useMemo(() => getMissingSeats(dept, liveRoles), [dept, liveRoles]);

  const outlookData = useMemo(() => {
    return months.map((m) => ({
      month: m.month,
      Board: m.board,
      Planned: m.planned,
      Actual: m.actual,
      isFuture: m.isFuture,
    }));
  }, [months]);

  const filteredTickets = useMemo(() => {
    return liveTickets.filter((t) => {
      if (dept !== 'all' && t.dept !== dept) return false;
      if (ticketFilter !== 'all' && t.status !== ticketFilter) return false;
      return true;
    });
  }, [dept, ticketFilter, liveTickets]);

  const filteredRoles = useMemo(() => {
    return liveRoles.filter((r) => {
      if (dept !== 'all' && r.dept !== dept) return false;
      if (roleTypeFilter !== 'all' && r.type !== roleTypeFilter) return false;
      return true;
    });
  }, [dept, roleTypeFilter, liveRoles]);

  const filteredDepts = useMemo(() => {
    if (dept === 'all') return liveDepts;
    return liveDepts.filter((d) => d.dept === dept);
  }, [dept, liveDepts]);

  const filteredAudit = useMemo(() => {
    return liveAudit.filter((e) => dept === 'all' || !e.dept || e.dept === dept);
  }, [dept, liveAudit]);

  const openTicketCount = liveTickets.filter((t) =>
    ['pending', 'hr_review', 'finance_review'].includes(t.status),
  ).length;

  const ticket = liveTickets.find((t) => t.id === selectedTicket) ?? liveTickets[0];
  const role = liveRoles.find((r) => r.id === selectedRole) ?? liveRoles[0];
  const activeSnap = snapshots.find((s) => s.id === snapshotId) ?? snapshots[2];

  const correctCurrentDept = liveDepts.find((d) => d.dept === correctDept);
  const correctCurrentValue = correctCurrentDept?.[correctField] ?? 0;

  const pageCopy: Record<AppView, { title: string; sub: string }> = {
    home: {
      title: persona === 'finance' ? 'Headcount reconciliation' : 'Hiring & backfill control',
      sub:
        persona === 'finance'
          ? 'Board vs planned vs actual as of July close — with named seats explaining the gap.'
          : 'Classify seats (backfill, not backfilling, pivot) and move requests through HR and finance.',
    },
    departments: {
      title: 'Departments',
      sub: 'Budget, headcount, and open seats by sector. Click any department to filter the whole product.',
    },
    approvals: {
      title: 'Approvals',
      sub: 'Manager requests flow Manager → HR → Finance. Filter by department or status.',
    },
    roles: {
      title: 'Roles',
      sub: 'Role-level board vs planned vs actual. Tag each seat as new, backfill, not backfilling, pivot, close, or steady.',
    },
    audit: {
      title: 'Audit & snapshots',
      sub: 'Timestamped as-of views and a change history finance can defend.',
    },
  };

  function selectDepartment(next: Dept | 'all') {
    setDept(next);
    if (next !== 'all' && view === 'departments') {
      setView('home');
    }
  }

  function demoAction(msg: string) {
    setToast(msg);
    window.setTimeout(() => setToast(null), 2200);
  }

  function openManual(mode: ManualMode = 'correct') {
    setManualMode(mode);
    if (dept !== 'all') {
      setCorrectDept(dept);
      setDepDept(dept);
      setAddDept(dept);
    }
    setManualOpen(true);
  }

  function stampNow() {
    const now = new Date();
    const hh = String(now.getHours()).padStart(2, '0');
    const mm = String(now.getMinutes()).padStart(2, '0');
    return `${VIEWING_AS_OF} ${hh}:${mm}`;
  }

  function recordManual(kind: string, summary: string, targetDept: Dept, auditAction: string) {
    const ts = stampNow();
    const id = `m-${Date.now()}`;
    setManualAdjs((prev) => [{ id, ts, kind, summary, dept: targetDept }, ...prev]);
    setLiveAudit((prev) => [
      {
        id: `a-${id}`,
        ts,
        actor: persona === 'finance' ? 'Alex Morgan (FP&A)' : 'Jordan Wells (HR)',
        action: auditAction,
        detail: summary,
        dept: targetDept,
      },
      ...prev,
    ]);
  }

  function saveCorrectNumbers() {
    const nextVal = Number(correctValue);
    if (!Number.isFinite(nextVal) || correctValue.trim() === '') {
      demoAction('Enter a valid number');
      return;
    }
    if (!correctReason.trim()) {
      demoAction('Reason is required');
      return;
    }
    const prev = correctCurrentValue;
    setLiveDepts((depts) =>
      depts.map((d) => {
        if (d.dept !== correctDept) return d;
        const updated = { ...d, [correctField]: nextVal };
        if (correctField === 'actual' && nextVal !== d.actual) {
          updated.varianceExplain = `${d.varianceExplain} Manual ${correctField}: ${prev} → ${nextVal} (${correctReason.trim()}).`;
        }
        return updated;
      }),
    );
    const summary = `${correctDept} ${correctField} ${prev} → ${nextVal}: ${correctReason.trim()}`;
    recordManual('Correct numbers', summary, correctDept, 'Manual number correction');
    demoAction(`Updated ${correctDept} ${correctField} to ${nextVal}`);
    setCorrectValue('');
    setCorrectReason('');
    setManualOpen(false);
  }

  function saveDeparture() {
    if (!depTitle.trim() || !depPerson.trim()) {
      demoAction('Role title and person are required');
      return;
    }
    if (depDisposition === 'pivot' && !depPivotTo.trim()) {
      demoAction('Pivoted to is required for pivot');
      return;
    }

    const plannedForType = depDisposition === 'backfill' ? 1 : 0;
    const roleId = `r-manual-${Date.now()}`;
    const newRole: RoleRow = {
      id: roleId,
      title: depTitle.trim(),
      dept: depDept,
      type: depDisposition,
      board: 1,
      planned: plannedForType,
      actual: 0,
      variance: -1,
      status: depDisposition === 'backfill' ? 'pending' : 'open',
      replacing: depPerson.trim(),
      pivotedTo: depDisposition === 'pivot' ? depPivotTo.trim() : undefined,
      note: depNote.trim() || `Departure effective ${depDate}`,
      manager: persona === 'hr' ? 'HR Ops' : 'FP&A',
    };

    setLiveRoles((prev) => [newRole, ...prev]);
    setLiveDepts((depts) =>
      depts.map((d) => {
        if (d.dept !== depDept) return d;
        const next: DeptBudget = {
          ...d,
          actual: Math.max(0, d.actual - 1),
        };
        if (depDisposition === 'backfill') {
          next.openBackfills = d.openBackfills + 1;
          next.openTickets = d.openTickets + 1;
        }
        next.varianceExplain = `${d.varianceExplain} Departure: ${depPerson.trim()} left ${depTitle.trim()} (${typeLabel(depDisposition)}).`;
        return next;
      }),
    );

    if (depDisposition === 'backfill') {
      const ticketId = `BF-M${String(Date.now()).slice(-4)}`;
      setLiveTickets((prev) => [
        {
          id: ticketId,
          title: `${depTitle.trim()} backfill`,
          dept: depDept,
          requestedBy: persona === 'hr' ? 'Jordan Wells (HR)' : 'Alex Morgan (FP&A)',
          createdAt: depDate,
          status: 'pending',
          type: 'backfill' as TicketType,
          replacing: depPerson.trim(),
          targetStart: depDate,
          rationale: depNote.trim() || `Backfill after ${depPerson.trim()} departure`,
          daysOpen: 0,
        },
        ...prev,
      ]);
    }

    const summary =
      depDisposition === 'pivot'
        ? `${depPerson.trim()} left ${depTitle.trim()} · pivoted to ${depPivotTo.trim()}`
        : `${depPerson.trim()} left ${depTitle.trim()} · ${typeLabel(depDisposition)}`;
    recordManual('Log departure', summary, depDept, 'Manual departure');
    demoAction(
      depDisposition === 'backfill'
        ? `Logged departure — ${depDept} actual −1, backfill ticket opened`
        : `Logged departure — ${depDept} actual −1 (${typeLabel(depDisposition)})`,
    );
    setDepTitle('');
    setDepPerson('');
    setDepPivotTo('');
    setDepNote('');
    setManualOpen(false);
  }

  function saveAddSeat() {
    if (!addTitle.trim()) {
      demoAction('Role title is required');
      return;
    }
    const roleId = `r-add-${Date.now()}`;
    const newRole: RoleRow = {
      id: roleId,
      title: addTitle.trim(),
      dept: addDept,
      type: addType,
      board: addBoard,
      planned: addPlanned,
      actual: addActual,
      variance: addActual - addBoard,
      status: addActual >= addBoard ? 'filled' : addType === 'backfill' || addType === 'new' ? 'pending' : 'open',
      note: addNote.trim() || undefined,
      manager: persona === 'hr' ? 'HR Ops' : 'FP&A',
    };
    setLiveRoles((prev) => [newRole, ...prev]);
    setSelectedRole(roleId);

    const bumpTickets = addType === 'new' || addType === 'backfill';
    setLiveDepts((depts) =>
      depts.map((d) => {
        if (d.dept !== addDept) return d;
        return {
          ...d,
          board: d.board + addBoard,
          planned: d.planned + addPlanned,
          actual: d.actual + addActual,
          openTickets: bumpTickets ? d.openTickets + 1 : d.openTickets,
          openBackfills: addType === 'backfill' ? d.openBackfills + 1 : d.openBackfills,
          varianceExplain: `${d.varianceExplain} Added seat: ${addTitle.trim()} (${typeLabel(addType)}).`,
        };
      }),
    );

    if (bumpTickets) {
      const ticketId = `${addType === 'backfill' ? 'BF' : 'RQ'}-M${String(Date.now()).slice(-4)}`;
      setLiveTickets((prev) => [
        {
          id: ticketId,
          title: addTitle.trim(),
          dept: addDept,
          requestedBy: persona === 'hr' ? 'Jordan Wells (HR)' : 'Alex Morgan (FP&A)',
          createdAt: VIEWING_AS_OF,
          status: 'pending',
          type: (addType === 'backfill' ? 'backfill' : 'new') as TicketType,
          targetStart: VIEWING_AS_OF,
          rationale: addNote.trim() || `Manual seat add — ${typeLabel(addType)}`,
          daysOpen: 0,
        },
        ...prev,
      ]);
    }

    const summary = `Added ${addTitle.trim()} · board ${addBoard} / planned ${addPlanned} / actual ${addActual}`;
    recordManual('Add seat', summary, addDept, 'Manual seat add');
    demoAction(`Added ${addTitle.trim()} in ${addDept}`);
    setAddTitle('');
    setAddNote('');
    setAddBoard(1);
    setAddPlanned(1);
    setAddActual(0);
    setManualOpen(false);
  }

  function saveManual() {
    if (manualMode === 'correct') saveCorrectNumbers();
    else if (manualMode === 'departure') saveDeparture();
    else saveAddSeat();
  }

  const yDomain = useMemo(() => {
    const vals = outlookData.flatMap((d) =>
      [d.Board, d.Planned, d.Actual].filter((v): v is number => typeof v === 'number'),
    );
    if (vals.length === 0) return [0, 100] as [number, number];
    const min = Math.max(0, Math.floor((Math.min(...vals) - 10) / 5) * 5);
    const max = Math.ceil((Math.max(...vals) + 10) / 5) * 5;
    return [min, max] as [number, number];
  }, [outlookData]);

  return (
    <div className="shell">
      <aside className="sidebar">
        <div>
          <div className="logo-wrap">
            <img className="logo" src="/bamboohr-logo.png" alt="BambooHR" />
          </div>
          <p className="product-name">Headcount Control Tower</p>
          <p className="product-sub">{COMPANY}</p>
        </div>

        <nav className="nav">
          {NAV.map((item) => (
            <button
              key={item.id}
              className={view === item.id ? 'active' : ''}
              onClick={() => setView(item.id)}
            >
              {item.label}
              {item.id === 'approvals' && <span className="count">{openTicketCount}</span>}
            </button>
          ))}
        </nav>

        <div className="sidebar-foot">
          <div className="field" style={{ marginBottom: 10 }}>
            <label style={{ color: '#8b9482' }}>Viewing as</label>
          </div>
          <div className="persona">
            <button
              className={persona === 'finance' ? 'active' : ''}
              onClick={() => {
                setPersona('finance');
                setChartMode('bridge');
                setView('home');
              }}
            >
              Finance
            </button>
            <button
              className={persona === 'hr' ? 'active' : ''}
              onClick={() => {
                setPersona('hr');
                setView('approvals');
                setTicketFilter('hr_review');
              }}
            >
              HR
            </button>
          </div>
        </div>
      </aside>

      <main className="main">
        <header className="topbar">
          <div>
            <h1>{pageCopy[view].title}</h1>
            <p>{pageCopy[view].sub}</p>
          </div>
          <div className="top-actions">
            <button className="btn primary" onClick={() => openManual('correct')}>
              Manual update
            </button>
            <button className="btn" onClick={() => setExportOpen(true)}>
              Export to Excel
            </button>
            <span className="chip green">Demo data</span>
            <span className="chip">
              Board lock <strong>{BOARD_AS_OF}</strong>
            </span>
            <span className="chip">
              July close · as of <strong>{CLOSE_AS_OF}</strong>
            </span>
            <span className="chip">
              Viewing <strong>{VIEWING_AS_OF}</strong>
            </span>
          </div>
        </header>

        <div className="filters">
          <div className="field">
            <label>Department</label>
            <select
              value={dept}
              onChange={(e) => setDept(e.target.value as Dept | 'all')}
            >
              <option value="all">All departments</option>
              {DEPT_LIST.map((d) => (
                <option key={d} value={d}>
                  {d}
                </option>
              ))}
            </select>
          </div>
          {view === 'roles' && (
            <div className="field">
              <label>Role type</label>
              <select
                value={roleTypeFilter}
                onChange={(e) => setRoleTypeFilter(e.target.value as 'all' | RoleType)}
              >
                <option value="all">All types</option>
                <option value="new">{typeLabel('new')}</option>
                <option value="backfill">{typeLabel('backfill')}</option>
                <option value="not_backfilling">{typeLabel('not_backfilling')}</option>
                <option value="pivot">{typeLabel('pivot')}</option>
                <option value="close">{typeLabel('close')}</option>
                <option value="steady">{typeLabel('steady')}</option>
              </select>
            </div>
          )}
          <div className="spacer" />
          {dept !== 'all' && (
            <button className="btn" onClick={() => setDept('all')}>
              Clear department
            </button>
          )}
        </div>

        {view === 'home' && (
          <>
            <section className="story">
              <div className="story-card">
                <div className="eyebrow">
                  {dept === 'all' ? 'Company story' : `${dept} story`} · {persona === 'finance' ? 'Finance' : 'HR'} lens
                  {' · '}as of {CLOSE_AS_OF} close
                </div>
                <h2>
                  {scope.variance === 0
                    ? 'On board plan'
                    : `${Math.abs(scope.variance)} seat${Math.abs(scope.variance) === 1 ? '' : 's'} ${scope.variance < 0 ? 'under' : 'over'} board`}
                </h2>
                <p>{scope.explain}</p>
                {missingSeats.length > 0 && (
                  <>
                    <h3 style={{ margin: '14px 0 0', fontSize: 13, fontWeight: 600 }}>What&apos;s missing</h3>
                    <div className="missing-list">
                      {missingSeats.slice(0, dept === 'all' ? 5 : 8).map((m) => (
                        <div className="missing-row" key={`${m.dept}-${m.title}-${m.disposition}`}>
                          <div>
                            <div className="title">{m.title}</div>
                            <div className="detail">
                              {dept === 'all' ? `${m.dept} · ` : ''}
                              {m.detail}
                            </div>
                          </div>
                          <span className={`type-pill ${m.disposition}`}>{typeLabel(m.disposition)}</span>
                          <span className={`delta ${m.delta < 0 ? 'neg' : m.delta > 0 ? 'pos' : 'neutral'}`}>
                            {m.delta > 0 ? '+' : ''}
                            {m.delta}
                          </span>
                        </div>
                      ))}
                    </div>
                  </>
                )}
              </div>
              <div className="story-card story-metrics" style={{ background: 'var(--white)', borderColor: 'var(--line)' }}>
                <div className="mini-stat">
                  <span>Board</span>
                  <strong>{scope.board}</strong>
                </div>
                <div className="mini-stat">
                  <span>Actual</span>
                  <strong>{scope.actual}</strong>
                </div>
                <div className="mini-stat">
                  <span>Planned</span>
                  <strong>{scope.planned}</strong>
                </div>
                <div className="mini-stat">
                  <span>Open tickets</span>
                  <strong>{scope.openTickets}</strong>
                </div>
              </div>
            </section>

            <p className="date-banner">
              Headcount figures are as of July 31 close ({CLOSE_AS_OF}), not calendar today ({VIEWING_AS_OF}).
            </p>

            {manualAdjs.length > 0 && (
              <div className="panel" style={{ marginBottom: 16, padding: '14px 16px' }}>
                <div className="panel-h" style={{ marginBottom: 8 }}>
                  <div>
                    <h2 style={{ fontSize: 15 }}>Manual updates this session</h2>
                    <p>Mid-cycle changes applied after July close — reflected in story, bridge, and audit.</p>
                  </div>
                </div>
                <div className="adj-list">
                  {manualAdjs.map((a) => (
                    <div className="adj-item" key={a.id}>
                      <time>{a.ts}</time>
                      <div>
                        <strong>
                          {a.kind} · {a.dept}
                        </strong>
                        <p>{a.summary}</p>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            <section className="kpi-grid">
              <div className="kpi">
                <div className="label">Board</div>
                <div className="value">{scope.board}</div>
                <div className="hint">Approved plan</div>
              </div>
              <div className="kpi">
                <div className="label">Planned</div>
                <div className="value">{scope.planned}</div>
                <div className="hint">Roll-forward</div>
              </div>
              <div className="kpi">
                <div className="label">Actual</div>
                <div className="value">{scope.actual}</div>
                <Variance n={scope.variance} />
              </div>
              <div className="kpi">
                <div className="label">Open tickets</div>
                <div className="value">{scope.openTickets}</div>
                <div className="hint">{scope.openBackfills} backfills</div>
              </div>
              <div className="kpi">
                <div className="label">FY-end plan</div>
                <div className="value">{scope.fyEndPlanned}</div>
                <div className="hint">
                  Board {scope.fyEndBoard} · Δ {scope.fyEndPlanned - scope.fyEndBoard}
                </div>
              </div>
            </section>

            <div className="grid-2">
              <div className="panel">
                <div className="panel-h">
                  <div>
                    <h2>{chartMode === 'outlook' ? 'Monthly outlook' : 'Variance bridge'}</h2>
                    <p>
                      {chartMode === 'outlook'
                        ? hist
                          ? `History ${hist.from}→${hist.to}: actual ${hist.startActual} → ${hist.endActual} (${hist.delta > 0 ? '+' : ''}${hist.delta}). Solid = actual · dashed = planned · bars = board.`
                          : 'Solid line = actual history. Dashed = planned future. Bars = board plan.'
                        : 'Start at board plan, then each change — land on current FTE.'}
                    </p>
                  </div>
                  <div className="btn-row">
                    <button
                      className={`btn ${chartMode === 'outlook' ? 'active' : ''}`}
                      onClick={() => setChartMode('outlook')}
                    >
                      Outlook
                    </button>
                    <button
                      className={`btn ${chartMode === 'bridge' ? 'active' : ''}`}
                      onClick={() => setChartMode('bridge')}
                    >
                      Bridge
                    </button>
                  </div>
                </div>

                {chartMode === 'outlook' && (
                  <div className="range-row">
                    <span className="range-label">Time range</span>
                    <div className="btn-row">
                      {OUTLOOK_RANGES.map((r) => (
                        <button
                          key={r.id}
                          className={`btn ${outlookRange === r.id ? 'active' : ''}`}
                          title={r.hint}
                          onClick={() => setOutlookRange(r.id)}
                        >
                          {r.label}
                        </button>
                      ))}
                    </div>
                  </div>
                )}

                {chartMode === 'outlook' ? (
                  <>
                    <div className={`chart-box ${outlookRange === 'h2' ? 'tall' : 'taller'}`}>
                      <ResponsiveContainer width="100%" height="100%">
                        <ComposedChart data={outlookData} margin={{ top: 12, right: 12, left: 0, bottom: 0 }}>
                          <defs>
                            <linearGradient id="actualFill" x1="0" y1="0" x2="0" y2="1">
                              <stop offset="0%" stopColor="#73c41d" stopOpacity={0.25} />
                              <stop offset="100%" stopColor="#73c41d" stopOpacity={0.02} />
                            </linearGradient>
                          </defs>
                          <CartesianGrid stroke="#eef1ea" vertical={false} />
                          <XAxis
                            dataKey="month"
                            tick={{ fill: '#6f7868', fontSize: outlookRange === 'h2' ? 12 : 10 }}
                            axisLine={false}
                            tickLine={false}
                            interval={outlookRange === '12m' ? 1 : 0}
                          />
                          <YAxis domain={yDomain} tick={{ fill: '#6f7868', fontSize: 12 }} axisLine={false} tickLine={false} width={40} />
                          <Tooltip
                            contentStyle={{ borderRadius: 10, border: '1px solid #e2e7dc', fontSize: 12 }}
                            formatter={(value, name) => [value ?? '—', String(name)]}
                          />
                          <Legend wrapperStyle={{ fontSize: 12 }} />
                          <Bar dataKey="Board" fill="#d7ebb8" radius={[4, 4, 0, 0]} maxBarSize={outlookRange === 'h2' ? 36 : 22} />
                          <Area
                            type="monotone"
                            dataKey="Actual"
                            stroke="#5aa314"
                            fill="url(#actualFill)"
                            strokeWidth={2.5}
                            connectNulls={false}
                            dot={{ r: outlookRange === 'h2' ? 3 : 2, fill: '#5aa314', strokeWidth: 0 }}
                          />
                          <Line
                            type="monotone"
                            dataKey="Planned"
                            stroke="#1c2118"
                            strokeWidth={2}
                            strokeDasharray="5 4"
                            dot={false}
                            connectNulls
                          />
                          <ReferenceLine
                            x="Jul"
                            stroke="#9a7408"
                            strokeDasharray="3 3"
                            label={{ value: 'Jul close', fill: '#9a7408', fontSize: 11, position: 'top' }}
                          />
                        </ComposedChart>
                      </ResponsiveContainer>
                    </div>
                    <div className="legend-row">
                      <span><i style={{ background: '#d7ebb8', height: 10 }} /> Board plan</span>
                      <span><i style={{ background: '#5aa314' }} /> Actual (to close)</span>
                      <span><i style={{ background: '#1c2118', height: 2 }} /> Planned forward</span>
                      <span className="range-hint">
                        {OUTLOOK_RANGES.find((r) => r.id === outlookRange)?.hint}
                      </span>
                    </div>
                  </>
                ) : (
                  <div className="bridge">
                    <div className="bridge-intro">
                      <div>
                        <span>Board plan</span>
                        <strong>{bridgeRows[0]?.value ?? scope.board}</strong>
                      </div>
                      <div className="bridge-arrow">→</div>
                      <div>
                        <span>Current FTE</span>
                        <strong>{bridgeRows[bridgeRows.length - 1]?.value ?? scope.actual}</strong>
                      </div>
                      <div className="bridge-arrow">=</div>
                      <div>
                        <span>Gap</span>
                        <strong className={scope.variance < 0 ? 'neg' : ''}>
                          {scope.variance > 0 ? '+' : ''}
                          {scope.variance}
                        </strong>
                      </div>
                    </div>
                    <div className="bridge-steps">
                      {bridgeRows.map((row, idx) => {
                        if (row.kind === 'start' || row.kind === 'end') {
                          return (
                            <div className={`bridge-row bookend ${row.kind}`} key={`${row.step}-${idx}`}>
                              <div className="bridge-label">
                                <strong>{row.step}</strong>
                                <span>{row.kind === 'start' ? 'Starting point' : 'As of July 31 close'}</span>
                              </div>
                              <div className="bridge-value">{row.value}</div>
                            </div>
                          );
                        }
                        const sign = row.value > 0 ? '+' : '';
                        return (
                          <div className="bridge-row delta-row" key={`${row.step}-${idx}`}>
                            <div className="bridge-label">
                              <strong>{row.step}</strong>
                              <span>Running total → {row.running}</span>
                            </div>
                            <div className="bridge-bar-wrap">
                              <div className="bridge-bar-track">
                                <div
                                  className={`bridge-bar ${row.kind === 'neg' ? 'neg' : 'pos'}`}
                                  style={{ width: `${row.barPct}%` }}
                                />
                              </div>
                            </div>
                            <div className={`bridge-delta ${row.kind === 'neg' ? 'neg' : 'pos'}`}>
                              {sign}
                              {row.value}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                    <p className="bridge-note">
                      Read top to bottom: start at board plan, apply each change, land on filled FTE as of July close.
                      Gaps are named by disposition — backfill, not backfilling, or pivot — so Finance and HR share one story.
                    </p>
                  </div>
                )}
              </div>

              <div className="panel">
                <div className="panel-h">
                  <div>
                    <h2>Departments</h2>
                    <p>Click to focus the whole app on one sector.</p>
                  </div>
                  <button className="btn" onClick={() => setView('departments')}>
                    View all
                  </button>
                </div>
                <div className="dept-list">
                  {filteredDepts.map((d) => {
                    const gap = Math.max(d.board - d.actual, 0);
                    const filledPct = Math.min(100, (d.actual / Math.max(d.board, 1)) * 100);
                    const gapPct = Math.min(100 - filledPct, (gap / Math.max(d.board, 1)) * 100);
                    const v = d.actual - d.board;
                    return (
                      <button
                        key={d.dept}
                        className={`dept-row ${dept === d.dept ? 'active' : ''}`}
                        onClick={() => selectDepartment(d.dept)}
                      >
                        <div>
                          <div className="name">{d.dept}</div>
                          <div className="meta">
                            {d.openTickets} open tickets · {fmtMoney(d.spendUsd)} spend
                          </div>
                        </div>
                        <div className="bar-track">
                          <div className="bar-fill" style={{ width: `${filledPct}%` }} />
                          <div className="bar-fill gap" style={{ width: `${gapPct}%` }} />
                        </div>
                        <div className={`dept-var ${v < 0 ? 'neg' : v > 0 ? 'pos' : ''}`}>
                          {v > 0 ? '+' : ''}
                          {v}
                        </div>
                      </button>
                    );
                  })}
                </div>
              </div>
            </div>

            <div className="grid-2 equal">
              <div className="panel">
                <div className="panel-h">
                  <div>
                    <h2>Needs attention</h2>
                    <p>Open approvals in this scope.</p>
                  </div>
                  <button className="btn primary" onClick={() => setView('approvals')}>
                    Open queue
                  </button>
                </div>
                <div className="ticket-list" style={{ maxHeight: 280 }}>
                  {filteredTickets.filter((t) => !['approved', 'rejected'].includes(t.status)).slice(0, 4).map((t) => (
                    <button
                      key={t.id}
                      className="ticket"
                      onClick={() => {
                        setSelectedTicket(t.id);
                        setView('approvals');
                      }}
                    >
                      <div className="ticket-top">
                        <span className="ticket-id">{t.id}</span>
                        <span className={`status ${t.status}`}>{statusLabel(t.status)}</span>
                      </div>
                      <h3>{t.title}</h3>
                      <div className="who">
                        {t.dept} · {t.daysOpen}d open · <span className={`type-pill ${t.type}`}>{typeLabel(t.type)}</span>
                      </div>
                    </button>
                  ))}
                  {filteredTickets.filter((t) => !['approved', 'rejected'].includes(t.status)).length === 0 && (
                    <div className="empty">No open tickets in this filter.</div>
                  )}
                </div>
              </div>

              <div className="panel">
                <div className="panel-h">
                  <div>
                    <h2>Budget pulse</h2>
                    <p>Spend vs budget for the current scope.</p>
                  </div>
                </div>
                <div className="detail-grid" style={{ gridTemplateColumns: '1fr 1fr', marginTop: 0 }}>
                  <div className="detail-stat">
                    <span>Budget</span>
                    <strong>{fmtMoney(scope.budgetUsd)}</strong>
                  </div>
                  <div className="detail-stat">
                    <span>Spend</span>
                    <strong>{fmtMoney(scope.spendUsd)}</strong>
                  </div>
                  <div className="detail-stat">
                    <span>Open backfills</span>
                    <strong>{scope.openBackfills}</strong>
                  </div>
                  <div className="detail-stat">
                    <span>Open tickets</span>
                    <strong>{scope.openTickets}</strong>
                  </div>
                </div>
                <div className="detail">
                  <h3>Definitions</h3>
                  <p>
                    <strong>Board</strong>, <strong>Planned</strong>, and <strong>Actual</strong> are as of July 31 close ({CLOSE_AS_OF}).
                    Dispositions: <strong>backfill</strong> = refill vacated seat; <strong>not backfilling</strong> = intentional open;
                    <strong> pivot</strong> = headcount moved to another role.
                  </p>
                </div>

                <div className="manual-callout">
                  <div>
                    <strong>Something change mid-month?</strong>
                    <p>Correct numbers, log a departure, or add a seat — updates flow into the story and bridge.</p>
                  </div>
                  <button className="btn primary" onClick={() => openManual('correct')}>
                    Open manual update
                  </button>
                </div>
              </div>
            </div>
          </>
        )}

        {view === 'departments' && (
          <div className="dept-grid">
            {liveDepts.map((d) => {
              const v = d.actual - d.board;
              return (
                <button
                  key={d.dept}
                  className={`dept-card ${dept === d.dept ? 'active' : ''}`}
                  onClick={() => selectDepartment(d.dept)}
                >
                  <h3>{d.dept}</h3>
                  <p style={{ margin: 0, fontSize: 13, color: 'var(--muted)', lineHeight: 1.4 }}>
                    {d.varianceExplain}
                  </p>
                  <div className="nums">
                    <div>
                      <span>Board</span>
                      <strong>{d.board}</strong>
                    </div>
                    <div>
                      <span>Actual</span>
                      <strong>{d.actual}</strong>
                    </div>
                    <div>
                      <span>Δ</span>
                      <strong style={{ color: v < 0 ? 'var(--neg)' : 'var(--pos)' }}>
                        {v > 0 ? '+' : ''}
                        {v}
                      </strong>
                    </div>
                  </div>
                  <div className="who" style={{ marginTop: 12 }}>
                    {d.openTickets} tickets · {fmtMoney(d.budgetUsd)} budget
                  </div>
                </button>
              );
            })}
          </div>
        )}

        {view === 'approvals' && (
          <>
            <div className="btn-row" style={{ marginBottom: 14 }}>
              {(['all', 'pending', 'hr_review', 'finance_review', 'approved', 'rejected'] as const).map((s) => (
                <button
                  key={s}
                  className={`btn ${ticketFilter === s ? 'active' : ''}`}
                  onClick={() => setTicketFilter(s)}
                >
                  {s === 'all' ? 'All' : statusLabel(s)}
                </button>
              ))}
            </div>

            {persona === 'hr' && (
              <div className="lanes" style={{ marginBottom: 14 }}>
                {(
                  [
                    ['pending', 'Submitted'],
                    ['hr_review', 'HR review'],
                    ['finance_review', 'Finance review'],
                    ['approved', 'Approved'],
                  ] as const
                ).map(([status, label]) => (
                  <div className="lane" key={status}>
                    <h3>
                      {label} (
                      {liveTickets.filter((t) => t.status === status && (dept === 'all' || t.dept === dept)).length})
                    </h3>
                    {liveTickets
                      .filter((t) => t.status === status && (dept === 'all' || t.dept === dept))
                      .map((t) => (
                        <button
                          key={t.id}
                          className={`lane-card ${selectedTicket === t.id ? 'active' : ''}`}
                          onClick={() => {
                            setSelectedTicket(t.id);
                            setTicketFilter(status);
                          }}
                        >
                          <h4>{t.title}</h4>
                          <p>
                            {t.id} · {t.dept}
                          </p>
                        </button>
                      ))}
                  </div>
                ))}
              </div>
            )}

            <div className="grid-2 equal">
              <div className="panel">
                <div className="panel-h">
                  <div>
                    <h2>Queue</h2>
                    <p>
                      {filteredTickets.length} request{filteredTickets.length === 1 ? '' : 's'}
                      {dept !== 'all' ? ` in ${dept}` : ''}
                    </p>
                  </div>
                </div>
                <div className="ticket-list">
                  {filteredTickets.length === 0 && <div className="empty">No tickets match these filters.</div>}
                  {filteredTickets.map((t) => (
                    <button
                      key={t.id}
                      className={`ticket ${selectedTicket === t.id ? 'active' : ''}`}
                      onClick={() => setSelectedTicket(t.id)}
                    >
                      <div className="ticket-top">
                        <span className="ticket-id">{t.id}</span>
                        <span className={`status ${t.status}`}>{statusLabel(t.status)}</span>
                      </div>
                      <h3>{t.title}</h3>
                      <div className="who">
                        {t.dept} · {t.requestedBy} · {t.daysOpen}d open
                      </div>
                      <div style={{ marginTop: 8 }}>
                        <span className={`type-pill ${t.type}`}>{typeLabel(t.type)}</span>
                      </div>
                    </button>
                  ))}
                </div>
              </div>

              <div className="panel">
                <div className="panel-h">
                  <div>
                    <h2>{ticket.title}</h2>
                    <p>
                      {ticket.id} · <span className={`type-pill ${ticket.type}`}>{typeLabel(ticket.type)}</span>
                    </p>
                  </div>
                  <span className={`status ${ticket.status}`}>{statusLabel(ticket.status)}</span>
                </div>
                <div className="detail-grid">
                  <div className="detail-stat">
                    <span>Department</span>
                    <strong style={{ fontSize: 14 }}>{ticket.dept}</strong>
                  </div>
                  <div className="detail-stat">
                    <span>Created</span>
                    <strong style={{ fontSize: 14 }}>{ticket.createdAt}</strong>
                  </div>
                  <div className="detail-stat">
                    <span>Target start</span>
                    <strong style={{ fontSize: 14 }}>{ticket.targetStart}</strong>
                  </div>
                  <div className="detail-stat">
                    <span>Days open</span>
                    <strong>{ticket.daysOpen}</strong>
                  </div>
                </div>
                <div className="detail">
                  <h3>Why this hire</h3>
                  <p>{ticket.rationale}</p>
                  {ticket.replacing && (
                    <p style={{ marginTop: 8 }}>
                      <strong>Replacing:</strong> {ticket.replacing}
                    </p>
                  )}
                  {ticket.pivotedTo && (
                    <p style={{ marginTop: 8 }}>
                      <strong>Pivoted to:</strong> {ticket.pivotedTo}
                    </p>
                  )}
                </div>
                <div className="detail">
                  <h3>Path</h3>
                  <p>Manager submits → HR classifies (new / backfill / not backfilling / pivot) → Finance checks board budget → Approved seat enters roll-forward.</p>
                  <div className="manual-callout compact">
                    <p>Need to log something outside the ticket path?</p>
                    <button className="btn" onClick={() => openManual('departure')}>
                      Manual update
                    </button>
                  </div>
                </div>
                <div className="btn-row" style={{ marginTop: 16 }}>
                  <button className="btn primary" onClick={() => demoAction(`Approved ${ticket.id} (demo)`)}>
                    Approve
                  </button>
                  <button className="btn" onClick={() => demoAction(`Changes requested on ${ticket.id}`)}>
                    Request changes
                  </button>
                  <button className="btn" onClick={() => demoAction(`Rejected ${ticket.id} (demo)`)}>
                    Reject
                  </button>
                  <button
                    className="btn"
                    onClick={() => {
                      setDept(ticket.dept);
                      setView('roles');
                    }}
                  >
                    View {ticket.dept} roles
                  </button>
                </div>
                {toast && (
                  <div className="chip green" style={{ marginTop: 12 }}>
                    {toast}
                  </div>
                )}
              </div>
            </div>
          </>
        )}

        {view === 'roles' && (
          <div className="panel full">
            <div className="panel-h">
              <div>
                <h2>Role reconciliation</h2>
                <p>
                  {filteredRoles.length} roles
                  {dept !== 'all' ? ` in ${dept}` : ''}
                  {roleTypeFilter !== 'all' ? ` · ${typeLabel(roleTypeFilter)}` : ''}
                </p>
              </div>
              <div className="btn-row">
                <button className="btn primary" onClick={() => openManual('add')}>
                  Add or edit…
                </button>
                {(['all', 'new', 'backfill', 'not_backfilling', 'pivot', 'close', 'steady'] as const).map((t) => (
                  <button
                    key={t}
                    className={`btn ${roleTypeFilter === t ? 'active' : ''}`}
                    onClick={() => setRoleTypeFilter(t)}
                  >
                    {t === 'all' ? 'All' : typeLabel(t)}
                  </button>
                ))}
              </div>
            </div>
            <div className="role-table-wrap">
              <table className="roles">
                <thead>
                  <tr>
                    <th>Role</th>
                    <th>Dept</th>
                    <th>Type</th>
                    <th>Board</th>
                    <th>Planned</th>
                    <th>Actual</th>
                    <th>Δ</th>
                    <th>Status</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredRoles.map((r) => (
                    <tr
                      key={r.id}
                      className={selectedRole === r.id ? 'active' : ''}
                      onClick={() => setSelectedRole(r.id)}
                    >
                      <td>
                        <strong>{r.title}</strong>
                        {r.replacing && (
                          <div style={{ fontSize: 11, color: 'var(--muted)', marginTop: 2 }}>
                            Replacing {r.replacing}
                          </div>
                        )}
                        {r.pivotedTo && (
                          <div style={{ fontSize: 11, color: 'var(--muted)', marginTop: 2 }}>
                            Pivoted to {r.pivotedTo}
                          </div>
                        )}
                        {r.pivotedFrom && (
                          <div style={{ fontSize: 11, color: 'var(--muted)', marginTop: 2 }}>
                            Pivoted from {r.pivotedFrom}
                          </div>
                        )}
                      </td>
                      <td>{r.dept}</td>
                      <td>
                        <span className={`type-pill ${r.type}`}>{typeLabel(r.type)}</span>
                      </td>
                      <td className="num">{r.board}</td>
                      <td className="num">{r.planned}</td>
                      <td className="num">{r.actual}</td>
                      <td className="num">{r.variance}</td>
                      <td>
                        <span className={`status ${r.status}`}>{statusLabel(r.status)}</span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {filteredRoles.length === 0 && <div className="empty">No roles match these filters.</div>}
            </div>
            <div className="detail">
              <h3>{role.title}</h3>
              <p>
                {role.note ?? 'No additional notes.'}
                {role.manager ? ` Manager: ${role.manager}.` : ''}
                {role.startMonth ? ` Target start: ${role.startMonth}.` : ''}
              </p>
              {role.pivotedTo && (
                <p style={{ marginTop: 8 }}>
                  <strong>Pivoted to:</strong> {role.pivotedTo}
                </p>
              )}
              {role.pivotedFrom && (
                <p style={{ marginTop: 8 }}>
                  <strong>Pivoted from:</strong> {role.pivotedFrom}
                </p>
              )}
              <div className="detail-grid">
                <div className="detail-stat">
                  <span>Board</span>
                  <strong>{role.board}</strong>
                </div>
                <div className="detail-stat">
                  <span>Planned</span>
                  <strong>{role.planned}</strong>
                </div>
                <div className="detail-stat">
                  <span>Actual</span>
                  <strong>{role.actual}</strong>
                </div>
                <div className="detail-stat">
                  <span>Type</span>
                  <strong style={{ fontSize: 14 }}>{typeLabel(role.type)}</strong>
                </div>
              </div>
            </div>
          </div>
        )}

        {view === 'audit' && (
          <div className="grid-2 equal">
            <div className="panel">
              <div className="panel-h">
                <div>
                  <h2>As-of snapshots</h2>
                  <p>Compare board lock, mid-month, and close.</p>
                </div>
              </div>
              <div className="snap-grid" style={{ gridTemplateColumns: '1fr 1fr' }}>
                {snapshots.map((s) => (
                  <button
                    key={s.id}
                    className={`snap ${snapshotId === s.id ? 'active' : ''}`}
                    onClick={() => setSnapshotId(s.id)}
                  >
                    <div className="label">{s.label}</div>
                    <div className="asof">As of {s.asOf}</div>
                    <div className="nums">
                      <div>
                        <span>Board</span>
                        <strong>{s.board}</strong>
                      </div>
                      <div>
                        <span>Actual</span>
                        <strong>{s.actual}</strong>
                      </div>
                      <div>
                        <span>Planned</span>
                        <strong>{s.planned}</strong>
                      </div>
                      <div>
                        <span>Δ Board</span>
                        <strong>{s.variance}</strong>
                      </div>
                    </div>
                  </button>
                ))}
              </div>
              <div className="detail">
                <h3>{activeSnap.label}</h3>
                <p>
                  At {activeSnap.asOf}: board {activeSnap.board}, planned {activeSnap.planned}, actual{' '}
                  {activeSnap.actual}. Variance vs board: {activeSnap.variance}.
                </p>
              </div>
            </div>

            <div className="panel">
              <div className="panel-h">
                <div>
                  <h2>Change history</h2>
                  <p>
                    {filteredAudit.length} events
                    {dept !== 'all' ? ` touching ${dept}` : ''}
                  </p>
                </div>
              </div>
              <div className="audit-list">
                {filteredAudit.map((e) => (
                  <div className="audit-item" key={e.id}>
                    <time>{e.ts}</time>
                    <div>
                      <strong>
                        {e.action} · {e.actor}
                      </strong>
                      <p>
                        {e.detail}
                        {e.dept ? ` · ${e.dept}` : ''}
                      </p>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        <p className="footer-note">
          {COMPANY} demo · BambooHR interview prototype · Filters apply across Home, Approvals, Roles, and Audit
        </p>
      </main>

      {exportOpen && (
        <div className="modal-backdrop" onClick={() => setExportOpen(false)} role="presentation">
          <div
            className="modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="export-title"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="modal-h">
              <div>
                <div className="excel-badge">.xlsx preview</div>
                <h2 id="export-title">Headcount reconciliation report</h2>
                <p>
                  {COMPANY} · As of July close {CLOSE_AS_OF}
                  {dept !== 'all' ? ` · ${dept}` : ' · All departments'} · Ready for finance / HR share-out
                </p>
              </div>
              <button className="btn" onClick={() => setExportOpen(false)}>
                Close
              </button>
            </div>

            <div className="excel-tabs">
              {(
                [
                  ['summary', 'Summary'],
                  ['departments', 'Departments'],
                  ['roles', 'Roles'],
                  ['tickets', 'Tickets'],
                ] as const
              ).map(([id, label]) => (
                <button
                  key={id}
                  className={`excel-tab ${exportSheet === id ? 'active' : ''}`}
                  onClick={() => setExportSheet(id)}
                >
                  {label}
                </button>
              ))}
            </div>

            <div className="excel-sheet">
              <div className="excel-filename">
                Northline_HC_Recon_{AS_OF.replace(/-/g, '')}
                {dept !== 'all' ? `_${dept.replace(/\s/g, '')}` : ''}.xlsx
                <span> · Sheet: {exportSheet}</span>
              </div>

              {exportSheet === 'summary' && (
                <table className="excel-table">
                  <thead>
                    <tr>
                      <th>Metric</th>
                      <th>Value</th>
                      <th>Notes</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td>Board FTE</td>
                      <td>{scope.board}</td>
                      <td>Board-approved headcount</td>
                    </tr>
                    <tr>
                      <td>Planned (roll-forward)</td>
                      <td>{scope.planned}</td>
                      <td>Living plan including approved adds</td>
                    </tr>
                    <tr>
                      <td>Actual FTE</td>
                      <td>{scope.actual}</td>
                      <td>Filled seats as of July close {CLOSE_AS_OF}</td>
                    </tr>
                    <tr>
                      <td>Variance vs board</td>
                      <td>
                        {scope.variance > 0 ? '+' : ''}
                        {scope.variance}
                      </td>
                      <td>{scope.explain}</td>
                    </tr>
                    <tr>
                      <td>Open tickets</td>
                      <td>{scope.openTickets}</td>
                      <td>Pending / in review</td>
                    </tr>
                    <tr>
                      <td>FY-end board</td>
                      <td>{scope.fyEndBoard}</td>
                      <td>December board plan</td>
                    </tr>
                    <tr>
                      <td>FY-end planned</td>
                      <td>{scope.fyEndPlanned}</td>
                      <td>December roll-forward</td>
                    </tr>
                  </tbody>
                </table>
              )}

              {exportSheet === 'departments' && (
                <table className="excel-table">
                  <thead>
                    <tr>
                      <th>Department</th>
                      <th>Board</th>
                      <th>Planned</th>
                      <th>Actual</th>
                      <th>Δ Board</th>
                      <th>Open tickets</th>
                      <th>Budget</th>
                      <th>Spend</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredDepts.map((d) => (
                      <tr key={d.dept}>
                        <td>{d.dept}</td>
                        <td>{d.board}</td>
                        <td>{d.planned}</td>
                        <td>{d.actual}</td>
                        <td>{d.actual - d.board}</td>
                        <td>{d.openTickets}</td>
                        <td>{fmtMoney(d.budgetUsd)}</td>
                        <td>{fmtMoney(d.spendUsd)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}

              {exportSheet === 'roles' && (
                <table className="excel-table">
                  <thead>
                    <tr>
                      <th>Role</th>
                      <th>Department</th>
                      <th>Type</th>
                      <th>Board</th>
                      <th>Planned</th>
                      <th>Actual</th>
                      <th>Δ</th>
                      <th>Status</th>
                      <th>Replacing</th>
                      <th>Pivoted to</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredRoles.slice(0, 25).map((r) => (
                      <tr key={r.id}>
                        <td>{r.title}</td>
                        <td>{r.dept}</td>
                        <td>{typeLabel(r.type)}</td>
                        <td>{r.board}</td>
                        <td>{r.planned}</td>
                        <td>{r.actual}</td>
                        <td>{r.variance}</td>
                        <td>{statusLabel(r.status)}</td>
                        <td>{r.replacing ?? '—'}</td>
                        <td>{r.pivotedTo ?? '—'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}

              {exportSheet === 'tickets' && (
                <table className="excel-table">
                  <thead>
                    <tr>
                      <th>Ticket ID</th>
                      <th>Title</th>
                      <th>Department</th>
                      <th>Type</th>
                      <th>Status</th>
                      <th>Requested by</th>
                      <th>Created</th>
                      <th>Target start</th>
                      <th>Days open</th>
                      <th>Rationale</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredTickets.map((t) => (
                      <tr key={t.id}>
                        <td>{t.id}</td>
                        <td>{t.title}</td>
                        <td>{t.dept}</td>
                        <td>{typeLabel(t.type)}</td>
                        <td>{statusLabel(t.status)}</td>
                        <td>{t.requestedBy}</td>
                        <td>{t.createdAt}</td>
                        <td>{t.targetStart}</td>
                        <td>{t.daysOpen}</td>
                        <td>{t.rationale}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>

            <div className="modal-foot">
              <p>Preview only — shows the workbook layout finance/HR would download. Filters on the page shape this report.</p>
              <div className="btn-row">
                <button
                  className="btn primary"
                  onClick={() =>
                    demoAction('Export queued (demo) — workbook would download as .xlsx')
                  }
                >
                  Download .xlsx
                </button>
                <button className="btn" onClick={() => setExportOpen(false)}>
                  Cancel
                </button>
              </div>
              {toast && <div className="chip green">{toast}</div>}
            </div>
          </div>
        </div>
      )}

      {manualOpen && (
        <div className="modal-backdrop" onClick={() => setManualOpen(false)} role="presentation">
          <div
            className="modal manual-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="manual-title"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="modal-h">
              <div>
                <h2 id="manual-title">Manual update</h2>
                <p>
                  Mid-cycle changes systems miss — edit numbers, log a departure, or add a seat. Updates flow into
                  the story, bridge, and audit.
                </p>
              </div>
              <button className="btn" onClick={() => setManualOpen(false)}>
                Close
              </button>
            </div>

            <div className="mode-tabs" role="tablist" aria-label="Manual update mode">
              {(
                [
                  ['correct', 'Correct numbers'],
                  ['departure', 'Log departure'],
                  ['add', 'Add seat'],
                ] as const
              ).map(([id, label]) => (
                <button
                  key={id}
                  type="button"
                  role="tab"
                  aria-selected={manualMode === id}
                  className={manualMode === id ? 'active' : ''}
                  onClick={() => setManualMode(id)}
                >
                  {label}
                </button>
              ))}
            </div>

            <div className="manual-body">
              {manualMode === 'correct' && (
                <div className="input-grid">
                  <div className="field">
                    <label>Department</label>
                    <select value={correctDept} onChange={(e) => setCorrectDept(e.target.value as Dept)}>
                      {DEPT_LIST.map((d) => (
                        <option key={d} value={d}>
                          {d}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div className="field">
                    <label>Field</label>
                    <select
                      value={correctField}
                      onChange={(e) => setCorrectField(e.target.value as CorrectField)}
                    >
                      <option value="actual">Actual</option>
                      <option value="planned">Planned</option>
                      <option value="board">Board</option>
                    </select>
                  </div>
                  <div className="field">
                    <label>New value</label>
                    <input
                      type="number"
                      value={correctValue}
                      onChange={(e) => setCorrectValue(e.target.value)}
                      placeholder={String(correctCurrentValue)}
                    />
                    <span className="field-hint">Current: {correctCurrentValue}</span>
                  </div>
                  <div className="field" style={{ gridColumn: '1 / -1' }}>
                    <label>Reason</label>
                    <input
                      type="text"
                      value={correctReason}
                      onChange={(e) => setCorrectReason(e.target.value)}
                      placeholder="Why this number changed mid-cycle"
                    />
                  </div>
                </div>
              )}

              {manualMode === 'departure' && (
                <div className="input-grid">
                  <div className="field">
                    <label>Role title</label>
                    <input
                      type="text"
                      value={depTitle}
                      onChange={(e) => setDepTitle(e.target.value)}
                      placeholder="e.g. Events Marketing Manager"
                    />
                  </div>
                  <div className="field">
                    <label>Department</label>
                    <select value={depDept} onChange={(e) => setDepDept(e.target.value as Dept)}>
                      {DEPT_LIST.map((d) => (
                        <option key={d} value={d}>
                          {d}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div className="field">
                    <label>Person leaving</label>
                    <input
                      type="text"
                      value={depPerson}
                      onChange={(e) => setDepPerson(e.target.value)}
                      placeholder="Full name"
                    />
                  </div>
                  <div className="field">
                    <label>Disposition</label>
                    <select
                      value={depDisposition}
                      onChange={(e) => setDepDisposition(e.target.value as DepartureDisposition)}
                    >
                      <option value="backfill">{typeLabel('backfill')}</option>
                      <option value="not_backfilling">{typeLabel('not_backfilling')}</option>
                      <option value="pivot">{typeLabel('pivot')}</option>
                      <option value="close">{typeLabel('close')}</option>
                    </select>
                  </div>
                  {depDisposition === 'pivot' && (
                    <div className="field" style={{ gridColumn: '1 / -1' }}>
                      <label>Pivoted to</label>
                      <input
                        type="text"
                        value={depPivotTo}
                        onChange={(e) => setDepPivotTo(e.target.value)}
                        placeholder="Role · department"
                      />
                    </div>
                  )}
                  <div className="field">
                    <label>Effective date</label>
                    <input type="date" value={depDate} onChange={(e) => setDepDate(e.target.value)} />
                  </div>
                  <div className="field" style={{ gridColumn: '1 / -1' }}>
                    <label>Note</label>
                    <textarea
                      value={depNote}
                      onChange={(e) => setDepNote(e.target.value)}
                      rows={2}
                      placeholder="Optional context"
                    />
                  </div>
                </div>
              )}

              {manualMode === 'add' && (
                <div className="input-grid">
                  <div className="field">
                    <label>Title</label>
                    <input
                      type="text"
                      value={addTitle}
                      onChange={(e) => setAddTitle(e.target.value)}
                      placeholder="Role title"
                    />
                  </div>
                  <div className="field">
                    <label>Department</label>
                    <select value={addDept} onChange={(e) => setAddDept(e.target.value as Dept)}>
                      {DEPT_LIST.map((d) => (
                        <option key={d} value={d}>
                          {d}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div className="field">
                    <label>Type</label>
                    <select value={addType} onChange={(e) => setAddType(e.target.value as AddSeatType)}>
                      <option value="new">{typeLabel('new')}</option>
                      <option value="backfill">{typeLabel('backfill')}</option>
                      <option value="not_backfilling">{typeLabel('not_backfilling')}</option>
                      <option value="pivot">{typeLabel('pivot')}</option>
                      <option value="close">{typeLabel('close')}</option>
                    </select>
                  </div>
                  <div className="field">
                    <label>Board seats</label>
                    <input
                      type="number"
                      min={0}
                      value={addBoard}
                      onChange={(e) => setAddBoard(Number(e.target.value) || 0)}
                    />
                  </div>
                  <div className="field">
                    <label>Planned</label>
                    <input
                      type="number"
                      min={0}
                      value={addPlanned}
                      onChange={(e) => setAddPlanned(Number(e.target.value) || 0)}
                    />
                  </div>
                  <div className="field">
                    <label>Actual</label>
                    <input
                      type="number"
                      min={0}
                      value={addActual}
                      onChange={(e) => setAddActual(Number(e.target.value) || 0)}
                    />
                  </div>
                  <div className="field" style={{ gridColumn: '1 / -1' }}>
                    <label>Note</label>
                    <textarea
                      value={addNote}
                      onChange={(e) => setAddNote(e.target.value)}
                      rows={2}
                      placeholder="Optional context"
                    />
                  </div>
                </div>
              )}
            </div>

            <div className="modal-foot">
              <p>
                {manualMode === 'correct' && 'Corrections update department totals and appear on the variance bridge.'}
                {manualMode === 'departure' &&
                  'Departures reduce actual headcount and create a named role row for the disposition.'}
                {manualMode === 'add' && 'New seats bump department board / planned / actual and appear on Roles.'}
              </p>
              <div className="btn-row">
                <button className="btn primary" onClick={saveManual}>
                  Save update
                </button>
                <button className="btn" onClick={() => setManualOpen(false)}>
                  Cancel
                </button>
              </div>
              {toast && <div className="chip green">{toast}</div>}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
