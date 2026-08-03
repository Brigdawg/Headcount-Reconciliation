export type Dept =
  | 'Engineering'
  | 'Product'
  | 'Design'
  | 'Sales'
  | 'Customer Success'
  | 'Marketing'
  | 'Finance'
  | 'People'
  | 'Operations';

export type RoleType = 'new' | 'backfill' | 'not_backfilling' | 'pivot' | 'close' | 'steady';
export type TicketType = 'backfill' | 'new' | 'not_backfilling' | 'pivot';
export type ApprovalStatus = 'pending' | 'hr_review' | 'finance_review' | 'approved' | 'rejected';
export type Persona = 'finance' | 'hr';
export type AppView = 'home' | 'departments' | 'approvals' | 'roles' | 'audit';

export interface MonthPoint {
  month: string; // chart label
  key: string; // YYYY-MM
  asOf: string;
  board: number;
  planned: number;
  actual: number | null;
  isFuture?: boolean;
}

export type OutlookRange = 'h2' | 'ytd' | '12m';

export const OUTLOOK_RANGES: { id: OutlookRange; label: string; hint: string }[] = [
  { id: 'h2', label: 'H2 outlook', hint: 'Jun–Dec · default close view' },
  { id: 'ytd', label: 'YTD 2026', hint: 'Jan–Dec with history + plan' },
  { id: '12m', label: '12 months', hint: 'Aug ’25–Jul ’26 history + forward' },
];


export interface DeptBudget {
  dept: Dept;
  board: number;
  planned: number;
  actual: number;
  openBackfills: number;
  openTickets: number;
  budgetUsd: number;
  spendUsd: number;
  varianceExplain: string;
}

export interface RoleRow {
  id: string;
  title: string;
  dept: Dept;
  type: RoleType;
  board: number;
  planned: number;
  actual: number;
  variance: number;
  status: ApprovalStatus | 'filled' | 'open';
  replacing?: string;
  /** For pivots: where the headcount moved */
  pivotedTo?: string;
  pivotedFrom?: string;
  note?: string;
  startMonth?: string;
  manager?: string;
}

export interface MissingSeat {
  title: string;
  dept: Dept;
  disposition: RoleType;
  detail: string;
  delta: number;
}

export interface BackfillTicket {
  id: string;
  title: string;
  dept: Dept;
  requestedBy: string;
  createdAt: string;
  status: ApprovalStatus;
  type: TicketType;
  replacing?: string;
  pivotedTo?: string;
  targetStart: string;
  rationale: string;
  daysOpen: number;
}

export interface AuditEvent {
  id: string;
  ts: string;
  actor: string;
  action: string;
  detail: string;
  dept?: Dept;
}

export interface Snapshot {
  id: string;
  label: string;
  asOf: string;
  board: number;
  planned: number;
  actual: number;
  variance: number;
}

export interface BridgeStep {
  step: string;
  value: number;
  kind: 'start' | 'neg' | 'pos' | 'end';
}

export const COMPANY = 'Northline Systems';
/** July month-end close — all headcount numbers in the recon are as of this date */
export const AS_OF = '2026-07-31';
export const CLOSE_AS_OF = AS_OF;
/** Calendar day the user is viewing the tower (live), distinct from close */
export const VIEWING_AS_OF = '2026-08-03';
export const BOARD_AS_OF = '2026-06-30';
export const DEPT_LIST: Dept[] = [
  'Engineering',
  'Product',
  'Design',
  'Sales',
  'Customer Success',
  'Marketing',
  'Finance',
  'People',
  'Operations',
];

/** Build a continuous monthly series with history + forward plan */
function buildSeries(
  points: Array<{
    key: string;
    board: number;
    planned: number;
    actual: number | null;
  }>,
): MonthPoint[] {
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return points.map((p) => {
    const [y, m] = p.key.split('-').map(Number);
    const lastDay = new Date(y, m, 0).getDate();
    const label =
      y < 2026 ? `${MONTHS[m - 1]} ’${String(y).slice(2)}` : MONTHS[m - 1];
    return {
      month: label,
      key: p.key,
      asOf: `${p.key}-${String(lastDay).padStart(2, '0')}`,
      board: p.board,
      planned: p.planned,
      actual: p.actual,
      isFuture: p.actual === null,
    };
  });
}

/** Company-wide: Aug 2025 → Dec 2026 */
export const companyMonths: MonthPoint[] = buildSeries([
  // FY25 H2 history
  { key: '2025-08', board: 248, planned: 248, actual: 246 },
  { key: '2025-09', board: 252, planned: 252, actual: 250 },
  { key: '2025-10', board: 258, planned: 258, actual: 255 },
  { key: '2025-11', board: 264, planned: 264, actual: 261 },
  { key: '2025-12', board: 270, planned: 270, actual: 268 },
  // FY26 YTD history
  { key: '2026-01', board: 275, planned: 275, actual: 272 },
  { key: '2026-02', board: 280, planned: 280, actual: 277 },
  { key: '2026-03', board: 286, planned: 286, actual: 282 },
  { key: '2026-04', board: 290, planned: 290, actual: 286 },
  { key: '2026-05', board: 294, planned: 294, actual: 289 },
  // Current close window
  { key: '2026-06', board: 298, planned: 298, actual: 292 },
  { key: '2026-07', board: 310, planned: 308, actual: 293 },
  // Forward plan
  { key: '2026-08', board: 315, planned: 312, actual: null },
  { key: '2026-09', board: 321, planned: 321, actual: null },
  { key: '2026-10', board: 328, planned: 326, actual: null },
  { key: '2026-11', board: 334, planned: 332, actual: null },
  { key: '2026-12', board: 340, planned: 337, actual: null },
]);

type SeriesSeed = {
  jul: { board: number; planned: number; actual: number };
  dec: { board: number; planned: number };
};

function deptSeries(seed: SeriesSeed): MonthPoint[] {
  const j = seed.jul;
  const d = seed.dec;
  const hist = [
    ['2025-08', 0.8],
    ['2025-09', 0.81],
    ['2025-10', 0.83],
    ['2025-11', 0.85],
    ['2025-12', 0.87],
    ['2026-01', 0.89],
    ['2026-02', 0.9],
    ['2026-03', 0.92],
    ['2026-04', 0.94],
    ['2026-05', 0.95],
    ['2026-06', 0.96],
  ] as const;

  const points: Array<{
    key: string;
    board: number;
    planned: number;
    actual: number | null;
  }> = hist.map(([key, r]) => {
    const board = Math.max(1, Math.round(j.board * r));
    const actual = Math.max(1, Math.round(j.actual * r));
    return { key, board, planned: board, actual };
  });

  points.push({
    key: '2026-07',
    board: j.board,
    planned: j.planned,
    actual: j.actual,
  });

  const fwd = [
    ['2026-08', 0.25],
    ['2026-09', 0.45],
    ['2026-10', 0.65],
    ['2026-11', 0.85],
    ['2026-12', 1],
  ] as const;
  for (const [key, t] of fwd) {
    const board = Math.round(j.board + (d.board - j.board) * t);
    const planned = Math.round(j.planned + (d.planned - j.planned) * t);
    points.push({ key, board, planned, actual: null });
  }

  return buildSeries(points);
}

/** Per-department monthly series so dept filters feel real */
export const deptMonths: Record<Dept, MonthPoint[]> = {
  Engineering: deptSeries({
    jul: { board: 112, planned: 110, actual: 104 },
    dec: { board: 122, planned: 121 },
  }),
  Product: deptSeries({
    jul: { board: 28, planned: 28, actual: 26 },
    dec: { board: 30, planned: 30 },
  }),
  Design: deptSeries({
    jul: { board: 18, planned: 18, actual: 17 },
    dec: { board: 20, planned: 20 },
  }),
  Sales: deptSeries({
    jul: { board: 64, planned: 66, actual: 62 },
    dec: { board: 74, planned: 72 },
  }),
  'Customer Success': deptSeries({
    jul: { board: 42, planned: 41, actual: 40 },
    dec: { board: 45, planned: 45 },
  }),
  Marketing: deptSeries({
    jul: { board: 22, planned: 21, actual: 20 },
    dec: { board: 23, planned: 23 },
  }),
  Finance: deptSeries({
    jul: { board: 12, planned: 12, actual: 12 },
    dec: { board: 13, planned: 13 },
  }),
  People: deptSeries({
    jul: { board: 8, planned: 8, actual: 8 },
    dec: { board: 9, planned: 9 },
  }),
  Operations: deptSeries({
    jul: { board: 4, planned: 4, actual: 4 },
    dec: { board: 4, planned: 4 },
  }),
};


export const companyBridge: BridgeStep[] = [
  { step: 'Board plan', value: 310, kind: 'start' },
  { step: 'Net new delayed', value: -8, kind: 'neg' },
  { step: 'Open backfills', value: -6, kind: 'neg' },
  { step: 'Not backfilling', value: -2, kind: 'neg' },
  { step: 'Early exits (still open)', value: -2, kind: 'neg' },
  { step: 'Offers accepted', value: 1, kind: 'pos' },
  { step: 'Current FTE', value: 293, kind: 'end' },
];

export const deptBridges: Partial<Record<Dept, BridgeStep[]>> = {
  Engineering: [
    { step: 'Board plan', value: 112, kind: 'start' },
    { step: 'Open backfills', value: -4, kind: 'neg' },
    { step: 'Delayed net new', value: -3, kind: 'neg' },
    { step: 'Pivot out → Product', value: -1, kind: 'neg' },
    { step: 'Current FTE', value: 104, kind: 'end' },
  ],
  Sales: [
    { step: 'Board plan', value: 64, kind: 'start' },
    { step: 'Rejected above-board', value: -1, kind: 'neg' },
    { step: 'Open seats', value: -2, kind: 'neg' },
    { step: 'Pivot in from Marketing', value: 1, kind: 'pos' },
    { step: 'Current FTE', value: 62, kind: 'end' },
  ],
  Product: [
    { step: 'Board plan', value: 28, kind: 'start' },
    { step: 'PM backfill open', value: -1, kind: 'neg' },
    { step: 'Analyst delayed', value: -1, kind: 'neg' },
    { step: 'Current FTE', value: 26, kind: 'end' },
  ],
  Marketing: [
    { step: 'Board plan', value: 22, kind: 'start' },
    { step: 'Not backfilling — Events Mgr', value: -1, kind: 'neg' },
    { step: 'Role closed — Lifecycle', value: -1, kind: 'neg' },
    { step: 'Current FTE', value: 20, kind: 'end' },
  ],
};

export const departments: DeptBudget[] = [
  {
    dept: 'Engineering',
    board: 112,
    planned: 110,
    actual: 104,
    openBackfills: 4,
    openTickets: 3,
    budgetUsd: 18_400_000,
    spendUsd: 17_100_000,
    varianceExplain:
      '4 open backfills, 3 delayed net-new, and 1 pivot out to Product (Frontend → Growth Ops). Missing seats are named on Home.',
  },
  {
    dept: 'Product',
    board: 28,
    planned: 28,
    actual: 26,
    openBackfills: 1,
    openTickets: 1,
    budgetUsd: 4_200_000,
    spendUsd: 3_900_000,
    varianceExplain: 'Growth PM backfill pending; one analyst start slipped to September.',
  },
  {
    dept: 'Design',
    board: 18,
    planned: 18,
    actual: 17,
    openBackfills: 1,
    openTickets: 1,
    budgetUsd: 2_400_000,
    spendUsd: 2_250_000,
    varianceExplain: 'Board-approved designer approved; start Aug 25.',
  },
  {
    dept: 'Sales',
    board: 64,
    planned: 66,
    actual: 62,
    openBackfills: 2,
    openTickets: 2,
    budgetUsd: 7_800_000,
    spendUsd: 7_500_000,
    varianceExplain:
      'Roll-forward +1 above board was rejected. Two AE seats still open. One headcount pivoted in from Marketing Content.',
  },
  {
    dept: 'Customer Success',
    board: 42,
    planned: 41,
    actual: 40,
    openBackfills: 1,
    openTickets: 1,
    budgetUsd: 4_600_000,
    spendUsd: 4_400_000,
    varianceExplain: 'TSM net-new queued for August; one attrition still open as a backfill.',
  },
  {
    dept: 'Marketing',
    board: 22,
    planned: 21,
    actual: 20,
    openBackfills: 0,
    openTickets: 1,
    budgetUsd: 3_100_000,
    spendUsd: 2_950_000,
    varianceExplain:
      'Down 2 vs board: (1) Events Marketing Manager — left, not backfilling; (2) Lifecycle Marketing Manager — role closed. Content Manager pivoted to Sales Growth Ops (net team change, not a second open seat).',
  },
  {
    dept: 'Finance',
    board: 12,
    planned: 12,
    actual: 12,
    openBackfills: 0,
    openTickets: 0,
    budgetUsd: 1_800_000,
    spendUsd: 1_750_000,
    varianceExplain: 'On plan.',
  },
  {
    dept: 'People',
    board: 8,
    planned: 8,
    actual: 8,
    openBackfills: 0,
    openTickets: 0,
    budgetUsd: 1_100_000,
    spendUsd: 1_050_000,
    varianceExplain: 'On plan. No open tickets.',
  },
  {
    dept: 'Operations',
    board: 4,
    planned: 4,
    actual: 4,
    openBackfills: 0,
    openTickets: 0,
    budgetUsd: 620_000,
    spendUsd: 600_000,
    varianceExplain: 'On plan.',
  },
];

export const roles: RoleRow[] = [
  // Engineering
  { id: 'r1', title: 'Staff Platform Engineer', dept: 'Engineering', type: 'new', board: 1, planned: 1, actual: 0, variance: -1, status: 'open', note: 'Net new — target Sep', startMonth: 'Sep', manager: 'Maya Ortiz' },
  { id: 'r2', title: 'Sr. Backend Engineer', dept: 'Engineering', type: 'backfill', board: 1, planned: 1, actual: 0, variance: -1, status: 'hr_review', replacing: 'Jordan Lee', note: 'Backfill after resignation', startMonth: 'Sep', manager: 'Maya Ortiz' },
  { id: 'r3', title: 'Frontend Engineer', dept: 'Engineering', type: 'pivot', board: 1, planned: 0, actual: 0, variance: -1, status: 'filled', replacing: 'Casey Ng', pivotedTo: 'Growth Ops Analyst · Sales', note: 'Left Eng; headcount pivoted to Sales Growth Ops', manager: 'Maya Ortiz' },
  { id: 'r4', title: 'Software Engineer II', dept: 'Engineering', type: 'steady', board: 24, planned: 24, actual: 23, variance: -1, status: 'filled', manager: 'Maya Ortiz' },
  { id: 'r5', title: 'Software Engineer I', dept: 'Engineering', type: 'steady', board: 18, planned: 18, actual: 18, variance: 0, status: 'filled', manager: 'Chris Park' },
  { id: 'r6', title: 'Principal Engineer', dept: 'Engineering', type: 'steady', board: 6, planned: 6, actual: 6, variance: 0, status: 'filled', manager: 'Maya Ortiz' },
  { id: 'r7', title: 'QA Manager', dept: 'Engineering', type: 'new', board: 1, planned: 1, actual: 0, variance: -1, status: 'pending', note: 'Net new — Oct', startMonth: 'Oct', manager: 'Chris Park' },
  { id: 'r8', title: 'QA Analyst', dept: 'Engineering', type: 'steady', board: 8, planned: 8, actual: 8, variance: 0, status: 'filled', manager: 'Chris Park' },
  { id: 'r11', title: 'Engineering Manager', dept: 'Engineering', type: 'steady', board: 8, planned: 8, actual: 8, variance: 0, status: 'filled', manager: 'VP Eng' },
  { id: 'r12', title: 'Director of Engineering', dept: 'Engineering', type: 'steady', board: 2, planned: 2, actual: 2, variance: 0, status: 'filled', manager: 'VP Eng' },
  // Product
  { id: 'r13', title: 'Product Manager — Growth', dept: 'Product', type: 'backfill', board: 1, planned: 1, actual: 0, variance: -1, status: 'pending', replacing: 'Avery Chen', startMonth: 'Oct', manager: 'Sam Rivera' },
  { id: 'r14', title: 'Product Manager — Platform', dept: 'Product', type: 'steady', board: 4, planned: 4, actual: 4, variance: 0, status: 'filled', manager: 'Sam Rivera' },
  { id: 'r15', title: 'Product Analyst', dept: 'Product', type: 'new', board: 1, planned: 1, actual: 0, variance: -1, status: 'open', note: 'Start slipped to Sep', startMonth: 'Sep', manager: 'Sam Rivera' },
  { id: 'r16', title: 'Senior Product Manager', dept: 'Product', type: 'steady', board: 6, planned: 6, actual: 6, variance: 0, status: 'filled', manager: 'Sam Rivera' },
  { id: 'r17', title: 'Head of Product', dept: 'Product', type: 'steady', board: 1, planned: 1, actual: 1, variance: 0, status: 'filled' },
  // Design
  { id: 'r19', title: 'Product Designer', dept: 'Design', type: 'new', board: 1, planned: 1, actual: 0, variance: -1, status: 'approved', note: 'Board-approved; Aug 25 start', startMonth: 'Aug', manager: 'Nina Brooks' },
  { id: 'r20', title: 'Senior Product Designer', dept: 'Design', type: 'steady', board: 6, planned: 6, actual: 6, variance: 0, status: 'filled', manager: 'Nina Brooks' },
  { id: 'r21', title: 'Design Manager', dept: 'Design', type: 'steady', board: 2, planned: 2, actual: 2, variance: 0, status: 'filled', manager: 'Nina Brooks' },
  { id: 'r22', title: 'Content Designer', dept: 'Design', type: 'steady', board: 3, planned: 3, actual: 3, variance: 0, status: 'filled', manager: 'Nina Brooks' },
  // Sales
  { id: 'r24', title: 'Outbound AE', dept: 'Sales', type: 'new', board: 2, planned: 3, actual: 1, variance: -1, status: 'rejected', note: 'RF +1 above board rejected', startMonth: 'Sep', manager: 'Taylor Kim' },
  { id: 'r25', title: 'Enterprise BDR', dept: 'Sales', type: 'new', board: 1, planned: 1, actual: 0, variance: -1, status: 'approved', startMonth: 'Aug', manager: 'Taylor Kim' },
  { id: 'r26', title: 'Account Executive', dept: 'Sales', type: 'steady', board: 28, planned: 28, actual: 27, variance: -1, status: 'filled', manager: 'Taylor Kim' },
  { id: 'r27', title: 'SDR', dept: 'Sales', type: 'steady', board: 16, planned: 16, actual: 16, variance: 0, status: 'filled', manager: 'Taylor Kim' },
  { id: 'r28', title: 'Sales Manager', dept: 'Sales', type: 'steady', board: 6, planned: 6, actual: 6, variance: 0, status: 'filled' },
  { id: 'r29', title: 'Growth Ops Analyst', dept: 'Sales', type: 'pivot', board: 1, planned: 1, actual: 1, variance: 0, status: 'filled', pivotedFrom: 'Content Manager · Marketing', replacing: 'Riley Quinn', note: 'Pivot hire — same person moved from Marketing Content to Sales Growth Ops', manager: 'Taylor Kim' },
  // CS
  { id: 'r30', title: 'Technical Success Manager', dept: 'Customer Success', type: 'new', board: 1, planned: 1, actual: 0, variance: -1, status: 'pending', startMonth: 'Aug', manager: 'Priya Shah' },
  { id: 'r31', title: 'Customer Success Manager', dept: 'Customer Success', type: 'steady', board: 18, planned: 18, actual: 18, variance: 0, status: 'filled', manager: 'Priya Shah' },
  { id: 'r32', title: 'Technical Account Manager', dept: 'Customer Success', type: 'backfill', board: 8, planned: 8, actual: 7, variance: -1, status: 'open', replacing: 'Morgan Ellis', manager: 'Priya Shah' },
  { id: 'r33', title: 'CS Director', dept: 'Customer Success', type: 'steady', board: 2, planned: 2, actual: 2, variance: 0, status: 'filled' },
  // Marketing — the −2 story, named
  { id: 'r35', title: 'Lifecycle Marketing Manager', dept: 'Marketing', type: 'close', board: 1, planned: 0, actual: 0, variance: -1, status: 'filled', note: 'Role closed — budget reallocated; not hiring this seat', manager: 'Dana Ortiz' },
  { id: 'r51', title: 'Events Marketing Manager', dept: 'Marketing', type: 'not_backfilling', board: 1, planned: 1, actual: 0, variance: -1, status: 'filled', replacing: 'Sam Brooks', note: 'Sam left Jul 10 — finance + HR agreed not to backfill', manager: 'Dana Ortiz' },
  { id: 'r37', title: 'Content Manager', dept: 'Marketing', type: 'pivot', board: 1, planned: 0, actual: 0, variance: -1, status: 'filled', replacing: 'Riley Quinn', pivotedTo: 'Growth Ops Analyst · Sales', note: 'Riley pivoted to Sales — Marketing seat closed on plan', manager: 'Dana Ortiz' },
  { id: 'r36', title: 'Product Marketing Manager', dept: 'Marketing', type: 'steady', board: 4, planned: 4, actual: 4, variance: 0, status: 'filled' },
  { id: 'r38', title: 'Demand Gen Manager', dept: 'Marketing', type: 'steady', board: 4, planned: 4, actual: 4, variance: 0, status: 'filled' },
  { id: 'r52', title: 'Brand Manager', dept: 'Marketing', type: 'steady', board: 3, planned: 3, actual: 3, variance: 0, status: 'filled' },
  { id: 'r53', title: 'Marketing Ops', dept: 'Marketing', type: 'steady', board: 4, planned: 4, actual: 4, variance: 0, status: 'filled' },
  { id: 'r54', title: 'Marketing Director', dept: 'Marketing', type: 'steady', board: 1, planned: 1, actual: 1, variance: 0, status: 'filled' },
  // Finance / People / Ops
  { id: 'r40', title: 'FP&A Analyst', dept: 'Finance', type: 'steady', board: 3, planned: 3, actual: 3, variance: 0, status: 'filled' },
  { id: 'r41', title: 'Controller', dept: 'Finance', type: 'steady', board: 1, planned: 1, actual: 1, variance: 0, status: 'filled' },
  { id: 'r42', title: 'Staff Accountant', dept: 'Finance', type: 'steady', board: 4, planned: 4, actual: 4, variance: 0, status: 'filled' },
  { id: 'r44', title: 'HR Generalist', dept: 'People', type: 'steady', board: 3, planned: 3, actual: 3, variance: 0, status: 'filled' },
  { id: 'r45', title: 'HRBP', dept: 'People', type: 'steady', board: 2, planned: 2, actual: 2, variance: 0, status: 'filled' },
  { id: 'r46', title: 'People Ops Manager', dept: 'People', type: 'steady', board: 1, planned: 1, actual: 1, variance: 0, status: 'filled' },
  { id: 'r47', title: 'Recruiter', dept: 'People', type: 'steady', board: 2, planned: 2, actual: 2, variance: 0, status: 'filled' },
  { id: 'r48', title: 'Office Manager', dept: 'Operations', type: 'steady', board: 2, planned: 2, actual: 2, variance: 0, status: 'filled' },
  { id: 'r49', title: 'IT Support', dept: 'Operations', type: 'steady', board: 2, planned: 2, actual: 2, variance: 0, status: 'filled' },
];

export const tickets: BackfillTicket[] = [
  {
    id: 'BF-1042',
    title: 'Sr. Backend Engineer',
    dept: 'Engineering',
    requestedBy: 'Maya Ortiz (Eng Manager)',
    createdAt: '2026-07-22',
    status: 'hr_review',
    type: 'backfill',
    replacing: 'Jordan Lee',
    targetStart: '2026-09-15',
    rationale: 'Critical path for platform reliability; departure Aug 8.',
    daysOpen: 9,
  },
  {
    id: 'BF-1048',
    title: 'Product Manager — Growth',
    dept: 'Product',
    requestedBy: 'Sam Rivera (Head of Product)',
    createdAt: '2026-07-25',
    status: 'pending',
    type: 'backfill',
    replacing: 'Avery Chen',
    targetStart: '2026-10-01',
    rationale: 'Preserve Q4 growth roadmap ownership.',
    daysOpen: 6,
  },
  {
    id: 'PV-3010',
    title: 'Growth Ops Analyst',
    dept: 'Sales',
    requestedBy: 'Taylor Kim (Sales VP)',
    createdAt: '2026-07-18',
    status: 'approved',
    type: 'pivot',
    replacing: 'Riley Quinn (Content Manager, Marketing)',
    pivotedTo: 'Sales · Growth Ops Analyst',
    targetStart: '2026-08-01',
    rationale: 'Same person, new seat: Marketing Content → Sales Growth Ops. Keeps company FTE flat while moving the head.',
    daysOpen: 13,
  },
  {
    id: 'NB-4401',
    title: 'Events Marketing Manager',
    dept: 'Marketing',
    requestedBy: 'Dana Ortiz (Marketing Dir)',
    createdAt: '2026-07-12',
    status: 'approved',
    type: 'not_backfilling',
    replacing: 'Sam Brooks',
    targetStart: '—',
    rationale: 'Sam left Jul 10. Finance and HR agreed: do not backfill. Seat stays on board until next reforecast so the −1 is explained.',
    daysOpen: 19,
  },
  {
    id: 'RQ-2214',
    title: 'Product Designer',
    dept: 'Design',
    requestedBy: 'Nina Brooks (Design Dir)',
    createdAt: '2026-07-12',
    status: 'approved',
    type: 'new',
    targetStart: '2026-08-25',
    rationale: 'Board-approved net new for Design system capacity.',
    daysOpen: 19,
  },
  {
    id: 'RQ-2220',
    title: 'Outbound AE',
    dept: 'Sales',
    requestedBy: 'Taylor Kim (Sales VP)',
    createdAt: '2026-07-28',
    status: 'rejected',
    type: 'new',
    targetStart: '2026-09-01',
    rationale: 'RF requested +1 above board; finance rejected until Q4 reforecast.',
    daysOpen: 3,
  },
  {
    id: 'BF-1051',
    title: 'Technical Success Manager',
    dept: 'Customer Success',
    requestedBy: 'Priya Shah (CS Dir)',
    createdAt: '2026-07-29',
    status: 'pending',
    type: 'new',
    targetStart: '2026-08-31',
    rationale: 'Support expanded enterprise book; aligned to Aug RF add.',
    daysOpen: 2,
  },
  {
    id: 'BF-1055',
    title: 'Technical Account Manager',
    dept: 'Customer Success',
    requestedBy: 'Priya Shah (CS Dir)',
    createdAt: '2026-07-20',
    status: 'hr_review',
    type: 'backfill',
    replacing: 'Morgan Ellis',
    targetStart: '2026-09-01',
    rationale: 'Maintain enterprise coverage after resignation.',
    daysOpen: 11,
  },
  {
    id: 'RQ-2228',
    title: 'QA Manager',
    dept: 'Engineering',
    requestedBy: 'Chris Park (Platform Lead)',
    createdAt: '2026-07-26',
    status: 'pending',
    type: 'new',
    targetStart: '2026-10-15',
    rationale: 'Board-approved net new for quality ownership.',
    daysOpen: 5,
  },
];

export const auditTrail: AuditEvent[] = [
  { id: 'a1', ts: '2026-07-31 16:42', actor: 'Alex Morgan (FP&A)', action: 'Snapshot locked', detail: 'July close — Board 310 / Planned 308 / Actual 293 as of 2026-07-31', dept: 'Finance' },
  { id: 'a2', ts: '2026-07-31 14:10', actor: 'Jordan Wells (HR)', action: 'Ticket advanced', detail: 'BF-1042 moved to HR review', dept: 'Engineering' },
  { id: 'a3', ts: '2026-07-30 11:05', actor: 'Alex Morgan (FP&A)', action: 'Variance explained', detail: 'Marketing −2 named: Events not backfilling + Lifecycle closed', dept: 'Marketing' },
  { id: 'a4', ts: '2026-07-28 09:40', actor: 'Finance Committee', action: 'Rejected request', detail: 'RQ-2220 Outbound AE exceeds board plan', dept: 'Sales' },
  { id: 'a5', ts: '2026-07-26 13:00', actor: 'Chris Park', action: 'Opened net-new', detail: 'RQ-2228 QA Manager submitted', dept: 'Engineering' },
  { id: 'a6', ts: '2026-07-25 15:22', actor: 'Sam Rivera', action: 'Opened backfill', detail: 'BF-1048 Product Manager — Growth', dept: 'Product' },
  { id: 'a7', ts: '2026-07-20 10:15', actor: 'Priya Shah', action: 'Opened backfill', detail: 'BF-1055 Technical Account Manager', dept: 'Customer Success' },
  { id: 'a8', ts: '2026-07-18 10:00', actor: 'Taylor Kim', action: 'Pivot approved', detail: 'PV-3010 Marketing Content → Sales Growth Ops', dept: 'Sales' },
  { id: 'a9', ts: '2026-07-12 09:30', actor: 'Dana Ortiz', action: 'Not backfilling', detail: 'NB-4401 Events Marketing Manager — Sam Brooks', dept: 'Marketing' },
  { id: 'a10', ts: '2026-06-30 17:00', actor: 'Board', action: 'Plan locked', detail: 'H2 board headcount locked: 310 (Jul) → 340 (Dec)' },
];

export const snapshots: Snapshot[] = [
  { id: 's0', label: 'FY25 close', asOf: '2025-12-31', board: 270, planned: 270, actual: 268, variance: -2 },
  { id: 's1', label: 'Q1 close', asOf: '2026-03-31', board: 286, planned: 286, actual: 282, variance: -4 },
  { id: 's2', label: 'Board lock', asOf: '2026-06-30', board: 298, planned: 298, actual: 292, variance: -6 },
  { id: 's3', label: 'July mid-month', asOf: '2026-07-15', board: 310, planned: 309, actual: 290, variance: -20 },
  { id: 's4', label: 'July close', asOf: '2026-07-31', board: 310, planned: 308, actual: 293, variance: -17 },
];

export const quotes = [
  {
    text: 'Reconciliation is critical, but the presentation tools are complicated and nobody understands them',
    attribution: 'Head of Finance',
  },
  {
    text: 'If this product existed, you would sell the finance department much easier',
    attribution: 'Head of Finance',
  },
];

/** Named seats that explain plan vs actual gaps — what Finance asks for first */
export function getMissingSeats(dept: Dept | 'all'): MissingSeat[] {
  const gapTypes: RoleType[] = ['new', 'backfill', 'not_backfilling', 'pivot', 'close'];
  return roles
    .filter((r) => {
      if (dept !== 'all' && r.dept !== dept) return false;
      if (!gapTypes.includes(r.type)) return false;
      if (r.type === 'steady') return false;
      // Show seats that contribute to a story gap (underfilled or dispositioned)
      if (r.type === 'pivot' && r.actual === 1 && r.pivotedFrom) return true; // destination side
      if (r.actual < r.board || r.type === 'close' || r.type === 'not_backfilling' || (r.type === 'pivot' && r.actual === 0)) {
        return true;
      }
      return false;
    })
    .map((r) => ({
      title: r.title,
      dept: r.dept,
      disposition: r.type,
      delta: r.actual - r.board,
      detail:
        r.note ??
        (r.pivotedTo
          ? `Pivoted to ${r.pivotedTo}`
          : r.pivotedFrom
            ? `Pivoted from ${r.pivotedFrom}`
            : r.replacing
              ? `Replacing ${r.replacing}`
              : 'Open vs board'),
    }))
    .sort((a, b) => a.delta - b.delta);
}

export function typeLabel(t: string) {
  return t.replace(/_/g, ' ');
}

export function getMonthsForDept(dept: Dept | 'all', range: OutlookRange = 'h2'): MonthPoint[] {
  const all = dept === 'all' ? companyMonths : deptMonths[dept];
  if (range === 'h2') {
    return all.filter((m) => m.key >= '2026-06' && m.key <= '2026-12');
  }
  if (range === 'ytd') {
    return all.filter((m) => m.key >= '2026-01' && m.key <= '2026-12');
  }
  // 12 months of history through today, plus remaining FY forward
  return all.filter((m) => m.key >= '2025-08' && m.key <= '2026-12');
}

export function historySummary(months: MonthPoint[]) {
  const past = months.filter((m) => m.actual !== null);
  if (past.length < 2) return null;
  const first = past[0];
  const last = past[past.length - 1];
  return {
    from: first.month,
    to: last.month,
    startActual: first.actual ?? 0,
    endActual: last.actual ?? 0,
    delta: (last.actual ?? 0) - (first.actual ?? 0),
  };
}

export function getBridgeForDept(dept: Dept | 'all'): BridgeStep[] {
  if (dept === 'all') return companyBridge;
  return deptBridges[dept] ?? [
    {
      step: 'Board plan',
      value: departments.find((d) => d.dept === dept)?.board ?? 0,
      kind: 'start',
    },
    {
      step: 'Current FTE',
      value: departments.find((d) => d.dept === dept)?.actual ?? 0,
      kind: 'end',
    },
  ];
}

export function summarizeScope(dept: Dept | 'all') {
  const depts = dept === 'all' ? departments : departments.filter((d) => d.dept === dept);
  const board = depts.reduce((s, d) => s + d.board, 0);
  const planned = depts.reduce((s, d) => s + d.planned, 0);
  const actual = depts.reduce((s, d) => s + d.actual, 0);
  const months = getMonthsForDept(dept, 'ytd');
  const fy = months[months.length - 1];
  const missing = getMissingSeats(dept).filter((m) =>
    ['not_backfilling', 'close', 'backfill', 'new'].includes(m.disposition),
  );

  const storyBits =
    dept === 'all'
      ? 'You are 17 seats under board as of July 31 close. The bridge names why: delayed net-new, open backfills, seats we are not backfilling, and early exits still open.'
      : depts[0]?.varianceExplain ?? '';

  return {
    board,
    planned,
    actual,
    variance: actual - board,
    openTickets: depts.reduce((s, d) => s + d.openTickets, 0),
    openBackfills: depts.reduce((s, d) => s + d.openBackfills, 0),
    budgetUsd: depts.reduce((s, d) => s + d.budgetUsd, 0),
    spendUsd: depts.reduce((s, d) => s + d.spendUsd, 0),
    fyEndBoard: fy.board,
    fyEndPlanned: fy.planned,
    missingCount: missing.length,
    explain: storyBits,
  };
}
