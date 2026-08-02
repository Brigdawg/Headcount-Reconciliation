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

export type RoleType = 'new' | 'backfill' | 'replace' | 'close' | 'steady';
export type ApprovalStatus = 'pending' | 'hr_review' | 'finance_review' | 'approved' | 'rejected';
export type EmpType = 'fte' | 'contractor';
export type Persona = 'finance' | 'hr';
export type AppView = 'home' | 'departments' | 'approvals' | 'roles' | 'audit';

export interface MonthPoint {
  month: string; // chart label
  key: string; // YYYY-MM
  asOf: string;
  board: number;
  planned: number;
  actual: number | null;
  contractors: number;
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
  contractors: number;
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
  empType: EmpType;
  board: number;
  planned: number;
  actual: number;
  variance: number;
  status: ApprovalStatus | 'filled' | 'open';
  replacing?: string;
  note?: string;
  startMonth?: string;
  manager?: string;
}

export interface BackfillTicket {
  id: string;
  title: string;
  dept: Dept;
  requestedBy: string;
  createdAt: string;
  status: ApprovalStatus;
  type: 'backfill' | 'new' | 'replace';
  replacing?: string;
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
  contractors: number;
  variance: number;
}

export interface BridgeStep {
  step: string;
  value: number;
  kind: 'start' | 'neg' | 'pos' | 'end';
}

export const COMPANY = 'Northline Systems';
export const AS_OF = '2026-07-31';
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
    contractors: number;
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
      contractors: p.contractors,
      isFuture: p.actual === null,
    };
  });
}

/** Company-wide: Aug 2025 → Dec 2026 */
export const companyMonths: MonthPoint[] = buildSeries([
  // FY25 H2 history
  { key: '2025-08', board: 248, planned: 248, actual: 246, contractors: 18 },
  { key: '2025-09', board: 252, planned: 252, actual: 250, contractors: 19 },
  { key: '2025-10', board: 258, planned: 258, actual: 255, contractors: 20 },
  { key: '2025-11', board: 264, planned: 264, actual: 261, contractors: 20 },
  { key: '2025-12', board: 270, planned: 270, actual: 268, contractors: 21 },
  // FY26 YTD history
  { key: '2026-01', board: 275, planned: 275, actual: 272, contractors: 21 },
  { key: '2026-02', board: 280, planned: 280, actual: 277, contractors: 22 },
  { key: '2026-03', board: 286, planned: 286, actual: 282, contractors: 22 },
  { key: '2026-04', board: 290, planned: 290, actual: 286, contractors: 22 },
  { key: '2026-05', board: 294, planned: 294, actual: 289, contractors: 23 },
  // Current close window
  { key: '2026-06', board: 298, planned: 298, actual: 292, contractors: 23 },
  { key: '2026-07', board: 310, planned: 308, actual: 293, contractors: 24 },
  // Forward plan
  { key: '2026-08', board: 315, planned: 312, actual: null, contractors: 24 },
  { key: '2026-09', board: 321, planned: 321, actual: null, contractors: 23 },
  { key: '2026-10', board: 328, planned: 326, actual: null, contractors: 21 },
  { key: '2026-11', board: 334, planned: 332, actual: null, contractors: 18 },
  { key: '2026-12', board: 340, planned: 337, actual: null, contractors: 17 },
]);

type SeriesSeed = {
  // Jul 2026 "current" board/planned/actual/contractors — scale history from these
  jul: { board: number; planned: number; actual: number; contractors: number };
  // Dec 2026 forward
  dec: { board: number; planned: number; contractors: number };
};

function deptSeries(seed: SeriesSeed): MonthPoint[] {
  // Approximate share of company growth; history scales smoothly into Jul actuals
  const j = seed.jul;
  const d = seed.dec;
  // History ratios vs Jul board (company Jul board = 310)
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
    contractors: number;
  }> = hist.map(([key, r]) => {
    const board = Math.max(1, Math.round(j.board * r));
    const actual = Math.max(1, Math.round(j.actual * r));
    const contractors = Math.max(0, Math.round(j.contractors * Math.min(1, r + 0.05)));
    return { key, board, planned: board, actual, contractors };
  });

  points.push({
    key: '2026-07',
    board: j.board,
    planned: j.planned,
    actual: j.actual,
    contractors: j.contractors,
  });

  // Smooth forward to Dec
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
    const contractors = Math.round(j.contractors + (d.contractors - j.contractors) * t);
    points.push({ key, board, planned, actual: null, contractors });
  }

  return buildSeries(points);
}

/** Per-department monthly series so dept filters feel real */
export const deptMonths: Record<Dept, MonthPoint[]> = {
  Engineering: deptSeries({
    jul: { board: 112, planned: 110, actual: 104, contractors: 12 },
    dec: { board: 122, planned: 121, contractors: 8 },
  }),
  Product: deptSeries({
    jul: { board: 28, planned: 28, actual: 26, contractors: 2 },
    dec: { board: 30, planned: 30, contractors: 1 },
  }),
  Design: deptSeries({
    jul: { board: 18, planned: 18, actual: 17, contractors: 1 },
    dec: { board: 20, planned: 20, contractors: 1 },
  }),
  Sales: deptSeries({
    jul: { board: 64, planned: 66, actual: 62, contractors: 3 },
    dec: { board: 74, planned: 72, contractors: 2 },
  }),
  'Customer Success': deptSeries({
    jul: { board: 42, planned: 41, actual: 40, contractors: 2 },
    dec: { board: 45, planned: 45, contractors: 2 },
  }),
  Marketing: deptSeries({
    jul: { board: 22, planned: 21, actual: 20, contractors: 2 },
    dec: { board: 23, planned: 23, contractors: 1 },
  }),
  Finance: deptSeries({
    jul: { board: 12, planned: 12, actual: 12, contractors: 1 },
    dec: { board: 13, planned: 13, contractors: 1 },
  }),
  People: deptSeries({
    jul: { board: 8, planned: 8, actual: 8, contractors: 0 },
    dec: { board: 9, planned: 9, contractors: 0 },
  }),
  Operations: deptSeries({
    jul: { board: 4, planned: 4, actual: 4, contractors: 1 },
    dec: { board: 4, planned: 4, contractors: 1 },
  }),
};


export const companyBridge: BridgeStep[] = [
  { step: 'Board plan', value: 310, kind: 'start' },
  { step: 'Net new delayed', value: -8, kind: 'neg' },
  { step: 'Backfills open', value: -7, kind: 'neg' },
  { step: 'Early exits', value: -3, kind: 'neg' },
  { step: 'Offers accepted', value: 1, kind: 'pos' },
  { step: 'Current FTE', value: 293, kind: 'end' },
];

export const deptBridges: Partial<Record<Dept, BridgeStep[]>> = {
  Engineering: [
    { step: 'Board plan', value: 112, kind: 'start' },
    { step: 'Open backfills', value: -4, kind: 'neg' },
    { step: 'Delayed net new', value: -3, kind: 'neg' },
    { step: 'Early exit', value: -1, kind: 'neg' },
    { step: 'Current FTE', value: 104, kind: 'end' },
  ],
  Sales: [
    { step: 'Board plan', value: 64, kind: 'start' },
    { step: 'Rejected above-board', value: -1, kind: 'neg' },
    { step: 'Open seats', value: -2, kind: 'neg' },
    { step: 'Offers accepted', value: 1, kind: 'pos' },
    { step: 'Current FTE', value: 62, kind: 'end' },
  ],
  Product: [
    { step: 'Board plan', value: 28, kind: 'start' },
    { step: 'PM backfill open', value: -1, kind: 'neg' },
    { step: 'Analyst delayed', value: -1, kind: 'neg' },
    { step: 'Current FTE', value: 26, kind: 'end' },
  ],
};

export const departments: DeptBudget[] = [
  {
    dept: 'Engineering',
    board: 112,
    planned: 110,
    actual: 104,
    contractors: 12,
    openBackfills: 4,
    openTickets: 3,
    budgetUsd: 18_400_000,
    spendUsd: 17_100_000,
    varianceExplain: '4 open backfills and 3 delayed net-new hires; 12 contractors sit outside board FTE.',
  },
  {
    dept: 'Product',
    board: 28,
    planned: 28,
    actual: 26,
    contractors: 2,
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
    contractors: 1,
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
    contractors: 3,
    openBackfills: 2,
    openTickets: 2,
    budgetUsd: 7_800_000,
    spendUsd: 7_500_000,
    varianceExplain: 'Roll-forward asked +1 above board (rejected). Two AE seats still open.',
  },
  {
    dept: 'Customer Success',
    board: 42,
    planned: 41,
    actual: 40,
    contractors: 2,
    openBackfills: 1,
    openTickets: 1,
    budgetUsd: 4_600_000,
    spendUsd: 4_400_000,
    varianceExplain: 'TSM net-new queued for August; one attrition not yet backfilled.',
  },
  {
    dept: 'Marketing',
    board: 22,
    planned: 21,
    actual: 20,
    contractors: 2,
    openBackfills: 0,
    openTickets: 0,
    budgetUsd: 3_100_000,
    spendUsd: 2_950_000,
    varianceExplain: 'Lifecycle role closed; budget reallocated. Plan sits 1 under board.',
  },
  {
    dept: 'Finance',
    board: 12,
    planned: 12,
    actual: 12,
    contractors: 1,
    openBackfills: 0,
    openTickets: 0,
    budgetUsd: 1_800_000,
    spendUsd: 1_750_000,
    varianceExplain: 'On plan. Contractor supports systems migration.',
  },
  {
    dept: 'People',
    board: 8,
    planned: 8,
    actual: 8,
    contractors: 0,
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
    contractors: 1,
    openBackfills: 0,
    openTickets: 0,
    budgetUsd: 620_000,
    spendUsd: 600_000,
    varianceExplain: 'On plan.',
  },
];

export const roles: RoleRow[] = [
  // Engineering
  { id: 'r1', title: 'Staff Platform Engineer', dept: 'Engineering', type: 'new', empType: 'fte', board: 1, planned: 1, actual: 0, variance: -1, status: 'open', note: 'Net new — target Sep', startMonth: 'Sep', manager: 'Maya Ortiz' },
  { id: 'r2', title: 'Sr. Backend Engineer', dept: 'Engineering', type: 'backfill', empType: 'fte', board: 1, planned: 1, actual: 0, variance: -1, status: 'hr_review', replacing: 'Jordan Lee', note: 'Backfill after resignation', startMonth: 'Sep', manager: 'Maya Ortiz' },
  { id: 'r3', title: 'DevOps Engineer', dept: 'Engineering', type: 'replace', empType: 'fte', board: 1, planned: 1, actual: 0, variance: -1, status: 'finance_review', replacing: 'Contract DevOps', note: 'Convert contractor → FTE', startMonth: 'Sep', manager: 'Chris Park' },
  { id: 'r4', title: 'Software Engineer II', dept: 'Engineering', type: 'steady', empType: 'fte', board: 24, planned: 24, actual: 23, variance: -1, status: 'filled', manager: 'Maya Ortiz' },
  { id: 'r5', title: 'Software Engineer I', dept: 'Engineering', type: 'steady', empType: 'fte', board: 18, planned: 18, actual: 18, variance: 0, status: 'filled', manager: 'Chris Park' },
  { id: 'r6', title: 'Principal Engineer', dept: 'Engineering', type: 'steady', empType: 'fte', board: 6, planned: 6, actual: 6, variance: 0, status: 'filled', manager: 'Maya Ortiz' },
  { id: 'r7', title: 'QA Manager', dept: 'Engineering', type: 'new', empType: 'fte', board: 1, planned: 1, actual: 0, variance: -1, status: 'pending', note: 'Net new — Oct', startMonth: 'Oct', manager: 'Chris Park' },
  { id: 'r8', title: 'QA Analyst', dept: 'Engineering', type: 'steady', empType: 'fte', board: 8, planned: 8, actual: 8, variance: 0, status: 'filled', manager: 'Chris Park' },
  { id: 'r9', title: 'Data Engineer (contract)', dept: 'Engineering', type: 'steady', empType: 'contractor', board: 0, planned: 3, actual: 3, variance: 0, status: 'filled', note: 'Outside board FTE', manager: 'Chris Park' },
  { id: 'r10', title: 'Platform Engineer (contract)', dept: 'Engineering', type: 'steady', empType: 'contractor', board: 0, planned: 9, actual: 9, variance: 0, status: 'filled', note: 'Outside board FTE', manager: 'Maya Ortiz' },
  { id: 'r11', title: 'Engineering Manager', dept: 'Engineering', type: 'steady', empType: 'fte', board: 8, planned: 8, actual: 8, variance: 0, status: 'filled', manager: 'VP Eng' },
  { id: 'r12', title: 'Director of Engineering', dept: 'Engineering', type: 'steady', empType: 'fte', board: 2, planned: 2, actual: 2, variance: 0, status: 'filled', manager: 'VP Eng' },
  // Product
  { id: 'r13', title: 'Product Manager — Growth', dept: 'Product', type: 'backfill', empType: 'fte', board: 1, planned: 1, actual: 0, variance: -1, status: 'pending', replacing: 'Avery Chen', startMonth: 'Oct', manager: 'Sam Rivera' },
  { id: 'r14', title: 'Product Manager — Platform', dept: 'Product', type: 'steady', empType: 'fte', board: 4, planned: 4, actual: 4, variance: 0, status: 'filled', manager: 'Sam Rivera' },
  { id: 'r15', title: 'Product Analyst', dept: 'Product', type: 'new', empType: 'fte', board: 1, planned: 1, actual: 0, variance: -1, status: 'open', note: 'Start slipped to Sep', startMonth: 'Sep', manager: 'Sam Rivera' },
  { id: 'r16', title: 'Senior Product Manager', dept: 'Product', type: 'steady', empType: 'fte', board: 6, planned: 6, actual: 6, variance: 0, status: 'filled', manager: 'Sam Rivera' },
  { id: 'r17', title: 'Head of Product', dept: 'Product', type: 'steady', empType: 'fte', board: 1, planned: 1, actual: 1, variance: 0, status: 'filled' },
  { id: 'r18', title: 'Research contractor', dept: 'Product', type: 'steady', empType: 'contractor', board: 0, planned: 2, actual: 2, variance: 0, status: 'filled', note: 'Outside board FTE' },
  // Design
  { id: 'r19', title: 'Product Designer', dept: 'Design', type: 'new', empType: 'fte', board: 1, planned: 1, actual: 0, variance: -1, status: 'approved', note: 'Board-approved; Aug 25 start', startMonth: 'Aug', manager: 'Nina Brooks' },
  { id: 'r20', title: 'Senior Product Designer', dept: 'Design', type: 'steady', empType: 'fte', board: 6, planned: 6, actual: 6, variance: 0, status: 'filled', manager: 'Nina Brooks' },
  { id: 'r21', title: 'Design Manager', dept: 'Design', type: 'steady', empType: 'fte', board: 2, planned: 2, actual: 2, variance: 0, status: 'filled', manager: 'Nina Brooks' },
  { id: 'r22', title: 'Content Designer', dept: 'Design', type: 'steady', empType: 'fte', board: 3, planned: 3, actual: 3, variance: 0, status: 'filled', manager: 'Nina Brooks' },
  { id: 'r23', title: 'Design contractor', dept: 'Design', type: 'steady', empType: 'contractor', board: 0, planned: 1, actual: 1, variance: 0, status: 'filled' },
  // Sales
  { id: 'r24', title: 'Outbound AE', dept: 'Sales', type: 'new', empType: 'fte', board: 2, planned: 3, actual: 1, variance: -1, status: 'rejected', note: 'RF +1 above board rejected', startMonth: 'Sep', manager: 'Taylor Kim' },
  { id: 'r25', title: 'Enterprise BDR', dept: 'Sales', type: 'new', empType: 'fte', board: 1, planned: 1, actual: 0, variance: -1, status: 'approved', startMonth: 'Aug', manager: 'Taylor Kim' },
  { id: 'r26', title: 'Account Executive', dept: 'Sales', type: 'steady', empType: 'fte', board: 28, planned: 28, actual: 27, variance: -1, status: 'filled', manager: 'Taylor Kim' },
  { id: 'r27', title: 'SDR', dept: 'Sales', type: 'steady', empType: 'fte', board: 16, planned: 16, actual: 16, variance: 0, status: 'filled', manager: 'Taylor Kim' },
  { id: 'r28', title: 'Sales Manager', dept: 'Sales', type: 'steady', empType: 'fte', board: 6, planned: 6, actual: 6, variance: 0, status: 'filled' },
  { id: 'r29', title: 'Solutions Architect (contract)', dept: 'Sales', type: 'steady', empType: 'contractor', board: 0, planned: 3, actual: 3, variance: 0, status: 'filled', note: 'Outside board FTE' },
  // CS
  { id: 'r30', title: 'Technical Success Manager', dept: 'Customer Success', type: 'new', empType: 'fte', board: 1, planned: 1, actual: 0, variance: -1, status: 'pending', startMonth: 'Aug', manager: 'Priya Shah' },
  { id: 'r31', title: 'Customer Success Manager', dept: 'Customer Success', type: 'steady', empType: 'fte', board: 18, planned: 18, actual: 18, variance: 0, status: 'filled', manager: 'Priya Shah' },
  { id: 'r32', title: 'Technical Account Manager', dept: 'Customer Success', type: 'backfill', empType: 'fte', board: 8, planned: 8, actual: 7, variance: -1, status: 'open', replacing: 'Morgan Ellis', manager: 'Priya Shah' },
  { id: 'r33', title: 'CS Director', dept: 'Customer Success', type: 'steady', empType: 'fte', board: 2, planned: 2, actual: 2, variance: 0, status: 'filled' },
  { id: 'r34', title: 'CS contractor', dept: 'Customer Success', type: 'steady', empType: 'contractor', board: 0, planned: 2, actual: 2, variance: 0, status: 'filled' },
  // Marketing
  { id: 'r35', title: 'Lifecycle Marketing Manager', dept: 'Marketing', type: 'close', empType: 'fte', board: 1, planned: 0, actual: 0, variance: 0, status: 'filled', note: 'Role closed — budget reallocated' },
  { id: 'r36', title: 'Product Marketing Manager', dept: 'Marketing', type: 'steady', empType: 'fte', board: 4, planned: 4, actual: 4, variance: 0, status: 'filled' },
  { id: 'r37', title: 'Content Manager', dept: 'Marketing', type: 'steady', empType: 'fte', board: 3, planned: 3, actual: 3, variance: 0, status: 'filled' },
  { id: 'r38', title: 'Demand Gen Manager', dept: 'Marketing', type: 'steady', empType: 'fte', board: 4, planned: 4, actual: 4, variance: 0, status: 'filled' },
  { id: 'r39', title: 'Marketing contractor', dept: 'Marketing', type: 'steady', empType: 'contractor', board: 0, planned: 2, actual: 2, variance: 0, status: 'filled' },
  // Finance / People / Ops
  { id: 'r40', title: 'FP&A Analyst', dept: 'Finance', type: 'steady', empType: 'fte', board: 3, planned: 3, actual: 3, variance: 0, status: 'filled' },
  { id: 'r41', title: 'Controller', dept: 'Finance', type: 'steady', empType: 'fte', board: 1, planned: 1, actual: 1, variance: 0, status: 'filled' },
  { id: 'r42', title: 'Staff Accountant', dept: 'Finance', type: 'steady', empType: 'fte', board: 4, planned: 4, actual: 4, variance: 0, status: 'filled' },
  { id: 'r43', title: 'Systems contractor', dept: 'Finance', type: 'steady', empType: 'contractor', board: 0, planned: 1, actual: 1, variance: 0, status: 'filled' },
  { id: 'r44', title: 'HR Generalist', dept: 'People', type: 'steady', empType: 'fte', board: 3, planned: 3, actual: 3, variance: 0, status: 'filled' },
  { id: 'r45', title: 'HRBP', dept: 'People', type: 'steady', empType: 'fte', board: 2, planned: 2, actual: 2, variance: 0, status: 'filled' },
  { id: 'r46', title: 'People Ops Manager', dept: 'People', type: 'steady', empType: 'fte', board: 1, planned: 1, actual: 1, variance: 0, status: 'filled' },
  { id: 'r47', title: 'Recruiter', dept: 'People', type: 'steady', empType: 'fte', board: 2, planned: 2, actual: 2, variance: 0, status: 'filled' },
  { id: 'r48', title: 'Office Manager', dept: 'Operations', type: 'steady', empType: 'fte', board: 2, planned: 2, actual: 2, variance: 0, status: 'filled' },
  { id: 'r49', title: 'IT Support', dept: 'Operations', type: 'steady', empType: 'fte', board: 2, planned: 2, actual: 2, variance: 0, status: 'filled' },
  { id: 'r50', title: 'Ops contractor', dept: 'Operations', type: 'steady', empType: 'contractor', board: 0, planned: 1, actual: 1, variance: 0, status: 'filled' },
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
    id: 'RQ-2210',
    title: 'DevOps Engineer',
    dept: 'Engineering',
    requestedBy: 'Chris Park (Platform Lead)',
    createdAt: '2026-07-18',
    status: 'finance_review',
    type: 'replace',
    replacing: 'Contract DevOps (6-mo)',
    targetStart: '2026-09-01',
    rationale: 'Convert contractor spend into board-approved FTE.',
    daysOpen: 13,
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
  { id: 'a1', ts: '2026-07-31 16:42', actor: 'Alex Morgan (FP&A)', action: 'Snapshot locked', detail: 'July close — Board 310 / Planned 306 / Actual 301', dept: 'Finance' },
  { id: 'a2', ts: '2026-07-31 14:10', actor: 'Jordan Wells (HR)', action: 'Ticket advanced', detail: 'BF-1042 moved to HR review', dept: 'Engineering' },
  { id: 'a3', ts: '2026-07-30 11:05', actor: 'Alex Morgan (FP&A)', action: 'Variance explained', detail: 'Engineering −8 vs board documented', dept: 'Engineering' },
  { id: 'a4', ts: '2026-07-28 09:40', actor: 'Finance Committee', action: 'Rejected request', detail: 'RQ-2220 Outbound AE exceeds board plan', dept: 'Sales' },
  { id: 'a5', ts: '2026-07-26 13:00', actor: 'Chris Park', action: 'Opened net-new', detail: 'RQ-2228 QA Manager submitted', dept: 'Engineering' },
  { id: 'a6', ts: '2026-07-25 15:22', actor: 'Sam Rivera', action: 'Opened backfill', detail: 'BF-1048 Product Manager — Growth', dept: 'Product' },
  { id: 'a7', ts: '2026-07-20 10:15', actor: 'Priya Shah', action: 'Opened backfill', detail: 'BF-1055 Technical Account Manager', dept: 'Customer Success' },
  { id: 'a8', ts: '2026-07-18 10:00', actor: 'Chris Park', action: 'Replace ticket', detail: 'RQ-2210 contractor → FTE conversion', dept: 'Engineering' },
  { id: 'a9', ts: '2026-07-12 09:30', actor: 'Nina Brooks', action: 'Approved net-new', detail: 'RQ-2214 Product Designer approved by finance', dept: 'Design' },
  { id: 'a10', ts: '2026-06-30 17:00', actor: 'Board', action: 'Plan locked', detail: 'H2 board headcount locked: 310 (Jul) → 336 (Dec)' },
];

export const snapshots: Snapshot[] = [
  { id: 's0', label: 'FY25 close', asOf: '2025-12-31', board: 270, planned: 270, actual: 268, contractors: 21, variance: -2 },
  { id: 's1', label: 'Q1 close', asOf: '2026-03-31', board: 286, planned: 286, actual: 282, contractors: 22, variance: -4 },
  { id: 's2', label: 'Board lock', asOf: '2026-06-30', board: 298, planned: 298, actual: 292, contractors: 23, variance: -6 },
  { id: 's3', label: 'July mid-month', asOf: '2026-07-15', board: 310, planned: 309, actual: 290, contractors: 24, variance: -20 },
  { id: 's4', label: 'July close', asOf: '2026-07-31', board: 310, planned: 308, actual: 293, contractors: 24, variance: -17 },
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

export function summarizeScope(dept: Dept | 'all', empType: 'all' | EmpType) {
  const depts = dept === 'all' ? departments : departments.filter((d) => d.dept === dept);
  const roleSet = roles.filter((r) => {
    if (dept !== 'all' && r.dept !== dept) return false;
    if (empType === 'fte' && r.empType !== 'fte') return false;
    if (empType === 'contractor' && r.empType !== 'contractor') return false;
    return true;
  });

  if (empType === 'contractor') {
    const contractors = roleSet.reduce((s, r) => s + r.actual, 0);
    return {
      board: 0,
      planned: roleSet.reduce((s, r) => s + r.planned, 0),
      actual: contractors,
      contractors,
      variance: 0,
      openTickets: 0,
      openBackfills: 0,
      budgetUsd: depts.reduce((s, d) => s + d.budgetUsd, 0),
      spendUsd: depts.reduce((s, d) => s + d.spendUsd, 0),
      fyEndBoard: 0,
      fyEndPlanned: roleSet.reduce((s, r) => s + r.planned, 0),
      explain: 'Contractors are tracked separately and do not count against board-approved FTE.',
    };
  }

  const board = depts.reduce((s, d) => s + d.board, 0);
  const planned = depts.reduce((s, d) => s + d.planned, 0);
  const actual = depts.reduce((s, d) => s + d.actual, 0);
  const contractors = depts.reduce((s, d) => s + d.contractors, 0);
  const months = getMonthsForDept(dept, 'ytd');
  const fy = months[months.length - 1];

  return {
    board,
    planned,
    actual,
    contractors: empType === 'fte' ? 0 : contractors,
    variance: actual - board,
    openTickets: depts.reduce((s, d) => s + d.openTickets, 0),
    openBackfills: depts.reduce((s, d) => s + d.openBackfills, 0),
    budgetUsd: depts.reduce((s, d) => s + d.budgetUsd, 0),
    spendUsd: depts.reduce((s, d) => s + d.spendUsd, 0),
    fyEndBoard: fy.board,
    fyEndPlanned: fy.planned,
    explain:
      dept === 'all'
        ? 'You are under board plan. Delayed net-new, open backfills, and early exits explain the gap — contractors are excluded from board FTE.'
        : depts[0]?.varianceExplain ?? '',
  };
}
