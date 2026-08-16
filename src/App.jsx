import React, { useState, useMemo, useEffect } from 'react';
import { BrowserRouter, Routes, Route, useLocation, useNavigate } from 'react-router-dom';
import * as XLSX from 'xlsx';
import {
  ChevronRight, ChevronDown, Save, RotateCcw, Trash2, X, ArrowLeftRight, History, Eye, EyeOff, Download, Archive,
  Plus, Edit3, GripVertical, Search, Network, Activity as ActivityIcon,
} from 'lucide-react';
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend,
} from 'recharts';

// Backend REST API — see the accompanying budget-server/ project. Locally this
// defaults to localhost:4000. When you deploy the frontend, set VITE_API_BASE to
// your deployed backend's URL (e.g. https://rosebud-server.fly.dev) as a build-time
// environment variable — no code changes needed.
const API_BASE = `${import.meta.env.VITE_API_BASE || 'http://localhost:4000'}/api`;

/* ----------------------------------------------------------------------
   DESIGN TOKENS
---------------------------------------------------------------------- */
const COLORS = {
  bgChrome: '#10151F',
  bgChrome2: '#161D2C',
  chromeBorder: '#242C3D',
  surface: '#FFFFFF',
  surfaceAlt: '#F4F6F9',
  border: '#E3E7EE',
  textDark: '#141A22',
  textMuted: '#6B7480',
  textOnDark: '#EAEEF4',
  textOnDarkMuted: '#8A93A6',
  jade: '#2F9E77',
  jadeSoft: '#E5F4EE',
  brick: '#C0503F',
  brickSoft: '#FBEAE6',
  amber: '#C4791F',
  amberSoft: '#FBF1DE',
  violet: '#5A57D6',
  violetSoft: '#EEEDFB',
};

/* ----------------------------------------------------------------------
   DIMENSIONS
---------------------------------------------------------------------- */
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const QUARTER_MONTHS = { Q1: ['Jan', 'Feb', 'Mar'], Q2: ['Apr', 'May', 'Jun'], Q3: ['Jul', 'Aug', 'Sep'], Q4: ['Oct', 'Nov', 'Dec'] };
const SCENARIOS = ['Budget', 'Forecast', 'Actual'];

const ENTITIES = [
  { id: 'sales', name: 'Sales' },
  { id: 'marketing', name: 'Marketing' },
  { id: 'engineering', name: 'Engineering' },
  { id: 'ga', name: 'G&A' },
];
const ENTITY_OPTIONS = [{ id: 'company', name: 'Company (Consolidated)' }, ...ENTITIES];

// Product hierarchy: Category -> Product. Only Product Revenue (and its two
// drivers, Units Sold / Unit Price) carry this dimension — every other line
// item is entity-only, matching the backend's model.
const PRODUCT_CATEGORIES = [
  { id: 'hardware', name: 'Hardware' },
  { id: 'software', name: 'Software' },
];
const PRODUCTS = [
  { id: 'core_widget', name: 'Core Widget', category: 'hardware' },
  { id: 'widget_mini', name: 'Widget Mini', category: 'hardware' },
  { id: 'platform_license', name: 'Platform License', category: 'software' },
  { id: 'addon_modules', name: 'Add-on Modules', category: 'software' },
];
const PRODUCT_DIMENSIONED_ACCOUNTS = ['units_sold', 'unit_price'];

// Account hierarchy. type: 'rollup' (sums structural children), 'formula' (computed from drivers),
// 'input' (editable leaf). isDriver rows are formula inputs shown nested under their formula parent.
const ACCOUNTS = [
  { id: 'revenue', name: 'Revenue', parentId: null, type: 'rollup', unit: '$' },
  { id: 'product_revenue', name: 'Product Revenue', parentId: 'revenue', type: 'formula', unit: '$' },
  { id: 'units_sold', name: 'Units Sold', parentId: 'product_revenue', type: 'input', unit: 'u', isDriver: true },
  { id: 'unit_price', name: 'Unit Price', parentId: 'product_revenue', type: 'input', unit: '$', isDriver: true },
  { id: 'service_revenue', name: 'Service Revenue', parentId: 'revenue', type: 'input', unit: '$' },

  { id: 'expenses', name: 'Expenses', parentId: null, type: 'rollup', unit: '$' },
  { id: 'personnel', name: 'Personnel Costs', parentId: 'expenses', type: 'formula', unit: '$' },
  { id: 'headcount', name: 'Headcount', parentId: 'personnel', type: 'input', unit: 'FTE', isDriver: true },
  { id: 'avg_salary', name: 'Avg Monthly Cost / FTE', parentId: 'personnel', type: 'input', unit: '$', isDriver: true },
  { id: 'marketing_spend', name: 'Marketing Spend', parentId: 'expenses', type: 'input', unit: '$' },
  { id: 'software', name: 'Software & Tools', parentId: 'expenses', type: 'input', unit: '$' },
  { id: 'travel', name: 'Travel & Entertainment', parentId: 'expenses', type: 'input', unit: '$' },
];
const ACCOUNTS_BY_ID = Object.fromEntries(ACCOUNTS.map((a) => [a.id, a]));
// Accounts sensible to pick as a standalone report metric when Account isn't
// the row axis — excludes driver rows (units_sold, unit_price, headcount,
// avg_salary), which only make sense nested under their formula parent.
const PIVOT_ACCOUNT_OPTIONS = ACCOUNTS.filter((a) => !a.isDriver);
// When Product is the row axis, only accounts that actually vary by product
// are meaningful to pick — everything else would show an identical value on
// every product row, which is confusing rather than wrong.
const PIVOT_ACCOUNT_OPTIONS_PRODUCT_ROWS = ACCOUNTS.filter((a) => a.id === 'product_revenue' || PRODUCT_DIMENSIONED_ACCOUNTS.includes(a.id));

// Metadata for the 5 pivotable dimensions. Account and Scenario deliberately
// have no "Total" member: summing the flat account list would double-count
// rollups against their own children, and summing across scenarios has no
// sensible meaning.
const ALL_PIVOT_DIMS = ['account', 'entity', 'product', 'scenario', 'time'];
const PIVOT_DIMENSIONS = {
  account: { label: 'Account' },
  entity: { label: 'Entity', hasTotal: true, totalId: 'company', totalName: 'Company (Total)' },
  product: { label: 'Product', hasTotal: true, totalId: 'all', totalName: 'All Products (Total)' },
  time: { label: 'Time', hasTotal: true, totalId: 'FY', totalName: 'FY (Total)' },
  scenario: { label: 'Scenario' },
};

// The list of members for a dimension when it's placed on an axis. Pure and
// reused for both rows and columns — whichever dimension lands where.
function pivotMembers(dim, granularity) {
  if (dim === 'account') return PIVOT_ACCOUNT_OPTIONS.map((a) => ({ id: a.id, name: a.name, unit: a.unit }));
  if (dim === 'entity') return ENTITIES.map((e) => ({ id: e.id, name: e.name }));
  if (dim === 'product') return PRODUCTS.map((p) => ({ id: p.id, name: p.name }));
  if (dim === 'scenario') return SCENARIOS.map((s) => ({ id: s, name: s }));
  if (dim === 'time') {
    if (granularity === 'Annual') return []; // only the FY total exists — no sub-periods to list
    if (granularity === 'Quarterly') return ['Q1', 'Q2', 'Q3', 'Q4'].map((q) => ({ id: q, name: q }));
    return MONTHS.map((m) => ({ id: m, name: m }));
  }
  return [];
}

// Computes one cell's value for any assignment of the 5 dimensions to
// {scenario, entity, account, period, product} — a thin wrapper around the
// same getPeriodValue used by the default Account×Time grid, just with every
// axis passed explicitly instead of some being fixed by closure.
function getPivotCellValue(allScenarioValues, scenario, entityId, accountId, period, productId) {
  const data = allScenarioValues[scenario] || {};
  return getPeriodValue(data, entityId, accountId, period, productId);
}

// Builds a tree for an ordered list of nested dimensions (e.g. ['account',
// 'entity'] = Account outer, Entity inner). Each node carries `path` — the
// full {dim: id} map from the root down to that node — which is exactly what
// a cell needs to resolve its value. Any dimension in `dims` beyond the
// first becomes nested children of every member (and Total, where that
// dimension has one) at the level above. `subsets` is an optional
// {dim: Set(memberIds)} map — when a dimension has one, only its selected
// members (and Total, if included) appear; dimensions with no entry show
// everything, same as before subsets existed.
function buildPivotTree(dims, granularity, pathSoFar, subsets) {
  if (dims.length === 0) return null;
  const [dim, ...rest] = dims;
  const meta = PIVOT_DIMENSIONS[dim];
  const subset = subsets && subsets[dim];
  let members = pivotMembers(dim, granularity).map((m) => ({ id: m.id, name: m.name, isTotal: false }));
  if (meta?.hasTotal) members.push({ id: meta.totalId, name: meta.totalName, isTotal: true });
  if (subset) members = members.filter((m) => subset.has(m.id));
  return members.map((m) => {
    const path = { ...pathSoFar, [dim]: m.id };
    return { dim, id: m.id, name: m.name, isTotal: m.isTotal, path, children: rest.length ? buildPivotTree(rest, granularity, path, subsets) : null };
  });
}
// Number of leaf columns/rows a node ultimately expands into — used as the
// colspan for column headers, and (implicitly, one row per leaf) for rows.
function countPivotLeaves(node) {
  if (!node.children || node.children.length === 0) return 1;
  return node.children.reduce((s, c) => s + countPivotLeaves(c), 0);
}
// Flattens a tree down to just its leaf nodes, each still carrying its full
// path — this is the actual list of columns (or, for rows, the base list
// before indentation/grouping is applied for display).
function flattenPivotLeaves(nodes) {
  const leaves = [];
  (nodes || []).forEach((node) => {
    if (node.children && node.children.length) leaves.push(...flattenPivotLeaves(node.children));
    else leaves.push(node);
  });
  return leaves;
}
// One header row per nesting level, each cell's colSpan equal to how many
// leaf columns it ultimately covers — the standard multi-level column
// header pattern, generalized to any depth.
function pivotColumnHeaderRows(tree) {
  const rows = [];
  (function walk(nodes, level) {
    if (!nodes || nodes.length === 0) return;
    rows[level] = rows[level] || [];
    nodes.forEach((node) => {
      rows[level].push({ name: node.name, isTotal: node.isTotal, colSpan: countPivotLeaves(node), path: node.path, hadChildren: !!node.hadChildren || !!node.children });
      if (node.children) walk(node.children, level + 1);
    });
  })(tree, 0);
  return rows;
}
// Returns a copy of the tree where any node whose path-key is in
// `collapsedSet` has its children stripped — collapsing it into a single
// leaf for colspan/flatten purposes, while remembering `hadChildren` so the
// header can still show a clickable chevron for it.
function pruneCollapsedPivotTree(nodes, collapsedSet) {
  if (!nodes) return nodes;
  return nodes.map((node) => {
    const key = JSON.stringify(node.path);
    if (collapsedSet.has(key)) return { ...node, children: null, hadChildren: true };
    return { ...node, children: node.children ? pruneCollapsedPivotTree(node.children, collapsedSet) : null, hadChildren: !!node.children };
  });
}
// Fills in any dimension in `dimOrder` missing from `path` with its "all
// members" total id, so a collapsed/group node can show a real rolled-up
// value instead of a blank. Returns null if a missing dimension has no such
// total (Account, Scenario) — rolling those up flatly would be misleading,
// so the caller shows "—" instead.
function rollupPathFor(path, dimOrder) {
  const filled = { ...path };
  for (const dim of dimOrder) {
    if (!(dim in filled)) {
      const meta = PIVOT_DIMENSIONS[dim];
      if (meta?.hasTotal) filled[dim] = meta.totalId;
      else return null;
    }
  }
  return filled;
}

/* ----------------------------------------------------------------------
   MODEL ENGINE  (pure functions — operate on a plain data object)
---------------------------------------------------------------------- */
function childrenOf(id) { return ACCOUNTS.filter((a) => a.parentId === id); }
function hasChildren(id) { return ACCOUNTS.some((a) => a.parentId === id); }
// A rollup row stays visible if any descendant is in the subset, even if the
// rollup itself isn't checked — otherwise a partial selection would leave
// its children showing with no parent row for context.
function accountMatchesSubset(accountId, subset) {
  if (subset.has(accountId)) return true;
  return childrenOf(accountId).some((c) => accountMatchesSubset(c.id, subset));
}
function topAncestor(id) {
  let a = ACCOUNTS_BY_ID[id];
  while (a && a.parentId) a = ACCOUNTS_BY_ID[a.parentId];
  return a ? a.id : id;
}
function computeFormula(accountId, get, getForProduct, product) {
  if (accountId === 'product_revenue') {
    if (product === 'all') {
      return PRODUCTS.reduce((s, p) => s + getForProduct('units_sold', p.id) * getForProduct('unit_price', p.id), 0);
    }
    return getForProduct('units_sold', product) * getForProduct('unit_price', product);
  }
  if (accountId === 'personnel') return get('headcount') * get('avg_salary');
  return 0;
}
// `product` selects which product's Units Sold / Unit Price / Product Revenue to
// show — it has no effect on any other line, and rollups (Revenue, Expenses) and
// the Company-consolidated view always sum every product in full regardless of
// what's selected, so totals are never a partial/misleading slice.
function getValueAt(data, entityId, accountId, month, product) {
  const prod = product || 'all';
  if (entityId === 'company') {
    // Rate/average measures (cost per FTE, price per unit) can't be summed
    // across entities the way dollar or unit totals can — a $9,500/FTE
    // department plus a $12,500/FTE department isn't a $22,000/FTE company.
    // These are derived as a true weighted average: the dependent dollar
    // total divided by the dependent volume total, both of which are
    // genuinely additive and correctly summed via the recursive call below.
    if (accountId === 'avg_salary') {
      const totalPersonnel = getValueAt(data, 'company', 'personnel', month, prod);
      const totalHeadcount = getValueAt(data, 'company', 'headcount', month, prod);
      return totalHeadcount !== 0 ? totalPersonnel / totalHeadcount : 0;
    }
    if (accountId === 'unit_price') {
      const totalRevenue = getValueAt(data, 'company', 'product_revenue', month, prod);
      const totalUnits = getValueAt(data, 'company', 'units_sold', month, prod);
      return totalUnits !== 0 ? totalRevenue / totalUnits : 0;
    }
    return ENTITIES.reduce((s, e) => s + getValueAt(data, e.id, accountId, month, prod), 0);
  }
  const acc = ACCOUNTS_BY_ID[accountId];
  if (!acc) return 0;
  if (acc.type === 'rollup') {
    // Rollups always consolidate every product — the selector never partially filters a total.
    return childrenOf(accountId).reduce((s, c) => s + getValueAt(data, entityId, c.id, month, 'all'), 0);
  }
  if (acc.type === 'formula') {
    return computeFormula(
      accountId,
      (depId) => getValueAt(data, entityId, depId, month, 'all'),
      (depId, p) => getValueAt(data, entityId, depId, month, p),
      prod,
    );
  }
  if (PRODUCT_DIMENSIONED_ACCOUNTS.includes(accountId)) {
    const byProduct = (data[entityId] && data[entityId][accountId]) || {};
    if (prod === 'all') {
      return PRODUCTS.reduce((s, p) => s + ((byProduct[p.id] && byProduct[p.id][month]) || 0), 0);
    }
    return (byProduct[prod] && byProduct[prod][month]) || 0;
  }
  return (data[entityId] && data[entityId][accountId] && data[entityId][accountId].none && data[entityId][accountId].none[month]) || 0;
}
function getPeriodValue(data, entityId, accountId, period, product) {
  if (MONTHS.includes(period)) return getValueAt(data, entityId, accountId, period, product);
  const periodMonths = period === 'FY' ? MONTHS : (QUARTER_MONTHS[period] || []);
  // Same non-additive problem, across time instead of across entities — an
  // average monthly rate isn't meaningful summed across 12 months either.
  // The weighted average here (sum of dollars ÷ sum of volume across the
  // period) is the mathematically correct time-weighted average.
  if (accountId === 'avg_salary') {
    const totalPersonnel = periodMonths.reduce((s, m) => s + getValueAt(data, entityId, 'personnel', m, product), 0);
    const totalHeadcount = periodMonths.reduce((s, m) => s + getValueAt(data, entityId, 'headcount', m, product), 0);
    return totalHeadcount !== 0 ? totalPersonnel / totalHeadcount : 0;
  }
  if (accountId === 'unit_price') {
    const totalRevenue = periodMonths.reduce((s, m) => s + getValueAt(data, entityId, 'product_revenue', m, product), 0);
    const totalUnits = periodMonths.reduce((s, m) => s + getValueAt(data, entityId, 'units_sold', m, product), 0);
    return totalUnits !== 0 ? totalRevenue / totalUnits : 0;
  }
  return periodMonths.reduce((s, m) => s + getValueAt(data, entityId, accountId, m, product), 0);
}

/* ----------------------------------------------------------------------
   FORMATTERS
---------------------------------------------------------------------- */
// Reads the payload out of our own JWT for UI purposes only (e.g. "is this row
// me?") — never used for authorization, which the backend always re-checks.
function decodeToken(token) {
  try {
    const payload = token.split('.')[1];
    return JSON.parse(atob(payload.replace(/-/g, '+').replace(/_/g, '/')));
  } catch {
    return null;
  }
}

function fmtMoney(v) {
  const r = Math.round(v);
  return `${r < 0 ? '-' : ''}$${Math.abs(r).toLocaleString()}`;
}
function fmtNum(v) { return Math.round(v).toLocaleString(); }
function fmtCell(unit, v) {
  if (unit === '$') return fmtMoney(v);
  if (unit === 'FTE') return `${fmtNum(v)} FTE`;
  if (unit === 'u') return `${fmtNum(v)} u`;
  return fmtNum(v);
}
function fmtPct(v) {
  if (!isFinite(v)) return '—';
  return `${v >= 0 ? '+' : ''}${v.toFixed(1)}%`;
}

/* ----------------------------------------------------------------------
   VISUAL HELPERS
---------------------------------------------------------------------- */
function railColor(account) {
  if (account.type === 'rollup') return COLORS.jade;
  if (account.type === 'formula') return COLORS.violet;
  if (account.isDriver) return COLORS.amber;
  return COLORS.border;
}
function rowBg(account) {
  if (account.type === 'rollup') return COLORS.surfaceAlt;
  if (account.type === 'formula') return COLORS.violetSoft;
  if (account.isDriver) return COLORS.amberSoft;
  return COLORS.surface;
}
function rowNameStyle(account) {
  if (account.type === 'rollup') return { fontWeight: 700, letterSpacing: '0.02em' };
  if (account.type === 'formula') return { fontWeight: 600 };
  if (account.isDriver) return { fontStyle: 'italic', color: COLORS.textMuted, fontWeight: 400 };
  return { fontWeight: 400 };
}

// A pill-shaped "chip" wrapping a real <select> (kept native for accessibility
// and keyboard support — only the visual chrome is custom). Called inline as
// a plain function, not rendered as a component, so it never remounts.
function chipSelect(label, value, onChange, optionsNode, opts) {
  return (
    <div
      title={opts?.title}
      style={{ display: 'flex', alignItems: 'center', gap: 6, background: COLORS.bgChrome2, border: `1px solid ${COLORS.chromeBorder}`, borderRadius: 999, padding: '5px 12px 5px 12px' }}
    >
      <span style={{ fontSize: 9.5, fontWeight: 700, letterSpacing: '.06em', color: COLORS.textOnDarkMuted, textTransform: 'uppercase' }}>{label}</span>
      <select
        className="lw-select"
        value={value}
        onChange={onChange}
        style={{ background: 'none', color: COLORS.textOnDark, border: 'none', fontSize: 12.5, fontWeight: 600, cursor: 'pointer', outline: 'none' }}
      >
        {optionsNode}
      </select>
    </div>
  );
}

/* ----------------------------------------------------------------------
   APP
---------------------------------------------------------------------- */
function Workspace({ token, workspaceName, onLogout, onAuthError, onSwitchWorkspace }) {
  const location = useLocation();
  const navigate = useNavigate();
  const showVersions = location.pathname === '/versions';
  const showMembers = location.pathname === '/members';
  const showBackups = location.pathname === '/backups';
  const showHierarchy = location.pathname === '/hierarchy';
  const showActivity = location.pathname === '/activity';

  const [values, setValues] = useState({});
  const [workspaceLoaded, setWorkspaceLoaded] = useState(false);
  const [loadError, setLoadError] = useState(null);
  const [saveStatus, setSaveStatus] = useState('idle'); // idle | saving | saved | error
  const [currentScenario, setCurrentScenario] = useState('Budget');
  const [currentEntity, setCurrentEntity] = useState('company');
  const [currentProduct, setCurrentProduct] = useState('all');
  // Ordered lists (outer to inner nesting) of which dimensions sit on each
  // axis. Whatever isn't in either list is a Filter. Scenario's filter value
  // is currentScenario itself (used elsewhere by Versions/Backups), and
  // Entity/Product reuse their existing long-standing state — only Account
  // and Time need new filter-value state below.
  const [rowsOrder, setRowsOrder] = useState(['account']);
  const [colsOrder, setColsOrder] = useState(['time']);
  const [filterAccount, setFilterAccount] = useState('revenue');
  const [filterTime, setFilterTime] = useState('FY');
  const [dragChip, setDragChip] = useState(null);
  // { [dim]: Set(memberIds) } — a dimension only appears here once the user
  // has actually customized its subset; absence means "show everything",
  // same as before this feature existed.
  const [pivotSubsets, setPivotSubsets] = useState({});
  const [subsetEditorDim, setSubsetEditorDim] = useState(null);
  const [subsetEditorDraft, setSubsetEditorDraft] = useState(null);
  const [pivotCollapsed, setPivotCollapsed] = useState(() => new Set());
  const [granularity, setGranularity] = useState('Monthly');
  const [expanded, setExpanded] = useState(() => new Set(['revenue', 'expenses']));
  const [compareMode, setCompareMode] = useState(false);
  const [compareTarget, setCompareTarget] = useState('Actual');
  const [comparePeriod, setComparePeriod] = useState('FY');
  const [versions, setVersions] = useState([]);
  const [versionLabel, setVersionLabel] = useState('');
  const [role, setRole] = useState('viewer');
  const [myWorkspaces, setMyWorkspaces] = useState([]);
  const [members, setMembers] = useState([]);
  const [activityEntries, setActivityEntries] = useState([]);
  const [activityLoading, setActivityLoading] = useState(false);
  const [inviteEmail, setInviteEmail] = useState('');
  const [inviteRole, setInviteRole] = useState('editor');
  const [membersError, setMembersError] = useState('');
  const [lastInviteLink, setLastInviteLink] = useState('');
  const [backups, setBackups] = useState([]);
  const [backupLabel, setBackupLabel] = useState('');
  const [backupsError, setBackupsError] = useState('');
  const [creatingBackup, setCreatingBackup] = useState(false);
  const [confirmingRestoreId, setConfirmingRestoreId] = useState(null);

  useEffect(() => {
    fetch(`${API_BASE}/workspace`, { headers: { Authorization: `Bearer ${token}` } })
      .then((r) => {
        if (r.status === 401) { onAuthError(); throw new Error('session expired'); }
        if (!r.ok) throw new Error(`Server responded ${r.status}`);
        return r.json();
      })
      .then((data) => {
        setValues(data.values);
        setVersions(data.versions);
        setRole(data.workspace.role);
        setWorkspaceLoaded(true);
      })
      .catch((err) => setLoadError(err.message));
    fetch(`${API_BASE}/auth/my-workspaces`, { headers: { Authorization: `Bearer ${token}` } })
      .then((r) => (r.ok ? r.json() : { workspaces: [] }))
      .then((data) => setMyWorkspaces(data.workspaces || []))
      .catch(() => {});
  }, [token]);

  const canManageMembers = role === 'admin';
  const canEditData = role === 'editor' || role === 'power' || role === 'admin';
  const canManageBackups = role === 'power' || role === 'admin';
  const canManageHierarchy = role === 'power' || role === 'admin';
  const myUserId = useMemo(() => decodeToken(token)?.userId, [token]);

  function loadMembers() {
    fetch(`${API_BASE}/workspace/members`, { headers: { Authorization: `Bearer ${token}` } })
      .then((r) => r.json())
      .then((data) => setMembers(data.members || []))
      .catch(() => {});
  }
  function loadActivity() {
    setActivityLoading(true);
    fetch(`${API_BASE}/workspace/activity`, { headers: { Authorization: `Bearer ${token}` } })
      .then((r) => r.json())
      .then((data) => { setActivityEntries(data.entries || []); setActivityLoading(false); })
      .catch(() => setActivityLoading(false));
  }
  function sendInvite() {
    setMembersError('');
    setLastInviteLink('');
    fetch(`${API_BASE}/workspace/invite`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify({ email: inviteEmail.trim(), role: inviteRole }),
    })
      .then((r) => r.json().then((data) => ({ ok: r.ok, data })))
      .then(({ ok, data }) => {
        if (!ok) { setMembersError(data.error || 'Could not send invite.'); return; }
        setInviteEmail('');
        setLastInviteLink(data.inviteLink);
        loadMembers();
      })
      .catch(() => setMembersError('Could not send invite.'));
  }
  function changeMemberRole(userId, newRole) {
    setMembersError('');
    fetch(`${API_BASE}/workspace/members/${userId}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify({ role: newRole }),
    })
      .then((r) => r.json().then((data) => ({ ok: r.ok, data })))
      .then(({ ok, data }) => { if (!ok) setMembersError(data.error || 'Could not update role.'); else loadMembers(); })
      .catch(() => setMembersError('Could not update role.'));
  }
  function removeMember(userId) {
    setMembersError('');
    fetch(`${API_BASE}/workspace/members/${userId}`, { method: 'DELETE', headers: { Authorization: `Bearer ${token}` } })
      .then((r) => r.json().then((data) => ({ ok: r.ok, data })))
      .then(({ ok, data }) => { if (!ok) setMembersError(data.error || 'Could not remove member.'); else loadMembers(); })
      .catch(() => setMembersError('Could not remove member.'));
  }
  function switchWorkspace(workspaceId) {
    fetch(`${API_BASE}/auth/switch-workspace`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify({ workspaceId }),
    })
      .then((r) => r.json())
      .then((data) => onSwitchWorkspace(data))
      .catch(() => setSaveStatus('error'));
  }
  function triggerJsonDownload(data, filename) {
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }
  function downloadBackup() {
    setMembersError('');
    fetch(`${API_BASE}/workspace/backup`, { headers: { Authorization: `Bearer ${token}` } })
      .then((r) => { if (!r.ok) throw new Error('Could not generate backup.'); return r.json(); })
      .then((backup) => {
        const dateStamp = new Date().toISOString().slice(0, 10);
        const safeName = (workspaceName || 'rosebud').toLowerCase().replace(/[^a-z0-9]+/g, '-');
        triggerJsonDownload(backup, `rosebud-backup-${safeName}-${dateStamp}.json`);
      })
      .catch(() => setMembersError('Could not download backup.'));
  }

  function loadBackups() {
    fetch(`${API_BASE}/workspace/backups`, { headers: { Authorization: `Bearer ${token}` } })
      .then((r) => r.json())
      .then((data) => setBackups(data.backups || []))
      .catch(() => {});
  }
  function createBackup() {
    setBackupsError('');
    setCreatingBackup(true);
    fetch(`${API_BASE}/workspace/backups`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify({ label: backupLabel.trim() }),
    })
      .then((r) => r.json().then((data) => ({ ok: r.ok, data })))
      .then(({ ok, data }) => {
        setCreatingBackup(false);
        if (!ok) { setBackupsError(data.error || 'Could not create backup.'); return; }
        setBackupLabel('');
        loadBackups();
      })
      .catch(() => { setCreatingBackup(false); setBackupsError('Could not create backup.'); });
  }
  function downloadStoredBackup(id, label) {
    setBackupsError('');
    fetch(`${API_BASE}/workspace/backups/${id}`, { headers: { Authorization: `Bearer ${token}` } })
      .then((r) => { if (!r.ok) throw new Error('Could not download backup.'); return r.json(); })
      .then((backup) => {
        const dateStamp = new Date().toISOString().slice(0, 10);
        const safeName = (label || 'backup').toLowerCase().replace(/[^a-z0-9]+/g, '-');
        triggerJsonDownload(backup, `rosebud-${safeName}-${dateStamp}.json`);
      })
      .catch(() => setBackupsError('Could not download backup.'));
  }
  function deleteBackup(id) {
    setBackupsError('');
    fetch(`${API_BASE}/workspace/backups/${id}`, { method: 'DELETE', headers: { Authorization: `Bearer ${token}` } })
      .then((r) => r.json().then((data) => ({ ok: r.ok, data })))
      .then(({ ok, data }) => { if (!ok) setBackupsError(data.error || 'Could not delete backup.'); else loadBackups(); })
      .catch(() => setBackupsError('Could not delete backup.'));
  }
  function restoreBackup(id) {
    if (confirmingRestoreId !== id) { setConfirmingRestoreId(id); return; }
    setConfirmingRestoreId(null);
    setBackupsError('');
    fetch(`${API_BASE}/workspace/backups/${id}/restore`, { method: 'PUT', headers: { Authorization: `Bearer ${token}` } })
      .then((r) => r.json().then((data) => ({ ok: r.ok, data })))
      .then(({ ok, data }) => {
        if (!ok) { setBackupsError(data.error || 'Could not restore backup.'); return; }
        setValues(data.values);
      })
      .catch(() => setBackupsError('Could not restore backup.'));
  }

  const liveData = values[currentScenario];
  const isDefaultConfig = rowsOrder.length === 1 && rowsOrder[0] === 'account' && colsOrder.length === 1 && colsOrder[0] === 'time';

  const columns = useMemo(() => {
    let cols;
    if (granularity === 'Monthly') cols = [...MONTHS, 'FY'];
    else if (granularity === 'Quarterly') cols = ['Q1', 'Q2', 'Q3', 'Q4', 'FY'];
    else cols = ['FY'];
    const subset = pivotSubsets.time;
    return subset ? cols.filter((c) => subset.has(c)) : cols;
  }, [granularity, pivotSubsets.time]);

  const visibleRows = useMemo(() => {
    const rows = [];
    const subset = pivotSubsets.account;
    (function walk(parentId, depth) {
      ACCOUNTS.filter((a) => a.parentId === parentId).forEach((a) => {
        if (subset && !accountMatchesSubset(a.id, subset)) return;
        rows.push({ ...a, depth });
        if (hasChildren(a.id) && expanded.has(a.id)) walk(a.id, depth + 1);
      });
    })(null, 0);
    return rows;
  }, [expanded, pivotSubsets.account]);

  // Moves a dimension to Rows, Columns, or Filter — it's removed from
  // wherever it was and appended as the innermost level of its new axis (if
  // moved to Rows/Columns), supporting any number of nested dimensions.
  function moveDimToZone(dim, targetZone) {
    setRowsOrder((prev) => {
      const without = prev.filter((d) => d !== dim);
      return targetZone === 'rows' ? [...without, dim] : without;
    });
    setColsOrder((prev) => {
      const without = prev.filter((d) => d !== dim);
      return targetZone === 'cols' ? [...without, dim] : without;
    });
    if (dim === 'product' && (targetZone === 'rows' || targetZone === 'cols') && !PIVOT_ACCOUNT_OPTIONS_PRODUCT_ROWS.some((a) => a.id === filterAccount)) {
      setFilterAccount('product_revenue');
    }
    setCompareMode(false);
  }
  // Swaps a dimension earlier/later within its own axis's nesting order.
  function reorderDim(dim, direction, order, setOrder) {
    setOrder((prev) => {
      const idx = prev.indexOf(dim);
      const swapWith = idx + direction;
      if (idx === -1 || swapWith < 0 || swapWith >= prev.length) return prev;
      const next = [...prev];
      [next[idx], next[swapWith]] = [next[swapWith], next[idx]];
      return next;
    });
  }
  function togglePivotExpand(path) {
    const key = JSON.stringify(path);
    setPivotCollapsed((prev) => { const n = new Set(prev); n.has(key) ? n.delete(key) : n.add(key); return n; });
  }

  // The full member list (including Total, where applicable) for a
  // dimension — what the subset editor's checklist is built from.
  function subsetAllMembers(dim) {
    const meta = PIVOT_DIMENSIONS[dim];
    const members = pivotMembers(dim, granularity).map((m) => ({ id: m.id, name: m.name }));
    if (meta?.hasTotal) members.push({ id: meta.totalId, name: meta.totalName });
    return members;
  }
  function openSubsetEditor(dim) {
    const allIds = subsetAllMembers(dim).map((m) => m.id);
    const current = pivotSubsets[dim];
    setSubsetEditorDraft(new Set(current ? [...current] : allIds));
    setSubsetEditorDim(dim);
  }
  function toggleSubsetDraftMember(id) {
    setSubsetEditorDraft((prev) => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n; });
  }
  function applySubsetEditor() {
    const allIds = subsetAllMembers(subsetEditorDim).map((m) => m.id);
    const isFullSet = allIds.length === subsetEditorDraft.size && allIds.every((id) => subsetEditorDraft.has(id));
    setPivotSubsets((prev) => {
      const next = { ...prev };
      if (isFullSet) delete next[subsetEditorDim];
      else next[subsetEditorDim] = subsetEditorDraft;
      return next;
    });
    setSubsetEditorDim(null);
    setSubsetEditorDraft(null);
  }
  function closeSubsetEditor() {
    setSubsetEditorDim(null);
    setSubsetEditorDraft(null);
  }

  // The small value-picker embedded in a chip when that dimension is sitting
  // in the Filter zone. stopPropagation on mousedown keeps clicking the
  // select from being interpreted as the start of a drag.
  function renderFilterValuePicker(dim) {
    const stop = (e) => e.stopPropagation();
    const style = { background: 'none', border: 'none', color: 'inherit', fontWeight: 700, fontSize: 12, cursor: 'pointer', outline: 'none' };
    const subset = pivotSubsets[dim];
    const inSubset = (id) => !subset || subset.has(id);
    if (dim === 'entity') {
      return (
        <select value={currentEntity} onChange={(e) => setCurrentEntity(e.target.value)} onMouseDown={stop} className="lw-select" style={style}>
          {ENTITY_OPTIONS.filter((e) => inSubset(e.id)).map((e) => <option key={e.id} value={e.id}>{e.name}</option>)}
        </select>
      );
    }
    if (dim === 'product') {
      return (
        <select value={currentProduct} onChange={(e) => setCurrentProduct(e.target.value)} onMouseDown={stop} className="lw-select" style={style}>
          {inSubset('all') && <option value="all">All</option>}
          {PRODUCT_CATEGORIES.map((cat) => {
            const inCat = PRODUCTS.filter((p) => p.category === cat.id && inSubset(p.id));
            return inCat.length ? (
              <optgroup key={cat.id} label={cat.name}>
                {inCat.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
              </optgroup>
            ) : null;
          })}
        </select>
      );
    }
    if (dim === 'account') {
      const baseOptions = (rowsOrder.includes('product') || colsOrder.includes('product')) ? PIVOT_ACCOUNT_OPTIONS_PRODUCT_ROWS : PIVOT_ACCOUNT_OPTIONS;
      return (
        <select value={filterAccount} onChange={(e) => setFilterAccount(e.target.value)} onMouseDown={stop} className="lw-select" style={style}>
          {baseOptions.filter((a) => inSubset(a.id)).map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
        </select>
      );
    }
    if (dim === 'scenario') {
      return (
        <select value={currentScenario} onChange={(e) => setCurrentScenario(e.target.value)} onMouseDown={stop} className="lw-select" style={style}>
          {SCENARIOS.filter((s) => inSubset(s)).map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
      );
    }
    if (dim === 'time') {
      return (
        <select value={filterTime} onChange={(e) => setFilterTime(e.target.value)} onMouseDown={stop} className="lw-select" style={style}>
          {[...MONTHS, 'Q1', 'Q2', 'Q3', 'Q4', 'FY'].filter((p) => inSubset(p)).map((p) => <option key={p} value={p}>{p}</option>)}
        </select>
      );
    }
    return null;
  }

  function toggleExpand(id) {
    setExpanded((prev) => {
      const n = new Set(prev);
      n.has(id) ? n.delete(id) : n.add(id);
      return n;
    });
  }

  function updateValue(entityId, accountId, month, raw, productOverride) {
    const num = raw === '' ? 0 : parseFloat(raw);
    const value = isNaN(num) ? 0 : num;
    const productKey = PRODUCT_DIMENSIONED_ACCOUNTS.includes(accountId) ? (productOverride ?? currentProduct) : 'none';

    // Optimistic local update so typing feels instant.
    setValues((prev) => {
      const next = { ...prev };
      const scen = { ...next[currentScenario] };
      const ent = { ...scen[entityId] };
      const accData = { ...ent[accountId] };
      accData[productKey] = { ...accData[productKey], [month]: value };
      ent[accountId] = accData;
      scen[entityId] = ent;
      next[currentScenario] = scen;
      return next;
    });

    setSaveStatus('saving');
    fetch(`${API_BASE}/cell`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify({ scenario: currentScenario, entity: entityId, account: accountId, month, value, product: productKey }),
    })
      .then((r) => { if (!r.ok) throw new Error('save failed'); setSaveStatus('saved'); })
      .catch(() => setSaveStatus('error'));
  }

  function resolveData(target) {
    if (target.startsWith('v:')) {
      const id = Number(target.slice(2));
      const v = versions.find((x) => x.id === id);
      return v ? v.data : {};
    }
    return values[target] || {};
  }
  function targetLabel(target) {
    if (target.startsWith('v:')) {
      const id = Number(target.slice(2));
      const v = versions.find((x) => x.id === id);
      return v ? `${v.label} (saved)` : 'Version';
    }
    return target;
  }

  function saveVersion() {
    const label = versionLabel.trim();
    fetch(`${API_BASE}/versions`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify({ scenario: currentScenario, label }),
    })
      .then((r) => r.json())
      .then((created) => setVersions((prev) => [...prev, created]))
      .catch(() => setSaveStatus('error'));
    setVersionLabel('');
  }
  function restoreVersion(v) {
    fetch(`${API_BASE}/versions/${v.id}/restore`, { method: 'PUT', headers: { Authorization: `Bearer ${token}` } })
      .then((r) => r.json())
      .then((res) => setValues((prev) => ({ ...prev, [res.scenario]: res.data })))
      .catch(() => setSaveStatus('error'));
  }
  function deleteVersion(id) {
    fetch(`${API_BASE}/versions/${id}`, { method: 'DELETE', headers: { Authorization: `Bearer ${token}` } })
      .then(() => setVersions((prev) => prev.filter((v) => v.id !== id)))
      .catch(() => setSaveStatus('error'));
  }
  function compareWithVersion(v) {
    setCompareMode(true);
    setCompareTarget(`v:${v.id}`);
    navigate('/');
  }

  if (loadError) {
    return (
      <div style={{ fontFamily: "'IBM Plex Sans', sans-serif", padding: 40, color: COLORS.brick }}>
        Couldn't reach the budget server at {API_BASE}. Make sure it's running (<code>npm start</code> in <code>budget-server/</code>), then refresh this page.
        <div style={{ color: COLORS.textMuted, fontSize: 13, marginTop: 8 }}>{loadError}</div>
      </div>
    );
  }
  if (!workspaceLoaded) {
    return (
      <div style={{ fontFamily: "'IBM Plex Sans', sans-serif", padding: 40, color: COLORS.textMuted }}>
        Loading workspace…
      </div>
    );
  }

  const canEdit = canEditData && currentEntity !== 'company' && granularity === 'Monthly' && !compareMode;
  // When Entity is the row axis, its filter chip is hidden — fall back to the
  // company-wide total for KPIs/chart rather than showing a stale, no-longer-
  // visible entity selection.
  // KPIs and the chart always show the company-wide total — a stable
  // headline summary that doesn't change meaning depending on how the grid
  // below happens to be pivoted at the moment.
  const kpiEntityContext = 'company';

  const kpiRevenue = getPeriodValue(liveData, kpiEntityContext, 'revenue', 'FY');
  const kpiExpenses = getPeriodValue(liveData, kpiEntityContext, 'expenses', 'FY');
  const kpiNet = kpiRevenue - kpiExpenses;
  const kpiMargin = kpiRevenue !== 0 ? (kpiNet / kpiRevenue) * 100 : 0;

  const chartData = MONTHS.map((m) => {
    const rev = getValueAt(liveData, kpiEntityContext, 'revenue', m);
    const exp = getValueAt(liveData, kpiEntityContext, 'expenses', m);
    return { month: m, Revenue: rev, Expenses: exp, 'Net Income': rev - exp };
  });

  const compareData = compareMode ? resolveData(compareTarget) : null;
  const compareOptions = [
    ...SCENARIOS.map((s) => ({ key: s, label: s })),
    ...versions.map((v) => ({ key: `v:${v.id}`, label: `${v.label} (saved)` })),
  ].filter((o) => o.key !== currentScenario || o.key.startsWith('v:'));

  return (
    <div style={{ fontFamily: "'IBM Plex Sans', sans-serif", background: COLORS.surfaceAlt, color: COLORS.textDark, minHeight: '100%', overflowX: 'hidden' }} className="w-full">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
        .lw-num { font-family: 'IBM Plex Mono', monospace; font-variant-numeric: tabular-nums; }
        .lw-cell-input {
          font-family: 'IBM Plex Mono', monospace; font-variant-numeric: tabular-nums;
          width: 84px; text-align: right; border: 1px solid transparent; border-radius: 4px;
          background: rgba(255,255,255,0.7); padding: 3px 5px; font-size: 12.5px;
          color: ${COLORS.textDark}; color-scheme: light;
        }
        .lw-cell-input:focus { outline: none; border-color: ${COLORS.jade}; background: #fff; }
        .lw-cell-input::-webkit-outer-spin-button,
        .lw-cell-input::-webkit-inner-spin-button { -webkit-appearance: none; margin: 0; }
        .lw-cell-input { -moz-appearance: textfield; }
        .lw-row:hover td { filter: brightness(0.98); }
        select.lw-select { font-family: 'IBM Plex Sans', sans-serif; }
        .hier-row:hover { background: ${COLORS.surfaceAlt}; }
        .hier-row.selected { background: ${COLORS.jadeSoft}; }
        .hier-actions { opacity: 0; }
        .hier-row:hover .hier-actions { opacity: 1; }
        .hier-drag-handle { opacity: 0; }
        .hier-row:hover .hier-drag-handle { opacity: 0.5; }
        .hier-row.drag-over { outline: 2px dashed ${COLORS.jade}; outline-offset: -2px; }
      `}</style>

      {/* ---------------- Top chrome ---------------- */}
      <div style={{ background: COLORS.bgChrome, borderBottom: `1px solid ${COLORS.chromeBorder}` }} className="px-6 py-4">
        <div className="flex items-center justify-between flex-wrap gap-3">
          <div>
            <div style={{ fontFamily: "'Space Grotesk', sans-serif", color: COLORS.textOnDark }} className="text-xl font-bold tracking-tight">
              Rosebud
            </div>
            <div style={{ color: COLORS.textOnDarkMuted }} className="text-xs">
              Driver-based budget modeling
            </div>
          </div>

          <div className="flex items-center gap-3 flex-wrap">
            <div style={{ color: COLORS.textOnDarkMuted }} className="text-xs flex items-center gap-2">
              {workspaceName}
              <span style={{ background: COLORS.bgChrome2, border: `1px solid ${COLORS.chromeBorder}`, borderRadius: 4, padding: '1px 6px', fontSize: 11, textTransform: 'capitalize' }}>{role}</span>
            </div>
            {myWorkspaces.length > 1 && (
              <select
                className="lw-select"
                value=""
                onChange={(e) => { if (e.target.value) switchWorkspace(Number(e.target.value)); }}
                style={{ background: COLORS.bgChrome2, color: COLORS.textOnDark, border: `1px solid ${COLORS.chromeBorder}`, borderRadius: 6, padding: '5px 8px', fontSize: 12.5 }}
              >
                <option value="">Switch workspace…</option>
                {myWorkspaces.map((w) => <option key={w.id} value={w.id}>{w.name} ({w.role})</option>)}
              </select>
            )}
            <button
              onClick={() => navigate('/')}
              style={{
                display: 'flex', alignItems: 'center', gap: 6, padding: '5px 10px', borderRadius: 6, fontSize: 12.5,
                background: location.pathname === '/' ? COLORS.violet : 'none', color: location.pathname === '/' ? '#fff' : COLORS.textOnDarkMuted,
                border: `1px solid ${location.pathname === '/' ? COLORS.violet : COLORS.chromeBorder}`, cursor: 'pointer',
              }}
            >
              Workspace
            </button>
            <button
              onClick={() => { if (showMembers) navigate('/'); else { navigate('/members'); loadMembers(); } }}
              style={{
                display: 'flex', alignItems: 'center', gap: 6, padding: '5px 10px', borderRadius: 6, fontSize: 12.5,
                background: showMembers ? COLORS.violet : 'none', color: showMembers ? '#fff' : COLORS.textOnDarkMuted,
                border: `1px solid ${showMembers ? COLORS.violet : COLORS.chromeBorder}`, cursor: 'pointer',
              }}
            >
              Members
            </button>
            <button
              onClick={() => { if (showActivity) navigate('/'); else { navigate('/activity'); loadActivity(); } }}
              style={{
                display: 'flex', alignItems: 'center', gap: 6, padding: '5px 10px', borderRadius: 6, fontSize: 12.5,
                background: showActivity ? COLORS.violet : 'none', color: showActivity ? '#fff' : COLORS.textOnDarkMuted,
                border: `1px solid ${showActivity ? COLORS.violet : COLORS.chromeBorder}`, cursor: 'pointer',
              }}
            >
              <ActivityIcon size={13} /> Activity
            </button>
            {canManageBackups && (
              <button
                onClick={() => { if (showBackups) navigate('/'); else { navigate('/backups'); loadBackups(); } }}
                style={{
                  display: 'flex', alignItems: 'center', gap: 6, padding: '5px 10px', borderRadius: 6, fontSize: 12.5,
                  background: showBackups ? COLORS.violet : 'none', color: showBackups ? '#fff' : COLORS.textOnDarkMuted,
                  border: `1px solid ${showBackups ? COLORS.violet : COLORS.chromeBorder}`, cursor: 'pointer',
                }}
              >
                <Archive size={13} /> Backups {backups.length > 0 ? `(${backups.length})` : ''}
              </button>
            )}
            {canManageHierarchy && (
              <button
                onClick={() => navigate(showHierarchy ? '/' : '/hierarchy')}
                style={{
                  display: 'flex', alignItems: 'center', gap: 6, padding: '5px 10px', borderRadius: 6, fontSize: 12.5,
                  background: showHierarchy ? COLORS.violet : 'none', color: showHierarchy ? '#fff' : COLORS.textOnDarkMuted,
                  border: `1px solid ${showHierarchy ? COLORS.violet : COLORS.chromeBorder}`, cursor: 'pointer',
                }}
              >
                <Network size={13} /> Hierarchy
              </button>
            )}
            <div className="flex items-center gap-1.5 text-xs" style={{ color: saveStatus === 'error' ? COLORS.brick : COLORS.textOnDarkMuted }}>
              <span style={{
                width: 6, height: 6, borderRadius: 6, display: 'inline-block',
                background: saveStatus === 'error' ? COLORS.brick : saveStatus === 'saving' ? COLORS.amber : COLORS.jade,
              }} />
              {saveStatus === 'saving' ? 'Saving…' : saveStatus === 'error' ? 'Save failed — check server' : 'Saved'}
            </div>
            <button
              onClick={onLogout}
              style={{ background: 'none', border: `1px solid ${COLORS.chromeBorder}`, color: COLORS.textOnDarkMuted, borderRadius: 6, padding: '5px 10px', fontSize: 12.5, cursor: 'pointer' }}
            >
              Log out
            </button>

            <div style={{ background: COLORS.bgChrome2, border: `1px solid ${COLORS.chromeBorder}`, borderRadius: 6, padding: 2 }} className="flex">
              {['Monthly', 'Quarterly', 'Annual'].map((g) => (
                <button
                  key={g}
                  onClick={() => setGranularity(g)}
                  style={{
                    padding: '5px 10px', fontSize: 12.5, borderRadius: 4, border: 'none',
                    background: granularity === g ? COLORS.jade : 'transparent',
                    color: granularity === g ? '#fff' : COLORS.textOnDarkMuted,
                    cursor: 'pointer',
                  }}
                >
                  {g}
                </button>
              ))}
            </div>

            {isDefaultConfig && (
              <button
                onClick={() => setCompareMode((c) => !c)}
                style={{
                  display: 'flex', alignItems: 'center', gap: 6, padding: '6px 10px', borderRadius: 6, fontSize: 13,
                  background: compareMode ? COLORS.violet : COLORS.bgChrome2, color: '#fff', border: `1px solid ${compareMode ? COLORS.violet : COLORS.chromeBorder}`, cursor: 'pointer',
                }}
              >
                <ArrowLeftRight size={14} /> Compare
              </button>
            )}

            <button
              onClick={() => navigate(showVersions ? '/' : '/versions')}
              style={{
                display: 'flex', alignItems: 'center', gap: 6, padding: '6px 10px', borderRadius: 6, fontSize: 13,
                background: showVersions ? COLORS.amber : COLORS.bgChrome2, color: '#fff', border: `1px solid ${showVersions ? COLORS.amber : COLORS.chromeBorder}`, cursor: 'pointer',
              }}
            >
              <History size={14} /> Versions {versions.length > 0 ? `(${versions.length})` : ''}
            </button>
          </div>
        </div>

        {compareMode && (
          <div className="flex items-center gap-3 flex-wrap mt-3 text-xs" style={{ color: COLORS.textOnDarkMuted }}>
            <span>Compare {currentScenario} against</span>
            <select
              className="lw-select"
              value={compareTarget}
              onChange={(e) => setCompareTarget(e.target.value)}
              style={{ background: COLORS.bgChrome2, color: COLORS.textOnDark, border: `1px solid ${COLORS.chromeBorder}`, borderRadius: 6, padding: '5px 8px', fontSize: 12.5 }}
            >
              {compareOptions.map((o) => <option key={o.key} value={o.key}>{o.label}</option>)}
            </select>
            <span>for period</span>
            <select
              className="lw-select"
              value={comparePeriod}
              onChange={(e) => setComparePeriod(e.target.value)}
              style={{ background: COLORS.bgChrome2, color: COLORS.textOnDark, border: `1px solid ${COLORS.chromeBorder}`, borderRadius: 6, padding: '5px 8px', fontSize: 12.5 }}
            >
              {[...MONTHS, 'Q1', 'Q2', 'Q3', 'Q4', 'FY'].map((p) => <option key={p} value={p}>{p}</option>)}
            </select>
            <button onClick={() => setCompareMode(false)} style={{ color: COLORS.textOnDarkMuted, background: 'none', border: 'none', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 4 }}>
              <X size={13} /> close
            </button>
          </div>
        )}
      </div>

      {location.pathname === '/' && (
      <>
      {/* ---------------- Pivot configuration ---------------- */}
      <div className="px-6 pt-4">
        <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 12 }} className="flex items-stretch gap-3 flex-wrap">
          {['rows', 'cols', 'filter'].map((zone) => {
            const order = zone === 'rows' ? rowsOrder : zone === 'cols' ? colsOrder : ALL_PIVOT_DIMS.filter((d) => !rowsOrder.includes(d) && !colsOrder.includes(d));
            const setOrder = zone === 'rows' ? setRowsOrder : zone === 'cols' ? setColsOrder : null;
            const isAxis = zone === 'rows' || zone === 'cols';
            return (
              <div
                key={zone}
                onDragOver={(e) => e.preventDefault()}
                onDrop={() => { if (dragChip) moveDimToZone(dragChip, zone); setDragChip(null); }}
                style={{
                  flex: zone === 'filter' ? 2 : 1, minWidth: zone === 'filter' ? 230 : 140,
                  border: `1.5px dashed ${COLORS.border}`, borderRadius: 8, padding: '8px 10px',
                  display: 'flex', flexDirection: 'column', gap: 6,
                }}
              >
                <div style={{ fontSize: 9.5, fontWeight: 700, letterSpacing: '.06em', color: COLORS.textMuted, textTransform: 'uppercase' }}>
                  {zone === 'rows' ? 'Rows (outer → inner)' : zone === 'cols' ? 'Columns (outer → inner)' : 'Filters'}
                </div>
                <div className="flex items-center gap-2 flex-wrap">
                  {order.map((dim, i) => (
                    <div
                      key={dim}
                      draggable
                      onDragStart={() => setDragChip(dim)}
                      onDoubleClick={(e) => { e.stopPropagation(); openSubsetEditor(dim); }}
                      title="Double-click to choose which members show"
                      style={{
                        display: 'flex', alignItems: 'center', gap: 3, background: COLORS.jadeSoft, color: COLORS.jade,
                        borderRadius: 999, padding: '5px 8px 5px 10px', fontSize: 12.5, fontWeight: 600, cursor: 'grab', userSelect: 'none',
                      }}
                    >
                      <GripVertical size={11} />
                      <span>{PIVOT_DIMENSIONS[dim].label}</span>
                      {pivotSubsets[dim] && (
                        <span style={{ fontSize: 10, fontWeight: 700, background: COLORS.jade, color: '#fff', borderRadius: 999, padding: '1px 5px' }}>
                          {pivotSubsets[dim].size}/{subsetAllMembers(dim).length}
                        </span>
                      )}
                      {zone === 'filter' && renderFilterValuePicker(dim)}
                      {isAxis && order.length > 1 && (
                        <span className="flex items-center" style={{ marginLeft: 2 }}>
                          <button
                            onClick={(e) => { e.stopPropagation(); reorderDim(dim, -1, order, setOrder); }}
                            disabled={i === 0}
                            title="Nest earlier (more outer)"
                            style={{ background: 'none', border: 'none', color: COLORS.jade, cursor: i === 0 ? 'default' : 'pointer', opacity: i === 0 ? 0.3 : 1, padding: '0 1px', fontSize: 11, lineHeight: 1 }}
                          >
                            ‹
                          </button>
                          <button
                            onClick={(e) => { e.stopPropagation(); reorderDim(dim, 1, order, setOrder); }}
                            disabled={i === order.length - 1}
                            title="Nest later (more inner)"
                            style={{ background: 'none', border: 'none', color: COLORS.jade, cursor: i === order.length - 1 ? 'default' : 'pointer', opacity: i === order.length - 1 ? 0.3 : 1, padding: '0 1px', fontSize: 11, lineHeight: 1 }}
                          >
                            ›
                          </button>
                        </span>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            );
          })}
        </div>
        <div style={{ color: COLORS.textMuted }} className="text-xs mt-1.5">
          Drag dimensions between Rows, Columns, and Filters — drop more than one onto the same axis to nest them, and use ‹ › to reorder. Only the default Account × Time arrangement (nothing nested) supports full editing and hierarchy drill-down; anything else is read-only.
        </div>
      </div>

      {/* ---------------- KPI strip ---------------- */}
      <div className="px-6 pt-5 grid grid-cols-3 gap-4">
        {[
          { label: `Revenue · FY (${currentScenario})`, value: fmtMoney(kpiRevenue), color: COLORS.jade },
          { label: `Expenses · FY (${currentScenario})`, value: fmtMoney(kpiExpenses), color: COLORS.textDark },
          { label: `Net Income · FY · ${kpiMargin.toFixed(1)}% margin`, value: fmtMoney(kpiNet), color: kpiNet >= 0 ? COLORS.jade : COLORS.brick },
        ].map((k) => (
          <div key={k.label} style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}` }} className="rounded-lg px-4 py-3">
            <div style={{ color: COLORS.textMuted }} className="text-xs mb-1">{k.label}</div>
            <div className="lw-num text-2xl" style={{ fontFamily: "'Space Grotesk', sans-serif", color: k.color }}>{k.value}</div>
          </div>
        ))}
      </div>

      {/* ---------------- Trend chart ---------------- */}
      <div className="px-6 pt-4">
        <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}` }} className="rounded-lg px-4 py-3">
          <div style={{ color: COLORS.textMuted }} className="text-xs mb-2">Monthly trend — Company (Consolidated), {currentScenario}</div>
          <div style={{ height: 200 }}>
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={chartData} margin={{ top: 5, right: 12, left: -12, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} />
                <XAxis dataKey="month" tick={{ fontSize: 11, fill: COLORS.textMuted }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
                <YAxis tick={{ fontSize: 11, fill: COLORS.textMuted }} axisLine={false} tickLine={false} tickFormatter={(v) => `${Math.round(v / 1000)}k`} />
                <Tooltip formatter={(v) => fmtMoney(v)} contentStyle={{ fontSize: 12, borderRadius: 8, border: `1px solid ${COLORS.border}` }} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                <Line type="monotone" dataKey="Revenue" stroke={COLORS.jade} strokeWidth={2} dot={false} />
                <Line type="monotone" dataKey="Expenses" stroke={COLORS.textMuted} strokeWidth={2} dot={false} />
                <Line type="monotone" dataKey="Net Income" stroke={COLORS.violet} strokeWidth={2} dot={false} />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      {/* ---------------- Grid ---------------- */}
      <div className="px-6 pt-4 pb-2" style={{ width: '100%', maxWidth: '100%', boxSizing: 'border-box' }}>
        {isDefaultConfig && !compareMode && (!canEdit || currentProduct === 'all') && (
          <div style={{ color: COLORS.textMuted }} className="text-xs mb-2">
            {currentEntity === 'company' && 'Select a specific entity to edit inputs. '}
            {granularity !== 'Monthly' && 'Switch to Monthly view to edit inputs. '}
            {currentProduct === 'all' && 'Select a specific product to edit Units Sold or Unit Price.'}
          </div>
        )}
        {isDefaultConfig && (
        <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 8, overflowX: 'auto', width: '100%', maxWidth: '100%' }}>
            <table style={{ minWidth: compareMode ? 640 : columns.length * 92 + 220, borderCollapse: 'collapse' }} className="text-sm">
              <thead>
                {!compareMode ? (
                  <tr style={{ background: COLORS.bgChrome }}>
                    <th style={{ position: 'sticky', left: 0, background: COLORS.bgChrome, color: COLORS.textOnDark, zIndex: 2 }} className="text-left px-3 py-2 text-xs font-medium whitespace-nowrap">
                      Account
                    </th>
                    {columns.map((c) => (
                      <th
                        key={c}
                        style={{ color: c === 'FY' ? '#fff' : COLORS.textOnDarkMuted, background: c === 'FY' ? '#232B3D' : 'transparent' }}
                        className="text-right px-3 py-2 text-xs font-medium whitespace-nowrap"
                      >
                        {c}
                      </th>
                    ))}
                  </tr>
                ) : (
                  <tr style={{ background: COLORS.bgChrome }}>
                    <th style={{ position: 'sticky', left: 0, background: COLORS.bgChrome, color: COLORS.textOnDark, zIndex: 2 }} className="text-left px-3 py-2 text-xs font-medium whitespace-nowrap">
                      Account
                    </th>
                    <th className="text-right px-3 py-2 text-xs font-medium whitespace-nowrap" style={{ color: COLORS.textOnDarkMuted }}>{currentScenario}</th>
                    <th className="text-right px-3 py-2 text-xs font-medium whitespace-nowrap" style={{ color: COLORS.textOnDarkMuted }}>{targetLabel(compareTarget)}</th>
                    <th className="text-right px-3 py-2 text-xs font-medium whitespace-nowrap" style={{ color: COLORS.textOnDarkMuted }}>Δ $</th>
                    <th className="text-right px-3 py-2 text-xs font-medium whitespace-nowrap" style={{ color: COLORS.textOnDarkMuted }}>Δ %</th>
                  </tr>
                )}
              </thead>

              <tbody>
                {visibleRows.map((row) => {
                  const bg = rowBg(row);
                  const nameStyle = rowNameStyle(row);
                  const kids = hasChildren(row.id);

                  if (!compareMode) {
                    return (
                      <tr key={row.id} className="lw-row" style={{ borderTop: `1px solid ${COLORS.border}` }}>
                        <td
                          style={{ position: 'sticky', left: 0, background: bg, borderLeft: `3px solid ${railColor(row)}`, zIndex: 1 }}
                          className="px-3 py-1.5 whitespace-nowrap"
                        >
                          <div style={{ paddingLeft: row.depth * 16, display: 'flex', alignItems: 'center', gap: 4, ...nameStyle }}>
                            {kids ? (
                              <button onClick={() => toggleExpand(row.id)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: COLORS.textMuted, display: 'flex' }}>
                                {expanded.has(row.id) ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
                              </button>
                            ) : <span style={{ width: 13, display: 'inline-block' }} />}
                            <span style={{ fontSize: 13 }}>{row.name}</span>
                            {row.isDriver && <span style={{ width: 5, height: 5, borderRadius: 5, background: COLORS.amber, display: 'inline-block', marginLeft: 2 }} />}
                          </div>
                        </td>
                        {columns.map((col) => {
                          const val = getPeriodValue(liveData, currentEntity, row.id, col, currentProduct);
                          const isProductAccount = PRODUCT_DIMENSIONED_ACCOUNTS.includes(row.id);
                          const productKey = isProductAccount ? currentProduct : 'none';
                          const editableHere = canEdit && row.type === 'input' && MONTHS.includes(col) && (!isProductAccount || currentProduct !== 'all');
                          return (
                            <td key={col} style={{ background: col === 'FY' ? '#EDEFF3' : bg }} className="px-3 py-1.5 text-right">
                              {editableHere ? (
                                <input
                                  type="number"
                                  className="lw-num lw-cell-input"
                                  value={liveData[currentEntity]?.[row.id]?.[productKey]?.[col] ?? 0}
                                  onChange={(e) => updateValue(currentEntity, row.id, col, e.target.value)}
                                />
                              ) : (
                                <span className="lw-num" style={{ fontSize: 12.5, fontWeight: col === 'FY' ? 600 : 400 }}>{fmtCell(row.unit, val)}</span>
                              )}
                            </td>
                          );
                        })}
                      </tr>
                    );
                  }

                  // Compare mode row
                  const valA = getPeriodValue(liveData, currentEntity, row.id, comparePeriod, currentProduct);
                  const valB = getPeriodValue(compareData, currentEntity, row.id, comparePeriod, currentProduct);
                  const diff = valA - valB;
                  const pct = valB !== 0 ? (diff / Math.abs(valB)) * 100 : (valA === 0 ? 0 : 100);
                  const favorable = topAncestor(row.id) === 'expenses' ? diff <= 0 : diff >= 0;
                  const diffColor = diff === 0 ? COLORS.textMuted : (favorable ? COLORS.jade : COLORS.brick);

                  return (
                    <tr key={row.id} className="lw-row" style={{ borderTop: `1px solid ${COLORS.border}` }}>
                      <td style={{ position: 'sticky', left: 0, background: bg, borderLeft: `3px solid ${railColor(row)}`, zIndex: 1 }} className="px-3 py-1.5 whitespace-nowrap">
                        <div style={{ paddingLeft: row.depth * 16, display: 'flex', alignItems: 'center', gap: 4, ...nameStyle }}>
                          {kids ? (
                            <button onClick={() => toggleExpand(row.id)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: COLORS.textMuted, display: 'flex' }}>
                              {expanded.has(row.id) ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
                            </button>
                          ) : <span style={{ width: 13, display: 'inline-block' }} />}
                          <span style={{ fontSize: 13 }}>{row.name}</span>
                        </div>
                      </td>
                      <td className="px-3 py-1.5 text-right lw-num" style={{ fontSize: 12.5, background: bg }}>{fmtCell(row.unit, valA)}</td>
                      <td className="px-3 py-1.5 text-right lw-num" style={{ fontSize: 12.5, background: bg }}>{fmtCell(row.unit, valB)}</td>
                      <td className="px-3 py-1.5 text-right lw-num" style={{ fontSize: 12.5, color: diffColor, background: bg }}>{fmtCell(row.unit, diff)}</td>
                      <td className="px-3 py-1.5 text-right lw-num" style={{ fontSize: 12.5, color: diffColor, background: bg }}>{fmtPct(pct)}</td>
                    </tr>
                  );
                })}

                {/* Net income summary row */}
                {(() => {
                  if (!compareMode) {
                    return (
                      <tr style={{ borderTop: `2px solid ${COLORS.textDark}` }}>
                        <td style={{ position: 'sticky', left: 0, background: '#1C2333', color: '#fff', zIndex: 1 }} className="px-3 py-2 font-bold whitespace-nowrap text-sm">
                          Net Income
                        </td>
                        {columns.map((col) => {
                          const rev = getPeriodValue(liveData, currentEntity, 'revenue', col);
                          const exp = getPeriodValue(liveData, currentEntity, 'expenses', col);
                          const net = rev - exp;
                          return (
                            <td key={col} className="px-3 py-2 text-right lw-num font-semibold" style={{ fontSize: 12.5, background: net >= 0 ? COLORS.jade : COLORS.brick, color: '#fff' }}>
                              {fmtMoney(net)}
                            </td>
                          );
                        })}
                      </tr>
                    );
                  }
                  const revA = getPeriodValue(liveData, currentEntity, 'revenue', comparePeriod);
                  const expA = getPeriodValue(liveData, currentEntity, 'expenses', comparePeriod);
                  const netA = revA - expA;
                  const revB = getPeriodValue(compareData, currentEntity, 'revenue', comparePeriod);
                  const expB = getPeriodValue(compareData, currentEntity, 'expenses', comparePeriod);
                  const netB = revB - expB;
                  const diff = netA - netB;
                  const pct = netB !== 0 ? (diff / Math.abs(netB)) * 100 : (netA === 0 ? 0 : 100);
                  const diffColor = diff === 0 ? '#DADFE8' : (diff >= 0 ? '#8FE3C4' : '#F4B7AC');
                  return (
                    <tr style={{ borderTop: `2px solid ${COLORS.textDark}` }}>
                      <td style={{ position: 'sticky', left: 0, background: '#1C2333', color: '#fff', zIndex: 1 }} className="px-3 py-2 font-bold whitespace-nowrap text-sm">Net Income</td>
                      <td style={{ background: '#1C2333', color: '#fff' }} className="px-3 py-2 text-right lw-num font-semibold" >{fmtMoney(netA)}</td>
                      <td style={{ background: '#1C2333', color: '#fff' }} className="px-3 py-2 text-right lw-num font-semibold">{fmtMoney(netB)}</td>
                      <td style={{ background: '#1C2333', color: diffColor }} className="px-3 py-2 text-right lw-num font-semibold">{fmtMoney(diff)}</td>
                      <td style={{ background: '#1C2333', color: diffColor }} className="px-3 py-2 text-right lw-num font-semibold">{fmtPct(pct)}</td>
                    </tr>
                  );
                })()}
              </tbody>
            </table>
        </div>
        )}
        {isDefaultConfig && (
        <div style={{ color: COLORS.textMuted }} className="text-xs mt-2">
          Rail colors: <span style={{ color: COLORS.jade }}>■</span> rollup &nbsp; <span style={{ color: COLORS.violet }}>■</span> formula &nbsp; <span style={{ color: COLORS.amber }}>■</span> driver input &nbsp; <span style={{ color: COLORS.border }}>■</span> input. Edits save to the shared workspace on the budget server as you type.
        </div>
        )}

        {!isDefaultConfig && (() => {
          const DIM_KEY = { entity: 'entity', account: 'account', time: 'period', product: 'product', scenario: 'scenario' };
          function filterValueFor(dim) {
            if (dim === 'entity') return currentEntity;
            if (dim === 'product') return currentProduct;
            if (dim === 'account') return filterAccount;
            if (dim === 'scenario') return currentScenario;
            if (dim === 'time') return filterTime;
            return null;
          }
          function resolveArgs(rowPath, colPath) {
            const combined = { ...rowPath, ...colPath };
            const args = {};
            Object.keys(DIM_KEY).forEach((dim) => {
              const key = DIM_KEY[dim];
              args[key] = dim in combined ? combined[dim] : filterValueFor(dim);
            });
            return args;
          }
          function valueAt(rowPath, colPath) {
            const a = resolveArgs(rowPath, colPath);
            return getPivotCellValue(values, a.scenario, a.entity, a.account, a.period, a.product);
          }
          function unitFor(rowPath, colPath) {
            const combined = { ...rowPath, ...colPath };
            if ('account' in combined) return ACCOUNTS_BY_ID[combined.account]?.unit || '$';
            return ACCOUNTS_BY_ID[filterAccount]?.unit || '$';
          }
          const bodyFont = { fontFamily: "'IBM Plex Sans', sans-serif" };

          // Both axes prune through the SAME collapsed-path set — collapsing
          // a node (row or column) strips its children for colspan/leaf
          // purposes while remembering it had them, so the header keeps a
          // clickable chevron and the cell shows a real rolled-up value
          // (via rollupPathFor) instead of going blank.
          const rawColTree = buildPivotTree(colsOrder, granularity, {}, pivotSubsets);
          const colTree = pruneCollapsedPivotTree(rawColTree, pivotCollapsed);
          const colHeaderLevels = pivotColumnHeaderRows(colTree);
          const colLeaves = flattenPivotLeaves(colTree);
          const rowTree = buildPivotTree(rowsOrder, granularity, {}, pivotSubsets);

          function cellValue(rowPath, rowHadChildren, colPath, colHadChildren) {
            const effRow = rowHadChildren ? rollupPathFor(rowPath, rowsOrder) : rowPath;
            const effCol = colHadChildren ? rollupPathFor(colPath, colsOrder) : colPath;
            if (!effRow || !effCol) return null;
            return valueAt(effRow, effCol);
          }

          function renderPivotRowNode(node, depth) {
            const key = JSON.stringify(node.path);
            const isExpanded = !pivotCollapsed.has(key);
            const isTotalStyle = node.isTotal;
            return (
              <React.Fragment key={key}>
                <tr
                  className={node.children ? undefined : 'lw-row'}
                  style={{ borderTop: `1px solid ${COLORS.border}`, background: node.children ? COLORS.surfaceAlt : COLORS.surface, cursor: node.children ? 'pointer' : undefined }}
                  onClick={node.children ? () => togglePivotExpand(node.path) : undefined}
                >
                  <td
                    style={{
                      position: 'sticky', left: 0, zIndex: 1, textAlign: 'left', ...bodyFont,
                      background: isTotalStyle ? '#1C2333' : (node.children ? COLORS.surfaceAlt : COLORS.surface),
                      color: isTotalStyle ? '#fff' : undefined,
                    }}
                    className="px-3 py-1.5 whitespace-nowrap text-sm"
                  >
                    <div style={{ paddingLeft: depth * 18, display: 'flex', alignItems: 'center', gap: 6, fontWeight: node.children ? 700 : (isTotalStyle ? 700 : 400) }}>
                      {node.children ? (isExpanded ? <ChevronDown size={13} /> : <ChevronRight size={13} />) : <span style={{ width: 13, flexShrink: 0 }} />}
                      {node.name}
                    </div>
                  </td>
                  {colLeaves.map((c) => {
                    const val = cellValue(node.path, !!node.children, c.path, c.hadChildren);
                    const cellBg = isTotalStyle ? COLORS.violet : (node.children ? COLORS.surfaceAlt : (c.isTotal || c.hadChildren ? COLORS.surfaceAlt : COLORS.surface));
                    return (
                      <td key={JSON.stringify(c.path)} className="px-3 py-1.5 text-right" style={{ background: cellBg, color: isTotalStyle ? '#fff' : undefined }}>
                        {val === null ? (
                          <span style={{ color: COLORS.textMuted, fontSize: 12.5 }}>—</span>
                        ) : (
                          <span className="lw-num" style={{ fontSize: 12.5, fontWeight: (isTotalStyle || c.isTotal || node.children || c.hadChildren) ? 600 : 400 }}>{fmtCell(unitFor(node.path, c.path), val)}</span>
                        )}
                      </td>
                    );
                  })}
                </tr>
                {node.children && isExpanded && node.children.map((child) => renderPivotRowNode(child, depth + 1))}
              </React.Fragment>
            );
          }

          return (
            <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 8, width: '100%', maxWidth: '100%', overflow: 'hidden' }}>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ minWidth: colLeaves.length * 100 + 220, borderCollapse: 'collapse', ...bodyFont }} className="text-sm">
                  <thead>
                    {colHeaderLevels.map((levelCells, levelIdx) => (
                      <tr key={levelIdx} style={{ background: COLORS.bgChrome }}>
                        {levelIdx === 0 && (
                          <th
                            rowSpan={colHeaderLevels.length}
                            style={{ position: 'sticky', left: 0, background: COLORS.bgChrome, color: COLORS.textOnDark, zIndex: 3, ...bodyFont }}
                            className="text-left px-3 py-2 text-xs font-medium whitespace-nowrap"
                          >
                            {rowsOrder.map((d) => PIVOT_DIMENSIONS[d].label).join(' › ')}
                          </th>
                        )}
                        {levelCells.map((cell, i) => (
                          <th
                            key={i}
                            colSpan={cell.colSpan}
                            onClick={cell.hadChildren ? () => togglePivotExpand(cell.path) : undefined}
                            style={{ color: cell.isTotal ? '#fff' : COLORS.textOnDarkMuted, background: cell.isTotal ? '#232B3D' : 'transparent', cursor: cell.hadChildren ? 'pointer' : undefined, ...bodyFont }}
                            className="text-right px-3 py-2 text-xs font-medium whitespace-nowrap"
                          >
                            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                              {cell.name}
                              {cell.hadChildren && (pivotCollapsed.has(JSON.stringify(cell.path)) ? <ChevronRight size={12} /> : <ChevronDown size={12} />)}
                            </span>
                          </th>
                        ))}
                      </tr>
                    ))}
                  </thead>
                  <tbody>
                    {rowTree.map((node) => renderPivotRowNode(node, 0))}
                  </tbody>
                </table>
              </div>
              <div className="text-xs px-3 py-2 border-t" style={{ color: COLORS.textMuted, borderColor: COLORS.border }}>
                Read-only in this arrangement. Switch Rows to Account and Columns to Time (nothing nested) to edit values. Click any column header with an arrow to collapse it.
              </div>
            </div>
          );
        })()}
      </div>

      {/* ---------------- Subset editor modal ---------------- */}
      {subsetEditorDim && (() => {
        const members = subsetAllMembers(subsetEditorDim);
        const allSelected = members.every((m) => subsetEditorDraft.has(m.id));
        return (
          <div
            style={{ position: 'fixed', inset: 0, background: 'rgba(15,20,30,0.55)', zIndex: 50, display: 'flex', alignItems: 'center', justifyContent: 'center' }}
            onClick={closeSubsetEditor}
          >
            <div
              onClick={(e) => e.stopPropagation()}
              style={{ background: COLORS.surface, borderRadius: 12, width: 340, maxHeight: '70vh', display: 'flex', flexDirection: 'column', boxShadow: '0 12px 40px rgba(0,0,0,0.3)' }}
            >
              <div style={{ padding: '14px 18px', borderBottom: `1px solid ${COLORS.border}` }} className="flex items-center justify-between">
                <div>
                  <div style={{ fontSize: 14, fontWeight: 700 }}>{PIVOT_DIMENSIONS[subsetEditorDim].label} subset</div>
                  <div style={{ fontSize: 11.5, color: COLORS.textMuted }}>Choose which members to show</div>
                </div>
                <button onClick={closeSubsetEditor} style={{ background: 'none', border: 'none', color: COLORS.textMuted, cursor: 'pointer', display: 'flex', padding: 4 }}>
                  <X size={16} />
                </button>
              </div>
              <div style={{ padding: '8px 18px', display: 'flex', gap: 12, borderBottom: `1px solid ${COLORS.border}` }}>
                <button onClick={() => setSubsetEditorDraft(new Set(members.map((m) => m.id)))} style={{ background: 'none', border: 'none', color: COLORS.violet, fontSize: 12, fontWeight: 600, cursor: 'pointer', padding: 0 }}>
                  Select all
                </button>
                <button onClick={() => setSubsetEditorDraft(new Set())} style={{ background: 'none', border: 'none', color: COLORS.violet, fontSize: 12, fontWeight: 600, cursor: 'pointer', padding: 0 }}>
                  Select none
                </button>
              </div>
              <div style={{ overflowY: 'auto', padding: '8px 10px', flex: 1 }}>
                {members.map((m) => (
                  <label key={m.id} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 8px', borderRadius: 6, cursor: 'pointer', fontSize: 13 }}>
                    <input type="checkbox" checked={subsetEditorDraft.has(m.id)} onChange={() => toggleSubsetDraftMember(m.id)} />
                    {m.name}
                  </label>
                ))}
              </div>
              <div style={{ padding: '12px 18px', borderTop: `1px solid ${COLORS.border}`, display: 'flex', gap: 8 }}>
                <button
                  onClick={closeSubsetEditor}
                  style={{ flex: 1, background: 'none', border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: 8, fontSize: 13, fontWeight: 600, cursor: 'pointer' }}
                >
                  Cancel
                </button>
                <button
                  onClick={applySubsetEditor}
                  disabled={subsetEditorDraft.size === 0}
                  style={{
                    flex: 1, background: subsetEditorDraft.size === 0 ? COLORS.border : COLORS.jade, color: '#fff', border: 'none', borderRadius: 8,
                    padding: 8, fontSize: 13, fontWeight: 600, cursor: subsetEditorDraft.size === 0 ? 'default' : 'pointer',
                  }}
                >
                  Apply{allSelected ? ' (all)' : ` (${subsetEditorDraft.size})`}
                </button>
              </div>
            </div>
          </div>
        );
      })()}
      </>
      )}


      {/* ---------------- Hierarchy editor page ---------------- */}
      {showHierarchy && (
        canManageHierarchy ? (
          <HierarchyEditor token={token} />
        ) : (
          <div className="px-6 pb-6">
            <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}` }} className="rounded-lg p-4 text-sm" >
              You need Power or Admin access to open the hierarchy editor.
            </div>
          </div>
        )
      )}

      {/* ---------------- Activity page ---------------- */}
      {showActivity && (
        <div className="px-6 pt-4 pb-6">
          <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 10, padding: 20 }}>
            <div className="flex items-center justify-between mb-1">
              <div className="font-semibold text-sm">Workspace activity</div>
            </div>
            <div style={{ color: COLORS.textMuted }} className="text-xs mb-4">
              Invites, role changes, version and backup saves/restores, and hierarchy edits. Routine cell edits aren't logged here to keep this readable.
            </div>
            {activityLoading ? (
              <div style={{ color: COLORS.textMuted }} className="text-xs">Loading…</div>
            ) : activityEntries.length === 0 ? (
              <div style={{ color: COLORS.textMuted }} className="text-xs">No activity yet.</div>
            ) : (
              <div className="flex flex-col">
                {activityEntries.map((e) => (
                  <div key={e.id} style={{ borderBottom: `1px solid ${COLORS.border}`, padding: '9px 0' }} className="flex items-start justify-between gap-4">
                    <div className="text-sm">
                      <span style={{ fontWeight: 600 }}>{e.userEmail}</span>{' '}
                      <span style={{ color: COLORS.textMuted }}>{e.action}</span>
                    </div>
                    <div style={{ color: COLORS.textMuted, fontFamily: "'IBM Plex Mono', monospace", whiteSpace: 'nowrap', fontSize: 11.5 }}>
                      {new Date(e.createdAt).toLocaleString()}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      )}

      {/* ---------------- Versions drawer ---------------- */}
      {showVersions && (
        <div className="px-6 pb-6">
          <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}` }} className="rounded-lg p-4">
            <div className="flex items-center justify-between mb-3">
              <div className="font-semibold text-sm">Saved versions</div>
              <button onClick={() => navigate('/')} style={{ background: 'none', border: 'none', cursor: 'pointer', color: COLORS.textMuted }}><X size={16} /></button>
            </div>

            {canEditData && (
              <div className="flex items-center gap-2 mb-4">
                <input
                  type="text"
                  placeholder={`Label (e.g. "${currentScenario} — Q1 sign-off")`}
                  value={versionLabel}
                  onChange={(e) => setVersionLabel(e.target.value)}
                  style={{ flex: 1, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '7px 10px', fontSize: 13 }}
                />
                <button
                  onClick={saveVersion}
                  style={{ display: 'flex', alignItems: 'center', gap: 6, background: COLORS.jade, color: '#fff', border: 'none', borderRadius: 6, padding: '7px 12px', fontSize: 13, cursor: 'pointer' }}
                >
                  <Save size={14} /> Save {currentScenario}
                </button>
              </div>
            )}

            {versions.length === 0 ? (
              <div style={{ color: COLORS.textMuted }} className="text-xs">No versions saved yet. Snapshot a scenario to compare it later or roll back to it.</div>
            ) : (
              <div className="flex flex-col gap-2">
                {versions.slice().reverse().map((v) => (
                  <div key={v.id} style={{ border: `1px solid ${COLORS.border}`, borderRadius: 8 }} className="flex items-center justify-between px-3 py-2">
                    <div>
                      <div className="text-sm font-medium">{v.label}</div>
                      <div style={{ color: COLORS.textMuted }} className="text-xs">{v.scenario} · saved {v.timestamp}</div>
                    </div>
                    <div className="flex items-center gap-1">
                      <button onClick={() => compareWithVersion(v)} title="Compare against this version" style={{ background: COLORS.violetSoft, color: COLORS.violet, border: 'none', borderRadius: 6, padding: '5px 8px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 4, fontSize: 12 }}>
                        <ArrowLeftRight size={12} /> Compare
                      </button>
                      {canEditData && (
                        <>
                          <button onClick={() => restoreVersion(v)} title="Restore this snapshot" style={{ background: COLORS.jadeSoft, color: COLORS.jade, border: 'none', borderRadius: 6, padding: '5px 8px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 4, fontSize: 12 }}>
                            <RotateCcw size={12} /> Restore
                          </button>
                          <button onClick={() => deleteVersion(v.id)} title="Delete" style={{ background: COLORS.brickSoft, color: COLORS.brick, border: 'none', borderRadius: 6, padding: '5px 8px', cursor: 'pointer', display: 'flex', alignItems: 'center' }}>
                            <Trash2 size={12} />
                          </button>
                        </>
                      )}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      )}

      {/* ---------------- Members drawer ---------------- */}
      {showMembers && (
        <div className="px-6 pb-6">
          <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}` }} className="rounded-lg p-4">
            <div className="flex items-center justify-between mb-3">
              <div className="font-semibold text-sm">Workspace members</div>
              <div className="flex items-center gap-2">
                {canManageBackups && (
                  <button
                    onClick={downloadBackup}
                    style={{ display: 'flex', alignItems: 'center', gap: 6, background: COLORS.surfaceAlt, color: COLORS.textDark, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '5px 10px', fontSize: 12, cursor: 'pointer' }}
                  >
                    <Download size={13} /> Download Backup
                  </button>
                )}
                <button onClick={() => navigate('/')} style={{ background: 'none', border: 'none', cursor: 'pointer', color: COLORS.textMuted }}><X size={16} /></button>
              </div>
            </div>

            {membersError && <div style={{ background: COLORS.brickSoft, color: COLORS.brick, borderRadius: 6, padding: '8px 10px', fontSize: 12.5, marginBottom: 12 }}>{membersError}</div>}
            {lastInviteLink && (
              <div style={{ background: COLORS.jadeSoft, color: COLORS.jade, borderRadius: 6, padding: '8px 10px', fontSize: 12, marginBottom: 12, wordBreak: 'break-all' }}>
                Invite sent. If email isn't configured yet, share this link directly: {lastInviteLink}
              </div>
            )}

            {canManageMembers && (
              <div className="flex items-center gap-2 mb-4">
                <input
                  type="email"
                  placeholder="teammate@company.com"
                  value={inviteEmail}
                  onChange={(e) => setInviteEmail(e.target.value)}
                  style={{ flex: 1, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '7px 10px', fontSize: 13 }}
                />
                <select
                  value={inviteRole}
                  onChange={(e) => setInviteRole(e.target.value)}
                  style={{ border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '7px 8px', fontSize: 13 }}
                >
                  <option value="viewer">Viewer</option>
                  <option value="editor">Editor</option>
                  <option value="power">Power user</option>
                  <option value="admin">Admin</option>
                </select>
                <button
                  onClick={sendInvite}
                  disabled={!inviteEmail.trim()}
                  style={{ display: 'flex', alignItems: 'center', gap: 6, background: COLORS.jade, color: '#fff', border: 'none', borderRadius: 6, padding: '7px 12px', fontSize: 13, cursor: inviteEmail.trim() ? 'pointer' : 'default', opacity: inviteEmail.trim() ? 1 : 0.6 }}
                >
                  Invite
                </button>
              </div>
            )}

            <div className="flex flex-col gap-2">
              {members.map((m) => {
                const isSelf = m.userId === myUserId;
                return (
                  <div key={m.userId} style={{ border: `1px solid ${COLORS.border}`, borderRadius: 8 }} className="flex items-center justify-between px-3 py-2">
                    <div>
                      <div className="text-sm font-medium">{m.email}{isSelf && <span style={{ color: COLORS.textMuted, fontWeight: 400 }}> (you)</span>}</div>
                      <div style={{ color: COLORS.textMuted }} className="text-xs">joined {new Date(m.joinedAt).toLocaleDateString()}</div>
                    </div>
                    {canManageMembers && !isSelf ? (
                      <div className="flex items-center gap-2">
                        <select
                          value={m.role}
                          onChange={(e) => changeMemberRole(m.userId, e.target.value)}
                          style={{ border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '4px 6px', fontSize: 12 }}
                        >
                          <option value="viewer">Viewer</option>
                          <option value="editor">Editor</option>
                          <option value="power">Power user</option>
                          <option value="admin">Admin</option>
                        </select>
                        <button onClick={() => removeMember(m.userId)} title="Remove from workspace" style={{ background: COLORS.brickSoft, color: COLORS.brick, border: 'none', borderRadius: 6, padding: '5px 8px', cursor: 'pointer', display: 'flex', alignItems: 'center' }}>
                          <Trash2 size={12} />
                        </button>
                      </div>
                    ) : (
                      <span style={{ background: COLORS.surfaceAlt, borderRadius: 4, padding: '2px 8px', fontSize: 11.5, textTransform: 'capitalize', color: COLORS.textMuted }}>{m.role}{isSelf && canManageMembers ? ' · ask another admin to change' : ''}</span>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      )}

      {/* ---------------- Backups drawer ---------------- */}
      {showBackups && canManageBackups && (
        <div className="px-6 pb-6">
          <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}` }} className="rounded-lg p-4">
            <div className="flex items-center justify-between mb-3">
              <div className="font-semibold text-sm">Workspace backups</div>
              <button onClick={() => navigate('/')} style={{ background: 'none', border: 'none', cursor: 'pointer', color: COLORS.textMuted }}><X size={16} /></button>
            </div>
            <div style={{ color: COLORS.textMuted }} className="text-xs mb-4">
              A backup captures every scenario, every product/entity/month value, every saved version, and the member list at this moment. Stored here so anyone with Power or Admin access can come back and download it later.
            </div>

            {backupsError && <div style={{ background: COLORS.brickSoft, color: COLORS.brick, borderRadius: 6, padding: '8px 10px', fontSize: 12.5, marginBottom: 12 }}>{backupsError}</div>}

            <div className="flex items-center gap-2 mb-4">
              <input
                type="text"
                placeholder='Label (e.g. "End of Q2")'
                value={backupLabel}
                onChange={(e) => setBackupLabel(e.target.value)}
                style={{ flex: 1, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '7px 10px', fontSize: 13 }}
              />
              <button
                onClick={createBackup}
                disabled={creatingBackup}
                style={{ display: 'flex', alignItems: 'center', gap: 6, background: COLORS.jade, color: '#fff', border: 'none', borderRadius: 6, padding: '7px 12px', fontSize: 13, cursor: creatingBackup ? 'default' : 'pointer', opacity: creatingBackup ? 0.7 : 1 }}
              >
                <Archive size={14} /> {creatingBackup ? 'Creating…' : 'Create Backup'}
              </button>
            </div>

            {backups.length === 0 ? (
              <div style={{ color: COLORS.textMuted }} className="text-xs">No backups saved yet. Create one before making a big change.</div>
            ) : (
              <div className="flex flex-col gap-2">
                {backups.map((b) => (
                  <div key={b.id} style={{ border: `1px solid ${confirmingRestoreId === b.id ? COLORS.amber : COLORS.border}`, borderRadius: 8 }} className="flex items-center justify-between px-3 py-2">
                    <div>
                      <div className="text-sm font-medium">{b.label}</div>
                      <div style={{ color: confirmingRestoreId === b.id ? COLORS.amber : COLORS.textMuted }} className="text-xs">
                        {confirmingRestoreId === b.id ? 'This overwrites current Budget/Forecast/Actual data with this backup. Saved versions and members are untouched.' : `by ${b.createdBy} · ${new Date(b.createdAt).toLocaleString()}`}
                      </div>
                    </div>
                    <div className="flex items-center gap-1">
                      {confirmingRestoreId === b.id ? (
                        <>
                          <button onClick={() => restoreBackup(b.id)} style={{ background: COLORS.amber, color: '#fff', border: 'none', borderRadius: 6, padding: '5px 10px', cursor: 'pointer', fontSize: 12, fontWeight: 600 }}>
                            Yes, restore
                          </button>
                          <button onClick={() => setConfirmingRestoreId(null)} style={{ background: 'none', color: COLORS.textMuted, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '5px 10px', cursor: 'pointer', fontSize: 12 }}>
                            Cancel
                          </button>
                        </>
                      ) : (
                        <>
                          <button onClick={() => downloadStoredBackup(b.id, b.label)} title="Download this backup" style={{ background: COLORS.jadeSoft, color: COLORS.jade, border: 'none', borderRadius: 6, padding: '5px 8px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 4, fontSize: 12 }}>
                            <Download size={12} /> Download
                          </button>
                          <button onClick={() => restoreBackup(b.id)} title="Restore this backup into the live workspace" style={{ background: COLORS.amberSoft, color: COLORS.amber, border: 'none', borderRadius: 6, padding: '5px 8px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 4, fontSize: 12 }}>
                            <RotateCcw size={12} /> Restore
                          </button>
                          <button onClick={() => deleteBackup(b.id)} title="Delete this backup" style={{ background: COLORS.brickSoft, color: COLORS.brick, border: 'none', borderRadius: 6, padding: '5px 8px', cursor: 'pointer', display: 'flex', alignItems: 'center' }}>
                            <Trash2 size={12} />
                          </button>
                        </>
                      )}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

/* ----------------------------------------------------------------------
   HIERARCHY EDITOR — Power/Admin only. Wired to the real backend: the tree
   below reflects your actual Accounts/Entities/Products structure and
   every edit (rename, add, delete, move, attributes) is saved for real.
   Renaming/deleting here does NOT yet change how the P&L grid itself
   calculates — see the in-page note for what that next step would mean.
---------------------------------------------------------------------- */
function hierFindNode(node, id) {
  if (!node) return null;
  if (node.id === id) return node;
  if (node.children) for (const c of node.children) { const f = hierFindNode(c, id); if (f) return f; }
  return null;
}
function hierFindParent(node, targetId, parent) {
  if (!node) return null;
  if (node.id === targetId) return parent;
  if (node.children) for (const c of node.children) { const f = hierFindParent(c, targetId, node); if (f) return f; }
  return null;
}
function hierCountLeaves(node) {
  if (node.type === 'leaf') return 1;
  return (node.children || []).reduce((s, c) => s + hierCountLeaves(c), 0);
}
function hierMatchesFilter(node, term) {
  if (!term) return true;
  if (node.name.toLowerCase().includes(term)) return true;
  return (node.children || []).some((c) => hierMatchesFilter(c, term));
}
function hierCloneAndUpdate(node, targetId, updater) {
  if (node.id === targetId) return { ...node, ...updater(node) };
  if (node.children) return { ...node, children: node.children.map((c) => hierCloneAndUpdate(c, targetId, updater)) };
  return node;
}

function HierarchyEditor({ token }) {
  const [currentDim, setCurrentDim] = useState('Products');
  const [tree, setTree] = useState(null);
  const [attributeDefs, setAttributeDefs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState('');
  const [selectedId, setSelectedId] = useState(null);
  const [expandedIds, setExpandedIds] = useState(() => new Set());
  const [searchTerm, setSearchTerm] = useState('');
  const [dragId, setDragId] = useState(null);
  const [dragOverId, setDragOverId] = useState(null);
  const [confirmingDelete, setConfirmingDelete] = useState(false);
  const [newAttrName, setNewAttrName] = useState('');
  const [usageCount, setUsageCount] = useState(null);
  const [saveError, setSaveError] = useState('');

  function fetchHierarchy(dimension) {
    setLoading(true);
    setLoadError('');
    fetch(`${API_BASE}/workspace/hierarchy/${dimension}`, { headers: { Authorization: `Bearer ${token}` } })
      .then((r) => r.json().then((data) => ({ ok: r.ok, data })))
      .then(({ ok, data }) => {
        if (!ok) { setLoadError(data.error || 'Could not load this dimension.'); setLoading(false); return; }
        setTree(data.tree);
        setAttributeDefs(data.attributeDefs);
        setExpandedIds(new Set([data.tree.id, ...(data.tree.children || []).map((c) => c.id)]));
        setLoading(false);
      })
      .catch(() => { setLoadError('Could not load this dimension.'); setLoading(false); });
  }

  useEffect(() => {
    setSelectedId(null);
    setSearchTerm('');
    setConfirmingDelete(false);
    fetchHierarchy(currentDim);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [currentDim]);

  const selectedNode = tree && selectedId ? hierFindNode(tree, selectedId) : null;

  useEffect(() => {
    if (!selectedNode) { setUsageCount(null); return; }
    fetch(`${API_BASE}/workspace/hierarchy/${currentDim}/usage/${selectedNode.id}`, { headers: { Authorization: `Bearer ${token}` } })
      .then((r) => r.json())
      .then((data) => setUsageCount(typeof data.count === 'number' ? data.count : null))
      .catch(() => setUsageCount(null));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedId, currentDim]);

  function toggleExpand(id) {
    setExpandedIds((prev) => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n; });
  }
  function selectNode(id) { setSelectedId(id); setConfirmingDelete(false); }
  function switchDim(name) { setCurrentDim(name); }

  function renameNode(node, name) {
    setTree((prev) => hierCloneAndUpdate(prev, node.id, () => ({ name })));
    setSaveError('');
    fetch(`${API_BASE}/workspace/hierarchy/${currentDim}/nodes/${node.id}`, {
      method: 'PUT', headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify({ name }),
    }).catch(() => setSaveError('Could not save the rename.'));
  }
  function setAttrValue(node, key, value) {
    const nextAttrs = { ...(node.attrs || {}), [key]: value };
    setTree((prev) => hierCloneAndUpdate(prev, node.id, () => ({ attrs: nextAttrs })));
    setSaveError('');
    fetch(`${API_BASE}/workspace/hierarchy/${currentDim}/nodes/${node.id}`, {
      method: 'PUT', headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify({ attrs: nextAttrs }),
    }).catch(() => setSaveError('Could not save the attribute.'));
  }
  function addCategoryToRoot() {
    setSaveError('');
    fetch(`${API_BASE}/workspace/hierarchy/${currentDim}/nodes`, {
      method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify({ parentId: tree.id, name: 'New category', type: 'cat' }),
    })
      .then((r) => r.json().then((data) => ({ ok: r.ok, data })))
      .then(({ ok, data }) => {
        if (!ok) { setSaveError(data.error || 'Could not add category.'); return; }
        setExpandedIds((prev) => new Set(prev).add(tree.id));
        setSelectedId(data.id);
        fetchHierarchy(currentDim);
      })
      .catch(() => setSaveError('Could not add category.'));
  }
  function addLeafTo(catNode) {
    setSaveError('');
    fetch(`${API_BASE}/workspace/hierarchy/${currentDim}/nodes`, {
      method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify({ parentId: catNode.id, name: 'New item', type: 'leaf' }),
    })
      .then((r) => r.json().then((data) => ({ ok: r.ok, data })))
      .then(({ ok, data }) => {
        if (!ok) { setSaveError(data.error || 'Could not add item.'); return; }
        setExpandedIds((prev) => new Set(prev).add(catNode.id));
        setSelectedId(data.id);
        fetchHierarchy(currentDim);
      })
      .catch(() => setSaveError('Could not add item.'));
  }
  function deleteSelected() {
    if (!selectedNode || selectedNode.id === tree.id) return;
    setSaveError('');
    fetch(`${API_BASE}/workspace/hierarchy/${currentDim}/nodes/${selectedNode.id}`, {
      method: 'DELETE', headers: { Authorization: `Bearer ${token}` },
    })
      .then((r) => r.json().then((data) => ({ ok: r.ok, data })))
      .then(({ ok, data }) => {
        if (!ok) { setSaveError(data.error || 'Could not delete.'); return; }
        setSelectedId(null);
        setConfirmingDelete(false);
        fetchHierarchy(currentDim);
      })
      .catch(() => setSaveError('Could not delete.'));
  }
  function handleDrop(targetCat) {
    if (targetCat.type !== 'cat' || !dragId || dragId === targetCat.id) { setDragId(null); setDragOverId(null); return; }
    setSaveError('');
    fetch(`${API_BASE}/workspace/hierarchy/${currentDim}/nodes/${dragId}/move`, {
      method: 'PUT', headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify({ newParentId: targetCat.id }),
    })
      .then((r) => r.json().then((data) => ({ ok: r.ok, data })))
      .then(({ ok, data }) => {
        if (!ok) { setSaveError(data.error || 'Could not move item.'); setDragId(null); setDragOverId(null); return; }
        setExpandedIds((prev) => new Set(prev).add(targetCat.id));
        setSelectedId(dragId);
        setDragId(null);
        setDragOverId(null);
        fetchHierarchy(currentDim);
      })
      .catch(() => { setSaveError('Could not move item.'); setDragId(null); setDragOverId(null); });
  }
  function addAttributeDef() {
    const name = newAttrName.trim();
    if (!name || attributeDefs.includes(name)) return;
    setSaveError('');
    fetch(`${API_BASE}/workspace/hierarchy/${currentDim}/attributes`, {
      method: 'PUT', headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify({ action: 'add', name }),
    })
      .then((r) => r.json().then((data) => ({ ok: r.ok, data })))
      .then(({ ok, data }) => {
        if (!ok) { setSaveError(data.error || 'Could not add attribute.'); return; }
        setNewAttrName('');
        fetchHierarchy(currentDim);
      })
      .catch(() => setSaveError('Could not add attribute.'));
  }
  function removeAttributeDef(name) {
    setSaveError('');
    fetch(`${API_BASE}/workspace/hierarchy/${currentDim}/attributes`, {
      method: 'PUT', headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify({ action: 'remove', name }),
    })
      .then((r) => r.json().then((data) => ({ ok: r.ok, data })))
      .then(({ ok, data }) => { if (!ok) setSaveError(data.error || 'Could not remove attribute.'); else fetchHierarchy(currentDim); })
      .catch(() => setSaveError('Could not remove attribute.'));
  }

  function flattenForExport() {
    const rows = [];
    (function walk(node, path) {
      if (node.children && node.children.length) {
        node.children.forEach((c) => walk(c, node === tree ? path : [...path, node.name]));
      } else {
        rows.push({ path, name: node.name, attrs: node.attrs || {} });
      }
    })(tree, []);
    const maxDepth = rows.reduce((m, r) => Math.max(m, r.path.length), 0);
    const levelHeaders = maxDepth === 0 ? [] : maxDepth === 1 ? ['Category'] : Array.from({ length: maxDepth }, (_, i) => `Level ${i + 1}`);
    const headers = [...levelHeaders, 'Name', ...attributeDefs];
    const body = rows.map((r) => {
      const levelCols = Array.from({ length: maxDepth }, (_, i) => r.path[i] || '');
      return [...levelCols, r.name, ...attributeDefs.map((d) => r.attrs[d] || '')];
    });
    return [headers, ...body];
  }
  function exportToExcel() {
    const aoa = flattenForExport();
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = aoa[0].map((_, i) => ({ wch: Math.max(12, ...aoa.map((r) => String(r[i] ?? '').length)) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, currentDim.slice(0, 31));
    XLSX.writeFile(wb, `rosebud-${currentDim.toLowerCase()}-export.xlsx`);
  }

  function renderRow(node, depth, ancestorContinues, isLast) {
    const term = searchTerm.trim().toLowerCase();
    if (!hierMatchesFilter(node, term)) return null;
    const isCat = node.type === 'cat';
    const isOpen = expandedIds.has(node.id);
    const isSelected = node.id === selectedId;
    return (
      <div key={node.id}>
        <div
          className={`hier-row${isSelected ? ' selected' : ''}${dragOverId === node.id ? ' drag-over' : ''}`}
          draggable={!isCat}
          onDragStart={() => setDragId(node.id)}
          onDragOver={(e) => { if (isCat && dragId && dragId !== node.id) { e.preventDefault(); setDragOverId(node.id); } }}
          onDragLeave={() => setDragOverId((cur) => (cur === node.id ? null : cur))}
          onDrop={(e) => { e.preventDefault(); handleDrop(node); }}
          onClick={() => selectNode(node.id)}
          style={{ display: 'flex', alignItems: 'center', gap: 0, padding: '4px 8px', borderRadius: 6, cursor: 'pointer' }}
        >
          {ancestorContinues.map((cont, i) => (
            <span key={i} style={{ width: 16, textAlign: 'center', color: COLORS.border, flexShrink: 0, fontFamily: 'monospace', fontSize: 13, userSelect: 'none' }}>
              {cont ? '│' : ''}
            </span>
          ))}
          {depth > 0 && (
            <span style={{ width: 16, textAlign: 'center', color: COLORS.border, flexShrink: 0, fontFamily: 'monospace', fontSize: 13, userSelect: 'none' }}>
              {isLast ? '└' : '├'}
            </span>
          )}
          {node.children ? (
            <span
              onClick={(e) => { e.stopPropagation(); toggleExpand(node.id); }}
              style={{
                width: 14, height: 14, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0,
                border: `1px solid ${COLORS.border}`, borderRadius: 2, background: COLORS.surface, color: COLORS.textMuted,
                fontSize: 10, fontWeight: 700, lineHeight: 1, cursor: 'pointer', marginRight: 6,
              }}
            >
              {isOpen ? '−' : '+'}
            </span>
          ) : (
            <span style={{ width: 14, flexShrink: 0, marginRight: 6 }} />
          )}
          {isCat ? (
            <span style={{ fontSize: 12, fontWeight: 700, color: COLORS.jade, marginRight: 5, flexShrink: 0, fontFamily: "'IBM Plex Mono', monospace" }} title="Consolidated (rollup)">Σ</span>
          ) : (
            <span style={{ width: 8, height: 8, borderRadius: 2, background: COLORS.violet, marginRight: 5, flexShrink: 0 }} title="Leaf item" />
          )}
          {!isCat && <GripVertical size={11} className="hier-drag-handle" style={{ color: COLORS.textMuted, flexShrink: 0, marginRight: 3 }} />}
          <span style={{ fontSize: 13, fontWeight: isCat ? 700 : 400, letterSpacing: isCat ? '.01em' : 0 }}>{node.name}</span>
          <span className="hier-actions" style={{ display: 'flex', gap: 2, marginLeft: 6 }}>
            {isCat && (
              <button onClick={(e) => { e.stopPropagation(); addLeafTo(node); }} title="Add item" style={{ background: 'none', border: 'none', color: COLORS.textMuted, cursor: 'pointer', display: 'flex', padding: 3, borderRadius: 4 }}>
                <Plus size={13} />
              </button>
            )}
            <button onClick={(e) => { e.stopPropagation(); selectNode(node.id); }} title="Edit" style={{ background: 'none', border: 'none', color: COLORS.textMuted, cursor: 'pointer', display: 'flex', padding: 3, borderRadius: 4 }}>
              <Edit3 size={13} />
            </button>
          </span>
        </div>
        {node.children && isOpen && node.children.map((c, i) => renderRow(c, depth + 1, [...ancestorContinues, depth > 0 ? !isLast : false], i === node.children.length - 1))}
      </div>
    );
  }

  if (loading) {
    return <div className="px-6 pt-4 pb-6"><div style={{ color: COLORS.textMuted, fontSize: 13 }}>Loading hierarchy…</div></div>;
  }
  if (loadError || !tree) {
    return <div className="px-6 pt-4 pb-6"><div style={{ color: COLORS.brick, fontSize: 13 }}>{loadError || 'Could not load this dimension.'}</div></div>;
  }

  const parentOfSelected = selectedNode && selectedNode.id !== tree.id ? hierFindParent(tree, selectedNode.id, null) : null;
  const breadcrumb = selectedNode
    ? (selectedNode.type === 'leaf' && parentOfSelected ? `${currentDim} / ${parentOfSelected.name} / ${selectedNode.name}` : `${currentDim} / ${selectedNode.name}`)
    : '';

  return (
    <div className="px-6 pt-4 pb-6">
      <div style={{ background: COLORS.violetSoft, color: COLORS.violet, borderRadius: 8, padding: '8px 12px', fontSize: 12.5, marginBottom: 12 }}>
        Wired to your real {currentDim.toLowerCase()} structure — edits here save for real. Renaming or restructuring doesn't yet change how the P&L grid itself calculates.
      </div>
      {saveError && <div style={{ background: COLORS.brickSoft, color: COLORS.brick, borderRadius: 8, padding: '8px 12px', fontSize: 12.5, marginBottom: 12 }}>{saveError}</div>}
      <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 10, overflow: 'hidden', display: 'flex', minHeight: 460 }}>
        <div style={{ flex: 1, padding: '16px 20px', borderRight: `1px solid ${COLORS.border}` }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 14 }}>
            <select value={currentDim} onChange={(e) => switchDim(e.target.value)} className="lw-select" style={{ border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '6px 10px', fontSize: 13 }}>
              <option>Products</option>
              <option>Accounts</option>
              <option>Entities</option>
              <option>Time</option>
            </select>
            <div style={{ position: 'relative', flex: 1, minWidth: 0 }}>
              <Search size={14} style={{ position: 'absolute', left: 10, top: 9, color: COLORS.textMuted }} />
              <input value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} placeholder="Filter..." style={{ width: '100%', border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: '8px 12px 8px 32px', fontSize: 13 }} />
            </div>
            <button onClick={addCategoryToRoot} style={{ display: 'flex', alignItems: 'center', gap: 5, color: COLORS.jade, fontSize: 12.5, fontWeight: 600, background: COLORS.jadeSoft, border: 'none', borderRadius: 6, padding: '8px 12px', cursor: 'pointer', whiteSpace: 'nowrap' }}>
              <Plus size={13} /> Add category
            </button>
            <button onClick={exportToExcel} style={{ display: 'flex', alignItems: 'center', gap: 5, color: COLORS.violet, fontSize: 12.5, fontWeight: 600, background: COLORS.violetSoft, border: 'none', borderRadius: 6, padding: '8px 12px', cursor: 'pointer', whiteSpace: 'nowrap' }}>
              <Download size={13} /> Export to Excel
            </button>
          </div>
          {renderRow(tree, 0, [], true)}
        </div>

        <div style={{ width: 300, padding: '16px 20px', background: COLORS.surfaceAlt }}>
          {!selectedNode ? (
            <div style={{ color: COLORS.textMuted, fontSize: 12.5, lineHeight: 1.6 }}>
              Select a category or item on the left to rename, move, or delete it. Drag any item onto a category to move it there.
            </div>
          ) : (
            <>
              <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 2 }}>
                <button onClick={() => selectNode(null)} title="Close" style={{ background: 'none', border: 'none', color: COLORS.textMuted, cursor: 'pointer', display: 'flex', padding: 4 }}>
                  <X size={16} />
                </button>
              </div>
              <div style={{ fontSize: 11, color: COLORS.textMuted, textTransform: 'uppercase', letterSpacing: '.04em', marginBottom: 4 }}>Name</div>
              <input
                value={selectedNode.name}
                onChange={(e) => renameNode(selectedNode, e.target.value)}
                style={{ width: '100%', border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: '8px 10px', fontSize: 14, marginBottom: 16, background: COLORS.surface }}
              />
              <span style={{ display: 'inline-block', fontSize: 11, padding: '3px 9px', borderRadius: 4, fontWeight: 600, marginBottom: 16, background: selectedNode.type === 'cat' ? COLORS.jadeSoft : COLORS.violetSoft, color: selectedNode.type === 'cat' ? COLORS.jade : COLORS.violet }}>
                {selectedNode.type === 'cat' ? 'Category' : 'Item'}
              </span>
              <div style={{ fontSize: 12, color: COLORS.textMuted, marginBottom: 16, fontFamily: "'IBM Plex Mono', monospace" }}>{breadcrumb}</div>

              <div style={{ marginBottom: 16 }}>
                <div style={{ fontSize: 11, color: COLORS.textMuted, textTransform: 'uppercase', letterSpacing: '.04em', marginBottom: 6 }}>Attributes</div>
                {attributeDefs.length === 0 && <div style={{ color: COLORS.textMuted, fontSize: 12.5, marginBottom: 8 }}>No attributes defined for this dimension yet.</div>}
                {attributeDefs.map((def) => (
                  <div key={def} style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 6 }}>
                    <span style={{ width: 80, flexShrink: 0, fontSize: 11.5, color: COLORS.textMuted }}>{def}</span>
                    <input
                      value={(selectedNode.attrs && selectedNode.attrs[def]) || ''}
                      onChange={(e) => setAttrValue(selectedNode, def, e.target.value)}
                      placeholder="—"
                      style={{ flex: 1, minWidth: 0, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '6px 8px', fontSize: 12.5, background: COLORS.surface }}
                    />
                    <button onClick={() => removeAttributeDef(def)} title="Remove this attribute from the whole dimension" style={{ background: 'none', border: 'none', color: COLORS.textMuted, cursor: 'pointer', display: 'flex', padding: 2, flexShrink: 0 }}>
                      <Trash2 size={12} />
                    </button>
                  </div>
                ))}
                <div style={{ display: 'flex', gap: 6, marginTop: 4 }}>
                  <input
                    value={newAttrName}
                    onChange={(e) => setNewAttrName(e.target.value)}
                    onKeyDown={(e) => { if (e.key === 'Enter') addAttributeDef(); }}
                    placeholder="New attribute name"
                    style={{ flex: 1, border: `1px dashed ${COLORS.border}`, borderRadius: 6, padding: '6px 8px', fontSize: 12, background: 'transparent' }}
                  />
                  <button onClick={addAttributeDef} style={{ background: COLORS.violetSoft, color: COLORS.violet, border: 'none', borderRadius: 6, padding: '0 10px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Add</button>
                </div>
              </div>

              <div style={{ fontSize: 12, color: COLORS.textMuted, background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: '10px 12px', marginBottom: 16, lineHeight: 1.5 }}>
                {usageCount === null ? 'Checking usage…' : (
                  <>Used in <b style={{ color: COLORS.textDark, fontFamily: "'IBM Plex Mono', monospace" }}>{usageCount}</b> real budget cells.</>
                )}
              </div>

              {selectedNode.id === tree.id ? (
                <div style={{ color: COLORS.textMuted, fontSize: 12.5 }}>{selectedNode.name} is the root of the {currentDim} dimension and can't be deleted.</div>
              ) : !confirmingDelete ? (
                <button onClick={() => setConfirmingDelete(true)} style={{ width: '100%', background: COLORS.brickSoft, color: COLORS.brick, border: 'none', borderRadius: 8, padding: 9, fontSize: 13, fontWeight: 600, cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6 }}>
                  <Trash2 size={14} /> Delete
                </button>
              ) : (
                <div style={{ background: COLORS.brick, color: '#fff', borderRadius: 8, padding: '10px 12px', fontSize: 12, lineHeight: 1.5 }}>
                  This removes "{selectedNode.name}"{selectedNode.type === 'cat' && selectedNode.children.length ? ' and everything inside it' : ''} from the hierarchy. Real budget data stays intact — only this structural metadata is deleted.
                  <div style={{ display: 'flex', gap: 8, marginTop: 8 }}>
                    <button onClick={deleteSelected} style={{ flex: 1, background: '#fff', color: COLORS.brick, border: 'none', borderRadius: 6, padding: 7, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Yes, delete</button>
                    <button onClick={() => setConfirmingDelete(false)} style={{ flex: 1, background: 'rgba(255,255,255,.18)', color: '#fff', border: 'none', borderRadius: 6, padding: 7, fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>Cancel</button>
                  </div>
                </div>
              )}
            </>
          )}
        </div>
      </div>
    </div>
  );
}

/* ----------------------------------------------------------------------
   AUTH SCREEN — login / signup / forgot password / reset password
---------------------------------------------------------------------- */
function AuthScreen({ onAuthenticated, initialResetToken, initialInviteToken }) {
  const [mode, setMode] = useState(initialResetToken ? 'reset' : (initialInviteToken ? 'signup' : 'login')); // login | signup | forgot | reset
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [workspaceName, setWorkspaceName] = useState('');
  const [resetToken] = useState(initialResetToken || '');
  const [inviteToken] = useState(initialInviteToken || '');
  const [inviteInfo, setInviteInfo] = useState(null);
  const [inviteError, setInviteError] = useState('');
  const [newPassword, setNewPassword] = useState('');
  const [busy, setBusy] = useState(false);
  const [showPassword, setShowPassword] = useState(false);
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');

  useEffect(() => {
    if (!inviteToken) return;
    fetch(`${API_BASE}/invites/${inviteToken}`)
      .then((r) => r.json().then((data) => ({ ok: r.ok, data })))
      .then(({ ok, data }) => {
        if (!ok) { setInviteError(data.error || 'This invite is no longer valid.'); return; }
        setInviteInfo(data);
        setEmail(data.email);
      })
      .catch(() => setInviteError('Could not check this invite — it may be no longer valid.'));
  }, [inviteToken]);

  const inputStyle = {
    width: '100%', border: `1px solid ${COLORS.chromeBorder}`, background: COLORS.bgChrome2, color: COLORS.textOnDark,
    borderRadius: 8, padding: '10px 12px', fontSize: 14, marginBottom: 12,
  };
  const labelStyle = { color: COLORS.textOnDarkMuted, fontSize: 12.5, marginBottom: 4, display: 'block' };
  const buttonStyle = {
    width: '100%', background: COLORS.jade, color: '#fff', border: 'none', borderRadius: 8,
    padding: '11px 12px', fontSize: 14, fontWeight: 600, cursor: busy ? 'default' : 'pointer', opacity: busy ? 0.7 : 1,
  };

  async function submit(e) {
    e.preventDefault();
    setError('');
    setNotice('');
    setBusy(true);
    try {
      if (mode === 'login' || mode === 'signup') {
        const path = mode === 'login' ? 'login' : 'signup';
        const body = mode === 'login' ? { email, password } : { email, password, workspaceName, inviteToken: inviteInfo ? inviteToken : undefined };
        const r = await fetch(`${API_BASE}/auth/${path}`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
        const data = await r.json();
        if (!r.ok) throw new Error(data.error || 'something went wrong');

        // Logging in doesn't auto-join an invited workspace the way signup does —
        // accept it explicitly with the fresh login token before continuing.
        if (mode === 'login' && inviteInfo) {
          const acceptR = await fetch(`${API_BASE}/invites/${inviteToken}/accept`, { method: 'POST', headers: { Authorization: `Bearer ${data.token}` } });
          const acceptData = await acceptR.json();
          if (!acceptR.ok) throw new Error(acceptData.error || 'Could not join that workspace.');
          onAuthenticated(acceptData);
        } else {
          onAuthenticated(data);
        }
      } else if (mode === 'forgot') {
        const r = await fetch(`${API_BASE}/auth/forgot-password`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ email }) });
        if (!r.ok) throw new Error('something went wrong');
        setNotice("If that email has an account, we've sent a reset link. (Running locally without email configured? Check the backend terminal — the link is logged there.)");
        setMode('login');
      } else if (mode === 'reset') {
        const r = await fetch(`${API_BASE}/auth/reset-password`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ token: resetToken, newPassword }) });
        const data = await r.json();
        if (!r.ok) throw new Error(data.error || 'something went wrong');
        setNotice('Password updated — log in with your new password.');
        setMode('login');
      }
    } catch (err) {
      setError(err.message);
    } finally {
      setBusy(false);
    }
  }

  return (
    <div style={{ fontFamily: "'IBM Plex Sans', sans-serif", background: COLORS.bgChrome, minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20 }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=IBM+Plex+Sans:wght@400;500;600&display=swap'); input { color-scheme: dark; }`}</style>
      <div style={{ width: 360, background: COLORS.bgChrome2, border: `1px solid ${COLORS.chromeBorder}`, borderRadius: 12, padding: 28 }}>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", color: COLORS.textOnDark }} className="text-2xl font-bold tracking-tight mb-1">Rosebud</div>
        <div style={{ color: COLORS.textOnDarkMuted }} className="text-xs mb-6">
          {mode === 'login' && 'Log in to your workspace'}
          {mode === 'signup' && (inviteInfo ? `Join ${inviteInfo.workspaceName}` : 'Create your workspace')}
          {mode === 'forgot' && 'Reset your password'}
          {mode === 'reset' && 'Choose a new password'}
        </div>

        {inviteInfo && (
          <div style={{ background: COLORS.violetSoft, color: COLORS.violet, borderRadius: 6, padding: '8px 10px', fontSize: 12.5, marginBottom: 12 }}>
            You've been invited to <b>{inviteInfo.workspaceName}</b> as a <b>{inviteInfo.role}</b>. Sign up, or log in if you already have an account, using {inviteInfo.email}.
          </div>
        )}
        {inviteError && <div style={{ background: COLORS.brickSoft, color: COLORS.brick, borderRadius: 6, padding: '8px 10px', fontSize: 12.5, marginBottom: 12 }}>{inviteError}</div>}
        {notice && <div style={{ background: COLORS.jadeSoft, color: COLORS.jade, borderRadius: 6, padding: '8px 10px', fontSize: 12.5, marginBottom: 12 }}>{notice}</div>}
        {error && <div style={{ background: COLORS.brickSoft, color: COLORS.brick, borderRadius: 6, padding: '8px 10px', fontSize: 12.5, marginBottom: 12 }}>{error}</div>}

        <form onSubmit={submit}>
          {mode === 'signup' && !inviteInfo && (
            <>
              <label style={labelStyle}>Workspace name (optional)</label>
              <input style={inputStyle} type="text" placeholder="Acme Inc." value={workspaceName} onChange={(e) => setWorkspaceName(e.target.value)} />
            </>
          )}

          {mode !== 'reset' && (
            <>
              <label style={labelStyle}>Email</label>
              <input
                style={{ ...inputStyle, opacity: inviteInfo ? 0.6 : 1 }}
                type="email" required value={email} onChange={(e) => setEmail(e.target.value)}
                placeholder="you@company.com" disabled={!!inviteInfo}
              />
            </>
          )}

          {(mode === 'login' || mode === 'signup') && (
            <>
              <label style={labelStyle}>Password</label>
              <div style={{ position: 'relative', marginBottom: 12 }}>
                <input
                  style={{ ...inputStyle, marginBottom: 0, paddingRight: 40 }}
                  type={showPassword ? 'text' : 'password'}
                  required minLength={8} value={password} onChange={(e) => setPassword(e.target.value)}
                  placeholder={mode === 'signup' ? 'At least 8 characters' : ''}
                />
                <button
                  type="button" onClick={() => setShowPassword((s) => !s)} tabIndex={-1}
                  style={{ position: 'absolute', right: 10, top: '50%', transform: 'translateY(-50%)', background: 'none', border: 'none', color: COLORS.textOnDarkMuted, cursor: 'pointer', display: 'flex', padding: 0 }}
                  aria-label={showPassword ? 'Hide password' : 'Show password'}
                >
                  {showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
                </button>
              </div>
            </>
          )}

          {mode === 'reset' && (
            <>
              <label style={labelStyle}>New password</label>
              <div style={{ position: 'relative', marginBottom: 12 }}>
                <input
                  style={{ ...inputStyle, marginBottom: 0, paddingRight: 40 }}
                  type={showPassword ? 'text' : 'password'}
                  required minLength={8} value={newPassword} onChange={(e) => setNewPassword(e.target.value)}
                  placeholder="At least 8 characters"
                />
                <button
                  type="button" onClick={() => setShowPassword((s) => !s)} tabIndex={-1}
                  style={{ position: 'absolute', right: 10, top: '50%', transform: 'translateY(-50%)', background: 'none', border: 'none', color: COLORS.textOnDarkMuted, cursor: 'pointer', display: 'flex', padding: 0 }}
                  aria-label={showPassword ? 'Hide password' : 'Show password'}
                >
                  {showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
                </button>
              </div>
            </>
          )}

          <button type="submit" style={buttonStyle} disabled={busy}>
            {busy ? 'Please wait…' : { login: 'Log in', signup: inviteInfo ? 'Join workspace' : 'Create workspace', forgot: 'Send reset link', reset: 'Update password' }[mode]}
          </button>
        </form>

        <div style={{ marginTop: 16, display: 'flex', justifyContent: 'space-between', fontSize: 12.5 }}>
          {mode === 'login' && (
            <>
              <button onClick={() => { setMode('signup'); setError(''); setNotice(''); }} style={{ background: 'none', border: 'none', color: COLORS.textOnDarkMuted, cursor: 'pointer' }}>{inviteInfo ? 'Sign up to join' : 'Create a workspace'}</button>
              <button onClick={() => { setMode('forgot'); setError(''); setNotice(''); }} style={{ background: 'none', border: 'none', color: COLORS.textOnDarkMuted, cursor: 'pointer' }}>Forgot password?</button>
            </>
          )}
          {mode !== 'login' && (
            <button onClick={() => { setMode('login'); setError(''); setNotice(''); }} style={{ background: 'none', border: 'none', color: COLORS.textOnDarkMuted, cursor: 'pointer' }}>← Back to log in</button>
          )}
        </div>
      </div>
    </div>
  );
}

/* ----------------------------------------------------------------------
   INVITE GATE — handles a ?inviteToken= link when the user is already
   logged into some account, so the invite is never silently ignored.
---------------------------------------------------------------------- */
function InviteGate({ currentToken, inviteToken, onAuthenticated, onLogout, onDismiss }) {
  const [inviteInfo, setInviteInfo] = useState(null);
  const [inviteError, setInviteError] = useState('');
  const [actionError, setActionError] = useState('');
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    fetch(`${API_BASE}/invites/${inviteToken}`)
      .then((r) => r.json().then((data) => ({ ok: r.ok, data })))
      .then(({ ok, data }) => { if (!ok) setInviteError(data.error || 'This invite is no longer valid.'); else setInviteInfo(data); })
      .catch(() => setInviteError('Could not check this invite — it may be no longer valid.'));
  }, [inviteToken]);

  // Not logged in at all — the normal auth screen already handles invites end to end.
  if (!currentToken) {
    return <AuthScreen onAuthenticated={onAuthenticated} initialResetToken="" initialInviteToken={inviteToken} />;
  }

  function acceptWithCurrentAccount() {
    setBusy(true);
    setActionError('');
    fetch(`${API_BASE}/invites/${inviteToken}/accept`, { method: 'POST', headers: { Authorization: `Bearer ${currentToken}` } })
      .then((r) => r.json().then((data) => ({ ok: r.ok, data })))
      .then(({ ok, data }) => {
        if (!ok) { setActionError(data.error || 'Could not accept this invite.'); setBusy(false); return; }
        onAuthenticated(data);
      })
      .catch(() => { setActionError('Could not accept this invite.'); setBusy(false); });
  }

  const btnPrimary = { width: '100%', background: COLORS.jade, color: '#fff', border: 'none', borderRadius: 8, padding: '11px 12px', fontSize: 14, fontWeight: 600, cursor: busy ? 'default' : 'pointer', opacity: busy ? 0.7 : 1, marginBottom: 10 };
  const btnSecondary = { width: '100%', background: 'none', color: COLORS.textOnDarkMuted, border: `1px solid ${COLORS.chromeBorder}`, borderRadius: 8, padding: '10px 12px', fontSize: 13, cursor: 'pointer', marginBottom: 10 };

  return (
    <div style={{ fontFamily: "'IBM Plex Sans', sans-serif", background: COLORS.bgChrome, minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20 }}>
      <div style={{ width: 380, background: COLORS.bgChrome2, border: `1px solid ${COLORS.chromeBorder}`, borderRadius: 12, padding: 28 }}>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", color: COLORS.textOnDark }} className="text-2xl font-bold tracking-tight mb-1">Rosebud</div>
        <div style={{ color: COLORS.textOnDarkMuted }} className="text-xs mb-5">You're already logged in — here's what to do with this invite.</div>

        {inviteError && <div style={{ background: COLORS.brickSoft, color: COLORS.brick, borderRadius: 6, padding: '8px 10px', fontSize: 12.5, marginBottom: 12 }}>{inviteError}</div>}
        {actionError && <div style={{ background: COLORS.brickSoft, color: COLORS.brick, borderRadius: 6, padding: '8px 10px', fontSize: 12.5, marginBottom: 12 }}>{actionError}</div>}

        {inviteInfo && (
          <>
            <div style={{ color: COLORS.textOnDark }} className="text-sm mb-5">
              You've been invited to <b>{inviteInfo.workspaceName}</b> as a <b>{inviteInfo.role}</b>, sent to <b>{inviteInfo.email}</b>.
            </div>
            <button onClick={acceptWithCurrentAccount} disabled={busy} style={btnPrimary}>
              {busy ? 'Please wait…' : 'Accept with my current account'}
            </button>
            <button onClick={onLogout} style={btnSecondary}>Log out to use a different account</button>
          </>
        )}
        <button onClick={onDismiss} style={{ width: '100%', background: 'none', color: COLORS.textOnDarkMuted, border: 'none', fontSize: 12.5, cursor: 'pointer' }}>
          Skip and go to my workspace
        </button>
      </div>
    </div>
  );
}

/* ----------------------------------------------------------------------
   APP  — owns the auth token, renders InviteGate, AuthScreen, or Workspace
---------------------------------------------------------------------- */
export default function App() {
  const [token, setToken] = useState(() => localStorage.getItem('rosebud_token'));
  const [workspaceName, setWorkspaceName] = useState(() => localStorage.getItem('rosebud_workspace_name') || '');
  const [resetToken] = useState(() => new URLSearchParams(window.location.search).get('resetToken') || '');
  const [inviteToken, setInviteToken] = useState(() => new URLSearchParams(window.location.search).get('inviteToken') || '');

  function clearInviteFromUrl() {
    setInviteToken('');
    window.history.replaceState({}, '', window.location.pathname);
  }
  function handleAuthenticated({ token: t, workspace }) {
    localStorage.setItem('rosebud_token', t);
    localStorage.setItem('rosebud_workspace_name', workspace.name);
    setToken(t);
    setWorkspaceName(workspace.name);
    if (resetToken || inviteToken) window.history.replaceState({}, '', window.location.pathname);
    setInviteToken('');
  }
  function handleLogout() {
    localStorage.removeItem('rosebud_token');
    localStorage.removeItem('rosebud_workspace_name');
    setToken(null);
  }

  // An invite link always takes priority, even for an already-logged-in user —
  // otherwise the invite gets silently ignored in favor of whatever account
  // happens to already be signed in on this browser.
  if (inviteToken) {
    return <InviteGate currentToken={token} inviteToken={inviteToken} onAuthenticated={handleAuthenticated} onLogout={handleLogout} onDismiss={clearInviteFromUrl} />;
  }
  if (!token) {
    return <AuthScreen onAuthenticated={handleAuthenticated} initialResetToken={resetToken} />;
  }
  // A single wildcard route keeps Workspace mounted continuously across "/",
  // "/members", "/backups", and "/versions" — it reads the current path itself
  // via useLocation(). Mapping each path to its own <Route> would cause React
  // Router to unmount/remount Workspace on every panel switch, wiping local
  // state (current entity/scenario selections, etc.) and re-fetching everything.
  return (
    <BrowserRouter>
      <Routes>
        <Route
          path="/*"
          element={<Workspace token={token} workspaceName={workspaceName} onLogout={handleLogout} onAuthError={handleLogout} onSwitchWorkspace={handleAuthenticated} />}
        />
      </Routes>
    </BrowserRouter>
  );
}
