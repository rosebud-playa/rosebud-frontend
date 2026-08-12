import React, { useState, useMemo, useEffect } from 'react';
import {
  ChevronRight, ChevronDown, Save, RotateCcw, Trash2, X, ArrowLeftRight, History, Eye, EyeOff,
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

/* ----------------------------------------------------------------------
   MODEL ENGINE  (pure functions — operate on a plain data object)
---------------------------------------------------------------------- */
function childrenOf(id) { return ACCOUNTS.filter((a) => a.parentId === id); }
function hasChildren(id) { return ACCOUNTS.some((a) => a.parentId === id); }
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
  if (period === 'FY') return MONTHS.reduce((s, m) => s + getPeriodValue(data, entityId, accountId, m, product), 0);
  return (QUARTER_MONTHS[period] || []).reduce((s, m) => s + getPeriodValue(data, entityId, accountId, m, product), 0);
}

/* ----------------------------------------------------------------------
   FORMATTERS
---------------------------------------------------------------------- */
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

/* ----------------------------------------------------------------------
   APP
---------------------------------------------------------------------- */
function Workspace({ token, workspaceName, onLogout, onAuthError, onSwitchWorkspace }) {
  const [values, setValues] = useState({});
  const [workspaceLoaded, setWorkspaceLoaded] = useState(false);
  const [loadError, setLoadError] = useState(null);
  const [saveStatus, setSaveStatus] = useState('idle'); // idle | saving | saved | error
  const [currentScenario, setCurrentScenario] = useState('Budget');
  const [currentEntity, setCurrentEntity] = useState('company');
  const [currentProduct, setCurrentProduct] = useState('all');
  const [granularity, setGranularity] = useState('Monthly');
  const [expanded, setExpanded] = useState(() => new Set(['revenue', 'expenses']));
  const [compareMode, setCompareMode] = useState(false);
  const [compareTarget, setCompareTarget] = useState('Actual');
  const [comparePeriod, setComparePeriod] = useState('FY');
  const [versions, setVersions] = useState([]);
  const [showVersions, setShowVersions] = useState(false);
  const [versionLabel, setVersionLabel] = useState('');
  const [role, setRole] = useState('viewer');
  const [myWorkspaces, setMyWorkspaces] = useState([]);
  const [showMembers, setShowMembers] = useState(false);
  const [members, setMembers] = useState([]);
  const [inviteEmail, setInviteEmail] = useState('');
  const [inviteRole, setInviteRole] = useState('editor');
  const [membersError, setMembersError] = useState('');
  const [lastInviteLink, setLastInviteLink] = useState('');

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

  function loadMembers() {
    fetch(`${API_BASE}/workspace/members`, { headers: { Authorization: `Bearer ${token}` } })
      .then((r) => r.json())
      .then((data) => setMembers(data.members || []))
      .catch(() => {});
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

  const liveData = values[currentScenario];

  const columns = useMemo(() => {
    if (granularity === 'Monthly') return [...MONTHS, 'FY'];
    if (granularity === 'Quarterly') return ['Q1', 'Q2', 'Q3', 'Q4', 'FY'];
    return ['FY'];
  }, [granularity]);

  const visibleRows = useMemo(() => {
    const rows = [];
    (function walk(parentId, depth) {
      ACCOUNTS.filter((a) => a.parentId === parentId).forEach((a) => {
        rows.push({ ...a, depth });
        if (hasChildren(a.id) && expanded.has(a.id)) walk(a.id, depth + 1);
      });
    })(null, 0);
    return rows;
  }, [expanded]);

  function toggleExpand(id) {
    setExpanded((prev) => {
      const n = new Set(prev);
      n.has(id) ? n.delete(id) : n.add(id);
      return n;
    });
  }

  function updateValue(accountId, month, raw) {
    const num = raw === '' ? 0 : parseFloat(raw);
    const value = isNaN(num) ? 0 : num;
    const productKey = PRODUCT_DIMENSIONED_ACCOUNTS.includes(accountId) ? currentProduct : 'none';

    // Optimistic local update so typing feels instant.
    setValues((prev) => {
      const next = { ...prev };
      const scen = { ...next[currentScenario] };
      const ent = { ...scen[currentEntity] };
      const accData = { ...ent[accountId] };
      accData[productKey] = { ...accData[productKey], [month]: value };
      ent[accountId] = accData;
      scen[currentEntity] = ent;
      next[currentScenario] = scen;
      return next;
    });

    setSaveStatus('saving');
    fetch(`${API_BASE}/cell`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify({ scenario: currentScenario, entity: currentEntity, account: accountId, month, value, product: productKey }),
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
    setShowVersions(false);
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

  const kpiRevenue = getPeriodValue(liveData, currentEntity, 'revenue', 'FY');
  const kpiExpenses = getPeriodValue(liveData, currentEntity, 'expenses', 'FY');
  const kpiNet = kpiRevenue - kpiExpenses;
  const kpiMargin = kpiRevenue !== 0 ? (kpiNet / kpiRevenue) * 100 : 0;

  const chartData = MONTHS.map((m) => {
    const rev = getValueAt(liveData, currentEntity, 'revenue', m);
    const exp = getValueAt(liveData, currentEntity, 'expenses', m);
    return { month: m, Revenue: rev, Expenses: exp, 'Net Income': rev - exp };
  });

  const compareData = compareMode ? resolveData(compareTarget) : null;
  const compareOptions = [
    ...SCENARIOS.map((s) => ({ key: s, label: s })),
    ...versions.map((v) => ({ key: `v:${v.id}`, label: `${v.label} (saved)` })),
  ].filter((o) => o.key !== currentScenario || o.key.startsWith('v:'));

  return (
    <div style={{ fontFamily: "'IBM Plex Sans', sans-serif", background: COLORS.surfaceAlt, color: COLORS.textDark, minHeight: '100%' }} className="w-full">
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
              onClick={() => { setShowMembers((s) => !s); if (!showMembers) loadMembers(); }}
              style={{
                display: 'flex', alignItems: 'center', gap: 6, padding: '5px 10px', borderRadius: 6, fontSize: 12.5,
                background: showMembers ? COLORS.violet : 'none', color: showMembers ? '#fff' : COLORS.textOnDarkMuted,
                border: `1px solid ${showMembers ? COLORS.violet : COLORS.chromeBorder}`, cursor: 'pointer',
              }}
            >
              Members
            </button>
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
            <select
              className="lw-select"
              value={currentEntity}
              onChange={(e) => setCurrentEntity(e.target.value)}
              style={{ background: COLORS.bgChrome2, color: COLORS.textOnDark, border: `1px solid ${COLORS.chromeBorder}`, borderRadius: 6, padding: '6px 10px', fontSize: 13 }}
            >
              {ENTITY_OPTIONS.map((e) => <option key={e.id} value={e.id}>{e.name}</option>)}
            </select>

            <select
              className="lw-select"
              value={currentProduct}
              onChange={(e) => setCurrentProduct(e.target.value)}
              title="Filters Product Revenue, Units Sold, and Unit Price only — every other line always shows the full total."
              style={{ background: COLORS.bgChrome2, color: COLORS.textOnDark, border: `1px solid ${COLORS.chromeBorder}`, borderRadius: 6, padding: '6px 10px', fontSize: 13 }}
            >
              <option value="all">All Products</option>
              {PRODUCT_CATEGORIES.map((cat) => (
                <optgroup key={cat.id} label={cat.name}>
                  {PRODUCTS.filter((p) => p.category === cat.id).map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
                </optgroup>
              ))}
            </select>

            <select
              className="lw-select"
              value={currentScenario}
              onChange={(e) => setCurrentScenario(e.target.value)}
              style={{ background: COLORS.bgChrome2, color: COLORS.textOnDark, border: `1px solid ${COLORS.chromeBorder}`, borderRadius: 6, padding: '6px 10px', fontSize: 13 }}
            >
              {SCENARIOS.map((s) => <option key={s} value={s}>{s}</option>)}
            </select>

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

            <button
              onClick={() => setCompareMode((c) => !c)}
              style={{
                display: 'flex', alignItems: 'center', gap: 6, padding: '6px 10px', borderRadius: 6, fontSize: 13,
                background: compareMode ? COLORS.violet : COLORS.bgChrome2, color: '#fff', border: `1px solid ${compareMode ? COLORS.violet : COLORS.chromeBorder}`, cursor: 'pointer',
              }}
            >
              <ArrowLeftRight size={14} /> Compare
            </button>

            <button
              onClick={() => setShowVersions((s) => !s)}
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
          <div style={{ color: COLORS.textMuted }} className="text-xs mb-2">Monthly trend — {ENTITY_OPTIONS.find((e) => e.id === currentEntity)?.name}, {currentScenario}</div>
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
      <div className="px-6 pt-4 pb-2">
        {!compareMode && (!canEdit || currentProduct === 'all') && (
          <div style={{ color: COLORS.textMuted }} className="text-xs mb-2">
            {currentEntity === 'company' && 'Select a specific entity to edit inputs. '}
            {granularity !== 'Monthly' && 'Switch to Monthly view to edit inputs. '}
            {currentProduct === 'all' && 'Select a specific product to edit Units Sold or Unit Price.'}
          </div>
        )}
        <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}` }} className="rounded-lg overflow-hidden">
          <div className="overflow-x-auto">
            <table style={{ minWidth: compareMode ? 640 : columns.length * 92 + 220, borderCollapse: 'collapse' }} className="w-full text-sm">
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
                                  onChange={(e) => updateValue(row.id, col, e.target.value)}
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
        </div>
        <div style={{ color: COLORS.textMuted }} className="text-xs mt-2">
          Rail colors: <span style={{ color: COLORS.jade }}>■</span> rollup &nbsp; <span style={{ color: COLORS.violet }}>■</span> formula &nbsp; <span style={{ color: COLORS.amber }}>■</span> driver input &nbsp; <span style={{ color: COLORS.border }}>■</span> input. Edits save to the shared workspace on the budget server as you type.
        </div>
      </div>

      {/* ---------------- Versions drawer ---------------- */}
      {showVersions && (
        <div className="px-6 pb-6">
          <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}` }} className="rounded-lg p-4">
            <div className="flex items-center justify-between mb-3">
              <div className="font-semibold text-sm">Saved versions</div>
              <button onClick={() => setShowVersions(false)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: COLORS.textMuted }}><X size={16} /></button>
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
              <button onClick={() => setShowMembers(false)} style={{ background: 'none', border: 'none', cursor: 'pointer', color: COLORS.textMuted }}><X size={16} /></button>
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
              {members.map((m) => (
                <div key={m.userId} style={{ border: `1px solid ${COLORS.border}`, borderRadius: 8 }} className="flex items-center justify-between px-3 py-2">
                  <div>
                    <div className="text-sm font-medium">{m.email}</div>
                    <div style={{ color: COLORS.textMuted }} className="text-xs">joined {new Date(m.joinedAt).toLocaleDateString()}</div>
                  </div>
                  {canManageMembers ? (
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
                    <span style={{ background: COLORS.surfaceAlt, borderRadius: 4, padding: '2px 8px', fontSize: 11.5, textTransform: 'capitalize', color: COLORS.textMuted }}>{m.role}</span>
                  )}
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
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
          {mode === 'signup' && 'Create your workspace'}
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
            {busy ? 'Please wait…' : { login: 'Log in', signup: 'Create workspace', forgot: 'Send reset link', reset: 'Update password' }[mode]}
          </button>
        </form>

        <div style={{ marginTop: 16, display: 'flex', justifyContent: 'space-between', fontSize: 12.5 }}>
          {mode === 'login' && (
            <>
              <button onClick={() => { setMode('signup'); setError(''); setNotice(''); }} style={{ background: 'none', border: 'none', color: COLORS.textOnDarkMuted, cursor: 'pointer' }}>Create a workspace</button>
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
   APP  — owns the auth token, renders AuthScreen or Workspace
---------------------------------------------------------------------- */
export default function App() {
  const [token, setToken] = useState(() => localStorage.getItem('rosebud_token'));
  const [workspaceName, setWorkspaceName] = useState(() => localStorage.getItem('rosebud_workspace_name') || '');
  const [resetToken] = useState(() => new URLSearchParams(window.location.search).get('resetToken') || '');
  const [inviteToken] = useState(() => new URLSearchParams(window.location.search).get('inviteToken') || '');

  function handleAuthenticated({ token: t, workspace }) {
    localStorage.setItem('rosebud_token', t);
    localStorage.setItem('rosebud_workspace_name', workspace.name);
    setToken(t);
    setWorkspaceName(workspace.name);
    if (resetToken || inviteToken) window.history.replaceState({}, '', window.location.pathname);
  }
  function handleLogout() {
    localStorage.removeItem('rosebud_token');
    localStorage.removeItem('rosebud_workspace_name');
    setToken(null);
  }

  if (!token) {
    return <AuthScreen onAuthenticated={handleAuthenticated} initialResetToken={resetToken} initialInviteToken={inviteToken} />;
  }
  return <Workspace token={token} workspaceName={workspaceName} onLogout={handleLogout} onAuthError={handleLogout} onSwitchWorkspace={handleAuthenticated} />;
}
