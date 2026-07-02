/* ============================================================
   RICE Delivery Tracker — client application
   ============================================================ */
"use strict";

// Color a raw Object Status by keyword — prototype status scale
// (grey → blues → green, plus red/amber for blocked/delayed).
function statusColor(s) {
  const k = (s || "").toLowerCase();
  if (k.includes("block")) return "#cf5064";                        // red
  if (k.includes("delay")) return "#e0a33e";                        // amber
  if (k.includes("complete") || k.includes("done")) return "#3f9d6b"; // green
  if (k.includes("f-spec") || k.includes("fspec") || k.includes("f spec")) return "#7b74d1"; // mid indigo
  if (k.includes("fut")) return "#9aa0ee";                          // light indigo
  if (k.includes("dev-not") || k.includes("dev not")) return "#aeb6c2"; // grey (before "not started")
  if (/(dev-in|dev in|in progress|progress|draft|review|wip)/.test(k)) return "#5b52e0"; // indigo
  if (k.includes("not started") || k.includes("not-started")) return "#dfe3e9"; // light grey
  return "#c3c9d4";                                                 // fallback grey
}
// Canonical display rank for statuses (lower = earlier): green, blues, amber, red, greys.
function statusRank(s) {
  const k = (s || "").toLowerCase();
  if (k.includes("complete") || k.includes("done")) return 1;
  if (k.includes("dev-in") || k.includes("dev in") || (k.includes("in progress") && !k.includes("f-spec") && !k.includes("fspec"))) return 2;
  if (k.includes("f-spec") || k.includes("fspec") || k.includes("f spec")) return 3;
  if (k.includes("fut")) return 4;
  if (/(progress|draft|review|wip)/.test(k)) return 4.5;
  if (k.includes("delay")) return 5;
  if (k.includes("block")) return 6;
  if (k.includes("dev-not") || k.includes("dev not")) return 7;
  if (k.includes("not started") || k.includes("not-started")) return 8;
  return 9;
}
// Distinct raw Object Status values present, ordered by canonical rank.
function statusOrder(recs) {
  return [...new Set(recs.map(r => r.object_status).filter(Boolean))]
    .sort((a, b) => (statusRank(a) - statusRank(b)) || a.localeCompare(b));
}
const TYPE_ORDER = ["Conversion", "Integration", "Report", "Extension"];
const TYPE_COLOR = {
  "Conversion": "#0097A9", "Integration": "#00A3E0", "Report": "#046A38",
  "Extension": "#6E2585", "Unspecified": "#97999B"
};
const PHASE_FILL = {
  "Sprint 1": "rgba(134,188,37,.16)", "Sprint 2": "rgba(0,163,224,.14)",
  "Sprint 3": "rgba(110,37,133,.13)", "SIT 1": "rgba(237,139,0,.16)",
  "SIT 2": "rgba(237,139,0,.16)", "UAT": "rgba(110,37,133,.20)",
  "Cutover": "rgba(218,41,28,.16)", "Post Go-Live": "rgba(0,163,224,.20)"
};
// Gap tiles have dynamic names ("Gap b/w X and Y") so they can't be keyed in
// PHASE_FILL directly — fall back on type instead.
function phaseFillFor(name, type) {
  return PHASE_FILL[name] || (type === "Gap" ? "rgba(117,120,123,.12)" : "rgba(0,0,0,.03)");
}

const State = {
  data: null,            // full payload
  filtered: [],          // records after all filters
  charts: {},            // donut chart instances by id
  gridApi: null,
  planView: "grid",
  quick: {},             // transient quick-filters {rice_type, object_status, assigned_sprint}
  choices: {},
  allOptions: { org: [], module: [] },  // full universe of values, for "Select all"
};

/* ---------------- helpers ---------------- */
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, c =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const fmtNum = (n) => (n == null || isNaN(n)) ? "—" : Number(n).toLocaleString();
const fmtDate = (iso) => {
  if (!iso) return "—";
  const d = new Date(iso + "T00:00:00");
  if (isNaN(d)) return "—";
  return d.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "2-digit" });
};
const parseISO = (iso) => iso ? new Date(iso + "T00:00:00") : null;
const today = () => { const d = new Date(); d.setHours(0, 0, 0, 0); return d; };
// Local YYYY-MM-DD (avoids toISOString's UTC shift, which can split/merge weeks).
const ymd = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
const slug = (s) => String(s).replace(/[^a-zA-Z0-9]+/g, "");

function toast(msg) {
  const t = $("#toast"); t.textContent = msg; t.classList.add("show");
  clearTimeout(t._t); t._t = setTimeout(() => t.classList.remove("show"), 2200);
}

/* ============================================================
   BOOT
   ============================================================ */
document.addEventListener("DOMContentLoaded", () => {
  initTheme();
  initUpload();
  initTopbar();
  initFilters();
  initPlanToggle();
  initModal();
  $("#exportBtn").addEventListener("click", exportCSV);
  $("#reuploadBtn").addEventListener("click", () => {
    $("#uploadScreen").classList.add("show");
    $("#dashboard").classList.remove("show");
    window.scrollTo(0, 0);
  });

  if ($("#dashboard").classList.contains("show")) loadData();
});

/* ---------------- theme ---------------- */
function initTheme() {
  const saved = (() => { try { return localStorage.getItem("rice-theme"); } catch { return null; } })();
  if (saved) document.documentElement.setAttribute("data-theme", saved);
  $("#themeBtn").addEventListener("click", () => {
    const cur = document.documentElement.getAttribute("data-theme") === "dark" ? "light" : "dark";
    document.documentElement.setAttribute("data-theme", cur);
    try { localStorage.setItem("rice-theme", cur); } catch {}
    if (State.data) { renderTypeCards(State.filtered); if (State.planView === "gantt") renderGantt(State.filtered); }
  });
}

/* ---------------- topbar scroll-spy ---------------- */
function initTopbar() {
  const links = $$("#navlinks a");
  const map = links.map(a => ({ a, el: $(a.getAttribute("href")) })).filter(x => x.el);
  window.addEventListener("scroll", () => {
    let cur = null;
    for (const m of map) {
      if (m.el.getBoundingClientRect().top <= 120) cur = m;
    }
    links.forEach(a => a.classList.remove("active"));
    if (cur) cur.a.classList.add("active");
  }, { passive: true });
}

/* ============================================================
   UPLOAD
   ============================================================ */
function initUpload() {
  const dz = $("#dropZone"), input = $("#fileInput");
  $("#browseBtn").addEventListener("click", () => input.click());
  input.addEventListener("change", () => { if (input.files[0]) uploadFile(input.files[0]); });
  ["dragenter", "dragover"].forEach(ev => dz.addEventListener(ev, e => {
    e.preventDefault(); dz.classList.add("drag");
  }));
  ["dragleave", "drop"].forEach(ev => dz.addEventListener(ev, e => {
    e.preventDefault(); dz.classList.remove("drag");
  }));
  dz.addEventListener("drop", e => {
    const f = e.dataTransfer.files[0]; if (f) uploadFile(f);
  });
}

async function uploadFile(file) {
  const msg = $("#uploadMsg");
  msg.className = "upload-msg"; msg.textContent = "Uploading & validating…";
  const fd = new FormData(); fd.append("file", file);
  try {
    const r = await fetch("/api/upload", { method: "POST", body: fd });
    const j = await r.json();
    if (!r.ok) { msg.className = "upload-msg err"; msg.textContent = j.message || "Upload failed."; return; }
    msg.className = "upload-msg ok"; msg.textContent = "✓ Loaded. Opening dashboard…";
    $("#uploadScreen").classList.remove("show");
    $("#dashboard").classList.add("show");
    await loadData();
    window.scrollTo(0, 0);
  } catch (e) {
    msg.className = "upload-msg err"; msg.textContent = "Network error: " + e.message;
  }
}

/* ============================================================
   DATA LOAD
   ============================================================ */
async function loadData() {
  try {
    const r = await fetch("/api/data");
    if (r.status === 404) {
      $("#uploadScreen").classList.add("show");
      $("#dashboard").classList.remove("show");
      return;
    }
    const j = await r.json();
    if (j.error) { toast("Error: " + j.message); return; }
    State.data = j;
    populateFilterOptions(j);
    $("#footMeta").textContent =
      `Source sheet: ${j.source_sheet} · ${j.record_count} objects · generated ${j.generated_at.replace("T", " ")}`;
    $("#brandSub").textContent = `${j.summary.total_in_scope} in-scope of ${j.record_count} RICE objects`;
    applyFilters();
  } catch (e) {
    toast("Failed to load data: " + e.message);
  }
}

/* ============================================================
   FILTERS
   ============================================================ */
/* Filter dimensions rendered as pills (multi-select), in display order. */
const FILTER_DIMS = [
  { key: "org", label: "Org" },
  { key: "module", label: "Module" },
  { key: "release", label: "Release" },
  { key: "sub_entity", label: "Sub Entity" },
];
// Selected values per multi-select dimension (empty set = "Any"). Scope is single-select.
State.sel = { org: new Set(), module: new Set(), release: new Set(), sub_entity: new Set() };
State.scope = "";
State.scopeOptions = [];

function initFilters() {
  // Close any open popover when clicking outside a pill, or on Escape.
  document.addEventListener("click", (e) => {
    if (!e.target.closest(".fpill-wrap")) closeAllPopovers();
  });
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") closeAllPopovers(); });

  // Search (debounced).
  let t;
  $("#fSearch").addEventListener("input", () => { clearTimeout(t); t = setTimeout(applyFilters, 180); });

  // Saved filters.
  const savedWrap = $(`.fpill-wrap[data-dim="saved"]`);
  $("[data-toggle]", savedWrap).addEventListener("click", (e) => { e.stopPropagation(); togglePopover(savedWrap); });
  $("#saveViewBtn").addEventListener("click", saveCurrentView);
  loadSavedViews();
}

function closeAllPopovers() {
  $$(".fpill-wrap.open").forEach(w => w.classList.remove("open"));
}

const CHEV_SVG = `<svg class="fpill-chev" width="12" height="12" viewBox="0 0 12 12" fill="none"><path d="M3 4.5 6 7.5 9 4.5" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
const CHECK_SVG = `<svg width="11" height="11" viewBox="0 0 12 12" fill="none"><path d="M2.5 6 5 8.5 9.5 3.5" stroke="#fff" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
const DOT_SVG = `<svg width="10" height="10" viewBox="0 0 10 10"><circle cx="5" cy="5" r="4" fill="#5b52e0"/></svg>`;

function multiPillHTML(d) {
  return `<div class="fpill-wrap" data-dim="${d.key}">
    <button type="button" class="fpill" data-toggle>
      <span class="fpill-key">${esc(d.label)}</span>
      <span class="fpill-val">Any</span>${CHEV_SVG}
    </button>
    <div class="fpop">
      <div class="fpop-search">
        <svg width="16" height="16" viewBox="0 0 18 18" fill="none"><circle cx="8" cy="8" r="5.2" stroke="#a3abba" stroke-width="1.6"/><path d="m12.5 12.5 3 3" stroke="#a3abba" stroke-width="1.6" stroke-linecap="round"/></svg>
        <input type="text" placeholder="Filter ${esc(d.label.toLowerCase())}…" data-search autocomplete="off">
      </div>
      <div class="fpop-seg">
        <button type="button" data-act="all">All</button>
        <button type="button" data-act="none">None</button>
        <button type="button" data-act="invert">Invert</button>
      </div>
      <div class="fpop-list" data-list></div>
      <div class="fpop-foot">
        <span class="fpop-count" data-count>0 selected</span>
        <button type="button" class="fpop-done" data-done>Done</button>
      </div>
    </div>
  </div>`;
}

function scopePillHTML() {
  return `<div class="fpill-wrap" data-dim="scope">
    <button type="button" class="fpill" data-toggle>
      <span class="fpill-key">In Scope</span>
      <span class="fpill-val">Any</span>${CHEV_SVG}
    </button>
    <div class="fpop fpop-menu" data-menu></div>
  </div>`;
}

function wireMultiPill(d) {
  const wrap = $(`.fpill-wrap[data-dim="${d.key}"]`);
  const search = $("[data-search]", wrap);
  $("[data-toggle]", wrap).addEventListener("click", (e) => { e.stopPropagation(); togglePopover(wrap); });
  search.addEventListener("input", () => renderMultiList(d, search.value.trim().toLowerCase()));
  $$("[data-act]", wrap).forEach(b => b.addEventListener("click", () => {
    const act = b.getAttribute("data-act");
    const all = State.allOptions[d.key] || [];
    const sel = State.sel[d.key];
    if (act === "all") all.forEach(v => sel.add(v));
    else if (act === "none") sel.clear();
    else { const inv = all.filter(v => !sel.has(v)); sel.clear(); inv.forEach(v => sel.add(v)); }
    renderMultiList(d, search.value.trim().toLowerCase());
    onFilterChange(d.key);
  }));
  $("[data-done]", wrap).addEventListener("click", () => togglePopover(wrap, false));
  renderMultiList(d, "");
}

function renderMultiList(d, q) {
  const wrap = $(`.fpill-wrap[data-dim="${d.key}"]`);
  const list = $("[data-list]", wrap);
  const sel = State.sel[d.key];
  const opts = (State.allOptions[d.key] || []).filter(v => !q || v.toLowerCase().includes(q));
  list.innerHTML = opts.map(v => {
    const on = sel.has(v);
    return `<label class="fpop-opt${on ? " on" : ""}" data-val="${esc(v)}">
      <span class="fpop-box">${on ? CHECK_SVG : ""}</span>
      <span class="fpop-optlabel">${esc(v)}</span>
    </label>`;
  }).join("") || `<div class="fpop-empty">No matches</div>`;
  $$(".fpop-opt", list).forEach(opt => opt.addEventListener("click", (e) => {
    e.preventDefault();
    const v = opt.getAttribute("data-val");
    if (sel.has(v)) sel.delete(v); else sel.add(v);
    opt.classList.toggle("on");
    $(".fpop-box", opt).innerHTML = sel.has(v) ? CHECK_SVG : "";
    updateFootCount(d);
    onFilterChange(d.key);
  }));
  updateFootCount(d);
}

function updateFootCount(d) {
  const c = $(`.fpill-wrap[data-dim="${d.key}"] [data-count]`);
  if (c) c.textContent = `${State.sel[d.key].size} selected`;
}

function wireScopePill() {
  const wrap = $(`.fpill-wrap[data-dim="scope"]`);
  $("[data-toggle]", wrap).addEventListener("click", (e) => { e.stopPropagation(); togglePopover(wrap); });
  renderScopeMenu();
}

function renderScopeMenu() {
  const menu = $(`.fpill-wrap[data-dim="scope"] [data-menu]`);
  const opts = [["", "Any"]].concat((State.scopeOptions || []).map(v => [v, v]));
  menu.innerHTML = opts.map(([val, label]) => {
    const on = State.scope === val;
    return `<button type="button" class="fpop-menu-item${on ? " on" : ""}" data-scope="${esc(val)}">
      <span class="fpop-radio">${on ? DOT_SVG : ""}</span><span>${esc(label)}</span>
    </button>`;
  }).join("");
  $$("[data-scope]", menu).forEach(b => b.addEventListener("click", () => {
    State.scope = b.getAttribute("data-scope");
    renderScopeMenu();
    onFilterChange("scope");
    togglePopover($(`.fpill-wrap[data-dim="scope"]`), false);
  }));
}

function togglePopover(wrap, force) {
  const willOpen = force === undefined ? !wrap.classList.contains("open") : force;
  closeAllPopovers();
  if (willOpen) {
    wrap.classList.add("open");
    const s = $("[data-search]", wrap);
    if (s) { s.value = ""; const d = FILTER_DIMS.find(x => x.key === wrap.getAttribute("data-dim")); if (d) renderMultiList(d, ""); setTimeout(() => s.focus(), 0); }
  }
}

function onFilterChange(key) {
  updatePillLabel(key);
  applyFilters();
}

function updateAllPillLabels() {
  FILTER_DIMS.forEach(d => updatePillLabel(d.key));
  updatePillLabel("scope");
}

function updatePillLabel(key) {
  const wrap = $(`.fpill-wrap[data-dim="${key}"]`);
  if (!wrap) return;
  const valEl = $(".fpill-val", wrap);
  let text = "Any", active = false;
  if (key === "scope") {
    active = !!State.scope;
    text = State.scope || "Any";
  } else {
    const sel = State.sel[key];
    const total = (State.allOptions[key] || []).length;
    if (sel.size === 0) { text = "Any"; }
    else if (total && sel.size === total) { text = `All (${total})`; active = true; }
    else if (sel.size === 1) { text = [...sel][0]; active = true; }
    else { text = `${sel.size} selected`; active = true; }
  }
  valEl.textContent = text;
  wrap.classList.toggle("active", active);
}

// Build the pills + popovers once the filter option universes are known.
function populateFilterOptions(j) {
  const f = j.filters;
  State.allOptions = {
    org: (f.accountable_org || []).slice(),
    module: (f.module || []).slice(),
    release: (f.release || []).slice(),
    sub_entity: (f.sub_entity || []).slice(),
  };
  State.scopeOptions = (f.in_scope || []).slice();

  $("#fbPills").innerHTML = FILTER_DIMS.map(multiPillHTML).join("") + scopePillHTML();
  FILTER_DIMS.forEach(wireMultiPill);
  wireScopePill();
  updateAllPillLabels();
}

const SEARCH_FIELDS = ["rice_id", "object_name", "description", "module", "source_system", "target_system",
  "functional_owner", "technical_owner", "rice_owner", "tech_spec_owner"];

function getBaseFilters() {
  return {
    org: [...State.sel.org],
    module: [...State.sel.module],
    release: [...State.sel.release],
    sub_entity: [...State.sel.sub_entity],
    scope: State.scope,
    search: ($("#fSearch").value || "").trim().toLowerCase(),
  };
}

function setQuick(obj) {
  State.quick = obj || {};
  applyFilters();
  document.getElementById("sec-plan").scrollIntoView({ behavior: "smooth", block: "start" });
}

const QUICK_LABELS = { rice_type: "Type", object_status: "Status", assigned_sprint: "Sprint" };

function clearQuick(key) {
  const q = { ...State.quick };
  delete q[key];
  State.quick = q;
  applyFilters();
}

function renderQuickPills() {
  const el = $("#quickPills");
  if (!el) return;
  const keys = Object.keys(State.quick).filter(k => State.quick[k]);
  el.innerHTML = keys.map(k =>
    `<span class="qchip">${esc(QUICK_LABELS[k] || k)}: <b>${esc(State.quick[k])}</b>
      <button type="button" class="qchip-x" data-key="${esc(k)}" aria-label="Clear ${esc(QUICK_LABELS[k] || k)} filter">×</button>
    </span>`).join("");
  $$(".qchip-x", el).forEach(btn => btn.addEventListener("click", () => clearQuick(btn.dataset.key)));
}

function applyFilters() {
  if (!State.data) return;
  const f = getBaseFilters();
  const q = State.quick;
  State.filtered = State.data.records.filter(r => {
    if (f.org.length && !f.org.includes(r.accountable_org)) return false;
    if (f.module.length && !f.module.includes(r.module)) return false;
    if (f.release.length && !f.release.includes(r.release)) return false;
    if (f.sub_entity.length && !f.sub_entity.includes(r.sub_entity)) return false;
    if (f.scope && r.in_scope !== f.scope) return false;
    if (f.search) {
      const hay = SEARCH_FIELDS.map(k => r[k] || "").join(" ").toLowerCase();
      if (!hay.includes(f.search)) return false;
    }
    if (q.rice_type && r.rice_type !== q.rice_type) return false;
    if (q.object_status && r.object_status !== q.object_status) return false;
    if (q.assigned_sprint && r.assigned_sprint !== q.assigned_sprint) return false;
    return true;
  });
  updateResultCount();
  renderQuickPills();
  renderAll(State.filtered);
}

function updateResultCount() {
  const el = $("#resultCount");
  if (el) el.innerHTML = `<span>${State.filtered.length}</span> / ${State.data.record_count} objects`;
}

/* ---------------- saved filter sets (persisted server-side) ---------------- */
let SavedViews = {};

async function loadSavedViews() {
  try {
    const r = await fetch("/api/saved-filters");
    SavedViews = r.ok ? (await r.json()) || {} : {};
  } catch { SavedViews = {}; }
  renderSavedMenu();
}

// Snapshot every active filter so it can be restored later.
function captureFilters() {
  const f = getBaseFilters();
  return { org: f.org, module: f.module, release: f.release, sub_entity: f.sub_entity,
    scope: f.scope, search: $("#fSearch").value, quick: { ...State.quick } };
}

function renderSavedMenu() {
  const names = Object.keys(SavedViews).sort((a, b) => a.localeCompare(b));
  const list = $("#savedList");
  list.innerHTML = names.length
    ? names.map(n => `<div class="fsaved-row">
        <button type="button" class="fsaved-apply" data-name="${esc(n)}">${esc(n)}</button>
        <button type="button" class="fsaved-del" data-del="${esc(n)}" title="Delete">✕</button>
      </div>`).join("")
    : `<div class="fsaved-empty">No saved filters yet</div>`;
  $$(".fsaved-apply", list).forEach(b => b.addEventListener("click", () => {
    applySavedView(b.getAttribute("data-name"));
    togglePopover($(`.fpill-wrap[data-dim="saved"]`), false);
  }));
  $$(".fsaved-del", list).forEach(b => b.addEventListener("click", (e) => {
    e.stopPropagation();
    deleteSavedView(b.getAttribute("data-del"));
  }));
  const val = $("#savedPillVal");
  if (val) val.textContent = names.length ? `${names.length} saved` : "None";
}

async function saveCurrentView() {
  const name = (prompt("Name this filter set:") || "").trim();
  if (!name) return;
  if (SavedViews[name] && !confirm(`Overwrite the saved filter "${name}"?`)) return;
  try {
    const r = await fetch("/api/saved-filters", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name, view: captureFilters() }),
    });
    if (!r.ok) throw new Error("save failed");
    SavedViews = (await r.json()).filters || {};
    renderSavedMenu();
    toast(`Saved filter "${name}"`);
  } catch { toast("Could not save filter"); }
}

async function deleteSavedView(name) {
  if (!confirm(`Delete saved filter "${name}"?`)) return;
  try {
    const r = await fetch(`/api/saved-filters/${encodeURIComponent(name)}`, { method: "DELETE" });
    if (!r.ok) throw new Error("delete failed");
    SavedViews = (await r.json()).filters || {};
    renderSavedMenu();
    toast(`Deleted "${name}"`);
  } catch { toast("Could not delete filter"); }
}

function applySavedView(name) {
  const v = SavedViews[name];
  if (!v) return;
  ["org", "module", "release", "sub_entity"].forEach(k => {
    State.sel[k] = new Set(v[k] || []);
    updatePillLabel(k);
  });
  State.scope = v.scope || "";
  updatePillLabel("scope");
  renderScopeMenu();
  $("#fSearch").value = v.search || "";
  State.quick = { ...(v.quick || {}) };
  applyFilters();
}

/* ============================================================
   RENDER ALL
   ============================================================ */
function renderAll(recs) {
  renderTypeCards(recs);
  renderRawStatus(recs);
  renderSprintSummary(recs);
  renderGrid(recs);
  if (State.planView === "gantt") renderGantt(recs);
  renderCapacity(recs);
  renderRisk(recs);
  renderOwnerFocus(recs);
  renderMatrix(recs);
  renderDataQuality(recs);
}

/* ---------- counting helpers ---------- */
function countBy(recs, key) {
  const m = {};
  recs.forEach(r => { const v = r[key] || "—"; m[v] = (m[v] || 0) + 1; });
  return m;
}
function typeBreakdown(recs) {
  const m = {};
  recs.forEach(r => { m[r.rice_type] = (m[r.rice_type] || 0) + 1; });
  return m;
}

/* ============================================================
   RICE TYPE CARDS + DONUTS
   ============================================================ */
function renderTypeCards(recs) {
  const types = TYPE_ORDER.concat(
    [...new Set(recs.map(r => r.rice_type))].filter(t => !TYPE_ORDER.includes(t))
  ).filter(t => recs.some(r => r.rice_type === t));

  // First tile = All Types (the filtered total + status mix), then one per type.
  const panels = [{ key: "all", label: "All Types", subset: recs, hero: true }]
    .concat(types.map(t => ({ key: slug(t), label: t, subset: recs.filter(r => r.rice_type === t), hero: false })));

  const C = 2 * Math.PI * 46;  // donut circumference

  $("#typeCards").innerHTML = panels.map(p => {
    const subset = p.subset;
    const sc = countByStatus(subset);
    const order = statusOrder(subset).filter(s => sc[s]);
    const total = subset.length;
    const completed = sc["Completed"] || 0;
    const pct = total ? Math.round(100 * completed / total) : 0;
    const arc = C * pct / 100;

    // Legend: one row per present status, counts right-aligned in mono.
    const legend = order.map(s =>
      `<div class="tc-row"><span class="tc-dot" style="background:${statusColor(s)}"></span>` +
      `<span class="tc-label">${esc(s)}</span><span class="tc-num">${sc[s]}</span></div>`
    ).join("");
    const note = (pct === 0 && total > 0) ? `<div class="tc-note">No work started yet</div>` : "";

    // Slim stacked status bar.
    const bar = total ? order.map(s =>
      `<div style="width:${(100 * sc[s] / total).toFixed(2)}%;background:${statusColor(s)}"></div>`
    ).join("") : `<div style="width:100%;background:#dfe3e9"></div>`;

    // Completion-only donut (single clean arc).
    const donutArc = pct > 0
      ? `<circle cx="60" cy="60" r="46" fill="none" stroke="#3f9d6b" stroke-width="13" stroke-linecap="round" stroke-dasharray="${arc.toFixed(1)} ${(C - arc).toFixed(1)}" transform="rotate(-90 60 60)"/>`
      : "";
    const pctFill = pct > 0 ? "#1c2230" : "#a3abba";

    return `<div class="tcard${p.hero ? " hero" : ""}">
      <div class="tcard-head">
        <span class="tcard-name">${esc(p.label)}</span>
        <span class="tcard-badge${p.hero ? " hero" : ""}">${total}</span>
      </div>
      <div class="tcard-body">
        <svg class="tcard-donut" width="112" height="112" viewBox="0 0 120 120">
          <circle cx="60" cy="60" r="46" fill="none" stroke="#eef0f3" stroke-width="13"/>
          ${donutArc}
          <text x="60" y="57" text-anchor="middle" font-family="'IBM Plex Mono',monospace" font-size="26" font-weight="600" fill="${pctFill}">${pct}%</text>
          <text x="60" y="74" text-anchor="middle" font-family="'Public Sans',sans-serif" font-size="9" letter-spacing="1.5" fill="#a3abba">COMPLETE</text>
        </svg>
        <div class="tcard-legend">${legend}${note}</div>
      </div>
      <div class="tcard-bar">${bar}</div>
    </div>`;
  }).join("");
}
function countByStatus(recs) {
  const m = {}; recs.forEach(r => m[r.object_status] = (m[r.object_status] || 0) + 1); return m;
}

/* ============================================================
   RAW OBJECT STATUS
   ============================================================ */
function renderRawStatus(recs) {
  const statuses = [...new Set(recs.map(r => r.object_status || "—"))].sort();
  const el = $("#rawStatusCards");
  el.innerHTML = statuses.map(s => {
    const subset = recs.filter(r => (r.object_status || "—") === s);
    const bd = typeBreakdown(subset);
    const breaks = TYPE_ORDER.filter(t => bd[t]).map(t => `<span>${t}<b>${bd[t]}</b></span>`).join("")
      || `<span class="muted">none</span>`;
    const real = s !== "—";
    return `<div class="card ${real ? "clickable" : ""}" ${real ? `data-raw="${esc(s)}"` : ""}>
      <div class="card-label">${esc(s)}</div>
      <div class="card-value">${subset.length}</div>
      <div class="card-break">${breaks}</div>
    </div>`;
  }).join("");
  $$(".card.clickable", el).forEach(c => c.addEventListener("click", () => setQuick({ object_status: c.dataset.raw })));
}

/* ============================================================
   SPRINT SUMMARY
   ============================================================ */
function renderSprintSummary(recs) {
  const phases = State.data.timeline.slice();
  const groups = {};
  phases.forEach(p => groups[p.name] = []);
  groups["Unscheduled"] = [];
  recs.forEach(r => { if (groups[r.assigned_sprint]) groups[r.assigned_sprint].push(r); });

  const cards = phases.concat([{ name: "Unscheduled", type: "Unscheduled", status: "—", start: null, end: null }]);
  $("#sprintSummary").innerHTML = cards.map(p => {
    const list = groups[p.name] || [];
    const sc = countByStatus(list);
    const order = statusOrder(list);
    const seg = order.filter(s => sc[s]).map(s =>
      `<i style="width:${(100 * sc[s] / (list.length || 1)).toFixed(1)}%;background:${statusColor(s)}" title="${esc(s)}: ${sc[s]}"></i>`).join("");
    const breaks = order.filter(s => sc[s]).map(s =>
      `<span>${esc(s)}<b> ${sc[s]}</b></span>`).join("") || `<span class="muted">no objects</span>`;
    const dates = p.start ? (p.open_ended ? `${fmtDate(p.start)} onward` : `${fmtDate(p.start)} – ${fmtDate(p.end)}`) : "Not date-assigned";
    return `<div class="sprint-card phase-${esc(p.type)}" data-sprint="${esc(p.name)}">
      <div class="sprint-head"><h3>${esc(p.name)}</h3><span class="sprint-status">${esc(p.status)}</span></div>
      <div class="sprint-dates">${dates}</div>
      <div class="sprint-count">${list.length}</div>
      <div class="sprint-bar">${seg}</div>
      <div class="sprint-breaks">${breaks}</div>
    </div>`;
  }).join("");
  $$(".sprint-card", $("#sprintSummary")).forEach(c =>
    c.addEventListener("click", () => setQuick({ assigned_sprint: c.dataset.sprint })));
}

/* ============================================================
   AG GRID (Delivery Plan)
   ============================================================ */
function riceBadge(p) { return `<span class="rice-badge rice-${slug(p.value)}">${esc(p.value)}</span>`; }
function pctRenderer(p) {
  if (p.value == null) return `<span class="muted">—</span>`;
  return `${Math.round(p.value)}%`;
}

function gridColumns() {
  return [
    { headerName: "RICE ID", field: "rice_id", pinned: "left", width: 130, cellClass: "cell-obj" },
    {
      headerName: "Object Name", field: "object_name", pinned: "left", width: 230, cellClass: "cell-obj",
      tooltipField: "description"
    },
    { headerName: "Type", field: "rice_type", pinned: "left", width: 110, cellRenderer: riceBadge },
    { headerName: "Design Sprint", field: "design_sprint", pinned: "left", width: 120, cellClass: "crumb" },
    { headerName: "Dev Sprint", field: "dev_sprint", pinned: "left", width: 130, cellClass: "crumb",
      valueFormatter: p => (p.value || "").replace(/\n/g, " → ") },
    { headerName: "Complexity / Build Hrs", colId: "complexity_hours", field: "complexity", pinned: "left", width: 175,
      valueGetter: p => `${p.data.complexity || "—"} · ${p.data.build_hours == null ? "—" : fmtNum(p.data.build_hours) + " hrs"}`,
      cellRenderer: p => `${esc(p.data.complexity || "—")} <span class="muted">· ${p.data.build_hours == null ? "—" : esc(fmtNum(p.data.build_hours)) + " hrs"}</span>` },
    { headerName: "Functional Owner", field: "functional_owner", pinned: "left", width: 160 },
    // scrolling columns
    { headerName: "Object Status", field: "object_status", width: 150 },
    { headerName: "Module", field: "module", width: 110 },
    { headerName: "Workstream", field: "workstream", width: 120 },
    { headerName: "Accountable Org", field: "accountable_org", width: 130 },
    { headerName: "In Scope", field: "in_scope", width: 100 },
    { headerName: "Technical Owner", field: "technical_owner", width: 150 },
    { headerName: "Spec %", field: "spec_pct", width: 90, cellRenderer: pctRenderer },
    { headerName: "Dev %", field: "dev_pct", width: 90, cellRenderer: pctRenderer },
    { headerName: "Hrs Used", field: "hours_consumed", width: 100, type: "numericColumn",
      valueFormatter: p => p.value == null ? "—" : fmtNum(p.value) },
    { headerName: "Hrs Left", field: "hours_left", width: 100, type: "numericColumn",
      valueFormatter: p => p.value == null ? "—" : fmtNum(p.value) },
    { headerName: "Spec Plan", field: "spec_effective", width: 110, valueFormatter: p => fmtDate(p.value) },
    { headerName: "Spec Actual", field: "spec_actual", width: 110, valueFormatter: p => fmtDate(p.value) },
    { headerName: "Delivery", field: "delivery_date", width: 110, valueFormatter: p => fmtDate(p.value) },
    { headerName: "Assigned Sprint", field: "assigned_sprint", width: 140 },
    { headerName: "Source", field: "source_system", width: 150 },
    { headerName: "Target", field: "target_system", width: 150 },
  ];
}

// Row coloring by Object Status — mirrors statusColor() so rows match the
// pills, donuts, and matrix. Classes are styled in styles.css.
const _statusKey = p => ((p.data && p.data.object_status) || "").toLowerCase();
const ROW_RULES = {
  "rice-row-blocked":    p => _statusKey(p).includes("block"),
  "rice-row-delayed":    p => _statusKey(p).includes("delay"),
  "rice-row-complete":   p => /complete|done/.test(_statusKey(p)),
  "rice-row-notstarted": p => /not started|not-started/.test(_statusKey(p)),
  "rice-row-progress":   p => { const k = _statusKey(p); return /progress|draft|fut|review|wip/.test(k) && !/delay/.test(k); },
  "rice-row-other":      p => { const k = _statusKey(p); return !!k && !/block|delay|complete|done|not started|not-started|progress|draft|fut|review|wip/.test(k); },
};

// Live count of rows currently shown in the grid (reflects the global filter
// bar AND any AG Grid column filters the user applies).
function updateGridCount() {
  const api = State.gridApi;
  const el = $("#gridCount");
  if (!api || !el) return;
  let total = 0;
  api.forEachNode(() => total++);            // all rows fed to the grid
  const shown = api.getDisplayedRowCount();  // after every active filter
  el.innerHTML = (shown === total)
    ? `<b>${total}</b> ${total === 1 ? "row" : "rows"}`
    : `Showing <b>${shown}</b> of ${total} rows`;
}

function renderGrid(recs) {
  if (!State.gridApi) {
    const options = {
      columnDefs: gridColumns(),
      rowData: recs,
      defaultColDef: { sortable: true, resizable: true, filter: true, suppressMovable: false },
      enableCellTextSelection: true,
      tooltipShowDelay: 300,
      animateRows: false,
      rowHeight: 34,
      rowClassRules: ROW_RULES,
      onModelUpdated: updateGridCount,
    };
    State.gridApi = agGrid.createGrid($("#riceGrid"), options);
  } else {
    State.gridApi.setGridOption("rowData", recs);
  }
}

function initPlanToggle() {
  $$("#planToggle button").forEach(b => b.addEventListener("click", () => {
    $$("#planToggle button").forEach(x => x.classList.remove("active"));
    b.classList.add("active");
    State.planView = b.dataset.view;
    if (State.planView === "grid") {
      $("#gridWrap").classList.remove("hidden"); $("#ganttWrap").classList.add("hidden");
    } else {
      $("#ganttWrap").classList.remove("hidden"); $("#gridWrap").classList.add("hidden");
      renderGantt(State.filtered);
    }
  }));
}

/* ============================================================
   GANTT  (signature visual)
   ============================================================ */
function renderGantt(recs) {
  const rows = recs.filter(r => r.gantt_start && r.gantt_delivery)
    .sort((a, b) => {
      const ta = TYPE_ORDER.indexOf(a.rice_type), tb = TYPE_ORDER.indexOf(b.rice_type);
      const ra = ta === -1 ? TYPE_ORDER.length : ta, rb = tb === -1 ? TYPE_ORDER.length : tb;
      if (ra !== rb) return ra - rb;
      if (a.rice_type !== b.rice_type) return (a.rice_type || "").localeCompare(b.rice_type || "");
      return (a.rice_id || "").localeCompare(b.rice_id || "", undefined, { numeric: true });
    });
  const wrap = $("#ganttScroll");
  if (!rows.length) { wrap.innerHTML = `<div class="risk-empty">No objects with scheduling dates in the current filter.</div>`; return; }

  const tl = State.data.timeline;
  // Open-ended phases (e.g. Post Go-Live) have a far-future sentinel end date
  // that would blow up the axis scale — anchor the baseline on the last
  // bounded phase instead and let real object dates extend it if needed.
  const boundedTl = tl.filter(p => !p.open_ended);
  let minD = parseISO(tl[0].start), maxD = parseISO((boundedTl[boundedTl.length - 1] || tl[tl.length - 1]).end);
  rows.forEach(r => {
    const s = parseISO(r.gantt_start), d = parseISO(r.gantt_delivery), sp = parseISO(r.gantt_spec);
    [s, d, sp].forEach(x => { if (x && x < minD) minD = x; if (x && x > maxD) maxD = x; });
  });
  // pad a week each side
  minD = new Date(minD.getTime() - 6 * 864e5); maxD = new Date(maxD.getTime() + 6 * 864e5);
  const span = maxD - minD;
  const LABEL_W = 240, ROW_H = 26, TOP = 72, PX_W = 1180;
  const x = (d) => LABEL_W + ((d - minD) / span) * PX_W;
  const H = TOP + rows.length * ROW_H + 12;
  const W = LABEL_W + PX_W + 20;

  // phase shading rects + labels. Labels are center-anchored; when one would
  // overlap the previous label (e.g. a wide Cutover band next to the 1-day
  // Go-Live milestone) it drops to a second row so the text never collides.
  // Gap tiles are numerous and narrow, so they get shading + a hover title
  // only (no inline label) rather than crowding the two label rows.
  let bg = "", axis = "";
  const lastRight = { 1: -1e9, 2: -1e9 };
  tl.forEach(p => {
    const ps = parseISO(p.start), pe = p.open_ended ? maxD : parseISO(p.end);
    const x1 = x(ps), x2 = Math.max(x(pe), x1 + 2);
    bg += `<rect x="${x1}" y="${TOP}" width="${x2 - x1}" height="${rows.length * ROW_H}" fill="${phaseFillFor(p.name, p.type)}"><title>${esc(p.name)}</title></rect>`;
    if (p.type === "Gap") return;
    const cx = (x1 + x2) / 2;
    const halfW = (esc(p.name).length * 6.4 + 6) / 2;   // approx label half-width
    const row = (cx - halfW < lastRight[1] + 4) ? 2 : 1;
    lastRight[row] = cx + halfW;
    const ly = row === 1 ? 30 : 44;
    axis += `<text x="${cx}" y="${ly}" text-anchor="middle" font-size="11" font-weight="700" fill="var(--text-2)">${esc(p.name)}</text>`;
    axis += `<line x1="${x1}" y1="${TOP}" x2="${x1}" y2="${H - 12}" stroke="var(--border)" stroke-dasharray="3 3"/>`;
  });
  // month gridlines
  let m = new Date(minD.getFullYear(), minD.getMonth(), 1);
  while (m < maxD) {
    const mx = x(m);
    if (mx > LABEL_W) {
      axis += `<line x1="${mx}" y1="${TOP - 6}" x2="${mx}" y2="${H - 12}" stroke="var(--border)" opacity=".6"/>`;
      axis += `<text x="${mx + 3}" y="${TOP - 14}" font-size="9.5" fill="var(--muted)">${m.toLocaleDateString(undefined, { month: "short", year: "2-digit" })}</text>`;
    }
    m = new Date(m.getFullYear(), m.getMonth() + 1, 1);
  }
  // current week overlay
  const t0 = today();
  const cwStart = new Date(t0); cwStart.setDate(t0.getDate() - ((t0.getDay() + 6) % 7));
  const cwEnd = new Date(cwStart.getTime() + 7 * 864e5);
  if (cwEnd > minD && cwStart < maxD) {
    const cx1 = x(cwStart < minD ? minD : cwStart), cx2 = x(cwEnd > maxD ? maxD : cwEnd);
    bg += `<rect x="${cx1}" y="${TOP}" width="${Math.max(cx2 - cx1, 3)}" height="${rows.length * ROW_H}" fill="var(--ph-current)"/>`;
    // "now" gets its own top lane with a connector line down to the band, so it
    // never collides with the phase or month labels below it.
    axis += `<line x1="${cx1}" y1="14" x2="${cx1}" y2="${TOP}" stroke="var(--dl-orange)" stroke-width="1" stroke-dasharray="2 2" opacity=".7"/>`;
    axis += `<text x="${cx1 + 3}" y="11" font-size="9" font-weight="700" fill="var(--dl-orange)">now</text>`;
  }

  // bars
  let bars = "";
  rows.forEach((r, i) => {
    const y = TOP + i * ROW_H, cy = y + ROW_H / 2;
    const xs = x(parseISO(r.gantt_start)), xd = x(parseISO(r.gantt_delivery));
    const barColor = TYPE_COLOR[r.rice_type] || "#00A3E0";
    if (i % 2 === 0) bars += `<rect x="${LABEL_W}" y="${y}" width="${PX_W + 20}" height="${ROW_H}" fill="var(--surface-2)" opacity=".5"/>`;
    // label
    const nm = (r.object_name || "").length > 30 ? r.object_name.slice(0, 29) + "…" : r.object_name;
    bars += `<text x="8" y="${cy - 3}" font-size="10.5" font-weight="700" fill="var(--text)">${esc(r.rice_id)}</text>`;
    bars += `<text x="8" y="${cy + 9}" font-size="9.5" fill="var(--text-2)">${esc(nm)}</text>`;
    // build bar
    bars += `<rect x="${xs}" y="${cy - 4}" width="${Math.max(xd - xs, 2)}" height="8" rx="3" fill="${barColor}" opacity="${r.gantt_delivery_estimated ? .5 : .9}">
      <title>${esc(r.rice_id)} — build ${fmtDate(r.gantt_start)} → ${fmtDate(r.gantt_delivery)}${r.gantt_delivery_estimated ? " (est)" : ""}</title></rect>`;
    // spec diamond
    if (r.gantt_spec) {
      const sx = x(parseISO(r.gantt_spec));
      bars += `<rect x="${sx - 5}" y="${cy - 5}" width="10" height="10" transform="rotate(45 ${sx} ${cy})" fill="#046A38"><title>Spec complete ${fmtDate(r.gantt_spec)}</title></rect>`;
    }
    // delivery dot
    bars += `<circle cx="${xd}" cy="${cy}" r="5" fill="#ED8B00" stroke="var(--surface)" stroke-width="1.5"><title>Delivery ${fmtDate(r.gantt_delivery)}</title></circle>`;
  });
  // frozen label backdrop
  const labelBg = `<rect x="0" y="0" width="${LABEL_W}" height="${H}" fill="var(--surface)"/>
    <line x1="${LABEL_W}" y1="0" x2="${LABEL_W}" y2="${H}" stroke="var(--border-2)"/>`;

  wrap.innerHTML =
    `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" style="display:block">
      ${bg}${axis}
      <g>${bars}</g>
      ${labelBg}
      <g>${rows.map((r, i) => {
        const cy = TOP + i * ROW_H + ROW_H / 2;
        const nm = (r.object_name || "").length > 30 ? r.object_name.slice(0, 29) + "…" : r.object_name;
        return `<text x="8" y="${cy - 3}" font-size="10.5" font-weight="700" fill="var(--text)">${esc(r.rice_id)}</text>
                <text x="8" y="${cy + 9}" font-size="9.5" fill="var(--text-2)">${esc(nm)}</text>`;
      }).join("")}</g>
    </svg>`;
}

/* ============================================================
   RESOURCE CAPACITY HEATMAP
   ============================================================ */
function weekKey(d) {
  const x = new Date(d); x.setHours(0, 0, 0, 0);
  x.setDate(x.getDate() - ((x.getDay() + 6) % 7)); // Monday
  return ymd(x);
}
function renderCapacity(recs) {
  // Conversions are excluded from capacity planning (row + All Types totals).
  recs = recs.filter(r => r.rice_type !== "Conversion");
  const HPW = State.data.hours_per_dev_week || 45;
  const tl = State.data.timeline;
  // Open-ended phases (e.g. Post Go-Live) have a far-future sentinel end date
  // that would blow up the week-column range — anchor on the last bounded phase.
  const boundedTl = tl.filter(p => !p.open_ended);
  let minD = parseISO(tl[0].start), maxD = parseISO((boundedTl[boundedTl.length - 1] || tl[tl.length - 1]).end);

  // build per-week, per-type hours
  const data = {}; // weekKey -> {type -> {hours, objs[]}}
  const addHours = (wk, type, hrs, obj) => {
    data[wk] = data[wk] || {};
    data[wk][type] = data[wk][type] || { hours: 0, objs: [] };
    data[wk][type].hours += hrs;
    if (!data[wk][type].objs.includes(obj)) data[wk][type].objs.push(obj);
  };

  recs.forEach(r => {
    if (!r.build_hours || !r.gantt_start || !r.gantt_delivery) return;
    const s = parseISO(r.gantt_start), d = parseISO(r.gantt_delivery);
    if (s < minD) minD = s; if (d > maxD) maxD = d;
    const weeks = [];
    let w = new Date(s); w.setDate(w.getDate() - ((w.getDay() + 6) % 7));
    while (w <= d) { weeks.push(weekKey(w)); w.setDate(w.getDate() + 7); }
    if (!weeks.length) weeks.push(weekKey(s));
    const per = r.build_hours / weeks.length;
    weeks.forEach(wk => addHours(wk, r.rice_type, per, r.rice_id));
  });

  // week column list
  const cols = [];
  let w = new Date(minD); w.setDate(w.getDate() - ((w.getDay() + 6) % 7));
  while (w <= maxD) { cols.push(weekKey(w)); w.setDate(w.getDate() + 7); }
  const curWk = weekKey(today());

  const phaseOf = (wk) => {
    const d = parseISO(wk);
    for (const p of tl) { if (parseISO(p.start) <= d && d <= parseISO(p.end)) return p; }
    return null;
  };

  const rowsTypes = TYPE_ORDER.filter(t => recs.some(r => r.rice_type === t));
  const cls = (dev) => dev === 0 ? "heat-0" : dev <= 2 ? "heat-low" : dev <= 4 ? "heat-mid" : "heat-high";

  let head = `<tr><th class="lbl">RICE Type \\ Week</th>` + cols.map(c => {
    const ph = phaseOf(c);
    const bg = ph ? phaseFillFor(ph.name, ph.type) : "";
    const mark = c === curWk ? ` style="background:var(--ph-current)"` : (bg ? ` style="background:${bg}"` : "");
    return `<th${mark} title="${ph ? esc(ph.name) : "—"}">${fmtDate(c)}</th>`;
  }).join("") + `</tr>`;

  const typeRow = (type) => {
    let tds = "";
    cols.forEach(c => {
      const cell = (data[c] && data[c][type]) || { hours: 0, objs: [] };
      const dev = Math.ceil(cell.hours / HPW);
      tds += `<td class="heat-cell ${cls(dev)}" data-wk="${c}" data-type="${type}" title="${Math.round(cell.hours)} hrs · ${cell.objs.length} objects">${dev || ""}</td>`;
    });
    return `<tr><td class="lbl">${esc(type)}</td>${tds}</tr>`;
  };

  const totalRow = () => {
    let tds = "";
    cols.forEach(c => {
      let hrs = 0, objs = [];
      Object.values(data[c] || {}).forEach(o => { hrs += o.hours; objs = objs.concat(o.objs); });
      const dev = Math.ceil(hrs / HPW);
      tds += `<td class="heat-cell ${cls(dev)}" data-wk="${c}" data-type="" title="${Math.round(hrs)} hrs · ${new Set(objs).size} objects"><b>${dev || ""}</b></td>`;
    });
    return `<tr style="font-weight:800"><td class="lbl">All Types</td>${tds}</tr>`;
  };

  $("#capacityHeat").innerHTML =
    `<div class="cap-legend">
       <span><i style="background:var(--surface);border:1px solid var(--border)"></i> 0</span>
       <span><i style="background:rgba(134,188,37,.45)"></i> 1–2</span>
       <span><i style="background:rgba(237,139,0,.55)"></i> 3–4</span>
       <span><i style="background:rgba(218,41,28,.6)"></i> 5+</span>
       <span class="muted">developers needed = ⌈weekly build hours ÷ ${HPW}⌉ · click a cell for objects</span>
     </div>
     <div class="cap-scroll">
       <table class="heat"><thead>${head}</thead>
         <tbody>${rowsTypes.map(typeRow).join("")}${totalRow()}</tbody></table>
     </div>`;

  $$(".heat-cell", $("#capacityHeat")).forEach(td => td.addEventListener("click", () => {
    const wk = td.dataset.wk, type = td.dataset.type;
    const cell = type ? (data[wk] && data[wk][type]) : null;
    let ids;
    if (type) ids = cell ? cell.objs : [];
    else { ids = []; Object.values(data[wk] || {}).forEach(o => ids = ids.concat(o.objs)); ids = [...new Set(ids)]; }
    const list = recs.filter(r => ids.includes(r.rice_id));
    openModal(`Week of ${fmtDate(wk)}${type ? " · " + type : ""} — ${list.length} objects`, list);
  }));
}

/* ============================================================
   RISK
   ============================================================ */
function renderRisk(recs) {
  const lean = recs.filter(r => r.lean_spec_risk);
  const build = recs.filter(r => r.build_risk);
  $("#leanRisk").innerHTML = riskItems(lean, "spec");
  $("#buildRisk").innerHTML = riskItems(build, "build");
  $$("#leanRisk .risk-item, #buildRisk .risk-item").forEach(el =>
    el.addEventListener("click", () => {
      const r = recs.find(x => x.rice_id === el.dataset.id);
      if (r) openModal(r.rice_id + " — " + r.object_name, [r], true);
    }));
}
function riskItems(list, kind) {
  if (!list.length) return `<div class="risk-empty">No objects flagged. ✓</div>`;
  return list.sort((a, b) => {
    const da = parseISO(kind === "spec" ? a.spec_effective : a.delivery_date) || new Date(8640e12);
    const db = parseISO(kind === "spec" ? b.spec_effective : b.delivery_date) || new Date(8640e12);
    return da - db;
  }).map(r => {
    const blocked = (r.object_status || "").toLowerCase().includes("block");
    const date = kind === "spec" ? r.spec_effective : r.delivery_date;
    const stat = kind === "spec" ? (r.fspec_status || r.object_status) : (r.dev_status || r.object_status);
    const pct = kind === "spec" ? r.spec_pct : r.dev_pct;
    return `<div class="risk-item ${blocked ? "blocked" : ""}" data-id="${esc(r.rice_id)}">
      <div class="ri-top"><span class="ri-name">${esc(r.object_name)}</span>
        <span class="st-pill" style="background:${statusColor(r.object_status)}">${esc(r.object_status || "—")}</span></div>
      <div class="ri-meta">
        <span class="ri-id">${esc(r.rice_id)}</span>
        <span>${esc(r.rice_type)}</span>
        <span>${kind === "spec" ? "Spec" : "Delivery"}: ${fmtDate(date)}</span>
        <span>${esc(stat)}${pct != null ? " · " + Math.round(pct) + "%" : ""}</span>
        <span>${esc(r.functional_owner || r.technical_owner || "—")}</span>
      </div></div>`;
  }).join("");
}

/* ============================================================
   OWNER FOCUS — by Functional Owner: delayed + lean spec due this week
   ============================================================ */
// Lean spec date = Revised if present, else Planned.
const leanSpecDate = (r) => r.spec_revised || r.spec_planned;

function renderOwnerFocus(recs) {
  const wkStart = parseISO(weekKey(today()));            // Monday of current week
  const wkEnd = new Date(wkStart); wkEnd.setDate(wkEnd.getDate() + 7);

  const isDelayed = (r) => (r.object_status || "").toLowerCase().includes("delay");
  const isDueThisWeek = (r) => {
    const d = parseISO(leanSpecDate(r));
    return d && d >= wkStart && d < wkEnd;
  };

  const matches = [];
  recs.forEach(r => {
    const reasons = [];
    if (isDelayed(r)) reasons.push("Delayed");
    if (isDueThisWeek(r)) reasons.push("Spec due this week");
    if (reasons.length) matches.push({ r, reasons });
  });

  const el = $("#ownerFocus");
  if (!matches.length) {
    el.innerHTML = `<div class="risk-empty">No delayed or lean-spec-due-this-week objects in the current filter. ✓</div>`;
    return;
  }

  const groups = {};
  matches.forEach(m => {
    const owner = m.r.functional_owner || "— Unassigned —";
    (groups[owner] = groups[owner] || []).push(m);
  });
  const owners = Object.keys(groups).sort((a, b) => a.localeCompare(b));

  el.innerHTML = owners.map(owner => {
    const items = groups[owner].sort((a, b) =>
      (parseISO(leanSpecDate(a.r)) || new Date(8640e12)) - (parseISO(leanSpecDate(b.r)) || new Date(8640e12)));
    const delayedN = items.filter(i => i.reasons.includes("Delayed")).length;
    const dueN = items.filter(i => i.reasons.includes("Spec due this week")).length;
    return `<div class="owner-card">
      <div class="owner-head">
        <span class="owner-name">${esc(owner)}</span>
        <span class="owner-counts">
          ${delayedN ? `<span class="oc-badge delayed">${delayedN} delayed</span>` : ""}
          ${dueN ? `<span class="oc-badge due">${dueN} due this week</span>` : ""}
        </span>
      </div>
      <div class="owner-items">
        ${items.map(({ r, reasons }) => {
          const sd = leanSpecDate(r);
          return `<div class="risk-item ${isDelayed(r) ? "blocked" : ""}" data-id="${esc(r.rice_id)}">
            <div class="ri-top">
              <span class="ri-name">${esc(r.object_name)}</span>
              <span class="reason-tags">${reasons.map(rs =>
                `<span class="oc-badge ${rs === "Delayed" ? "delayed" : "due"}">${rs}</span>`).join("")}</span>
            </div>
            <div class="ri-meta">
              <span class="ri-id">${esc(r.rice_id)}</span>
              <span>${esc(r.rice_type)}</span>
              <span class="st-pill" style="background:${statusColor(r.object_status)}">${esc(r.object_status || "—")}</span>
              <span>Spec: ${fmtDate(sd)}${r.spec_revised ? " (revised)" : (r.spec_planned ? " (planned)" : "")}</span>
            </div>
            ${r.description ? `<div class="ri-desc">${esc(r.description)}</div>` : ""}
          </div>`;
        }).join("")}
      </div>
    </div>`;
  }).join("");

  $$("#ownerFocus .risk-item").forEach(item => item.addEventListener("click", () => {
    const r = recs.find(x => x.rice_id === item.dataset.id);
    if (r) openModal(r.rice_id + " — " + r.object_name, [r], true);
  }));
}

/* ============================================================
   MATRIX
   ============================================================ */
function renderMatrix(recs) {
  const types = TYPE_ORDER.filter(t => recs.some(r => r.rice_type === t))
    .concat([...new Set(recs.map(r => r.rice_type))].filter(t => !TYPE_ORDER.includes(t)));
  const rawStatuses = [...new Set(recs.map(r => r.object_status).filter(Boolean))].sort();
  buildMatrix($("#matrixRaw"), recs, types, rawStatuses, "object_status");
}
function buildMatrix(container, recs, rows, cols, field) {
  const count = (t, c) => recs.filter(r => r.rice_type === t && r[field] === c).length;
  let head = `<tr><th class="lbl">RICE Type</th>` +
    cols.map(c => `<th>${esc(c)}</th>`).join("") + `<th class="total-col">Total</th></tr>`;
  let body = rows.map(t => {
    let rowTotal = 0;
    const tds = cols.map(c => {
      const n = count(t, c); rowTotal += n;
      const color = field === "object_status" ? statusColor(c) : "";
      return n ? `<td class="cell" data-type="${esc(t)}" data-col="${esc(c)}" data-field="${field}" style="${color ? `color:${color}` : ""}">${n}</td>`
        : `<td class="zero">·</td>`;
    }).join("");
    return `<tr><td class="lbl"><span class="rice-badge rice-${slug(t)}">${esc(t)}</span></td>${tds}<td class="total-col">${rowTotal}</td></tr>`;
  }).join("");
  // totals row
  const totals = cols.map(c => recs.filter(r => r[field] === c).length);
  body += `<tr class="total-row"><td class="lbl">Total</td>${totals.map(n => `<td>${n}</td>`).join("")}<td class="total-col">${recs.length}</td></tr>`;
  container.innerHTML = `<table class="matrix"><thead>${head}</thead><tbody>${body}</tbody></table>`;
  $$(".cell", container).forEach(td => td.addEventListener("click", () => {
    const q = { rice_type: td.dataset.type };
    q[td.dataset.field] = td.dataset.col;
    setQuick(q);
  }));
}

/* ============================================================
   DATA QUALITY (recomputed on filtered set)
   ============================================================ */
function renderDataQuality(recs) {
  const total = recs.length || 1;
  const checks = {
    "Sprint": r => !r.design_sprint && !r.dev_sprint,
    "Spec Date": r => !(r.spec_planned || r.spec_revised || r.spec_actual),
    "Build Hours": r => r.build_hours == null,
    "Delivery Date": r => !(r.build_planned || r.build_actual || r.delivery_date),
    "Functional Owner": r => !r.functional_owner,
    "Object Status": r => !r.object_status,
    "RICE Type": r => !r.rice_type || r.rice_type === "Unspecified",
    "Module": r => !r.module,
  };
  const items = Object.entries(checks).map(([field, fn]) => {
    const miss = recs.filter(fn).length;
    return { field, miss, pct: Math.round(1000 * miss / total) / 10 };
  });
  $("#dataQuality").innerHTML = items.map(it =>
    `<div class="dq-item">
      <div class="dq-top"><span class="dq-field">${it.field}</span>
        <span class="dq-pct" style="color:${it.pct > 40 ? "var(--dl-red)" : it.pct > 15 ? "var(--dl-orange)" : "var(--dl-green-dark)"}">${it.pct}%</span></div>
      <div class="dq-bar"><i style="width:${it.pct}%;background:${it.pct > 40 ? "var(--dl-red)" : it.pct > 15 ? "var(--dl-orange)" : "var(--dl-green)"}"></i></div>
      <div class="dq-sub">${it.miss} of ${total} objects missing</div>
    </div>`).join("");
}

/* ============================================================
   MODAL
   ============================================================ */
function initModal() {
  $("#modalClose").addEventListener("click", closeModal);
  $("#modal").addEventListener("click", e => { if (e.target.id === "modal") closeModal(); });
  document.addEventListener("keydown", e => { if (e.key === "Escape") closeModal(); });
}
function openModal(title, list, detailed) {
  $("#modalTitle").textContent = title;
  if (detailed && list.length === 1) {
    const r = list[0];
    const f = (k, v) => `<div class="mi"><b>${k}</b> ${esc(v == null || v === "" ? "—" : v)}</div>`;
    $("#modalBody").innerHTML = `<div class="modal-list">
      ${f("Type", r.rice_type)} ${f("Module", r.module)} ${f("Workstream", r.workstream)}
      ${f("Object Status", r.object_status)}
      ${f("Description", r.description)}
      ${f("Functional Owner", r.functional_owner)} ${f("Technical Owner", r.technical_owner)}
      ${f("Complexity", r.complexity)} ${f("Build Hours", r.build_hours)}
      ${f("Spec %", r.spec_pct)} ${f("Dev %", r.dev_pct)}
      ${f("Spec (planned)", fmtDate(r.spec_effective))} ${f("Spec (actual)", fmtDate(r.spec_actual))}
      ${f("Delivery", fmtDate(r.delivery_date))} ${f("Assigned Sprint", r.assigned_sprint)}
      ${f("Source → Target", (r.source_system || "—") + " → " + (r.target_system || "—"))}
      ${r.comments ? f("Comments", r.comments) : ""}
    </div>`;
  } else if (list.length) {
    $("#modalBody").innerHTML = `<div class="modal-list">` + list.map(r =>
      `<div class="mi"><b>${esc(r.rice_id)}</b> ${esc(r.object_name)} —
        <span class="st-pill" style="background:${statusColor(r.object_status)};font-size:10px">${esc(r.object_status || "—")}</span>
        <span class="muted"> ${esc(r.rice_type)} · ${r.build_hours == null ? "—" : r.build_hours + " hrs"} · ${esc(r.functional_owner || "no owner")}</span></div>`
    ).join("") + `</div>`;
  } else {
    $("#modalBody").innerHTML = `<div class="risk-empty">No objects.</div>`;
  }
  $("#modal").classList.remove("hidden");
}
function closeModal() { $("#modal").classList.add("hidden"); }

/* ============================================================
   CSV EXPORT (filtered set)
   ============================================================ */
function exportCSV() {
  const recs = State.filtered;
  if (!recs.length) { toast("Nothing to export."); return; }
  const cols = ["rice_id", "rice_type", "object_name", "module", "workstream", "accountable_org", "in_scope",
    "object_status", "complexity", "build_hours", "spec_pct", "dev_pct",
    "hours_consumed", "hours_left", "functional_owner", "technical_owner", "design_sprint", "dev_sprint",
    "spec_effective", "spec_actual", "delivery_date", "assigned_sprint", "lean_spec_risk", "build_risk",
    "source_system", "target_system"];
  const head = cols.join(",");
  const esc2 = v => { v = v == null ? "" : String(v); return /[",\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v; };
  const lines = recs.map(r => cols.map(c => esc2(r[c])).join(","));
  const blob = new Blob([head + "\n" + lines.join("\n")], { type: "text/csv" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "rice_tracker_filtered.csv";
  a.click(); URL.revokeObjectURL(a.href);
  toast(`Exported ${recs.length} rows.`);
}
