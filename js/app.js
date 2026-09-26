'use strict';
/* ============================================================
   Storage Sizer — storage capacity planner for VMware refreshes
   100% client-side. No uploads, no storage, no network calls
   carrying customer data. Everything lives in page memory.
   ============================================================ */

/* ================= Helpers ================= */
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fmtInt = (n) => Math.round(n || 0).toLocaleString('en-US');
const fmt1 = (n) => (n == null || !isFinite(n)) ? '—' : (Math.round(n * 10) / 10).toLocaleString('en-US');
const fmtTB = (tb) => (tb == null || !isFinite(tb)) ? '—' : (tb >= 100 ? fmtInt(tb) : fmt1(tb)) + ' TB';
function parseNum(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const n = parseFloat(String(v).replace(/,/g, ''));
  return isFinite(n) ? n : 0;
}
function pick(row, aliases) {
  for (const a of aliases) {
    if (row[a] !== undefined && row[a] !== '' && row[a] != null) return row[a];
  }
  return '';
}
function lcg(seed) {
  let s = seed >>> 0;
  return () => { s = (s * 1664525 + 1013904223) >>> 0; return s / 4294967296; };
}

/* ================= RVTools column aliases (modern + legacy) ================= */
const COL = {
  vm: {
    name: ['VM'], power: ['Powerstate', 'Power state'], template: ['Template'],
    provMB: ['Provisioned MB', 'Provisioned MiB'],
    usedMB: ['In Use MB', 'In use MiB', 'In Use MiB', 'Used MB', 'Used MiB'],
    dc: ['Datacenter', 'DC'], cluster: ['Cluster'],
  },
  ds: {
    name: ['Name', 'Datastore'], type: ['Type'],
    capMB: ['Capacity MB', 'Capacity MiB'],
    provMB: ['Provisioned MB', 'Provisioned MiB'],
    usedMB: ['In Use MB', 'In Use MiB', 'In use MiB', 'Used MB'],
  },
  snap: { sizeMB: ['Size MB', 'Size MiB', 'Snapshot Size MB'] },
  disk: { thin: ['Thin', 'Thin provisioned', 'Thin Provisioned'] },
};
const TAB_NAMES = {
  vInfo: ['vInfo'], vDatastore: ['vDatastore'], vSnapshot: ['vSnapshot'], vDisk: ['vDisk'],
};
function findSheet(wb, candidates) {
  const lower = wb.SheetNames.map((s) => s.toLowerCase());
  for (const c of candidates) {
    const i = lower.indexOf(c.toLowerCase());
    if (i >= 0) return wb.SheetNames[i];
  }
  return null;
}
function sheetRows(ws) {
  if (!ws) return [];
  const asArray = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
  if (!asArray.length) return [];
  let headerIdx = 0;
  for (let i = 0; i < Math.min(asArray.length, 5); i++) {
    const joined = asArray[i].join(' ').toLowerCase();
    if (/(^|\s)(vm|host|name|cluster|datastore|snapshot|disk)(s|\s|$)/.test(' ' + joined + ' ')) { headerIdx = i; break; }
  }
  return XLSX.utils.sheet_to_json(ws, { defval: '', raw: true, range: headerIdx });
}
function normPower(v) {
  const s = String(v || '').toLowerCase().replace(/[\s_-]/g, '');
  if (s === 'poweredon') return 'on';
  if (s === 'poweredoff') return 'off';
  if (s === 'suspended') return 'suspended';
  return 'unknown';
}
function isTruthy(v) { return /true|yes|1/i.test(String(v).trim()); }

/* ================= Demand model ================= */
const MB_TO_TB = 1 / 1048576;
function clKey(dc, cluster) { return dc + ' / ' + (cluster || 'Standalone'); }

function buildPools(vms, datastores, snapshots, disks) {
  const byCl = {};
  // Storage counts regardless of power state — only templates are excluded.
  vms.filter((v) => !v.template).forEach((v) => {
    const k = clKey(v.dc, v.cluster);
    if (!byCl[k]) byCl[k] = { name: k, vms: 0, provTB: 0, usedTB: 0 };
    const p = byCl[k];
    p.vms++;
    p.provTB += (v.provMB || 0) * MB_TO_TB;
    p.usedTB += (v.usedMB || 0) * MB_TO_TB;
  });
  let skipped = 0;
  const pools = Object.keys(byCl).sort().map((k, i) => {
    const p = byCl[k];
    if (p.vms === 0 || (p.provTB <= 0 && p.usedTB <= 0)) { skipped++; return null; }
    return {
      id: 'p' + i, name: k, vms: p.vms,
      provisionedTB: Math.round(p.provTB * 10) / 10,
      usedTB: Math.round(p.usedTB * 10) / 10,
      usableTB: null, // RVTools pools have no per-pool usable capacity
      source: 'rvtools',
    };
  }).filter(Boolean);
  const validation = {
    dsCapTB: datastores.reduce((a, d) => a + d.capMB, 0) * MB_TO_TB,
    dsUsedTB: datastores.reduce((a, d) => a + d.usedMB, 0) * MB_TO_TB,
    dsCount: datastores.length,
    snapTB: snapshots.reduce((a, s) => a + s, 0) * MB_TO_TB,
    disksN: disks.length,
    disksThin: disks.filter(Boolean).length,
  };
  return { pools, skipped, validation };
}

/* ================= Demo data (synthetic, in-memory only) ================= */
function genDemoPools() {
  const rnd = lcg(20260926);
  const mk = (name, vms, prov, used, usable) => ({
    id: 'd' + Math.floor(rnd() * 1e6), name, vms,
    provisionedTB: prov, usedTB: used, usableTB: usable, source: 'demo',
  });
  return [
    mk('DC-East / Prod-General', 184, 68, 46, 120),
    mk('DC-East / Prod-DB (snapshot-heavy)', 58, 44, 31, 60),
    mk('DC-West / ROBO-Edge', 26, 11, 8, 24),
  ];
}
function genDemoValidation() {
  return { dsCapTB: 204, dsUsedTB: 96, dsCount: 9, snapTB: 6.4, disksN: 268, disksThin: 166 };
}

/* ================= Sizing math ================= */
const PROT = {
  none:    { label: 'None',            factor: 1.0 },
  mirror2: { label: '2-way mirror',    factor: 2.0 },
  mirror3: { label: '3-way mirror',    factor: 3.0 },
  raid5:   { label: 'RAID-5 / 4+1',    factor: 1.33 },
  raid6:   { label: 'RAID-6 / 6+2',    factor: 1.5 },
  ec83:    { label: 'EC 8+3',          factor: 1.375 },
};

function defaultCfg() {
  return {
    basis: 'used',          // 'used' | 'provisioned'
    reduction: 1.8,         // data reduction ratio
    protection: 'mirror2',  // key into PROT
    remote: 'none',         // 'none' | 'async'
    remoteCopies: 1,        // 1–2, only when remote === 'async'
    snapCopies: 1,          // 0–5
    snapPct: 25,            // % of base per copy, 5–100
    overhead: 10,           // % system overhead, 0–30
    spare: 15,              // % free-space headroom, 0–40
    useGlobal: true,
    ownGrowth: 20,          // pool's own growth rate (in global mode units)
  };
}
function defaultGlobal() { return { horizon: 3, mode: 'compound', value: 20 }; }

function effGrowth(cfg, G) {
  return cfg.useGlobal ? G.value : (cfg.ownGrowth != null && cfg.ownGrowth !== '' ? parseNum(cfg.ownGrowth) : G.value);
}
function growthLabel(g, mode) {
  return mode === 'compound' ? g + '%/yr compound' : fmt1(g) + ' TB/yr static';
}

/* Per-pool, per-year chain. y = 0..horizon. */
function sizePoolYear(pool, cfg, G, y) {
  const basisTB = cfg.basis === 'provisioned' ? pool.provisionedTB : pool.usedTB;
  const g = effGrowth(cfg, G);
  const logical = G.mode === 'compound'
    ? basisTB * Math.pow(1 + g / 100, y)
    : Math.max(0, basisTB + y * g);
  const reduced = cfg.reduction > 0 ? logical / cfg.reduction : logical;
  const snapFactor = 1 + cfg.snapCopies * (cfg.snapPct / 100);
  const withSnaps = reduced * snapFactor;
  const localF = PROT[cfg.protection] ? PROT[cfg.protection].factor : 1.0;
  const withLocal = withSnaps * localF;
  const remoteF = cfg.remote === 'async' ? (1 + Math.min(Math.max(cfg.remoteCopies, 1), 2)) : 1;
  const withRemote = withLocal * remoteF;
  const withOH = withRemote * (1 + cfg.overhead / 100);
  const raw = withOH / Math.max(1 - cfg.spare / 100, 0.5);
  return { basisTB, g, logical, reduced, snapFactor, withSnaps, localF, withLocal, remoteF, withRemote, withOH, raw };
}

/* ================= App state ================= */
const APP = { pools: [], cfgs: {}, global: defaultGlobal(), source: null, fileName: null, validation: null, results: null };

function setStatus(msg) { const s = $('parseStatus'); s.hidden = false; s.innerHTML = msg; }
function showError(msg) { const e = $('fileError'); e.hidden = false; e.innerHTML = msg; }
function clearMsgs() { $('fileError').hidden = true; $('parseStatus').hidden = true; }

function setStep(n) {
  $('stepData').hidden = n !== 1;
  $('stepConfig').hidden = n !== 2;
  $('stepResults').hidden = n !== 3;
  document.querySelectorAll('#stepper .step').forEach((el) => {
    const s = parseInt(el.dataset.step);
    el.classList.toggle('active', s === n);
    el.classList.toggle('done', s < n);
  });
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function startWizard() {
  $('landing').hidden = true;
  $('wizard').hidden = false;
  setStep(1);
  window.scrollTo({ top: 0 });
}

/* ================= Data loading ================= */
function parseWorkbook(wb, fileName) {
  const vms = sheetRows(wb.Sheets[findSheet(wb, TAB_NAMES.vInfo)] || {}).map((r) => ({
    power: normPower(pick(r, COL.vm.power)),
    template: isTruthy(pick(r, COL.vm.template)),
    provMB: parseNum(pick(r, COL.vm.provMB)),
    usedMB: parseNum(pick(r, COL.vm.usedMB)),
    dc: String(pick(r, COL.vm.dc) || '—'),
    cluster: String(pick(r, COL.vm.cluster) || ''),
  }));
  if (!vms.length) {
    showError('<strong>No VM rows found.</strong> Make sure this is an RVTools export with a <code>vInfo</code> tab (File → Export all to Excel).');
    return;
  }
  const dsWs = findSheet(wb, TAB_NAMES.vDatastore);
  const datastores = dsWs ? sheetRows(wb.Sheets[dsWs]).map((r) => ({
    name: String(pick(r, COL.ds.name) || 'unknown'),
    type: String(pick(r, COL.ds.type) || '—'),
    capMB: parseNum(pick(r, COL.ds.capMB)),
    provMB: parseNum(pick(r, COL.ds.provMB)),
    usedMB: parseNum(pick(r, COL.ds.usedMB)),
  })).filter((d) => d.capMB > 0 || d.usedMB > 0) : [];
  const snWs = findSheet(wb, TAB_NAMES.vSnapshot);
  const snapshots = snWs ? sheetRows(wb.Sheets[snWs]).map((r) => parseNum(pick(r, COL.snap.sizeMB))) : [];
  const dkWs = findSheet(wb, TAB_NAMES.vDisk);
  const disks = dkWs ? sheetRows(wb.Sheets[dkWs]).map((r) => isTruthy(pick(r, COL.disk.thin))) : [];

  const { pools, skipped, validation } = buildPools(vms, datastores, snapshots, disks);
  if (!pools.length) { showError('<strong>No pools with storage demand found</strong> in this export.'); return; }
  APP.pools = pools;
  APP.validation = validation;
  APP.source = 'rvtools';
  APP.fileName = fileName;
  setStatus('Parsed <strong>' + fmtInt(vms.length) + '</strong> VMs → <strong>' + pools.length + '</strong> storage pools.' +
    (validation.dsCount ? ' <strong>' + validation.dsCount + '</strong> datastores for validation.' : ' No vDatastore tab — capacity validation unavailable.') +
    (skipped ? ' (' + skipped + ' empty skipped)' : ''));
  renderInventory();
}

async function handleFile(file) {
  clearMsgs();
  if (!file) return;
  if (!/\.(xlsx|xls)$/i.test(file.name)) { showError('<strong>Not a spreadsheet.</strong> Drop the <code>.xlsx</code> from RVTools (File → Export all to Excel).'); return; }
  setStatus('Reading <strong>' + esc(file.name) + '</strong>…');
  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    parseWorkbook(wb, file.name);
  } catch (err) {
    showError('<strong>Could not parse that file.</strong> ' + esc(err.message || 'Unknown error.'));
  }
}

function loadDemo() {
  clearMsgs();
  APP.pools = genDemoPools();
  APP.validation = genDemoValidation();
  APP.source = 'demo';
  APP.fileName = 'demo-environment';
  setStatus('Loaded <strong>demo environment</strong>: 3 synthetic pools with different profiles (general, snapshot-heavy DB, ROBO edge) plus synthetic datastore context. Generated in your browser — nothing uploaded.');
  renderInventory();
}

/* ---- Manual entry ---- */
function manualRowHTML() {
  return '<tr>' +
    '<td><input data-f="name" placeholder="e.g. Prod-General" value=""></td>' +
    '<td><input class="num" data-f="vms" type="number" min="0" value=""></td>' +
    '<td><input class="num" data-f="prov" type="number" min="0" step="0.1" value=""></td>' +
    '<td><input class="num" data-f="used" type="number" min="0" step="0.1" value=""></td>' +
    '<td><input class="num" data-f="usable" type="number" min="0" step="0.1" value=""></td>' +
    '<td><button class="btn ghost" data-del style="padding:6px 10px">✕</button></td></tr>';
}
function wireManualEditor() {
  const body = $('manualBody');
  if (!body.children.length) { body.innerHTML = manualRowHTML() + manualRowHTML(); }
  body.querySelectorAll('[data-del]').forEach((b) => {
    b.onclick = () => { if (body.children.length > 1) b.closest('tr').remove(); };
  });
}
function applyManual() {
  clearMsgs();
  const rows = [...$('manualBody').querySelectorAll('tr')];
  const pools = [];
  rows.forEach((tr, i) => {
    const g = (f) => tr.querySelector('[data-f="' + f + '"]').value.trim();
    const name = g('name') || ('Pool ' + (i + 1));
    const vms = parseNum(g('vms')), prov = parseNum(g('prov')), used = parseNum(g('used'));
    if (!vms && !prov && !used) return; // skip empty rows
    const usable = parseNum(g('usable')) || null;
    pools.push({
      id: 'm' + i, name, vms,
      provisionedTB: Math.round(prov * 10) / 10,
      usedTB: Math.round(used * 10) / 10,
      usableTB: usable, source: 'manual',
    });
  });
  if (!pools.length) { showError('<strong>No usable rows.</strong> Fill in at least pool name, VMs, provisioned, and used TB for one pool.'); return; }
  APP.pools = pools;
  APP.validation = null;
  APP.source = 'manual';
  APP.fileName = 'manual-entry';
  setStatus('Using <strong>' + pools.length + '</strong> manually entered pool' + (pools.length > 1 ? 's' : '') + '. Datastore validation is unavailable for manual entry unless you fill in current usable TB.');
  renderInventory();
}

function renderInventory() {
  const w = $('inventoryWrap');
  w.hidden = false;
  $('inventoryMeta').textContent = APP.source === 'rvtools' ? 'from ' + APP.fileName : APP.source === 'demo' ? 'synthetic demo data' : 'manual entry';
  $('inventoryBody').innerHTML = APP.pools.map((p) =>
    '<tr><td><strong>' + esc(p.name) + '</strong></td>' +
    '<td class="num">' + fmtInt(p.vms) + '</td>' +
    '<td class="num">' + fmtTB(p.provisionedTB) + '</td>' +
    '<td class="num">' + fmtTB(p.usedTB) + '</td>' +
    '<td class="num">' + (p.usableTB != null ? fmtTB(p.usableTB) : '—') + '</td></tr>'
  ).join('');
  const v = APP.validation;
  const bits = [];
  if (v && v.dsCount) {
    const pct = v.dsCapTB > 0 ? Math.round(v.dsUsedTB / v.dsCapTB * 100) : 0;
    bits.push('<strong>vDatastore totals:</strong> ' + fmtTB(v.dsCapTB) + ' capacity · ' + fmtTB(v.dsUsedTB) + ' consumed (' + pct + '% used)');
  }
  if (v && v.snapTB > 0.05) bits.push('<strong>vSnapshot:</strong> ' + fmtTB(v.snapTB) + ' in snapshots today');
  if (v && v.disksN) {
    const thinPct = Math.round(v.disksThin / v.disksN * 100);
    bits.push('<strong>vDisk:</strong> ' + fmtInt(v.disksThin) + ' of ' + fmtInt(v.disksN) + ' VMDKs thin (' + thinPct + '%)');
  }
  $('validationNote').innerHTML = bits.length ? bits.join(' &nbsp;·&nbsp; ') : 'No datastore/snapshot/disk context detected — validation findings will be limited.';
  w.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/* ================= Step 2: per-pool configuration ================= */
function getCfg(id) { if (!APP.cfgs[id]) APP.cfgs[id] = defaultCfg(); return APP.cfgs[id]; }

function segHTML(seg, opts, cur) {
  return '<div class="seg" data-seg="' + seg + '">' + opts.map((o) =>
    '<button type="button" data-val="' + o[0] + '"' + (o[0] === cur ? ' class="active"' : '') + (o[2] ? ' disabled title="' + esc(o[2]) + '"' : '') + '>' + o[1] + '</button>'
  ).join('') + '</div>';
}

function renderGlobalGrowth() {
  const G = APP.global;
  $('gHorizon').value = G.horizon;
  document.querySelector('[data-gl="horizon"]').textContent = G.horizon + ' year' + (G.horizon > 1 ? 's' : '');
  $('gMode').querySelectorAll('button').forEach((b) => b.classList.toggle('active', b.dataset.val === G.mode));
  $('gValue').value = G.value;
  $('gValueNum').textContent = G.value;
  $('gValueUnit').textContent = G.mode === 'compound' ? '%/yr' : 'TB/yr';
}

function renderConfig() {
  renderGlobalGrowth();
  const wrap = $('poolCards');
  wrap.innerHTML = APP.pools.map((p, i) => {
    const cfg = getCfg(p.id);
    const protOpts = Object.keys(PROT).map((k) =>
      '<option value="' + k + '"' + (k === cfg.protection ? ' selected' : '') + '>' + esc(PROT[k].label) + ' (' + PROT[k].factor + '×)</option>').join('');
    return '<div class="ccard' + (i === 0 ? ' open' : '') + '" data-id="' + p.id + '">' +
      '<div class="ccard-head">' +
        '<div><div class="ccard-title">' + esc(p.name) + '</div>' +
        '<div class="ccard-sub">' + fmtInt(p.vms) + ' VMs · ' + fmtTB(p.provisionedTB) + ' provisioned · ' + fmtTB(p.usedTB) + ' used' + (p.usableTB != null ? ' · ' + fmtTB(p.usableTB) + ' usable today' : '') + '</div></div>' +
        '<div class="ccard-preview"><div class="hosts"><span data-pv="raw">—</span> <small>TB raw</small></div><div class="binding" data-pv="sub"></div></div>' +
        '<div class="ccard-toggle">▾</div>' +
      '</div>' +
      '<div class="ccard-body"><div class="cfg-grid">' +
        '<div class="cfg-field"><label>Demand basis</label>' +
          segHTML('basis', [['used', 'Used'], ['provisioned', 'Provisioned']], cfg.basis) +
          '<div class="cfg-note">Used = actual written data. Provisioned = allocated incl. thin commitments (conservative).</div></div>' +
        '<div class="cfg-field"><label>Data reduction ratio: <strong data-lb="reduction">' + cfg.reduction.toFixed(1) + '×</strong></label>' +
          '<input type="range" data-cfg="reduction" min="1" max="4" step="0.1" value="' + cfg.reduction + '">' +
          '<div class="cfg-note">Dedupe + compression. Logical ÷ ratio = physical.</div></div>' +
        '<div class="cfg-field"><label>Local protection</label>' +
          '<select data-cfg="protection">' + protOpts + '</select>' +
          '<div class="cfg-note">Resilience multiplier applied to reduced data.</div></div>' +
        '<div class="cfg-field"><label>Remote replication</label>' +
          segHTML('remote', [['none', 'None'], ['async', 'Async copy']], cfg.remote) +
          '<div class="cfg-note">Async adds full replica copies of the protected data.</div></div>' +
        '<div class="cfg-field" data-remote-only' + (cfg.remote === 'async' ? '' : ' hidden') + '><label>Replica copies: <strong data-lb="remoteCopies">' + cfg.remoteCopies + '</strong></label>' +
          '<input type="number" data-cfg="remoteCopies" min="1" max="2" step="1" value="' + cfg.remoteCopies + '"></div>' +
        '<div class="cfg-field"><label>Snapshot copies retained: <strong data-lb="snapCopies">' + cfg.snapCopies + '</strong></label>' +
          '<input type="range" data-cfg="snapCopies" min="0" max="5" step="1" value="' + cfg.snapCopies + '">' +
          '<div class="cfg-note">How many point-in-time copies you keep around.</div></div>' +
        '<div class="cfg-field"><label>Size per snapshot copy: <strong data-lb="snapPct">' + cfg.snapPct + '%</strong></label>' +
          '<input type="range" data-cfg="snapPct" min="5" max="100" step="5" value="' + cfg.snapPct + '">' +
          '<div class="cfg-note">Each copy as % of the pool&#39;s reduced base (change rate).</div></div>' +
        '<div class="cfg-field"><label>System overhead: <strong data-lb="overhead">' + cfg.overhead + '%</strong></label>' +
          '<input type="range" data-cfg="overhead" min="0" max="30" step="1" value="' + cfg.overhead + '">' +
          '<div class="cfg-note">Formatting, metadata, system reserve.</div></div>' +
        '<div class="cfg-field"><label>Free-space headroom: <strong data-lb="spare">' + cfg.spare + '%</strong></label>' +
          '<input type="range" data-cfg="spare" min="0" max="40" step="1" value="' + cfg.spare + '">' +
          '<div class="cfg-note">Slack for rebuilds and snapshot bursts. Keep ≥15% on most arrays.</div></div>' +
        '<div class="cfg-field"><label>Growth</label>' +
          '<div class="cfg-row" style="align-items:center"><input type="checkbox" data-cfg="useGlobal" style="width:auto;flex:none"' + (cfg.useGlobal ? ' checked' : '') + '><span class="muted" style="font-size:.85rem">Use global growth</span></div>' +
          '<div data-own-only' + (cfg.useGlobal ? ' hidden' : '') + ' style="margin-top:8px"><input type="number" data-cfg="ownGrowth" min="0" step="1" value="' + cfg.ownGrowth + '" placeholder="Pool rate"></div>' +
          '<div class="cfg-note" data-lb="growthNote"></div></div>' +
      '</div>' +
      '<div class="spec-line" data-pv="spec"></div>' +
      '</div></div>';
  }).join('');

  APP.pools.forEach((p) => {
    const card = wrap.querySelector('.ccard[data-id="' + p.id + '"]');
    card.querySelector('.ccard-head').addEventListener('click', (e) => {
      if (e.target.closest('button,select,input')) return;
      card.classList.toggle('open');
    });
    card.querySelectorAll('[data-cfg]').forEach((el) => {
      el.addEventListener('input', () => onCfgInput(p.id, card));
      el.addEventListener('change', () => onCfgInput(p.id, card));
    });
    card.querySelectorAll('[data-seg]').forEach((seg) => {
      seg.querySelectorAll('button:not([disabled])').forEach((b) => {
        b.addEventListener('click', () => {
          seg.querySelectorAll('button').forEach((x) => x.classList.remove('active'));
          b.classList.add('active');
          onCfgInput(p.id, card);
        });
      });
    });
  });
  refreshPreviews();
}

function onCfgInput(pid, card) {
  const cfg = getCfg(pid);
  const val = (s) => { const el = card.querySelector('[data-cfg="' + s + '"]'); return el ? (el.type === 'checkbox' ? el.checked : el.value) : null; };
  const segVal = (s) => { const b = card.querySelector('[data-seg="' + s + '"] button.active'); return b ? b.dataset.val : null; };
  cfg.basis = segVal('basis') || 'used';
  cfg.reduction = Math.min(Math.max(parseFloat(val('reduction')) || 1.8, 1.0), 4.0);
  cfg.protection = val('protection') || 'mirror2';
  cfg.remote = segVal('remote') || 'none';
  cfg.remoteCopies = Math.min(Math.max(parseInt(val('remoteCopies')) || 1, 1), 2);
  cfg.snapCopies = Math.min(Math.max(parseInt(val('snapCopies')) || 0, 0), 5);
  cfg.snapPct = Math.min(Math.max(parseInt(val('snapPct')) || 25, 5), 100);
  cfg.overhead = Math.min(Math.max(parseInt(val('overhead')) || 0, 0), 30);
  cfg.spare = Math.min(Math.max(parseInt(val('spare')) || 0, 0), 40);
  cfg.useGlobal = !!val('useGlobal');
  const og = parseNum(val('ownGrowth'));
  cfg.ownGrowth = cfg.useGlobal ? cfg.ownGrowth : (og > 0 || val('ownGrowth') === '0' ? og : cfg.ownGrowth);
  card.querySelectorAll('[data-remote-only]').forEach((el) => { el.hidden = cfg.remote !== 'async'; });
  card.querySelectorAll('[data-own-only]').forEach((el) => { el.hidden = cfg.useGlobal; });
  refreshPreviews();
}

function onGlobalInput() {
  const G = APP.global;
  G.horizon = Math.min(Math.max(parseInt($('gHorizon').value) || 3, 1), 5);
  G.value = Math.max(parseNum($('gValue').value), 0);
  renderGlobalGrowth();
  refreshPreviews();
}

function refreshPreviews() {
  const G = APP.global;
  APP.pools.forEach((p) => {
    const card = document.querySelector('.ccard[data-id="' + p.id + '"]');
    if (!card) return;
    const cfg = getCfg(p.id);
    const rN = sizePoolYear(p, cfg, G, G.horizon);
    const r0 = sizePoolYear(p, cfg, G, 0);
    const set = (k, v) => { const el = card.querySelector('[data-pv="' + k + '"]'); if (el) el.innerHTML = v; };
    const lbl = (k, v) => { const el = card.querySelector('[data-lb="' + k + '"]'); if (el) el.textContent = v; };
    lbl('reduction', cfg.reduction.toFixed(1) + '×');
    lbl('remoteCopies', cfg.remoteCopies);
    lbl('snapCopies', cfg.snapCopies);
    lbl('snapPct', cfg.snapPct + '%');
    lbl('overhead', cfg.overhead + '%');
    lbl('spare', cfg.spare + '%');
    lbl('growthNote', 'Effective: ' + growthLabel(effGrowth(cfg, G), G.mode) + ' · horizon ' + G.horizon + ' yr');
    set('raw', fmt1(rN.raw));
    set('sub', 'Year-' + G.horizon + ' · ' + esc(PROT[cfg.protection].label) + (cfg.remote === 'async' ? ' +async×' + cfg.remoteCopies : ''));
    set('spec', '<strong>Year ' + G.horizon + ':</strong> ' + fmtTB(rN.logical) + ' logical → ÷' + cfg.reduction.toFixed(1) + ' reduction → ×' + (PROT[cfg.protection].factor) + ' ' + esc(PROT[cfg.protection].label) +
      (cfg.remote === 'async' ? ' → ×' + (1 + cfg.remoteCopies) + ' remote' : '') +
      ' → <strong>' + fmtTB(rN.raw) + ' raw</strong> <span class="muted">(today: ' + fmtTB(r0.raw) + ')</span>');
  });
}

/* ================= Step 3: results ================= */
function computeResults() {
  const G = APP.global;
  return APP.pools.map((p) => {
    const cfg = getCfg(p.id);
    const years = [];
    for (let y = 0; y <= G.horizon; y++) years.push(sizePoolYear(p, cfg, G, y));
    return { p, cfg, years };
  });
}

function renderResults() {
  APP.results = computeResults();
  $('dashFileName').textContent = APP.fileName || '—';
  $('dashMeta').textContent = APP.source === 'rvtools' ? ' · RVTools export' : APP.source === 'demo' ? ' · demo data' : ' · manual entry';
  renderPlanTab(APP.results);
  renderPoolsTab(APP.results);
  renderFindingsTab(APP.results);
  renderReportTab();
  switchTab('plan');
}

function switchTab(name) {
  document.querySelectorAll('#tabs button').forEach((b) => b.classList.toggle('active', b.dataset.tab === name));
  ['plan', 'pools', 'findings', 'report'].forEach((t) => { $('tab-' + t).hidden = t !== name; });
}

function mathStep(n, formula, result, isResult) {
  return '<div class="math-step' + (isResult ? ' result' : '') + '"><div class="ms-num">' + n + '</div>' +
    '<div class="ms-formula">' + formula + '</div><div class="ms-result">' + result + '</div></div>';
}

function protShort(cfg) {
  return PROT[cfg.protection].label + ' (' + PROT[cfg.protection].factor + '×)' + (cfg.remote === 'async' ? ' + async ×' + cfg.remoteCopies : '');
}

function renderPlanTab(res) {
  const G = APP.global;
  const N = G.horizon;
  const sum = (f) => res.reduce((a, x) => a + f(x), 0);
  const raw0 = sum((x) => x.years[0].raw);
  const rawN = sum((x) => x.years[N].raw);
  const growthTB = rawN - raw0;
  const v = APP.validation;

  // Year-by-year projection (aggregated across pools)
  const rows = [];
  for (let y = 0; y <= N; y++) {
    const t = { logical: 0, reduced: 0, withSnaps: 0, withRemote: 0, raw: 0 };
    res.forEach((x) => {
      const s = x.years[y];
      t.logical += s.logical; t.reduced += s.reduced; t.withSnaps += s.withSnaps; t.withRemote += s.withRemote; t.raw += s.raw;
    });
    rows.push('<tr' + (y === N ? ' style="font-weight:700"' : '') + '><td>Year ' + y + (y === N ? ' <span class="muted small">← horizon</span>' : '') + '</td>' +
      '<td class="num">' + fmtTB(t.logical) + '</td><td class="num">' + fmtTB(t.reduced) + '</td>' +
      '<td class="num">' + fmtTB(t.withSnaps) + '</td><td class="num">' + fmtTB(t.withRemote) + '</td>' +
      '<td class="num"><strong>' + fmtTB(t.raw) + '</strong></td></tr>');
  }

  // Per-pool summary
  const poolRows = res.map(({ p, cfg, years }) => {
    const r0 = years[0], rN = years[N];
    const basisTB = cfg.basis === 'provisioned' ? p.provisionedTB : p.usedTB;
    return '<tr><td><strong>' + esc(p.name) + '</strong><br><span class="muted small">' + fmtInt(p.vms) + ' VMs</span></td>' +
      '<td class="num">' + fmtTB(basisTB) + '<br><span class="muted small">' + (cfg.basis === 'provisioned' ? 'provisioned' : 'used') + '</span></td>' +
      '<td class="num">' + cfg.reduction.toFixed(1) + '×</td>' +
      '<td>' + esc(protShort(cfg)) + '</td>' +
      '<td class="num">' + fmtTB(r0.raw) + '</td>' +
      '<td class="num"><strong>' + fmtTB(rN.raw) + '</strong></td></tr>';
  }).join('');

  const stats =
    '<div class="stat"><div class="v blue">' + fmtTB(raw0) + '</div><div class="l">Raw needed today</div></div>' +
    '<div class="stat"><div class="v purple">' + fmtTB(rawN) + '</div><div class="l">Raw needed at Year ' + N + '</div></div>' +
    '<div class="stat"><div class="v ' + (growthTB > 0 ? 'amber' : 'green') + '">' + (growthTB >= 0 ? '+' : '') + fmtTB(growthTB).replace(' TB', '') + ' TB</div><div class="l">Net growth (raw)</div></div>' +
    '<div class="stat"><div class="v green">' + res.length + '</div><div class="l">Pools sized</div></div>' +
    (v && v.dsCount ? '<div class="stat"><div class="v">' + fmtTB(v.dsCapTB) + '</div><div class="l">Current estate capacity</div></div>' : '');

  $('tab-plan').innerHTML =
    '<div class="stat-grid">' + stats + '</div>' +
    '<div class="panel"><h3>Year-by-year projection <span class="sub">all pools · ' + (G.mode === 'compound' ? 'compounding' : 'static') + ' growth</span></h3>' +
    '<div class="table-scroll"><table class="data"><thead><tr><th>Year</th><th class="num">Logical</th><th class="num">After reduction</th><th class="num">+ Snapshots</th><th class="num">+ Protection</th><th class="num">Raw needed</th></tr></thead>' +
    '<tbody>' + rows.join('') + '</tbody></table></div>' +
    '<p class="note">+Protection includes local protection and any remote replica copies. Raw needed adds system overhead and divides by free-space headroom.</p></div>' +
    '<div class="panel"><h3>Per-pool summary</h3><div class="table-scroll"><table class="data">' +
    '<thead><tr><th>Pool</th><th class="num">Basis TB</th><th class="num">Reduction</th><th>Protection</th><th class="num">Raw today</th><th class="num">Raw at Year ' + N + '</th></tr></thead>' +
    '<tbody>' + poolRows + '</tbody></table></div></div>';
}

function workedChain(p, cfg, G, y, label) {
  const s = sizePoolYear(p, cfg, G, y);
  const g = effGrowth(cfg, G);
  const growFormula = G.mode === 'compound'
    ? fmtTB(s.basisTB) + ' × (1 + ' + g + '%)<sup>' + y + '</sup>'
    : fmtTB(s.basisTB) + ' + ' + y + ' × ' + fmt1(g) + ' TB';
  let steps = '', n = 0;
  steps += mathStep(++n, 'demand basis (' + (cfg.basis === 'provisioned' ? 'provisioned' : 'used') + ', ' + fmtInt(p.vms) + ' VMs)', fmtTB(s.basisTB));
  steps += mathStep(++n, 'growth → logical, ' + label + ' = ' + growFormula, fmtTB(s.logical));
  steps += mathStep(++n, 'data reduction ÷ ' + cfg.reduction.toFixed(1), fmtTB(s.reduced));
  steps += mathStep(++n, 'snapshots × (1 + ' + cfg.snapCopies + ' × ' + cfg.snapPct + '%)', fmtTB(s.withSnaps));
  steps += mathStep(++n, 'local protection × ' + s.localF + ' (' + esc(PROT[cfg.protection].label) + ')', fmtTB(s.withLocal));
  steps += mathStep(++n, 'remote replication × ' + s.remoteF + (cfg.remote === 'async' ? ' (' + cfg.remoteCopies + ' async cop' + (cfg.remoteCopies > 1 ? 'ies' : 'y') + ')' : ' (none)'), fmtTB(s.withRemote));
  steps += mathStep(++n, 'system overhead × ' + (1 + cfg.overhead / 100).toFixed(2), fmtTB(s.withOH));
  steps += mathStep(++n, 'free-space headroom ÷ (1 − ' + cfg.spare + '%)', '<strong>' + fmtTB(s.raw) + ' raw</strong>', true);
  return '<h3 style="margin-top:1.6rem">' + esc(label) + '</h3>' + steps;
}

function renderPoolsTab(res) {
  const G = APP.global;
  const N = G.horizon;
  $('tab-pools').innerHTML = res.map(({ p, cfg }, idx) =>
    '<div class="panel"><h3>' + (idx + 1) + '. ' + esc(p.name) +
    ' <span class="sub">' + fmtInt(p.vms) + ' VMs · basis ' + (cfg.basis === 'provisioned' ? 'provisioned' : 'used') + ' · ' + esc(protShort(cfg)) + '</span></h3>' +
    workedChain(p, cfg, G, 0, 'Today (Year 0)') +
    workedChain(p, cfg, G, N, 'Horizon (Year ' + N + ')') + '</div>'
  ).join('');
}

/* ================= Findings ================= */
function buildFindings(res) {
  const F = [];
  const G = APP.global;
  const N = G.horizon;
  const v = APP.validation;
  const raw0 = res.reduce((a, x) => a + x.years[0].raw, 0);
  const rawN = res.reduce((a, x) => a + x.years[N].raw, 0);

  F.push({ sev: 'info', icon: '📋', title: 'Plan totals',
    body: '<strong>' + fmtTB(raw0) + '</strong> raw today across <strong>' + res.length + '</strong> pool' + (res.length > 1 ? 's' : '') + ' → <strong>' + fmtTB(rawN) + '</strong> raw at Year ' + N + '. Tune anything on the Configure step — this page recomputes when you come back.' });

  // Estate validation vs current capacity
  if (v && v.dsCount && v.dsCapTB > 0) {
    const pct = Math.round(v.dsUsedTB / v.dsCapTB * 100);
    if (rawN > v.dsCapTB) {
      F.push({ sev: 'crit', icon: '🔴', title: 'Horizon demand exceeds current capacity',
        body: 'The estate has <strong>' + fmtTB(v.dsCapTB) + '</strong> usable today but the plan needs <strong>' + fmtTB(rawN) + '</strong> by Year ' + N + ' — a <strong>' + fmtTB(rawN - v.dsCapTB) + '</strong> shortfall. That gap is the core of the refresh conversation.' });
    } else {
      F.push({ sev: 'info', icon: '📦', title: 'Fits inside current capacity — for now',
        body: 'The estate has ' + fmtTB(v.dsCapTB) + ' usable (' + pct + '% consumed) and the plan needs ' + fmtTB(rawN) + ' by Year ' + N + ', leaving ' + fmtTB(v.dsCapTB - rawN) + ' of headroom. The question is timing: when does the growth curve cross capacity?' });
    }
    const sumUsed = res.reduce((a, x) => a + x.p.usedTB, 0);
    if (sumUsed > 0 && v.dsUsedTB > sumUsed * 1.2) {
      F.push({ sev: 'warn', icon: '🔍', title: 'Datastore usage exceeds VM sums by ' + fmtTB(v.dsUsedTB - sumUsed),
        body: 'vDatastore shows ' + fmtTB(v.dsUsedTB) + ' consumed but the VM inventory only accounts for ' + fmtTB(sumUsed) + '. Templates, ISOs, or orphaned VMDKs likely account for the gap — worth quantifying before sizing, since it is real capacity.' });
    }
  }

  res.forEach(({ p, cfg, years }) => {
    const r0 = years[0], rN = years[N];
    const localF = PROT[cfg.protection].factor;

    if (localF >= 2 && cfg.reduction < 1.5) {
      const saving = rN.raw * (0.5 / (cfg.reduction + 0.5));
      F.push({ sev: 'warn', icon: '💾', title: esc(p.name) + ': mirroring without reduction is expensive',
        body: 'Local protection of ' + localF + '× applied to data reducing at only ' + cfg.reduction.toFixed(1) + '×. Data reduction is the biggest lever on this pool — every 0.5× improvement saves roughly <strong>' + fmtTB(saving) + '</strong> of raw at Year ' + N + '.' });
    }
    if (cfg.reduction >= 3) {
      F.push({ sev: 'warn', icon: '⚠️', title: esc(p.name) + ': high reduction assumption',
        body: cfg.reduction.toFixed(1) + '× reduction is workload-dependent — validate with the vendor\'s sizing tool before quoting their number. If it lands at 2×, raw at Year ' + N + ' moves from ' + fmtTB(rN.raw) + ' to ' + fmtTB(rN.raw * cfg.reduction / 2) + '.' });
    }
    if (cfg.snapCopies >= 3) {
      const snapTB = rN.withSnaps - rN.reduced;
      F.push({ sev: 'warn', icon: '📸', title: esc(p.name) + ': snapshot retention costs ' + fmtTB(snapTB),
        body: cfg.snapCopies + ' retained copies at ' + cfg.snapPct + '% each add ' + Math.round((rN.snapFactor - 1) * 100) + '% overhead on the reduced base. Shorter retention or a smaller change-rate assumption is the lever.' });
    }
    if (cfg.spare < 10) {
      F.push({ sev: 'warn', icon: '🧯', title: esc(p.name) + ': headroom below 10%',
        body: 'Only ' + cfg.spare + '% free-space headroom — rebuilds and snapshot bursts need slack. Raising it to 15% would move raw at Year ' + N + ' from ' + fmtTB(rN.raw) + ' to ' + fmtTB(rN.withOH / 0.85) + '.' });
    } else if (cfg.spare > 30) {
      const at20 = rN.withOH / 0.8;
      F.push({ sev: 'info', icon: '🛟', title: esc(p.name) + ': generous headroom',
        body: cfg.spare + '% headroom is conservative — trimming to 20% would save ' + fmtTB(rN.raw - at20) + ' of raw at Year ' + N + '. Keep it if the workload is bursty; otherwise it is quotable capacity.' });
    }
    const g = effGrowth(cfg, G);
    if (G.mode === 'compound' && g >= 25) {
      const dbl = Math.log(2) / Math.log(1 + g / 100);
      F.push({ sev: 'warn', icon: '📈', title: esc(p.name) + ': ' + g + '%/yr compounding doubles demand every ~' + fmt1(dbl) + ' years',
        body: 'At this rate the pool\'s raw need grows from ' + fmtTB(r0.raw) + ' to ' + fmtTB(rN.raw) + ' by Year ' + N + '. Make sure the customer has signed off on that trajectory — compounding punishes optimism.' });
    }
    if (G.mode === 'static' && g > 0 && g > r0.basisTB * 0.2) {
      F.push({ sev: 'warn', icon: '📈', title: esc(p.name) + ': static growth outpaces the base',
        body: 'Adding ' + fmt1(g) + ' TB/yr to a ' + fmtTB(r0.basisTB) + ' base is more than 20%/yr equivalent. Consider whether compounding growth fits this pool better.' });
    }
    if (p.provisionedTB > 0 && p.usedTB > 0 && p.provisionedTB / p.usedTB > 2) {
      F.push({ sev: 'warn', icon: '🧹', title: esc(p.name) + ': ' + fmt1(p.provisionedTB / p.usedTB) + ':1 overprovisioned',
        body: fmtTB(p.provisionedTB) + ' provisioned vs ' + fmtTB(p.usedTB) + ' used. Thin-provisioning cleanup or reclaim candidates live here — the gap is capacity you are protecting but nobody is writing to.' });
    }
    if (cfg.remote === 'async') {
      const repTB = rN.withRemote - rN.withLocal;
      F.push({ sev: 'info', icon: '🌐', title: esc(p.name) + ': async replication adds ' + fmtTB(repTB),
        body: cfg.remoteCopies + ' replica cop' + (cfg.remoteCopies > 1 ? 'ies' : 'y') + ' at 100% of the protected size. Bandwidth and RPO are not modeled — only this capacity footprint.' });
    }
    if (p.usableTB != null && p.usableTB > 0) {
      if (rN.raw > p.usableTB) {
        F.push({ sev: 'crit', icon: '🔴', title: esc(p.name) + ': shortfall of ' + fmtTB(rN.raw - p.usableTB) + ' by Year ' + N,
          body: 'The pool owns ' + fmtTB(p.usableTB) + ' usable but the plan needs ' + fmtTB(rN.raw) + ' at the horizon. ' + (r0.raw > p.usableTB ? 'It is already over capacity today by ' + fmtTB(r0.raw - p.usableTB) + '.' : 'Day one fits with ' + fmtTB(p.usableTB - r0.raw) + ' to spare — the growth curve is the problem.') });
      } else {
        F.push({ sev: 'info', icon: '✅', title: esc(p.name) + ': ' + fmtTB(p.usableTB - rN.raw) + ' of headroom at Year ' + N,
          body: 'Current usable (' + fmtTB(p.usableTB) + ') covers the plan through the horizon. The expansion question is timing, not if.' });
      }
    }
  });

  // Snapshot hygiene from the export
  if (v && v.snapTB > 0.05) {
    const sumUsed = res.reduce((a, x) => a + x.p.usedTB, 0);
    if (sumUsed > 0 && v.snapTB > sumUsed * 0.1) {
      F.push({ sev: 'warn', icon: '🧹', title: 'Snapshot hygiene: ' + fmtTB(v.snapTB) + ' sitting in snapshots',
        body: 'That is more than 10% of used VM storage. Consolidate stale snapshots before sizing — they inflate the "used" basis and the retention math on top of it.' });
    } else {
      F.push({ sev: 'info', icon: '📸', title: fmtTB(v.snapTB) + ' in snapshots today',
        body: 'Factored into the retention tunables as context — your per-copy percentage on each card is the planning assumption going forward.' });
    }
  }
  // Thin vs thick
  if (v && v.disksN > 0) {
    const thinPct = Math.round(v.disksThin / v.disksN * 100);
    if (thinPct < 30) {
      F.push({ sev: 'info', icon: '💽', title: 'Thick-heavy estate: only ' + thinPct + '% of VMDKs are thin',
        body: fmtInt(v.disksThin) + ' of ' + fmtInt(v.disksN) + ' virtual disks are thin-provisioned. Converting thick to thin recovers the gap between provisioned and used — that is the overprovisioning finding, with a concrete lever.' });
    } else {
      F.push({ sev: 'info', icon: '💽', title: thinPct + '% of VMDKs are thin-provisioned',
        body: fmtInt(v.disksThin) + ' of ' + fmtInt(v.disksN) + ' disks. The provisioned-vs-used gap is mostly real future demand, not lazy thick disks.' });
    }
  }
  return F;
}

function renderFindingsTab(res) {
  const F = buildFindings(res);
  $('tab-findings').innerHTML = '<div class="panel"><h3>SE talking points <span class="sub">' + F.length + ' findings</span></h3>' +
    F.map((f) => '<div class="finding ' + f.sev + '"><div class="sev">' + f.icon + '</div><div><strong>' + f.title + '</strong><p>' + f.body + '</p></div></div>').join('') + '</div>';
}

function renderReportTab() {
  $('tab-report').innerHTML = '<div class="panel"><h3>📄 Customer-ready briefing</h3>' +
    '<p class="muted">Generates a standalone HTML report — growth plan stats, the year-by-year projection, per-pool worked math, estate validation, findings, and methodology. Self-contained (no external dependencies), safe to email. Scrub customer names from pool labels first if it leaves your org.</p>' +
    '<div class="toolbar"><button class="btn primary" id="dlReportBtn2">⬇ Download HTML report</button></div></div>';
  $('dlReportBtn2').onclick = downloadReport;
}

/* ================= Report ================= */
function buildReportHTML(res) {
  const date = new Date().toISOString().slice(0, 10);
  const G = APP.global;
  const N = G.horizon;
  const srcLbl = APP.source === 'rvtools' ? 'RVTools export (' + APP.fileName + ')' : APP.source === 'demo' ? 'Demo data (synthetic)' : 'Manual entry';
  const v = APP.validation;
  const raw0 = res.reduce((a, x) => a + x.years[0].raw, 0);
  const rawN = res.reduce((a, x) => a + x.years[N].raw, 0);
  const F = buildFindings(res);
  const css = 'body{font-family:-apple-system,"Segoe UI",Inter,Roboto,Helvetica,Arial,sans-serif;margin:0;color:#1a1f28;line-height:1.55}' +
    '.wrap{max-width:960px;margin:0 auto;padding:32px 24px}' +
    'h1{font-size:1.9rem;margin:0 0 4px}h2{font-size:1.35rem;margin:2.2rem 0 .8rem;border-bottom:2px solid #4f8cff;padding-bottom:6px}h3{font-size:1.1rem;margin:1.6rem 0 .6rem}' +
    '.meta{color:#5b6572;font-size:.9rem;margin-bottom:1.5rem}' +
    'table{width:100%;border-collapse:collapse;font-size:.88rem;margin:.8rem 0}' +
    'th{text-align:left;color:#5b6572;font-size:.75rem;text-transform:uppercase;letter-spacing:.04em;padding:8px 10px;border-bottom:2px solid #d5dbe4}' +
    'td{padding:8px 10px;border-bottom:1px solid #e6ebf1;vertical-align:top}' +
    '.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}' +
    '.stat{display:inline-block;background:#f2f5fa;border:1px solid #dbe2ec;border-radius:10px;padding:12px 20px;margin:0 10px 10px 0}' +
    '.stat .v{font-size:1.5rem;font-weight:800}.stat .l{font-size:.8rem;color:#5b6572}' +
    '.finding{border:1px solid #dbe2ec;border-left:4px solid #4f8cff;border-radius:8px;padding:12px 16px;margin-bottom:10px;background:#fafbfe}' +
    '.finding.warn{border-left-color:#f5a623}.finding.crit{border-left-color:#ff6b6b}.finding.info{border-left-color:#4f8cff}' +
    '.mono{font-family:ui-monospace,Menlo,Consolas,monospace;font-size:.85em;background:#f2f5fa;padding:1px 6px;border-radius:5px}' +
    '.disclaimer{background:#fff8ec;border:1px solid #f0d9a8;border-radius:10px;padding:14px 18px;margin-top:2rem;font-size:.9rem}' +
    '@media print{.wrap{padding:0}}';

  const projRows = [];
  for (let y = 0; y <= N; y++) {
    const t = { logical: 0, reduced: 0, withSnaps: 0, withRemote: 0, raw: 0 };
    res.forEach((x) => { const s = x.years[y]; t.logical += s.logical; t.reduced += s.reduced; t.withSnaps += s.withSnaps; t.withRemote += s.withRemote; t.raw += s.raw; });
    projRows.push('<tr><td>Year ' + y + '</td><td class="num">' + fmtTB(t.logical) + '</td><td class="num">' + fmtTB(t.reduced) + '</td><td class="num">' + fmtTB(t.withSnaps) + '</td><td class="num">' + fmtTB(t.withRemote) + '</td><td class="num"><strong>' + fmtTB(t.raw) + '</strong></td></tr>');
  }

  const poolRows = res.map(({ p, cfg, years }) =>
    '<tr><td><strong>' + esc(p.name) + '</strong></td><td class="num">' + fmtTB(cfg.basis === 'provisioned' ? p.provisionedTB : p.usedTB) + '</td><td class="num">' + cfg.reduction.toFixed(1) + '×</td><td>' + esc(protShort(cfg)) + '</td><td class="num">' + fmtTB(years[0].raw) + '</td><td class="num"><strong>' + fmtTB(years[N].raw) + '</strong></td></tr>').join('');

  const perPool = res.map(({ p, cfg }, idx) => {
    const chain = (y, label) => {
      const s = sizePoolYear(p, cfg, G, y);
      const g = effGrowth(cfg, G);
      const growFormula = G.mode === 'compound'
        ? fmtTB(s.basisTB) + ' × (1 + ' + g + '%)<sup>' + y + '</sup>'
        : fmtTB(s.basisTB) + ' + ' + y + ' × ' + fmt1(g) + ' TB';
      const rows = [
        ['Demand basis', fmtTB(s.basisTB) + ' (' + (cfg.basis === 'provisioned' ? 'provisioned' : 'used') + ', ' + fmtInt(p.vms) + ' VMs)'],
        ['Growth → logical (' + label + ')', growFormula + ' = ' + fmtTB(s.logical)],
        ['Data reduction', '÷ ' + cfg.reduction.toFixed(1) + ' = ' + fmtTB(s.reduced)],
        ['Snapshots', '× (1 + ' + cfg.snapCopies + ' × ' + cfg.snapPct + '%) = ' + fmtTB(s.withSnaps)],
        ['Local protection', '× ' + s.localF + ' (' + esc(PROT[cfg.protection].label) + ') = ' + fmtTB(s.withLocal)],
        ['Remote replication', '× ' + s.remoteF + (cfg.remote === 'async' ? ' (' + cfg.remoteCopies + ' async copies)' : ' (none)') + ' = ' + fmtTB(s.withRemote)],
        ['System overhead', '× ' + (1 + cfg.overhead / 100).toFixed(2) + ' = ' + fmtTB(s.withOH)],
        ['<strong>Free-space headroom</strong>', '<strong>÷ (1 − ' + cfg.spare + '%) = ' + fmtTB(s.raw) + ' raw</strong>'],
      ];
      return '<h3>' + esc(label) + '</h3><table><tbody>' + rows.map((x) => '<tr><td style="width:38%;color:#5b6572">' + x[0] + '</td><td>' + x[1] + '</td></tr>').join('') + '</tbody></table>';
    };
    return '<h3>' + (idx + 1) + '. ' + esc(p.name) + ' <span style="color:#5b6572;font-weight:400;font-size:.85rem">· ' + fmtInt(p.vms) + ' VMs · ' + esc(protShort(cfg)) + '</span></h3>' +
      chain(0, 'Today (Year 0)') + chain(N, 'Horizon (Year ' + N + ')');
  }).join('');

  const estate = (v && v.dsCount)
    ? '<h2>2. Estate validation</h2><table><thead><tr><th>Metric</th><th class="num">Value</th></tr></thead><tbody>' +
      '<tr><td>Datastores detected</td><td class="num">' + v.dsCount + '</td></tr>' +
      '<tr><td>Current estate capacity</td><td class="num">' + fmtTB(v.dsCapTB) + '</td></tr>' +
      '<tr><td>Consumed today</td><td class="num">' + fmtTB(v.dsUsedTB) + '</td></tr>' +
      '<tr><td>Snapshot footprint</td><td class="num">' + (v.snapTB > 0.05 ? fmtTB(v.snapTB) : '—') + '</td></tr>' +
      '<tr><td>Thin-provisioned VMDKs</td><td class="num">' + (v.disksN ? fmtInt(v.disksThin) + ' of ' + fmtInt(v.disksN) : '—') + '</td></tr>' +
      '<tr><td><strong>Plan raw at Year ' + N + ' vs estate capacity</strong></td><td class="num"><strong>' + fmtTB(rawN) + ' vs ' + fmtTB(v.dsCapTB) + '</strong></td></tr>' +
      '</tbody></table>' : '';

  const findings = F.map((f) => '<div class="finding ' + f.sev + '"><strong>' + f.title + '</strong><br><span style="color:#5b6572">' + f.body + '</span></div>').join('');

  return '<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>Storage Sizer — Growth Plan (' + date + ')</title><style>' + css + '</style></head><body><div class="wrap">' +
    '<h1>Storage Growth Plan</h1><div class="meta">Generated ' + date + ' · Source: ' + esc(srcLbl) + ' · Storage Sizer (client-side sizing tool)</div>' +
    '<div><div class="stat"><div class="v">' + fmtTB(raw0) + '</div><div class="l">Raw needed today</div></div>' +
    '<div class="stat"><div class="v">' + fmtTB(rawN) + '</div><div class="l">Raw needed at Year ' + N + '</div></div>' +
    '<div class="stat"><div class="v">' + res.length + '</div><div class="l">Pools sized</div></div></div>' +
    '<h2>1. Year-by-year projection</h2><table><thead><tr><th>Year</th><th class="num">Logical</th><th class="num">After reduction</th><th class="num">+ Snapshots</th><th class="num">+ Protection</th><th class="num">Raw needed</th></tr></thead><tbody>' + projRows.join('') + '</tbody></table>' +
    '<h2>2. Per-pool summary</h2><table><thead><tr><th>Pool</th><th class="num">Basis TB</th><th class="num">Reduction</th><th>Protection</th><th class="num">Raw today</th><th class="num">Raw at Year ' + N + '</th></tr></thead><tbody>' + poolRows + '</tbody></table>' +
    estate +
    '<h2>3. Per-pool worked math</h2>' + perPool +
    '<h2>4. Findings</h2>' + findings +
    '<h2>5. Methodology</h2><p style="color:#5b6572;font-size:.9rem">Per pool and per year: logical demand = basis (used or provisioned storage) grown ' + (G.mode === 'compound' ? 'compounding at the configured %/yr' : 'by the configured TB/yr') + ' over ' + N + ' years; physical = logical ÷ data-reduction ratio; then multiplied through snapshot retention (1 + copies × size%), local protection factor, remote replica copies, system overhead (1 + %); finally divided by (1 − free-space headroom %). Demand per pool from RVTools vInfo (all non-template VMs, regardless of power state) or manual entry. This is capacity sizing, not performance sizing: it does not model IOPS, latency, or workload behavior. Reduction ratios, snapshot change rates, and growth rates are planner assumptions, not measurements.</p>' +
    '<div class="disclaimer">⚠️ <strong>Indicative analysis, not a quote.</strong> Array pricing, data-reduction guarantees, software licensing, and partner programs affect real cost. Validate all figures against an official quote before committing to purchases.</div>' +
    '</div></body></html>';
}

function downloadReport() {
  if (!APP.results) return;
  const html = buildReportHTML(APP.results);
  const blob = new Blob([html], { type: 'text/html' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'storage-sizer-growth-plan-' + new Date().toISOString().slice(0, 10) + '.html';
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
}

/* ================= Wiring ================= */
function clearSession() {
  APP.pools = []; APP.cfgs = {}; APP.global = defaultGlobal(); APP.source = null; APP.fileName = null; APP.validation = null; APP.results = null;
  $('inventoryWrap').hidden = true;
  $('manualEditor').hidden = true;
  $('manualBody').innerHTML = '';
  $('fileInput').value = '';
  clearMsgs();
  $('wizard').hidden = true;
  $('landing').hidden = false;
  window.scrollTo({ top: 0 });
}

/* ================= Changelog ================= */
function renderChangelog() {
  const body = $('changelog-body');
  if (!body) return;
  fetch('CHANGELOG.md', { cache: 'no-store' })
    .then((res) => { if (!res.ok) throw new Error('bad status'); return res.text(); })
    .then((md) => {
      let html = '', inList = false;
      const closeList = () => { if (inList) { html += '</ul>'; inList = false; } };
      for (const line of md.split('\n')) {
        if (line.startsWith('## ')) { closeList(); html += '<h4>' + esc(line.slice(3).trim()) + '</h4>'; }
        else if (line.startsWith('- ')) { if (!inList) { html += '<ul>'; inList = true; } html += '<li>' + esc(line.slice(2).trim()) + '</li>'; }
        else if (line.trim() === '' || line.startsWith('# ')) { closeList(); }
        else { closeList(); html += '<p>' + esc(line.trim()) + '</p>'; }
      }
      closeList();
      body.innerHTML = html;
    })
    .catch(() => { body.innerHTML = "<p class='muted'>Changelog unavailable.</p>"; });
}

function wireApp() {
  document.querySelector('.cta').addEventListener('click', (e) => { e.preventDefault(); startWizard(); });
  $('brandHome').addEventListener('click', (e) => { e.preventDefault(); $('wizard').hidden = true; $('landing').hidden = false; window.scrollTo({ top: 0 }); });

  const dz = $('dropzone'), fi = $('fileInput');
  dz.addEventListener('click', () => fi.click());
  dz.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fi.click(); } });
  ['dragover', 'dragenter'].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add('drag'); }));
  ['dragleave', 'drop'].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove('drag'); }));
  dz.addEventListener('drop', (e) => { const f = e.dataTransfer.files && e.dataTransfer.files[0]; if (f) handleFile(f); });
  fi.addEventListener('change', () => { if (fi.files[0]) handleFile(fi.files[0]); });

  $('demoBtn').onclick = loadDemo;
  $('manualBtn').onclick = () => { clearMsgs(); const me = $('manualEditor'); me.hidden = !me.hidden; if (!me.hidden) { wireManualEditor(); me.scrollIntoView({ behavior: 'smooth' }); } };
  $('manualAddRow').onclick = () => { $('manualBody').insertAdjacentHTML('beforeend', manualRowHTML()); wireManualEditor(); };
  $('manualApply').onclick = applyManual;

  $('backToStartBtn').onclick = () => { APP.pools = []; APP.source = null; APP.validation = null; $('inventoryWrap').hidden = true; clearMsgs(); window.scrollTo({ top: 0, behavior: 'smooth' }); };
  $('toConfigBtn').onclick = () => { if (!APP.pools.length) return; APP.pools.forEach((p) => getCfg(p.id)); renderConfig(); setStep(2); };
  $('backToDataBtn').onclick = () => setStep(1);
  $('toResultsBtn').onclick = () => { if (!APP.pools.length) return; renderResults(); setStep(3); };
  $('backToConfigBtn').onclick = () => setStep(2);

  // Global growth controls
  $('gHorizon').addEventListener('input', onGlobalInput);
  $('gValue').addEventListener('input', onGlobalInput);
  $('gMode').querySelectorAll('button').forEach((b) => {
    b.addEventListener('click', () => {
      const mode = b.dataset.val;
      if (mode === APP.global.mode) return;
      APP.global.mode = mode;
      APP.global.value = mode === 'compound' ? 20 : 5;
      onGlobalInput();
    });
  });

  document.querySelectorAll('#tabs button').forEach((b) => b.addEventListener('click', () => switchTab(b.dataset.tab)));
  $('printBtn').onclick = () => window.print();
  $('dlReportBtn').onclick = downloadReport;
  $('clearBtn').onclick = clearSession;
  renderChangelog();
}

document.addEventListener('DOMContentLoaded', wireApp);
