'use strict';
/* ============================================================
   Network Sizer — TOR & FC SAN switch planner for VMware refreshes
   100% client-side. No uploads, no storage, no network calls
   carrying customer data. Everything lives in page memory.
   ============================================================ */

/* ================= Helpers ================= */
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fmtInt = (n) => Math.round(n || 0).toLocaleString('en-US');
const fmt1 = (n) => (n == null || !isFinite(n)) ? '—' : (Math.round(n * 10) / 10).toLocaleString('en-US');
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

/* ================= Switch presets =================
   TOR: dl = downlink ports, dlSpeed = Gbps, ul/ulSpeed = dedicated
   uplinks. sharedPorts = uplinks carve out of the downlink pool. */
const TOR = {
  nx93180:  { label: 'Cisco Nexus 93180YC-FX3', dl: 48, dlSpeed: 25, ul: 6, ulSpeed: 100, note: '48×25G SFP28 + 6×100G QSFP28' },
  arista:   { label: 'Arista 7050SX3-48YC8',     dl: 48, dlSpeed: 25, ul: 8, ulSpeed: 100, note: '48×25G + 8×100G, wire-speed' },
  dell5248: { label: 'Dell PowerSwitch S5248F-ON', dl: 48, dlSpeed: 25, ul: 4, ulSpeed: 100, note: '48×25G + 4×100G (+2×200G not modeled)' },
  nx9364:   { label: 'Cisco Nexus 9364C-GX',     dl: 64, dlSpeed: 100, ul: 0, ulSpeed: 100, sharedPorts: true, note: '64×100G shared pool; uplinks carve from downlinks; 4×25G breakout' },
  custom:   { label: 'Custom switch', custom: true, note: 'Enter your own port counts' },
};
const FC = {
  g720:     { label: 'Brocade G720',    ports: 64,  maxSpeed: 64, note: 'Gen 7 · 16/32/64G auto-sensing · 1RU' },
  g730:     { label: 'Brocade G730',    ports: 128, maxSpeed: 64, note: 'Gen 7 · 128×64G · 2RU' },
  mds9148t: { label: 'Cisco MDS 9148T', ports: 48,  maxSpeed: 32, note: '32G · 24→48 ports via license' },
  mds9396t: { label: 'Cisco MDS 9396T', ports: 96,  maxSpeed: 32, note: '32G · 48→96 ports via license' },
  mds9132t: { label: 'Cisco MDS 9132T', ports: 32,  maxSpeed: 32, note: '32G · 8→32 ports via license' },
  custom:   { label: 'Custom FC switch', custom: true, note: 'Enter your own port count' },
};

/* ================= RVTools column aliases ================= */
const COL = {
  host: {
    name: ['Host', 'Hostname', 'Host Name'],
    cluster: ['Cluster'],
    dc: ['Datacenter', 'DC'],
  },
  vm: {
    name: ['VM'],
    dc: ['Datacenter', 'DC'],
    cluster: ['Cluster'],
  },
};
const TAB_NAMES = { vHost: ['vHost'], vInfo: ['vInfo'] };
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
    if (/(^|\s)(vm|host|name|cluster|datacenter|datastore)(s|\s|$)/.test(' ' + joined + ' ')) { headerIdx = i; break; }
  }
  return XLSX.utils.sheet_to_json(ws, { defval: '', raw: true, range: headerIdx });
}
function clKey(dc, cluster) { return dc + ' / ' + (cluster || 'Standalone'); }

function buildGroups(hosts, vms) {
  const byCl = {};
  hosts.forEach((h) => {
    const k = clKey(h.dc, h.cluster);
    if (!byCl[k]) byCl[k] = { name: k, hosts: 0, vms: 0 };
    byCl[k].hosts++;
  });
  vms.forEach((v) => {
    const k = clKey(v.dc, v.cluster);
    if (!byCl[k]) byCl[k] = { name: k, hosts: 0, vms: 0 };
    byCl[k].vms++;
  });
  let skipped = 0;
  return Object.keys(byCl).sort().map((k, i) => {
    const g = byCl[k];
    if (g.hosts === 0) { skipped++; return null; }
    return { id: 'g' + i, name: k, hosts: g.hosts, vms: g.vms, source: 'rvtools' };
  }).filter(Boolean).map((g, i) => (g.id = 'g' + i, g));
}

/* ================= Demo data ================= */
function genDemoGroups() {
  return [
    { id: 'd1', name: 'DC-East / Prod-General', hosts: 12, vms: 340, source: 'demo' },
    { id: 'd2', name: 'DC-East / VDI', hosts: 6, vms: 900, source: 'demo' },
    { id: 'd3', name: 'DC-West / Edge', hosts: 3, vms: 80, source: 'demo' },
  ];
}

/* ================= Sizing math (pure, unit-tested) ================= */
function resolveTor(eth) {
  const P = TOR[eth.preset] || TOR.nx93180;
  if (P.custom) return { label: 'Custom', dl: Math.max(1, eth.cDl | 0), dlSpeed: eth.cDlS, ulSpeed: eth.cUlS, sharedPorts: false, note: 'custom' };
  return P;
}
function resolveFc(fc) {
  const P = FC[fc.preset] || FC.g720;
  if (P.custom) return { label: 'Custom', ports: Math.max(1, fc.cPorts | 0), maxSpeed: fc.cSpeed, note: 'custom' };
  return P;
}

/* Ethernet TOR sizing.
   prof: {mgmtN,mgmtS,dataN,dataS,storN,storS} speeds in Gbps
   eth: {preset,cDl,cDlS,cUlS,uplinks,overTarget,dual,spare,breakout} */
function sizeEthernet(hosts, prof, eth) {
  const P = resolveTor(eth);
  const uplinks = Math.max(0, eth.uplinks | 0);
  const usableDl = P.sharedPorts ? Math.max(0, P.dl - uplinks) : P.dl;
  const classes = [
    { n: prof.mgmtN | 0, s: prof.mgmtS },
    { n: prof.dataN | 0, s: prof.dataS },
    { n: prof.storN | 0, s: prof.storS },
  ];
  let servedPorts = 0, servedBW = 0, unserved = 0, subPorts = 0, fullPorts = 0;
  classes.forEach((c) => {
    if (!c.n) return;
    const cnt = hosts * c.n;
    if (P.sharedPorts && P.dlSpeed === 100 && eth.breakout && c.s <= 25) {
      subPorts += cnt; servedPorts += cnt; servedBW += cnt * c.s;
    } else if (c.s <= P.dlSpeed) {
      if (P.sharedPorts && P.dlSpeed === 100) fullPorts += cnt;
      else servedPorts += cnt;
      servedBW += cnt * c.s;
    } else {
      unserved += cnt;
    }
  });
  // For shared 100G ports: full-speed ports consume whole ports, sub-ports consume 1/4
  let switchesRaw;
  if (P.sharedPorts && P.dlSpeed === 100) {
    const portCap = usableDl, subCap = usableDl * 4;
    switchesRaw = Math.max(
      fullPorts ? Math.ceil(fullPorts / Math.max(1, portCap)) : 0,
      subPorts ? Math.ceil(subPorts / Math.max(1, subCap)) : 0
    );
    servedPorts = fullPorts + subPorts;
  } else {
    switchesRaw = servedPorts ? Math.ceil(servedPorts / Math.max(1, usableDl)) : 0;
  }
  let switches = switchesRaw;
  if (eth.dual && switchesRaw > 0) switches = Math.max(2, switchesRaw + (switchesRaw % 2));
  if (eth.spare && switchesRaw > 0) switches += 1;
  // portEquiv: physical ports consumed (sub-ports collapse 4:1 with breakout on)
  const portEquiv = (P.sharedPorts && P.dlSpeed === 100)
    ? fullPorts + (eth.breakout ? subPorts / 4 : subPorts)
    : servedPorts;
  const capPorts = switches * usableDl;
  const util = capPorts > 0 ? portEquiv / capPorts : 0;
  const uplinkBW = switches * uplinks * P.ulSpeed; // Gbps
  const over = uplinkBW > 0 ? servedBW / uplinkBW : (servedBW > 0 ? Infinity : 0);
  return {
    P, uplinks, usableDl, hosts,
    servedPorts, servedBW, unserved, portEquiv, fullPorts, subPorts,
    switchesRaw, switches, util, uplinkBW, over,
    mgmt1G: (prof.mgmtN | 0) > 0 && prof.mgmtS === 1 && P.dlSpeed >= 25,
  };
}

/* FC SAN sizing.
   fc: {enabled,preset,cPorts,cSpeed,isl,dual,arrays,targets,targetSpeed} */
function sizeFC(hosts, hbaN, hbaSpeed, fc) {
  if (!fc.enabled || !(hbaN > 0)) {
    return { enabled: false, switches: 0, hostPorts: 0, targetPorts: 0, devicePorts: 0 };
  }
  const P = resolveFc(fc);
  const hostPorts = hosts * (hbaN | 0);
  const targetPorts = Math.max(0, fc.arrays | 0) * Math.max(0, fc.targets | 0);
  const devicePorts = hostPorts + targetPorts;
  const usable = Math.max(1, P.ports - Math.max(0, fc.isl | 0));
  const perFabric = fc.dual ? Math.ceil(devicePorts / 2) : devicePorts;
  const perFabricSw = devicePorts > 0 ? Math.max(1, Math.ceil(perFabric / usable)) : 0;
  const switches = fc.dual ? perFabricSw * 2 : perFabricSw;
  const effSpeed = Math.min(hbaSpeed, fc.targetSpeed, P.maxSpeed);
  const islBW = Math.max(0, fc.isl | 0) * effSpeed; // Gbps per switch reserved
  const ratio = targetPorts > 0 ? hostPorts / targetPorts : null;
  const util = switches > 0 ? perFabric / (perFabricSw * usable || 1) : 0;
  return {
    enabled: true, P, hosts, hostPorts, targetPorts, devicePorts,
    usable, perFabric, perFabricSw, switches, util,
    islBW, ratio, effSpeed,
    hbaCapped: hbaSpeed > P.maxSpeed,
    targetCapped: fc.targetSpeed > P.maxSpeed,
    oddHba: fc.dual && ((hbaN | 0) % 2 === 1),
  };
}

/* Net-new vs existing inventory.
   ethInv: [{preset,count,cDl}], fcInv: [{preset,count,cPorts}] */
function netNewEthernet(ethRes, eth, ethInv) {
  let reuse = 0;
  (ethInv || []).forEach((r) => {
    const P = TOR[r.preset] || TOR.nx93180;
    reuse += (P.custom ? Math.max(0, r.cDl | 0) : P.dl) * Math.max(0, r.count | 0);
  });
  const remaining = Math.max(0, ethRes.portEquiv - reuse);
  const P = ethRes.P;
  const usableDl = P.sharedPorts ? Math.max(0, P.dl - ethRes.uplinks) : P.dl;
  let raw = remaining > 0 ? Math.ceil(remaining / Math.max(1, usableDl)) : 0;
  let sw = raw;
  if (eth.dual && raw > 0) sw = Math.max(2, raw + (raw % 2));
  if (eth.spare && raw > 0) sw += 1;
  return { reuse, remaining, switches: sw };
}
function netNewFC(fcRes, fc, fcInv) {
  if (!fcRes.enabled) return { reuse: 0, switches: 0 };
  let reuse = 0;
  (fcInv || []).forEach((r) => {
    const P = FC[r.preset] || FC.g720;
    reuse += (P.custom ? Math.max(0, r.cPorts | 0) : P.ports) * Math.max(0, r.count | 0);
  });
  const P = fcRes.P;
  const usable = Math.max(1, P.ports - Math.max(0, fc.isl | 0));
  const perFabricReuse = fc.dual ? Math.floor(reuse / 2) : reuse;
  const perFabricRem = Math.max(0, fcRes.perFabric - perFabricReuse);
  const perFabricSw = perFabricRem > 0 ? Math.max(1, Math.ceil(perFabricRem / usable)) : 0;
  return { reuse, switches: fc.dual ? perFabricSw * 2 : perFabricSw };
}

// Export for node unit tests (guarded — undefined in the browser)
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { sizeEthernet, sizeFC, netNewEthernet, netNewFC, TOR, FC, parseNum, resolveTor, resolveFc, serializeState, projectEnvelope, validProject };
}

/* ================= App state ================= */
const APP = {
  groups: [], ethInv: [], fcInv: [],
  prof: null, eth: null, fc: null,
  source: null, fileName: null, results: null,
  projectName: 'Untitled project',
};
function defaultProf() { return { hosts: 0, mgmtN: 2, mgmtS: 1, dataN: 2, dataS: 25, storN: 2, storS: 25, hbaN: 2, hbaS: 32 }; }
function defaultEth() { return { preset: 'nx93180', cDl: 48, cDlS: 25, cUlS: 100, uplinks: 6, overTarget: 3, dual: true, spare: false, breakout: true }; }
function defaultFc() { return { enabled: true, preset: 'g720', cPorts: 64, cSpeed: 64, isl: 8, arrays: 2, targets: 4, targetSpeed: 32, dual: true }; }
function detectedHosts() { return APP.groups.reduce((a, g) => a + (g.hosts || 0), 0); }
function ensureCfg() {
  if (!APP.prof) APP.prof = defaultProf();
  if (!APP.eth) APP.eth = defaultEth();
  if (!APP.fc) APP.fc = defaultFc();
  if (!APP.prof.hosts) APP.prof.hosts = detectedHosts();
}

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
  const hostWs = findSheet(wb, TAB_NAMES.vHost);
  if (!hostWs) {
    showError('<strong>No <code>vHost</code> tab found.</strong> Network sizing starts from hosts — make sure this is an RVTools export with a <code>vHost</code> tab (File → Export all to Excel).');
    return;
  }
  const hosts = sheetRows(wb.Sheets[hostWs]).map((r) => ({
    name: String(pick(r, COL.host.name) || 'unknown'),
    dc: String(pick(r, COL.host.dc) || '—'),
    cluster: String(pick(r, COL.host.cluster) || ''),
  })).filter((h) => h.name && h.name !== 'unknown');
  if (!hosts.length) { showError('<strong>No host rows found</strong> in the <code>vHost</code> tab.'); return; }
  const infoWs = findSheet(wb, TAB_NAMES.vInfo);
  const vms = infoWs ? sheetRows(wb.Sheets[infoWs]).map((r) => ({
    dc: String(pick(r, COL.vm.dc) || '—'),
    cluster: String(pick(r, COL.vm.cluster) || ''),
  })) : [];
  const groups = buildGroups(hosts, vms);
  if (!groups.length) { showError('<strong>No host groups found</strong> in this export.'); return; }
  APP.groups = groups;
  APP.source = 'rvtools';
  APP.fileName = fileName;
  APP.prof = null; APP.eth = null; APP.fc = null;
  const totalH = groups.reduce((a, g) => a + g.hosts, 0);
  const totalV = groups.reduce((a, g) => a + g.vms, 0);
  setStatus('Parsed <strong>' + fmtInt(hosts.length) + '</strong> hosts → <strong>' + groups.length + '</strong> group' + (groups.length > 1 ? 's' : '') +
    (vms.length ? ' · <strong>' + fmtInt(vms.length) + '</strong> VMs for context.' : '. No vInfo tab — VM context unavailable.'));
  renderInventory();
  queueAutosave();
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
  APP.groups = genDemoGroups();
  APP.source = 'demo';
  APP.fileName = 'demo-environment';
  APP.prof = null; APP.eth = null; APP.fc = null;
  setStatus('Loaded <strong>demo environment</strong>: 3 synthetic clusters (prod, VDI, edge) — 21 hosts total. Generated in your browser — nothing uploaded.');
  renderInventory();
  queueAutosave();
}

/* ---- Manual entry ---- */
function manualRowHTML() {
  return '<tr>' +
    '<td><input data-f="name" placeholder="e.g. DC-East / Prod" value=""></td>' +
    '<td><input class="num" data-f="hosts" type="number" min="0" value=""></td>' +
    '<td><input class="num" data-f="vms" type="number" min="0" value=""></td>' +
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
  const groups = [];
  rows.forEach((tr, i) => {
    const g = (f) => tr.querySelector('[data-f="' + f + '"]').value.trim();
    const name = g('name') || ('Group ' + (i + 1));
    const hosts = parseNum(g('hosts'));
    if (!hosts) return;
    groups.push({ id: 'm' + i, name, hosts: Math.round(hosts), vms: Math.round(parseNum(g('vms'))), source: 'manual' });
  });
  if (!groups.length) { showError('<strong>No usable rows.</strong> Fill in at least a group name and host count for one group.'); return; }
  APP.groups = groups;
  APP.source = 'manual';
  APP.fileName = 'manual-entry';
  APP.prof = null; APP.eth = null; APP.fc = null;
  setStatus('Using <strong>' + groups.length + '</strong> manually entered group' + (groups.length > 1 ? 's' : '') + ' — <strong>' + fmtInt(groups.reduce((a, g) => a + g.hosts, 0)) + '</strong> hosts total.');
  renderInventory();
  queueAutosave();
}

function renderInventory() {
  $('inventoryWrap').hidden = false;
  $('existingWrap').hidden = false;
  $('inventoryMeta').textContent = APP.source === 'rvtools' ? 'from ' + APP.fileName : APP.source === 'demo' ? 'synthetic demo data' : 'manual entry';
  $('inventoryBody').innerHTML = APP.groups.map((g) =>
    '<tr><td><strong>' + esc(g.name) + '</strong></td>' +
    '<td class="num">' + fmtInt(g.hosts) + '</td>' +
    '<td class="num">' + (g.vms ? fmtInt(g.vms) : '—') + '</td></tr>'
  ).join('');
  const totalH = APP.groups.reduce((a, g) => a + g.hosts, 0);
  const totalV = APP.groups.reduce((a, g) => a + g.vms, 0);
  $('validationNote').innerHTML = '<strong>' + fmtInt(totalH) + ' hosts</strong> total' + (totalV ? ' · ' + fmtInt(totalV) + ' VMs' : '') +
    ' — sizing starts from hosts; VMs are context only.';
  $('inventoryWrap').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/* ---- Existing inventory rows ---- */
function torOptions(cur) {
  return Object.keys(TOR).map((k) =>
    '<option value="' + k + '"' + (k === cur ? ' selected' : '') + '>' + esc(TOR[k].label) + ' — ' + esc(TOR[k].note) + '</option>').join('');
}
function fcOptions(cur) {
  return Object.keys(FC).map((k) =>
    '<option value="' + k + '"' + (k === cur ? ' selected' : '') + '>' + esc(FC[k].label) + ' — ' + esc(FC[k].note) + '</option>').join('');
}
function invRowHTML(kind, row) {
  const opts = kind === 'eth' ? torOptions(row.preset) : fcOptions(row.preset);
  const customPorts = (kind === 'eth' ? TOR[row.preset].custom : FC[row.preset].custom);
  return '<tr>' +
    '<td><select data-f="preset">' + opts + '</select>' +
    (customPorts ? ' <input class="num" data-f="cPorts" type="number" min="1" value="' + (row.cPorts || '') + '" placeholder="ports" style="max-width:90px;margin-top:6px">' : '') + '</td>' +
    '<td><input class="num" data-f="count" type="number" min="0" value="' + (row.count || '') + '"></td>' +
    '<td><button class="btn ghost" data-del style="padding:6px 10px">✕</button></td></tr>';
}
function renderInvTable(kind) {
  queueAutosave();
  const body = $(kind === 'eth' ? 'ethInvBody' : 'fcInvBody');
  const rows = kind === 'eth' ? APP.ethInv : APP.fcInv;
  body.innerHTML = rows.map((r) => invRowHTML(kind, r)).join('');
  body.querySelectorAll('tr').forEach((tr, i) => {
    const sel = tr.querySelector('[data-f="preset"]');
    sel.addEventListener('change', () => { rows[i].preset = sel.value; renderInvTable(kind); });
    const cnt = tr.querySelector('[data-f="count"]');
    cnt.addEventListener('input', () => { rows[i].count = Math.max(0, parseNum(cnt.value)); });
    const cp = tr.querySelector('[data-f="cPorts"]');
    if (cp) cp.addEventListener('input', () => { rows[i].cPorts = Math.max(0, parseNum(cp.value)); });
    tr.querySelector('[data-del]').onclick = () => { rows.splice(i, 1); renderInvTable(kind); };
  });
}

/* ================= Step 2: configure ================= */
function segHTML(seg, opts, cur) {
  return '<div class="seg" data-seg="' + seg + '">' + opts.map((o) =>
    '<button type="button" data-val="' + o[0] + '"' + (o[0] === cur ? ' class="active"' : '') + '>' + o[1] + '</button>'
  ).join('') + '</div>';
}
function segCfg(kind, field, opts, cur) {
  // Segmented pill group bound to a config field (Physgun vibe); the chosen value lives in data-segval.
  return '<div class="seg" data-' + kind + 'cfg="' + field + '" data-segval="' + esc(cur) + '">' + opts.map((o) =>
    '<button type="button" data-val="' + o[0] + '"' + (String(o[0]) === String(cur) ? ' class="active"' : '') + '>' + o[1] + '</button>'
  ).join('') + '</div>';
}
function pgField(label, badgeAttr, badgeVal, slider, scale, note) {
  // Physgun slider block: label row w/ live badge, glowing slider, min/mid/max scale.
  return '<div class="cfg-field"><div class="pg-lab"><span class="pg-lab-t">' + label + '</span>' +
    '<span class="pg-val" ' + badgeAttr + '>' + badgeVal + '</span></div>' + slider +
    '<div class="pg-scale">' + scale.map((s) => '<span>' + s + '</span>').join('') + '</div>' +
    (note ? '<div class="cfg-note">' + note + '</div>' : '') + '</div>';
}
function pgRange(field, kind, min, max, step, val, label, scale, note, displayVal) {
  return pgField(label, 'data-lb="' + field + '"', displayVal != null ? displayVal : val,
    '<input type="range" class="pg-slider" data-' + kind + 'cfg="' + field + '" min="' + min + '" max="' + max +
    '" step="' + step + '" value="' + val + '">', scale, note);
}
function sliderFill(el) {
  const min = parseFloat(el.min) || 0, max = parseFloat(el.max) || 100;
  const v = Math.min(Math.max(parseFloat(el.value) || 0, min), max);
  return max > min ? ((v - min) / (max - min)) * 100 : 0;
}
function paintSlider(el) {
  if (el && el.classList && el.classList.contains('pg-slider') && el.style) {
    el.style.setProperty('--fill', sliderFill(el) + '%');
  }
}
function paintSliders(root) {
  (root || document).querySelectorAll('.pg-slider').forEach(paintSlider);
}
function speedSeg(cur) {
  return [[1, '1G'], [10, '10G'], [25, '25G'], [40, '40G'], [100, '100G']]
    .map((o) => '<option value="' + o[0] + '"' + (o[0] === cur ? ' selected' : '') + '>' + o[1] + '</option>').join('');
}

function renderConfig() {
  ensureCfg();
  const P = APP.prof;
  const gMax = Math.max(128, Math.ceil(detectedHosts() * 2 / 25) * 25);
  const gEl = $('gHosts');
  gEl.max = gMax;
  $('gHostsMid').textContent = Math.round(gMax / 2);
  $('gHostsMax').textContent = gMax;
  if (P.hosts > gMax) P.hosts = gMax;
  gEl.value = P.hosts;
  $('gHostsNum').textContent = fmtInt(P.hosts);
  paintSlider(gEl);
  $('pMgmtN').value = P.mgmtN; $('pMgmtS').value = P.mgmtS;
  $('pDataN').value = P.dataN; $('pDataS').value = P.dataS;
  $('pStorN').value = P.storN; $('pStorS').value = P.storS;
  $('pHbaN').value = P.hbaN; $('pHbaS').value = P.hbaS;

  const eth = APP.eth, fc = APP.fc;
  const torOpts = torOptions(eth.preset), fcOpts = fcOptions(fc.preset);
  const wrap = $('fabricCards');
  wrap.innerHTML =
    '<div class="ccard open" data-card="eth">' +
      '<div class="ccard-head"><div><div class="ccard-title"><span class="pg-glyph">🌐</span><span>Ethernet fabric — top-of-rack</span></div>' +
      '<div class="ccard-sub" data-pv="ethSub"></div></div>' +
      '<div class="ccard-preview"><div class="hosts"><span data-pv="ethSw">—</span> <small>switches</small></div><div class="binding" data-pv="ethOver"></div></div>' +
      '<div class="ccard-toggle">▾</div></div>' +
      '<div class="ccard-body"><div class="cfg-grid">' +
        '<div class="cfg-field"><label>TOR switch preset</label><select data-ecfg="preset">' + torOpts + '</select><div class="cfg-note" data-pv="ethNote"></div></div>' +
        '<div class="cfg-field" data-eth-custom' + (TOR[eth.preset].custom ? '' : ' hidden') + '><label>Custom ports</label><div class="cfg-row"><input type="number" data-ecfg="cDl" min="1" value="' + eth.cDl + '"><select data-ecfg="cDlS">' + speedSeg(eth.cDlS) + '</select></div>' +
        '<div class="cfg-row" style="margin-top:8px"><select data-ecfg="cUlS">' + [[40, '40G uplink'], [100, '100G uplink'], [400, '400G uplink']].map((o) => '<option value="' + o[0] + '"' + (o[0] === eth.cUlS ? ' selected' : '') + '>' + o[1] + '</option>').join('') + '</select></div><div class="cfg-note">Downlink count/speed and uplink speed for a custom SKU.</div></div>' +
        pgRange('uplinks', 'e', 0, 12, 1, eth.uplinks, '🔌 Uplinks / switch', ['0', '6', '12'], 'TOR → spine links per switch (at <span data-pv="ethUlSpeed"></span>).') +
        pgRange('overTarget', 'e', 1, 8, 0.5, eth.overTarget, '📊 Oversubscription target', ['1', '4.5', '8'], 'Downlink ÷ uplink bandwidth. 3:1 is the classic rule of thumb.', eth.overTarget + ':1') +
        '<div class="cfg-field"><div class="pg-lab"><span class="pg-lab-t">🔗 A/B dual-homing</span></div>' + segCfg('e', 'dual', [['1', 'Yes'], ['0', 'No']], eth.dual ? '1' : '0') + '<div class="cfg-note">Rounds the switch count up to pairs.</div></div>' +
        '<div class="cfg-field"><div class="pg-lab"><span class="pg-lab-t">➕ N+1 spare</span></div>' + segCfg('e', 'spare', [['1', 'Yes'], ['0', 'No']], eth.spare ? '1' : '0') + '<div class="cfg-note">Adds one spare switch to the BOM.</div></div>' +
        '<div class="cfg-field" data-eth-breakout' + ((TOR[eth.preset].sharedPorts) ? '' : ' hidden') + '><div class="pg-lab"><span class="pg-lab-t">⚡ 100G breakout</span></div>' + segCfg('e', 'breakout', [['1', 'On'], ['0', 'Off']], eth.breakout ? '1' : '0') + '<div class="cfg-note">Lets ≤25G hosts share one 100G port four ways.</div></div>' +
      '</div><div class="spec-line" data-pv="ethSpec"></div></div></div>' +
    '<div class="ccard open" data-card="fc">' +
      '<div class="ccard-head"><div><div class="ccard-title"><span class="pg-glyph">🔌</span><span>FC fabric — SAN switches</span></div>' +
      '<div class="ccard-sub" data-pv="fcSub"></div></div>' +
      '<div class="ccard-preview"><div class="hosts"><span data-pv="fcSw">—</span> <small>switches</small></div><div class="binding" data-pv="fcFab"></div></div>' +
      '<div class="ccard-toggle">▾</div></div>' +
      '<div class="ccard-body"><div class="cfg-grid">' +
        '<div class="cfg-field"><div class="pg-lab"><span class="pg-lab-t">🔌 Size an FC SAN fabric</span></div>' + segCfg('f', 'enabled', [['1', 'On'], ['0', 'Off']], fc.enabled ? '1' : '0') + '<div class="cfg-note">Switch off for IP-storage-only designs.</div></div>' +
        '<div class="cfg-field"><label>FC switch preset</label><select data-fcfg="preset">' + fcOpts + '</select><div class="cfg-note" data-pv="fcNote"></div></div>' +
        '<div class="cfg-field" data-fc-custom' + (FC[fc.preset].custom ? '' : ' hidden') + '><label>Custom FC ports</label><div class="cfg-row"><input type="number" data-fcfg="cPorts" min="1" value="' + fc.cPorts + '"><select data-fcfg="cSpeed">' + [[16, '16G'], [32, '32G'], [64, '64G']].map((o) => '<option value="' + o[0] + '"' + (o[0] === fc.cSpeed ? ' selected' : '') + '>' + o[1] + '</option>').join('') + '</select></div><div class="cfg-note">Total ports and max speed for a custom SKU.</div></div>' +
        pgRange('isl', 'f', 0, 32, 1, fc.isl, '🔗 ISL ports reserved / switch', ['0', '16', '32'], 'Ports held back for inter-switch links (not usable by hosts/targets).') +
        pgRange('arrays', 'f', 0, 8, 1, fc.arrays, '🗄️ Storage arrays', ['0', '4', '8'], 'Arrays zoned into the fabric.') +
        pgRange('targets', 'f', 0, 16, 1, fc.targets, '🎯 Target ports / array', ['0', '8', '16'], 'Front-end FC ports per array.') +
        '<div class="cfg-field"><label>Target port speed</label><select data-fcfg="targetSpeed">' + [[16, '16G'], [32, '32G'], [64, '64G']].map((o) => '<option value="' + o[0] + '"' + (o[0] === fc.targetSpeed ? ' selected' : '') + '>' + o[1] + '</option>').join('') + '</select><div class="cfg-note">Array front-end speed.</div></div>' +
        '<div class="cfg-field"><div class="pg-lab"><span class="pg-lab-t">🛡️ Fabric redundancy</span></div>' + segCfg('f', 'dual', [['1', 'A/B dual'], ['0', 'Single']], fc.dual ? '1' : '0') + '<div class="cfg-note">Standard SAN posture: every HBA path survives a fabric failure.</div></div>' +
      '</div><div class="spec-line" data-pv="fcSpec"></div></div></div>';

  wrap.querySelectorAll('.ccard').forEach((card) => {
    card.querySelector('.ccard-head').addEventListener('click', (e) => {
      if (e.target.closest('button,select,input')) return;
      card.classList.toggle('open');
    });
  });
  wrap.querySelectorAll('[data-ecfg],[data-fcfg]').forEach((el) => {
    el.addEventListener('input', (e) => { paintSlider(e.target); onCfgInput(); });
    el.addEventListener('change', (e) => { paintSlider(e.target); onCfgInput(); });
  });
  wrap.querySelectorAll('.seg[data-ecfg],.seg[data-fcfg]').forEach((seg) => {
    seg.querySelectorAll('button').forEach((b) => b.addEventListener('click', () => {
      seg.querySelectorAll('button').forEach((x) => x.classList.remove('active'));
      b.classList.add('active');
      seg.dataset.segval = b.dataset.val;
      onCfgInput();
    }));
  });
  paintSliders(wrap);
  refreshPreviews();
}

function readProfileInputs() {
  const P = APP.prof;
  const gMax = parseNum($('gHosts').max) || 128;
  P.hosts = Math.min(Math.max(Math.round(parseNum($('gHosts').value)), 0), gMax) || detectedHosts();
  P.mgmtN = Math.min(Math.max(Math.round(parseNum($('pMgmtN').value)), 0), 8);
  P.mgmtS = parseNum($('pMgmtS').value) || 1;
  P.dataN = Math.min(Math.max(Math.round(parseNum($('pDataN').value)), 0), 8);
  P.dataS = parseNum($('pDataS').value) || 25;
  P.storN = Math.min(Math.max(Math.round(parseNum($('pStorN').value)), 0), 8);
  P.storS = parseNum($('pStorS').value) || 25;
  P.hbaN = Math.min(Math.max(Math.round(parseNum($('pHbaN').value)), 0), 4);
  P.hbaS = parseNum($('pHbaS').value) || 32;
}

function onCfgInput() {
  readProfileInputs();
  const card = document.querySelector('.ccard[data-card="eth"]');
  const E = APP.eth;
  const ev = (s) => { const el = card.querySelector('[data-ecfg="' + s + '"]'); if (!el) return null; if (el.type === 'checkbox') return el.checked; if (el.dataset && el.dataset.segval !== undefined) return el.dataset.segval; return el.value; };
  const newPreset = ev('preset') || 'nx93180';
  if (newPreset !== E.preset) {
    E.preset = newPreset;
    const NP = TOR[E.preset];
    if (!NP.custom) {
      // shared-pool presets have no dedicated uplinks (ul=0) — seed a sane carve-out instead of 0
      E.uplinks = NP.sharedPorts ? 4 : NP.ul;
      const ulEl = card.querySelector('[data-ecfg="uplinks"]');
      if (ulEl) { ulEl.value = E.uplinks; paintSlider(ulEl); }
    }
  }
  E.cDl = Math.max(1, Math.round(parseNum(ev('cDl')) || 48));
  E.cDlS = parseNum(ev('cDlS')) || 25;
  E.cUlS = parseNum(ev('cUlS')) || 100;
  E.uplinks = Math.min(Math.max(Math.round(parseNum(ev('uplinks')) || 0), 0), 12);
  E.overTarget = Math.min(Math.max(parseNum(ev('overTarget')) || 3, 1), 8);
  E.dual = ev('dual') === '1';
  E.spare = ev('spare') === '1';
  E.breakout = ev('breakout') === '1';

  const fcard = document.querySelector('.ccard[data-card="fc"]');
  const F = APP.fc;
  const fv = (s) => { const el = fcard.querySelector('[data-fcfg="' + s + '"]'); if (!el) return null; if (el.type === 'checkbox') return el.checked; if (el.dataset && el.dataset.segval !== undefined) return el.dataset.segval; return el.value; };
  F.enabled = fv('enabled') === '1';
  F.preset = fv('preset') || 'g720';
  F.cPorts = Math.max(1, Math.round(parseNum(fv('cPorts')) || 64));
  F.cSpeed = parseNum(fv('cSpeed')) || 64;
  F.isl = Math.min(Math.max(Math.round(parseNum(fv('isl')) || 0), 0), 32);
  F.arrays = Math.min(Math.max(Math.round(parseNum(fv('arrays')) || 0), 0), 8);
  F.targets = Math.min(Math.max(Math.round(parseNum(fv('targets')) || 0), 0), 16);
  F.targetSpeed = parseNum(fv('targetSpeed')) || 32;
  F.dual = fv('dual') === '1';

  // Show/hide custom + breakout rows
  card.querySelectorAll('[data-eth-custom]').forEach((el) => { el.hidden = !TOR[E.preset].custom; });
  card.querySelectorAll('[data-eth-breakout]').forEach((el) => { el.hidden = !TOR[E.preset].sharedPorts; });
  fcard.querySelectorAll('[data-fc-custom]').forEach((el) => { el.hidden = !FC[F.preset].custom; });
  refreshPreviews();
  queueAutosave();
}

function profileSummary() {
  const P = APP.prof;
  const per = P.mgmtN + P.dataN + P.storN;
  return '<strong>' + fmtInt(P.hosts) + ' hosts</strong> × ' + per + ' eth ports/host (' +
    P.mgmtN + '×' + P.mgmtS + 'G mgmt, ' + P.dataN + '×' + P.dataS + 'G data, ' + P.storN + '×' + P.storS + 'G storage IP)' +
    (P.hbaN > 0 ? ' + <strong>' + P.hbaN + '×' + P.hbaS + 'G FC HBAs</strong>/host' : ' · no FC HBAs');
}

function refreshPreviews() {
  readProfileInputs();
  $('gHostsNum').textContent = fmtInt(APP.prof.hosts);
  $('profileSpec').innerHTML = profileSummary();
  const P = APP.prof, E = APP.eth, F = APP.fc;
  const ethR = sizeEthernet(P.hosts, P, E);
  const fcR = sizeFC(P.hosts, P.hbaN, P.hbaS, F);
  const card = document.querySelector('.ccard[data-card="eth"]');
  const fcard = document.querySelector('.ccard[data-card="fc"]');
  const set = (c, k, v) => { const el = c.querySelector('[data-pv="' + k + '"]'); if (el) el.innerHTML = v; };
  const lbl = (c, k, v) => { const el = c.querySelector('[data-lb="' + k + '"]'); if (el) el.textContent = v; };

  lbl(card, 'uplinks', E.uplinks);
  lbl(card, 'overTarget', E.overTarget + ':1');
  set(card, 'ethSw', ethR.switches);
  set(card, 'ethOver', isFinite(ethR.over) ? fmt1(ethR.over) + ':1 oversub' : 'no uplinks');
  set(card, 'ethSub', esc(ethR.P.label) + ' · ' + fmtInt(ethR.servedPorts) + ' downlink ports');
  set(card, 'ethNote', esc(ethR.P.note || ''));
  set(card, 'ethUlSpeed', ethR.P.ulSpeed + 'G');
  set(card, 'ethSpec', '<strong>' + ethR.switches + '× ' + esc(ethR.P.label) + '</strong> — ' +
    fmtInt(ethR.servedPorts) + ' ports ÷ ' + ethR.usableDl + ' usable/switch' +
    (E.dual ? ' → pair' : '') + (E.spare ? ' + spare' : '') +
    ' · util ' + Math.round(ethR.util * 100) + '%' +
    (ethR.unserved ? ' · <strong style="color:#ff9d9d">' + fmtInt(ethR.unserved) + ' ports need &gt;' + ethR.P.dlSpeed + 'G — unserved!</strong>' : ''));

  lbl(fcard, 'isl', F.isl);
  lbl(fcard, 'arrays', F.arrays);
  lbl(fcard, 'targets', F.targets);
  set(fcard, 'fcSw', fcR.enabled ? fcR.switches : '—');
  set(fcard, 'fcFab', fcR.enabled ? (F.dual ? 'Fabric A/B · ' + fcR.perFabricSw + ' per fabric' : 'single fabric') : 'disabled');
  set(fcard, 'fcSub', fcR.enabled ? esc(fcR.P.label) + ' · ' + fmtInt(fcR.devicePorts) + ' device ports (' + fmtInt(fcR.hostPorts) + ' host + ' + fmtInt(fcR.targetPorts) + ' target)' : 'FC fabric disabled — IP storage only');
  set(fcard, 'fcNote', esc((fcR.P && fcR.P.note) || ''));
  set(fcard, 'fcSpec', fcR.enabled
    ? '<strong>' + fcR.switches + '× ' + esc(fcR.P.label) + '</strong> — ' + fmtInt(fcR.devicePorts) + ' device ports' +
      (F.dual ? ' ÷ 2 fabrics' : '') + ' ÷ ' + fcR.usable + ' usable/switch (' + fcR.P.ports + ' − ' + F.isl + ' ISL)' +
      ' · ' + fmtInt(fcR.islBW) + ' Gbps ISL reserve/switch'
    : 'Enable the FC fabric to size SAN switches, or leave off for an all-IP storage design.');
}

/* ================= Step 3: results ================= */
function computeResults() {
  const P = APP.prof, E = APP.eth, F = APP.fc;
  const ethR = sizeEthernet(P.hosts, P, E);
  const fcR = sizeFC(P.hosts, P.hbaN, P.hbaS, F);
  const nnEth = netNewEthernet(ethR, E, APP.ethInv);
  const nnFc = netNewFC(fcR, F, APP.fcInv);
  return { P: Object.assign({}, P), E: Object.assign({}, E), F: Object.assign({}, F), ethR, fcR, nnEth, nnFc };
}

function renderResults() {
  APP.results = computeResults();
  $('dashFileName').textContent = APP.fileName || '—';
  $('dashMeta').textContent = APP.source === 'rvtools' ? ' · RVTools export' : APP.source === 'demo' ? ' · demo data' : ' · manual entry';
  renderPlanTab();
  renderFabricsTab();
  renderFindingsTab();
  renderReportTab();
  switchTab('plan');
}

function switchTab(name) {
  document.querySelectorAll('#tabs button').forEach((b) => b.classList.toggle('active', b.dataset.tab === name));
  ['plan', 'fabrics', 'findings', 'report'].forEach((t) => { $('tab-' + t).hidden = t !== name; });
}

function mathStep(n, formula, result, isResult) {
  return '<div class="math-step' + (isResult ? ' result' : '') + '"><div class="ms-num">' + n + '</div>' +
    '<div class="ms-formula">' + formula + '</div><div class="ms-result">' + result + '</div></div>';
}
function overStr(o) { return isFinite(o) ? fmt1(o) + ':1' : (o === Infinity ? '∞ (no uplinks)' : '—'); }

function renderPlanTab() {
  const { P, E, F, ethR, fcR, nnEth, nnFc } = APP.results;
  const bom = [];
  bom.push(['🌐 Ethernet', '<strong>' + ethR.switches + '×</strong> ' + esc(ethR.P.label),
    esc(ethR.P.note || ''), ethR.switches + ' required']);
  if (nnEth.reuse > 0) {
    bom.push(['🌐 Ethernet', 'Existing reuse: ' + fmtInt(nnEth.reuse) + ' ports',
      'subtracted from demand', '<strong>' + nnEth.switches + ' net-new</strong>']);
  }
  if (fcR.enabled) {
    bom.push(['🔌 Fibre Channel', '<strong>' + fcR.switches + '×</strong> ' + esc(fcR.P.label),
      esc(fcR.P.note || '') + (F.dual ? ' · Fabric A/B' : ' · single fabric'), fcR.switches + ' required']);
    if (nnFc.reuse > 0) {
      bom.push(['🔌 Fibre Channel', 'Existing reuse: ' + fmtInt(nnFc.reuse) + ' ports',
        'subtracted from demand', '<strong>' + nnFc.switches + ' net-new</strong>']);
    }
  }
  const bomRows = bom.map((r) => '<tr><td>' + r[0] + '</td><td>' + r[1] + '</td><td class="muted small">' + r[2] + '</td><td class="num">' + r[3] + '</td></tr>').join('');

  const ethUtilPct = Math.round(ethR.util * 100);
  const fcUtilPct = fcR.enabled ? Math.round(fcR.util * 100) : 0;
  const utilBar = (label, pct) => {
    const cls = pct > 95 ? 'red' : pct > 80 ? 'amber' : 'green';
    return '<div class="bar-row"><div class="bar-label">' + label + '</div>' +
      '<div class="bar-track"><div class="bar-fill ' + cls + '" style="width:' + Math.min(100, pct) + '%"></div></div>' +
      '<div class="bar-val">' + pct + '%</div></div>';
  };

  const stats =
    '<div class="stat"><div class="v blue grad">' + ethR.switches + '</div><div class="l">TOR switches</div></div>' +
    '<div class="stat"><div class="v purple grad">' + (fcR.enabled ? fcR.switches : '—') + '</div><div class="l">FC switches</div></div>' +
    '<div class="stat"><div class="v">' + fmtInt(ethR.servedPorts) + '</div><div class="l">Ethernet downlink ports</div></div>' +
    '<div class="stat"><div class="v ' + (ethR.over > E.overTarget ? 'amber' : 'green') + '">' + overStr(ethR.over) + '</div><div class="l">Oversubscription (target ' + E.overTarget + ':1)</div></div>' +
    (fcR.enabled ? '<div class="stat"><div class="v">' + fmtInt(fcR.devicePorts) + '</div><div class="l">FC device ports</div></div>' : '') +
    ((nnEth.reuse > 0 || nnFc.reuse > 0)
      ? '<div class="stat"><div class="v green">' + (nnEth.switches + nnFc.switches) + '</div><div class="l">Net-new switches</div></div>' : '');

  // Stacked oversubscription meter + downlink bandwidth mix (Physgun vibe: animated bars)
  const dlBW = ethR.servedBW, ulBW = ethR.uplinkBW;
  let overMeter = '';
  if (dlBW > 0) {
    const covPct = ulBW >= dlBW ? 100 : (ulBW / dlBW) * 100;
    const excPct = 100 - covPct;
    overMeter =
      '<div class="meter"><div class="m-blue" style="width:' + covPct.toFixed(1) + '%"></div>' +
      (excPct > 0.05 ? '<div class="m-red" style="width:' + excPct.toFixed(1) + '%"></div>' : '') + '</div>' +
      '<div class="meter-legend">' +
      '<span><span class="sw" style="background:#52aaff"></span>Uplink covers <strong>' + fmtInt(ulBW) + ' Gbps</strong></span>' +
      (excPct > 0.05
        ? '<span><span class="sw" style="background:#ff6b6b"></span>Downlink exceeds uplink by <strong>' + fmtInt(dlBW - ulBW) + ' Gbps</strong></span>'
        : '<span><span class="sw" style="background:#3ddc84"></span>Fully covered — <strong>' + overStr(ethR.over) + '</strong></span>') +
      '</div>';
  }
  const mixClasses = [
    ['Mgmt', P.mgmtN, P.mgmtS, ''],
    ['Data / vMotion', P.dataN, P.dataS, 'purple'],
    ['Storage IP', P.storN, P.storS, 'green'],
  ];
  const mixBars = mixClasses
    .map((c) => [c[0], c[2] <= ethR.P.dlSpeed ? P.hosts * c[1] * c[2] : 0, c[3]])
    .filter((r) => r[1] > 0)
    .map((r) =>
      '<div class="bar-row"><div class="bar-label">' + r[0] + '</div>' +
      '<div class="bar-track"><div class="bar-fill ' + r[2] + '" style="width:' + (dlBW > 0 ? Math.min(100, (r[1] / dlBW) * 100).toFixed(1) : 0) + '%"></div></div>' +
      '<div class="bar-val">' + fmtInt(r[1]) + ' Gbps</div></div>'
    ).join('');
  const bwPanel = '<div class="panel"><h3><span class="pg-glyph">🌊</span><span>Bandwidth &amp; oversubscription</span></h3>' +
    overMeter + mixBars +
    '<p class="note">Oversubscription <strong>' + overStr(ethR.over) + '</strong> vs target <strong>' + E.overTarget + ':1</strong> — downlink demand <strong>' + fmtInt(dlBW) + ' Gbps</strong>, uplink capacity <strong>' + fmtInt(ulBW) + ' Gbps</strong>. Bars animate as you re-tune on the Configure step.</p></div>';

  $('tab-plan').innerHTML =
    '<div class="stat-grid">' + stats + '</div>' +
    '<div class="panel"><h3><span class="pg-glyph">📦</span><span>Bill of materials <span class="sub">switch counts — optics, cables &amp; licenses via your VAR quoting</span></span></h3>' +
    '<div class="table-scroll"><table class="data"><thead><tr><th>Fabric</th><th>Item</th><th>Detail</th><th class="num">Qty</th></tr></thead>' +
    '<tbody>' + bomRows + '</tbody></table></div></div>' +
    '<div class="grid2">' +
    '<div class="panel"><h3><span class="pg-glyph">📊</span><span>Port utilization</span></h3>' +
      utilBar('Ethernet downlinks', ethUtilPct) +
      (fcR.enabled ? utilBar('FC usable ports / fabric', fcUtilPct) : '') +
      '<p class="note">Utilization = used ports ÷ usable ports across the sized switches. High 90s means no room for adds/moves/changes.</p></div>' +
    '<div class="panel"><h3><span class="pg-glyph">📋</span><span>Demand summary</span></h3><div class="spec-line" style="margin-top:0">' + profileSummary() + '</div>' +
      '<p class="muted small" style="margin-bottom:0">Ethernet downlink bandwidth: <strong>' + fmtInt(ethR.servedBW) + ' Gbps</strong> · uplink bandwidth: <strong>' + fmtInt(ethR.uplinkBW) + ' Gbps</strong>' +
      (fcR.enabled ? '<br>FC device ports: <strong>' + fmtInt(fcR.hostPorts) + '</strong> host + <strong>' + fmtInt(fcR.targetPorts) + '</strong> target' +
        (fcR.ratio != null ? ' (' + fmt1(fcR.ratio) + ':1 host:target)' : '') +
        ' · ISL reserve <strong>' + fmtInt(fcR.islBW) + ' Gbps</strong>/switch' : '') + '</p></div>' +
    '</div>' + bwPanel;
}

function ethWorked() {
  const { P, E, ethR } = APP.results;
  let s = '', n = 0;
  const cls = [['mgmt', P.mgmtN, P.mgmtS], ['data', P.dataN, P.dataS], ['storage IP', P.storN, P.storS]];
  cls.forEach((c) => {
    if (c[1] > 0) s += mathStep(++n, c[0] + ' ports = ' + P.hosts + ' hosts × ' + c[1] + ' × ' + c[2] + 'G', fmtInt(P.hosts * c[1]) + ' ports');
  });
  s += mathStep(++n, 'total downlink ports (served at ≤' + ethR.P.dlSpeed + 'G)',
    '<strong>' + fmtInt(ethR.servedPorts) + ' ports</strong>');
  if (ethR.unserved) s += mathStep(++n, 'ports needing &gt;' + ethR.P.dlSpeed + 'G — no matching downlink on ' + esc(ethR.P.label),
    '<strong style="color:#ff9d9d">' + fmtInt(ethR.unserved) + ' UNSERVED</strong>');
  s += mathStep(++n, 'usable downlinks / switch = ' + ethR.P.dl + (ethR.P.sharedPorts ? ' − ' + ethR.uplinks + ' uplinks (shared pool)' : ' (dedicated uplinks)'),
    fmtInt(ethR.usableDl) + (ethR.P.sharedPorts && ethR.P.dlSpeed === 100 ? ' ports / ' + fmtInt(ethR.usableDl * 4) + ' sub-ports' : ''));
  s += mathStep(++n, 'raw switches = ceil(' + fmtInt(ethR.servedPorts) + ' ÷ ' + ethR.usableDl + ')' +
    (E.dual ? ' → round to pair (A/B)' : '') + (E.spare ? ' → +1 spare' : ''),
    '<strong>' + ethR.switches + ' switches</strong>');
  s += mathStep(++n, 'downlink bandwidth = Σ ports × speed', fmtInt(ethR.servedBW) + ' Gbps');
  s += mathStep(++n, 'uplink bandwidth = ' + ethR.switches + ' × ' + ethR.uplinks + ' × ' + ethR.P.ulSpeed + 'G', fmtInt(ethR.uplinkBW) + ' Gbps');
  const overOK = ethR.over <= E.overTarget;
  s += mathStep(++n, 'oversubscription = ' + fmtInt(ethR.servedBW) + ' ÷ ' + fmtInt(ethR.uplinkBW) + ' (target ≤ ' + E.overTarget + ':1)',
    '<strong>' + overStr(ethR.over) + '</strong>' + (overOK ? '' : ' <span style="color:#f5a623">— over target</span>'), !overOK);
  return s;
}

function fcWorked() {
  const { P, F, fcR } = APP.results;
  if (!fcR.enabled) return '<p class="muted">FC fabric disabled — no SAN switches sized.</p>';
  let s = '', n = 0;
  s += mathStep(++n, 'host ports = ' + P.hosts + ' hosts × ' + P.hbaN + ' HBAs', fmtInt(fcR.hostPorts) + ' ports');
  s += mathStep(++n, 'target ports = ' + F.arrays + ' arrays × ' + F.targets + ' targets', fmtInt(fcR.targetPorts) + ' ports');
  s += mathStep(++n, 'device ports = host + target', '<strong>' + fmtInt(fcR.devicePorts) + ' ports</strong>');
  s += mathStep(++n, 'per fabric = ' + fmtInt(fcR.devicePorts) + (F.dual ? ' ÷ 2 (A/B)' : ' (single fabric)'), fmtInt(fcR.perFabric) + ' ports');
  s += mathStep(++n, 'usable / switch = ' + fcR.P.ports + ' − ' + F.isl + ' ISL reserve', fmtInt(fcR.usable) + ' ports');
  s += mathStep(++n, 'switches / fabric = ceil(' + fmtInt(fcR.perFabric) + ' ÷ ' + fcR.usable + ')' + (F.dual ? ' × 2 fabrics' : ''),
    '<strong>' + fcR.switches + ' switches</strong>', true);
  s += mathStep(++n, 'ISL reserve bandwidth / switch = ' + F.isl + ' × ' + fcR.effSpeed + 'G', fmtInt(fcR.islBW) + ' Gbps');
  if (fcR.ratio != null) s += mathStep(++n, 'host:target ratio = ' + fmtInt(fcR.hostPorts) + ' ÷ ' + fmtInt(fcR.targetPorts), fmt1(fcR.ratio) + ':1');
  return s;
}

function renderFabricsTab() {
  const { fcR } = APP.results;
  $('tab-fabrics').innerHTML =
    '<div class="panel"><h3><span class="pg-glyph">🌐</span><span>Ethernet fabric — worked math <span class="sub">' + esc(APP.results.ethR.P.label) + '</span></span></h3>' + ethWorked() + '</div>' +
    '<div class="panel"><h3><span class="pg-glyph">🔌</span><span>FC fabric — worked math <span class="sub">' + (fcR.enabled ? esc(fcR.P.label) : 'disabled') + '</span></span></h3>' + fcWorked() + '</div>';
}

/* ================= Findings ================= */
function buildFindings() {
  const { P, E, F, ethR, fcR, nnEth, nnFc } = APP.results;
  const Fl = [];
  const add = (sev, icon, title, body) => Fl.push({ sev, icon, title, body });

  add('info', '📋', 'Plan totals',
    '<strong>' + ethR.switches + '×</strong> ' + esc(ethR.P.label) + ' for <strong>' + fmtInt(P.hosts) + '</strong> hosts' +
    (fcR.enabled ? ' and <strong>' + fcR.switches + '×</strong> ' + esc(fcR.P.label) + ' across Fabric A/B' : ' — no FC fabric') +
    '. Tune anything on the Configure step — this page recomputes when you come back.');

  // Ethernet oversubscription
  if (ethR.over > E.overTarget) {
    const needUl = Math.ceil(ethR.servedBW / E.overTarget / Math.max(1, ethR.switches) / ethR.P.ulSpeed);
    add('warn', '📊', 'Oversubscription ' + overStr(ethR.over) + ' exceeds the ' + E.overTarget + ':1 target',
      'Downlink demand is ' + fmtInt(ethR.servedBW) + ' Gbps against ' + fmtInt(ethR.uplinkBW) + ' Gbps of uplink. ' +
      (ethR.uplinks < (ethR.P.ul || 6) + 6
        ? 'Raising uplinks toward ~' + needUl + '/switch would bring it to target — or accept the ratio and note it as a design decision.'
        : 'All uplink ports are committed — the lever left is a lower oversubscription target acceptance or fewer downlink ports per host.'));
  } else {
    add('info', '✅', 'Oversubscription ' + overStr(ethR.over) + ' within target',
      fmtInt(ethR.servedBW) + ' Gbps of downlink against ' + fmtInt(ethR.uplinkBW) + ' Gbps uplink. East-west-heavy estates (vMotion, vSAN) may still want headroom — the ratio is capacity, not a traffic model.');
  }

  if (!E.dual && ethR.switches > 0) {
    add('crit', '🔴', 'No A/B dual-homing — single TOR failure takes the rack',
      'Every host is single-homed. One switch failure or firmware upgrade takes down all ' + fmtInt(P.hosts) + ' hosts\' network paths. Enable dual-homing unless this is a deliberately non-redundant edge design.');
  }
  if (E.spare) add('info', '🛟', 'N+1 spare included', 'One extra TOR covers a switch failure without re-cabling. Remember the spare needs its own optics and uplink ports to be useful on day one.');
  if (ethR.unserved > 0) {
    add('crit', '🔴', fmtInt(ethR.unserved) + ' ports need faster downlinks than ' + ethR.P.dlSpeed + 'G',
      'Part of the host profile runs above the TOR downlink speed — those connections have no home on ' + esc(ethR.P.label) + '. Move to a faster preset (e.g. 100G TOR) or split the fast hosts onto their own switches.');
  }
  if (ethR.mgmt1G) {
    add('info', '💡', '1G mgmt on ' + ethR.P.dlSpeed + 'G TOR ports',
      'It works — SFP28 ports accept 1G — but you\'re burning ' + fmtInt(P.hosts * P.mgmtN) + ' high-speed ports on mgmt. Many shops hang mgmt/IPMI off a cheap dedicated 1G switch instead; worth pricing both ways.');
  }
  const utilPct = Math.round(ethR.util * 100);
  if (utilPct >= 95 && ethR.switches > 0) {
    add('warn', '🧯', 'TOR ports ' + utilPct + '% full — no room for adds/moves',
      'The switches are essentially full on day one. Any host added later means another switch (or pair). Consider sizing +1 switch of headroom now — it\'s cheaper than a return visit.');
  } else if (utilPct < 50 && ethR.switches > 0) {
    add('info', '🪶', 'TOR ports only ' + utilPct + '% utilized',
      'Headroom is comfortable. If the estate won\'t grow, a smaller/cheaper TOR SKU could do the same job — or keep the headroom and call it growth capacity in the proposal.');
  }
  if (P.storN === 0 && !(fcR.enabled && P.hbaN > 0)) {
    add('warn', '🗄️', 'No storage network modeled',
      'Zero storage-IP ports and no FC HBAs — the hosts have no storage path in this plan. If storage is local/vSAN or out of scope, fine; otherwise add the storage class back.');
  }
  if (ethR.P.sharedPorts && E.breakout && (P.dataS <= 25 || P.storS <= 25)) {
    add('info', '🔀', 'Breakout is doing real work here',
      '≤25G host ports ride 4-per-100G-port via breakout — that\'s what keeps the 64-port switch count down. Budget breakout cables/DACs accordingly; they\'re a line item, not free.');
  }
  if (ethR.P.sharedPorts && !E.breakout) {
    add('warn', '🔀', 'Breakout off: ≤25G hosts burn whole 100G ports',
      'Without breakout each 25G host connection consumes a full 100G port. Enable 4×25G breakout (or accept the waste) — the switch count above assumes no sharing.');
  }

  // FC findings
  if (fcR.enabled) {
    if (!F.dual) {
      add('crit', '🔴', 'Single FC fabric — no path redundancy',
        'One fabric means one ISL failure or switch outage away from losing storage paths. Dual Fabric A/B is standard SAN posture — the tool sizes pairs by default for a reason.');
    }
    if (fcR.hbaCapped) {
      add('warn', '⚠️', 'HBA speed (' + P.hbaS + 'G) exceeds switch max (' + fcR.P.maxSpeed + 'G)',
        'Links will negotiate down to ' + fcR.P.maxSpeed + 'G. Either the HBAs are over-spec\'d for ' + esc(fcR.P.label) + ' or the fabric should move to a 64G-capable switch.');
    }
    if (fcR.targetCapped) {
      add('warn', '⚠️', 'Array target speed (' + F.targetSpeed + 'G) exceeds switch max (' + fcR.P.maxSpeed + 'G)',
        'Array front-ends will run at ' + fcR.P.maxSpeed + 'G. Match the fabric to the arrays, not just the hosts.');
    }
    if (fcR.oddHba && P.hbaN > 0) {
      add('warn', '⚖️', 'Odd HBA count (' + P.hbaN + ') with dual fabrics',
        'HBAs don\'t split evenly across Fabric A/B — one fabric carries more paths. Prefer 2 or 4 HBAs per host for symmetric fabrics.');
    }
    if (fcR.ratio != null && fcR.ratio > 7) {
      add('warn', '📊', 'Host:target ratio ' + fmt1(fcR.ratio) + ':1 is aggressive',
        fmtInt(fcR.hostPorts) + ' host ports against ' + fmtInt(fcR.targetPorts) + ' array targets. Past ~7:1, target ports become the congestion point under burst — consider more targets per array.');
    }
    if (F.arrays === 0) {
      add('warn', '🗄️', 'FC fabric sized with zero storage arrays',
        'The fabric carries host HBAs but no array targets are defined. Add the arrays (count + target ports each) so the device-port math reflects reality.');
    }
    add('info', '🔌', 'ISL reserve: ' + fmtInt(fcR.islBW) + ' Gbps per switch',
      F.isl + ' ports × ' + fcR.effSpeed + 'G held for inter-switch links on each of the ' + fcR.switches + ' switches. That\'s a port reservation, not a bandwidth design — validate against expected east-west traffic.');
  } else {
    add('info', '🔌', 'No FC fabric sized',
      P.hbaN > 0 ? 'FC fabric toggle is off — HBAs are defined but no SAN switches are in the plan.' : 'All-IP storage design: no HBAs, no SAN switches. The Ethernet fabric carries storage traffic — keep an eye on that oversubscription ratio.');
  }

  // Net-new
  if (nnEth.reuse > 0 || nnFc.reuse > 0) {
    const parts = [];
    if (nnEth.reuse > 0) parts.push('<strong>' + nnEth.switches + ' net-new TOR</strong> (' + fmtInt(nnEth.reuse) + ' existing ports reused)');
    if (nnFc.reuse > 0) parts.push('<strong>' + nnFc.switches + ' net-new FC</strong> (' + fmtInt(nnFc.reuse) + ' existing ports reused)');
    add('info', '♻️', 'Existing inventory offsets the buy', parts.join(' · ') + '. Assumes existing ports are reusable at the needed speeds — verify against installed optics before committing.');
  }
  return Fl;
}

function renderFindingsTab() {
  const Fl = buildFindings();
  $('tab-findings').innerHTML = '<div class="panel"><h3><span class="pg-glyph">💡</span><span>SE talking points <span class="sub">' + Fl.length + ' findings</span></span></h3>' +
    Fl.map((f) => '<div class="finding ' + f.sev + '"><div class="sev">' + f.icon + '</div><div><strong>' + f.title + '</strong><p>' + f.body + '</p></div></div>').join('') + '</div>';
}

function renderReportTab() {
  $('tab-report').innerHTML = '<div class="panel"><h3><span class="pg-glyph">📄</span><span>Customer-ready briefing</span></h3>' +
    '<p class="muted">Generates a standalone HTML report — BOM, worked fabric math, oversubscription, findings, and methodology. Self-contained (no external dependencies), safe to email. Scrub customer names from group labels first if it leaves your org.</p>' +
    '<div class="toolbar"><button class="btn primary" id="dlReportBtn2">⬇ Download HTML report</button></div></div>';
  $('dlReportBtn2').onclick = downloadReport;
}

/* ================= Report ================= */
function buildReportHTML() {
  const { P, E, F, ethR, fcR, nnEth, nnFc } = APP.results;
  const date = new Date().toISOString().slice(0, 10);
  const srcLbl = APP.source === 'rvtools' ? 'RVTools export (' + APP.fileName + ')' : APP.source === 'demo' ? 'Demo data (synthetic)' : 'Manual entry';
  const Fl = buildFindings();
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
    '.step{display:flex;gap:10px;padding:7px 0;border-bottom:1px dashed #e6ebf1;font-size:.9rem}' +
    '.step:last-child{border-bottom:none}' +
    '.step .n{flex:none;width:24px;height:24px;border-radius:50%;background:#eef2f8;color:#5b6572;font-size:.72rem;display:flex;align-items:center;justify-content:center}' +
    '.step .f{flex:1;color:#5b6572}.step .r{flex:none;font-weight:700}' +
    '.disclaimer{background:#fff8ec;border:1px solid #f0d9a8;border-radius:10px;padding:14px 18px;margin-top:2rem;font-size:.9rem}' +
    '@media print{.wrap{padding:0}}';

  const bomRows = [
    ['Ethernet', ethR.switches + '× ' + ethR.P.label, ethR.P.note || '', ethR.switches + ' required'],
  ];
  if (nnEth.reuse > 0) bomRows.push(['Ethernet', 'Existing reuse: ' + fmtInt(nnEth.reuse) + ' ports', 'subtracted from demand', nnEth.switches + ' net-new']);
  if (fcR.enabled) {
    bomRows.push(['Fibre Channel', fcR.switches + '× ' + fcR.P.label, (fcR.P.note || '') + (F.dual ? ' · Fabric A/B' : ''), fcR.switches + ' required']);
    if (nnFc.reuse > 0) bomRows.push(['Fibre Channel', 'Existing reuse: ' + fmtInt(nnFc.reuse) + ' ports', 'subtracted from demand', nnFc.switches + ' net-new']);
  }

  const ethSteps = [
    ['Host port profile', P.hosts + ' hosts × (' + P.mgmtN + '×' + P.mgmtS + 'G mgmt + ' + P.dataN + '×' + P.dataS + 'G data + ' + P.storN + '×' + P.storS + 'G storage IP)', fmtInt(ethR.servedPorts) + ' downlink ports'],
    ['Usable downlinks / switch', ethR.P.dl + (ethR.P.sharedPorts ? ' − ' + ethR.uplinks + ' uplinks (shared pool)' : ' (dedicated uplinks)'), fmtInt(ethR.usableDl)],
    ['Switches', 'ceil(' + fmtInt(ethR.servedPorts) + ' ÷ ' + ethR.usableDl + ')' + (E.dual ? ' → A/B pair' : '') + (E.spare ? ' → +1 spare' : ''), '<strong>' + ethR.switches + ' switches</strong>'],
    ['Oversubscription', fmtInt(ethR.servedBW) + ' Gbps downlink ÷ ' + fmtInt(ethR.uplinkBW) + ' Gbps uplink (target ≤ ' + E.overTarget + ':1)', '<strong>' + overStr(ethR.over) + '</strong>'],
  ];
  const fcSteps = fcR.enabled ? [
    ['Device ports', P.hosts + '×' + P.hbaN + ' HBAs + ' + F.arrays + '×' + F.targets + ' array targets', fmtInt(fcR.devicePorts) + ' ports'],
    ['Per fabric', F.dual ? '÷ 2 (A/B)' : 'single fabric', fmtInt(fcR.perFabric) + ' ports'],
    ['Usable / switch', fcR.P.ports + ' − ' + F.isl + ' ISL reserve', fmtInt(fcR.usable) + ' ports'],
    ['Switches', 'ceil(' + fmtInt(fcR.perFabric) + ' ÷ ' + fcR.usable + ')' + (F.dual ? ' × 2' : ''), '<strong>' + fcR.switches + ' switches</strong>'],
  ] : [['FC fabric', 'disabled', '—']];

  const stepRows = (steps) => steps.map((x, i) =>
    '<div class="step"><div class="n">' + (i + 1) + '</div><div class="f">' + x[0] + ' — <span class="mono">' + x[1] + '</span></div><div class="r">' + x[2] + '</div></div>').join('');
  const findings = Fl.map((f) => '<div class="finding ' + f.sev + '"><strong>' + f.title + '</strong><br><span style="color:#5b6572">' + f.body + '</span></div>').join('');

  return '<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>Network Sizer — Network Plan (' + date + ')</title><style>' + css + '</style></head><body><div class="wrap">' +
    '<h1>Network Plan — TOR &amp; FC SAN</h1><div class="meta">Generated ' + date + ' · Source: ' + esc(srcLbl) + ' · Network Sizer (client-side sizing tool)</div>' +
    '<div><div class="stat"><div class="v">' + ethR.switches + '</div><div class="l">TOR switches</div></div>' +
    '<div class="stat"><div class="v">' + (fcR.enabled ? fcR.switches : '—') + '</div><div class="l">FC switches</div></div>' +
    '<div class="stat"><div class="v">' + overStr(ethR.over) + '</div><div class="l">Oversubscription</div></div>' +
    '<div class="stat"><div class="v">' + fmtInt(P.hosts) + '</div><div class="l">Target hosts</div></div></div>' +
    '<h2>1. Bill of materials</h2><table><thead><tr><th>Fabric</th><th>Item</th><th>Detail</th><th class="num">Qty</th></tr></thead><tbody>' +
    bomRows.map((r) => '<tr><td>' + r[0] + '</td><td>' + r[1] + '</td><td style="color:#5b6572;font-size:.85rem">' + r[2] + '</td><td class="num">' + r[3] + '</td></tr>').join('') +
    '</tbody></table>' +
    '<h2>2. Ethernet fabric — worked math</h2><h3>' + esc(ethR.P.label) + '</h3>' + stepRows(ethSteps) +
    '<h2>3. FC fabric — worked math</h2><h3>' + (fcR.enabled ? esc(fcR.P.label) : 'Disabled') + '</h3>' + stepRows(fcSteps) +
    '<h2>4. Findings</h2>' + findings +
    '<h2>5. Methodology</h2><p style="color:#5b6572;font-size:.9rem">Ethernet: downlink ports = target hosts × per-host port profile (mgmt + data/vMotion + storage IP); TOR switches = ceil(downlink ports ÷ usable downlinks per switch), rounded to pairs for A/B dual-homing with optional N+1 spare; oversubscription = total downlink bandwidth ÷ total uplink bandwidth. FC: device ports = host HBAs + array target ports, split across Fabric A/B; usable ports per switch = total ports − ISL reserve; switches per fabric = ceil(device ports per fabric ÷ usable), total × 2. Host counts from RVTools vHost or manual entry; VMs are context only. This is port-count sizing, not traffic modeling: it does not model per-workload bandwidth, burst behavior, or latency. Optics, cables, licenses, and the spine layer are out of scope.</p>' +
    '<div class="disclaimer">⚠️ <strong>Indicative analysis, not a quote.</strong> Switch pricing, optics, licensing, and partner programs affect real cost. Validate all figures against an official quote before committing to purchases.</div>' +
    '</div></body></html>';
}

function downloadReport() {
  if (!APP.results) return;
  const html = buildReportHTML();
  const blob = new Blob([html], { type: 'text/html' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'network-sizer-plan-' + new Date().toISOString().slice(0, 10) + '.html';
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
}

/* ================= Wiring ================= */
function clearSession() {
  APP.groups = []; APP.ethInv = []; APP.fcInv = [];
  APP.prof = null; APP.eth = null; APP.fc = null;
  APP.source = null; APP.fileName = null; APP.results = null;
  APP.projectName = 'Untitled project';
  clearAutosave();
  updateProjName(); $('projSaved').textContent = '';
  $('inventoryWrap').hidden = true;
  $('existingWrap').hidden = true;
  $('manualEditor').hidden = true;
  $('manualBody').innerHTML = '';
  $('fileInput').value = '';
  clearMsgs();
  $('wizard').hidden = true;
  $('landing').hidden = false;
  window.scrollTo({ top: 0 });
}


/* ================= Projects: save / load / export / import ================= */
const LS_AUTO = 'network-sizer:autosave';
const LS_PROJECTS = 'network-sizer:projects';
const PROJECT_VERSION = 1;

function serializeState() {
  return {
    groups: APP.groups, ethInv: APP.ethInv, fcInv: APP.fcInv,
    prof: APP.prof, eth: APP.eth, fc: APP.fc,
    source: APP.source, fileName: APP.fileName,
  };
}
function projectEnvelope(name, state) {
  return {
    app: 'network-sizer', version: PROJECT_VERSION,
    name: (name || 'Untitled project').slice(0, 60),
    savedAt: new Date().toISOString(),
    state: state || serializeState(),
  };
}
function validProject(d) {
  return !!(d && d.app === 'network-sizer' && d.state &&
    Array.isArray(d.state.groups) && typeof d.version === 'number' && d.version <= PROJECT_VERSION);
}
function slugify(s) { return String(s || 'project').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 40) || 'project'; }
function fmtTime(iso) {
  try { return new Date(iso).toLocaleString(undefined, { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' }); }
  catch (e) { return ''; }
}

let toastT = null;
function showToast(html, ms) {
  const t = $('projToast');
  t.innerHTML = html; t.hidden = false;
  clearTimeout(toastT);
  toastT = setTimeout(() => { t.hidden = true; }, ms || 6000);
}
function updateProjName() { $('projName').textContent = APP.projectName || 'Untitled project'; }

function applyProject(env) {
  const s = env.state || {};
  APP.groups = Array.isArray(s.groups) ? s.groups : [];
  APP.ethInv = Array.isArray(s.ethInv) ? s.ethInv : [];
  APP.fcInv = Array.isArray(s.fcInv) ? s.fcInv : [];
  APP.prof = s.prof || null; APP.eth = s.eth || null; APP.fc = s.fc || null;
  APP.source = s.source || null; APP.fileName = s.fileName || null; APP.results = null;
  APP.projectName = env.name || 'Untitled project';
  if (!APP.groups.length) { showToast('That project has no demand data \u2014 nothing to restore.'); return; }
  clearMsgs();
  $('landing').hidden = true; $('wizard').hidden = false;
  renderInventory();
  renderInvTable('eth'); renderInvTable('fc');
  if (APP.prof && APP.eth && APP.fc) renderConfig();
  setStep(1);
  updateProjName();
  queueAutosave();
}

/* ---- autosave (this browser only) ---- */
let autosaveT = null;
function queueAutosave() { clearTimeout(autosaveT); autosaveT = setTimeout(autosaveNow, 900); }
function autosaveNow() {
  if (!APP.groups.length) return;
  try {
    localStorage.setItem(LS_AUTO, JSON.stringify(projectEnvelope(APP.projectName, serializeState())));
    $('projSaved').textContent = '\u00B7 autosaved ' + fmtTime(new Date().toISOString());
  } catch (e) { /* private mode / quota — non-fatal */ }
}
function clearAutosave() { try { localStorage.removeItem(LS_AUTO); } catch (e) {} }

/* ---- named projects (this browser only) ---- */
function getProjects() { try { return JSON.parse(localStorage.getItem(LS_PROJECTS) || '[]'); } catch (e) { return []; } }
function setProjects(list) { try { localStorage.setItem(LS_PROJECTS, JSON.stringify(list.slice(0, 30))); } catch (e) {} }
function renderProjList() {
  const list = getProjects();
  const box = $('projList');
  if (!list.length) { box.innerHTML = '<p class="muted" style="font-size:.85rem">No saved projects yet \u2014 name it above and hit <strong>Save project</strong>.</p>'; return; }
  box.innerHTML = list.map((p) => {
    const hosts = p.state && p.state.groups ? p.state.groups.reduce((a, g) => a + (g.hosts || 0), 0) : 0;
    return '<div class="proj-item"><div><div class="nm">' + esc(p.name || 'Untitled project') + '</div>' +
      '<div class="meta">saved ' + esc(fmtTime(p.savedAt)) + ' \u00B7 ' + fmtInt(hosts) + ' hosts</div></div>' +
      '<div class="ops"><button class="btn ghost" data-load="' + p.id + '">Load</button>' +
      '<button class="btn danger-ghost" data-delp="' + p.id + '">Delete</button></div></div>';
  }).join('');
  box.querySelectorAll('[data-load]').forEach((b) => { b.onclick = () => {
    const p = getProjects().find((x) => x.id === b.dataset.load);
    if (p && validProject(p)) { $('projPanel').hidden = true; applyProject(p); showToast('Loaded project <strong>' + esc(p.name || '') + '</strong>.'); }
    else showToast('Could not load that project \u2014 the saved data looks invalid.');
  }; });
  box.querySelectorAll('[data-delp]').forEach((b) => { b.onclick = () => {
    setProjects(getProjects().filter((x) => x.id !== b.dataset.delp));
    renderProjList();
  }; });
}
function saveNamedProject() {
  if (!APP.groups.length) { showToast('Load some demand first \u2014 there is nothing to save yet.'); return; }
  const input = $('projNameInput').value.trim();
  const name = (input || APP.projectName || 'Untitled project').slice(0, 60);
  const list = getProjects();
  const env = projectEnvelope(name, serializeState());
  env.id = 'p' + Date.now().toString(36);
  const ix = list.findIndex((p) => (p.name || '') === name);
  if (ix >= 0) { env.id = list[ix].id; list[ix] = env; } else list.unshift(env);
  setProjects(list);
  APP.projectName = name; updateProjName();
  $('projNameInput').value = '';
  renderProjList();
  showToast('Project <strong>' + esc(name) + '</strong> saved in this browser.');
  queueAutosave();
}

/* ---- export / import (.json) ---- */
function exportProject() {
  if (!APP.groups.length) { showToast('Load some demand first \u2014 there is nothing to export yet.'); return; }
  const env = projectEnvelope(APP.projectName, serializeState());
  const blob = new Blob([JSON.stringify(env, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'network-sizer-' + slugify(env.name) + '.json';
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  showToast('Exported <strong>' + esc(a.download) + '</strong> \u2014 keep it with the engagement files.');
}
function importProjectFile(file) {
  if (!file) return;
  const rd = new FileReader();
  rd.onload = () => {
    try {
      const d = JSON.parse(rd.result);
      if (!validProject(d)) { showToast('<strong>Not a network-sizer project file.</strong> Pick a JSON exported from this app.'); return; }
      $('projPanel').hidden = true;
      applyProject(d);
      showToast('Imported project <strong>' + esc(d.name || 'Untitled') + '</strong>.');
    } catch (e) { showToast('<strong>Could not read that file.</strong> ' + esc(e.message || '')); }
  };
  rd.readAsText(file);
}

function wireProjects() {
  const toggle = () => {
    const p = $('projPanel');
    p.hidden = !p.hidden;
    if (!p.hidden) {
      $('projNameInput').value = APP.projectName === 'Untitled project' ? '' : APP.projectName;
      renderProjList();
      $('projNameInput').focus();
    }
  };
  $('projBtn').onclick = toggle;
  $('projDoSave').onclick = saveNamedProject;
  $('projNameInput').addEventListener('keydown', (e) => { if (e.key === 'Enter') saveNamedProject(); });
  $('projExportBtn').onclick = exportProject;
  $('projImportBtn').onclick = () => $('projImportFile').click();
  $('projImportFile').addEventListener('change', (e) => { importProjectFile(e.target.files[0]); e.target.value = ''; });
  // restore last session, if any
  try {
    const raw = localStorage.getItem(LS_AUTO);
    if (raw) {
      const d = JSON.parse(raw);
      if (validProject(d) && d.state.groups.length) {
        applyProject(d);
        showToast('Restored your last session \u2014 <strong>' + esc(d.name || '') + '</strong> &nbsp;·&nbsp; <a id="toastFresh">start fresh</a>', 5000);
        const f = $('toastFresh');
        if (f) f.onclick = () => { clearSession(); $('projToast').hidden = true; };
      }
    }
  } catch (e) { /* corrupted autosave — start clean */ }
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
  wireProjects();
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

  $('ethInvAdd').onclick = () => { APP.ethInv.push({ preset: 'nx93180', count: 0, cDl: 48 }); renderInvTable('eth'); queueAutosave(); };
  $('fcInvAdd').onclick = () => { APP.fcInv.push({ preset: 'g720', count: 0, cPorts: 64 }); renderInvTable('fc'); queueAutosave(); };

  $('backToStartBtn').onclick = () => {
    APP.groups = []; APP.ethInv = []; APP.fcInv = []; APP.prof = null; APP.eth = null; APP.fc = null;
    APP.source = null;
    clearAutosave(); APP.projectName = 'Untitled project'; updateProjName(); $('projSaved').textContent = '';
    $('inventoryWrap').hidden = true; $('existingWrap').hidden = true;
    clearMsgs(); window.scrollTo({ top: 0, behavior: 'smooth' });
  };
  $('toConfigBtn').onclick = () => { if (!APP.groups.length) return; ensureCfg(); renderConfig(); setStep(2); };
  $('backToDataBtn').onclick = () => setStep(1);
  $('toResultsBtn').onclick = () => { if (!APP.groups.length) return; ensureCfg(); renderResults(); setStep(3); };
  $('backToConfigBtn').onclick = () => setStep(2);

  ['gHosts', 'pMgmtN', 'pMgmtS', 'pDataN', 'pDataS', 'pStorN', 'pStorS', 'pHbaN', 'pHbaS'].forEach((id) => {
    $(id).addEventListener('input', () => { readProfileInputs(); refreshPreviews(); paintSlider($(id)); queueAutosave(); });
    $(id).addEventListener('change', () => { readProfileInputs(); if (id === 'gHosts') $('gHosts').value = APP.prof.hosts; refreshPreviews(); paintSlider($(id)); queueAutosave(); });
  });

  document.querySelectorAll('#tabs button').forEach((b) => b.addEventListener('click', () => switchTab(b.dataset.tab)));
  $('printBtn').onclick = () => window.print();
  $('dlReportBtn').onclick = downloadReport;
  $('clearBtn').onclick = clearSession;
  // scroll-reveal on landing (Physgun vibe) — only when IntersectionObserver exists
  if (typeof window !== 'undefined' && 'IntersectionObserver' in window) {
    const io = new IntersectionObserver((es) => es.forEach((e) => {
      if (e.isIntersecting) { e.target.classList.add('in'); io.unobserve(e.target); }
    }), { threshold: 0.1 });
    document.querySelectorAll('#landing .cards3 .card, #landing .trust, #landing .faq-list details').forEach((el) => {
      el.classList.add('reveal'); io.observe(el);
    });
  }
  renderChangelog();
}

if (typeof document !== 'undefined') document.addEventListener('DOMContentLoaded', wireApp);
