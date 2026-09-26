// Node smoke tests for network-sizer: sizing math (unchanged by the Physgun
// reskin) + the new slider/seg helpers. Run: node tests/sizing.test.mjs
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import assert from 'node:assert/strict';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const src = readFileSync(join(root, 'js/app.js'), 'utf8');
// app.js is a plain script; its trailing DOMContentLoaded hook is guarded by
// `typeof document !== 'undefined'`, so it loads cleanly in node.
const api = new Function(
  src + '\nreturn { sizeEthernet, sizeFC, sliderFill, paintSlider, segCfg, pgRange, overStr, defaultProf, defaultEth, defaultFc, resolveTor };'
)();

let passed = 0;
const t = (name, fn) => { fn(); passed++; console.log('ok -', name); };

/* ---- worked example from the landing page: 12 hosts, 2x1G mgmt + 2x25G data + 2x25G stor, nx93180 ---- */
t('landing worked example: ethernet math', () => {
  const prof = { hosts: 12, mgmtN: 2, mgmtS: 1, dataN: 2, dataS: 25, storN: 2, storS: 25, hbaN: 0, hbaS: 32 };
  const eth = { preset: 'nx93180', cDl: 48, cDlS: 25, cUlS: 100, uplinks: 6, overTarget: 3, dual: true, spare: false, breakout: true };
  const r = api.sizeEthernet(12, prof, eth);
  assert.equal(r.servedPorts, 72);
  assert.equal(r.usableDl, 48);
  assert.equal(r.switches, 2);
  assert.ok(Math.abs(r.util - 0.75) < 1e-9, 'util 75%');
  assert.equal(r.servedBW, 1224);
  assert.equal(r.uplinkBW, 1200);
  assert.ok(Math.abs(r.over - 1.02) < 0.005, 'over ~1.02:1');
});

/* ---- FC side of the same example: 2 HBAs/host, 2 arrays x 4 targets, G720 ---- */
t('landing worked example: FC math', () => {
  const fc = { enabled: true, preset: 'g720', cPorts: 64, cSpeed: 64, isl: 8, arrays: 2, targets: 4, targetSpeed: 32, dual: true };
  const r = api.sizeFC(12, 2, 32, fc);
  assert.equal(r.devicePorts, 32);
  assert.equal(r.perFabric, 16);
  assert.equal(r.usable, 56);
  assert.equal(r.switches, 2);
  assert.equal(r.islBW, 256);
});

/* ---- shared-pool preset: 9364C-GX seeds 4 uplinks, math still works; 0 uplinks = infinite over ---- */
t('nx9364 shared pool: 4 seeded uplinks', () => {
  const prof = { hosts: 12, mgmtN: 2, mgmtS: 1, dataN: 2, dataS: 25, storN: 2, storS: 25, hbaN: 0, hbaS: 32 };
  const eth = { preset: 'nx9364', cDl: 64, cDlS: 100, cUlS: 100, uplinks: 4, overTarget: 3, dual: true, spare: false, breakout: true };
  const P = api.resolveTor(eth);
  assert.equal(P.sharedPorts, true);
  const r = api.sizeEthernet(12, prof, eth);
  assert.equal(r.usableDl, 60);
  assert.equal(r.switches, 2);
  assert.ok(isFinite(r.over), 'finite oversubscription with seeded uplinks');
});
t('nx9364 with 0 uplinks: oversubscription is infinite (the bug the seed fix avoids)', () => {
  const prof = { hosts: 12, mgmtN: 2, mgmtS: 1, dataN: 2, dataS: 25, storN: 2, storS: 25, hbaN: 0, hbaS: 32 };
  const eth = { preset: 'nx9364', cDl: 64, cDlS: 100, cUlS: 100, uplinks: 0, overTarget: 3, dual: true, spare: false, breakout: true };
  const r = api.sizeEthernet(12, prof, eth);
  assert.equal(r.over, Infinity);
  assert.equal(api.overStr(r.over), '∞ (no uplinks)');
});

/* ---- pairing / spare rounding unchanged ---- */
t('dual-homing rounds odd raw counts to pairs', () => {
  const prof = { hosts: 20, mgmtN: 2, mgmtS: 1, dataN: 2, dataS: 25, storN: 2, storS: 25, hbaN: 0, hbaS: 32 };
  const eth = { preset: 'nx93180', cDl: 48, cDlS: 25, cUlS: 100, uplinks: 6, overTarget: 3, dual: true, spare: false, breakout: true };
  const r = api.sizeEthernet(20, prof, eth);
  assert.equal(r.switchesRaw, 3); // 120 ports / 48
  assert.equal(r.switches, 4);    // rounded to pair
});
t('spare adds one switch', () => {
  const prof = { hosts: 12, mgmtN: 2, mgmtS: 1, dataN: 2, dataS: 25, storN: 2, storS: 25, hbaN: 0, hbaS: 32 };
  const eth = { preset: 'nx93180', cDl: 48, cDlS: 25, cUlS: 100, uplinks: 6, overTarget: 3, dual: true, spare: true, breakout: true };
  assert.equal(api.sizeEthernet(12, prof, eth).switches, 3);
});
t('FC disabled when no HBAs', () => {
  const fc = api.defaultFc();
  assert.equal(api.sizeFC(12, 0, 32, fc).enabled, false);
});

/* ---- slider fill helper (drives the --fill CSS var) ---- */
t('sliderFill percentages', () => {
  const el = (min, max, value) => ({ min: String(min), max: String(max), value: String(value) });
  assert.equal(api.sliderFill(el(0, 12, 6)), 50);
  assert.equal(api.sliderFill(el(1, 8, 1)), 0);
  assert.equal(api.sliderFill(el(0, 32, 32)), 100);
  assert.ok(Math.abs(api.sliderFill(el(1, 8, 3.5)) - (2.5 / 7) * 100) < 1e-9);
});
t('paintSlider sets --fill only on .pg-slider elements', () => {
  const seen = {};
  const good = { min: '0', max: '12', value: '6', classList: { contains: (c) => c === 'pg-slider' }, style: { setProperty: (k, v) => { seen[k] = v; } } };
  api.paintSlider(good);
  assert.equal(seen['--fill'], '50%');
  let threw = false;
  try { api.paintSlider({ classList: { contains: () => false }, style: { setProperty: () => { throw new Error('must not paint'); } } }); }
  catch (e) { threw = true; }
  assert.equal(threw, false);
  api.paintSlider(null); // no-op, no throw
});

/* ---- segmented control markup ---- */
t('segCfg emits data-segval + active button', () => {
  const html = api.segCfg('e', 'dual', [['1', 'Yes'], ['0', 'No']], '1');
  assert.ok(html.includes('data-ecfg="dual"'));
  assert.ok(html.includes('data-segval="1"'));
  assert.ok(html.includes('data-val="1" class="active"'));
  const html2 = api.segCfg('f', 'enabled', [['1', 'On'], ['0', 'Off']], '0');
  assert.ok(html2.includes('data-val="0" class="active"'));
});
t('pgRange emits slider + badge + scale', () => {
  const html = api.pgRange('uplinks', 'e', 0, 12, 1, 6, 'Uplinks / switch', ['0', '6', '12'], 'note', null);
  assert.ok(html.includes('class="pg-slider"'));
  assert.ok(html.includes('data-ecfg="uplinks"'));
  assert.ok(html.includes('data-lb="uplinks"'));
  assert.ok(html.includes('pg-scale'));
});

console.log(`\n${passed} tests passed.`);
