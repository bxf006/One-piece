const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const C = require('../src/core.js');
const S = require('../src/sheet-reader.js');
const R = require('../src/config-store.js');
const documents = () => Object.fromEntries(C.CONFIG_TYPES.map(type => [type, JSON.parse(fs.readFileSync(path.join(__dirname, '../configs', type + '.json'), 'utf8'))]));
function browserCore() {
  const context = vm.createContext({ TextEncoder });
  for (const file of ['core.js', 'config-store.js']) vm.runInContext(fs.readFileSync(path.join(__dirname, '../src', file), 'utf8'), context);
  return { core: context.CurveCore, store: context.CurveConfigStore };
}

test('browser core has no built-in default configuration before the JSON documents load', () => {
  const { core } = browserCore();
  assert.equal(core.defaultsReady, false); assert.equal(core.DEFAULT_HEADER_RULE, undefined);
  assert.equal(core.RANGES.length, 0); assert.equal(core.PALETTE.length, 0);
  assert.throws(() => core.readSheet(S.readPasted('时间\t值\n0\t1'), '粘贴数据', S), /默认配置尚未载入/);
});

test('changing only JSON defaults changes grouping, regions, padding, headers, time and colors', async () => {
  const files = documents(), { core, store } = browserCore();
  files.axes.defaults.timeMode = 'text';
  files.axes.defaults.axisGroups.find(g => g.id === 'small').max = 20;
  files.axes.defaults.axisGroups.find(g => g.id === 'small').label = '自定义小范围';
  files.axes.defaults.axisRegions.small = { start: 15, end: 35 };
  files.axes.defaults.autoBounds.paddingPercent = 10;
  files.regex.defaults.headerRule = { enabled: true, pattern: '^PREFIX/(.*)$', flags: '', replacement: '改名_$1' };
  files.selection.defaults.palette = ['#123456'];
  files.selection.defaults.fallbackColor = { hueStep: 90, saturation: 50, lightness: 40 };
  const loaded = await store.loadCloud(async url => ({ ok: true, status: 200, text: async () => JSON.stringify(files[url.split('/').at(-1).replace('.json', '')]) }));
  core.configureDefaults(loaded.defaults);
  const data = core.readSheet(S.readPasted('时间\tPREFIX/压力\n8\t2\n3\t20'), '粘贴数据', S);
  assert.equal(data.kind, 'category'); assert.equal(data.rows[0].x, 0);
  assert.equal(data.series[0].name, '改名_压力'); assert.equal(data.series[0].color, '#123456');
  assert.equal(data.series[0].range.id, 'small');
  const axis = core.axesFor(data.series)[0];
  assert.equal(axis.groupLabel, '自定义小范围'); assert.equal(axis.min, 0); assert.equal(axis.max, 22);
  assert.deepEqual(Array.from(axis.band), [0.15, 0.35]);
  assert.equal(core.colorAt(1), 'hsl(90, 50%, 40%)');
  const exported = JSON.parse(store.bundle([], 'selection').text);
  assert.deepEqual(exported.defaults, files.selection.defaults);
});

test('invalid default updates preserve the last complete default set', () => {
  const { core } = browserCore(), values = Object.fromEntries(Object.entries(documents()).map(([type, doc]) => [type, doc.defaults]));
  core.configureDefaults(values); const before = JSON.stringify(core.getDefaults());
  const mutations = [
    d => { delete d.axes.axisRegions.medium; },
    d => { d.axes.axisGroups[1].id = 'small'; },
    d => { d.axes.autoBounds.paddingPercent = -1; },
    d => { d.regex.headerRule.pattern = '('; },
    d => { d.selection.palette = []; },
    d => { d.selection.fallbackColor.lightness = 120; }
  ];
  for (const mutate of mutations) {
    const invalid = structuredClone(values); mutate(invalid);
    assert.throws(() => core.configureDefaults(invalid)); assert.equal(JSON.stringify(core.getDefaults()), before);
  }
});

test('configuration document upload is recognized as configuration and local import does not change defaults', async () => {
  const doc = documents().regex, before = C.getDefaults();
  doc.defaults.headerRule.replacement = 'changed_$1';
  doc.configs.push({ schema: 'curve-viewer', version: 2, type: 'regex', name: '本地导入', headerRule: doc.defaults.headerRule });
  const text = JSON.stringify(doc);
  await assert.rejects(async () => S.readDataFile(Buffer.from(text), 'regex.json'), error => error.code === 'DISPLAY_CONFIG');
  assert.equal(R.parseBundle(text)[0].name, '本地导入');
  assert.deepEqual(C.getDefaults(), before);
});
