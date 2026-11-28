const { test } = require('node:test');
const assert = require('node:assert/strict');
const C = require('../src/core.js');
const S = require('../src/sheet-reader.js');

test('editable header regex supports captures, flags, disabled rules and literal replacement text', () => {
  assert.equal(C.normalizeHeader('HPODU/设备/告警1 (ID 34000) [1]'), '告警1[1]');
  const rule = { enabled: true, pattern: '^.*?/([^/]+?)\\s*\\(.*?\\)\\s*\\[(.*?)\\]$', flags: 'i', replacement: '$2:$1' };
  assert.equal(C.normalizeHeader('设备/温度 (ID 2) [℃]', rule), '℃:温度');
  const global = C.createHeaderFormatter({ enabled: true, pattern: 'a', flags: 'gi', replacement: 'B' });
  assert.equal(global('AaA'), 'BBB'); assert.equal(global('AaA'), 'BBB');
  assert.equal(C.normalizeHeader('  普通标题  ', rule), '普通标题');
  assert.equal(C.normalizeHeader('设备/温度 (ID 2) [℃]', { ...rule, enabled: false }), '设备/温度 (ID 2) [℃]');
  assert.equal(C.normalizeHeader('温度', { ...rule, pattern: '^.*$', replacement: '' }), '温度');
  assert.equal(C.normalizeHeader('温度', { ...rule, pattern: '^.*$', replacement: '<b>$&</b>' }), '<b>温度</b>');
});

test('invalid regex and flags are rejected before any header changes', () => {
  for (const value of [null, {}, { ...C.DEFAULT_HEADER_RULE, pattern: '(' }, { ...C.DEFAULT_HEADER_RULE, flags: 'gg' }, { ...C.DEFAULT_HEADER_RULE, flags: 'x' }, { ...C.DEFAULT_HEADER_RULE, pattern: '' }, { ...C.DEFAULT_HEADER_RULE, replacement: 1 }]) assert.throws(() => C.validateHeaderRule(value));
});

test('renaming always uses original headers and preserves rows, duplicates, numeric data and source identities', () => {
  const wb = S.readPasted('时间\t设备/温度 (ID 1) [℃]\t设备/压力 (ID 2) [bar]\t设备/温度 (ID 1) [℃]\n2\t20\t1\t21\n0\t30\t2\t31');
  const original = C.readSheet(wb, '粘贴数据', S);
  const rule = { enabled: true, pattern: '^.*$', flags: '', replacement: '同名' };
  const renamed = C.renameHeaders(original, rule);
  assert.equal(renamed.rows, original.rows);
  assert.deepEqual(renamed.series.map(s => [s.name, s.occurrence]), [['同名',1],['同名',2],['同名',3]]);
  assert.deepEqual(renamed.series.map(s => [s.sourceName,s.sourceOccurrence,s.column,s.min,s.max]), original.series.map(s => [s.sourceName,s.sourceOccurrence,s.column,s.min,s.max]));
  const restored = C.renameHeaders(renamed, C.DEFAULT_HEADER_RULE);
  assert.deepEqual(restored, original);
  assert.equal(C.renameHeaders(renamed, rule).series[0].name, '同名');
  assert.deepEqual(C.readSheet(wb,'粘贴数据',S,'auto',rule), renamed);
});

test('each v2 profile owns only its fields and names may repeat in separate categories', () => {
  const base = { schema: 'curve-viewer', version: 2, name: '设备一' };
  const configs = [
    { ...base, type: 'axes', timeMode: 'number', axisLimits: { small: { min: -1, max: 5 } }, axisRegions: { small: { start: 50, end: 100 } } },
    { ...base, type: 'regex', headerRule: C.DEFAULT_HEADER_RULE },
    { ...base, type: 'selection', selectedSeries: [{ name:'压力',occurrence:1,sourceName:'原始压力',sourceOccurrence:1,color:'#123456' }], seriesAxes: [{ name:'压力',occurrence:1,sourceName:'原始压力',sourceOccurrence:1,axisId:'small' }] }
  ];
  for(const config of configs) assert.deepEqual(C.validateProfile(JSON.parse(JSON.stringify(config))),config);
  for(const [type,field,value] of [['axes','selectedSeries',[]],['axes','seriesAxes',[]],['selection','axisLimits',{}],['selection','headerRule',C.DEFAULT_HEADER_RULE],['regex','timeMode','auto']]) {
    assert.throws(()=>C.validateProfile({...configs.find(c=>c.type===type),[field]:value}),/混入/);
  }
  assert.throws(()=>C.validateProfile({...configs[2],selectedSeries:[{name:'压力',occurrence:1,sourceName:'原始压力',sourceOccurrence:0}]}),/原始列名/);
});

test('legacy combined configs migrate to independent axes and selection profiles without losing bounds or colors', () => {
  const legacy = { schema:'curve-viewer',version:1,name:'旧配置',timeMode:'excel',selectedSeries:[{name:'温度',occurrence:1,color:'#123456'}],axisLimits:{thermal:{min:-40,max:60}},axisRegions:{thermal:{start:20,end:70}},seriesAxes:[{name:'温度',occurrence:1,axisId:'thermal'}] };
  const [axes,selection] = C.splitConfig(legacy);
  assert.equal(axes.type,'axes'); assert.equal(axes.selectedSeries,undefined); assert.equal(axes.headerRule,undefined);
  assert.equal(axes.timeMode,'excel'); assert.deepEqual(axes.axisLimits,legacy.axisLimits); assert.deepEqual(axes.axisRegions,legacy.axisRegions); assert.equal(axes.seriesAxes,undefined); assert.deepEqual(selection.seriesAxes,legacy.seriesAxes);
  assert.equal(selection.type,'selection'); assert.deepEqual(selection.selectedSeries,legacy.selectedSeries); assert.equal(selection.axisLimits,undefined); assert.equal(selection.timeMode,undefined);
  assert.deepEqual(C.splitConfig(axes),[axes]); assert.deepEqual(C.splitConfig(selection),[selection]);
});

test('old axis-owned assignments migrate into matching legacy selections and preserve new explicit assignments', () => {
  const base = {schema:'curve-viewer',version:2,name:'设备'};
  const seriesAxes = [{name:'压力',occurrence:1,sourceName:'原始压力',sourceOccurrence:1,axisId:'auto'}];
  const oldAxes = {...base,type:'axes',timeMode:'auto',seriesAxes};
  const oldSelection = {...base,type:'selection',selectedSeries:[{name:'温度',occurrence:1,color:'#123456'}]};
  const migrated = C.migrateProfiles([oldAxes,oldSelection]);
  assert.equal(migrated.length,2);assert.equal(migrated[0].seriesAxes,undefined);
  assert.deepEqual(migrated[1].seriesAxes,seriesAxes);assert.deepEqual(migrated[1].selectedSeries,oldSelection.selectedSeries);
  const explicit = {...oldSelection,seriesAxes:[]};
  const protectedProfiles = C.migrateProfiles([oldAxes,explicit]);
  assert.deepEqual(protectedProfiles.find(c=>c.type==='selection'&&c.name==='设备').seriesAxes,[]);
  const backup = protectedProfiles.find(c=>c.name==='设备（旧坐标分配）');
  assert.deepEqual(backup.selectedSeries,[]);assert.deepEqual(backup.seriesAxes,seriesAxes);
  assert.deepEqual(C.migrateProfiles(protectedProfiles),protectedProfiles);
  assert.equal(C.migrateProfiles([oldAxes])[1].name,'设备（旧坐标分配）');
  assert.throws(()=>C.migrateProfiles([{...oldAxes,seriesAxes:[{...seriesAxes[0],axisId:'unknown'}]}]),/分配无效/);
});

test('the last legacy axis choice determines assignments when the last selection has a different name', () => {
  const base = {schema:'curve-viewer',version:2};
  const oldSelection = {...base,type:'selection',name:'曲线',selectedSeries:[{name:'温度',occurrence:1}]};
  const assigned = {...base,type:'axes',name:'曲线',seriesAxes:[{name:'温度',occurrence:1,axisId:'medium'}]};
  const automatic = {...base,type:'axes',name:'自动',seriesAxes:[]};
  const last = {axes:{source:'local',name:'自动'},selection:{source:'local',name:'曲线'}};
  const migrated=C.migrateProfiles([assigned,automatic,oldSelection],last);
  assert.deepEqual(migrated.find(c=>c.type==='selection').seriesAxes,[]);
  assert.deepEqual(migrated.find(c=>c.name==='曲线（旧坐标分配）').seriesAxes,assigned.seriesAxes);
});
