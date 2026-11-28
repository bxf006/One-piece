// Exercise the real app handlers with a small DOM/chart test double.
// This checks state transitions, not browser layout or native clipboard behavior.
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const C = require('../src/core.js');
const S = require('../src/sheet-reader.js');
const { makePMSD, FILENAME } = require('./pmsd-fixture.cjs');
const { X, excelBook, excelBytes } = require('./excel-fixture.cjs');
const R = require('../src/config-store.js');
const { IDBFactory } = require('fake-indexeddb');

async function setup(configs = [], lastConfigName = '', options = {}) {
  const runtime = { TextEncoder, TextDecoder };
  vm.createContext(runtime);
  for (const file of ['core.js', 'config-store.js']) vm.runInContext(fs.readFileSync(path.join(__dirname, '../src', file), 'utf8'), runtime, { filename: file });
  const appCore = runtime.CurveCore, appStore = runtime.CurveConfigStore;
  const documents = options.cloudDocuments || Object.fromEntries(C.CONFIG_TYPES.map(type => [type, JSON.parse(fs.readFileSync(path.join(__dirname, '../configs', type + '.json'), 'utf8'))]));
  const fetchCloud = async (url, init) => {
    const type = url.split('/').at(-1).replace('.json', '');
    if (!options.fetch) return { ok: true, status: 200, text: async () => JSON.stringify(documents[type]) };
    const response = await options.fetch(url, init);
    if (!response.ok || options.rawCloud) return response;
    const text = await response.text(); let data;
    try { data = JSON.parse(text); } catch { return { ...response, text: async () => text }; }
    return { ...response, text: async () => JSON.stringify(Array.isArray(data) ? { ...documents[type], configs: data } : data) };
  };
  let focused;
  const stored = new Map([['curve-explorer.configs.v1', JSON.stringify(configs)], ['curve-explorer.last-config.v1', lastConfigName]]);
  if (configs.some(c => c.version === 2) || typeof lastConfigName === 'object') {
    stored.set('curve-explorer.configs.v2', JSON.stringify(configs));
    stored.set('curve-explorer.last-config.v2', JSON.stringify(typeof lastConfigName === 'object' ? lastConfigName : { selection: lastConfigName }));
  }
  const downloads = [], downloadNames = [];
  const idb = Object.hasOwn(options, 'indexedDB') ? options.indexedDB : new IDBFactory();
  const repository = appStore.create(idb);
  let cached = {configs:[],last:{}};
  const store = { ...repository, snapshot: async () => { cached = await repository.snapshot(); return cached; }, select: async (...args) => { await repository.select(...args); cached = await repository.snapshot(); } }; 
  let deferFrames = false, nextFrame = 0, onResize;
  const frames = new Map();
  const errors = [];
  const windowEvents = {};
  class Element {
    constructor(tag = 'div') {
      this.tagName = tag; this.children = []; this.attributes = {}; this.style = {}; this.value = ''; this.hidden = false; this.events = {}; this.offsetWidth = 240; this.offsetHeight = 100;
      const classes = new Set();
      this.classList = { toggle: (name, yes) => yes ? classes.add(name) : classes.delete(name), contains: name => classes.has(name) };
    }
    set textContent(value) { this.text = String(value); this.children = []; }
    get textContent() { return this.text || this.children.map(child => child.textContent ?? child).join(''); }
    setAttribute(key, value) { this.attributes[key] = String(value); }
    getAttribute(key) { return this.attributes[key]; }
    append(...children) { this.children.push(...children.flatMap(child => child?.tagName === 'fragment' ? child.children : [child])); }
    replaceChildren(...children) { this.text = ''; this.children = []; this.append(...children); if (this.tagName === 'select') this.value = this.children.find(x => x.selected)?.value || this.children[0]?.value || ''; }
    addEventListener(name, handler) { this.events[name] = handler; }
    get options() { return this.children; }
    cloneNode() { return Object.assign(new Element(this.tagName), { value: this.value, text: this.text, selected: this.selected }); }
    focus() { focused = this; }
    showModal() { this.open = true; }
    close() { this.open = false; }
    click() { if(this.download) downloadNames.push(this.download); return this.onclick?.(); }
    remove() {}
    getBoundingClientRect() { return { left: 100, top: 50 }; }
  }
  const template = fs.readFileSync(path.join(__dirname, '../src/template.html'), 'utf8');
  const elements = new Map();
  for (const match of template.matchAll(/<([a-z][\w-]*)\b([^>]*\bid="([^"]+)"[^>]*)>/gi)) {
    const element = new Element(match[1]); element.hidden = /\bhidden\b/.test(match[2]); elements.set(match[3], element);
  }
  const get = id => { assert.ok(elements.has(id), `Missing template element: ${id}`); return elements.get(id); };
  get('time-mode').replaceChildren(...C.TIME_MODES.map(value => Object.assign(new Element('option'), { value, text: value })));
  const chart = {
    width: 800, height: 500, resizeCount: 0, handlers: {},
    clear() { this.option = null; }, setOption(option) { this.option = option; }, resize() { this.resizeCount++; },
    getZr: () => ({ on() {} }), on(name, callback) { this.handlers[name] = callback; },
    dispatchAction(action) { if (!this.option) return; Object.assign(this.option.dataZoom[0], { start: action.start, end: action.end }); this.handlers.datazoom?.(action); },
    getWidth() { return this.width; }, getHeight() { return this.height; },
    domain() { const [min, max] = this.option.domain, zoom = this.option.dataZoom[0]; return [min + (max - min) * zoom.start / 100, min + (max - min) * zoom.end / 100]; },
    convertToPixel(_, x) { const [lo, hi] = this.domain(), g = this.option.grid; return g.left + (x - lo) / (hi - lo) * (this.width - g.left - g.right); },
    convertFromPixel(_, x) { const [lo, hi] = this.domain(), g = this.option.grid; return lo + (x - g.left) / (this.width - g.left - g.right) * (hi - lo); },
    containPixel(_, [x, y]) { if (!this.option) return false; const g = this.option.grid; return x >= g.left && x <= this.width - g.right && y >= g.top && y <= this.height - g.bottom; }
  };
  const context = {
    CurveCore: appCore, SheetReader: S, JSZip: {}, CurvePlot: { init: () => chart },
    CurveConfigStore: { ...appStore, create: () => store }, indexedDB: idb, location: { protocol: options.protocol || 'https:' }, fetch: fetchCloud, AbortController,
    document: { body: new Element('body'), getElementById: id => elements.get(id) || null, createElement: tag => new Element(tag), createDocumentFragment: () => new Element('fragment'), createTextNode: text => ({ textContent: String(text) }) },
    console: { error: (...args) => errors.push(args) },
    localStorage: { getItem: key => stored.get(key) ?? null, setItem(key, value) { stored.set(key, value); } },
    Blob, URL: { createObjectURL(blob) { downloads.push(blob); return 'blob:test'; }, revokeObjectURL() {} },
    ResizeObserver: class { constructor(callback) { onResize = callback; } observe() {} },
    requestAnimationFrame(callback) { const id = ++nextFrame; if (deferFrames) frames.set(id, callback); else callback(); return id; },
    setTimeout: () => 1, clearTimeout() {}, addEventListener: (name, callback) => { windowEvents[name] = callback; }
  };
  context.window = context;
  vm.runInNewContext(fs.readFileSync(path.join(__dirname, '../src/app.js'), 'utf8'), context, { filename: 'app.js' });
  await context.CurveAppReady;
  const enter = text => { get('paste-input').value = text; get('paste-input').oninput(); };
  return { get, chart, enter, errors, windowEvents, downloads, downloadNames, repository, store, idb, core: appCore, savedConfigs: () => structuredClone(cached.configs), lastConfig: () => cached.last.selection?.name || '', lastConfigs: () => Object.fromEntries(C.CONFIG_TYPES.map(type=>[type,cached.last[type]?.name || ''])), choices: () => structuredClone(cached.last), removeElement: id => elements.delete(id), focused: () => focused,
    resize: () => onResize(), deferFrames: value => { deferFrames = value; }, flushFrames() { const callbacks = [...frames.values()]; frames.clear(); callbacks.forEach(callback => callback()); } };
}

async function saveCategory(app, type, name) {
  app.get('config-tab-' + type).onclick(); app.get('config-name').value = name; await app.get('save-config').onclick();
  return app.savedConfigs().find(c => c.type === type && c.name === name);
}

function editHeaderRule(get, rule, apply = true) {
  get('config-tab-regex').onclick(); get('regex-settings').onclick();
  for (const [key, value] of Object.entries(rule)) {
    if (key === 'enabled') get('regex-enabled').checked = value;
    else get('regex-' + key).value = value;
  }
  get('regex-pattern').oninput();
  if (apply) get('regex-form').onsubmit({ preventDefault() {} });
}

test('saving updates IndexedDB without downloads and exports all local configs of one type in one JSON array', async () => {
  const app=await setup(),{get,enter}=app;
  enter('时间\t温度\t压力\n0\t20\t2');await get('paste-load').onclick();get('series-search').value='温度';get('select-all').onclick();
  await saveCategory(app,'selection','温度配置');get('select-none').onclick();get('series-search').value='压力';get('select-all').onclick();
  await saveCategory(app,'selection','压力配置');await saveCategory(app,'selection','温度配置');
  assert.equal(app.downloads.length,0);assert.equal(app.savedConfigs().length,2);assert.equal(app.savedConfigs()[0].selectedSeries[0].name,'压力');
  await get('export-configs').onclick();assert.equal(app.downloadNames[0],'selection.json');assert.deepEqual(JSON.parse(await app.downloads[0].text()).configs,app.savedConfigs());
  const restored=await setup([], '', {indexedDB:app.idb});assert.deepEqual(restored.savedConfigs(),app.savedConfigs());
  await restored.get('config-list').children[0].children[1].events.click();
  const again=await setup([], '', {indexedDB:app.idb});assert.deepEqual(again.savedConfigs().map(c=>c.name),['压力配置']);
});

test('cloud profiles load read-only, keep separate names from local profiles and persist the selected source', async () => {
  const local={schema:'curve-viewer',version:2,type:'selection',name:'同名',selectedSeries:[{name:'温度',occurrence:1}]};
  const cloud={...local,selectedSeries:[{name:'压力',occurrence:1}]};
  const fetch=async url=>({ok:true,status:200,text:async()=>JSON.stringify(url.endsWith('selection.json')?[cloud]:[])});
  const app=await setup([local],'同名',{protocol:'http:',fetch}),{get,chart,enter}=app;
  assert.equal(get('config-list').children[0].children.length,2);assert.equal(get('cloud-config-list').children[0].children.length,1);
  enter('时间\t温度\t压力\n0\t20\t2');await get('paste-load').onclick();assert.equal(chart.option.series[0].name,'温度');
  await get('cloud-config-list').children[0].children[0].events.click();assert.equal(chart.option.series[0].name,'压力');
  assert.deepEqual(app.choices().selection,{source:'cloud',name:'同名'});assert.equal(app.savedConfigs()[0].selectedSeries[0].name,'温度');
  const restored=await setup([], '', {indexedDB:app.idb,protocol:'http:',fetch});restored.enter('时间\t温度\t压力\n0\t20\t2');await restored.get('paste-load').onclick();assert.equal(restored.chart.option.series[0].name,'压力');
  await get('config-list').children[0].children[1].events.click();assert.equal(app.savedConfigs().length,0);assert.deepEqual(app.choices().selection,{source:'cloud',name:'同名'});
  await get('save-config').onclick();assert.equal(app.savedConfigs().length,1);assert.equal(app.savedConfigs()[0].selectedSeries[0].name,'压力');assert.equal(get('cloud-config-list').children[0].children.length,1);assert.equal(app.choices().selection.source,'local');
});

test('startup and refresh read the fixed cloud directory and invalid refresh retains the previous cloud list and plot', async () => {
  const profile={schema:'curve-viewer',version:2,type:'selection',name:'云端',selectedSeries:[{name:'温度',occurrence:1}]};
  let cloudText=JSON.stringify([profile]);const calls=[];
  const fetch=async(url,options)=>{calls.push(url);assert.equal(options.cache,'no-store');return{ok:true,status:200,text:async()=>url.endsWith('selection.json')?cloudText:'[]'};};
  const app=await setup([], '', {protocol:'http:',fetch}),{get,enter,chart}=app;
  assert.deepEqual(calls,['./configs/axes.json','./configs/regex.json','./configs/selection.json']);
  assert.equal(app.savedConfigs().length,0);assert.equal(get('cloud-config-list').children[0].children.length,1);
  enter('时间\t温度\t压力\n0\t20\t2');await get('paste-load').onclick();assert.equal(chart.option.series.length,0);
  await get('cloud-config-list').children[0].children[0].events.click();const before=chart.option;
  cloudText='{bad';await get('refresh-cloud').onclick();
  assert.equal(chart.option,before);assert.match(get('cloud-config-status').textContent,/失败.*保留/);assert.match(get('cloud-config-list').textContent,/云端/);
  const next={...profile,selectedSeries:[{name:'压力',occurrence:1}]};
  cloudText=JSON.stringify([next]);await get('refresh-cloud').onclick();
  assert.equal(calls.length,9);assert.ok(calls.every(url=>url.startsWith('./configs/')));
  assert.equal(chart.option,before);await get('cloud-config-list').children[0].children[0].events.click();assert.equal(chart.option.series[0].name,'压力');
  await get('config-input').onchange({target:{files:[dataFile('selection.json',JSON.stringify([profile,next].map((c,i)=>({...c,name:'本地'+i}))))]}});
  assert.equal(app.savedConfigs().length,2);await get('export-configs').onclick();assert.deepEqual(JSON.parse(await app.downloads[0].text()).configs.map(c=>c.name),['本地0','本地1']);
});

test('static website startup failure preserves local curves and retry plus page reload restore the selected cloud config', async () => {
  const local={schema:'curve-viewer',version:2,type:'selection',name:'本地',selectedSeries:[{name:'温度',occurrence:1}]};
  const cloud={...local,name:'网站配置',selectedSeries:[{name:'压力',occurrence:1}]};
  let fail=true;const requests=[];
  const fetch=async url=>{requests.push(url);if(fail)throw new Error('network unavailable');return{ok:true,status:200,text:async()=>JSON.stringify(url.endsWith('selection.json')?[cloud]:[])};};
  const app=await setup([local],'本地',{fetch}),{get,enter,chart}=app;
  assert.match(get('cloud-config-status').textContent,/读取失败/);assert.equal(app.savedConfigs().length,1);
  enter('时间\t温度\t压力\n0\t20\t2');await get('paste-load').onclick();assert.equal(chart.option,null);assert.match(get('paste-status').textContent,/默认配置尚未载入/);
  fail=false;await get('refresh-cloud').onclick();assert.match(get('cloud-config-list').textContent,/网站配置/);
  await get('paste-load').onclick();assert.equal(chart.option.series[0].name,'温度');
  await get('cloud-config-list').children[0].children[0].events.click();assert.equal(chart.option.series[0].name,'压力');
  const reopened=await setup([], '', {indexedDB:app.idb,fetch});
  assert.deepEqual(reopened.choices().selection,{source:'cloud',name:'网站配置'});
  reopened.enter('时间\t温度\t压力\n0\t25\t3');await reopened.get('paste-load').onclick();assert.equal(reopened.chart.option.series[0].name,'压力');
  assert.equal(requests.length,9);assert.ok(requests.every(url=>url.startsWith('./configs/')));
});

test('JSON defaults control initial selection, axes and reset actions and export with local profiles', async () => {
  const documents=Object.fromEntries(C.CONFIG_TYPES.map(type=>[type,JSON.parse(fs.readFileSync(path.join(__dirname,'../configs',type+'.json'),'utf8'))]));
  documents.axes.defaults.axisRegions.medium={start:5,end:55};documents.axes.defaults.axisLimits.medium={min:-10,max:60};
  documents.selection.defaults.seriesAxes=[{name:'温度',occurrence:1,axisId:'medium'}];
  documents.regex.defaults.headerRule={enabled:true,pattern:'^(.+)$',flags:'',replacement:'R_$1'};
  documents.selection.defaults.selectedSeries=[{name:'温度',occurrence:1}];documents.selection.defaults.palette=['#123456'];
  const app=await setup([], '', {cloudDocuments:documents}),{get,enter,chart}=app;
  enter('时间\t温度\t压力\n0\t20\t2\n1\t30\t3');await get('paste-load').onclick();
  assert.equal(chart.option.series.length,1);assert.equal(chart.option.series[0].name,'R_温度');assert.equal(chart.option.series[0].lineStyle.color,'#123456');
  assert.equal(chart.option.yAxis[0].min,-10);assert.equal(chart.option.yAxis[0].max,60);assert.deepEqual(Array.from(chart.option.yAxis[0].band),[0.05,0.55]);
  editHeaderRule(get,{enabled:true,pattern:'^(.+)$',flags:'',replacement:'X_$1'});
  get('regex-settings').onclick();get('regex-reset').onclick();assert.equal(get('regex-replacement').value,'R_$1');get('regex-form').onsubmit({preventDefault(){}});
  get('config-tab-axes').onclick();get('axis-region-settings').onclick();regionInput(get,'medium','start',10);regionInput(get,'medium','end',30);applyRegions(get);
  get('axis-region-settings').onclick();get('axis-region-reset-all').onclick();applyRegions(get);assert.deepEqual(Array.from(chart.option.yAxis[0].band),[0.05,0.55]);
  await saveCategory(app,'selection','保存默认曲线');await get('export-configs').onclick();
  const exported=JSON.parse(await app.downloads.at(-1).text());assert.deepEqual(exported.defaults,documents.selection.defaults);assert.equal(exported.configs[0].name,'保存默认曲线');
  const before=chart.option;await get('data-input').onchange({target:{files:[dataFile('selection.json',JSON.stringify(exported))]}});assert.equal(chart.option,before);assert.equal(app.savedConfigs().length,1);
});

test('last chosen profiles override JSON defaults and refreshed JSON drives subsequent resets', async () => {
  const documents=Object.fromEntries(C.CONFIG_TYPES.map(type=>[type,JSON.parse(fs.readFileSync(path.join(__dirname,'../configs',type+'.json'),'utf8'))]));
  documents.selection.defaults.selectedSeries=[{name:'温度',occurrence:1}];
  const local={schema:'curve-viewer',version:2,type:'selection',name:'压力选择',selectedSeries:[{name:'压力',occurrence:1}]};
  const app=await setup([local],'压力选择',{cloudDocuments:documents}),{get,enter,chart}=app;
  enter('时间\t温度\t压力\n0\t20\t2');await get('paste-load').onclick();assert.equal(chart.option.series[0].name,'压力');const before=chart.option;
  documents.regex.defaults.headerRule={enabled:true,pattern:'^(.+)$',flags:'',replacement:'NEW_$1'};
  documents.axes.defaults.axisRegions.small={start:40,end:70};
  await get('refresh-cloud').onclick();assert.equal(chart.option,before);
  get('config-tab-regex').onclick();get('regex-settings').onclick();get('regex-reset').onclick();assert.equal(get('regex-replacement').value,'NEW_$1');get('regex-form').onsubmit({preventDefault(){}});
  get('config-tab-axes').onclick();get('axis-region-settings').onclick();get('axis-region-reset-all').onclick();applyRegions(get);assert.deepEqual(Array.from(chart.option.yAxis[0].band),[0.4,0.7]);
  assert.equal(chart.option.series[0].name,'NEW_压力');
});

test('file URL explains static deployment without making blocked configuration requests', async () => {
  let requests=0;const app=await setup([], '', {protocol:'file:',fetch:async()=>{requests++;throw new TypeError('Failed to fetch');}});
  assert.equal(requests,0);assert.equal(app.core.defaultsReady,false);
  assert.match(app.get('cloud-config-status').textContent,/file:\/\/.*静态网站.*HTTP\/HTTPS/);
  assert.doesNotMatch(app.get('cloud-config-status').textContent,/start-viewer|cmd|Failed to fetch|已保留/);
  await app.get('refresh-cloud').onclick();assert.equal(requests,0);assert.equal(app.get('save-config').disabled,true);
});

test('failed local saves and deletes preserve lists and never claim success', async () => {
  const app=await setup(),{get,enter}=app;enter('时间\t温度\n0\t20');await get('paste-load').onclick();get('select-all').onclick();
  const put=app.store.putMany;app.store.putMany=async()=>{throw new Error('QuotaExceededError');};
  get('config-name').value='失败项';await get('save-config').onclick();assert.equal(app.savedConfigs().length,0);assert.match(get('toast').textContent,/保存失败/);assert.equal(app.downloads.length,0);
  app.store.putMany=put;await get('save-config').onclick();assert.equal(app.savedConfigs().length,1);
  const remove=app.store.remove;app.store.remove=async()=>{throw new Error('transaction aborted');};
  await get('config-list').children[0].children[1].events.click();assert.equal(app.savedConfigs().length,1);assert.match(get('toast').textContent,/删除失败/);
  app.store.remove=remove;await get('config-list').children[0].children[1].events.click();assert.equal(app.savedConfigs().length,0);
});

test('data and cloud viewing still work when IndexedDB is unavailable while local saving remains disabled', async () => {
  const cloud={schema:'curve-viewer',version:2,type:'selection',name:'云端',selectedSeries:[{name:'温度',occurrence:1}]};
  const app=await setup([], '', {indexedDB:undefined,protocol:'https:',fetch:async url=>({ok:true,status:200,text:async()=>JSON.stringify(url.endsWith('selection.json')?[cloud]:[])})});
  const {get,enter,chart}=app;assert.match(get('config-storage-status').textContent,/IndexedDB 不可用/);assert.equal(get('save-config').disabled,true);
  enter('时间\t温度\n0\t20');await get('paste-load').onclick();await get('cloud-config-list').children[0].children[0].events.click();assert.equal(chart.option.series[0].name,'温度');assert.match(get('toast').textContent,/未能保存/);
  await get('save-config').onclick();assert.equal(app.savedConfigs().length,0);assert.match(get('toast').textContent,/未保存/);
});

test('returning to a page refreshes local IndexedDB records changed by another page', async () => {
  const first=await setup(),second=await setup([], '', {indexedDB:first.idb});
  editHeaderRule(first.get,{enabled:true,pattern:'^(.+)$',flags:'',replacement:'R_$1'});await saveCategory(first,'regex','跨页面');
  await second.windowEvents.focus();assert.equal(second.savedConfigs().length,1);second.get('config-tab-regex').onclick();assert.match(second.get('config-list').textContent,/跨页面/);
  await first.get('config-list').children[0].children[1].events.click();await second.windowEvents.focus();assert.equal(second.savedConfigs().length,0);
});

test('three profile categories export separately, keep equal names and restore independent last choices', async () => {
  const app=await setup(),{get,chart,enter}=app;
  enter('时间\t温度\t压力\n2\t20\t2\n0\t30\t3'); await get('paste-load').onclick();
  assert.equal(get('save-config').disabled,false);
  get('series-search').value='温度';get('select-all').onclick();
  get('axis-settings').onclick();axisInput(get,'左1','下限',-40);axisInput(get,'左1','上限',80);applyAxes(get);
  get('axis-region-settings').onclick();regionInput(get,'thermal','start',10);regionInput(get,'thermal','end',60);applyRegions(get);
  const selection=await saveCategory(app,'selection','同名'),axes=await saveCategory(app,'axes','同名');
  editHeaderRule(get,{enabled:true,pattern:'^(.+)$',flags:'',replacement:'测点_$1'});
  const regex=await saveCategory(app,'regex','同名');
  assert.equal(selection.axisLimits,undefined);assert.equal(selection.headerRule,undefined);assert.equal(axes.selectedSeries,undefined);assert.equal(regex.selectedSeries,undefined);assert.equal(regex.axisRegions,undefined);
  assert.deepEqual(app.savedConfigs().map(c=>c.type),['selection','axes','regex']);
  assert.deepEqual(app.lastConfigs(),{axes:'同名',regex:'同名',selection:'同名'});
  assert.equal(app.downloads.length,0); // Saving to IndexedDB no longer downloads a JSON file.
  for(const config of [selection,axes,regex]){get('config-tab-'+config.type).onclick();await get('export-configs').onclick();assert.deepEqual(JSON.parse(await app.downloads.at(-1).text()).configs,[config]);}
  assert.deepEqual(Array.from(chart.option.series,s=>s.name),['测点_温度']);
  assert.equal(chart.option.yAxis[0].min,-40);assert.deepEqual(Array.from(chart.option.yAxis[0].band),[0.1,0.6]);
  const restored=await setup(app.savedConfigs(),app.lastConfigs());
  restored.enter('时间\t温度\t压力\n3\t40\t4\n1\t50\t5');await restored.get('paste-load').onclick();
  assert.deepEqual(Array.from(restored.chart.option.series,s=>s.name),['测点_温度']);assert.equal(restored.chart.option.yAxis[0].max,80);
  assert.deepEqual(Array.from(restored.chart.option.series[0].data,row=>row[0]),[3,1]);
  restored.get('clear-data').onclick();restored.enter('时间\t温度\n0\t20');await restored.get('paste-load').onclick();
  assert.equal(restored.chart.option.series[0].name,'测点_温度');assert.equal(restored.chart.option.yAxis[0].min,-40);
});

test('editing regex validates and previews safely, cancel preserves data and reset starts from raw headers', async () => {
  const app=await setup(),{get,chart,enter}=app, raw='HPODU/设备/温度 (ID 7) [℃]';
  enter(`时间\t${raw}\n1\t20\n0\t30`);await get('paste-load').onclick();get('select-all').onclick();
  const before=chart.option;
  editHeaderRule(get,{pattern:'('},false);assert.equal(get('regex-apply').disabled,true);assert.match(get('regex-error').textContent,/无效/);
  get('regex-form').onsubmit({preventDefault(){}});assert.equal(chart.option,before);assert.equal(get('regex-dialog').open,true);
  get('regex-cancel').onclick();get('regex-settings').onclick();assert.equal(get('regex-pattern').value,C.DEFAULT_HEADER_RULE.pattern);
  editHeaderRule(get,{pattern:'^(.+)$',flags:'',replacement:'<b>$1</b>'},false);
  assert.match(get('regex-preview').textContent,/<b>HPODU/);assert.equal(get('regex-preview').children.length,0);
  get('regex-cancel').onclick();assert.equal(chart.option,before);
  editHeaderRule(get,{enabled:false});assert.equal(chart.option.series[0].name,raw);
  editHeaderRule(get,{enabled:true,pattern:'^(.+)$',replacement:'新_$1'});assert.equal(chart.option.series[0].name,'新_'+raw);
  get('regex-settings').onclick();get('regex-reset').onclick();get('regex-form').onsubmit({preventDefault(){}});
  assert.equal(chart.option.series[0].name,'温度[℃]');assert.deepEqual(Array.from(chart.option.series[0].data,p=>Array.from(p)),[[1,20],[0,30]]);
});

test('regex name collisions keep exact source columns, selections and manual axes through save and reload', async () => {
  const app=await setup(),{get,chart,enter}=app;
  const a='HPODU/设备一/告警 (ID 1) [1]',b='HPODU/设备二/告警 (ID 2) [1]';
  const pasted=`时间\t${a}\t${b}\t${a}\n2\t1\t2\t3\n0\t4\t5\t6`;
  enter(pasted);await get('paste-load').onclick();
  for(const index of [1,2]){const checkbox=get('series-list').children[index].children[0];checkbox.checked=true;checkbox.events.change();}
  get('series-axis-settings').onclick();curveAxis(get,'告警[1] (2)','medium');curveAxis(get,'告警[1] (3)','auto');applyCurveAxes(get);
  const selected=await saveCategory(app,'selection','两列'),axes=await saveCategory(app,'axes','两轴');
  editHeaderRule(get,{enabled:true,pattern:'^.*$',flags:'',replacement:'同名'});await saveCategory(app,'regex','合并名称');
  assert.deepEqual(Array.from(chart.option.series,s=>s.data[0][1]),[2,3]);
  assert.deepEqual(Array.from(chart.option.series,s=>s.name),['同名 (2)','同名 (3)']);
  assert.deepEqual(Array.from(chart.option.yAxis,a=>Array.from(a.band)),[[0,0.25],[0,1]]);
  const restored=await setup(app.savedConfigs(),app.lastConfigs());restored.enter(pasted);await restored.get('paste-load').onclick();
  assert.deepEqual(Array.from(restored.chart.option.series,s=>s.data[0][1]),[2,3]);
  restored.get('config-tab-selection').onclick();restored.get('select-all').onclick();await restored.get('config-list').children[0].children[0].events.click();
  assert.deepEqual(Array.from(restored.chart.option.series,s=>s.data[0][1]),[2,3]);
  assert.equal(selected.selectedSeries[0].sourceName,b);assert.equal(selected.selectedSeries[1].sourceOccurrence,2);assert.equal(selected.seriesAxes[0].sourceName,b);assert.equal(axes.seriesAxes,undefined);
});

test('legacy migration and category changes do not overwrite the other selected profile or its settings', async () => {
  const legacy={schema:'curve-viewer',version:1,name:'旧温度',timeMode:'auto',selectedSeries:[{name:'温度',occurrence:1}],axisLimits:{small:{min:0,max:9}}};
  const other={...legacy,name:'旧压力',selectedSeries:[{name:'压力',occurrence:1}],axisLimits:{small:{min:-1,max:8}}};
  const app=await setup([legacy,other],'旧温度'),{get,chart,enter}=app;
  assert.equal(app.savedConfigs().length,4);assert.deepEqual(app.lastConfigs(),{axes:'旧温度',regex:'',selection:'旧温度'});
  enter('时间\t温度\t压力\n0\t20\t2');await get('paste-load').onclick();
  await get('config-list').children[1].children[0].events.click();assert.equal(chart.option.series[0].name,'压力');assert.equal(chart.option.yAxis[0].max,9);
  get('config-tab-axes').onclick();await get('config-list').children[1].children[0].events.click();assert.equal(chart.option.series[0].name,'压力');assert.equal(chart.option.yAxis[0].max,8);
  assert.deepEqual(app.lastConfigs(),{axes:'旧压力',regex:'',selection:'旧压力'});
  get('config-tab-selection').onclick();await get('config-list').children[0].children[0].events.click();
  assert.deepEqual(app.lastConfigs(),{axes:'旧压力',regex:'',selection:'旧温度'});
});

test('regex and axes can be configured before loading and imported mixed profiles are atomic per file', async () => {
  const app=await setup(),{get,chart}=app;
  editHeaderRule(get,{enabled:true,pattern:'^(.+)$',flags:'',replacement:'R_$1'});
  const regex=await saveCategory(app,'regex','前缀');assert.equal(chart.option,null);
  get('config-tab-axes').onclick();get('axis-region-settings').onclick();regionInput(get,'thermal','start',5);applyRegions(get);
  const axes=await saveCategory(app,'axes','区域');assert.equal(axes.axisRegions.thermal.start,5);
  const other=await setup();await other.get('data-input').onchange({target:{files:[dataFile('regex.json',JSON.stringify(regex))]}});
  assert.equal(other.savedConfigs()[0].type,'regex');assert.equal(other.lastConfigs().regex,'');
  other.get('config-tab-regex').onclick();await other.get('config-list').children[0].children[0].events.click();
  const before=other.savedConfigs();
  await other.get('config-input').onchange({target:{files:[dataFile('bad.json',JSON.stringify([axes,{...regex,headerRule:{...regex.headerRule,pattern:'('}}]))]}});
  assert.deepEqual(other.savedConfigs(),before);assert.match(other.get('toast').textContent,/失败/);
  other.enter('时间\t温度\n0\t20');await other.get('paste-preview-button').onclick();
  assert.equal(other.get('paste-preview-table').children[0].children[0].children[1].textContent,'R_温度');
  await other.get('paste-load').onclick();assert.equal(other.chart.option.series.length,0);other.get('select-all').onclick();assert.equal(other.chart.option.series[0].name,'R_温度');
});

test('the selected header rule applies to CSV JSON TXT Excel and PMSD uploads', async () => {
  const app=await setup(),{get,chart}=app;
  editHeaderRule(get,{enabled:true,pattern:'^(.+)$',flags:'',replacement:'R_$1'});await saveCategory(app,'regex','前缀');
  const files=[dataFile('a.csv','时间,温度\n0,20'),dataFile('a.txt','时间\t温度\n0\t20'),dataFile('a.json',JSON.stringify([{时间:0,温度:20}])),binaryFile('a.xlsb',await excelBytes('xlsb')),binaryFile(FILENAME,makePMSD())];
  for(const file of files){await get('data-input').onchange({target:{files:[file]}});get('select-all').onclick();assert.ok(chart.option.series.length>0,file.name);assert.ok(chart.option.series.every(s=>s.name.startsWith('R_')),file.name);}
});

test('modified state and removal stay in their own category', async () => {
  const app=await setup(),{get,enter}=app;
  enter('时间\t温度\n0\t20');await get('paste-load').onclick();get('select-all').onclick();
  await saveCategory(app,'selection','选择');await saveCategory(app,'axes','坐标');await saveCategory(app,'regex','标题');
  get('select-none').onclick();assert.doesNotMatch(get('config-active').textContent,/已修改/);
  get('config-tab-selection').onclick();assert.match(get('config-active').textContent,/已修改/);
  get('config-tab-axes').onclick();assert.doesNotMatch(get('config-active').textContent,/已修改/);
  get('axis-region-settings').onclick();regionInput(get,'thermal','start',10);applyRegions(get);assert.match(get('config-active').textContent,/已修改/);
  get('config-tab-regex').onclick();await get('config-list').children[0].children[1].events.click();
  assert.deepEqual(app.lastConfigs(),{axes:'坐标',regex:'',selection:'选择'});assert.equal(app.savedConfigs().length,2);
});

test('tab navigation preserves the draft and chart, with keyboard focus and resize on return', async () => {
  const { get, chart, enter, focused } = await setup();
  const original = chart.option;
  get('tab-paste').onclick();
  enter('时间\t温度\n0\t20');
  assert.equal(get('paste-panel').hidden, false);
  assert.equal(get('view-panel').hidden, true);
  get('tab-paste').onkeydown({ key: 'Home', preventDefault() {} });
  assert.equal(get('view-panel').hidden, false);
  assert.equal(get('tab-view').getAttribute('aria-selected'), 'true');
  assert.equal(focused(), get('tab-view'));
  assert.ok(chart.resizeCount > 0);
  assert.equal(get('paste-input').value, '时间\t温度\n0\t20');
  assert.equal(chart.option, original);
});

test('preview is bounded and renders cell text without changing the current chart', async () => {
  const { get, chart, enter } = await setup();
  const original = chart.option;
  get('tab-paste').onclick();
  const header = ['时间', '<img src=x onerror=alert(1)>', ...Array(8).fill('温度')].join('\t');
  enter(header + '\n' + Array.from({ length: 12 }, (_, i) => [i, ...Array(9).fill(i + 1)].join('\t')).join('\n'));
  await get('paste-preview-button').onclick();
  const [head, body] = get('paste-preview-table').children;
  assert.equal(head.children[0].children.length, 8);
  assert.equal(head.children[0].children[1].textContent, '<img src=x onerror=alert(1)>');
  assert.equal(body.children.length, 8);
  assert.equal(get('paste-preview').hidden, false);
  assert.equal(chart.option, original);
  assert.equal(get('loading').hidden, true);
});

test('pasted data replaces the chart, clears the series filter and synchronizes time format', async () => {
  const { get, chart, enter } = await setup();
  get('tab-paste').onclick();
  get('series-search').value = '不存在';
  enter('时间\t温度\n46023.5\t20\n46024\t30');
  get('paste-time-mode').value = 'excel';
  await get('paste-load').onclick();
  get('select-all').onclick();
  assert.equal(chart.option.series[0].data.length, 2);
  assert.match(get('chart-subtitle').textContent, /粘贴数据/);
  assert.equal(get('time-mode').value, 'excel');
  assert.equal(get('series-search').value, '');
  assert.equal(chart.option.series[0].data[0][0], Date.UTC(2026, 0, 1, 12));
  assert.equal(get('view-panel').hidden, false);
  assert.equal(get('paste-input').getAttribute('aria-invalid'), 'false');
  assert.ok(get('paste-input').value.length > 0);
});

test('invalid paste preserves the chart and time mode; clearing allows another import', async () => {
  const { get, chart, enter } = await setup();
  const original = chart.option;
  get('tab-paste').onclick();
  enter('时间\t值\n####\t20'); get('paste-time-mode').value = 'date';
  await get('paste-load').onclick();
  assert.equal(chart.option, original);
  assert.equal(get('time-mode').value, 'auto');
  assert.equal(get('paste-panel').hidden, false);
  assert.equal(get('paste-input').getAttribute('aria-invalid'), 'true');
  assert.equal(get('loading').hidden, true);
  get('paste-clear').onclick();
  assert.equal(get('paste-input').value, '');
  assert.equal(get('paste-preview').hidden, true);
  assert.equal(get('paste-input').getAttribute('aria-invalid'), 'false');
  enter('时间\t压力\t温度\t电流\t转速\n0\t1\t20\t200\t2000\n1\t2\t21\t300\t3000');
  get('paste-time-mode').value = 'auto';
  await get('paste-load').onclick();
  get('select-all').onclick();
  assert.equal(chart.option.series[0].data.length, 2);
  assert.equal(chart.option.yAxis.length, 4);
});

test('saved series configuration remains usable after pasting new data', async () => {
  const config = { schema: 'curve-viewer', version: 1, name: '压力', timeMode: 'auto', selectedSeries: [{ name: '入口压力', occurrence: 1, color: '#123456' }] };
  const { get, chart, enter } = await setup([config]);
  enter('时间\t入口压力\t温度\n0\t1\t10');
  await get('paste-load').onclick();
  await get('config-list').children[0].children[0].events.click();
  get('tab-paste').onclick();
  enter('时间\t入口压力\t温度\n0\t3\t20\n1\t4\t30');
  await get('paste-load').onclick();
  assert.equal(chart.option.series.length, 1);
  assert.equal(chart.option.series[0].name, '入口压力');
  assert.equal(chart.option.series[0].lineStyle.color, '#123456');
});

test('startup has no data and empty-state controls do not create curves or crash', async () => {
  const { get, chart, errors } = await setup();
  assert.equal(chart.option, null);
  assert.equal(get('selection-count').textContent, '0 / 0');
  assert.equal(get('chart-empty').hidden, false);
  assert.match(get('chart-empty').textContent, /请上传/);
  assert.equal(get('paste-input').value, '');
  for (const id of ['series-search', 'select-all', 'select-none', 'save-config', 'sheet-select', 'time-mode', 'reset-zoom', 'axis-settings', 'series-axis-settings', 'clear-data']) assert.equal(get(id).disabled, true);
  get('select-all').onclick(); get('select-none').onclick(); get('sheet-select').onchange(); get('time-mode').onchange(); get('reset-zoom').onclick(); await get('save-config').onclick();
  assert.equal(chart.option, null);
  assert.match(get('toast').textContent, /请先载入数据/);
  assert.equal(errors.length, 0);
});

test('missing preview elements report the element and preserve an error stack for diagnosis', async () => {
  const { get, chart, enter, errors, removeElement } = await setup();
  const original = chart.option;
  get('tab-paste').onclick();
  enter('时间\t温度\n0\t20');
  removeElement('paste-preview-summary');
  await get('paste-load').onclick();
  assert.match(get('paste-status').textContent, /生成预览失败.*#paste-preview-summary/);
  assert.equal(chart.option, original);
  assert.equal(get('loading').hidden, true);
  assert.equal(errors.length, 1);
  assert.match(errors[0][1].stack, /renderPastePreview/);
});

test('wide telemetry paste with alarm columns, trailing header spaces and millisecond timestamps loads', async () => {
  const { get, chart, enter, errors } = await setup();
  const names = ['时间', ...Array.from({ length: 16 }, (_, i) => `告警${i + 1} `), ...Array.from({ length: 29 }, (_, i) => `测量${i + 1} `)];
  const values = [...Array(16).fill(0), 5, 100, 35, 1.4, -18.9, 1, 43, 43, -18, -2.6, 3.5, 14.4, -17.9, -14.7, 0, 0, 1.6, 1.6, -18, -17.7, 15.4, 21.2, 0, 0, 226, 0, 320, 0, -14];
  enter([names.join('\t'), ['2026-11-25 18:00:40.964', ...values].join('\t'), ['2026-11-25 18:00:41.601', ...values].join('\t')].join('\n'));
  get('tab-paste').onclick();
  await get('paste-load').onclick();
  assert.equal(errors.length, 0);
  assert.equal(get('view-panel').hidden, false);
  get('select-all').onclick();
  assert.equal(chart.option.series[0].data.length, 2);
  assert.equal(chart.option.series.length, 45);
  assert.equal(chart.option.series[44].data[0][1], -14);
  assert.equal(chart.option.series[0].data[1][0] - chart.option.series[0].data[0][0], 637);
});

const dataFile = (name, text) => ({ name, size: Buffer.byteLength(text), arrayBuffer: async () => new TextEncoder().encode(text).buffer, text: async () => text });
const binaryFile = (name, bytes) => ({ name, size: bytes.length, arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) });

test('Excel upload skips a non-data sheet, restores the last config and permits worksheet switching', async () => {
  const config = { schema: 'curve-viewer', version: 1, name: '只看温度', timeMode: 'auto', selectedSeries: [{ name: '温度[℃]', occurrence: 1 }] };
  const { get, chart, errors } = await setup([config], config.name), book = excelBook();
  X.utils.book_append_sheet(book, X.utils.aoa_to_sheet([['时间', '温度[℃]', '压力'], [2, 80, 3], [0, 20, 1]]), '另一组');
  await get('data-input').onchange({ target: { files: [binaryFile('数据.xlsx', await excelBytes('xlsx', book))] } });
  assert.equal(get('sheet-select').value, '运行数据'); assert.equal(get('sheet-select').hidden, false);
  assert.equal(chart.option.series.length, 1); assert.equal(chart.option.series[0].name, '温度[℃]');
  assert.deepEqual(Array.from(chart.option.series[0].data, r => r[1]), [30, -20, 40]);
  assert.equal(C.formatTime(chart.option.series[0].data[1][0]), '2026-01-01 12:00:00');
  get('sheet-select').value = '另一组'; get('sheet-select').onchange();
  assert.equal(chart.option.series.length, 1); assert.deepEqual(Array.from(chart.option.series[0].data, r => Array.from(r)), [[2, 80], [0, 20]]);
  get('sheet-select').value = '说明'; get('sheet-select').onchange(); assert.equal(get('sheet-select').value, '另一组');
  assert.equal(errors.length, 0);
});

test('Excel drag and drop loads real XLS/XLSB/template bytes; invalid workbooks preserve the chart', async () => {
  const { get, chart, windowEvents } = await setup();
  for (const extension of ['xls', 'xlsb', 'xltm']) {
    get('tab-paste').onclick();
    await windowEvents.drop({ preventDefault() {}, dataTransfer: { files: [binaryFile(`data.${extension}`, await excelBytes(extension))] } });
    assert.equal(get('view-panel').hidden, false); assert.equal(chart.option.series.length, 0);
    get('select-all').onclick(); assert.equal(chart.option.series.length, 2);
  }
  const original = chart.option;
  for (const file of [binaryFile('broken.xlsx', Buffer.from([0x50, 0x4b, 3, 4])), dataFile('fake.xls', '时间,值\n0,2')]) {
    await get('data-input').onchange({ target: { files: [file] } });
    assert.equal(chart.option, original); assert.match(get('toast').textContent, /导入失败/); assert.equal(get('loading').hidden, true);
  }
});

test('PMSD uploads and drops match the chosen config, preserve source timestamps and report timezone fallback', async () => {
  const config = { schema: 'curve-viewer', version: 1, name: '温度配置', timeMode: 'auto', selectedSeries: [{ name: '温度[℃]', occurrence: 1 }] };
  const { get, chart, windowEvents, errors } = await setup([config], '温度配置');
  await get('data-input').onchange({ target: { files: [binaryFile(FILENAME, makePMSD())] } });
  assert.equal(chart.option.series.length, 1); assert.equal(chart.option.series[0].name, '温度[℃]');
  assert.equal(chart.option.series[0].data[0][1], -18.9);
  assert.equal(C.formatTime(chart.option.series[0].data[0][0]), '2026-11-25 18:00:40.964');
  assert.equal(chart.option.series[0].data[1][0] - chart.option.series[0].data[0][0], 637);
  assert.match(get('toast').textContent, /UTC\+08:00/);
  get('tab-paste').onclick();
  await windowEvents.drop({ preventDefault() {}, dataTransfer: { files: [binaryFile('renamed.pmsd', makePMSD())] } });
  assert.equal(get('view-panel').hidden, false); assert.equal(chart.option.series.length, 1);
  assert.match(get('data-notice').textContent, /浏览器本地时区/); assert.equal(errors.length, 0);
});

test('completed PMSD uploads and drops accept END1 without adding an extra point', async () => {
  const config = { schema:'curve-viewer',version:2,type:'selection',name:'压力',selectedSeries:[{name:'压力[bar]',occurrence:1}] };
  const {get,chart,windowEvents,errors}=await setup([config],{selection:'压力'});
  const input=makePMSD({endRecord:true});
  await get('data-input').onchange({target:{files:[binaryFile(FILENAME,input)]}});
  assert.equal(chart.option.series[0].data.length,2);assert.match(get('toast').textContent,/已导入.*2 个时间点/);assert.doesNotMatch(get('toast').textContent,/失败/);
  assert.deepEqual(Array.from(chart.option.series[0].data,row=>row[1]),[1.4,1.6]);
  await windowEvents.drop({preventDefault(){},dataTransfer:{files:[binaryFile(FILENAME,input)]}});
  assert.equal(chart.option.series[0].data.length,2);assert.equal(errors.length,0);
});

test('damaged PMSD uploads keep the existing curves and report validation failures', async () => {
  const { get, chart, enter } = await setup(); enter('时间\t值\n0\t2\n1\t3'); await get('paste-load').onclick(); get('select-all').onclick();
  const original = chart.option, broken = makePMSD(); broken[broken.length - 10] ^= 1;
  const brokenEnd = makePMSD({ endRecord: true }); brokenEnd[brokenEnd.length - 1] ^= 1;
  for (const input of [broken, makePMSD({ version: 3 }), brokenEnd]) {
    await get('data-input').onchange({ target: { files: [binaryFile(FILENAME, input)] } });
    assert.equal(chart.option, original); assert.match(get('toast').textContent, /导入失败.*校验失败|导入失败.*版本/);
    assert.equal(get('loading').hidden, true); assert.equal(get('data-input').value, '');
  }
});

test('new data is unselected until a config is chosen, then only matching configured series load', async () => {
  const config = { schema: 'curve-viewer', version: 1, name: '只看温度', timeMode: 'auto', selectedSeries: [{ name: '温度', occurrence: 1 }, { name: '不存在', occurrence: 1 }] };
  const app = await setup([config]), { get, chart, enter } = app;
  enter('时间\t温度\t压力\n2\t20\t2\n0\t30\t3'); await get('paste-load').onclick();
  assert.equal(chart.option.series.length, 0); assert.equal(get('selection-count').textContent, '0 / 2');
  await get('config-list').children[0].children[0].events.click();
  assert.deepEqual(Array.from(chart.option.series, s => s.name), ['温度']); assert.equal(app.lastConfig(), '只看温度');
  get('select-all').onclick(); assert.equal(chart.option.series.length, 2);
  await get('paste-load').onclick();
  assert.deepEqual(Array.from(chart.option.series, s => s.name), ['温度']); assert.match(get('toast').textContent, /未匹配/);
  await get('data-input').onchange({ target: { files: [dataFile('pressure.csv', '时间,压力\n0,2\n1,3')] } });
  assert.equal(chart.option.series.length, 0); assert.equal(get('selection-count').textContent, '0 / 1');
  await get('config-list').children[0].children[1].events.click();
  assert.equal(app.lastConfig(), '');
  enter('时间\t温度\n0\t20'); await get('paste-load').onclick(); assert.equal(chart.option.series.length, 0);
});

test('last chosen config is remembered across page loads and can be chosen before any data', async () => {
  const configs = ['温度', '压力'].map(name => ({ schema: 'curve-viewer', version: 1, name, timeMode: 'number', selectedSeries: [{ name, occurrence: 1 }] }));
  const first = await setup(configs);
  await first.get('config-list').children[1].children[0].events.click();
  assert.equal(first.chart.option, null); assert.equal(first.lastConfig(), '压力');
  const restored = await setup(first.savedConfigs(), first.lastConfigs());
  assert.equal(restored.get('config-list').children[1].children[0].getAttribute('aria-pressed'), 'true');
  restored.enter('时间\t温度\t压力\n2\t20\t2\n0\t30\t3'); await restored.get('paste-load').onclick();
  assert.deepEqual(Array.from(restored.chart.option.series, s => s.name), ['压力']);
  assert.equal(restored.get('time-mode').value, 'auto'); // Choosing a selection profile does not change time/axis settings.
  const missing = await setup(configs, '已删除');
  missing.enter('时间\t温度\n0\t20'); await missing.get('paste-load').onclick();
  assert.equal(missing.chart.option.series.length, 0);
});

test('source row order and first duplicate hover remain correct through paste, config and time changes', async () => {
  const config = { schema: 'curve-viewer', version: 1, name: '值', timeMode: 'auto', selectedSeries: [{ name: '值', occurrence: 1 }] };
  const { get, chart, enter } = await setup([config], '值');
  const rows = [[8, 80], [20, 200], [0, 0], [8, 81], [5, 50]];
  enter('时间\t值\n' + rows.map(row => row.join('\t')).join('\n')); await get('paste-load').onclick();
  const points = () => Array.from(chart.option.series[0].data, point => Array.from(point));
  assert.deepEqual(points(), rows); assert.deepEqual(Array.from(chart.option.domain), [0, 20]);
  mouse(get, 'click', chart.convertToPixel({}, 8), 150, { ctrlKey: true });
  assert.match(get('hover-card').innerHTML, /第 2 行/); assert.match(get('hover-card').innerHTML, />80</);
  get('time-mode').value = 'number'; get('time-mode').onchange(); assert.deepEqual(points(), rows);
  await get('config-list').children[0].children[0].events.click(); assert.deepEqual(points(), rows);
});

test('clear data removes curves, paste, ranges and queued hover while keeping saved and last chosen configs', async () => {
  const config = { schema: 'curve-viewer', version: 1, name: '温度', timeMode: 'auto', selectedSeries: [{ name: '温度', occurrence: 1 }], axisLimits: { thermal: { min: -25, max: 50 } } };
  const app = await setup([config], '温度'), { get, chart, enter } = app;
  enter('时间\t温度\n0\t20\n10\t30\n20\t40'); await get('paste-load').onclick();
  get('axis-settings').onclick(); axisInput(get, '左1', '上限', 70); applyAxes(get);
  mouse(get, 'click', 400, 150, { ctrlKey: true });
  app.deferFrames(true); mouse(get, 'mousemove', 700);
  get('clear-data').onclick(); app.flushFrames(); app.deferFrames(false);
  assert.equal(chart.option, null); assert.equal(get('selection-count').textContent, '0 / 0');
  assert.equal(get('paste-input').value, ''); assert.equal(get('paste-preview').hidden, true); assert.equal(get('paste-preview-table').children.length, 0);
  for (const id of ['cursor-line', 'cursor-label', 'hover-card']) assert.equal(get(id).hidden, true);
  assert.equal(get('hover-card').innerHTML, ''); assert.equal(get('axis-legend').children.length, 0);
  assert.equal(get('sheet-select').children.length, 0); assert.equal(get('data-notice').hidden, true);
  assert.equal(get('clear-data').disabled, true); assert.match(get('hover-status').textContent, /跟随：关/);
  assert.deepEqual(app.savedConfigs(), C.splitConfig(config)); assert.equal(app.lastConfig(), '温度');
  await get('data-input').onchange({ target: { files: [dataFile('again.csv', '时间,温度\n2,35\n1,25')] } });
  assert.equal(chart.option.series[0].name, '温度'); assert.equal(chart.option.yAxis[0].max, 50);
  assert.deepEqual(Array.from(chart.option.series[0].data, row => row[0]), [2, 1]); assert.equal(get('clear-data').disabled, false);
});

test('clear is disabled during a pending load and works after that load completes', async () => {
  const { get, chart, enter } = await setup();
  enter('时间\t值\n0\t1\n1\t2'); await get('paste-load').onclick();
  let release;
  const next = dataFile('next.csv', '时间,值\n2,3\n0,4');
  const pending = get('data-input').onchange({ target: { files: [{ ...next, arrayBuffer: () => new Promise(resolve => { release = resolve; }) }] } });
  await Promise.resolve(); // The loader awaits initial IndexedDB reads before acquiring its busy state.
  assert.equal(get('clear-data').disabled, true);
  const before = chart.option; get('clear-data').onclick(); assert.equal(chart.option, before);
  assert.equal(get('save-config').disabled,true);assert.equal(get('regex-settings').disabled,true);
  release(await next.arrayBuffer()); await pending;
  assert.equal(get('save-config').disabled,false);assert.equal(get('regex-settings').disabled,false);
  assert.equal(get('clear-data').disabled, false); get('clear-data').onclick(); assert.equal(chart.option, null);
});

test('paste previews and plot labels shorten headers while original-name configs still select the correct duplicate', async () => {
  const first = 'HPODU/设备/告警1 (ID 34000) [1]', second = 'HPODU/设备2/告警1 (ID 35000) [1]';
  const oldConfig = { schema: 'curve-viewer', version: 1, name: '原始列名配置', timeMode: 'auto', selectedSeries: [{ name: second, occurrence: 1 }, { name: first, occurrence: 2 }] };
  const app = await setup([oldConfig]), { get, chart, enter } = app;
  const pasted = ['时间\t' + [first, second, first].join('\t'), '0\t1\t2\t3', '1\t4\t5\t6'].join('\n');
  enter(pasted); await get('paste-preview-button').onclick();
  const [head, body] = get('paste-preview-table').children;
  assert.equal(head.children[0].children[1].textContent, '告警1[1]');
  assert.equal(body.children[0].children[1].textContent, '1');
  assert.equal(get('paste-input').value, pasted);
  await get('paste-load').onclick();
  get('select-all').onclick();
  assert.deepEqual(Array.from(chart.option.series, s => s.name), ['告警1[1]', '告警1[1] (2)', '告警1[1] (3)']);
  await get('config-list').children[0].children[0].events.click();
  assert.deepEqual(Array.from(chart.option.series, s => s.data[0][1]), [2, 3]);
  await get('paste-load').onclick();
  assert.deepEqual(Array.from(chart.option.series, s => s.data[0][1]), [2, 3]);
  assert.doesNotMatch(get('toast').textContent, /未匹配/);
  get('config-name').value = '简称配置'; await get('save-config').onclick();
  const saved = app.savedConfigs().find(c => c.name === '简称配置');
  assert.deepEqual(saved.selectedSeries.map(s => [s.name, s.occurrence]), [['告警1[1]', 2], ['告警1[1]', 3]]);
  get('select-all').onclick();
  await get('config-list').children[1].children[0].events.click();
  assert.deepEqual(Array.from(chart.option.series, s => s.data[0][1]), [2, 3]);
});

function mouse(get, name, x = 100, y = 150, options = {}) {
  let prevented = false;
  const event = { clientX: x + 100, clientY: y + 50, button: 0, ctrlKey: false, ...options, preventDefault() { prevented = true; } };
  get('plot-wrap').events[name]?.(event);
  return prevented;
}
async function hoverData(app) {
  app.enter('时间\t温度\n0\t20\n10\t30\n20\t40');
  await app.get('paste-load').onclick();
  app.get('select-all').onclick();
}

test('Ctrl-left toggles following while plain moves, clicks and out-of-plot shortcuts do not', async () => {
  const app = await setup(), { get } = app;
  assert.equal(mouse(get, 'click', 100, 150, { ctrlKey: true }), false);
  await hoverData(app);
  mouse(get, 'mousemove', 700); mouse(get, 'click', 700);
  assert.equal(get('cursor-line').hidden, true);
  assert.equal(mouse(get, 'click', 10, 150, { ctrlKey: true }), false);
  assert.equal(mouse(get, 'click', 100, 150, { ctrlKey: true }), true);
  assert.equal(get('cursor-line').hidden, false); assert.equal(get('cursor-label').textContent, '0');
  assert.match(get('hover-status').textContent, /跟随：开/);
  mouse(get, 'mousemove', 700); assert.equal(get('cursor-label').textContent, '20');
  const fixed = get('cursor-line').style.cssText;
  mouse(get, 'click', 700, 150, { ctrlKey: true }); mouse(get, 'mousemove', 100);
  assert.equal(get('cursor-line').style.cssText, fixed); assert.equal(get('cursor-label').textContent, '20');
  assert.match(get('hover-status').textContent, /跟随：关/);
  mouse(get, 'click', 100, 150, { ctrlKey: true }); assert.equal(get('cursor-label').textContent, '0');
});

test('leaving the plot keeps markers; Ctrl-right only toggles labels and keeps the dashed line', async () => {
  const app = await setup(), { get } = app; await hoverData(app);
  mouse(get, 'click', 400, 150, { ctrlKey: true });
  const fixed = get('cursor-line').style.cssText, html = get('hover-card').innerHTML;
  mouse(get, 'mousemove', 5); mouse(get, 'mousemove', 400, 480); mouse(get, 'mouseleave');
  assert.equal(get('cursor-line').style.cssText, fixed);
  assert.equal(get('cursor-line').hidden, false); assert.equal(get('cursor-label').hidden, false); assert.equal(get('hover-card').hidden, false);
  assert.equal(get('hover-card').innerHTML, html);
  assert.equal(mouse(get, 'contextmenu', 400, 150, { button: 2 }), false);
  assert.equal(mouse(get, 'contextmenu', 5, 150, { button: 2, ctrlKey: true }), false);
  assert.equal(mouse(get, 'contextmenu', 400, 150, { button: 2, ctrlKey: true }), true);
  assert.equal(get('cursor-label').hidden, true); assert.equal(get('hover-card').hidden, true);
  assert.equal(get('cursor-line').hidden, false); assert.equal(get('cursor-line').style.cssText, fixed);
  mouse(get, 'mousemove', 700); assert.notEqual(get('cursor-line').style.cssText, fixed);
  assert.equal(get('hover-card').hidden, true);
  mouse(get, 'click', 700, 150, { ctrlKey: true }); // Freeze at the last data point.
  const lastLine = get('cursor-line').style.cssText;
  mouse(get, 'contextmenu', 100, 150, { button: 2, ctrlKey: true });
  assert.equal(get('cursor-line').style.cssText, lastLine);
  assert.equal(get('cursor-label').hidden, false); assert.equal(get('cursor-label').textContent, '20');
  assert.equal(get('hover-card').hidden, false); assert.match(get('hover-card').innerHTML, />40</);
  mouse(get, 'mouseleave'); assert.equal(get('hover-card').hidden, false);
});

test('queued moves cannot revive a stopped cursor or markers from previous data', async () => {
  const app = await setup(), { get } = app; await hoverData(app);
  mouse(get, 'click', 100, 150, { ctrlKey: true });
  app.deferFrames(true);
  mouse(get, 'mousemove', 400); mouse(get, 'mousemove', 5); mouse(get, 'mouseleave'); app.flushFrames();
  assert.equal(get('cursor-label').textContent, '10'); assert.equal(get('hover-card').hidden, false);
  mouse(get, 'mousemove', 700); mouse(get, 'click', 700, 150, { ctrlKey: true }); app.flushFrames();
  assert.equal(get('cursor-label').textContent, '10');
  mouse(get, 'click', 100, 150, { ctrlKey: true }); mouse(get, 'mousemove', 700);
  app.deferFrames(false);
  await get('data-input').onchange({ target: { files: [dataFile('new.csv', '时间,值\n100,2\n200,3')] } });
  get('select-all').onclick();
  app.flushFrames();
  assert.equal(get('cursor-line').hidden, true); assert.equal(get('cursor-label').hidden, true); assert.equal(get('hover-card').hidden, true);
  mouse(get, 'mousemove', 700); assert.equal(get('cursor-label').textContent, '200');
});

test('a fixed marker stays anchored through resizing, axis changes, tabs and time zoom', async () => {
  const app = await setup(), { get, chart } = app; await hoverData(app);
  mouse(get, 'click', 400, 150, { ctrlKey: true }); mouse(get, 'click', 400, 150, { ctrlKey: true });
  const firstPosition = get('cursor-line').style.cssText;
  chart.width = 1000; app.resize();
  assert.notEqual(get('cursor-line').style.cssText, firstPosition); assert.equal(get('cursor-label').textContent, '10');
  get('axis-settings').onclick(); axisInput(get, '左1', '下限', -20); applyAxes(get);
  assert.equal(get('cursor-line').hidden, false); assert.equal(get('cursor-label').textContent, '10');
  get('tab-paste').onclick(); assert.equal(get('cursor-line').hidden, true);
  get('tab-view').onclick(); assert.equal(get('cursor-line').hidden, false);
  chart.dispatchAction({ start: 75, end: 100 }); assert.equal(get('cursor-line').hidden, true);
  get('reset-zoom').onclick(); assert.equal(get('cursor-line').hidden, false); assert.equal(get('cursor-label').textContent, '10');
  const beforeError = get('cursor-line').style.cssText;
  await get('data-input').onchange({ target: { files: [dataFile('bad.csv', '时间,值')] } });
  assert.equal(get('cursor-line').style.cssText, beforeError); assert.equal(get('cursor-line').hidden, false);
});

test('collapsing the sidebar resizes the plot and preserves data, selection and fixed labels', async () => {
  const app = await setup(), { get, chart, focused } = app; await hoverData(app);
  mouse(get, 'click', 400, 150, { ctrlKey: true }); mouse(get, 'click', 400, 150, { ctrlKey: true });
  const original = chart.option, selected = get('selection-count').textContent, sizeCount = chart.resizeCount;
  assert.equal(get('series-sidebar').hidden, false);
  get('toggle-sidebar').onclick();
  assert.equal(get('series-sidebar').hidden, true);
  assert.equal(get('view-panel').classList.contains('sidebar-collapsed'), true);
  assert.equal(get('toggle-sidebar').getAttribute('aria-expanded'), 'false');
  assert.ok(chart.resizeCount > sizeCount);
  assert.equal(chart.option, original); assert.equal(get('selection-count').textContent, selected);
  assert.equal(get('cursor-line').hidden, false); assert.equal(get('cursor-label').textContent, '10');
  get('toggle-sidebar').onclick();
  assert.equal(get('series-sidebar').hidden, false); assert.equal(get('toggle-sidebar').getAttribute('aria-expanded'), 'true');
  get('series-sidebar').events.keydown({ key: 'Escape' });
  assert.equal(get('series-sidebar').hidden, true); assert.equal(focused(), get('toggle-sidebar'));
});

function findChild(root, predicate) {
  if (predicate(root)) return root;
  for (const child of root.children || []) { const found = findChild(child, predicate); if (found) return found; }
}
function axisInput(get, position, bound, value) {
  const input = findChild(get('axis-fields'), node => node.attributes?.['aria-label'] === `${position} ${bound}`);
  assert.ok(input, `${position} ${bound} input exists`);
  if (value !== undefined) { input.value = String(value); input.oninput(); }
  return input;
}
const applyAxes = get => get('axis-form').onsubmit({ preventDefault() {} });

function curveAxis(get, name, value) {
  const select = findChild(get('series-axis-list'), node => node.attributes?.['aria-label'] === `${name} 坐标系`);
  assert.ok(select, `${name} coordinate selector exists`);
  if (value !== undefined) { select.value = value; select.onchange(); }
  return select;
}
const applyCurveAxes = get => get('series-axis-form').onsubmit({ preventDefault() {} });

function regionInput(get, id, key, value) {
  const group = C.AXIS_GROUPS.find(g => g.id === id), label = key === 'start' ? '起始比例 (%)' : '终止比例 (%)';
  const input = findChild(get('axis-region-fields'), node => node.attributes?.['aria-label'] === `${group.label}组 ${label}`);
  assert.ok(input, `${id} ${key} region input exists`);
  if (value !== undefined) { input.value = String(value); input.oninput(); }
  return input;
}
const applyRegions = get => get('axis-region-form').onsubmit({ preventDefault() {} });

test('config region settings apply to each group while preserving limits, selection, zoom and markers', async () => {
  const { get, chart, enter } = await setup();
  get('axis-region-settings').onclick(); assert.equal(get('axis-region-dialog').open, true); get('axis-region-cancel').onclick();
  enter('时间\t压力\t温度\t电压\n0\t1\t-20\t226\n10\t3\t30\t320'); await get('paste-load').onclick(); get('select-all').onclick();
  get('axis-settings').onclick(); axisInput(get, '左1', '下限', 0.5); axisInput(get, '左1', '上限', 4.5); applyAxes(get);
  chart.handlers.datazoom({ start: 20, end: 70 });
  mouse(get, 'click', 400, 150, { ctrlKey: true }); mouse(get, 'click', 400, 150, { ctrlKey: true });
  const original = chart.option, marker = get('cursor-label').textContent;
  get('axis-region-settings').onclick(); assert.equal(get('axis-region-fields').children.length, 5);
  regionInput(get, 'small', 'start', 70); regionInput(get, 'thermal', 'start', 25.5); regionInput(get, 'thermal', 'end', 75.5);
  regionInput(get, 'medium', 'start', 5); regionInput(get, 'medium', 'end', 20); regionInput(get, 'auto', 'end', 85);
  applyRegions(get);
  assert.equal(get('axis-region-dialog').open, false);
  assert.deepEqual(Array.from(chart.option.yAxis, a => Array.from(a.band)), [[0.7, 1], [0.255, 0.755], [0.05, 0.2]]);
  assert.deepEqual(Array.from(chart.option.yAxis, a => [a.min, a.max]), Array.from(original.yAxis, a => [a.min, a.max]));
  assert.deepEqual(chart.option.series, original.series); assert.equal(get('selection-count').textContent, '3 / 3');
  assert.equal(chart.option.dataZoom[0].start, 20); assert.equal(chart.option.dataZoom[0].end, 70);
  assert.equal(get('cursor-label').textContent, marker);
  assert.match(get('axis-legend').children[2].title, /25.5%～75.5%/);
  get('series-axis-settings').onclick(); assert.match(curveAxis(get, '温度').children.find(c => c.value === 'thermal').textContent, /25.5%～75.5%/); get('series-axis-cancel').onclick();
  get('time-mode').value = 'number'; get('time-mode').onchange(); get('reset-zoom').onclick();
  assert.deepEqual(Array.from(chart.option.yAxis[1].band), [0.255, 0.755]);
  get('select-none').onclick(); get('select-all').onclick(); assert.deepEqual(Array.from(chart.option.yAxis[0].band), [0.7, 1]);
});

test('invalid region percentages and cancelled reset drafts leave the chart unchanged', async () => {
  const app = await setup(), { get, chart } = app; await hoverData(app);
  const original = chart.option;
  get('axis-region-settings').onclick(); regionInput(get, 'small', 'start', 65);
  for (const [start, end] of [['', 50], ['abc', 80], [0, 'Infinity'], [-1, 80], [20, 101], [50, 50], [70, 20]]) {
    regionInput(get, 'thermal', 'start', start); regionInput(get, 'thermal', 'end', end); applyRegions(get);
    assert.equal(chart.option, original); assert.equal(get('axis-region-dialog').open, true); assert.ok(get('axis-region-error').textContent.length > 0);
  }
  get('axis-region-cancel').onclick(); get('axis-region-settings').onclick();
  assert.equal(regionInput(get, 'small', 'start').value, '80');
  regionInput(get, 'thermal', 'start', 10); regionInput(get, 'thermal', 'end', 60); regionInput(get, 'small', 'start', 65); applyRegions(get);
  const customized = chart.option;
  get('axis-region-settings').onclick(); get('axis-region-reset-all').onclick(); get('axis-region-cancel').onclick(); assert.equal(chart.option, customized);
  get('axis-region-settings').onclick();
  findChild(get('axis-region-fields'), node => node.attributes?.['aria-label'] === `${C.RANGES[1].label}组 恢复默认`).onclick(); applyRegions(get);
  assert.deepEqual(Array.from(chart.option.yAxis[0].band), [0.25, 0.8]);
  get('axis-region-settings').onclick(); assert.equal(regionInput(get, 'small', 'start').value, '65');
  get('axis-region-reset-all').onclick(); applyRegions(get); get('axis-region-settings').onclick();
  assert.equal(regionInput(get, 'small', 'start').value, '80');
});

test('region config persists inactive groups across export, import, reload and clear; legacy config restores defaults', async () => {
  const legacy = { schema: 'curve-viewer', version: 1, name: '旧版', timeMode: 'auto', selectedSeries: [{ name: '温度', occurrence: 1 }] };
  const app = await setup([legacy]), { get, chart, enter } = app;
  const pasted = '时间\t温度\n0\t-20\n10\t30'; enter(pasted); await get('paste-load').onclick();
  get('axis-region-settings').onclick(); regionInput(get, 'thermal', 'start', 10.5); regionInput(get, 'thermal', 'end', 60.5); regionInput(get, 'auto', 'end', 90); applyRegions(get);
  assert.equal(chart.option.series.length, 0); // Editing regions never selects data.
  get('select-all').onclick(); await saveCategory(app, 'selection', '区域曲线');
  const saved = await saveCategory(app, 'axes', '区域配置');
  assert.deepEqual(saved.axisRegions, { thermal: { start: 10.5, end: 60.5 }, auto: { start: 0, end: 90 } });
  await get('export-configs').onclick();assert.deepEqual(JSON.parse(await app.downloads.at(-1).text()).configs.find(c=>c.name===saved.name), saved);
  get('axis-region-settings').onclick(); get('axis-region-reset-all').onclick(); applyRegions(get);
  await get('paste-load').onclick(); assert.deepEqual(Array.from(chart.option.yAxis[0].band), [0.105, 0.605]);
  get('axis-region-settings').onclick(); get('clear-data').onclick();
  assert.equal(get('axis-region-dialog').open, false); assert.equal(get('axis-region-settings').disabled, false); assert.equal(get('axis-region-fields').children.length, 0);
  enter(pasted); await get('paste-load').onclick(); assert.deepEqual(Array.from(chart.option.yAxis[0].band), [0.105, 0.605]);
  const restarted = await setup(app.savedConfigs(), app.lastConfigs()); restarted.enter(pasted); await restarted.get('paste-load').onclick();
  assert.deepEqual(Array.from(restarted.chart.option.yAxis[0].band), [0.105, 0.605]);
  restarted.get('axis-region-settings').onclick(); assert.equal(regionInput(restarted.get, 'auto', 'end').value, '90');
  const other = await setup(); other.enter(pasted); await other.get('paste-load').onclick();
  await other.get('config-input').onchange({ target: { files: [dataFile('regions.json', JSON.stringify(saved))] } });
  other.get('select-all').onclick(); other.get('config-tab-axes').onclick();
  await other.get('config-list').children[0].children[0].events.click(); assert.deepEqual(Array.from(other.chart.option.yAxis[0].band), [0.105, 0.605]);
  const before = other.chart.option;
  await other.get('config-input').onchange({ target: { files: [dataFile('invalid-regions.json', JSON.stringify({ ...saved, axisRegions: { thermal: { start: 80, end: 20 } } }))] } });
  assert.equal(other.chart.option, before); assert.match(other.get('toast').textContent, /起始 < 终止/);
  await get('config-list').children[0].children[0].events.click(); assert.deepEqual(Array.from(chart.option.yAxis[0].band), [0.25, 0.8]);
  get('axis-region-settings').onclick(); assert.equal(regionInput(get, 'auto', 'end').value, '100');
});

test('loaded axes use data extrema and requested bands; manual curve assignments preserve zoom and values', async () => {
  const { get, chart, enter } = await setup();
  enter('时间\t压力\t温度\t电压\t转速\t超范围\n2\t1.4\t-18.9\t226\t2100\t-80\n0\t1.6\t43\t320\t3500\t-60');
  await get('paste-load').onclick(); get('select-all').onclick();
  assert.deepEqual(Array.from(chart.option.yAxis, a => [a.min, a.max, Array.from(a.band)]), [
    [1, 2, [0.8, 1]], [-22, 47, [0.25, 0.8]], [221, 325, [0, 0.25]], [2030, 3570, [0, 1]], [-81, -59, [0, 1]]
  ]);
  const values = JSON.stringify(chart.option.series.map(s => s.data));
  chart.handlers.datazoom({ start: 20, end: 70 });
  get('series-axis-settings').onclick(); curveAxis(get, '压力', 'medium'); applyCurveAxes(get);
  assert.equal(chart.option.yAxis.length, 4);
  const pressureAxis = chart.option.yAxis[chart.option.series[0].yAxisIndex];
  assert.deepEqual([pressureAxis.min, pressureAxis.max, Array.from(pressureAxis.band)], [-15, 336, [0, 0.25]]);
  assert.equal(chart.option.series[0].yAxisIndex, chart.option.series[2].yAxisIndex);
  assert.equal(JSON.stringify(chart.option.series.map(s => s.data)), values);
  assert.equal(get('selection-count').textContent, '5 / 5');
  assert.equal(chart.option.dataZoom[0].start, 20); assert.equal(chart.option.dataZoom[0].end, 70);
  get('axis-settings').onclick(); axisInput(get, '右1', '下限', 100); axisInput(get, '右1', '上限', 250); applyAxes(get);
  assert.deepEqual([chart.option.yAxis[1].min, chart.option.yAxis[1].max, Array.from(chart.option.yAxis[1].band)], [100, 250, [0, 0.25]]);
});

test('curve coordinate drafts survive search, cancel atomically and can be set before selecting curves', async () => {
  const { get, chart, enter } = await setup();
  get('series-axis-settings').onclick(); assert.ok(!get('series-axis-dialog').open);
  enter('时间\t压力\t温度\t空列\n0\t1\t20\t\n1\t3\t30\t'); await get('paste-load').onclick();
  const original = chart.option;
  assert.equal(get('series-axis-settings').disabled, false);
  get('series-axis-settings').onclick();
  assert.equal(get('series-axis-list').children.length, 2);
  curveAxis(get, '压力', 'auto');
  get('series-axis-search').value = '温度'; get('series-axis-search').oninput();
  assert.equal(get('series-axis-list').children.length, 1); curveAxis(get, '温度', 'medium');
  get('series-axis-search').value = ''; get('series-axis-search').oninput(); assert.equal(curveAxis(get, '压力').value, 'auto');
  get('series-axis-cancel').onclick(); assert.equal(chart.option, original);
  get('series-axis-settings').onclick(); assert.equal(curveAxis(get, '压力').value, ''); assert.equal(curveAxis(get, '温度').value, '');
  curveAxis(get, '压力', 'auto'); curveAxis(get, '温度', 'medium'); applyCurveAxes(get);
  assert.equal(chart.option.series.length, 0); assert.equal(get('selection-count').textContent, '0 / 3');
  get('select-all').onclick();
  assert.deepEqual(Array.from(chart.option.yAxis, a => Array.from(a.band)), [[0, 0.25], [0, 1]]);
  const assigned = chart.option;
  get('series-axis-settings').onclick(); get('series-axis-reset').onclick(); get('series-axis-cancel').onclick(); assert.equal(chart.option, assigned);
  get('series-axis-settings').onclick(); assert.equal(curveAxis(get, '压力').value, 'auto');
  get('series-axis-reset').onclick(); applyCurveAxes(get);
  assert.deepEqual(Array.from(chart.option.yAxis, a => Array.from(a.band)), [[0.8, 1], [0.25, 0.8]]);
});

test('coordinate assignments save unchecked curves, export and import, and restore after clearing or reloading', async () => {
  const app = await setup(), { get, chart, enter } = app;
  const pasted = '时间\t压力\t温度\n0\t1\t20\n1\t3\t30';
  enter(pasted); await get('paste-load').onclick();
  get('series-axis-settings').onclick(); curveAxis(get, '压力', 'auto'); curveAxis(get, '温度', 'medium'); applyCurveAxes(get);
  get('series-search').value = '温度'; get('series-search').oninput(); get('select-all').onclick();
  const selection = await saveCategory(app, 'selection', '只看温度');
  const saved = await saveCategory(app, 'axes', '分区配置');
  assert.deepEqual(selection.selectedSeries.map(s => s.name), ['温度']); assert.equal(saved.selectedSeries, undefined);
  assert.equal(saved.seriesAxes, undefined);
  assert.deepEqual(selection.seriesAxes, [{ name: '压力', occurrence: 1, sourceName: '压力', sourceOccurrence: 1, axisId: 'auto' }, { name: '温度', occurrence: 1, sourceName: '温度', sourceOccurrence: 1, axisId: 'medium' }]);
  await get('export-configs').onclick();assert.deepEqual(JSON.parse(await app.downloads.at(-1).text()).configs, [saved]);
  get('series-axis-settings').onclick(); get('series-axis-reset').onclick(); applyCurveAxes(get);
  await get('paste-load').onclick();
  assert.equal(chart.option.series.length, 1); assert.deepEqual(Array.from(chart.option.yAxis[0].band), [0, 0.25]);
  get('series-axis-settings').onclick(); get('clear-data').onclick();
  assert.equal(get('series-axis-dialog').open, false); assert.equal(get('series-axis-settings').disabled, true); assert.equal(get('series-axis-list').children.length, 0);
  enter(pasted); await get('paste-load').onclick(); assert.deepEqual(Array.from(chart.option.yAxis[0].band), [0, 0.25]);
  const restored = await setup(app.savedConfigs(), app.lastConfigs());
  restored.enter(pasted); await restored.get('paste-load').onclick();
  restored.get('select-all').onclick(); assert.deepEqual(Array.from(restored.chart.option.yAxis, a => Array.from(a.band)), [[0, 0.25], [0, 1]]);
  const imported = await setup(); imported.enter(pasted); await imported.get('paste-load').onclick();
  await imported.get('config-input').onchange({ target: { files: [dataFile('bands.json', JSON.stringify(selection))] } });
  imported.get('config-tab-selection').onclick();
  await imported.get('config-list').children[0].children[0].events.click(); assert.deepEqual(Array.from(imported.chart.option.yAxis[0].band), [0, 0.25]);
});

test('selection profiles store only checked or manually assigned series and axis profiles never change their assignments', async () => {
  const app=await setup(),{get,enter,chart}=app;
  const pasted='时间\t已选自动\t已选手动\t未选手动\t未选自动\n0\t1\t2\t3\t4\n1\t2\t3\t4\t5';
  enter(pasted);await get('paste-load').onclick();
  for(const index of [0,1]){const checkbox=get('series-list').children[index].children[0];checkbox.checked=true;checkbox.events.change();}
  await saveCategory(app,'selection','组合');const axes=await saveCategory(app,'axes','坐标');
  get('series-axis-settings').onclick();curveAxis(get,'已选手动','medium');curveAxis(get,'未选手动','auto');applyCurveAxes(get);
  assert.doesNotMatch(get('config-active').textContent,/已修改/);
  get('config-tab-selection').onclick();assert.match(get('config-active').textContent,/已修改/);
  const profile=await saveCategory(app,'selection','组合');
  assert.deepEqual(profile.selectedSeries.map(s=>s.name),['已选自动','已选手动']);
  assert.deepEqual(profile.seriesAxes.map(s=>[s.name,s.axisId]),[['已选手动','medium'],['未选手动','auto']]);
  assert.ok(!JSON.stringify(profile).includes('未选自动'));assert.equal(axes.seriesAxes,undefined);
  await get('export-configs').onclick();const exported=await app.downloads.at(-1).text();assert.equal(app.downloadNames.at(-1),'selection.json');
  get('config-tab-axes').onclick();await get('config-list').children[0].children[0].events.click();
  assert.deepEqual(Array.from(chart.option.yAxis,a=>Array.from(a.band)),[[0.8,1],[0,0.25]]);
  get('series-axis-settings').onclick();assert.equal(curveAxis(get,'未选手动').value,'auto');get('series-axis-cancel').onclick();
  await get('export-configs').onclick();assert.ok(!(await app.downloads.at(-1).text()).includes('seriesAxes'));
  const restored=await setup([], '', {indexedDB:app.idb});restored.enter(pasted);await restored.get('paste-load').onclick();
  assert.deepEqual(Array.from(restored.chart.option.series,s=>s.name),['已选自动','已选手动']);
  restored.get('series-axis-settings').onclick();assert.equal(curveAxis(restored.get,'未选手动').value,'auto');assert.equal(curveAxis(restored.get,'未选自动').value,'');restored.get('series-axis-cancel').onclick();
  const imported=await setup();await imported.get('config-input').onchange({target:{files:[dataFile('selection.json',exported)]}});
  await imported.get('config-list').children[0].children[0].events.click();imported.enter(pasted);await imported.get('paste-load').onclick();
  assert.deepEqual(Array.from(imported.chart.option.series,s=>s.name),['已选自动','已选手动']);
  imported.get('select-none').onclick();const none=await saveCategory(imported,'selection','只保存分配');assert.deepEqual(none.selectedSeries,[]);assert.equal(none.seriesAxes.length,2);
  await imported.get('paste-load').onclick();assert.equal(imported.chart.option.series.length,0);
  imported.get('series-axis-settings').onclick();assert.equal(curveAxis(imported.get,'未选手动').value,'auto');imported.get('series-axis-reset').onclick();applyCurveAxes(imported.get);
  const empty=await saveCategory(imported,'selection','全部自动');assert.deepEqual(empty.selectedSeries,[]);assert.deepEqual(empty.seriesAxes,[]);
});

test('manual assignments match original device headers and duplicate columns and legacy configs reset them', async () => {
  const first = 'HPODU/设备/告警1 (ID 34000) [1]', second = 'HPODU/设备2/告警1 (ID 35000) [1]';
  const selectedSeries = [{ name: second, occurrence: 1 }, { name: first, occurrence: 2 }];
  const config = { schema: 'curve-viewer', version: 1, name: '分配', timeMode: 'auto', selectedSeries,
    seriesAxes: [{ name: second, occurrence: 1, axisId: 'medium' }, { name: first, occurrence: 2, axisId: 'auto' }] };
  const legacy = { schema: 'curve-viewer', version: 1, name: '旧配置', timeMode: 'auto', selectedSeries };
  const { get, chart, enter } = await setup([config, legacy], '分配');
  enter(['时间\t' + [first, second, first].join('\t'), '0\t1\t2\t3', '1\t4\t5\t6'].join('\n')); await get('paste-load').onclick();
  assert.deepEqual(Array.from(chart.option.series, s => [s.name, Array.from(chart.option.yAxis[s.yAxisIndex].band)]), [['告警1[1] (2)', [0, 0.25]], ['告警1[1] (3)', [0, 1]]]);
  get('time-mode').value = 'number'; get('time-mode').onchange(); assert.equal(chart.option.yAxis.length, 2);
  get('config-tab-selection').onclick(); await get('config-list').children[1].children[0].events.click();
  assert.equal(chart.option.yAxis.length, 1); assert.deepEqual(Array.from(chart.option.yAxis[0].band), [0.8, 1]);
  assert.equal(chart.option.yAxis[0].min, 1); assert.equal(chart.option.yAxis[0].max, 7);
});

test('axis dialog applies independent bounds without changing data, selection or time zoom', async () => {
  const { get, chart, enter } = await setup();
  get('axis-settings').onclick(); assert.ok(!get('axis-dialog').open);
  enter('时间\t压力\t温度\t转速\n0\t1\t-20\t12000\n10\t3\t30\t24000');
  await get('paste-load').onclick();
  get('select-all').onclick();
  const original = chart.option;
  chart.handlers.datazoom({ start: 20, end: 70 });
  get('axis-settings').onclick();
  assert.equal(get('axis-dialog').open, true);
  axisInput(get, '左1', '下限', -0.5); axisInput(get, '左1', '上限', 2.5);
  axisInput(get, '左2', '下限', 15000); axisInput(get, '左2', '上限', 22000);
  applyAxes(get);
  assert.equal(get('axis-dialog').open, false);
  assert.equal(chart.option.yAxis[0].min, -0.5); assert.equal(chart.option.yAxis[0].max, 2.5);
  assert.equal(chart.option.yAxis[1].min, -23); assert.equal(chart.option.yAxis[1].max, 33);
  assert.equal(chart.option.yAxis[2].min, 15000); assert.equal(chart.option.yAxis[2].max, 22000);
  assert.deepEqual(chart.option.series, original.series); assert.deepEqual(chart.option.domain, original.domain);
  assert.equal(chart.option.dataZoom[0].start, 20); assert.equal(chart.option.dataZoom[0].end, 70);
  get('select-none').onclick(); assert.equal(get('axis-settings').disabled, true);
  get('select-all').onclick(); assert.equal(chart.option.yAxis[0].min, -0.5);
  get('time-mode').value = 'number'; get('time-mode').onchange();
  assert.equal(chart.option.yAxis[0].max, 2.5);
  get('reset-zoom').onclick(); assert.equal(chart.option.yAxis[0].min, -0.5);
  get('axis-legend').children[1].onclick();
  assert.equal(axisInput(get, '左1', '下限').value, '-0.5');
});

test('invalid bounds and cancelled drafts preserve the existing chart; reset can be cancelled', async () => {
  const { get, chart, enter } = await setup();
  enter('时间\t压力\t温度\n0\t1\t-20\n1\t3\t30'); await get('paste-load').onclick();
  get('select-all').onclick();
  const original = chart.option;
  get('axis-settings').onclick(); axisInput(get, '左1', '上限', 6);
  for (const [min, max] of [['', '4'], ['0', 'abc'], ['0', 'Infinity'], ['5', '5'], ['6', '5'], ['-1e308', '1e308']]) {
    axisInput(get, '右1', '下限', min); axisInput(get, '右1', '上限', max); applyAxes(get);
    assert.equal(chart.option, original); assert.equal(get('axis-dialog').open, true);
    assert.match(get('axis-error').textContent, /右1：/);
  }
  get('axis-cancel').onclick(); get('axis-settings').onclick();
  assert.equal(axisInput(get, '左1', '上限').value, '4');
  axisInput(get, '左1', '上限', 6); applyAxes(get);
  const manual = chart.option;
  get('axis-settings').onclick(); get('axis-reset-all').onclick(); get('axis-cancel').onclick();
  assert.equal(chart.option, manual);
  get('axis-settings').onclick(); assert.equal(axisInput(get, '左1', '上限').value, '6');
  get('axis-reset-all').onclick(); applyAxes(get);
  assert.equal(chart.option.yAxis[0].max, 4); assert.equal(chart.option.yAxis[1].min, -23);
});

test('single-axis reset keeps other manual ranges; loading new data clears unsaved ranges', async () => {
  const { get, chart, enter } = await setup();
  enter('时间\t压力\t温度\n0\t1\t-20\n1\t3\t30'); await get('paste-load').onclick();
  get('select-all').onclick();
  get('axis-settings').onclick(); axisInput(get, '左1', '上限', 6); axisInput(get, '右1', '下限', -25); applyAxes(get);
  get('axis-settings').onclick();
  findChild(get('axis-fields'), node => node.attributes?.['aria-label'] === '左1 恢复自动').onclick(); applyAxes(get);
  assert.equal(chart.option.yAxis[0].max, 4); assert.equal(chart.option.yAxis[1].min, -25);
  await get('data-input').onchange({ target: { files: [dataFile('new.csv', '时间,压力,温度\n0,2,20\n1,4,30')] } });
  get('select-all').onclick();
  assert.equal(chart.option.yAxis[1].min, 19);
});

test('manual ranges survive config export, import and clean reload; legacy config restores automatic ranges', async () => {
  const legacy = { schema: 'curve-viewer', version: 1, name: '自动', timeMode: 'auto', selectedSeries: [{ name: '温度', occurrence: 1 }] };
  const app = await setup([legacy]), { get, chart, enter } = app;
  enter('时间\t温度\n0\t-20\n1\t30'); await get('paste-load').onclick();
  get('select-all').onclick();
  get('axis-settings').onclick(); axisInput(get, '左1', '下限', -25.5); axisInput(get, '左1', '上限', 40.25); applyAxes(get);
  await saveCategory(app, 'selection', '温度选择');
  const saved = await saveCategory(app, 'axes', '手动范围');
  assert.deepEqual(saved.axisLimits, { thermal: { min: -25.5, max: 40.25 } });
  await get('export-configs').onclick();assert.deepEqual(JSON.parse(await app.downloads.at(-1).text()).configs.find(c=>c.name===saved.name), saved);
  get('axis-settings').onclick(); applyAxes(get); // Applying an unchanged dialog keeps the saved config active.
  enter('时间\t温度\n0\t-15\n2\t35'); await get('paste-load').onclick();
  assert.equal(chart.option.yAxis[0].min, -25.5);
  await get('config-list').children[0].children[0].events.click();
  assert.equal(chart.option.yAxis[0].min, -18); assert.equal(chart.option.yAxis[0].max, 38);
  await get('config-list').children[1].children[0].events.click();
  assert.equal(chart.option.yAxis[0].max, 40.25);
  const other = await setup(); other.enter('时间\t温度\n0\t-10\n1\t20'); await other.get('paste-load').onclick();
  await other.get('config-input').onchange({ target: { files: [dataFile('ranges.json', JSON.stringify(saved))] } });
  other.get('select-all').onclick(); other.get('config-tab-axes').onclick();
  await other.get('config-list').children[0].children[0].events.click();
  assert.equal(other.chart.option.yAxis[0].min, -25.5);
  const before = other.chart.option;
  await other.get('config-input').onchange({ target: { files: [dataFile('bad.json', JSON.stringify({ ...saved, axisLimits: { thermal: { min: 2, max: 1 } } }))] } });
  assert.equal(other.chart.option, before); assert.match(other.get('toast').textContent, /下限必须小于上限/);
});

test('file upload loads CSV, JSON and TXT into the chart and rejects unsupported file extensions', async () => {
  const { get, chart } = await setup();
  for (const file of [dataFile('test.csv', '时间,值\n0,2\n1,3'), dataFile('test.txt', '时间\t值\n0\t2\n1\t3'), dataFile('test.json', '[{"time":0,"value":2},{"time":1,"value":3}]')]) {
    get('tab-paste').onclick();
    await get('data-input').onchange({ target: { files: [file] } });
    get('select-all').onclick();
    assert.equal(chart.option.series[0].data.length, 2);
    assert.match(get('chart-subtitle').textContent, /数据/);
    assert.equal(get('view-panel').hidden, false);
    assert.equal(chart.option.series[0].data[1][1], 3);
    assert.equal(get('data-input').value, '');
  }
  const original = chart.option;
  await get('data-input').onchange({ target: { files: [dataFile('file.pdf', 'anything')] } });
  assert.equal(chart.option, original);
  assert.match(get('toast').textContent, /支持 Excel、CSV、JSON、TXT/);
});

test('JSON configuration uploads and drops keep the current chart while data JSON draws curves', async () => {
  const { get, chart, windowEvents } = await setup();
  const original = chart.option;
  const config = JSON.stringify({ schema: 'curve-viewer', version: 1, name: '空配置', timeMode: 'auto', selectedSeries: [] });
  await get('data-input').onchange({ target: { files: [dataFile('config.json', config)] } });
  assert.equal(chart.option, original);
  assert.equal(get('config-count').textContent, '1');
  await windowEvents.drop({ preventDefault() {}, dataTransfer: { files: [dataFile('data.json', '[["时间","值"],[0,2],[1,4]]')] } });
  get('select-all').onclick();
  assert.match(get('chart-subtitle').textContent, /数据/);
  assert.equal(chart.option.series[0].data[1][1], 4);
  assert.equal(get('config-count').textContent, '1');
  const current = chart.option;
  await windowEvents.drop({ preventDefault() {}, dataTransfer: { files: [dataFile('data.xls', 'unsupported')] } });
  assert.equal(chart.option, current);
});
