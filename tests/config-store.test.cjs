const { test } = require('node:test');
const assert = require('node:assert/strict');
const { IDBFactory } = require('fake-indexeddb');
const R = require('../src/config-store.js');
const C = require('../src/core.js');
const selection = (name, value = '温度') => ({ schema:'curve-viewer',version:2,type:'selection',name,selectedSeries:[{name:value,occurrence:1}],seriesAxes:[] });
const axes = name => C.validateProfile({schema:'curve-viewer',version:2,type:'axes',name,timeMode:'auto'});
const regex = name => ({schema:'curve-viewer',version:2,type:'regex',name,headerRule:C.validateHeaderRule()});
const document = (type, configs = []) => ({schema:'curve-viewer-bundle',version:1,type,defaults:C.getDefaults(type),configs});
const fileType = url => url.split('/').at(-1).replace('.json','');

test('IndexedDB saves, updates, reads after reopening and deletes a single local profile', async () => {
  const idb=new IDBFactory(),a=R.create(idb); await a.migrate();
  await a.putMany([selection('同名'),axes('同名'),selection('另一个')],true);
  await a.putMany([selection('同名','压力')],true);
  const snapshot=await a.snapshot();assert.equal(snapshot.configs.length,3);assert.equal(snapshot.configs[0].selectedSeries[0].name,'压力');
  assert.deepEqual(snapshot.last,{selection:{source:'local',name:'同名'},axes:{source:'local',name:'同名'}});
  await a.close();const b=R.create(idb);assert.deepEqual(await b.snapshot(),snapshot);
  await b.remove('selection','同名');const after=await b.snapshot();
  assert.deepEqual(after.configs.map(c=>[c.type,c.name]),[['axes','同名'],['selection','另一个']]);assert.equal(after.last.selection,undefined);assert.equal(after.last.axes.name,'同名');
  await b.close();
});

test('legacy migration is once-only and never resurrects deleted or overwrites edited profiles', async () => {
  const db=R.create(new IDBFactory()),old=selection('旧配置');
  await db.migrate([old],{selection:'旧配置'});await db.putMany([selection('旧配置','压力')]);
  await db.migrate([old],{selection:'旧配置'});assert.equal((await db.snapshot()).configs[0].selectedSeries[0].name,'压力');
  await db.remove('selection','旧配置');await db.migrate([old],{selection:'旧配置'});
  assert.deepEqual(await db.snapshot(),{configs:[],last:{}});await db.close();
});

test('different connections update independent records without replacing the whole database', async () => {
  const idb=new IDBFactory(),a=R.create(idb),b=R.create(idb);
  await Promise.all([a.putMany([selection('甲')]),b.putMany([selection('乙')])]);
  await a.putMany([selection('甲','压力')]);
  const rows=(await b.snapshot()).configs;assert.equal(rows.length,2);assert.equal(rows.find(c=>c.name==='乙').selectedSeries[0].name,'温度');
  await a.remove('selection','甲');assert.equal((await b.snapshot()).configs[0].name,'乙');await a.close();await b.close();
});

test('transaction abort does not report success or leave partial imported profiles', async () => {
  const store=R.create(new IDBFactory());await store.putMany([selection('已有')]);const before=await store.snapshot();
  const db=await store.open(),transaction=db.transaction.bind(db);
  db.transaction=(names,mode)=>{const tx=transaction(names,mode);if(mode==='readwrite')queueMicrotask(()=>tx.abort());return tx;};
  await assert.rejects(()=>store.putMany([selection('甲'),axes('乙')],true));
  await assert.rejects(()=>store.remove('selection','已有'));
  db.transaction=transaction;assert.deepEqual(await store.snapshot(),before);await store.close();
});

test('save resolves on transaction completion and invalid batch leaves storage unchanged', async () => {
  const store=R.create(new IDBFactory()),db=await store.open(),transaction=db.transaction.bind(db);let committed=false;
  db.transaction=(names,mode)=>{const tx=transaction(names,mode);tx.addEventListener('complete',()=>{committed=true;});return tx;};
  await store.putMany([selection('正确')]);assert.equal(committed,true);
  await assert.rejects(()=>store.putMany([selection('不应写入'),{...axes('错误'),axisLimits:{small:{min:2,max:1}}}]));
  assert.deepEqual((await store.snapshot()).configs.map(c=>c.name),['正确']);await store.close();
});

test('last choices distinguish cloud and local names and local deletion never clears the cloud choice', async () => {
  const store=R.create(new IDBFactory());await store.putMany([selection('同名')],true);
  await store.select('selection',{source:'cloud',name:'同名'});await store.remove('selection','同名');
  assert.deepEqual((await store.snapshot()).last.selection,{source:'cloud',name:'同名'});
  await store.putMany([regex('规则')],true);assert.equal((await store.snapshot()).last.regex.source,'local');await store.close();
});

test('a deleted local profile cannot be persisted as a newly selected profile', async () => {
  const store=R.create(new IDBFactory());await store.putMany([selection('删除项')],true);await store.remove('selection','删除项');
  await assert.rejects(()=>store.select('selection',{source:'local',name:'删除项'}));assert.equal((await store.snapshot()).last.selection,undefined);await store.close();
});

test('unavailable IndexedDB fails explicitly without a localStorage replacement', async () => {
  const store=R.create(undefined);await assert.rejects(()=>store.snapshot(),/IndexedDB/);await assert.rejects(()=>store.putMany([selection('配置')]),/IndexedDB/);
});

test('one exported JSON document contains current defaults and all local profiles of its type', () => {
  const input=[selection('甲'),axes('坐标'),selection('乙'),regex('正则')],output=R.bundle(input,'selection');
  assert.equal(output.filename,'selection.json');assert.deepEqual(JSON.parse(output.text).configs,[input[0],input[2]]);assert.deepEqual(R.parseBundle(output.text,'selection'),[input[0],input[2]]);
  assert.deepEqual(R.parseCloudDocument(output.text,'selection').defaults,C.getDefaults('selection'));
  assert.deepEqual(JSON.parse(R.bundle([],'regex').text).configs,[]);
  assert.deepEqual(JSON.parse(R.bundle([],'regex').text).defaults,C.getDefaults('regex'));
  assert.throws(()=>R.parseBundle(JSON.stringify([input[0],input[0]])),/重复/);
  assert.throws(()=>R.parseBundle(JSON.stringify([input[1]]),'selection'),/只能包含/);
  assert.throws(()=>R.parseBundle('[]'.repeat(R.MAX_BYTES)),/10 MB/);
});

test('cloud fetch reads three no-cache default documents and rejects missing or corrupt categories', async () => {
  const calls=[],items=[axes('轴'),regex('规则'),selection('曲线')];
  const fetcher=async(url,options)=>{calls.push([url,options]);return{ok:true,status:200,text:async()=>JSON.stringify(document(fileType(url),[items.find(c=>c.type===fileType(url))]))};};
  const loaded=await R.loadCloud(fetcher);assert.deepEqual(loaded.configs,items);assert.deepEqual(loaded.defaults,C.getDefaults());assert.equal(calls.length,3);assert.ok(calls.every(([,options])=>options.cache==='no-store'&&!options.method));
  await assert.rejects(()=>R.loadCloud(async()=>({ok:false,status:404})),/HTTP 404/);
  await assert.rejects(()=>R.loadCloud(async()=>({ok:true,status:200,text:async()=>'{bad'})),/json/);
  await assert.rejects(()=>R.loadCloud(async()=>({ok:false,status:500})),/HTTP 500/);
  await assert.rejects(()=>R.loadCloud(async()=>({ok:true,status:200,text:async()=>'[]'})),/defaults/);
});

test('cloud reads are fixed to the project directory and reject incorrect types or duplicate names', async () => {
  const calls=[];
  const profiles=await R.loadCloud(async(url)=>{calls.push(url);return {ok:true,status:200,text:async()=>JSON.stringify(document(fileType(url),url.endsWith('axes.json')?[axes('轴')]:[]))};});
  assert.deepEqual(calls,['./configs/axes.json','./configs/regex.json','./configs/selection.json']);assert.deepEqual(profiles.configs,[axes('轴')]);
  await assert.rejects(()=>R.loadCloud(async url=>({ok:true,status:200,text:async()=>JSON.stringify(document(fileType(url),[selection('错类')]))})),/只能包含/);
  await assert.rejects(()=>R.loadCloud(async url=>({ok:true,status:200,text:async()=>JSON.stringify(document(fileType(url),url.endsWith('axes.json')?[axes('重复'),axes('重复')]:[]))})),/重复/);
});

test('cloud request failures identify the failed JSON, resolved address and original cause', async () => {
  const cause = new TypeError('Failed to fetch');
  await assert.rejects(() => R.loadCloud(async () => { throw cause; }, 'https://example.test/tools/curve/index.html'), error => {
    assert.equal(error.code, 'CLOUD_REQUEST_FAILED'); assert.equal(error.cause, cause);
    assert.equal(error.url, 'https://example.test/tools/curve/configs/axes.json');
    assert.match(error.message, /axes\.json.*网络请求失败.*Failed to fetch.*https:/); return true;
  });
  const aborted = new Error('aborted'); aborted.name = 'AbortError';
  await assert.rejects(() => R.loadCloud(async () => { throw aborted; }), /axes\.json.*取消或超时.*\.\/configs\/axes\.json/);
});

test('relative cloud requests resolve next to the page at website root or a nested deployment path', async () => {
  for (const base of ['https://example.test/index.html', 'https://example.test/tools/curve/index.html']) {
    const paths = [], expectedDirectory = new URL('./configs/', base).pathname;
    const profiles = await R.loadCloud(async (relative, options) => {
      assert.equal(options.cache, 'no-store');
      const url = new URL(relative, base); paths.push(url.pathname);
      assert.equal(url.origin, 'https://example.test');
      assert.equal(url.pathname, expectedDirectory + relative.split('/').at(-1));
      return { ok: true, status: 200, text: async () => JSON.stringify(document(fileType(relative),url.pathname.endsWith('/selection.json') ? [selection('网站曲线')] : [])) };
    });
    assert.equal(paths.length, 3); assert.equal(profiles.configs[0].name, '网站曲线');
  }
});

test('existing IndexedDB records migrate axis assignments once and retain the active selection across reopen', async () => {
  const idb = new IDBFactory(), store = R.create(idb), db = await store.open();
  const seriesAxes = [{name:'压力',occurrence:1,sourceName:'压力',sourceOccurrence:1,axisId:'auto'}];
  const oldAxes = {...axes('坐标'),seriesAxes}, oldSelection = selection('曲线');delete oldSelection.seriesAxes;
  await new Promise((resolve,reject)=>{
    const tx=db.transaction(['profiles','settings'],'readwrite');
    for(const [index,config] of [oldAxes,oldSelection].entries())tx.objectStore('profiles').put({id:R.keyOf(config),config,order:index+1,updatedAt:123});
    for(const [type,name] of [['axes','坐标'],['selection','曲线']])tx.objectStore('settings').put({key:'last.'+type,value:{source:'local',name}});
    tx.objectStore('settings').put({key:'legacy-migrated',value:true});
    tx.oncomplete=resolve;tx.onabort=()=>reject(tx.error);
  });
  const migrated=await store.snapshot(), active=migrated.configs.find(c=>c.type==='selection'&&c.name==='曲线');
  assert.deepEqual(active.selectedSeries,oldSelection.selectedSeries);assert.deepEqual(active.seriesAxes,seriesAxes);
  assert.equal(migrated.configs.find(c=>c.type==='axes').seriesAxes,undefined);
  assert.deepEqual(migrated.last.selection,{source:'local',name:'曲线'});
  const raw=await new Promise((resolve,reject)=>{const tx=db.transaction('profiles','readonly'),request=tx.objectStore('profiles').getAll();tx.oncomplete=()=>resolve(request.result);tx.onabort=()=>reject(tx.error);});
  assert.ok(raw.every(row=>row.config.type!=='axes'||!Object.hasOwn(row.config,'seriesAxes')));
  await store.putMany([{...active,seriesAxes:[]}]);
  await store.remove('selection','坐标（旧坐标分配）');await store.close();
  const reopened=R.create(idb), after=await reopened.snapshot();
  assert.equal(after.configs.length,2);assert.deepEqual(after.configs.find(c=>c.type==='selection').seriesAxes,[]);
  await reopened.close();
});

test('old cloud axes and defaults move to selection while exports contain no axis-owned series entries', async () => {
  const seriesAxes=[{name:'压力',occurrence:1,axisId:'auto'}], oldSelection=selection('设备');delete oldSelection.seriesAxes;
  const docs=Object.fromEntries(C.CONFIG_TYPES.map(type=>[type,document(type)]));
  docs.axes.defaults.seriesAxes=seriesAxes;delete docs.selection.defaults.seriesAxes;
  docs.axes.configs=[{...axes('设备'),seriesAxes}];docs.selection.configs=[oldSelection];
  const loaded=await R.loadCloud(async url=>({ok:true,status:200,text:async()=>JSON.stringify(docs[fileType(url)])}));
  assert.equal(loaded.configs.length,2);assert.equal(loaded.defaults.axes.seriesAxes,undefined);assert.deepEqual(loaded.defaults.selection.seriesAxes,seriesAxes);
  const migrated=loaded.configs.find(c=>c.type==='selection');assert.deepEqual(migrated.selectedSeries,oldSelection.selectedSeries);assert.deepEqual(migrated.seriesAxes,seriesAxes);
  assert.ok(!R.bundle(loaded.configs,'axes').text.includes('seriesAxes'));
  assert.deepEqual(JSON.parse(R.bundle(loaded.configs,'selection').text).configs[0].seriesAxes,seriesAxes);
  const imported=R.parseBundle(JSON.stringify(docs.axes));assert.equal(imported[1].name,'设备（旧坐标分配）');assert.deepEqual(imported[1].selectedSeries,[]);
});
