/* Local IndexedDB profiles and read-only project configuration bundles. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./core.js'));
  else root.CurveConfigStore = factory(root.CurveCore);
})(typeof globalThis !== 'undefined' ? globalThis : this, function (C) {
  'use strict';
  const DATABASE = 'curve-explorer-configs', MAX_BYTES = 10 * 1024 * 1024;
  const FILES = Object.freeze({ axes: 'axes.json', regex: 'regex.json', selection: 'selection.json' });
  const keyOf = config => JSON.stringify([config.type, config.name]);
  function reference(value) {
    if (!value || !['local', 'cloud'].includes(value.source) || typeof value.name !== 'string' || !value.name || value.name.length > 120) throw new Error('配置来源或名称无效。');
    return { source: value.source, name: value.name };
  }
  function decodeBundle(text, expectedType) {
    if (typeof text !== 'string' || new TextEncoder().encode(text).length > MAX_BYTES) throw new Error('配置文件超过 10 MB。');
    const value = JSON.parse(text.replace(/^\uFEFF/, ''));
    const document = value?.schema === 'curve-viewer-bundle';
    if (document && (value.version !== 1 || !C.CONFIG_TYPES.includes(value.type) || !Array.isArray(value.configs) || (expectedType && value.type !== expectedType))) throw new Error('配置文档的版本、分类或 configs 数组无效。');
    const defaults = document ? C.validateDefaults(value.type, value.defaults) : null;
    const entries = document ? value.configs : Array.isArray(value) ? value : [value];
    if (entries.length > 1000) throw new Error('单个配置文件最多支持 1,000 项配置。');
    const configs = entries.flatMap(C.splitConfig), seen = new Set();
    for (const config of configs) {
      const type = expectedType || (document ? value.type : null);
      if (type && config.type !== type) throw new Error(`${FILES[type]} 只能包含 ${type} 类型的配置。`);
      const key = keyOf(config);
      if (seen.has(key)) throw new Error(`同类配置名称重复：${config.name}`);
      seen.add(key);
    }
    return { configs, defaults, type: document ? value.type : expectedType };
  }
  function parseBundle(text, expectedType) { return decodeBundle(text, expectedType).configs; }
  function parseCloudDocument(text, type) {
    const document = decodeBundle(text, type);
    if (!document.defaults) throw new Error(`${FILES[type]} 需要包含 defaults 和 configs 的配置文档。`);
    return document;
  }
  function bundle(configs, type) {
    if (!C.CONFIG_TYPES.includes(type)) throw new Error('配置类型无效。');
    const profiles = configs.filter(c => c.type === type).map(C.validateProfile);
    const text = JSON.stringify({ schema: 'curve-viewer-bundle', version: 1, type, defaults: C.getDefaults(type), configs: profiles }, null, 2) + '\n';
    parseBundle(text, type); return { filename: FILES[type], text };
  }
  async function loadCloud(fetcher, baseURL) {
    const results = await Promise.all(C.CONFIG_TYPES.map(async type => {
      const name = FILES[type], relative = './configs/' + name;
      let address = relative;
      if (baseURL) { try { address = new URL(relative, baseURL).href; } catch { /* Keep the original relative address for diagnostics. */ } }
      let response;
      try { response = await fetcher(relative, { cache: 'no-store' }); }
      catch (cause) {
        const reason = cause?.name === 'AbortError' ? '请求已取消或超时' : `网络请求失败（${cause?.message || '无法连接'}）`;
        const error = new Error(`${name} ${reason}。请求地址：${address}`, { cause });
        error.code = 'CLOUD_REQUEST_FAILED'; error.url = address; throw error;
      }
      if (!response.ok) throw new Error(`${name} 读取失败（HTTP ${response.status}）。请求地址：${address}`);
      try { return parseCloudDocument(await response.text(), type); } catch (error) { throw new Error(`${name}：${error.message}`); }
    }));
    return { configs: results.flatMap(document => document.configs), defaults: Object.fromEntries(results.map(document => [document.type, document.defaults])) };
  }
  function create(indexedDB, name = DATABASE) {
    let connection;
    function open() {
      if (connection) return connection;
      connection = new Promise((resolve, reject) => {
        if (!indexedDB?.open) { reject(new Error('当前浏览器无法使用 IndexedDB。')); return; }
        let request, failed = false;
        try { request = indexedDB.open(name, 1); } catch (error) { reject(error); return; }
        request.onupgradeneeded = () => {
          const db = request.result;
          if (!db.objectStoreNames.contains('profiles')) db.createObjectStore('profiles', { keyPath: 'id' });
          if (!db.objectStoreNames.contains('settings')) db.createObjectStore('settings', { keyPath: 'key' });
        };
        request.onerror = () => reject(request.error || new Error('IndexedDB 打开失败。'));
        request.onblocked = () => { failed = true; reject(new Error('配置数据库正在被其他旧页面占用，请关闭旧页面后重试。')); };
        request.onsuccess = () => {
          const db = request.result;
          if (failed) { db.close(); return; }
          db.onversionchange = () => { db.close(); connection = null; };
          db.onclose = () => { connection = null; };
          resolve(db);
        };
      }).catch(error => { connection = null; throw error; });
      return connection;
    }
    async function transaction(mode, action) {
      const db = await open();
      return new Promise((resolve, reject) => {
        let result, tx;
        try { tx = db.transaction(['profiles', 'settings'], mode); }
        catch (error) { reject(error); return; }
        tx.oncomplete = () => { try { resolve(typeof result === 'function' ? result() : result); } catch (error) { reject(error); } };
        tx.onabort = () => reject(tx.error || new Error('配置写入已取消，未保存。'));
        tx.onerror = () => { /* The abort event reports the transaction failure. */ };
        try { result = action(tx.objectStore('profiles'), tx.objectStore('settings'), tx); }
        catch (error) { tx.abort(); reject(error); }
      });
    }
    async function snapshot() {
      return transaction('readonly', (profiles, settings) => {
        const a = profiles.getAll(), b = settings.getAll();
        return () => {
          const configs = a.result.sort((x, y) => (x.order || 0) - (y.order || 0)).map(entry => C.validateProfile(entry.config)), last = {};
          for (const entry of b.result) if (entry.key.startsWith('last.')) last[entry.key.slice(5)] = reference(entry.value);
          return { configs, last };
        };
      });
    }
    async function migrate(configs = [], last = {}) {
      const validated = configs.map(C.validateProfile);
      return transaction('readwrite', (profiles, settings) => {
        const marker = settings.get('legacy-migrated');
        marker.onsuccess = () => {
          if (marker.result) return;
          const existing = profiles.getAll();
          existing.onsuccess = () => {
            const ids = new Set(existing.result.map(entry => entry.id));
            let order = Math.max(0, ...existing.result.map(entry => entry.order || 0)) + 1;
            for (const config of validated) if (!ids.has(keyOf(config))) { ids.add(keyOf(config)); profiles.put({ id: keyOf(config), config, order: order++, updatedAt: Date.now() }); }
          };
          for (const type of C.CONFIG_TYPES) if (last[type] && validated.some(c => c.type === type && c.name === last[type])) {
            const key = 'last.' + type, current = settings.get(key);
            current.onsuccess = () => { if (!current.result) settings.put({ key, value: { source: 'local', name: last[type] } }); };
          }
          settings.put({ key: 'legacy-migrated', value: true });
        };
      });
    }
    async function putMany(configs, choose = false) {
      const validated = configs.map(C.validateProfile);
      return transaction('readwrite', (profiles, settings) => {
        const existing = profiles.getAll();
        existing.onsuccess = () => {
          const entries = new Map(existing.result.map(entry => [entry.id, entry]));
          let order = Math.max(0, ...existing.result.map(entry => entry.order || 0)) + 1;
          for (const config of validated) {
            const id = keyOf(config), entry = { id, config, order: entries.get(id)?.order || order++, updatedAt: Date.now() };
            profiles.put(entry); entries.set(id, entry);
            if (choose) settings.put({ key: 'last.' + config.type, value: { source: 'local', name: config.name } });
          }
        };
      });
    }
    async function remove(type, name) {
      if (!C.CONFIG_TYPES.includes(type)) throw new Error('配置类型无效。');
      return transaction('readwrite', (profiles, settings) => {
        profiles.delete(keyOf({ type, name }));
        const key = 'last.' + type, current = settings.get(key);
        current.onsuccess = () => { if (current.result?.value.source === 'local' && current.result.value.name === name) settings.delete(key); };
      });
    }
    async function select(type, value) {
      if (!C.CONFIG_TYPES.includes(type)) throw new Error('配置类型无效。');
      const ref = reference(value);
      return transaction('readwrite', (profiles, settings, tx) => {
        if (ref.source === 'local') {
          const config = profiles.get(keyOf({ type, name: ref.name }));
          config.onsuccess = () => { if (config.result) settings.put({ key: 'last.' + type, value: ref }); else tx.abort(); };
        } else settings.put({ key: 'last.' + type, value: ref });
      });
    }
    async function close() { if (connection) { (await connection).close(); connection = null; } }
    return { open, snapshot, migrate, putMany, remove, select, close };
  }
  return { DATABASE, MAX_BYTES, FILES, keyOf, parseBundle, parseCloudDocument, bundle, loadCloud, create };
});
