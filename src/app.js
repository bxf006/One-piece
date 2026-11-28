(function () {
  'use strict';
  const C = window.CurveCore;
  const XLSX = window.SheetReader;
  const R = window.CurveConfigStore, configStore = R?.create(window.indexedDB);
  const $ = id => {
    const element = document.getElementById(id);
    if (!element) throw new Error(`页面元素 #${id} 不存在`);
    return element;
  };
  const CONFIG_LABELS = { axes: '坐标轴', regex: '标题正则', selection: '序列' };
  const state = { workbook: null, data: null, selected: new Set(), configs: [], active: null, axisActive: null, regexActive: null, dirty: false, axisDirty: false, regexDirty: false, configType: 'selection', configNames: { axes: '', regex: '', selection: '' }, headerRule: null, filename: '', zoom: [0, 100], axes: [], axisLimits: {}, axisAssignments: {}, axisRegions: {}, busy: false, storageWarning: false };
  Object.assign(state, { cloudConfigs: [], activeSources: {}, lastChoices: {}, loadingConfigs: true, configBusy: false, storageReady: false, storageError: '', cloudMessage: '正在读取云端配置…' });
  let configsReady;
  const activeKey = type => ({ axes: 'axisActive', regex: 'regexActive', selection: 'active' })[type];
  const dirtyKey = type => ({ axes: 'axisDirty', regex: 'regexDirty', selection: 'dirty' })[type];
  const hover = { following: false, labels: true, point: null };
  let hoverPending = false, hoverEpoch = 0, pointer = null;
  let toastTimer;
  function notify(message, error = false) {
    const box = $('toast'); box.textContent = message; box.classList.toggle('error', error); box.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(() => { box.hidden = true; }, error ? 9000 : 5000);
  }
  function escapeHTML(value) { return String(value).replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[ch]); }
  function busy(value) {
    state.busy = value; $('loading').hidden = !(value || state.configBusy); $('clear-data').disabled = value || !state.data;
    $('axis-region-settings').disabled = $('regex-settings').disabled = value; renderConfigs();
  }
  function configBusy(value) { state.configBusy = value; $('loading').hidden = !(state.busy || value); renderConfigs(); }
  function configLocked() { return state.configBusy || state.loadingConfigs; }
  function download(blob, filename) {
    const url = URL.createObjectURL(blob), a = document.createElement('a');
    a.href = url; a.download = filename; document.body.append(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }
  function legacyStore() {
    const configs = [], last = {};
    try {
      const stored = localStorage.getItem('curve-explorer.configs.v2'), legacy = stored === null;
      const data = JSON.parse((legacy ? localStorage.getItem('curve-explorer.configs.v1') : stored) || '[]');
      if (Array.isArray(data)) for (const item of data) { try { C.splitConfig(item); configs.push(item); } catch { /* Preserve other valid old profiles. */ } }
      const oldName = localStorage.getItem('curve-explorer.last-config.v1');
      Object.assign(last, legacy ? { axes: oldName, selection: oldName } : JSON.parse(localStorage.getItem('curve-explorer.last-config.v2') || '{}'));
    }
    catch { /* IndexedDB still works when old localStorage is unavailable. */ }
    return { configs, last };
  }
  function reconcileProfiles(source) {
    const profiles = source === 'cloud' ? state.cloudConfigs : state.configs;
    for (const type of C.CONFIG_TYPES) {
      const active = state[activeKey(type)];
      if (!active || state.activeSources[type] !== source) continue;
      const current = profiles.find(c => c.type === type && c.name === active.name);
      if (!current) { state[activeKey(type)] = null; state.activeSources[type] = null; state[dirtyKey(type)] = false; }
      else if (JSON.stringify(current) !== JSON.stringify(active)) state[dirtyKey(type)] = true;
    }
  }
  async function reloadLocalProfiles() {
    const snapshot = await configStore.snapshot(); state.configs = snapshot.configs; state.lastChoices = snapshot.last;
    state.storageReady = true; state.storageError = ''; reconcileProfiles('local'); return snapshot;
  }
  function restoreChoices() {
    for (const type of C.CONFIG_TYPES) {
      const ref = state.lastChoices[type];
      if (!ref) continue;
      const profiles = ref.source === 'cloud' ? state.cloudConfigs : state.configs;
      const active = profiles.find(c => c.type === type && c.name === ref.name);
      if (!active) continue;
      state[activeKey(type)] = active; state.activeSources[type] = ref.source; state.configNames[type] = active.name;
    }
    $('config-name').value = state.configNames[state.configType];
    if (!C.defaultsReady) return;
    const axes = state.axisActive || C.getDefaults('axes');
    $('time-mode').value = $('paste-time-mode').value = axes.timeMode; state.axisLimits = C.validateAxisLimits(axes.axisLimits); state.axisRegions = C.validateAxisRegions(axes.axisRegions);
    state.headerRule = C.validateHeaderRule(state.regexActive?.headerRule || C.getDefaults('regex').headerRule);
  }
  async function refreshCloud(initial = false) {
    if (!initial) { await configsReady; if (state.busy || configLocked()) return; configBusy(true); }
    try {
      if (window.location?.protocol === 'file:') throw new Error('当前通过 file:// 本地文件打开，浏览器不允许自动读取相邻 JSON。请将 index.html、vendor 和 configs 部署到静态网站，通过 HTTP/HTTPS 网址访问');
      const controller = new AbortController(), timer = setTimeout(() => controller.abort(), 8000);
      let profiles;
      try { profiles = await R.loadCloud((url, options) => window.fetch(url, { ...options, signal: controller.signal }), document.baseURI || window.location?.href); }
      finally { clearTimeout(timer); }
      C.configureDefaults(profiles.defaults); state.cloudConfigs = profiles.configs; reconcileProfiles('cloud');
      if (!initial && !state.data) restoreChoices();
      state.cloudMessage = `已读取默认值及 ${profiles.configs.length} 项云端配置 · 只读`;
      if (!initial) notify('云端配置已刷新；点击对应配置应用。');
    } catch (error) { state.cloudMessage = `云端读取失败：${error.message}；${C.defaultsReady ? '已保留原默认值和列表' : '默认配置尚未载入'}。`; console.error('[曲线查看：云端配置读取]', error); if (!initial) notify(state.cloudMessage, true); }
    finally { if (!initial) configBusy(false); renderConfigs(); }
  }
  async function initializeConfigs() {
    const local = async () => {
      try { const old = legacyStore(); await configStore.migrate(old.configs, old.last); await reloadLocalProfiles(); }
      catch (error) { state.storageError = `本地 IndexedDB 不可用：${error.message}`; }
    };
    await Promise.all([local(), refreshCloud(true)]);
    restoreChoices(); state.loadingConfigs = false; renderConfigs();
  }
  async function refreshLocal() {
    await configsReady; if (state.busy || configLocked()) return;
    try { await reloadLocalProfiles(); renderConfigs(); }
    catch (error) { state.storageError = `本地配置读取失败：${error.message}`; renderConfigs(); }
  }
  const chart = window.CurvePlot?.init($('chart'));
  if (!chart || !XLSX || !window.JSZip || !R) { notify('组件未能加载，请保留 HTML 同目录下完整的 vendor 文件夹。', true); return; }
  function activateTab(name, focus = false) {
    for (const key of ['view', 'paste']) {
      const selected = key === name, tab = $('tab-' + key);
      tab.setAttribute('aria-selected', String(selected)); tab.tabIndex = selected ? 0 : -1;
      $(key + '-panel').hidden = !selected;
    }
    cancelHoverMove();
    if (name !== 'view') hideHover();
    if (focus) $('tab-' + name).focus();
    if (name === 'view') requestAnimationFrame(() => { chart.resize(); refreshHover(); });
  }
  function setSidebarCollapsed(collapsed) {
    $('series-sidebar').hidden = collapsed;
    $('view-panel').classList.toggle('sidebar-collapsed', collapsed);
    const button = $('toggle-sidebar');
    button.setAttribute('aria-expanded', String(!collapsed));
    button.title = collapsed ? '展开序列与配置' : '收起序列与配置'; button.setAttribute('aria-label', button.title);
    cancelHoverMove();
    requestAnimationFrame(() => { if (!$('view-panel').hidden) { chart.resize(); refreshHover(); } });
  }
  function pasteStatus(message, error = false) {
    $('paste-status').textContent = message;
    $('paste-status').classList.toggle('error', error);
    $('paste-input').setAttribute('aria-invalid', String(error));
  }
  function resetPastePreview() {
    $('paste-preview').hidden = true;
    $('paste-preview-table').replaceChildren();
    pasteStatus('粘贴后可先预览，再载入曲线。支持最多 250,000 行数据、512 列。');
  }
  function renderPastePreview(workbook) {
    const sheet = workbook.Sheets['粘贴数据'], bounds = XLSX.utils.decode_range(sheet['!ref']);
    const rows = XLSX.utils.sheet_to_json(sheet, { range: { s: { r: 0, c: 0 }, e: { r: Math.min(8, bounds.e.r), c: Math.min(7, bounds.e.c) } } });
    const head = document.createElement('thead'), body = document.createElement('tbody');
    rows.forEach((row, index) => {
      const tr = document.createElement('tr');
      row.forEach(value => { const cell = document.createElement(index ? 'td' : 'th'); cell.textContent = value == null || value === '' ? '—' : index ? String(value) : C.normalizeHeader(value, state.headerRule); if (!index) cell.scope = 'col'; tr.append(cell); });
      (index ? body : head).append(tr);
    });
    $('paste-preview-table').replaceChildren(head, body);
    $('paste-preview-summary').textContent = `${bounds.e.r.toLocaleString('zh-CN')} 行数据 · ${bounds.e.c + 1} 列（含时间列）· 预览前 8 行、8 列，载入时使用全部数据。`;
    $('paste-preview').hidden = false;
  }
  async function processPaste(load = false) {
    await configsReady; if (state.busy || configLocked()) return;
    let stage = '读取表格';
    busy(true);
    try {
      C.requireDefaults();
      await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      const workbook = XLSX.readPasted($('paste-input').value);
      stage = '生成预览';
      renderPastePreview(workbook);
      if (!load) { pasteStatus('已识别表格。确认列名与数据后，点击“载入曲线”。'); return; }
      const mode = $('paste-time-mode').value;
      stage = '载入曲线';
      replaceData(workbook, '粘贴数据', 'Excel 粘贴数据', false, mode);
      $('time-mode').value = mode;
      if (state.axisActive && state.axisActive.timeMode !== mode) { state.axisDirty = true; renderConfigs(); }
      $('series-search').value = ''; renderSeries();
      const missing = state.active ? state.active.selectedSeries.filter(s => !findConfigSeries(s)).length : 0;
      pasteStatus(`已载入 ${state.data.rows.length.toLocaleString('zh-CN')} 个时间点。粘贴内容保留，可修改后再次载入。`);
      activateTab('view', true);
      notify(`已载入粘贴数据，共 ${state.data.rows.length} 个时间点。${missing ? ` 当前配置有 ${missing} 条序列未匹配。` : ''}`, missing > 0);
    } catch (error) {
      console.error(`[曲线查看：${stage}]`, error);
      pasteStatus(`${stage}失败：${String(error.message || '请检查表格内容').replace(/[。.]$/, '')}。粘贴内容已保留。`, true);
    } finally { busy(false); }
  }
  function selectedSeries() { return state.data ? state.data.series.filter(s => state.selected.has(s.key)) : []; }
  function timeText(row, short = false) {
    if (!row) return '—';
    return state.data.kind === 'time' ? C.formatTime(row.x, short) : state.data.kind === 'value' ? C.formatNumber(row.x) : row.label;
  }
  function matchedSeries() {
    const q = $('series-search').value.trim().toLocaleLowerCase();
    return (state.data?.series || []).filter(s => s.label.toLocaleLowerCase().includes(q));
  }
  function renderSeries() {
    const fragment = document.createDocumentFragment(), list = matchedSeries();
    $('selection-count').textContent = `${state.selected.size} / ${state.data?.series.length || 0}`;
    for (const id of ['series-search', 'select-all', 'select-none', 'sheet-select', 'time-mode', 'reset-zoom']) $(id).disabled = !state.data;
    $('clear-data').disabled = state.busy || !state.data;
    $('series-axis-settings').disabled = !state.data?.series.some(s => s.count);
    $('axis-region-settings').disabled = state.busy;
    $('series-count').textContent = $('series-search').value ? `匹配 ${list.length} 条` : `全部 ${list.length} 条`;
    $('select-all').textContent = $('series-search').value ? '选中匹配' : '全选';
    $('select-none').textContent = $('series-search').value ? '清空匹配' : '清空';
    for (const s of list) {
      const label = document.createElement('label'); label.className = 'series-item';
      const input = document.createElement('input'); input.type = 'checkbox'; input.checked = state.selected.has(s.key); input.setAttribute('aria-label', s.label);
      input.addEventListener('change', () => { input.checked ? state.selected.add(s.key) : state.selected.delete(s.key); state.dirty = true; renderSeries(); renderConfigs(); renderChart(); });
      const content = document.createElement('div'); content.className = 'series-content';
      const title = document.createElement('div'); title.className = 'series-name';
      const swatch = document.createElement('span'); swatch.className = 'swatch'; swatch.style.background = s.color;
      const name = document.createElement('strong'); name.textContent = s.label;
      const group = C.AXIS_GROUPS.find(g => g.id === C.axisIdFor(s, state.axisAssignments));
      name.title = s.count ? `${s.label}｜数据 ${C.formatNumber(s.min)} ~ ${C.formatNumber(s.max)}｜坐标组 ${group.label} · ${C.axisRegionLabel(group, state.axisRegions)}` : `${s.label}｜无有效数值`;
      title.append(swatch, name);
      const meta = document.createElement('div'); meta.className = 'series-meta';
      const minmax = document.createElement('span'); minmax.textContent = s.count ? `${C.formatNumber(s.min)} ~ ${C.formatNumber(s.max)}` : '无有效数值';
      const range = document.createElement('span'); range.className = 'range'; range.textContent = s.count ? group.label : '无坐标轴';
      meta.append(minmax, range); content.append(title, meta); label.append(input, content); fragment.append(label);
    }
    if (!list.length) { const p = document.createElement('p'); p.className = 'empty-list'; p.textContent = state.data ? '没有匹配的序列' : '上传或粘贴数据后，序列将显示在这里。'; fragment.append(p); }
    $('series-list').replaceChildren(fragment);
  }
  function hideHover() { $('hover-card').hidden = $('cursor-line').hidden = $('cursor-label').hidden = true; }
  function cancelHoverMove() { hoverEpoch++; hoverPending = false; pointer = null; }
  function clearHover() { cancelHoverMove(); hover.point = null; hideHover(); }
  function updateHoverStatus() { $('hover-status').textContent = `跟随：${hover.following ? '开' : '关'} · 标签：${hover.labels ? '开' : '关'}`; }
  function axisPosition(index) { return `${index % 2 === 0 ? '左' : '右'}${Math.floor(index / 2) + 1}`; }
  let axisEditors = [], resetAllLimits = false, axisAssignmentDraft = {}, regionEditors = [];
  function clearAxisError() {
    $('axis-error').textContent = '';
    for (const editor of axisEditors) for (const input of [editor.min, editor.max]) input.setAttribute('aria-invalid', 'false');
  }
  function resetAxisEditor(editor) {
    editor.min.value = String(editor.automatic.min); editor.max.value = String(editor.automatic.max);
    editor.manual = false; editor.status.textContent = '自动';
  }
  function openAxisSettings(focusId) {
    if (!state.axes.length) return;
    const automatic = C.axesFor(selectedSeries(), {}, state.axisAssignments, state.axisRegions), fragment = document.createDocumentFragment();
    resetAllLimits = false;
    axisEditors = state.axes.map((axis, index) => {
      const row = document.createElement('fieldset'); row.className = 'axis-editor';
      const heading = document.createElement('legend'); heading.style.color = axis.color;
      const status = document.createElement('span'); status.className = 'axis-mode'; status.textContent = axis.manual ? '手动' : '自动';
      heading.append(`${axisPosition(index)} · ${axis.groupLabel}组 · ${axis.region}`, status);
      const names = document.createElement('p'); names.className = 'axis-series';
      names.textContent = selectedSeries().filter(s => s.count && C.axisIdFor(s, state.axisAssignments) === axis.id).map(s => s.label).join('、'); names.title = names.textContent;
      const fields = document.createElement('div'); fields.className = 'axis-fields';
      const editor = { id: axis.id, position: axisPosition(index), automatic: automatic.find(a => a.id === axis.id), manual: !!axis.manual, status };
      for (const [key, label] of [['min', '下限'], ['max', '上限']]) {
        const field = document.createElement('label'), input = document.createElement('input');
        input.id = `axis-${key}-${axis.id}`; input.type = 'number'; input.step = 'any'; input.value = String(axis[key]);
        input.setAttribute('aria-label', `${editor.position} ${label}`); input.setAttribute('aria-describedby', 'axis-error');
        input.oninput = () => { editor.manual = true; status.textContent = '手动'; clearAxisError(); };
        field.append(label, input); fields.append(field); editor[key] = input;
      }
      const reset = document.createElement('button'); reset.type = 'button'; reset.textContent = '恢复自动';
      reset.setAttribute('aria-label', `${editor.position} 恢复自动`);
      reset.onclick = () => { resetAxisEditor(editor); clearAxisError(); };
      fields.append(reset); row.append(heading, names, fields); fragment.append(row);
      return editor;
    });
    $('axis-fields').replaceChildren(fragment); clearAxisError();
    $('axis-dialog').showModal();
    if (focusId) axisEditors.find(editor => editor.id === focusId)?.min.focus();
  }
  function applyAxisSettings(event) {
    event.preventDefault(); clearAxisError();
    const limits = resetAllLimits ? {} : { ...state.axisLimits };
    for (const editor of axisEditors) {
      if (!editor.manual) { delete limits[editor.id]; continue; }
      const min = C.numeric(editor.min.value), max = C.numeric(editor.max.value);
      try { Object.assign(limits, C.validateAxisLimits({ [editor.id]: { min, max } })); }
      catch (error) {
        $('axis-error').textContent = `${editor.position}：${error.message}`;
        editor.min.setAttribute('aria-invalid', 'true'); editor.max.setAttribute('aria-invalid', 'true');
        (min === null || min >= max ? editor.min : editor.max).focus(); return;
      }
    }
    const changed = [...new Set([...Object.keys(state.axisLimits), ...Object.keys(limits)])].some(id => state.axisLimits[id]?.min !== limits[id]?.min || state.axisLimits[id]?.max !== limits[id]?.max);
    state.axisLimits = limits; if (changed) state.axisDirty = true;
    renderChart(); renderConfigs(); $('axis-dialog').close();
    notify('坐标范围已更新。');
  }
  function renderSeriesAxes() {
    const query = $('series-axis-search').value.trim().toLocaleLowerCase(), fragment = document.createDocumentFragment();
    const list = (state.data?.series || []).filter(s => s.count && s.label.toLocaleLowerCase().includes(query));
    for (const s of list) {
      const row = document.createElement('label'); row.className = 'series-axis-row';
      const description = document.createElement('span'), name = document.createElement('strong'), status = document.createElement('small');
      name.textContent = s.label; name.title = s.label; status.textContent = state.selected.has(s.key) ? '已显示' : '未勾选'; description.append(name, status);
      const select = document.createElement('select'); select.className = 'select-control'; select.setAttribute('aria-label', `${s.label} 坐标系`);
      const choices = [{ id: '', label: `自动匹配（${s.range?.label || '自动'}组）` }, ...C.AXIS_GROUPS.map(group => ({ id: group.id, label: `${group.label}组 · ${C.axisRegionLabel(group, state.axisRegions)}` }))];
      select.replaceChildren(...choices.map(choice => Object.assign(document.createElement('option'), { value: choice.id, textContent: choice.label })));
      select.value = axisAssignmentDraft[s.key] || '';
      select.onchange = () => { if (select.value) axisAssignmentDraft[s.key] = select.value; else delete axisAssignmentDraft[s.key]; };
      row.append(description, select); fragment.append(row);
    }
    if (!list.length) { const empty = document.createElement('p'); empty.className = 'empty-list'; empty.textContent = '没有匹配的数值序列'; fragment.append(empty); }
    $('series-axis-list').replaceChildren(fragment);
  }
  function openSeriesAxes() {
    if (!state.data?.series.some(s => s.count)) return;
    axisAssignmentDraft = { ...state.axisAssignments }; $('series-axis-search').value = '';
    renderSeriesAxes(); $('series-axis-dialog').showModal();
  }
  function applySeriesAxes(event) {
    event.preventDefault();
    if (!state.data) return;
    const keys = new Set([...Object.keys(state.axisAssignments), ...Object.keys(axisAssignmentDraft)]);
    if ([...keys].some(key => state.axisAssignments[key] !== axisAssignmentDraft[key])) state.dirty = true;
    state.axisAssignments = { ...axisAssignmentDraft };
    renderSeries(); renderChart(); renderConfigs(); $('series-axis-dialog').close(); notify('曲线坐标系已更新；保存序列配置后可复用。');
  }
  function clearRegionError() {
    $('axis-region-error').textContent = '';
    for (const editor of regionEditors) for (const key of ['start', 'end']) editor[key].setAttribute('aria-invalid', 'false');
  }
  function resetRegionEditor(editor) {
    const defaults = C.axisRegionFor(editor.group);
    editor.start.value = String(defaults.start); editor.end.value = String(defaults.end);
  }
  function openAxisRegions() {
    if (state.busy || configLocked() || !C.defaultsReady) return;
    const fragment = document.createDocumentFragment();
    regionEditors = C.AXIS_GROUPS.map(group => {
      const row = document.createElement('fieldset'); row.className = 'axis-editor';
      const legend = document.createElement('legend'); legend.textContent = `${group.label}组`;
      const fields = document.createElement('div'); fields.className = 'axis-fields';
      const editor = { group }, region = C.axisRegionFor(group, state.axisRegions);
      for (const [key, label] of [['start', '起始比例 (%)'], ['end', '终止比例 (%)']]) {
        const field = document.createElement('label'), input = document.createElement('input');
        input.type = 'number'; input.min = '0'; input.max = '100'; input.step = 'any'; input.value = String(region[key]);
        input.setAttribute('aria-label', `${group.label}组 ${label}`); input.setAttribute('aria-describedby', 'axis-region-error');
        input.oninput = clearRegionError; field.append(label, input); fields.append(field); editor[key] = input;
      }
      const reset = document.createElement('button'); reset.type = 'button'; reset.textContent = '恢复默认';
      reset.setAttribute('aria-label', `${group.label}组 恢复默认`);
      reset.onclick = () => { resetRegionEditor(editor); clearRegionError(); };
      fields.append(reset); row.append(legend, fields); fragment.append(row); return editor;
    });
    $('axis-region-fields').replaceChildren(fragment); clearRegionError(); $('axis-region-dialog').showModal();
  }
  function applyAxisRegions(event) {
    event.preventDefault(); if (state.busy) return; clearRegionError();
    const regions = {};
    for (const editor of regionEditors) {
      const { group } = editor, start = C.numeric(editor.start.value), end = C.numeric(editor.end.value);
      try { C.validateAxisRegions({ [group.id]: { start, end } }); }
      catch (error) {
        $('axis-region-error').textContent = `${group.label}组：${error.message}`;
        editor.start.setAttribute('aria-invalid', 'true'); editor.end.setAttribute('aria-invalid', 'true');
        (start === null || start < 0 || start >= end ? editor.start : editor.end).focus(); return;
      }
      const defaults = C.axisRegionFor(group);
      if (start !== defaults.start || end !== defaults.end) regions[group.id] = { start, end };
    }
    if (C.AXIS_GROUPS.some(group => {
      const previous = C.axisRegionFor(group, state.axisRegions), next = C.axisRegionFor(group, regions);
      return previous.start !== next.start || previous.end !== next.end;
    })) state.axisDirty = true;
    state.axisRegions = regions; renderSeries(); renderChart(); renderConfigs(); $('axis-region-dialog').close();
    notify('坐标区域已应用；保存为坐标轴配置后可在下次加载时复用。');
  }
  function renderChart() {
    cancelHoverMove();
    if (!state.data) {
      clearHover();
      state.axes = []; chart.clear(); $('axis-legend').replaceChildren(); $('axis-legend').hidden = true; $('axis-settings').disabled = true;
      $('chart-empty').hidden = false;
      $('chart-empty').textContent = '请上传 Excel、CSV、JSON、TXT、PMSD 文件，或在“粘贴数据”中载入数据';
      return;
    }
    const { data } = state, series = selectedSeries();
    const axes = C.axesFor(series, state.axisLimits, state.axisAssignments, state.axisRegions); state.axes = axes; $('axis-settings').disabled = !axes.length; $('axis-legend').hidden = !axes.length;
    const leftCount = Math.ceil(axes.length / 2), rightCount = Math.floor(axes.length / 2);
    const grid = { top: 32, bottom: 66, left: 58 + Math.max(0, leftCount - 1) * 69, right: 28 + rightCount * 66 };
    $('chart-empty').hidden = series.length > 0 && axes.length > 0;
    $('chart-empty').textContent = series.length ? '所选序列没有有效数值' : '勾选左侧序列，开始查看曲线';
    const legend = document.createDocumentFragment();
    const caption = document.createElement('span'); caption.className = 'axis-caption'; caption.textContent = '纵轴范围'; legend.append(caption);
    axes.forEach((a, i) => {
      const tag = document.createElement('button'); tag.type = 'button'; tag.className = 'axis-tag';
      tag.title = `${a.groupLabel}组 · ${a.region}；点击调整上下限`; tag.setAttribute('aria-label', `${axisPosition(i)}，${tag.title}，当前 ${a.label}${a.manual ? '，手动' : '，自动'}`);
      tag.onclick = () => openAxisSettings(a.id);
      const dot = document.createElement('i'); dot.style.background = a.color;
      tag.append(dot, document.createTextNode(`${axisPosition(i)} ${a.groupLabel}组：${a.label}${a.manual ? ' · 手动' : ''}`)); legend.append(tag);
    });
    $('axis-legend').replaceChildren(legend);
    const allAxes = axes.length ? axes : [{ ...C.RANGES[0], id: 'empty', band: [0, 1] }];
    chart.setOption({
      animation: false, useUTC: true, textStyle: { fontFamily: 'Segoe UI, Microsoft YaHei, sans-serif' },
      grid, domain: data.domain, tooltip: { show: false },
      xAxis: { type: data.kind, ...(data.kind === 'category' ? { data: data.rows.map(r => r.label), boundaryGap: false } : { min: 'dataMin', max: 'dataMax' }),
        axisLine: { lineStyle: { color: '#dfe5ef' } }, axisTick: { show: false }, splitLine: { show: true, lineStyle: { color: '#f1f3f8', type: 'dashed' } },
        axisLabel: { color: '#929caf', fontSize: 11, margin: 13, hideOverlap: true, ...(data.kind === 'time' ? { formatter: value => C.formatTime(value, data.domain[1] - data.domain[0] < 86400000) } : {}) },
        axisPointer: { show: false } },
      yAxis: allAxes.map((a, i) => ({ type: 'value', min: a.min, max: a.max, name: a.label, band: a.band || [0, 1], nameLocation: 'end', nameGap: 18,
        nameTextStyle: { color: a.color, fontSize: 10 }, position: i % 2 === 0 ? 'left' : 'right', offset: Math.floor(i / 2) * 69, splitNumber: 5,
        axisLine: { show: true, lineStyle: { color: '#e0e5ef' } }, axisTick: { show: false },
        axisLabel: { color: a.color, fontSize: 11, formatter: value => C.formatNumber(value) },
        splitLine: { show: i === 0, lineStyle: { color: '#edf0f6', type: 'dashed' } } })),
      dataZoom: [
        { type: 'inside', xAxisIndex: 0, filterMode: 'none', start: state.zoom[0], end: state.zoom[1], zoomOnMouseWheel: 'ctrl', moveOnMouseMove: true },
        { type: 'slider', xAxisIndex: 0, filterMode: 'none', start: state.zoom[0], end: state.zoom[1], bottom: 20, height: 25, borderColor: '#e5e9f3', backgroundColor: '#f6f8fc', fillerColor: '#5465e81a', dataBackground: { lineStyle: { color: '#bcc5e9' }, areaStyle: { color: '#e4e8f6' } }, selectedDataBackground: { lineStyle: { color: '#8c9ae6' }, areaStyle: { color: '#cdd4f4' } }, handleStyle: { color: '#fff', borderColor: '#b5bfe4' }, moveHandleSize: 0, showDetail: false, brushSelect: false }
      ],
      series: series.filter(s => s.count).map(s => {
        const column = data.series.indexOf(s), axisIndex = axes.findIndex(a => a.id === C.axisIdFor(s, state.axisAssignments));
        return { id: s.key, name: s.label, type: 'line', yAxisIndex: axisIndex, showSymbol: data.rows.length < 2, symbolSize: 5, connectNulls: false, clip: true,
          lineStyle: { width: 1.8, color: s.color }, itemStyle: { color: s.color }, emphasis: { disabled: true },
          data: data.rows.map((row, index) => [data.kind === 'category' ? index : row.x, row.values[column]]) };
      })
    }, { notMerge: true });
    state.plot = grid;
    refreshHover();
  }
  function updateSummary() {
    const d = state.data;
    if (!d) {
      $('chart-subtitle').textContent = '等待上传或粘贴数据';
      $('time-summary').textContent = '暂无数据';
      $('data-notice').textContent = ''; $('data-notice').hidden = true;
      return;
    }
    const labels = { excel: 'Excel 日期', date: '日期 / 时间', number: '数值时间', text: '文本顺序', 'unix-s': 'Unix 秒', 'unix-ms': 'Unix 毫秒' };
    $('chart-subtitle').textContent = `${d.sheetName} · ${d.timeName} · ${labels[d.detected]}`;
    $('time-summary').textContent = `原始首行 ${timeText(d.rows[0])} · 末行 ${timeText(d.rows.at(-1))} · ${d.rows.length.toLocaleString('zh-CN')} 个时间点`;
    const warnings = [];
    if (d.skipped) warnings.push(`${d.skipped} 行因时间为空或无效而跳过`);
    if (d.missing) warnings.push(`${d.missing} 个空白或非数值单元格按缺失值处理，曲线保留断点`);
    if (state.workbook?.ImportInfo?.timeZoneSource === 'local') warnings.push('PMSD 文件名未能识别原时区，时间按浏览器本地时区显示');
    if (d.detected === 'text') warnings.push('时间未转为数值，当前按文本原始顺序等距显示，可手动更改时间格式');
    $('data-notice').textContent = warnings.join('；'); $('data-notice').hidden = !warnings.length;
  }
  function replaceData(workbook, sheet, filename, keepSelection = false, mode = keepSelection ? $('time-mode').value : state.axisActive?.timeMode || $('time-mode').value) {
    const data = C.readSheet(workbook, sheet, XLSX, mode, state.headerRule); // Parse before changing current state.
    clearHover();
    const oldSelection = state.selected;
    if (!keepSelection) { const axes = state.axisActive || C.getDefaults('axes'); state.axisLimits = C.validateAxisLimits(axes.axisLimits); state.axisRegions = C.validateAxisRegions(axes.axisRegions); state.axisAssignments = {}; state.dirty = false; state.axisDirty = !!state.axisActive && state.axisActive.timeMode !== mode; }
    state.workbook = workbook; state.data = data; state.filename = filename; state.zoom = [0, 100]; state.lastMode = mode;
    state.selected = keepSelection ? new Set(data.series.filter(s => oldSelection.has(s.key)).map(s => s.key)) : new Set();
    $('time-mode').value = mode;
    $('sheet-select').replaceChildren(...workbook.SheetNames.map(name => Object.assign(document.createElement('option'), { value: name, textContent: name, selected: name === sheet })));
    $('sheet-select').hidden = workbook.SheetNames.length < 2;
    if (!keepSelection) matchConfig(state.active || C.getDefaults('selection'), false);
    updateSummary(); renderSeries(); renderChart(); renderConfigs();
  }
  function clearData() {
    if (state.busy || configLocked()) return;
    clearHover();
    state.workbook = state.data = null;
    state.selected = new Set(); state.axes = []; state.axisLimits = {}; state.axisAssignments = {}; state.axisRegions = {}; state.zoom = [0, 100];
    state.filename = ''; state.dirty = state.axisDirty = false; state.lastMode = state.axisActive?.timeMode || (C.defaultsReady ? C.getDefaults('axes').timeMode : '');
    hover.following = false; hover.labels = true;
    for (const id of ['series-search', 'data-input', 'paste-input']) $(id).value = '';
    $('config-name').value = state.configNames[state.configType];
    $('time-mode').value = $('paste-time-mode').value = state.lastMode;
    $('sheet-select').replaceChildren(); $('sheet-select').hidden = true;
    $('hover-card').innerHTML = ''; $('cursor-label').textContent = '';
    axisEditors = []; resetAllLimits = false; $('axis-fields').replaceChildren(); clearAxisError();
    if ($('axis-dialog').open) $('axis-dialog').close();
    axisAssignmentDraft = {}; $('series-axis-list').replaceChildren(); $('series-axis-search').value = '';
    if ($('series-axis-dialog').open) $('series-axis-dialog').close();
    regionEditors = []; $('axis-region-fields').replaceChildren(); clearRegionError();
    if ($('axis-region-dialog').open) $('axis-region-dialog').close();
    if ($('regex-dialog').open) $('regex-dialog').close();
    resetPastePreview(); updateSummary(); renderSeries(); renderChart(); renderConfigs(); updateHoverStatus();
    notify('数据已清空。');
  }
  async function loadDataFile(file) {
    await configsReady; if (!file || state.busy || configLocked()) return;
    if (!XLSX.supportsDataFile(file.name)) { $('data-input').value = ''; return notify('支持 Excel、CSV、JSON、TXT、PMSD 文件，请选择支持的扩展名。', true); }
    if (file.size > 50 * 1024 * 1024) return notify('文件超过 50 MB，请先减少行数或拆分工作表。', true);
    busy(true);
    try {
      const bytes = await file.arrayBuffer();
      await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      const workbook = await XLSX.readDataFile(bytes, file.name);
      if (!workbook.SheetNames.length) throw new Error('文件中没有工作表。');
      let firstError, loaded = false;
      for (const sheet of workbook.SheetNames) {
        try { replaceData(workbook, sheet, file.name, false); loaded = true; break; }
        catch (error) { firstError ??= error; }
      }
      if (!loaded) throw firstError;
      activateTab('view');
      const missing = state.active && !state.dirty ? state.active.selectedSeries.length - state.selected.size : 0;
      const timeZone = workbook.ImportInfo?.timeZone;
      notify(`已导入 ${file.name}，共 ${state.data.rows.length} 个时间点。${timeZone ? ` 时间：${timeZone}。` : ''}${missing ? ` 当前配置有 ${missing} 条序列未匹配。` : ''}`, missing > 0);
    } catch (error) {
      if (error.code === 'DISPLAY_CONFIG') await importConfigs([file]);
      else { console.error('[曲线查看：文件导入]', error); notify(`导入失败：${error.message || '无法解析此文件'}`, true); }
    }
    finally { busy(false); $('data-input').value = ''; }
  }
  function renderConfigs() {
    const type = state.configType, active = state[activeKey(type)], dirty = state[dirtyKey(type)];
    const configs = state.configs.filter(config => config.type === type);
    const cloud = state.cloudConfigs.filter(config => config.type === type), locked = state.busy || configLocked();
    $('config-count').textContent = configs.length + cloud.length;
    for (const category of C.CONFIG_TYPES) $('config-tab-' + category).setAttribute('aria-pressed', String(category === type));
    $('config-name').placeholder = `${CONFIG_LABELS[type]}配置名称`;
    $('save-config').disabled = locked || !state.storageReady || !C.defaultsReady || (type === 'selection' && !state.data);
    $('save-config').title = `保存到本地 IndexedDB：同类同名更新`;
    $('config-active').textContent = `当前：${active ? (state.activeSources[type] === 'cloud' ? '云端 · ' : '本地 · ') + active.name : '未选择'}${dirty ? ' · 已修改' : ''}`;
    $('config-storage-status').textContent = state.loadingConfigs ? '正在读取本地配置…' : state.storageError;
    $('config-storage-status').hidden = !state.loadingConfigs && !state.storageError;
    $('cloud-config-status').textContent = state.cloudMessage;
    for (const id of ['import-config', 'export-configs']) $(id).disabled = locked || !state.storageReady;
    $('export-configs').disabled ||= !C.defaultsReady;
    $('axis-region-settings').disabled = $('regex-settings').disabled = locked || !C.defaultsReady;
    $('refresh-cloud').disabled = locked;
    $('axis-region-settings').hidden = type !== 'axes';
    $('regex-settings').hidden = type !== 'regex';
    function list(profiles, source, target) {
      const fragment = document.createDocumentFragment();
      for (const config of profiles) {
      const chosen = active?.name === config.name && state.activeSources[type] === source;
      const row = document.createElement('div'); row.className = 'config-row' + (chosen ? ' active' : '');
      const select = document.createElement('button'); select.className = 'config-choice';
      select.disabled = locked; select.setAttribute('aria-pressed', chosen ? 'true' : 'false');
      const name = document.createElement('span'); name.textContent = config.name + (chosen && dirty ? ' · 已修改' : ''); name.title = name.textContent;
      const count = document.createElement('small'); count.textContent = type === 'selection' ? `${config.selectedSeries.length} 勾选 · ${config.seriesAxes.length} 分配` : type === 'axes' ? '轴' : config.headerRule.enabled ? '启用' : '停用';
      select.append(name, count); select.addEventListener('click', () => applyConfig(config, source)); row.append(select);
      if (source === 'local') {
        const remove = document.createElement('button'); remove.className = 'remove-config'; remove.textContent = '×'; remove.disabled = locked || !state.storageReady; remove.setAttribute('aria-label', `删除本地配置 ${config.name}`); remove.title = '从本地 IndexedDB 删除此项';
        remove.addEventListener('click', () => removeLocalConfig(config)); row.append(remove);
      }
      fragment.append(row);
      }
      if (!profiles.length) { const p = document.createElement('p'); p.className = 'empty-list'; p.textContent = source === 'local' ? `暂无本地${CONFIG_LABELS[type]}配置。` : `暂无云端${CONFIG_LABELS[type]}配置。`; fragment.append(p); }
      $(target).replaceChildren(fragment);
    }
    list(configs, 'local', 'config-list'); list(cloud, 'cloud', 'cloud-config-list');
  }
  function switchConfigType(type) {
    state.configNames[state.configType] = $('config-name').value;
    state.configType = type; $('config-name').value = state.configNames[type]; renderConfigs();
  }
  function findConfigSeries(entry) {
    if (entry.sourceName !== undefined) {
      const original = state.data.series.find(s => s.sourceName === entry.sourceName && s.sourceOccurrence === entry.sourceOccurrence);
      if (original) return original;
    }
    const key = C.keyOf(entry.name, entry.occurrence);
    return state.data.series.find(s => s.key === key) || state.data.series.find(s => s.sourceName === entry.name && s.sourceOccurrence === entry.occurrence) || state.data.series.find(s => s.defaultName === entry.name && s.defaultOccurrence === entry.occurrence);
  }
  function matchConfig(config, report = true) {
    const missing = [], selected = new Set();
    for (const entry of config.selectedSeries) {
      const s = findConfigSeries(entry);
      if (!s) missing.push(entry.name); else { selected.add(s.key); if (entry.color) s.color = entry.color; }
    }
    state.selected = selected; matchSeriesAxes(config);
    if (missing.length && report) notify(`已匹配 ${selected.size} 条序列，${missing.length} 条未在当前工作表找到：${missing.slice(0, 4).join('、')}${missing.length > 4 ? '…' : ''}`, true);
    return missing;
  }
  function matchSeriesAxes(config) {
    state.axisAssignments = {};
    if (state.data) for (const entry of config.seriesAxes || []) { const s = findConfigSeries(entry); if (s) state.axisAssignments[s.key] = entry.axisId; }
  }
  function replaceCurrentHeaders(data) {
    const old = new Map(state.data.series.map(s => [s.column, { selected: state.selected.has(s.key), axis: state.axisAssignments[s.key], color: s.color }]));
    const selected = new Set(), assignments = {};
    for (const s of data.series) {
      const previous = old.get(s.column);
      if (previous?.selected) selected.add(s.key);
      if (previous?.axis) assignments[s.key] = previous.axis;
      if (previous) s.color = previous.color;
    }
    clearHover(); state.data = data; state.selected = selected; state.axisAssignments = assignments;
  }
  function applyHeaderRule(rule) {
    const validated = C.validateHeaderRule(rule);
    const renamed = state.data ? C.renameHeaders(state.data, validated) : null;
    if (renamed) replaceCurrentHeaders(renamed);
    state.headerRule = validated; resetPastePreview();
  }
  async function applyConfig(config, source = 'local') {
    await configsReady; if (state.busy || configLocked()) return;
    configBusy(true);
    try {
      // Validate/reparse before persisting the selection or changing the plot.
      const data = config.type === 'axes' && state.data && config.timeMode !== state.lastMode ? C.readSheet(state.workbook, state.data.sheetName, XLSX, config.timeMode, state.headerRule) : null;
      let remembered = false;
      if (state.storageReady) {
        try { await configStore.select(config.type, { source, name: config.name }); remembered = true; state.lastChoices[config.type] = { source, name: config.name }; }
        catch (error) { state.storageError = `无法记住配置选择：${error.message}`; }
      }
      let missing = [];
      if (config.type === 'regex') applyHeaderRule(config.headerRule);
      else if (config.type === 'axes') {
        if (data) { replaceCurrentHeaders(data); state.zoom = [0, 100]; }
        state.axisLimits = C.validateAxisLimits(config.axisLimits); state.axisRegions = C.validateAxisRegions(config.axisRegions);
        $('time-mode').value = $('paste-time-mode').value = config.timeMode; state.lastMode = config.timeMode;
      } else if (state.data) missing = matchConfig(config);
      state[activeKey(config.type)] = config; state[dirtyKey(config.type)] = false;
      state.activeSources[config.type] = source;
      state.configNames[config.type] = config.name; $('config-name').value = config.name;
      updateSummary(); renderSeries(); renderChart(); renderConfigs();
      if (!missing.length) notify(`已${state.data ? '应用' : '选择'}${source === 'cloud' ? '云端' : '本地'}${CONFIG_LABELS[config.type]}配置“${config.name}”。${remembered ? '' : '当前选择未能保存到本地数据库。'}`, !remembered);
    } catch (error) { notify(`配置未应用：${error.message}`, true); }
    finally { configBusy(false); }
  }
  async function saveLocalConfig() {
    await configsReady; if (state.busy || configLocked()) return;
    if (!state.storageReady) return notify('本地 IndexedDB 不可用，配置未保存。', true);
    if (!C.defaultsReady) return notify('请先成功读取 configs 中的默认配置。', true);
    const type = state.configType;
    if (type === 'selection' && !state.data) return notify('请先载入数据，再保存序列配置。', true);
    const name = $('config-name').value.trim() || `${CONFIG_LABELS[type]}配置 ${state.configs.filter(c => c.type === type).length + 1}`;
    const identity = s => ({ name: s.name, occurrence: s.occurrence, sourceName: s.sourceName, sourceOccurrence: s.sourceOccurrence });
    const fields = type === 'regex' ? { headerRule: state.headerRule } : type === 'selection' ? {
      selectedSeries: selectedSeries().map(s => ({ ...identity(s), color: s.color })),
      seriesAxes: state.data.series.filter(s => state.axisAssignments[s.key]).map(s => ({ ...identity(s), axisId: state.axisAssignments[s.key] }))
    } : { timeMode: $('time-mode').value, axisLimits: state.axisLimits, axisRegions: state.axisRegions };
    let config;
    try { config = C.validateProfile({ schema: 'curve-viewer', version: 2, type, name, ...fields }); } catch (error) { notify(error.message, true); return; }
    configBusy(true);
    try {
      await configStore.putMany([config], true); await reloadLocalProfiles();
      state[activeKey(type)] = config; state.activeSources[type] = 'local'; state[dirtyKey(type)] = false; state.configNames[type] = name; $('config-name').value = name;
      notify(`${CONFIG_LABELS[type]}配置“${name}”已保存到本地 IndexedDB；同类同名项已更新。`);
    } catch (error) { state.storageError = `本地保存失败：${error.message}`; notify(state.storageError, true); }
    finally { configBusy(false); }
  }
  async function removeLocalConfig(config) {
    await configsReady; if (state.busy || configLocked() || !state.storageReady) return;
    if (!state.configs.includes(config)) return;
    configBusy(true);
    try { await configStore.remove(config.type, config.name); await reloadLocalProfiles(); notify(`已删除本地配置“${config.name}”。`); }
    catch (error) { state.storageError = `本地删除失败：${error.message}`; notify(state.storageError, true); }
    finally { configBusy(false); }
  }
  async function exportLocalType() {
    await configsReady; if (state.busy || configLocked() || !state.storageReady) return;
    const type = state.configType; configBusy(true);
    try {
      await reloadLocalProfiles(); const output = R.bundle(state.configs, type);
      download(new Blob([output.text], { type: 'application/json;charset=utf-8' }), output.filename);
      notify(`已导出本地全部 ${CONFIG_LABELS[type]}配置至 ${output.filename}，可手动替换 configs 中的同名文件。`);
    } catch (error) { notify(`导出失败：${error.message}`, true); }
    finally { configBusy(false); }
  }
  async function importConfigs(files) {
    await configsReady; if (configLocked()) return;
    if (!state.storageReady) return notify('本地 IndexedDB 不可用，无法导入本地配置。', true);
    const jsonFiles = Array.from(files).filter(f => /\.json$/i.test(f.name)).sort((a, b) => a.name.localeCompare(b.name));
    if (!jsonFiles.length) return notify('没有找到 JSON 配置文件。', true);
    let count = 0; const errors = []; configBusy(true);
    try {
    for (const file of jsonFiles) {
      try {
        if (file.size > R.MAX_BYTES) throw new Error('配置文件超过 10 MB');
        const configs = R.parseBundle(await file.text());
        await configStore.putMany(configs); count += configs.length;
      } catch (error) { errors.push(`${file.name}：${error.message}`); }
    }
    await reloadLocalProfiles();
    notify(`已导入 ${count} 个本地配置${errors.length ? `；${errors.length} 个文件失败。${errors[0]}` : '，点击列表中的配置即可应用。'}`, !!errors.length);
    } catch (error) { state.storageError = `本地配置读取失败：${error.message}`; notify(state.storageError, true); }
    finally { $('config-input').value = ''; configBusy(false); }
  }
  function regexDraft() { return C.validateHeaderRule({ enabled: $('regex-enabled').checked, pattern: $('regex-pattern').value, flags: $('regex-flags').value, replacement: $('regex-replacement').value }); }
  function previewRegex() {
    try {
      const format = C.createHeaderFormatter(regexDraft());
      $('regex-preview').textContent = $('regex-sample').value.split(/\r?\n/).filter(Boolean).slice(0,12).map(text => `${text} → ${format(text)}`).join('\n') || '输入原始标题即可预览。';
      $('regex-error').textContent = ''; $('regex-apply').disabled = false;
    } catch (error) { $('regex-error').textContent = error.message; $('regex-preview').textContent = ''; $('regex-apply').disabled = true; }
  }
  function fillRegex(rule) {
    $('regex-enabled').checked = rule.enabled; $('regex-pattern').value = rule.pattern; $('regex-flags').value = rule.flags; $('regex-replacement').value = rule.replacement; previewRegex();
  }
  function openRegex() {
    if (state.busy || configLocked() || !C.defaultsReady) return;
    $('regex-sample').value = state.data ? [state.data.sourceTimeName, ...state.data.series.slice(0,5).map(s => s.sourceName)].join('\n') : '';
    fillRegex(state.headerRule); $('regex-dialog').showModal();
  }
  function applyRegex(event) {
    event.preventDefault(); if (state.busy) return;
    try {
      const rule = regexDraft(), changed = JSON.stringify(rule) !== JSON.stringify(state.headerRule);
      applyHeaderRule(rule); if (changed) state.regexDirty = true;
      updateSummary(); renderSeries(); renderChart(); renderConfigs(); $('regex-dialog').close();
      notify('标题规则已应用，保存为标题正则配置后可复用。');
    } catch (error) { $('regex-error').textContent = error.message; }
  }
  function insidePlot(event) {
    return state.axes.length > 0 && !$('view-panel').hidden && chart.containPixel({ gridIndex: 0 }, [event.offsetX, event.offsetY]);
  }
  function captureHover(event) {
    if (!insidePlot(event)) return;
    const data = state.data;
    const x = chart.convertFromPixel({ xAxisIndex: 0 }, event.offsetX);
    const index = data.kind === 'category' ? Math.max(0, Math.min(data.rows.length - 1, Math.round(x))) : C.nearest(data.rows, x);
    const { top, bottom } = state.plot;
    hover.point = { data, index, yRatio: (event.offsetY - top) / Math.max(1, chart.getHeight() - top - bottom) };
    refreshHover();
  }
  function refreshHover() {
    const point = hover.point;
    if (!point || point.data !== state.data || !state.axes.length || $('view-panel').hidden) return hideHover();
    const data = state.data, row = data.rows[point.index];
    const pixelX = chart.convertToPixel({ xAxisIndex: 0 }, data.kind === 'category' ? point.index : row.x);
    const { top, bottom, left, right } = state.plot, height = chart.getHeight(), width = chart.getWidth();
    if (!Number.isFinite(pixelX) || height <= top + bottom || pixelX < left - 1 || pixelX > width - right + 1) return hideHover();
    const line = $('cursor-line'); line.hidden = false; line.style.cssText = `left:${pixelX}px;top:${top}px;height:${height - top - bottom}px`;
    const label = $('cursor-label'), card = $('hover-card');
    label.hidden = card.hidden = !hover.labels;
    if (!hover.labels) return;
    label.textContent = timeText(row, true); label.style.left = `${Math.min(width - 95, Math.max(95, pixelX))}px`; label.style.top = `${height - bottom + 2}px`;
    card.style.maxHeight = `${height - 16}px`;
    const selected = selectedSeries();
    card.innerHTML = `<div class="hover-time">${escapeHTML(timeText(row))} <span style="font-weight:400;color:#909aaa">· 第 ${row.sourceRow} 行</span></div>` + selected.map(s => `<div class="hover-row"><span class="swatch" style="background:${s.color}"></span><span class="hover-name">${escapeHTML(s.label)}</span><strong>${C.formatNumber(row.values[data.series.indexOf(s)])}</strong></div>`).join('');
    // Keep the saved data point anchored when resizing or changing the view.
    const pixelY = top + point.yRatio * (height - top - bottom);
    const cardWidth = card.offsetWidth;
    card.style.left = `${pixelX + cardWidth + 20 < width ? pixelX + 18 : Math.max(8, pixelX - cardWidth - 18)}px`;
    card.style.top = `${Math.max(6, Math.min(pixelY - 24, height - card.offsetHeight - 8))}px`;
  }
  function plotPointer(event) {
    const rect = $('chart').getBoundingClientRect();
    return { offsetX: event.clientX - rect.left, offsetY: event.clientY - rect.top };
  }
  // Delegate to the wrapper so shortcuts also work over the retained value card.
  $('plot-wrap').addEventListener('mousemove', event => {
    const position = plotPointer(event);
    if (!hover.following || !insidePlot(position)) return;
    pointer = position;
    if (!hoverPending) {
      hoverPending = true; const epoch = hoverEpoch;
      requestAnimationFrame(() => {
        if (epoch !== hoverEpoch) return;
        hoverPending = false;
        if (hover.following && pointer) captureHover(pointer);
      });
    }
  });
  $('plot-wrap').addEventListener('click', event => {
    const position = plotPointer(event);
    if (!event.ctrlKey || event.button !== 0 || !insidePlot(position)) return;
    event.preventDefault(); cancelHoverMove(); hover.following = !hover.following;
    if (hover.following) captureHover(position);
    updateHoverStatus();
  });
  $('plot-wrap').addEventListener('contextmenu', event => {
    if (!event.ctrlKey || !insidePlot(plotPointer(event))) return;
    event.preventDefault(); hover.labels = !hover.labels; refreshHover(); updateHoverStatus();
  });
  chart.on('datazoom', event => { const z = event.batch?.[0] || event; state.zoom = [z.start ?? state.zoom[0], z.end ?? state.zoom[1]]; cancelHoverMove(); refreshHover(); });
  new ResizeObserver(() => { if (!$('view-panel').hidden) { chart.resize(); refreshHover(); } }).observe($('chart'));
  for (const name of ['view', 'paste']) {
    $('tab-' + name).onclick = () => activateTab(name);
    $('tab-' + name).onkeydown = event => {
      if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
      event.preventDefault();
      activateTab(event.key === 'Home' ? 'view' : event.key === 'End' ? 'paste' : name === 'view' ? 'paste' : 'view', true);
    };
  }
  $('paste-time-mode').replaceChildren(...Array.from($('time-mode').options, option => option.cloneNode(true)));
  $('paste-input').oninput = resetPastePreview;
  $('paste-input').onkeydown = event => { if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') { event.preventDefault(); processPaste(true); } };
  $('paste-clear').onclick = () => { $('paste-input').value = ''; resetPastePreview(); $('paste-input').focus(); };
  $('paste-preview-button').onclick = () => processPaste(false);
  $('paste-load').onclick = () => processPaste(true);
  $('upload-button').onclick = () => $('data-input').click();
  $('data-input').onchange = e => loadDataFile(e.target.files[0]);
  $('series-search').oninput = renderSeries;
  $('select-all').onclick = () => { matchedSeries().forEach(s => state.selected.add(s.key)); state.dirty = true; renderSeries(); renderChart(); renderConfigs(); };
  $('select-none').onclick = () => { matchedSeries().forEach(s => state.selected.delete(s.key)); state.dirty = true; renderSeries(); renderChart(); renderConfigs(); };
  $('sheet-select').onchange = () => {
    if (!state.data) return;
    const previous = state.data.sheetName;
    try { replaceData(state.workbook, $('sheet-select').value, state.filename); }
    catch (error) { $('sheet-select').value = previous; notify(error.message, true); }
  };
  $('time-mode').onchange = () => {
    if (!state.data) return;
    const previousMode = state.lastMode || 'auto', oldDirty = state.axisDirty;
    try { state.axisDirty = true; replaceData(state.workbook, state.data.sheetName, state.filename, true); state.lastMode = $('time-mode').value; }
    catch (error) { $('time-mode').value = previousMode; state.axisDirty = oldDirty; notify(error.message, true); }
  };
  $('reset-zoom').onclick = () => { if (!state.data) return; state.zoom = [0, 100]; chart.dispatchAction({ type: 'dataZoom', start: 0, end: 100 }); refreshHover(); };
  $('clear-data').onclick = clearData;
  $('toggle-sidebar').onclick = () => setSidebarCollapsed(!$('series-sidebar').hidden);
  $('series-sidebar').addEventListener('keydown', event => { if (event.key === 'Escape') { setSidebarCollapsed(true); $('toggle-sidebar').focus(); } });
  $('axis-settings').onclick = () => openAxisSettings();
  $('series-axis-settings').onclick = openSeriesAxes;
  $('series-axis-search').oninput = renderSeriesAxes;
  $('series-axis-form').onsubmit = applySeriesAxes;
  $('series-axis-cancel').onclick = () => $('series-axis-dialog').close();
  $('series-axis-reset').onclick = () => { axisAssignmentDraft = {}; renderSeriesAxes(); };
  $('axis-region-settings').onclick = openAxisRegions;
  $('axis-region-form').onsubmit = applyAxisRegions;
  $('axis-region-cancel').onclick = () => $('axis-region-dialog').close();
  $('axis-region-reset-all').onclick = () => { regionEditors.forEach(resetRegionEditor); clearRegionError(); };
  $('axis-form').onsubmit = applyAxisSettings;
  $('axis-cancel').onclick = () => $('axis-dialog').close();
  $('axis-reset-all').onclick = () => { resetAllLimits = true; axisEditors.forEach(resetAxisEditor); clearAxisError(); };
  $('save-config').onclick = saveLocalConfig;
  $('export-configs').onclick = exportLocalType;
  $('refresh-cloud').onclick = () => refreshCloud();
  for (const type of C.CONFIG_TYPES) $('config-tab-' + type).onclick = () => switchConfigType(type);
  $('regex-settings').onclick = openRegex;
  $('regex-form').onsubmit = applyRegex;
  $('regex-cancel').onclick = () => $('regex-dialog').close();
  $('regex-reset').onclick = () => fillRegex(C.DEFAULT_HEADER_RULE);
  for (const id of ['regex-enabled', 'regex-pattern', 'regex-flags', 'regex-replacement', 'regex-sample']) $(id).oninput = previewRegex;
  $('config-name').addEventListener('keydown', e => { if (e.key === 'Enter') return saveLocalConfig(); });
  $('import-config').onclick = () => $('config-input').click();
  $('config-input').onchange = e => importConfigs(e.target.files);
  window.addEventListener('focus', refreshLocal);
  $('help-top').onclick = () => $('help-dialog').showModal();
  $('close-help').onclick = () => $('help-dialog').close();
  let dragDepth = 0;
  window.addEventListener('dragenter', e => { if (!e.dataTransfer?.types.includes('Files')) return; e.preventDefault(); dragDepth++; $('drop-overlay').hidden = false; });
  window.addEventListener('dragover', e => { if (e.dataTransfer?.types.includes('Files')) e.preventDefault(); });
  window.addEventListener('dragleave', e => { e.preventDefault(); if (--dragDepth <= 0) { dragDepth = 0; $('drop-overlay').hidden = true; } });
  window.addEventListener('drop', async e => {
    e.preventDefault(); dragDepth = 0; $('drop-overlay').hidden = true;
    const files = Array.from(e.dataTransfer?.files || []);
    for (const file of files) await loadDataFile(file);
  });
  updateSummary(); renderSeries(); renderConfigs(); renderChart(); updateHoverStatus();
  configsReady = initializeConfigs(); window.CurveAppReady = configsReady;
  $('sheet-select').hidden = true; setSidebarCollapsed((window.innerWidth || 1280) <= 800);
})();
