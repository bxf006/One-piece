/* Curve Viewer — data and configuration helpers. */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) {
    api.configureDefaults(Object.fromEntries(api.CONFIG_TYPES.map(type => [type, require('../configs/' + type + '.json').defaults])));
    module.exports = api;
  }
  else root.CurveCore = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  const AXIS_IDS = ['small', 'thermal', 'medium', 'large', 'auto'];
  const PALETTE = [], RANGES = [], AXIS_GROUPS = [];
  let defaults = null, DEFAULT_HEADER_RULE;
  function requireDefaults() {
    if (!defaults) throw new Error('默认配置尚未载入，请检查网站 configs 中的三个 JSON 文件并刷新云端配置。');
    return defaults;
  }
  function getDefaults(type) {
    const value = requireDefaults();
    return JSON.parse(JSON.stringify(type ? value[type] : value));
  }
  function numeric(value) {
    if (typeof value === 'number') return Number.isFinite(value) ? value : null;
    if (typeof value !== 'string' || !value.trim()) return null;
    const s = value.trim();
    if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(s)) return null;
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  }
  function rangeFor(min, max) {
    return RANGES.find(range => min >= range.min && max <= range.max) || null;
  }
  function colorAt(index) {
    const color = requireDefaults().selection.fallbackColor;
    return PALETTE[index] || `hsl(${Math.round(index * color.hueStep % 360)}, ${color.saturation}%, ${color.lightness}%)`;
  }
  function keyOf(name, occurrence = 1) { return JSON.stringify([name, occurrence]); }
  function validateHeaderRule(value = DEFAULT_HEADER_RULE) {
    if (!value || typeof value !== 'object' || Array.isArray(value) || typeof value.enabled !== 'boolean' || typeof value.pattern !== 'string' || !value.pattern || value.pattern.length > 1000 || typeof value.flags !== 'string' || !/^[gimsu]*$/.test(value.flags) || typeof value.replacement !== 'string' || value.replacement.length > 1000) throw new Error('标题规则需要启用状态、1～1000 字符的正则、g/i/m/s/u 标志及替换文本。');
    try { new RegExp(value.pattern, value.flags); } catch (error) { throw new Error(`正则表达式无效：${error.message}`); }
    return { enabled: value.enabled, pattern: value.pattern, flags: value.flags, replacement: value.replacement };
  }
  function createHeaderFormatter(value = DEFAULT_HEADER_RULE) {
    const rule = validateHeaderRule(value), regex = new RegExp(rule.pattern, rule.flags);
    return value => {
      const text = String(value ?? '').trim();
      if (!rule.enabled) return text;
      regex.lastIndex = 0;
      return text.replace(regex, rule.replacement).trim() || text;
    };
  }
  function normalizeHeader(value, rule = DEFAULT_HEADER_RULE) { return createHeaderFormatter(rule)(value); }
  function renameHeaders(data, rule) {
    const format = createHeaderFormatter(rule), occurrences = new Map();
    const series = data.series.map(s => {
      const name = format(s.sourceName), occurrence = (occurrences.get(name) || 0) + 1;
      occurrences.set(name, occurrence);
      return { ...s, name, occurrence, key: keyOf(name, occurrence), label: name + (occurrence > 1 ? ` (${occurrence})` : '') };
    });
    return { ...data, series, timeName: format(data.sourceTimeName) };
  }
  function parseTextTime(value) {
    if (value instanceof Date) return Number.isFinite(value.getTime()) ? value.getTime() : null;
    if (typeof value !== 'string') return null;
    const s = value.trim();
    const clock = /^(\d{1,3}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?$/.exec(s);
    if (clock) {
      const [, h, m, sec = '0', ms = '0'] = clock;
      if (+m > 59 || +sec > 59) return null;
      return ((+h * 60 + +m) * 60 + +sec) * 1000 + +ms.padEnd(3, '0');
    }
    const date = /^(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})日?(?:[ T](\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?)?$/.exec(s);
    if (date) {
      const [, y, mo, d, h = '0', m = '0', sec = '0', ms = '0'] = date;
      if (+mo < 1 || +mo > 12 || +d < 1 || +h > 23 || +m > 59 || +sec > 59) return null;
      const t = Date.UTC(+y, +mo - 1, +d, +h, +m, +sec, +ms.padEnd(3, '0'));
      const check = new Date(t);
      return check.getUTCMonth() === +mo - 1 && check.getUTCDate() === +d ? t : null;
    }
    if (/^\d{4}-\d{2}-\d{2}T.*(?:Z|[+-]\d{2}:?\d{2})$/i.test(s)) {
      const t = Date.parse(s); return Number.isFinite(t) ? t : null;
    }
    return null;
  }
  function excelTime(value, XLSX, date1904) {
    const n = numeric(value);
    if (n === null) return parseTextTime(value);
    const d = XLSX.SSF.parse_date_code(n, { date1904 });
    if (!d) return null;
    return Date.UTC(d.y, d.m - 1, d.d, d.H, d.M, d.S, Math.round((d.u || 0) * 1000));
  }
  function readSheet(workbook, sheetName, XLSX, mode = requireDefaults().axes.timeMode, headerRule = DEFAULT_HEADER_RULE) {
    requireDefaults();
    const formatHeader = createHeaderFormatter(headerRule);
    const sheet = workbook.Sheets[sheetName];
    if (!sheet || !sheet['!ref']) throw new Error('工作表为空，请选择含有数据的工作表。');
    const bounds = XLSX.utils.decode_range(sheet['!ref']);
    if (bounds.e.r > 250000 || bounds.e.c > 511) throw new Error('单张工作表最多支持 250,000 行和 512 列，请先缩小数据范围。');
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true, blankrows: true, range: { s: { r: 0, c: 0 }, e: bounds.e } });
    if (raw.length < 2 || raw[0].length < 2) throw new Error('需要至少两列：第一列为时间，第二列起为数值序列；第一行填写名称。');
    const headers = raw[0];
    const occurrences = new Map(), sourceOccurrences = new Map(), defaultOccurrences = new Map(), defaultFormat = createHeaderFormatter();
    const series = [];
    for (let c = 1; c <= bounds.e.c; c++) {
      if (headers[c] == null && !raw.slice(1).some(row => row[c] != null)) continue;
      const sourceName = String(headers[c] ?? `序列 ${c}`).trim() || `序列 ${c}`;
      const name = formatHeader(sourceName);
      const defaultName = defaultFormat(sourceName), defaultOccurrence = (defaultOccurrences.get(defaultName) || 0) + 1;
      defaultOccurrences.set(defaultName, defaultOccurrence);
      const sourceOccurrence = (sourceOccurrences.get(sourceName) || 0) + 1;
      sourceOccurrences.set(sourceName, sourceOccurrence);
      const occurrence = (occurrences.get(name) || 0) + 1;
      occurrences.set(name, occurrence);
      series.push({ key: keyOf(name, occurrence), name, occurrence, sourceName, sourceOccurrence, defaultName, defaultOccurrence, label: name + (occurrence > 1 ? ` (${occurrence})` : ''), column: c, color: colorAt(series.length), min: Infinity, max: -Infinity, count: 0 });
    }
    if (!series.length) throw new Error('没有找到数值序列。');
    const firstIndex = raw.findIndex((r, i) => i > 0 && r[0] != null && String(r[0]).trim() !== '');
    if (firstIndex < 0) throw new Error('第一列没有有效时间。');
    const first = raw[firstIndex][0];
    const firstCell = sheet[XLSX.utils.encode_cell({ r: firstIndex, c: 0 })];
    let detected = mode;
    if (mode === 'auto') {
      if (firstCell?.z && XLSX.SSF.is_date(firstCell.z) && typeof first === 'number') detected = 'excel';
      else if (numeric(first) !== null) detected = 'number';
      else if (parseTextTime(first) !== null) detected = 'date';
      else detected = 'text';
    }
    const date1904 = !!workbook.Workbook?.WBProps?.date1904;
    const kind = ['excel', 'date', 'unix-s', 'unix-ms'].includes(detected) ? 'time' : detected === 'text' ? 'category' : 'value';
    const rows = [];
    let skipped = 0, missing = 0, minX = Infinity, maxX = -Infinity;
    for (let r = 1; r < raw.length; r++) {
      const row = raw[r];
      if (row.every(v => v == null || v === '')) continue;
      let x = null;
      if (row[0] != null && String(row[0]).trim() !== '') {
        if (detected === 'text') x = rows.length;
        else if (detected === 'excel') x = excelTime(row[0], XLSX, date1904);
        else if (detected === 'date') x = parseTextTime(row[0]);
        else { x = numeric(row[0]); if (x !== null && detected === 'unix-s') x *= 1000; }
      }
      if (x === null || !Number.isFinite(x) || (kind === 'time' && Math.abs(x) > 8640000000000000)) { skipped++; continue; }
      const values = series.map(s => {
        const cell = sheet[XLSX.utils.encode_cell({ r, c: s.column })];
        const value = cell?.t === 'e' ? null : numeric(row[s.column]);
        if (value === null) missing++;
        else { s.min = Math.min(s.min, value); s.max = Math.max(s.max, value); s.count++; }
        return value;
      });
      rows.push({ x, label: firstCell && kind === 'time' ? formatTime(x, false) : String(row[0]), values, sourceRow: r + 1 });
      minX = Math.min(minX, x); maxX = Math.max(maxX, x);
    }
    if (!rows.length) throw new Error('没有可绘制的数据。请检查第一列，或更改时间格式。');
    if (!series.some(s => s.count)) throw new Error('第二列起没有有效数值。请检查单元格内容。');
    for (const s of series) { s.range = s.count ? rangeFor(s.min, s.max) : null; if (!s.count) s.min = s.max = null; }
    const sourceTimeName = String(headers[0] || '时间');
    return { rows, series, domain: [minX, maxX], kind, detected, skipped, missing, sourceTimeName, timeName: formatHeader(sourceTimeName), sheetName };
  }
  function formatNumber(n) {
    if (n == null || !Number.isFinite(n)) return '—';
    return new Intl.NumberFormat('zh-CN', { maximumSignificantDigits: 8 }).format(n);
  }
  function formatTime(t, short = false) {
    const d = new Date(t);
    const clock = d.toISOString().slice(11, 23).replace(/\.000$/, '');
    if (short || (t >= 0 && t < 86400000)) return clock;
    return d.toISOString().slice(0, 10) + ' ' + clock;
  }
  function dataRange(min, max) {
    const rule = requireDefaults().axes.autoBounds, factor = rule.paddingPercent / 100, span = max - min;
    const padding = min === max ? Math.max(rule.minimumConstantPadding, Math.abs(min) * factor) : Number.isFinite(span) ? span * factor : max * factor - min * factor;
    const lower = Number.isFinite(min - padding) ? min - padding : min, upper = Number.isFinite(max + padding) ? max + padding : max;
    return { min: rule.roundOutward ? Math.floor(lower) : lower, max: rule.roundOutward ? Math.ceil(upper) : upper };
  }
  function validateAxisLimits(value = {}) {
    if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('坐标范围必须是对象。');
    const limits = {}, ids = new Set(AXIS_IDS);
    for (const [id, range] of Object.entries(value)) {
      if (!ids.has(id)) throw new Error('配置包含未知坐标轴。');
      if (!range || typeof range !== 'object' || Array.isArray(range) || !Number.isFinite(range.min) || !Number.isFinite(range.max)) throw new Error('坐标上下限必须是有效数字。');
      if (range.min >= range.max) throw new Error('坐标下限必须小于上限。');
      if (!Number.isFinite(range.max - range.min)) throw new Error('坐标范围过大，请缩小上下限。');
      limits[id] = { min: range.min, max: range.max };
    }
    return limits;
  }
  function axisIdFor(series, assignments = {}) { return assignments[series.key] || series.range?.id || 'auto'; }
  function validateAxisRegions(value = {}) {
    if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('坐标区域必须是对象。');
    const regions = {}, ids = new Set(AXIS_IDS);
    for (const [id, region] of Object.entries(value)) {
      if (!ids.has(id)) throw new Error('配置包含未知坐标组。');
      if (!region || typeof region !== 'object' || Array.isArray(region) || !Number.isFinite(region.start) || !Number.isFinite(region.end)) throw new Error('区域起始、终止比例必须是有效数字。');
      if (region.start < 0 || region.end > 100 || region.start >= region.end) throw new Error('区域比例须满足 0 ≤ 起始 < 终止 ≤ 100。');
      regions[id] = { start: region.start, end: region.end };
    }
    return regions;
  }
  function axisRegionFor(group, regions = {}) { return regions[group.id] || { start: group.band[0] * 100, end: group.band[1] * 100 }; }
  function axisRegionLabel(group, regions = {}) {
    const region = axisRegionFor(group, regions);
    return `${formatNumber(region.start)}%～${formatNumber(region.end)}%`;
  }
  function axesFor(series, limits = {}, assignments = {}, regions = {}) {
    return AXIS_GROUPS.flatMap(group => {
      const members = series.filter(s => s.count && axisIdFor(s, assignments) === group.id);
      if (!members.length) return [];
      const range = limits[group.id] || dataRange(Math.min(...members.map(s => s.min)), Math.max(...members.map(s => s.max)));
      const region = axisRegionFor(group, regions);
      return [{ ...group, ...range, band: [region.start / 100, region.end / 100], region: axisRegionLabel(group, regions), groupLabel: group.label, manual: !!limits[group.id], label: `${formatNumber(range.min)} ~ ${formatNumber(range.max)}` }];
    });
  }
  function nearest(rows, x) {
    let index = -1, distance = Infinity;
    for (let i = 0; i < rows.length; i++) {
      const current = Math.abs(rows[i].x - x);
      if (index < 0 || current < distance) { index = i; distance = current; }
    }
    return index;
  }
  const TIME_MODES = ['auto', 'excel', 'date', 'number', 'text', 'unix-s', 'unix-ms'];
  function validateConfig(value) {
    if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('配置必须是 JSON 对象。');
    if (value.schema !== 'curve-viewer' || value.version !== 1) throw new Error('不支持的配置格式，需要 curve-viewer v1 配置。');
    if (typeof value.name !== 'string' || !value.name.trim() || value.name.length > 120) throw new Error('配置名称应为 1～120 个字符。');
    if (!Array.isArray(value.selectedSeries) || value.selectedSeries.length > 512) throw new Error('selectedSeries 必须是最多 512 项的数组。');
    const timeMode = value.timeMode || defaults?.axes.timeMode || TIME_MODES[0];
    if (!TIME_MODES.includes(timeMode)) throw new Error('配置中的时间格式无效。');
    const keys = new Set();
    const selectedSeries = value.selectedSeries.map(s => {
      if (!s || typeof s.name !== 'string' || !s.name || s.name.length > 1000 || !Number.isInteger(s.occurrence) || s.occurrence < 1 || s.occurrence > 512) throw new Error('配置中的序列名称或序号无效。');
      const key = keyOf(s.name, s.occurrence);
      if (keys.has(key)) throw new Error('配置中存在重复序列。'); keys.add(key);
      return { name: s.name, occurrence: s.occurrence, ...(validColor(s.color) ? { color: s.color } : {}) };
    });
    const axisLimits = validateAxisLimits(value.axisLimits);
    const axisRegions = validateAxisRegions(value.axisRegions);
    const assignments = value.seriesAxes === undefined ? [] : value.seriesAxes;
    if (!Array.isArray(assignments) || assignments.length > 512) throw new Error('seriesAxes 必须是最多 512 项的数组。');
    const assignmentKeys = new Set();
    const seriesAxes = assignments.map(entry => {
      if (!entry || typeof entry.name !== 'string' || !entry.name || entry.name.length > 1000 || !Number.isInteger(entry.occurrence) || entry.occurrence < 1 || entry.occurrence > 512 || !AXIS_IDS.includes(entry.axisId)) throw new Error('配置中的曲线坐标分配无效。');
      const key = keyOf(entry.name, entry.occurrence);
      if (assignmentKeys.has(key)) throw new Error('配置中存在重复的曲线坐标分配。'); assignmentKeys.add(key);
      return { name: entry.name, occurrence: entry.occurrence, axisId: entry.axisId };
    });
    return { schema: 'curve-viewer', version: 1, name: value.name.trim(), timeMode, selectedSeries,
      ...(Object.keys(axisLimits).length ? { axisLimits } : {}), ...(Object.keys(axisRegions).length ? { axisRegions } : {}), ...(seriesAxes.length ? { seriesAxes } : {}) };
  }
  const CONFIG_TYPES = ['axes', 'regex', 'selection'];
  function validateProfile(value) {
    if (!value || value.schema !== 'curve-viewer' || value.version !== 2 || !CONFIG_TYPES.includes(value.type)) throw new Error('需要 curve-viewer v2 分类配置。');
    if (typeof value.name !== 'string' || !value.name.trim() || value.name.length > 120) throw new Error('配置名称应为 1～120 个字符。');
    const profile = { schema: 'curve-viewer', version: 2, type: value.type, name: value.name.trim() };
    const fields = { axes: ['timeMode', 'axisLimits', 'axisRegions', 'seriesAxes'], regex: ['headerRule'], selection: ['selectedSeries'] };
    for (const type of CONFIG_TYPES) if (type !== value.type) for (const field of fields[type]) {
      if (Object.hasOwn(value, field)) throw new Error(`分类配置不能混入其他类型的字段：${field}`);
    }
    if (value.type === 'regex') {
      if (!value.headerRule) throw new Error('缺少标题正则规则。');
      return { ...profile, headerRule: validateHeaderRule(value.headerRule) };
    }
    const validated = validateConfig({ ...value, version: 1, selectedSeries: value.type === 'selection' ? value.selectedSeries : [] });
    function identities(entries, originals) {
      return entries.map((entry, index) => {
        const original = originals[index];
        if (original.sourceName === undefined && original.sourceOccurrence === undefined) return entry;
        if (typeof original.sourceName !== 'string' || !original.sourceName || original.sourceName.length > 1000 || !Number.isInteger(original.sourceOccurrence) || original.sourceOccurrence < 1 || original.sourceOccurrence > 512) throw new Error('配置中的原始列名或序号无效。');
        return { ...entry, sourceName: original.sourceName, sourceOccurrence: original.sourceOccurrence };
      });
    }
    if (value.type === 'selection') return { ...profile, selectedSeries: identities(validated.selectedSeries, value.selectedSeries) };
    return { ...profile, timeMode: validated.timeMode, axisLimits: validated.axisLimits || {}, axisRegions: validated.axisRegions || {}, seriesAxes: identities(validated.seriesAxes || [], value.seriesAxes || []) };
  }
  function splitConfig(value) {
    if (value?.version === 2) return [validateProfile(value)];
    const legacy = validateConfig(value), base = { schema: 'curve-viewer', version: 2, name: legacy.name };
    return [validateProfile({ ...base, type: 'axes', timeMode: legacy.timeMode, axisLimits: legacy.axisLimits || {}, axisRegions: legacy.axisRegions || {}, seriesAxes: legacy.seriesAxes || [] }), validateProfile({ ...base, type: 'selection', selectedSeries: legacy.selectedSeries })];
  }
  function validColor(value) {
    if (typeof value !== 'string') return false;
    if (/^#[0-9a-f]{6}$/i.test(value)) return true;
    const match = /^hsl\((\d+(?:\.\d+)?), (\d+(?:\.\d+)?)%, (\d+(?:\.\d+)?)%\)$/.exec(value);
    return !!match && +match[1] <= 360 && +match[2] <= 100 && +match[3] <= 100;
  }
  function validateDefaults(type, value) {
    if (!CONFIG_TYPES.includes(type) || !value || typeof value !== 'object' || Array.isArray(value)) throw new Error(`${type}.json 缺少有效的 defaults。`);
    if (type === 'regex') return { headerRule: validateHeaderRule(value.headerRule || null) };
    if (type === 'selection') {
      const selection = validateProfile({ schema: 'curve-viewer', version: 2, type, name: 'defaults', selectedSeries: value.selectedSeries });
      if (!Array.isArray(value.palette) || !value.palette.length || value.palette.length > 512 || !value.palette.every(validColor)) throw new Error('selection.json 的默认配色 palette 无效。');
      const color = value.fallbackColor;
      if (!color || !Number.isFinite(color.hueStep) || color.hueStep <= 0 || color.hueStep > 360 || !Number.isFinite(color.saturation) || color.saturation < 0 || color.saturation > 100 || !Number.isFinite(color.lightness) || color.lightness < 0 || color.lightness > 100) throw new Error('selection.json 的 fallbackColor 无效。');
      return { selectedSeries: selection.selectedSeries, palette: [...value.palette], fallbackColor: { hueStep: color.hueStep, saturation: color.saturation, lightness: color.lightness } };
    }
    if (!TIME_MODES.includes(value.timeMode)) throw new Error('axes.json 的默认时间格式无效。');
    const axes = validateProfile({ schema: 'curve-viewer', version: 2, type, name: 'defaults', timeMode: value.timeMode, axisLimits: value.axisLimits, axisRegions: value.axisRegions, seriesAxes: value.seriesAxes });
    if (!AXIS_IDS.every(id => axes.axisRegions[id])) throw new Error('axes.json 的 defaults.axisRegions 必须包含所有坐标组。');
    if (!Array.isArray(value.axisGroups) || value.axisGroups.length !== AXIS_IDS.length) throw new Error('axes.json 的 defaults.axisGroups 必须包含所有坐标组。');
    const seen = new Set(), groups = value.axisGroups.map(group => {
      if (!group || !AXIS_IDS.includes(group.id) || seen.has(group.id) || typeof group.label !== 'string' || !group.label.trim() || group.label.length > 120 || !validColor(group.color)) throw new Error('默认坐标组的标识、名称或颜色无效。');
      seen.add(group.id);
      if (group.id !== 'auto' && (!Number.isFinite(group.min) || !Number.isFinite(group.max) || group.min >= group.max)) throw new Error('默认坐标组的范围无效。');
      return { id: group.id, ...(group.id === 'auto' ? {} : { min: group.min, max: group.max }), label: group.label, color: group.color };
    });
    const bounds = value.autoBounds;
    if (!bounds || !Number.isFinite(bounds.paddingPercent) || bounds.paddingPercent < 0 || bounds.paddingPercent > 100 || !Number.isFinite(bounds.minimumConstantPadding) || bounds.minimumConstantPadding <= 0 || typeof bounds.roundOutward !== 'boolean') throw new Error('axes.json 的默认自动上下限规则无效。');
    return { timeMode: axes.timeMode, axisLimits: axes.axisLimits, axisRegions: axes.axisRegions, seriesAxes: axes.seriesAxes, axisGroups: groups,
      autoBounds: { paddingPercent: bounds.paddingPercent, minimumConstantPadding: bounds.minimumConstantPadding, roundOutward: bounds.roundOutward } };
  }
  function configureDefaults(value) {
    const next = Object.fromEntries(CONFIG_TYPES.map(type => [type, validateDefaults(type, value?.[type])]));
    const groups = next.axes.axisGroups.map(group => ({ ...group, band: [next.axes.axisRegions[group.id].start / 100, next.axes.axisRegions[group.id].end / 100] }));
    defaults = next; DEFAULT_HEADER_RULE = Object.freeze({ ...next.regex.headerRule });
    AXIS_GROUPS.splice(0, AXIS_GROUPS.length, ...groups); RANGES.splice(0, RANGES.length, ...groups.filter(g => g.id !== 'auto')); PALETTE.splice(0, PALETTE.length, ...next.selection.palette);
  }
  return { RANGES, AXIS_GROUPS, PALETTE, TIME_MODES, CONFIG_TYPES, get DEFAULT_HEADER_RULE() { return DEFAULT_HEADER_RULE; }, get defaultsReady() { return !!defaults; }, requireDefaults, getDefaults, configureDefaults, validateDefaults, validateHeaderRule, createHeaderFormatter, renameHeaders, validateProfile, splitConfig, numeric, rangeFor, keyOf, normalizeHeader, colorAt, parseTextTime, readSheet, axesFor, axisIdFor, axisRegionFor, axisRegionLabel, nearest, formatNumber, formatTime, validateConfig, validateAxisLimits, validateAxisRegions };
});
