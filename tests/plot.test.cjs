const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const C = require('../src/core.js');

function setupPlot() {
  const calls = [];
  const windowEvents = {};
  const ctx = new Proxy({}, { get: (_, name) => (...args) => calls.push({ name, args }), set: () => true });
  class Element {
    constructor() { this.style = {}; this.children = []; this.clientWidth = 800; this.clientHeight = 500; this.events = {}; }
    append(...children) { this.children.push(...children); }
    setAttribute() {}
    addEventListener(name, callback) { this.events[name] = callback; }
    getContext() { return ctx; }
  }
  const context = { CurveCore: C, document: { createElement: () => new Element() }, addEventListener(name, callback) { windowEvents[name] = callback; } };
  context.window = context;
  vm.runInNewContext(fs.readFileSync(require.resolve('../src/plot.js'), 'utf8'), context);
  const chart = context.CurvePlot.init(new Element());
  const option = {
    grid: { left: 60, right: 30, top: 50, bottom: 80 }, domain: [0, 10],
    xAxis: { type: 'value' }, yAxis: [{ min: -0.5, max: 1.5, name: '手动', offset: 0, axisLabel: { color: '#000' } }],
    dataZoom: [{ start: 0, end: 100 }],
    series: [{ yAxisIndex: 0, lineStyle: { color: '#000', width: 1 }, data: [[0, -0.5], [5, 0.5], [10, 1.5]] }]
  };
  return { calls, chart, option, windowEvents };
}

test('canvas uses exact manual bounds for positions and retains very small tick labels', () => {
  const { calls, chart, option } = setupPlot();
  chart.setOption(option);
  const clipAt = calls.findIndex(call => call.name === 'clip');
  const curve = calls.slice(clipAt + 1).filter(call => call.name === 'moveTo' || call.name === 'lineTo');
  assert.deepEqual(curve.map(call => call.args), [[60, 420], [415, 235], [770, 50]]);
  calls.length = 0;
  chart.setOption({ ...option, series: [], yAxis: [{ ...option.yAxis[0], min: 0, max: 1e-13 }] });
  const labels = calls.filter(call => call.name === 'fillText' && call.args[1] === 51).map(call => call.args[0]);
  assert.equal(labels.at(-1), C.formatNumber(1e-13));
  assert.equal(new Set(labels).size, 6);
});

test('zoomed canvas keeps source-order segments that return from outside the visible time range', () => {
  const { calls, chart, option } = setupPlot();
  const points = [[8, 0], [9, 1], [5, 2], [1, 3], [0, 4], [6, 5]];
  chart.setOption({ ...option, dataZoom: [{ start: 40, end: 60 }], yAxis: [{ ...option.yAxis[0], min: 0, max: 5 }], series: [{ ...option.series[0], data: points }] });
  const clipAt = calls.findIndex(call => call.name === 'clip');
  const curve = calls.slice(clipAt + 1).filter(call => call.name === 'moveTo' || call.name === 'lineTo');
  assert.deepEqual(curve.map(call => call.name), ['moveTo', 'lineTo', 'lineTo', 'lineTo', 'lineTo', 'lineTo']);
  assert.deepEqual(curve.map(call => call.args[0]), points.map(([x]) => chart.convertToPixel({}, x)));
  assert.equal(curve[2].args[0], 415); assert.equal(curve[5].args[0], 770);
});

test('each coordinate band maps minima, midpoints and maxima to its own height with readable ticks', () => {
  for (const band of [[0, 0.25], [0.25, 0.8], [0.8, 1], [0, 1], [0.105, 0.605]]) {
    const { calls, chart, option } = setupPlot();
    chart.setOption({ ...option, yAxis: [{ ...option.yAxis[0], band }] });
    const top = 50 + 370 * band[0], bottom = 50 + 370 * band[1], height = 370 * (band[1] - band[0]);
    const clipAt = calls.findIndex(call => call.name === 'clip');
    const curve = calls.slice(clipAt + 1).filter(call => call.name === 'moveTo' || call.name === 'lineTo');
    const expectedY = [bottom, bottom - height / 2, bottom - height];
    curve.forEach((call, i) => assert.ok(Math.abs(call.args[1] - expectedY[i]) < 1e-9));
    const rectangles = calls.filter(call => call.name === 'rect');
    assert.deepEqual(rectangles[1].args, [60, top, 710, height]);
    const ticks = calls.filter(call => call.name === 'fillText' && call.args[1] === 51);
    assert.equal(ticks[0].args[0], '-0.5'); assert.equal(ticks.at(-1).args[0], '1.5');
    for (let i = 1; i < ticks.length; i++) assert.ok(ticks[i - 1].args[2] - ticks[i].args[2] >= 24 - 1e-9);
  }
});

test('manual bounds clip overflowing curves to their own band and resize preserves proportions', () => {
  const { calls, chart, option } = setupPlot();
  const band = [0, 0.25];
  chart.setOption({ ...option, yAxis: [{ ...option.yAxis[0], min: 0, max: 1, band }],
    series: [{ ...option.series[0], data: [[0, -1], [5, 0.5], [10, 2]] }] });
  const area = chart.axisBounds(chart.option.yAxis[0]);
  const clipAt = calls.findIndex(call => call.name === 'clip');
  const curve = calls.slice(clipAt + 1).filter(call => call.name === 'moveTo' || call.name === 'lineTo');
  assert.ok(curve[0].args[1] > area.bottom); assert.ok(curve[2].args[1] < area.top);
  const seriesClip = calls.findIndex((call, i) => i > clipAt && call.name === 'clip');
  assert.ok(seriesClip > clipAt && seriesClip < calls.indexOf(curve[0]));
  assert.deepEqual(calls[seriesClip - 1].args, [60, area.top, 710, area.height]);
  chart.element.clientHeight = 900; chart.resize();
  const resized = chart.axisBounds(chart.option.yAxis[0]);
  assert.equal(resized.top, 50); assert.equal(resized.bottom, 50 + (900 - 50 - 80) * 0.25);
});

test('Ctrl-left does not pan the canvas; ordinary dragging and Ctrl-wheel zoom still work', () => {
  const { chart, option, windowEvents } = setupPlot();
  chart.setOption({ ...option, dataZoom: [{ start: 20, end: 80 }] });
  chart.canvas.events.mousedown({ button: 0, ctrlKey: true, offsetX: 300, offsetY: 200 });
  chart.canvas.events.mousemove({ buttons: 1, offsetX: 400, offsetY: 200 });
  assert.equal(chart.drag, undefined); assert.equal(chart.start, 20); assert.equal(chart.end, 80);
  chart.canvas.events.mousedown({ button: 0, ctrlKey: false, offsetX: 300, offsetY: 200 });
  chart.canvas.events.mousemove({ buttons: 1, offsetX: 400, offsetY: 200 });
  assert.ok(chart.start < 20); assert.ok(chart.end < 80);
  windowEvents.mouseup(); assert.equal(chart.drag, null);
  let prevented = false;
  const span = chart.end - chart.start;
  chart.canvas.events.wheel({ ctrlKey: true, offsetX: 400, offsetY: 200, deltaY: -1, preventDefault() { prevented = true; } });
  assert.equal(prevented, true); assert.ok(chart.end - chart.start < span);
});
