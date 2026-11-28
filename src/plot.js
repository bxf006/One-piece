/* Canvas time-series renderer, kept local with no remote dependencies. */
(function () {
  'use strict';
  class Plot {
    constructor(element) {
      this.element = element; this.handlers = {}; this.zrHandlers = {};
      this.canvas = document.createElement('canvas'); this.canvas.setAttribute('aria-hidden', 'true');
      this.element.style.position = 'relative'; this.element.append(this.canvas); this.ctx = this.canvas.getContext('2d');
      this.controls = document.createElement('div'); this.controls.className = 'zoom-controls'; this.controls.hidden = true;
      this.startInput = document.createElement('input'); this.endInput = document.createElement('input');
      [this.startInput, this.endInput].forEach((input, i) => { input.type = 'range'; input.min = 0; input.max = 100; input.step = .1; input.value = i ? 100 : 0; input.setAttribute('aria-label', i ? '结束时间范围' : '开始时间范围'); });
      const startLabel = document.createElement('label'), endLabel = document.createElement('label'); startLabel.append('起点', this.startInput); endLabel.append('终点', this.endInput); this.controls.append(startLabel, endLabel); this.element.append(this.controls);
      this.startInput.oninput = () => this.zoom(Math.min(+this.startInput.value, this.end - .1), this.end);
      this.endInput.oninput = () => this.zoom(this.start, Math.max(+this.endInput.value, this.start + .1));
      this.canvas.addEventListener('mousemove', e => {
        if (this.drag && e.buttons === 1) {
          const delta = (e.offsetX - this.drag.x) / this.plotWidth() * (this.drag.end - this.drag.start);
          let start = this.drag.start - delta, end = this.drag.end - delta;
          if (start < 0) { end -= start; start = 0; } if (end > 100) { start -= end - 100; end = 100; }
          this.zoom(start, end);
        }
        this.zrHandlers.mousemove?.({ offsetX: e.offsetX, offsetY: e.offsetY });
      });
      this.canvas.addEventListener('mousedown', e => { if (e.button === 0 && !e.ctrlKey && this.containPixel({}, [e.offsetX, e.offsetY])) { this.drag = { x: e.offsetX, start: this.start, end: this.end }; this.canvas.style.cursor = 'grabbing'; } });
      const endDrag = () => { this.drag = null; this.canvas.style.cursor = 'crosshair'; };
      window.addEventListener('mouseup', endDrag); this.canvas.addEventListener('mouseleave', endDrag);
      this.canvas.addEventListener('wheel', e => {
        if (!e.ctrlKey || !this.containPixel({}, [e.offsetX, e.offsetY])) return;
        e.preventDefault();
        const ratio = Math.max(0, Math.min(1, (e.offsetX - this.option.grid.left) / this.plotWidth())), span = this.end - this.start;
        const newSpan = Math.max(.1, Math.min(100, span * (e.deltaY > 0 ? 1.15 : .87))), center = this.start + ratio * span;
        let start = Math.max(0, center - ratio * newSpan), end = Math.min(100, start + newSpan); start = Math.max(0, end - newSpan); this.zoom(start, end);
      }, { passive: false });
      this.resize();
    }
    getZr() { return { on: (name, callback) => { this.zrHandlers[name] = callback; } }; }
    on(name, callback) { this.handlers[name] = callback; }
    getWidth() { return this.width; } getHeight() { return this.height; }
    plotWidth() { return this.width - this.option.grid.left - this.option.grid.right; }
    resize() {
      this.width = this.element.clientWidth; this.height = this.element.clientHeight;
      const dpr = Math.min(window.devicePixelRatio || 1, 3); this.canvas.width = Math.round(this.width * dpr); this.canvas.height = Math.round(this.height * dpr);
      this.canvas.style.width = `${this.width}px`; this.canvas.style.height = `${this.height}px`; this.ctx.setTransform(dpr, 0, 0, dpr, 0, 0); if (this.option) this.draw();
    }
    clear() {
      this.option = null; this.drag = null; this.controls.hidden = true;
      this.ctx.clearRect(0, 0, this.width, this.height);
    }
    setOption(option) {
      this.controls.hidden = false;
      this.option = option; this.start = option.dataZoom[0].start; this.end = option.dataZoom[0].end;
      // Domain is supplied independently from selected series so empty selections still retain time.
      this.min = option.domain[0]; this.max = option.domain[1];
      if (this.min === this.max) { this.min -= .5; this.max += .5; }
      this.syncInputs(); this.draw();
    }
    syncInputs() { this.startInput.value = this.start; this.endInput.value = this.end; }
    zoom(start, end) { this.start = Math.max(0, Math.min(99.9, start)); this.end = Math.min(100, Math.max(this.start + .1, end)); this.syncInputs(); this.draw(); this.handlers.datazoom?.({ start: this.start, end: this.end }); }
    dispatchAction(action) { if (this.option && action.type === 'dataZoom') this.zoom(action.start, action.end); }
    domain() { const span = this.max - this.min; return [this.min + this.start / 100 * span, this.min + this.end / 100 * span]; }
    convertFromPixel(_, x) { const [lo, hi] = this.domain(); return lo + (x - this.option.grid.left) / this.plotWidth() * (hi - lo); }
    convertToPixel(_, x) { const [lo, hi] = this.domain(); return this.option.grid.left + (x - lo) / (hi - lo) * this.plotWidth(); }
    containPixel(_, [x, y]) { if (!this.option) return false; const g = this.option.grid; return x >= g.left && x <= this.width - g.right && y >= g.top && y <= this.height - g.bottom; }
    axisBounds(axis) {
      const g = this.option.grid, fullHeight = this.height - g.top - g.bottom, band = axis.band || [0, 1];
      const height = fullHeight * (band[1] - band[0]);
      return { top: g.top + fullHeight * band[0], bottom: g.top + fullHeight * band[1], height, ticks: Math.max(1, Math.min(5, Math.floor(height / 24))) };
    }
    draw() {
      const o = this.option, ctx = this.ctx, g = o.grid, top = g.top, bottom = this.height - g.bottom, right = this.width - g.right, height = bottom - top;
      ctx.clearRect(0, 0, this.width, this.height); ctx.lineWidth = 1; ctx.font = '11px "Segoe UI", "Microsoft YaHei", sans-serif';
      const line = (x1, y1, x2, y2, color, dashed = false) => { ctx.strokeStyle = color; ctx.setLineDash(dashed ? [3, 4] : []); ctx.beginPath(); ctx.moveTo(x1, y1); ctx.lineTo(x2, y2); ctx.stroke(); };
      const grids = new Set();
      for (const axis of o.yAxis) {
        const area = this.axisBounds(axis), key = `${area.top}:${area.bottom}`;
        if (grids.has(key)) continue; grids.add(key);
        for (let i = 0; i <= area.ticks; i++) line(g.left, area.top + area.height * i / area.ticks, right, area.top + area.height * i / area.ticks, '#e9edf5', true);
      }
      o.yAxis.forEach((axis, index) => {
        const area = this.axisBounds(axis);
        const isLeft = index % 2 === 0, x = isLeft ? g.left - axis.offset : right + axis.offset;
        line(x, area.top, x, area.bottom, '#e0e5ef'); ctx.fillStyle = axis.axisLabel.color;
        ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillText(axis.name, x, area.top - 15);
        ctx.textAlign = isLeft ? 'right' : 'left'; ctx.textBaseline = 'middle';
        for (let tick = 0; tick <= area.ticks; tick++) { const ratio = tick / area.ticks, value = axis.min * (1 - ratio) + axis.max * ratio; ctx.fillText(CurveCore.formatNumber(value), x + (isLeft ? -9 : 9), area.bottom - area.height * ratio); }
      });
      const [lo, hi] = this.domain(), xCount = Math.max(2, Math.floor(this.plotWidth() / (o.xAxis.type === 'time' && hi - lo > 86400000 ? 170 : 100)));
      ctx.textBaseline = 'top'; ctx.textAlign = 'center'; ctx.fillStyle = '#929caf';
      for (let i = 0; i <= xCount; i++) {
        let value = lo + (hi - lo) * i / xCount;
        if (o.xAxis.type === 'category') value = Math.round(value);
        const x = this.convertToPixel({}, value); line(x, top, x, bottom, '#f0f2f7', true);
        const label = o.xAxis.type === 'category' ? o.xAxis.data[value] ?? '' : o.xAxis.type === 'time' ? o.xAxis.axisLabel.formatter(value) : CurveCore.formatNumber(value);
        ctx.fillText(label.length > 25 ? label.slice(0, 24) + '…' : label, x, bottom + 12, 160);
      }
      line(g.left, bottom, right, bottom, '#dfe5ef');
      ctx.save(); ctx.beginPath(); ctx.rect(g.left, top - 2, this.plotWidth(), height + 4); ctx.clip();
      for (const series of o.series) {
        const axis = o.yAxis[series.yAxisIndex], area = this.axisBounds(axis);
        ctx.save(); ctx.beginPath(); ctx.rect(g.left, area.top, this.plotWidth(), area.height); ctx.clip();
        ctx.strokeStyle = series.lineStyle.color; ctx.fillStyle = series.lineStyle.color; ctx.lineWidth = series.lineStyle.width; ctx.setLineDash([]); ctx.beginPath();
        let active = false;
        for (let i = 0; i < series.data.length; i++) {
          const [x, y] = series.data[i];
          if (y === null || !Number.isFinite(y)) { active = false; continue; }
          // Rows retain source order; later points may return into the visible interval.
          const px = this.convertToPixel({}, x), py = area.bottom - (y - axis.min) / (axis.max - axis.min) * area.height;
          if (active) ctx.lineTo(px, py); else ctx.moveTo(px, py); active = true;
          const isolated = (i === 0 || series.data[i - 1][1] === null) && (i === series.data.length - 1 || series.data[i + 1][1] === null);
          if (isolated) { ctx.fillRect(px - 2, py - 2, 4, 4); }
        }
        ctx.stroke(); ctx.restore();
      }
      ctx.restore();
      // A miniature time overview and the currently visible interval.
      const previewY = this.height - 39, previewH = 12;
      ctx.fillStyle = '#f1f4fb'; ctx.fillRect(g.left, previewY, this.plotWidth(), previewH);
      ctx.fillStyle = '#5465e82a'; ctx.fillRect(g.left + this.start / 100 * this.plotWidth(), previewY, (this.end - this.start) / 100 * this.plotWidth(), previewH);
      this.controls.style.left = `${g.left}px`; this.controls.style.right = `${g.right}px`;
    }
  }
  window.CurvePlot = { init: element => new Plot(element) };
})();
