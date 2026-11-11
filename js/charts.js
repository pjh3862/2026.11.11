/* 작은 SVG 그래프 — 외부 라이브러리 없이 막대(묶음)·꺾은선 두 가지만 (data09-25 의 charts.js 를 고쳐 씀).
   · 음수도 그립니다(예상재고가 0 아래로 내려가는 날). 0 선과 기준선(안전재고·가용 인시)을 따로 긋습니다.
   · 막대·점에 마우스를 올리면 값이 보입니다(title). 색은 CSS 변수(--series-* …)를 씁니다.
   · 가로 글자가 많으면 몇 칸 건너 하나씩만 적어 좁은 화면에서도 겹치지 않게 합니다. */
(function (root) {
  'use strict';
  var NS = 'http://www.w3.org/2000/svg';
  function el(tag, attrs, text) {
    var e = document.createElementNS(NS, tag);
    Object.keys(attrs || {}).forEach(function (k) { e.setAttribute(k, attrs[k]); });
    if (text != null) e.textContent = text;
    return e;
  }
  function nice(v) {
    if (!(v > 0)) return 1;
    var p = Math.pow(10, Math.floor(Math.log10(v))), n = v / p;
    return (n <= 1 ? 1 : n <= 2 ? 2 : n <= 2.5 ? 2.5 : n <= 5 ? 5 : 10) * p;
  }
  function fmt(v, unit) {
    if (v == null) return '-';
    if (unit === '%') return (Math.round(v * 1000) / 10) + '%';
    return (Math.round(v * 10) / 10).toLocaleString('ko-KR') + (unit || '');
  }

  /* opts: { labels:[], series:[{name, values:[], color}], unit, type:'bar'|'line', height, refs:[{value, name, color}] } */
  function chart(opts) {
    var W = 640, H = opts.height || 220, padL = 52, padR = 12, padT = 14, padB = 30;
    var n = opts.labels.length, unit = opts.unit || '';
    var all = [];
    opts.series.forEach(function (s) { s.values.forEach(function (v) { if (v != null && isFinite(v)) all.push(v); }); });
    (opts.refs || []).forEach(function (r) { if (r.value != null && isFinite(r.value)) all.push(r.value); });
    var hi = Math.max.apply(null, all.concat([0])), lo = Math.min.apply(null, all.concat([0]));
    var span = nice(Math.max(hi - lo, 1) / 4) * 4;
    var step0 = span / 4;
    var max = Math.ceil(hi / step0) * step0, min = Math.floor(lo / step0) * step0;
    if (max === min) max = min + step0 * 4;
    var svg = el('svg', { viewBox: '0 0 ' + W + ' ' + H, class: 'chart', role: 'img', 'aria-label': opts.title || '' });
    var plotW = W - padL - padR, plotH = H - padT - padB;
    function y(v) { return padT + plotH - ((v - min) / (max - min)) * plotH; }
    for (var gv = min; gv <= max + 1e-9; gv += step0) {
      var gy = y(gv);
      svg.appendChild(el('line', { x1: padL, x2: W - padR, y1: gy, y2: gy, class: Math.abs(gv) < 1e-9 ? 'axis' : 'grid' }));
      svg.appendChild(el('text', { x: padL - 6, y: gy + 4, 'text-anchor': 'end', class: 'tick' }, fmt(gv, unit === '%' ? '%' : '')));
    }
    var step = plotW / Math.max(n, 1);
    var every = Math.max(1, Math.ceil(n / 10));
    opts.labels.forEach(function (lb, i) {
      if (i % every) return;
      svg.appendChild(el('text', { x: padL + step * (i + 0.5), y: H - 10, 'text-anchor': 'middle', class: 'tick' }, lb));
    });
    if (!all.length) {
      svg.appendChild(el('text', { x: W / 2, y: H / 2, 'text-anchor': 'middle', class: 'tick' }, '표시할 값이 없습니다'));
      return svg;
    }
    if (opts.type === 'line') {
      opts.series.forEach(function (s) {
        var d = '', pen = false;
        s.values.forEach(function (v, i) {
          if (v == null || !isFinite(v)) { pen = false; return; }
          d += (pen ? 'L' : 'M') + (padL + step * (i + 0.5)).toFixed(1) + ' ' + y(v).toFixed(1) + ' ';
          pen = true;
        });
        svg.appendChild(el('path', { d: d, class: 'ln', style: 'stroke:' + s.color }));
        s.values.forEach(function (v, i) {
          if (v == null || !isFinite(v)) return;
          var c = el('circle', { cx: padL + step * (i + 0.5), cy: y(v), r: n > 20 ? 3 : 4.5, class: 'dot', style: 'fill:' + (v < 0 && s.negColor ? s.negColor : s.color) });
          c.appendChild(el('title', null, opts.labels[i] + ' · ' + s.name + ' ' + fmt(v, unit)));
          svg.appendChild(c);
        });
      });
    } else {
      var k = opts.series.length, gap = 2, bw = Math.min(30, (step * 0.7 - gap * (k - 1)) / k);
      opts.labels.forEach(function (lb, i) {
        var x0 = padL + step * (i + 0.5) - (bw * k + gap * (k - 1)) / 2;
        opts.series.forEach(function (s, j) {
          var v = s.values[i];
          if (v == null || !isFinite(v)) return;
          var top = y(Math.max(v, 0)), bot = y(Math.min(v, 0));
          var b = el('rect', { x: x0 + j * (bw + gap), y: top, width: bw, height: Math.max(0, bot - top), rx: 3, class: 'bar', style: 'fill:' + ((s.colors && s.colors[i]) || s.color) });
          b.appendChild(el('title', null, lb + ' · ' + s.name + ' ' + fmt(v, unit)));
          svg.appendChild(b);
        });
      });
    }
    (opts.refs || []).forEach(function (r) {
      if (r.value == null || !isFinite(r.value)) return;
      var ry = y(r.value);
      svg.appendChild(el('line', { x1: padL, x2: W - padR, y1: ry, y2: ry, class: 'ref', style: 'stroke:' + r.color }));
      svg.appendChild(el('text', { x: W - padR - 4, y: ry - 5, 'text-anchor': 'end', class: 'tick ref-t' }, r.name + ' ' + fmt(r.value, unit)));
    });
    return svg;
  }
  function legend(series, refs) {
    var box = document.createElement('div'); box.className = 'legend';
    var all = series.map(function (s) { return { name: s.name, color: s.color }; }).concat((refs || []).map(function (r) { return { name: r.name + '(선)', color: r.color, line: true }; }));
    if (all.length < 2) return box;
    all.forEach(function (s) {
      var i = document.createElement('span'); i.className = 'lg';
      var sw = document.createElement('i'); sw.style.background = s.color; if (s.line) sw.className = 'line'; i.appendChild(sw);
      i.appendChild(document.createTextNode(s.name)); box.appendChild(i);
    });
    return box;
  }
  function figure(title, opts) {
    var f = document.createElement('figure'); f.className = 'fig';
    var cap = document.createElement('figcaption'); cap.textContent = title; f.appendChild(cap);
    opts.title = title;
    f.appendChild(legend(opts.series, opts.refs));
    f.appendChild(chart(opts));
    return f;
  }
  root.PPCharts = { chart: chart, figure: figure };
})(window);
