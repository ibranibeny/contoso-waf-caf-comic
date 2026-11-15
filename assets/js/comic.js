/* Land It Well, Dewi! · motion comic reader + animated diagram engine (no dependencies) */
(function () {
  'use strict';
  var LIW = window.LIW = window.LIW || {};
  document.documentElement.classList.remove('no-js');
  function $(s, r) { return (r || document).querySelector(s); }
  function $$(s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); }
  function el(tag, cls, html) { var e = document.createElement(tag); if (cls) e.className = cls; if (html != null) e.innerHTML = html; return e; }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  var store = {
    get: function (k, d) { try { var v = localStorage.getItem(k); return v === null ? d : v; } catch (e) { return d; } },
    set: function (k, v) { try { localStorage.setItem(k, v); } catch (e) { } }
  };
  LIW.store = store;
  var params = new URLSearchParams(location.search);

  /* ------------------------------------------------------------ motion preference */
  var MKEY = 'liw-motion';
  var mq = window.matchMedia ? window.matchMedia('(prefers-reduced-motion: reduce)') : null;
  function motionPref() { var v = store.get(MKEY, ''); if (v === 'on') return true; if (v === 'off') return false; return !(mq && mq.matches); }
  LIW.motion = motionPref();
  function applyMotion() {
    var h = document.documentElement;
    h.classList.toggle('motion', LIW.motion); h.classList.toggle('still', !LIW.motion);
    $$('[data-motion-label]').forEach(function (b) { b.textContent = LIW.motion ? '✦ Motion: on' : '✧ Motion: off'; b.setAttribute('aria-pressed', LIW.motion ? 'true' : 'false'); });
    document.dispatchEvent(new CustomEvent('liw-motion'));
  }
  LIW.setMotion = function (on, remember) { LIW.motion = !!on; if (remember !== false) store.set(MKEY, on ? 'on' : 'off'); applyMotion(); };
  LIW.toggleMotion = function () { LIW.setMotion(!LIW.motion); var n = $('.motion-notice'); if (n) n.remove(); };
  function motionNotice() {
    if (store.get(MKEY, '') || !(mq && mq.matches)) return;
    var n = el('div', 'motion-notice', '<p><b>Your device asks for reduced motion, so the comic is standing still.</b> This story is animated: moving panels, characters and flying packets in the diagrams.</p>' +
      '<div class="mn-btns"><button class="btn primary" data-mn="on">▶ Play animations</button><button class="btn" data-mn="off">Keep it still</button></div>' +
      '<p class="mn-hint">Your choice is remembered. Change it any time with the Motion button or the <span class="kbd">M</span> key.</p>');
    n.setAttribute('role', 'dialog'); n.setAttribute('aria-label', 'Animation preference');
    document.body.appendChild(n);
    n.addEventListener('click', function (ev) { var b = ev.target.closest('[data-mn]'); if (!b) return; LIW.setMotion(b.getAttribute('data-mn') === 'on'); n.remove(); });
  }
  if (params.get('motion') === 'on') LIW.setMotion(true, true);
  if (params.get('motion') === 'off') LIW.setMotion(false, true);
  applyMotion();
  document.addEventListener('DOMContentLoaded', function () {
    applyMotion(); motionNotice();
    $$('[data-motion-label]').forEach(function (b) { b.addEventListener('click', function () { LIW.toggleMotion(); }); });
  });

  /* ------------------------------------------------------------ animated diagrams */
  /* ------------------------------------------------------------ PowerPoint export (code loads on first use) */
  var pptxLoading = null;
  LIW.exportPptx = function (scope) {
    var meta = $('meta[name="liw-pptx"]'); if (!meta) return;
    var epEl = $('#ep-data'), ep = null;
    if (epEl) { try { ep = JSON.parse(epEl.textContent).id; } catch (e) { } }
    function go() { LIW.pptx.open({ scope: scope || (ep ? 'episode' : 'all'), ep: ep, data: meta.getAttribute('data-comic'), icons: meta.getAttribute('data-icons') }); }
    if (LIW.pptx) { go(); return; }
    if (!pptxLoading) pptxLoading = new Promise(function (res, rej) { var sc = document.createElement('script'); sc.src = meta.getAttribute('content'); sc.onload = res; sc.onerror = function () { pptxLoading = null; rej(); }; document.head.appendChild(sc); });
    pptxLoading.then(go, function () { alert('The PowerPoint exporter could not load. Check your connection and try again.'); });
  };
  document.addEventListener('click', function (ev) {
    var b = ev.target.closest && ev.target.closest('[data-pptx]'); if (!b) return;
    ev.preventDefault(); ev.stopPropagation(); LIW.exportPptx(b.getAttribute('data-pptx'));
  });

  var SVGNS = 'http://www.w3.org/2000/svg', XL = 'http://www.w3.org/1999/xlink';
  var KIND = { req: '#06b6d4', resp: '#f59e0b', block: '#ef4444', policy: '#8b5cf6', data: '#10b981', alert: '#f97316', money: '#ca8a04' };
  LIW.KIND = KIND;
  function sv(tag, attrs, parent) { var e = document.createElementNS(SVGNS, tag); for (var k in attrs) { if (attrs[k] != null) e.setAttribute(k, attrs[k]); } if (parent) parent.appendChild(e); return e; }
  function f1(v) { return Math.round(v * 10) / 10; }
  function iconRef(k) { return k.indexOf('g-') === 0 ? k : 'ic-' + k; }
  function exitPt(N, T) {
    if (!N.hw && !N.ht) return { x: N.x, y: N.y };
    var dx = T.x - N.x, dy = T.y - N.y, pad = 6;
    var tx = dx > 0 ? (N.hw + pad) / dx : (dx < 0 ? -(N.hw + pad) / dx : Infinity);
    var ty = dy > 0 ? (N.hb + pad) / dy : (dy < 0 ? -(N.ht + pad) / dy : Infinity);
    var t = Math.min(tx, ty); if (!(t < 0.95)) t = 0.5;
    return { x: N.x + dx * t, y: N.y + dy * t };
  }
  function roundPath(p, r) {
    var s = 'M' + f1(p[0].x) + ' ' + f1(p[0].y);
    for (var i = 1; i < p.length - 1; i++) {
      var a = p[i - 1], b = p[i], c = p[i + 1];
      var l1 = Math.hypot(b.x - a.x, b.y - a.y) || 1, l2 = Math.hypot(c.x - b.x, c.y - b.y) || 1, rr = Math.min(r, l1 / 2, l2 / 2);
      s += ' L' + f1(b.x - (b.x - a.x) / l1 * rr) + ' ' + f1(b.y - (b.y - a.y) / l1 * rr) + ' Q' + f1(b.x) + ' ' + f1(b.y) + ' ' + f1(b.x + (c.x - b.x) / l2 * rr) + ' ' + f1(b.y + (c.y - b.y) / l2 * rr);
    }
    var z = p[p.length - 1]; return s + ' L' + f1(z.x) + ' ' + f1(z.y);
  }
  var dgUid = 0;
  function Diagram(host, def, opt) {
    this.host = host; this.def = def; this.opt = opt || {}; this.k = -1; this.run = 0; this.playing = false;
    this.speed = this.opt.speed || 1; this.id = 'dg' + (++dgUid);
    this.build();
    var self = this;
    document.addEventListener('liw-motion', function () { if (!self.playing) self.render(self.k, false); });
  }
  Diagram.prototype.build = function () {
    var d = this.def, self = this;
    var root = this.root = el('div', 'dg');
    root.innerHTML = '<div class="dg-head"><span class="kick">Animated architecture · ' + d.steps.length + ' steps</span><h3></h3></div>';
    $('h3', root).textContent = d.title;
    var svg = this.svg = sv('svg', { viewBox: '0 0 ' + d.vb[0] + ' ' + d.vb[1], 'class': 'dg-svg', role: 'img', 'aria-label': d.title + '. The steps are described below the picture.' });
    var defs = sv('defs', {}, svg);
    var flt = sv('filter', { id: this.id + '-glow', x: '-60%', y: '-60%', width: '220%', height: '220%' }, defs);
    sv('feGaussianBlur', { stdDeviation: '3.2', result: 'b' }, flt);
    var fm = sv('feMerge', {}, flt); sv('feMergeNode', { 'in': 'b' }, fm); sv('feMergeNode', { 'in': 'SourceGraphic' }, fm);
    ['def'].concat(Object.keys(KIND)).forEach(function (k) {
      var mk = sv('marker', { id: self.id + '-a-' + k, viewBox: '0 0 10 10', refX: '8', refY: '5', markerWidth: '6.5', markerHeight: '6.5', orient: 'auto-start-reverse' }, defs);
      sv('path', { d: 'M0,0 L10,5 L0,10 z', fill: k === 'def' ? '#94a3b8' : KIND[k] }, mk);
    });
    var gz = sv('g', {}, svg), ge = sv('g', {}, svg), gn = sv('g', {}, svg), gl = sv('g', {}, svg);
    this.gp = sv('g', {}, svg);
    this.zones = {}; this.nodes = {}; this.edges = {};
    d.zones.forEach(function (z) {
      var g = sv('g', { 'class': 'zone ' + (z.cls || '') }, gz);
      sv('rect', { x: z.x, y: z.y, width: z.w, height: z.h, rx: 14 }, g);
      var t = sv('text', { x: z.x + 12, y: z.y + 19 }, g); t.textContent = z.label;
      self.zones[z.id] = g;
    });
    d.nodes.forEach(function (nd) {
      var cls = nd.cls || '', lane = nd.icon.indexOf('lane:') === 0, sm = cls.indexOf('sm') >= 0;
      var g = sv('g', { 'class': 'node ' + cls, transform: 'translate(' + nd.x + ',' + nd.y + ')' }, gn);
      var info = { x: nd.x, y: nd.y, g: g };
      if (lane) {
        var col = nd.icon.slice(5), w = 150, h = 40;
        sv('rect', { 'class': 'box', x: -w / 2, y: -h / 2, width: w, height: h, rx: 18, style: 'stroke:' + col + ';stroke-width:2.5' }, g);
        sv('circle', { cx: -w / 2 + 18, cy: 0, r: 7, fill: col }, g);
        var tt = sv('text', { x: 9, y: 5 }, g); tt.textContent = nd.label;
        info.hw = w / 2; info.ht = h / 2; info.hb = h / 2;
      } else {
        var half = sm ? 26 : 31, ic = sm ? 36 : 44, lines = String(nd.label).split('\n');
        sv('rect', { 'class': 'box', x: -half, y: -half, width: half * 2, height: half * 2, rx: 13 }, g);
        var u = sv('use', { href: '#' + iconRef(nd.icon), x: -ic / 2, y: -ic / 2, width: ic, height: ic }, g);
        u.setAttributeNS(XL, 'xlink:href', '#' + iconRef(nd.icon));
        var top = cls.indexOf('lt') >= 0;
        lines.forEach(function (ln, i) { var t = sv('text', { x: 0, y: top ? -half - 8 - (lines.length - 1 - i) * 14 : half + 16 + i * 14 }, g); t.textContent = ln; });
        info.hw = half; info.ht = top ? half + 6 + lines.length * 14 : half; info.hb = top ? half : half + 6 + lines.length * 14;
      }
      var tl = sv('title', {}, g); tl.textContent = String(nd.label).replace(/\n/g, ' ');
      self.nodes[nd.id] = info;
    });
    d.edges.forEach(function (ed) {
      var A = self.nodes[ed.a], B = typeof ed.b === 'string' ? self.nodes[ed.b] : { x: ed.b[0], y: ed.b[1], hw: 0, ht: 0, hb: 0 };
      if (!A || !B) return;
      var tree = (ed.cls || '').indexOf('tree') >= 0, pts;
      if (tree) {
        var y1 = A.y + A.hb - 2, y2 = B.y - B.ht - 4, mid = Math.round((y1 + y2) / 2);
        pts = [{ x: A.x, y: y1 }, { x: A.x, y: mid }, { x: B.x, y: mid }, { x: B.x, y: y2 }];
      } else {
        pts = [{ x: A.x, y: A.y }].concat((ed.via || []).map(function (p) { return { x: p[0], y: p[1] }; })).concat([{ x: B.x, y: B.y }]);
        pts[0] = exitPt(A, pts[1]); pts[pts.length - 1] = exitPt(B, pts[pts.length - 2]);
      }
      var g = sv('g', { 'class': 'edge ' + (ed.cls || '') }, ge);
      var p = sv('path', { 'class': 'ln', d: roundPath(pts, 14) }, g);
      self.edges[ed.id] = { g: g, p: p, tree: tree, label: ed.label, lp: ed.lp == null ? 0.5 : ed.lp, gl: gl };
      self.mark(ed.id, 'def');
    });
    root.appendChild(svg);
    var cap = this.capEl = el('div', 'dg-cap', '<span class="num">▶</span><p></p>');
    root.appendChild(cap);
    var ctrl = el('div', 'dg-ctrl');
    ctrl.innerHTML = '<button class="btn sm" data-a="prev" aria-label="Previous step">◀</button>' +
      '<button class="btn sm primary" data-a="play">▶ Play the flow</button>' +
      '<button class="btn sm" data-a="next" aria-label="Next step">▶</button>' +
      '<button class="btn sm" data-a="replay" aria-label="Replay from the start">↺ Replay</button><div class="dg-dots"></div>';
    var dots = $('.dg-dots', ctrl);
    d.steps.forEach(function (s, i) { var b = el('button', '', String(i + 1)); b.setAttribute('aria-label', 'Step ' + (i + 1)); b.dataset.k = i; dots.appendChild(b); });
    root.appendChild(ctrl);
    if (d.note) root.appendChild(el('p', 'dg-note', '✎ ' + esc(d.note)));
    var leg = el('div', 'legend', '<span><i style="background:' + KIND.req + '"></i>request</span><span><i style="background:' + KIND.resp + '"></i>response</span><span><i style="background:' + KIND.data + '"></i>data / ops</span><span><i style="background:' + KIND.policy + '"></i>policy</span><span><i style="background:' + KIND.alert + '"></i>signal / alert</span><span><i style="background:' + KIND.block + '"></i>blocked</span>');
    root.appendChild(leg);
    ctrl.addEventListener('click', function (ev) {
      var b = ev.target.closest('button'); if (!b) return; ev.stopPropagation();
      if (b.dataset.k != null) { self.pause(); self.render(+b.dataset.k, true); return; }
      var a = b.dataset.a;
      if (a === 'play') { if (self.playing) self.pause(); else self.playAll(); }
      if (a === 'next') { self.pause(); if (self.k < d.steps.length - 1) self.render(self.k + 1, true); }
      if (a === 'prev') { self.pause(); if (self.k > 0) self.render(self.k - 1, true); else self.render(-1, false); }
      if (a === 'replay') { self.pause(); self.render(-1, false); self.playAll(); }
    });
    this.host.innerHTML = ''; this.host.appendChild(root);
    // edge labels need layout
    Object.keys(this.edges).forEach(function (id) {
      var E = self.edges[id]; if (!E.label) return;
      var L = E.p.getTotalLength(), pt = E.p.getPointAtLength(L * E.lp);
      var g = sv('g', { 'class': 'elabel' }, E.gl), t = sv('text', { x: pt.x, y: pt.y + 4, 'text-anchor': 'middle' }, g); t.textContent = E.label;
      var bb = t.getBBox(); sv('rect', { x: bb.x - 5, y: bb.y - 2, width: bb.width + 10, height: bb.height + 4, rx: 6 }, g).parentNode.insertBefore(g.lastChild, t);
    });
    this.render(-1, false);
  };
  Diagram.prototype.mark = function (id, k) { var E = this.edges[id]; if (!E || E.tree) return; E.p.setAttribute('marker-end', 'url(#' + this.id + '-a-' + k + ')'); };
  Diagram.prototype.state = function (k) {
    var down = {}, ghost = {}, d = this.def;
    d.nodes.forEach(function (n) { if ((n.cls || '').indexOf('ghost') >= 0) ghost[n.id] = 1; });
    for (var i = 0; i <= k && i < d.steps.length; i++) {
      var s = d.steps[i];
      s.down.forEach(function (x) { down[x] = 1; }); s.up.forEach(function (x) { delete down[x]; });
      s.show.forEach(function (x) { delete ghost[x]; }); s.hide.forEach(function (x) { ghost[x] = 1; });
    }
    return { down: down, ghost: ghost };
  };
  Diagram.prototype.light = function (id, kind) {
    var rev = id.charAt(0) === '-'; if (rev) id = id.slice(1);
    var E = this.edges[id]; if (!E) return null;
    E.g.classList.add('lit'); E.g.style.setProperty('--k', KIND[kind] || KIND.req); this.mark(id, kind);
    return E;
  };
  Diagram.prototype.endPoint = function (ids) {
    var id = ids[ids.length - 1], rev = id.charAt(0) === '-'; if (rev) id = id.slice(1);
    var E = this.edges[id]; if (!E) return null; var L = E.p.getTotalLength();
    return E.p.getPointAtLength(rev ? 0 : L);
  };
  Diagram.prototype.xmark = function (pt, animate) {
    if (!pt) return;
    var g = sv('g', {}, this.gp);
    if (animate) {
      var c = sv('circle', { 'class': 'burst', cx: pt.x, cy: pt.y, r: 6 }, g), t0 = null;
      var fr = function (ts) { if (t0 === null) t0 = ts; var u = Math.min(1, (ts - t0) / 600); c.setAttribute('r', 6 + 26 * u); c.setAttribute('opacity', 1 - u); if (u < 1) requestAnimationFrame(fr); };
      requestAnimationFrame(fr);
    }
    var t = sv('text', { 'class': 'xmark', x: pt.x + 10, y: pt.y - 10 }, g); t.textContent = '✖';
  };
  Diagram.prototype.fly = function (ids, kind, block, run, done) {
    var self = this, col = KIND[kind] || KIND.req, i = 0, last = null;
    var pkt = sv('g', { 'class': 'pkt', filter: 'url(#' + this.id + '-glow)' }, this.gp);
    sv('circle', { 'class': 'halo', r: 12, fill: col }, pkt); sv('circle', { r: 6, fill: col, stroke: '#fff', 'stroke-width': 2 }, pkt);
    function next() {
      if (run !== self.run) { pkt.remove(); return; }
      if (i >= ids.length) {
        if (block) { self.xmark(last, true); }
        setTimeout(function () { pkt.remove(); if (run === self.run && done) done(); }, block ? 700 : 150);
        return;
      }
      var id = ids[i++], rev = id.charAt(0) === '-', E = self.light(id, kind);
      if (!E) { next(); return; }
      var L = E.p.getTotalLength(), dur = Math.max(420, L / (230 * self.speed) * 1000), t0 = null;
      function frame(ts) {
        if (run !== self.run) { pkt.remove(); return; }
        if (t0 === null) t0 = ts;
        var u = Math.min(1, (ts - t0) / dur), e = u < 0.5 ? 2 * u * u : -1 + (4 - 2 * u) * u;
        var pt = E.p.getPointAtLength((rev ? 1 - e : e) * L); last = pt;
        pkt.setAttribute('transform', 'translate(' + f1(pt.x) + ',' + f1(pt.y) + ')');
        if (u < 1) requestAnimationFrame(frame); else next();
      }
      requestAnimationFrame(frame);
    }
    next();
  };
  Diagram.prototype.render = function (k, animate, done) {
    var self = this, d = this.def; this.k = k; var run = ++this.run;
    while (this.gp.firstChild) this.gp.removeChild(this.gp.firstChild);
    var st = this.state(k), step = d.steps[k];
    Object.keys(this.nodes).forEach(function (id) {
      var g = self.nodes[id].g; g.classList.toggle('down', !!st.down[id]); g.classList.toggle('ghost', !!st.ghost[id]); g.classList.remove('hl');
    });
    Object.keys(this.zones).forEach(function (id) { self.zones[id].classList.toggle('down', !!st.down[id]); });
    Object.keys(this.edges).forEach(function (id) { var E = self.edges[id]; E.g.classList.remove('lit'); self.mark(id, 'def'); });
    var num = $('.num', this.capEl), cp = $('p', this.capEl);
    $$('.dg-dots button', this.root).forEach(function (b, i) { b.classList.toggle('on', i === k); });
    if (!step) {
      this.root.style.setProperty('--k', '#0ea5e9'); num.textContent = '▶';
      cp.textContent = 'Press "Play the flow" to watch the traffic move hop by hop, or pick a step number.';
      if (done) done(); return;
    }
    var col = KIND[step.k] || KIND.req; this.root.style.setProperty('--k', col);
    num.textContent = 'Step ' + (k + 1) + ' / ' + d.steps.length; cp.textContent = step.c;
    step.hl.forEach(function (id) { var N = self.nodes[id]; if (N) { N.g.classList.add('hl'); N.g.style.setProperty('--k', col); } });
    if (animate && LIW.motion && step.e.length) { this.fly(step.e, step.k, step.block, run, done); }
    else {
      step.e.forEach(function (id) { self.light(id, step.k); });
      if (step.block && step.e.length) this.xmark(this.endPoint(step.e), false);
      if (done) done();
    }
  };
  Diagram.prototype.btn = function () { var b = $('[data-a="play"]', this.root); if (b) b.textContent = this.playing ? '❚❚ Pause' : (this.k >= this.def.steps.length - 1 ? '↺ Play again' : '▶ Play the flow'); };
  Diagram.prototype.playAll = function (cb) {
    var self = this, n = this.def.steps.length; this.playing = true; this.btn();
    var start = this.k >= n - 1 ? 0 : this.k + 1;
    (function go(k) {
      if (!self.playing) return;
      if (k >= n) { self.playing = false; self.btn(); if (cb) cb(); return; }
      self.render(k, true, function () {
        if (!self.playing) return;
        self.timer = setTimeout(function () { go(k + 1); }, (1100 + self.def.steps[k].c.length * 30) / self.speed);
      });
      self.btn();
    })(start);
  };
  Diagram.prototype.pause = function () { this.playing = false; clearTimeout(this.timer); this.btn(); };
  Diagram.prototype.stop = function () { this.pause(); this.run++; };
  LIW.Diagram = Diagram;

  function dgData() { var s = $('#dg-data'); return s ? JSON.parse(s.textContent) : {}; }
  LIW.mountDiagrams = function () {
    var data = dgData();
    $$('[data-dg]').forEach(function (host) {
      var def = data[host.getAttribute('data-dg')]; if (!def) return;
      var dg = new Diagram(host, def, {});
      if ('IntersectionObserver' in window && host.hasAttribute('data-autoplay')) {
        var io = new IntersectionObserver(function (es) { es.forEach(function (e) { if (e.isIntersecting && LIW.motion && !dg.playing && dg.k < 0) dg.playAll(); }); }, { threshold: 0.45 });
        io.observe(host);
      }
    });
  };

  /* ------------------------------------------------------------ the reader */
  var STOPS = { cap: 1, sfx: 1, say: 1, card: 1, diagram: 1, quiz: 1, end: 1 };
  LIW.episode = function () {
    var D = JSON.parse($('#ep-data').textContent), dgs = dgData();
    var beats = D.beats.concat([{ t: 'end' }]), stops = [];
    beats.forEach(function (b, i) { if (STOPS[b.t]) stops.push(i); });
    var A = D.assets, chars = D.chars;
    var stage = $('#stage'), sprites = $('#sprites'), bubble = $('#bubble'), cap = $('#cap'), scene = $('#scene'), sfx = $('#sfx');
    var ov = $('#ov'), board = $('#board'), logList = $('#logList');
    var bgA = $('#bgA'), bgB = $('#bgB'), panel = $('#panel'), flash = $('#flash');
    var cur = -1, auto = false, speed = 1, timer = null, typing = null, dg = null, quizDone = false, curBg = null, curPanel = null, frontBg = bgA;
    var spr = {};
    // preload sprite images
    Object.keys(chars).forEach(function (c) { (chars[c].ex || []).forEach(function (x) { var im = new Image(); im.src = A + chars[c].dir + '/' + c + '-' + x + '.webp'; }); });
    function autoPos(st, who) {
      var used = {}; Object.keys(st.sprites).forEach(function (k) { used[st.sprites[k].pos] = 1; });
      var pref = (chars[who] && chars[who].pref) || ['L', 'C', 'R'];
      for (var i = 0; i < pref.length; i++) if (!used[pref[i]]) return pref[i];
      return pref[0];
    }
    function stateAt(si) {
      var st = { bg: null, panel: null, mode: 'scene', scene: '', sprites: {}, enter: {} };
      for (var j = 0; j <= stops[si]; j++) {
        var b = beats[j];
        if (b.t === 'bg') { st.bg = b.v; st.mode = 'scene'; st.panel = null; st.sprites = {}; }
        else if (b.t === 'panel') { st.panel = b.v; st.mode = 'panel'; }
        else if (b.t === 'scene') st.scene = b.v;
        else if (b.t === 'enter') { var o = st.sprites[b.who]; st.sprites[b.who] = { ex: b.ex || (o && o.ex) || 'base', pos: b.pos || (o && o.pos) || autoPos(st, b.who) }; }
        else if (b.t === 'exit') delete st.sprites[b.who];
        else if (b.t === 'say' && st.mode === 'scene') {
          var s0 = st.sprites[b.who];
          if (!s0) st.sprites[b.who] = { ex: b.ex || 'base', pos: b.pos || autoPos(st, b.who) };
          else if (b.ex) s0.ex = b.ex;
        }
      }
      st.beat = beats[stops[si]];
      return st;
    }
    function sm() { return stage.clientWidth < 700 ? '-sm' : ''; }
    function setBg(key) {
      if (key === curBg) return; curBg = key;
      var back = frontBg === bgA ? bgB : bgA;
      back.style.backgroundImage = key ? 'url(' + A + 'bg/' + key + sm() + '.webp)' : 'none';
      back.classList.add('on', 'kb'); frontBg.classList.remove('on'); frontBg = back;
      stage.classList.toggle('sky', !!(D.sky && D.sky.indexOf(key) >= 0));
      stage.classList.toggle('glow', !!(D.glow && D.glow.indexOf(key) >= 0));
    }
    function setPanel(key) {
      if (key === curPanel) return; curPanel = key;
      if (!key) { panel.classList.remove('on'); return; }
      var img = $('img', panel); img.src = A + 'panels/' + key + sm() + '.webp';
      img.alt = (D.panelAlt && D.panelAlt[key]) || '';
      panel.classList.remove('on'); void panel.offsetWidth; panel.classList.add('on');
      flash.classList.remove('go'); void flash.offsetWidth; flash.classList.add('go');
    }
    function renderSprites(st, speaker, mode) {
      var want = st.mode === 'scene' ? st.sprites : {};
      sprites.classList.toggle('hidden', st.mode !== 'scene');
      Object.keys(spr).forEach(function (k) { if (!want[k]) { var e = spr[k]; e.style.opacity = 0; setTimeout(function () { e.remove(); }, 450); delete spr[k]; } });
      Object.keys(want).forEach(function (k) {
        var w = want[k], c = chars[k] || {}, e = spr[k], fresh = false;
        if (!e) {
          e = spr[k] = el('div', 'sprite enter'); e.dataset.who = k; var im = el('img'); im.alt = c.name || k; im.decoding = 'async'; e.appendChild(im); sprites.appendChild(e); fresh = true;
          setTimeout(function () { e.classList.remove('enter'); }, 700);
        }
        e.className = 'sprite ' + (c.kind || '') + ' pos-' + w.pos + (fresh ? ' enter' : '');
        var src = A + c.dir + '/' + k + '-' + w.ex + '.webp', im2 = $('img', e);
        if (im2.getAttribute('src') !== src) im2.setAttribute('src', src);
        if (speaker) { if (k === speaker) { e.classList.add('speaking'); if (mode === 'shout') e.classList.add('shake'); } else e.classList.add('dim'); }
      });
    }
    function bubblePos(st, b) {
      if (b.pos) return b.pos;
      var s = st.sprites[b.who]; if (st.mode === 'scene' && s) return s.pos;
      return (chars[b.who] && chars[b.who].side) || 'L';
    }
    function typeText(node, text, cb) {
      clearInterval(typing); typing = null;
      if (!LIW.motion) { node.textContent = text; if (cb) cb(); return; }
      var i = 0, step = Math.max(1, Math.round(text.length / 90)); node.textContent = '';
      typing = setInterval(function () {
        i = Math.min(text.length, i + step); node.textContent = text.slice(0, i);
        if (i >= text.length) { clearInterval(typing); typing = null; if (cb) cb(); }
      }, 22 / speed);
    }
    function finishTyping() { if (!typing) return false; clearInterval(typing); typing = null; var b = beats[stops[cur]]; var t = $('.txt', bubble) || $('.txt', cap); if (t && b && b.v) t.textContent = b.v; afterShow(); return true; }
    function hideTalk() { bubble.classList.remove('on'); cap.classList.remove('on'); sfx.classList.remove('on'); }
    function closeOverlay() { if (dg) { dg.stop(); dg = null; } ov.classList.remove('on'); board.innerHTML = ''; }
    function progress() {
      var pct = Math.round((cur + 1) / stops.length * 100); $('#prog').style.width = pct + '%';
      $('#count').textContent = (cur + 1) + ' / ' + stops.length;
    }
    function go(si, fromUser) {
      if (si < 0 || si >= stops.length) return;
      clearTimeout(timer); clearInterval(typing); typing = null;
      closeOverlay(); cur = si; quizDone = false;
      var st = stateAt(si), b = st.beat;
      if (st.mode === 'panel') setPanel(st.panel); else setPanel(null);
      setBg(st.bg);
      scene.textContent = st.scene ? '📍 ' + st.scene : ''; scene.classList.toggle('on', !!st.scene);
      hideTalk();
      var speaker = b.t === 'say' ? b.who : null;
      renderSprites(st, speaker, b.mode);
      progress();
      if (b.t === 'say') {
        var c = chars[b.who] || { name: b.who, color: '#334155' }, pos = bubblePos(st, b);
        bubble.className = 'bubble at-' + pos + ' ' + (b.mode || 'say') + (c.kind === 'villain' ? ' villain' : '');
        bubble.style.setProperty('--c', c.color);
        bubble.innerHTML = '<span class="nm">' + esc(c.name) + '</span><span class="txt"></span><span class="more">▼</span>';
        void bubble.offsetWidth; bubble.classList.add('on');
        typeText($('.txt', bubble), b.v, afterShow);
        addLog(c, b);
      } else if (b.t === 'cap') {
        cap.innerHTML = '<span class="txt"></span>'; cap.classList.add('on');
        typeText($('.txt', cap), b.v, afterShow); addLog(null, b);
      } else if (b.t === 'sfx') {
        sfx.textContent = b.v; void sfx.offsetWidth; sfx.classList.add('on');
        if (LIW.motion) { stage.animate && stage.animate([{ transform: 'translate(0,0)' }, { transform: 'translate(-6px,3px)' }, { transform: 'translate(5px,-3px)' }, { transform: 'translate(0,0)' }], { duration: 380 }); }
        afterShow();
      } else if (b.t === 'card') { showCard(b); }
      else if (b.t === 'diagram') { showDiagram(b); }
      else if (b.t === 'quiz') { showQuiz(b); }
      else if (b.t === 'end') { showEnd(); }
      try { history.replaceState(null, '', '#' + (si + 1)); } catch (e) { }
    }
    function afterShow() {
      if (!auto) return;
      var b = beats[stops[cur]], ms = 1500;
      if (b.t === 'say') ms = 1300 + b.v.length * 40;
      if (b.t === 'cap') ms = 1300 + b.v.length * 36;
      if (b.t === 'sfx') ms = 1200;
      if (b.t === 'card') ms = Math.min(26000, 5000 + (b.len || 400) * 22);
      schedule(ms);
    }
    function schedule(ms) { clearTimeout(timer); timer = setTimeout(function () { next(); }, ms / speed); }
    function next() {
      if (finishTyping()) return;
      var b = beats[stops[cur]];
      if (b && b.t === 'quiz' && !quizDone) { revealQuiz(-1); return; }
      if (cur < stops.length - 1) go(cur + 1, true);
    }
    function prev() { if (cur > 0) go(cur - 1, true); }
    function boardGo(extra) {
      return '<div class="go">' + (extra || '') + '<button class="btn" data-go="prev">◀ Back</button><button class="btn primary" data-go="next">Continue ▶</button></div>';
    }
    function showBoard(html, accent) {
      board.innerHTML = html; board.style.borderTopColor = accent || D.accent;
      ov.classList.add('on'); ov.scrollTop = 0;
      var nb = $('[data-go="next"]', board); if (nb) setTimeout(function () { try { nb.focus({ preventScroll: true }); } catch (e) { } }, 60);
    }
    function showCard(b) { showBoard(b.v + boardGo()); afterShow(); }
    function showDiagram(b) {
      showBoard('<div class="dg-host"></div>' + (b.src || '') + boardGo());
      dg = new Diagram($('.dg-host', board), dgs[b.v], { speed: speed });
      if (auto) dg.playAll(function () { schedule(1400); });
      else if (LIW.motion) dg.playAll();
    }
    function showQuiz(b) {
      var L = 'ABCDEF', h = '<div class="quiz"><span class="kick">Checkpoint</span><p class="q">' + esc(b.q) + '</p><div class="opts">';
      b.opts.forEach(function (o, i) { h += '<button class="opt" data-i="' + i + '"><span class="l">' + L[i] + '</span><span>' + esc(o.t) + '</span></button>'; });
      h += '</div><div class="fb" id="fb"></div></div>';
      showBoard(h + boardGo('<button class="btn" data-go="reveal">Show answer</button>'));
      if (auto) timer = setTimeout(function () { revealQuiz(-1); }, 4200 / speed);
    }
    function revealQuiz(pick) {
      var b = beats[stops[cur]]; if (quizDone) return; quizDone = true;
      var ok = -1; b.opts.forEach(function (o, i) { if (o.ok) ok = i; });
      $$('.opt', board).forEach(function (o, i) { o.disabled = true; if (i === ok) o.classList.add('right'); else if (i === pick) o.classList.add('wrong'); });
      var fb = $('#fb', board), good = pick === ok;
      fb.className = 'fb on ' + (pick < 0 ? '' : (good ? 'ok' : 'no'));
      fb.innerHTML = '<b>' + (pick < 0 ? 'Answer: ' + 'ABCDEF'[ok] + '. ' : (good ? '✔ Correct! ' : '✘ Not quite. The answer is ' + 'ABCDEF'[ok] + '. ')) + '</b>' + esc(b.ex) + (b.src || '');
      var rv = $('[data-go="reveal"]', board); if (rv) rv.remove();
      if (auto) schedule(3200 + b.ex.length * 32);
    }
    function showEnd() {
      var done = JSON.parse(store.get('liw-done', '[]')); if (done.indexOf(D.id) < 0) { done.push(D.id); store.set('liw-done', JSON.stringify(done)); }
      var nx = D.next ? '<div class="nx"><img src="' + A + 'panels/' + D.next.thumb + '-sm.webp" alt=""><div style="text-align:left"><span class="kick">Up next · ' + esc(D.next.n) + '</span><h3>' + esc(D.next.title) + '</h3><a class="btn primary" href="' + D.next.href + (auto ? '?autoplay=1&speed=' + speed : '') + '">Read it ▶</a></div></div>' : '<p class="big">🎉 You finished the whole story!</p>';
      showBoard('<div class="endc"><span class="kick">' + esc(D.n) + ' complete</span><p class="big">✔ ' + esc(D.title) + '</p>' + nx +
        '<div class="go" style="justify-content:center"><button class="btn" data-go="restart">↺ Read again</button><a class="btn" href="index.html#episodes">All episodes</a><a class="btn" href="cheatsheet.html">Cheat sheet</a><a class="btn" href="sources.html#' + D.id + '">Sources</a><button class="btn" data-pptx="episode">⬇ PowerPoint</button></div></div>');
      if (auto && D.next) { timer = setTimeout(function () { location.href = D.next.href + '?autoplay=1&speed=' + speed; }, 6000 / speed); }
    }
    function addLog(c, b) {
      if (logList.querySelector('[data-s="' + cur + '"]')) return;
      var p = el('p'); p.dataset.s = cur;
      p.innerHTML = c ? '<b style="--c:' + c.color + '">' + esc(c.name) + '</b> ' + esc(b.v) : '<i>' + esc(b.v) + '</i>';
      logList.appendChild(p);
    }
    function setAuto(on) {
      auto = on; var b = $('#bAuto'); b.setAttribute('aria-pressed', on ? 'true' : 'false'); b.textContent = on ? '❚❚ Pause auto-play' : '▷ Auto-play';
      clearTimeout(timer);
      if (on) {
        var bt = beats[stops[cur]];
        if (bt.t === 'diagram' && dg) dg.playAll(function () { schedule(1200); });
        else if (bt.t === 'quiz' && !quizDone) timer = setTimeout(function () { revealQuiz(-1); }, 2500 / speed);
        else if (!typing) afterShow();
      } else if (dg) dg.pause();
    }
    // events
    stage.addEventListener('click', function () { next(); });
    $('#talk').addEventListener('click', function (ev) { if (document.body.classList.contains('compact')) next(); });
    $('#bNext').addEventListener('click', next); $('#bPrev').addEventListener('click', prev);
    $('#bAuto').addEventListener('click', function () { setAuto(!auto); });
    $('#speed').addEventListener('change', function (e) { speed = parseFloat(e.target.value) || 1; if (dg) dg.speed = speed; });
    function toggleLog() { $('#log').classList.toggle('on'); }
    function toggleText() { document.body.classList.toggle('text-mode'); if (document.body.classList.contains('text-mode')) { setAuto(false); $('#transcript').scrollIntoView(); } }
    $('#bLog').addEventListener('click', toggleLog); $('#bLogX').addEventListener('click', toggleLog);
    $('#bText').addEventListener('click', toggleText);
    $$('[data-act="comic"]').forEach(function (b) { b.addEventListener('click', toggleText); });
    function full() { var d = document; if (!d.fullscreenElement) { (d.documentElement.requestFullscreen || function () { }).call(d.documentElement); } else d.exitFullscreen(); }
    $('#bFull').addEventListener('click', full);
    $('#bPptx').addEventListener('click', function () { setAuto(false); });
    board.addEventListener('click', function (ev) {
      var o = ev.target.closest('.opt'); if (o && !o.disabled) { clearTimeout(timer); revealQuiz(+o.dataset.i); return; }
      var g = ev.target.closest('[data-go]'); if (!g) return;
      var a = g.dataset.go;
      if (a === 'next') next(); if (a === 'prev') prev(); if (a === 'reveal') { clearTimeout(timer); revealQuiz(-1); }
      if (a === 'restart') { logList.innerHTML = ''; go(0); }
    });
    document.addEventListener('keydown', function (ev) {
      if (ev.altKey || ev.ctrlKey || ev.metaKey) return;
      if (document.querySelector('.xp-ov')) return; // the export dialog has the keyboard
      var t = ev.target; if (t && (t.tagName === 'SELECT' || t.tagName === 'INPUT' || t.tagName === 'TEXTAREA')) return;
      var k = ev.key;
      if (k === 'ArrowRight' || k === ' ' || k === 'PageDown') { if (t && t.tagName === 'BUTTON' && k === ' ') return; ev.preventDefault(); next(); }
      else if (k === 'Enter') { if (t && (t.tagName === 'BUTTON' || t.tagName === 'A')) return; ev.preventDefault(); next(); }
      else if (k === 'ArrowLeft' || k === 'PageUp') { ev.preventDefault(); prev(); }
      else if (k === 'a' || k === 'A') setAuto(!auto);
      else if (k === 'l' || k === 'L') toggleLog();
      else if (k === 't' || k === 'T') toggleText();
      else if (k === 'f' || k === 'F') full();
      else if (k === 'm' || k === 'M') LIW.toggleMotion();
      else if (k === 'p' || k === 'P') { setAuto(false); LIW.exportPptx('episode'); }
      else if (k === 'Escape') { $('#log').classList.remove('on'); document.body.classList.remove('text-mode'); }
    });
    document.addEventListener('liw-motion', function () { if (cur >= 0 && !ov.classList.contains('on')) { /* keep state */ } });
    function compactCheck() { var w = $('.stage-wrap').getBoundingClientRect().width; document.body.classList.toggle('compact', w < 640); }
    window.addEventListener('resize', compactCheck); compactCheck();
    var sp = parseFloat(params.get('speed')); if (sp >= 0.5 && sp <= 3) { speed = sp; var sel = $('#speed'); if (![].some.call(sel.options, function (o) { return +o.value === sp; })) { var op = el('option', '', sp + '×'); op.value = sp; sel.appendChild(op); } sel.value = String(sp); }
    var h = parseInt((location.hash || '').slice(1), 10);
    go(h > 0 && h <= stops.length ? h - 1 : 0);
    if (params.get('autoplay') === '1') setAuto(true);
  };

  /* ------------------------------------------------------------ home page */
  LIW.home = function () {
    var done = JSON.parse(store.get('liw-done', '[]'));
    $$('.ep[data-id]').forEach(function (a) { if (done.indexOf(a.dataset.id) >= 0) a.classList.add('is-done'); });
    var firstUndone = $$('.ep[data-id]').filter(function (a) { return done.indexOf(a.dataset.id) < 0; })[0];
    var c = $('#continue');
    if (c && firstUndone && done.length) { c.href = firstUndone.getAttribute('href'); c.innerHTML = '▶ Continue: ' + esc($('h3', firstUndone).textContent); c.hidden = false; }
    LIW.mountDiagrams();
  };
  LIW.gallery = function () { LIW.mountDiagrams(); };
})();
