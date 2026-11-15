/* Land It Well, Dewi! · Export to PowerPoint
   Builds an animated .pptx right in the browser: one slide per comic beat with Morph transitions and typed
   speech, lesson cards and quizzes, editable architecture diagrams with the official Azure icons whose packets
   fly along motion paths, and speaker notes on every slide. No server and no third-party libraries. */
(function () {
  'use strict';
  var LIW = window.LIW = window.LIW || {};
  if (LIW.pptx) return;

  /* ============================================================ units, colours, fonts */
  var IN = 914400, PT = 12700, SW = 12192000, SH = 6858000;
  var FONT = 'Calibri', FONT_D = 'Trebuchet MS', FONT_E = 'Segoe UI Emoji';
  var INK = '0B1430', BODY = '334155', MUTED = '475569', PAPER = 'FFFDF8', NIGHT = '0A1022', STAGE = '0D1530', LINK = '1D4ED8';
  var KIND = { req: '06B6D4', resp: 'F59E0B', block: 'EF4444', policy: '8B5CF6', data: '10B981', alert: 'F97316', money: 'CA8A04' };
  var KIND_NAME = { req: 'request', resp: 'response', block: 'blocked', policy: 'policy', data: 'data / ops', alert: 'signal / alert', money: 'money' };
  var LEGEND = ['req', 'resp', 'data', 'policy', 'alert', 'block'];
  var ZONE = {
    _: ['FFFFFF', '9FB3CF', 'dash'], caf: ['E0F2FE', '38BDF8', 'dash'], waf: ['EDE9FE', 'A78BFA', 'dash'], lz: ['ECFEFF', '06B6D4', 'dash'],
    ongoing: ['FFF7ED', 'FB923C', 'dash'], hub: ['E0F2FE', '0284C7', null], spoke: ['F0FDF4', '22C55E', null], onprem: ['F1F5F9', '64748B', 'dash'],
    azure: ['F8FBFF', '3B82F6', 'dash'], test: ['FEFCE8', 'EAB308', 'dash'], mgmt: ['FAF5FF', 'A855F7', 'dash'], gov: ['FDF2F8', 'EC4899', 'dash'],
    protect: ['F0F9FF', '0EA5E9', 'dash'], paas: ['FFF7ED', 'F97316', 'dash'], subnet: ['FFFFFF', '86EFAC', 'dash'], az: ['EFF6FF', '60A5FA', 'sysDash'],
    lanes: ['FAFAF9', 'A8A29E', 'dash']
  };
  var RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  var NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
  var XH = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  var GRP = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
  var HLCLR = '<a:extLst><a:ext uri="{A12FA001-AC4F-418D-AE19-62706E023703}"><ahyp:hlinkClr xmlns:ahyp="http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor" val="tx"/></a:ext></a:extLst>';

  function I(v) { return Math.round(v * IN); }
  function R(v) { return Math.round(v); }
  function hx(c) { return String(c || '000000').replace('#', '').toUpperCase(); }
  function xe(s) { return String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function clamp(v, a, b) { return Math.max(a, Math.min(b, v)); }
  function plain(runs) { return typeof runs === 'string' ? runs : (runs || []).map(function (r) { return r.t; }).join(''); }
  function f4(v) { return (Math.round(v * 10000) / 10000).toString(); }
  function guid() { var h = '0123456789ABCDEF', s = '{'; for (var i = 0; i < 32; i++) { if (i === 8 || i === 12 || i === 16 || i === 20) s += '-'; s += h[Math.floor(Math.random() * 16)]; } return s + '}'; }
  function rgb(h) { h = hx(h); return [0, 2, 4].map(function (i) { return parseInt(h.substr(i, 2), 16); }); }
  function lumi(h) { var w = [0.2126, 0.7152, 0.0722]; return rgb(h).reduce(function (a, v, i) { v /= 255; return a + w[i] * (v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4)); }, 0); }
  function contrast(a, b) { var x = lumi(a), y = lumi(b); return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05); }
  function shade(h, f) { return rgb(h).map(function (v) { var t = Math.round(v * (1 - f)).toString(16); return t.length < 2 ? '0' + t : t; }).join('').toUpperCase(); }
  function inkOn(h, bg, min) { min = min || 4.6; var c = hx(h); for (var i = 1; i <= 19 && contrast(c, bg) < min; i++) c = shade(h, i * 0.05); return c; }
  function deep(h) { return inkOn(h, 'FFFFFF'); } // a fill dark enough to carry white text
  function slug(t) { return String(t).replace(/[^A-Za-z0-9]+/g, '-').replace(/^-|-$/g, ''); }

  /* ============================================================ text measuring */
  var mctx = null, mcache = {};
  // Measure with the deck's own fonts where the browser has them; otherwise fall back to a WIDER face so boxes err on the roomy side.
  var STACK = { 'Calibri': '"Calibri","Carlito","Helvetica Neue",Arial,sans-serif', 'Trebuchet MS': '"Trebuchet MS",Arial,"Liberation Sans",sans-serif', 'Consolas': 'Consolas,"Liberation Mono","DejaVu Sans Mono",monospace' };
  function fstack(font) { return STACK[font] || ('"' + font + '",Arial,sans-serif'); }
  function tw(text, pt, bold, font, slack) { // EMU, with a little slack for font differences between systems
    font = font || FONT; var k = (bold ? 1 : 0) + font + '\u0001' + text, w = mcache[k];
    if (w == null) {
      if (!mctx) mctx = document.createElement('canvas').getContext('2d');
      mctx.font = (bold ? '700 ' : '400 ') + '100px ' + fstack(font);
      w = mcache[k] = mctx.measureText(text).width / 100;
    }
    return w * pt * PT * (slack || 1.06);
  }
  function lineCount(runs, pt, font, width, boldAll) {
    if (typeof runs === 'string') runs = [{ t: runs }];
    var lines = 1, x = 0, space = tw(' ', pt, false, font, 1.02);
    runs.forEach(function (r) {
      String(r.t).split('\n').forEach(function (part, pi) {
        if (pi > 0) { lines++; x = 0; }
        part.split(/\s+/).forEach(function (wd) {
          if (!wd) return;
          var ww = tw(wd, pt, boldAll || r.b, r.c ? 'Consolas' : font, 1.02);
          if (x > 0 && x + space + ww > width) { lines++; x = ww; } else x += (x > 0 ? space : 0) + ww;
          while (x > width * 1.02 && ww > width) { lines++; x -= width; }
        });
      });
    });
    return lines;
  }

  /* ============================================================ DrawingML snippets */
  function clr(c, a) { return '<a:srgbClr val="' + hx(c) + '"' + (a != null && a < 1 ? '><a:alpha val="' + R(a * 100000) + '"/></a:srgbClr>' : '/>'); }
  function solid(c, a) { return '<a:solidFill>' + clr(c, a) + '</a:solidFill>'; }
  function grad(stops, ang, radial) {
    return '<a:gradFill rotWithShape="1"><a:gsLst>' + stops.map(function (s) { return '<a:gs pos="' + R(s[0] * 100000) + '">' + clr(s[1], s[2]) + '</a:gs>'; }).join('') + '</a:gsLst>' +
      (radial ? '<a:path path="circle"><a:fillToRect l="50000" t="40000" r="50000" b="60000"/></a:path>' : '<a:lin ang="' + R((ang || 0) * 60000) + '" scaled="0"/>') + '</a:gradFill>';
  }
  function lnXml(l) {
    if (!l) return '<a:ln><a:noFill/></a:ln>';
    return '<a:ln w="' + R(l.w || PT) + '"' + (l.cap ? ' cap="' + l.cap + '"' : '') + '>' + solid(l.c, l.a) + (l.dash ? '<a:prstDash val="' + l.dash + '"/>' : '') + '<a:round/>' +
      (l.tail ? '<a:tailEnd type="' + l.tail + '" w="lg" len="lg"/>' : '') + '</a:ln>';
  }
  function shadowXml(d) { return '<a:outerShdw blurRad="' + R(d.blur || 0) + '" dist="' + R(d.dist || 0) + '" dir="' + R((d.dir == null ? 90 : d.dir) * 60000) + '" algn="ctr" rotWithShape="0">' + clr(d.c || '000000', d.a) + '</a:outerShdw>'; }
  function fxXml(o) {
    var s = '';
    if (o.glow) s += '<a:glow rad="' + R(o.glow.r) + '">' + clr(o.glow.c, o.glow.a) + '</a:glow>';
    if (o.shadow) s += shadowXml(o.shadow);
    return s ? '<a:effectLst>' + s + '</a:effectLst>' : '';
  }
  var EMOJI = /([\u{1F000}-\u{1FAFF}][\u{FE0F}\u{200D}\u{1F000}-\u{1FAFF}]*|[\u2600-\u27BF]\uFE0F)/u;
  function splitEmoji(r) {
    var t = String(r.t); if (!EMOJI.test(t)) return [r];
    return t.split(EMOJI).filter(function (p) { return p; }).map(function (p) { var o = {}; for (var k in r) o[k] = r[k]; o.t = p; if (EMOJI.test(p)) o.font = FONT_E; return o; });
  }
  function runXml(s, r, d) {
    var sz = r.sz || d.sz || 14, font = r.font || (r.c ? 'Consolas' : d.font) || FONT, link = r.link || r.u;
    var a = ' lang="en-US" sz="' + R(sz * 100) + '"';
    if (r.b != null ? r.b : d.b) a += ' b="1"';
    if (r.i || d.i) a += ' i="1"';
    var spc = r.spc != null ? r.spc : d.spc; if (spc) a += ' spc="' + R(spc) + '"';
    if (r.caps || d.caps) a += ' cap="all"';
    a += ' dirty="0"';
    var x = '', ol = r.outline || d.outline, gl = r.glow || d.glow, sh = r.shadow || d.shadow;
    if (ol) x += '<a:ln w="' + R(ol.w) + '">' + solid(ol.c) + '<a:round/></a:ln>';
    x += solid(r.color || (link ? d.linkColor : null) || d.color || '111111', r.alpha != null ? r.alpha : d.alpha);
    if (gl || sh) x += '<a:effectLst>' + (gl ? '<a:glow rad="' + R(gl.r) + '">' + clr(gl.c, gl.a) + '</a:glow>' : '') + (sh ? shadowXml(sh) : '') + '</a:effectLst>';
    x += '<a:latin typeface="' + xe(font) + '"/><a:ea typeface="' + xe(font) + '"/><a:cs typeface="' + xe(font) + '"/>';
    if (link) x += '<a:hlinkClick r:id="' + s.link(link) + '">' + HLCLR + '</a:hlinkClick>';
    else if (r.jump) x += '<a:hlinkClick r:id="' + s.jump(r.jump) + '" action="ppaction://hlinksldjump">' + HLCLR + '</a:hlinkClick>';
    return '<a:r><a:rPr' + a + '>' + x + '</a:rPr><a:t>' + xe(r.t) + '</a:t></a:r>';
  }
  function paraXml(s, p) {
    var a = (p.marL ? ' marL="' + R(p.marL) + '" indent="' + R(p.ind || 0) + '"' : '') + ' algn="' + (p.algn || 'l') + '"';
    var x = '';
    if (p.sp) x += '<a:lnSpc><a:spcPts val="' + R(p.sp * 100) + '"/></a:lnSpc>';
    if (p.bef) x += '<a:spcBef><a:spcPts val="' + R(p.bef * 100) + '"/></a:spcBef>';
    x += p.bullet ? '<a:buClr><a:srgbClr val="' + hx(p.bullet.c || '111111') + '"/></a:buClr><a:buSzPct val="100000"/><a:buFont typeface="Arial"/><a:buChar char="' + xe(p.bullet.ch || '•') + '"/>' : '<a:buNone/>';
    var runs = []; (p.runs || []).forEach(function (r) { if (r && r.t !== '' && r.t != null) runs = runs.concat(splitEmoji(r)); });
    return '<a:p><a:pPr' + a + '>' + x + '</a:pPr>' + runs.map(function (r) { return runXml(s, r, p); }).join('') + '<a:endParaRPr lang="en-US" sz="' + R((p.sz || 14) * 100) + '" dirty="0"/></a:p>';
  }
  function P(runs, o) { var p = {}; for (var k in o) p[k] = o[k]; p.runs = typeof runs === 'string' ? [{ t: runs }] : (runs || []); return p; }
  function bodyXml(s, paras, o) {
    o = o || {}; var ins = o.ins || [0, 0, 0, 0];
    return '<p:txBody><a:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="overflow" horzOverflow="overflow" vert="horz" wrap="' + (o.wrap === false ? 'none' : 'square') +
      '" lIns="' + R(ins[0]) + '" tIns="' + R(ins[1]) + '" rIns="' + R(ins[2]) + '" bIns="' + R(ins[3]) + '" numCol="1" anchor="' + (o.anchor || 't') + '" anchorCtr="0" rtlCol="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>' +
      (paras && paras.length ? paras.map(function (p) { return paraXml(s, p); }).join('') : '<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>') + '</p:txBody>';
  }

  /* ============================================================ deck + slides */
  function Deck(D, opts) { this.D = D; this.opts = opts || {}; this.slides = []; this.media = {}; this.mediaList = []; this.sections = []; this.anchors = {}; }
  Deck.prototype.addMedia = function (key, bytes, ext, w, h) {
    var m = this.media[key]; if (m) return m;
    m = this.media[key] = { key: key, file: 'image' + (this.mediaList.length + 1) + '.' + ext, bytes: bytes, w: w, h: h };
    this.mediaList.push(m); return m;
  };
  Deck.prototype.slide = function (o) { var s = new Slide(this, o || {}); this.slides.push(s); return s; };
  Deck.prototype.section = function (name) {
    var last = this.sections[this.sections.length - 1];
    if (last && last.start === this.slides.length) last.name = name; else this.sections.push({ name: name, start: this.slides.length });
  };
  function Slide(deck, o) {
    this.deck = deck; this.nid = 2; this.parts = []; this.fx = []; this.pics = {}; this.relKey = {};
    this.rels = [{ id: 'rId1', type: RT + 'slideLayout', target: '../slideLayouts/slideLayout1.xml' }];
    this.bg = o.bg || NIGHT; this.ctx = o.ctx || null; this.notes = ''; this.adv = 0; this.trans = 'fade';
  }
  Slide.prototype.rel = function (type, target, ext) {
    var k = type + '|' + (typeof target === 'string' ? target : '#' + target.anchor);
    if (this.relKey[k]) return this.relKey[k];
    var id = 'rId' + (this.rels.length + 1); this.rels.push({ id: id, type: type, target: target, ext: !!ext }); this.relKey[k] = id; return id;
  };
  Slide.prototype.link = function (url) { return this.rel(RT + 'hyperlink', url, true); };
  Slide.prototype.jump = function (anchor) { return this.rel(RT + 'slide', { anchor: anchor }); };
  Slide.prototype.sp = function (o) {
    var id = this.nid++, w = Math.max(R(o.w), 1), h = Math.max(R(o.h), 1);
    var rot = o.rot ? ' rot="' + R((((o.rot % 360) + 360) % 360) * 60000) + '"' : '';
    var geom = o.cust || ('<a:prstGeom prst="' + (o.geom || 'rect') + '"><a:avLst>' + (o.adj ? Object.keys(o.adj).map(function (k) { return '<a:gd name="' + k + '" fmla="val ' + R(o.adj[k]) + '"/>'; }).join('') : '') + '</a:avLst></a:prstGeom>');
    var fill = o.fillXml || (o.fill ? solid(o.fill, o.fa) : '<a:noFill/>');
    this.parts.push('<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="' + xe(o.name || ((o.txBox ? 'TextBox ' : 'Shape ') + id)) + '"' + (o.descr ? ' descr="' + xe(o.descr) + '"' : '') + '/><p:cNvSpPr' + (o.txBox ? ' txBox="1"' : '') + '/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm' + rot + '><a:off x="' + R(o.x) + '" y="' + R(o.y) + '"/><a:ext cx="' + w + '" cy="' + h + '"/></a:xfrm>' + geom + fill + lnXml(o.line) + fxXml(o) + '</p:spPr>' +
      bodyXml(this, o.paras, o.tx) + '</p:sp>');
    return id;
  };
  Slide.prototype.text = function (x, y, w, h, paras, o) {
    o = o || {};
    return this.sp({ x: x, y: y, w: w, h: h, paras: paras, txBox: true, name: o.name, rot: o.rot, fill: o.fill, fa: o.fa, line: o.line, tx: { anchor: o.anchor || 't', ins: o.ins, wrap: o.wrap } });
  };
  Slide.prototype.pic = function (o) {
    var m = o.media; if (!m) return null;
    var id = this.nid++, rid = this.rel(RT + 'image', '../media/' + m.file); m.used = true;
    var ext = m.svg ? '<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="' + this.rel(RT + 'image', '../media/' + m.svg.file) + '"/></a:ext></a:extLst>' : '';
    var bfx = (o.alpha != null && o.alpha < 1 ? '<a:alphaModFix amt="' + R(o.alpha * 100000) + '"/>' : '') + (o.lum ? '<a:lum bright="' + R(o.lum) + '"/>' : '');
    var c = o.crop, src = c ? '<a:srcRect l="' + R(c.l * 100000) + '" t="' + R(c.t * 100000) + '" r="' + R(c.r * 100000) + '" b="' + R(c.b * 100000) + '"/>' : '';
    this.parts.push('<p:pic><p:nvPicPr><p:cNvPr id="' + id + '" name="' + xe(o.name || ('Picture ' + id)) + '" descr="' + xe(o.descr || '') + '"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="' + rid + '">' + bfx + ext + '</a:blip>' + src + '<a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="' + R(o.x) + '" y="' + R(o.y) + '"/><a:ext cx="' + Math.max(1, R(o.w)) + '" cy="' + Math.max(1, R(o.h)) + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' + fxXml(o) + '</p:spPr></p:pic>');
    this.pics[id] = 1; return id;
  };
  Slide.prototype.cover = function (m, x, y, w, h, o) { // like CSS object-fit: cover
    if (!m) return null;
    var ir = m.w / m.h, fr = w / h, c = { l: 0, t: 0, r: 0, b: 0 }, k;
    if (ir > fr) { k = (1 - fr / ir) / 2; c.l = c.r = k; } else { k = (1 - ir / fr) / 2; c.t = c.b = k; }
    o = o || {}; o.media = m; o.x = x; o.y = y; o.w = w; o.h = h; o.crop = c; return this.pic(o);
  };
  Slide.prototype.contain = function (m, x, y, w, h, o) {
    if (!m) return null;
    var ir = m.w / m.h, ww = w, hh = w / ir; if (hh > h) { hh = h; ww = h * ir; }
    o = o || {}; o.media = m; o.x = x + (w - ww) / 2; o.y = y + (h - hh) / (o.bottom ? 1 : 2); o.w = ww; o.h = hh; return this.pic(o);
  };
  Slide.prototype.path = function (cmds, o) { // cmds in EMU: {t:'M'|'L', p} or {t:'Q', c, p}
    var xs = [], ys = [];
    cmds.forEach(function (c) { xs.push(c.p.x); ys.push(c.p.y); if (c.c) { xs.push(c.c.x); ys.push(c.c.y); } });
    var x0 = Math.min.apply(null, xs), y0 = Math.min.apply(null, ys);
    var w = Math.max(R(Math.max.apply(null, xs) - x0), 2), h = Math.max(R(Math.max.apply(null, ys) - y0), 2);
    function pt(p) { return '<a:pt x="' + R(p.x - x0) + '" y="' + R(p.y - y0) + '"/>'; }
    o.cust = '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="' + w + '" h="' + h + '" fill="none" extrusionOk="0">' +
      cmds.map(function (c) { return c.t === 'M' ? '<a:moveTo>' + pt(c.p) + '</a:moveTo>' : c.t === 'L' ? '<a:lnTo>' + pt(c.p) + '</a:lnTo>' : '<a:quadBezTo>' + pt(c.c) + pt(c.p) + '</a:quadBezTo>'; }).join('') +
      '</a:path></a:pathLst></a:custGeom>';
    o.x = x0; o.y = y0; o.w = w; o.h = h;
    return this.sp(o);
  };
  Slide.prototype.table = function (o) {
    var id = this.nid++, self = this;
    function edge(tag, l) { return l ? '<a:' + tag + ' w="' + R(l.w) + '" cap="flat" cmpd="sng" algn="ctr">' + solid(l.c) + '<a:prstDash val="solid"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:' + tag + '>' : '<a:' + tag + ' w="12700" cmpd="sng"><a:noFill/></a:' + tag + '>'; }
    var rows = o.rows.map(function (row) {
      return '<a:tr h="' + R(row.h) + '">' + row.cells.map(function (c) {
        return '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>' + c.paras.map(function (p) { return paraXml(self, p); }).join('') + '</a:txBody><a:tcPr marL="' + R(c.m[0]) + '" marR="' + R(c.m[2]) + '" marT="' + R(c.m[1]) + '" marB="' + R(c.m[3]) + '">' +
          edge('lnL', null) + edge('lnR', null) + edge('lnT', c.top) + edge('lnB', null) + (c.fill ? solid(c.fill) : '<a:noFill/>') + '</a:tcPr></a:tc>';
      }).join('') + '</a:tr>';
    }).join('');
    var h = o.rows.reduce(function (t, r) { return t + r.h; }, 0);
    this.parts.push('<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="' + id + '" name="' + xe(o.name || ('Table ' + id)) + '"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>' +
      '<p:xfrm><a:off x="' + R(o.x) + '" y="' + R(o.y) + '"/><a:ext cx="' + R(o.w) + '" cy="' + R(h) + '"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>' +
      o.cols.map(function (cw) { return '<a:gridCol w="' + R(cw) + '"/>'; }).join('') + '</a:tblGrid>' + rows + '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>');
    return id;
  };

  /* ============================================================ animation timeline -> <p:timing>
     fx: {k:'appear'|'fade'|'zoom'|'out'|'path', id, t (ms after the slide starts), dur, lt (ms per letter), path, loop, rev, ease} */
  function timingXml(s) {
    if (!s.fx.length) return '';
    var nid = 4, slots = {}, grp = {}, bld = [];
    s.fx.forEach(function (f) { (slots[f.t] = slots[f.t] || []).push(f); });
    function tgt(id) { return '<p:tgtEl><p:spTgt spid="' + id + '"/></p:tgtEl>'; }
    function setVis(id, v, delay) { return '<p:set><p:cBhvr><p:cTn id="' + (nid++) + '" dur="1" fill="hold"><p:stCondLst><p:cond delay="' + (delay || 0) + '"/></p:stCondLst></p:cTn>' + tgt(id) + '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="' + v + '"/></p:to></p:set>'; }
    function grow(id, attr, d) { return '<p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="' + (nid++) + '" dur="' + d + '" fill="hold"/>' + tgt(id) + '<p:attrNameLst><p:attrName>' + attr + '</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:fltVal val="0"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#' + attr + '"/></p:val></p:tav></p:tavLst></p:anim>'; }
    function fade(id, how, d) { return '<p:animEffect transition="' + how + '" filter="fade"><p:cBhvr><p:cTn id="' + (nid++) + '" dur="' + d + '"/>' + tgt(id) + '</p:cBhvr></p:animEffect>'; }
    function eff(f, nt) {
      var id = nid++, head, body, extra = '', g = null;
      g = grp[f.id] = grp[f.id] == null ? 0 : grp[f.id] + 1; if (!s.pics[f.id]) bld.push('<p:bldP spid="' + f.id + '" grpId="' + g + '" animBg="1"/>');
      var d = R(f.dur || 400);
      if (f.k === 'appear') { head = ' presetID="1" presetClass="entr" presetSubtype="0"'; if (f.lt) extra = '<p:iterate type="lt"><p:tmAbs val="' + R(f.lt) + '"/></p:iterate>'; body = setVis(f.id, 'visible'); }
      else if (f.k === 'fade') { head = ' presetID="10" presetClass="entr" presetSubtype="0"'; body = setVis(f.id, 'visible') + fade(f.id, 'in', d); }
      else if (f.k === 'zoom') { head = ' presetID="53" presetClass="entr" presetSubtype="16"'; body = setVis(f.id, 'visible') + grow(f.id, 'ppt_w', d) + grow(f.id, 'ppt_h', d) + fade(f.id, 'in', d); }
      else if (f.k === 'out') { head = ' presetID="10" presetClass="exit" presetSubtype="0"'; body = fade(f.id, 'out', d) + setVis(f.id, 'hidden', d - 1); }
      else if (f.k === 'hide') { head = ' presetID="1" presetClass="exit" presetSubtype="0"'; body = setVis(f.id, 'hidden'); }
      else {
        head = ' presetID="0" presetClass="path" presetSubtype="0" accel="' + (f.ease ? 50000 : 0) + '" decel="' + (f.ease ? 50000 : 0) + '"';
        body = '<p:animMotion origin="layout" path="' + f.path + '" pathEditMode="relative" rAng="0"><p:cBhvr><p:cTn id="' + (nid++) + '" dur="' + d + '" fill="hold"/>' + tgt(f.id) +
          '<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr><p:rCtr x="0" y="0"/></p:animMotion>';
      }
      var loop = f.loop ? ' repeatCount="' + R(f.loop * 1000) + '"' + (f.rev ? ' autoRev="1"' : '') : '';
      return '<p:par><p:cTn id="' + id + '"' + head + loop + ' fill="hold" grpId="' + g + '" nodeType="' + nt + '"><p:stCondLst><p:cond delay="0"/></p:stCondLst>' + extra + '<p:childTnLst>' + body + '</p:childTnLst></p:cTn></p:par>';
    }
    var inner = Object.keys(slots).map(Number).sort(function (a, b) { return a - b; }).map(function (t) {
      var sid = nid++, effs = slots[t].map(function (f, i) { return eff(f, i === 0 ? 'afterEffect' : 'withEffect'); }).join('');
      return '<p:par><p:cTn id="' + sid + '" fill="hold"><p:stCondLst><p:cond delay="' + R(t) + '"/></p:stCondLst><p:childTnLst>' + effs + '</p:childTnLst></p:cTn></p:par>';
    }).join('');
    return '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>' +
      '<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>' + inner + '</p:childTnLst></p:cTn></p:par>' +
      '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>' +
      (bld.length ? '<p:bldLst>' + bld.join('') + '</p:bldLst>' : '') + '</p:timing>';
  }
  function transXml(s) {
    var adv = s.adv ? ' advTm="' + R(s.adv) + '"' : '';
    if (s.trans === 'morph') return '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159"><p:transition spd="slow"' + adv + '><p159:morph option="byObject"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"' + adv + '><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>';
    if (s.trans === 'fade') return '<p:transition spd="med"' + adv + '><p:fade/></p:transition>';
    return adv ? '<p:transition' + adv + '/>' : '';
  }
  function slideXml(s) {
    return XH + '<p:sld ' + NS + '><p:cSld><p:bg><p:bgPr>' + solid(s.bg) + '<a:effectLst/></p:bgPr></p:bg><p:spTree>' + GRP + s.parts.join('') + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' + transXml(s) + timingXml(s) + '</p:sld>';
  }
  function notesXml(text) {
    var paras = String(text || '').split('\n').map(function (l) { return l ? '<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>' + xe(l) + '</a:t></a:r></a:p>' : '<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>'; }).join('');
    return XH + '<p:notes ' + NS + '><p:cSld><p:spTree>' + GRP +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>' +
      (paras || '<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>') + '</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>';
  }

  /* ============================================================ package parts */
  var CLRMAP = '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
  function LVL1(sz, mj, bu) { return '<a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">' + (bu ? '<a:buNone/>' : '') + '<a:defRPr sz="' + sz + '" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+' + mj + '-lt"/><a:ea typeface="+' + mj + '-ea"/><a:cs typeface="+' + mj + '-cs"/></a:defRPr></a:lvl1pPr>'; }
  var THEME = XH + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Land It Well"><a:themeElements><a:clrScheme name="Land It Well">' +
    '<a:dk1><a:srgbClr val="0F172A"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="0A1022"/></a:dk2><a:lt2><a:srgbClr val="FFFDF8"/></a:lt2>' +
    '<a:accent1><a:srgbClr val="0EA5E9"/></a:accent1><a:accent2><a:srgbClr val="8B5CF6"/></a:accent2><a:accent3><a:srgbClr val="FBBF24"/></a:accent3><a:accent4><a:srgbClr val="2DD4BF"/></a:accent4>' +
    '<a:accent5><a:srgbClr val="F472B6"/></a:accent5><a:accent6><a:srgbClr val="16A34A"/></a:accent6><a:hlink><a:srgbClr val="1D4ED8"/></a:hlink><a:folHlink><a:srgbClr val="7C3AED"/></a:folHlink></a:clrScheme>' +
    '<a:fontScheme name="Land It Well"><a:majorFont><a:latin typeface="Trebuchet MS"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Segoe UI"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>' +
    '<a:fmtScheme name="Land It Well"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>' +
    '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>' +
    '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
    '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>' +
    '</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>';
  var MASTER = XH + '<p:sldMaster ' + NS + '><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>' + GRP + '</p:spTree></p:cSld>' + CLRMAP +
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle>' + LVL1(4400, 'mj', 1) + '</p:titleStyle><p:bodyStyle>' + LVL1(1800, 'mn', 1) + '</p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>' + LVL1(1800, 'mn', 0) + '</p:otherStyle></p:txStyles></p:sldMaster>';
  var LAYOUT = XH + '<p:sldLayout ' + NS + ' type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>' + GRP + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>';
  var NOTES_MASTER = XH + '<p:notesMaster ' + NS + '><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>' + GRP +
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="1143000"/><a:ext cx="5486400" cy="3086100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr></p:sp>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="4400550"/><a:ext cx="5486400" cy="3600450"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>Notes</a:t></a:r></a:p></p:txBody></p:sp>' +
    '</p:spTree></p:cSld>' + CLRMAP + '<p:notesStyle>' + LVL1(1200, 'mn', 0) + '</p:notesStyle></p:notesMaster>';
  function relsXml(rels, deck) {
    return XH + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rels.map(function (r) {
      var t = r.target;
      if (typeof t !== 'string') { var to = deck.anchors[t.anchor]; t = to ? 'slide' + (deck.slides.indexOf(to) + 1) + '.xml' : 'slide1.xml'; }
      return '<Relationship Id="' + r.id + '" Type="' + r.type + '" Target="' + xe(t) + '"' + (r.ext ? ' TargetMode="External"' : '') + '/>';
    }).join('') + '</Relationships>';
  }
  function R1(id, type, target) { return { id: id, type: RT + type, target: target }; }
  function pack(deck, prog) {
    var D = deck.D, n = deck.slides.length, files = [], ct = [];
    function add(name, data, type) { files.push({ name: name, data: data }); if (type) ct.push('<Override PartName="/' + name + '" ContentType="application/vnd.openxmlformats-officedocument.' + type + '"/>'); }
    var title = deck.opts.title || D.title;
    add('_rels/.rels', XH + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + RT + 'officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="' + RT + 'extended-properties" Target="docProps/app.xml"/></Relationships>');
    var now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    files.push({ name: 'docProps/core.xml', data: XH + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>' + xe(title) + '</dc:title><dc:subject>' + xe(D.tag) + '</dc:subject><dc:creator>' + xe(D.title) + '</dc:creator><cp:keywords>Azure; Cloud Adoption Framework; Well-Architected Framework; anime comic</cp:keywords><dc:description>' + xe(D.site) + '</dc:description><dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + now + '</dcterms:modified></cp:coreProperties>' });
    ct.push('<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>');
    add('docProps/app.xml', XH + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>' + xe(D.title) + ' PowerPoint export</Application><PresentationFormat>Widescreen</PresentationFormat><Slides>' + n + '</Slides><Notes>' + n + '</Notes></Properties>', 'extended-properties+xml');
    var sldIds = '', presRels = [R1('rId1', 'slideMaster', 'slideMasters/slideMaster1.xml'), R1('rId2', 'notesMaster', 'notesMasters/notesMaster1.xml'), R1('rId3', 'presProps', 'presProps.xml'), R1('rId4', 'viewProps', 'viewProps.xml'), R1('rId5', 'theme', 'theme/theme1.xml'), R1('rId6', 'tableStyles', 'tableStyles.xml')];
    deck.slides.forEach(function (s, i) { sldIds += '<p:sldId id="' + (256 + i) + '" r:id="rId' + (10 + i) + '"/>'; presRels.push(R1('rId' + (10 + i), 'slide', 'slides/slide' + (i + 1) + '.xml')); });
    var sec = '';
    if (deck.sections.length) {
      if (deck.sections[0].start > 0) deck.sections.unshift({ name: 'Start', start: 0 });
      sec = '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">' + deck.sections.map(function (sc, k) {
        var end = k + 1 < deck.sections.length ? deck.sections[k + 1].start : n, l = '';
        for (var i = sc.start; i < end; i++) l += '<p14:sldId id="' + (256 + i) + '"/>';
        return '<p14:section name="' + xe(sc.name) + '" id="' + guid() + '"><p14:sldIdLst>' + l + '</p14:sldIdLst></p14:section>';
      }).join('') + '</p14:sectionLst></p:ext></p:extLst>';
    }
    add('ppt/presentation.xml', XH + '<p:presentation ' + NS + ' saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:notesMasterIdLst><p:notesMasterId r:id="rId2"/></p:notesMasterIdLst><p:sldIdLst>' + sldIds + '</p:sldIdLst><p:sldSz cx="' + SW + '" cy="' + SH + '"/><p:notesSz cx="6858000" cy="9144000"/>' +
      '<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>' + LVL1(1800, 'mn', 0) + '</p:defaultTextStyle>' + sec + '</p:presentation>', 'presentationml.presentation.main+xml');
    add('ppt/_rels/presentation.xml.rels', relsXml(presRels, deck));
    add('ppt/presProps.xml', XH + '<p:presentationPr ' + NS + '><p:showPr showNarration="1" useTimings="' + (deck.opts.auto ? 1 : 0) + '"><p:present/><p:sldAll/><p:penClr><a:prstClr val="red"/></p:penClr></p:showPr></p:presentationPr>', 'presentationml.presProps+xml');
    add('ppt/viewProps.xml', XH + '<p:viewPr ' + NS + '><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="80000"/></p:normalViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>', 'presentationml.viewProps+xml');
    add('ppt/tableStyles.xml', XH + '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>', 'presentationml.tableStyles+xml');
    add('ppt/theme/theme1.xml', THEME, 'theme+xml'); add('ppt/theme/theme2.xml', THEME, 'theme+xml');
    add('ppt/slideMasters/slideMaster1.xml', MASTER, 'presentationml.slideMaster+xml');
    add('ppt/slideMasters/_rels/slideMaster1.xml.rels', relsXml([R1('rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml'), R1('rId2', 'theme', '../theme/theme1.xml')], deck));
    add('ppt/slideLayouts/slideLayout1.xml', LAYOUT, 'presentationml.slideLayout+xml');
    add('ppt/slideLayouts/_rels/slideLayout1.xml.rels', relsXml([R1('rId1', 'slideMaster', '../slideMasters/slideMaster1.xml')], deck));
    add('ppt/notesMasters/notesMaster1.xml', NOTES_MASTER, 'presentationml.notesMaster+xml');
    add('ppt/notesMasters/_rels/notesMaster1.xml.rels', relsXml([R1('rId1', 'theme', '../theme/theme2.xml')], deck));
    deck.slides.forEach(function (s, i) {
      var k = i + 1;
      s.rel(RT + 'notesSlide', '../notesSlides/notesSlide' + k + '.xml');
      add('ppt/slides/slide' + k + '.xml', slideXml(s), 'presentationml.slide+xml');
      add('ppt/slides/_rels/slide' + k + '.xml.rels', relsXml(s.rels, deck));
      add('ppt/notesSlides/notesSlide' + k + '.xml', notesXml(s.notes), 'presentationml.notesSlide+xml');
      add('ppt/notesSlides/_rels/notesSlide' + k + '.xml.rels', relsXml([R1('rId1', 'notesMaster', '../notesMasters/notesMaster1.xml'), R1('rId2', 'slide', '../slides/slide' + k + '.xml')], deck));
    });
    deck.mediaList.filter(function (m) { return m.used; }).forEach(function (m) { files.push({ name: 'ppt/media/' + m.file, data: m.bytes, store: true }); if (m.svg) files.push({ name: 'ppt/media/' + m.svg.file, data: m.svg.bytes }); });
    files.unshift({ name: '[Content_Types].xml', data: XH + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="svg" ContentType="image/svg+xml"/>' + ct.join('') + '</Types>' });
    return zip(files, prog);
  }

  /* ============================================================ zip (stored, or deflated with CompressionStream) */
  var CRC = (function () { var t = new Uint32Array(256); for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function crc32(u) { var c = 0xFFFFFFFF; for (var i = 0; i < u.length; i++) c = CRC[(c ^ u[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  var enc = new TextEncoder();
  function deflate(u) {
    if (typeof CompressionStream === 'undefined') return Promise.resolve(null);
    try { return new Response(new Blob([u]).stream().pipeThrough(new CompressionStream('deflate-raw'))).arrayBuffer().then(function (b) { return new Uint8Array(b); }, function () { return null; }); }
    catch (e) { return Promise.resolve(null); }
  }
  function zip(files, prog) {
    var chunks = [], central = [], offset = 0, i = 0, d = new Date();
    var dt = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate(), tm = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
    function next() {
      if (i >= files.length) return Promise.resolve(finish());
      var f = files[i++], data = typeof f.data === 'string' ? enc.encode(f.data) : f.data, crc = crc32(data);
      return (f.store ? Promise.resolve(null) : deflate(data)).then(function (def) {
        var method = def && def.length < data.length ? 8 : 0, body = method ? def : data, name = enc.encode(f.name);
        var h = new DataView(new ArrayBuffer(30));
        h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, method, true); h.setUint16(10, tm, true); h.setUint16(12, dt, true);
        h.setUint32(14, crc, true); h.setUint32(18, body.length, true); h.setUint32(22, data.length, true); h.setUint16(26, name.length, true); h.setUint16(28, 0, true);
        chunks.push(new Uint8Array(h.buffer), name, body);
        var c = new DataView(new ArrayBuffer(46));
        c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint16(10, method, true); c.setUint16(12, tm, true); c.setUint16(14, dt, true);
        c.setUint32(16, crc, true); c.setUint32(20, body.length, true); c.setUint32(24, data.length, true); c.setUint16(28, name.length, true); c.setUint32(42, offset, true);
        central.push(new Uint8Array(c.buffer), name);
        offset += 30 + name.length + body.length;
        if (prog && (i % 25 === 0 || i === files.length)) prog(i / files.length);
        return next();
      });
    }
    function finish() {
      var size = central.reduce(function (s, u) { return s + u.length; }, 0), e = new DataView(new ArrayBuffer(22));
      e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true); e.setUint32(12, size, true); e.setUint32(16, offset, true);
      return new Blob(chunks.concat(central, [new Uint8Array(e.buffer)]), { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
    }
    return next();
  }

  /* ============================================================ pictures */
  function fetchBitmap(url) {
    return fetch(url).then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.blob(); }).then(function (b) {
      if (window.createImageBitmap) return createImageBitmap(b);
      return new Promise(function (res, rej) { var im = new Image(), u = URL.createObjectURL(b); im.onload = function () { res(im); }; im.onerror = rej; im.src = u; });
    });
  }
  function encode(src, type, q, bg) {
    var c = document.createElement('canvas'); c.width = src.width; c.height = src.height;
    var g = c.getContext('2d'); if (bg) { g.fillStyle = bg; g.fillRect(0, 0, c.width, c.height); } g.drawImage(src, 0, 0);
    return new Promise(function (res, rej) { c.toBlob(function (b) { if (!b) { rej(new Error('encode')); return; } b.arrayBuffer().then(function (ab) { res({ bytes: new Uint8Array(ab), w: c.width, h: c.height }); }, rej); }, type, q); });
  }
  function svgDoc(ic, size) { return '<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="' + ic.vb + '" width="' + size + '" height="' + size + '">' + ic.svg + '</svg>'; }
  function rasterSvg(text, size) {
    return new Promise(function (res, rej) {
      var im = new Image(), u = URL.createObjectURL(new Blob([text], { type: 'image/svg+xml' }));
      im.onload = function () { var c = document.createElement('canvas'); c.width = c.height = size; c.getContext('2d').drawImage(im, 0, 0, size, size); URL.revokeObjectURL(u); c.toBlob(function (b) { b.arrayBuffer().then(function (ab) { res(new Uint8Array(ab)); }, rej); }, 'image/png'); };
      im.onerror = function () { URL.revokeObjectURL(u); rej(new Error('icon')); };
      im.src = u;
    });
  }
  function loadMedia(deck, icons, keys, prog) {
    var list = Object.keys(keys), done = 0, at = 0;
    function one(key) {
      var p;
      if (key.indexOf('icon/') === 0) {
        var ic = icons[key.slice(5)];
        if (!ic) return Promise.resolve();
        var svg = svgDoc(ic, 96);
        p = rasterSvg(svgDoc(ic, 128), 128).then(function (png) { var m = deck.addMedia(key, png, 'png', 128, 128); m.svg = { file: m.file.replace('.png', '.svg'), bytes: enc.encode(svg) }; });
      } else {
        var photo = /^(bg|panels|img)\//.test(key);
        p = fetchBitmap('assets/' + key + '.webp').then(function (bm) { return encode(bm, photo ? 'image/jpeg' : 'image/png', 0.86, photo ? '#0d1530' : null); })
          .then(function (r) { deck.addMedia(key, r.bytes, photo ? 'jpeg' : 'png', r.w, r.h); });
      }
      return p.catch(function () { /* a missing picture leaves a gap rather than failing the whole deck */ }).then(function () { done++; prog(done / list.length); });
    }
    function worker() { if (at >= list.length) return Promise.resolve(); return one(list[at++]).then(worker); }
    return Promise.all([worker(), worker(), worker(), worker()]);
  }

  /* ============================================================ story state (mirrors the comic engine) */
  var STOPS = { cap: 1, sfx: 1, say: 1, card: 1, diagram: 1, quiz: 1, end: 1 };
  function framesOf(D, ep) {
    var beats = ep.beats.concat([{ t: 'end' }]), out = [], st = { bg: null, panel: null, mode: 'scene', scene: '', sprites: {} };
    function autoPos(who) {
      var used = {}; Object.keys(st.sprites).forEach(function (k) { used[st.sprites[k].pos] = 1; });
      var pref = (D.chars[who] && D.chars[who].pref) || ['L', 'C', 'R'];
      for (var i = 0; i < pref.length; i++) if (!used[pref[i]]) return pref[i];
      return pref[0];
    }
    beats.forEach(function (b) {
      if (b.t === 'bg') { st.bg = b.v; st.mode = 'scene'; st.panel = null; st.sprites = {}; }
      else if (b.t === 'panel') { st.panel = b.v; st.mode = 'panel'; }
      else if (b.t === 'scene') st.scene = b.v;
      else if (b.t === 'enter') { var o = st.sprites[b.who]; st.sprites[b.who] = { ex: b.ex || (o && o.ex) || 'base', pos: b.pos || (o && o.pos) || autoPos(b.who) }; }
      else if (b.t === 'exit') delete st.sprites[b.who];
      else if (b.t === 'say' && st.mode === 'scene') { var s0 = st.sprites[b.who]; if (!s0) st.sprites[b.who] = { ex: b.ex || 'base', pos: b.pos || autoPos(b.who) }; else if (b.ex) s0.ex = b.ex; }
      if (STOPS[b.t]) out.push({ beat: b, st: JSON.parse(JSON.stringify(st)) });
    });
    return out;
  }
  function ctxOf(st) { return st.mode === 'panel' ? 'p:' + st.panel : 'b:' + st.bg; }
  function epById(D, id) { for (var i = 0; i < D.eps.length; i++) if (D.eps[i].id === id) return D.eps[i]; return null; }
  function nextEp(D, ep) { var i = D.eps.indexOf(ep); return i >= 0 && i + 1 < D.eps.length ? D.eps[i + 1] : null; }

  /* ============================================================ plan: which slides, which pictures */
  function makePlan(D, o) {
    var items = [];
    function episode(ep, whole) {
      if (whole) items.push({ k: 'section', name: ep.n + ' · ' + ep.title });
      items.push({ k: 'epTitle', ep: ep });
      var fr = framesOf(D, ep);
      fr.forEach(function (f, i) {
        var b = f.beat, base = { ep: ep, fr: f, prog: (i + 1) / fr.length, whole: whole };
        function push(k) { var x = { k: k }; for (var kk in base) x[kk] = base[kk]; items.push(x); return x; }
        if (b.t === 'say' || b.t === 'cap' || b.t === 'sfx') push('scene');
        else if (b.t === 'card') push('card');
        else if (b.t === 'diagram') push('diagram').did = b.v;
        else if (b.t === 'quiz') { push('quiz'); push('answer'); }
        else if (b.t === 'end') push('end');
      });
      items.push({ k: 'sources', ep: ep });
    }
    if (o.scope === 'all') {
      items.push({ k: 'section', name: 'Welcome' }, { k: 'cover' }, { k: 'cast' }, { k: 'villains' }, { k: 'contents' });
      D.eps.forEach(function (ep) { episode(ep, true); });
      items.push({ k: 'section', name: 'The end' }, { k: 'closing' });
    } else if (o.scope === 'diagrams') {
      items.push({ k: 'section', name: 'Welcome' }, { k: 'dgCover' }, { k: 'dgContents' });
      D.dgOrder.forEach(function (x) { items.push({ k: 'section', name: D.diagrams[x.id].title }, { k: 'diagram', did: x.id, ep: epById(D, x.ep), standalone: true }); });
      items.push({ k: 'section', name: 'Sources' }, { k: 'dgSources' });
    } else episode(epById(D, o.ep) || D.eps[0], false);
    return items;
  }
  function slideEstimate(D, items) {
    return items.reduce(function (n, it) { return it.k === 'section' ? n : n + (it.k === 'diagram' ? D.diagrams[it.did].steps.length + 1 : it.k === 'dgSources' ? 3 : 1); }, 0);
  }
  function mediaKeys(D, items) {
    var keys = {};
    function st(s, cast) { if (s.mode === 'panel' && s.panel) keys['panels/' + s.panel] = 1; else if (s.bg) keys['bg/' + s.bg] = 1; if (cast && s.mode === 'scene') Object.keys(s.sprites).forEach(function (w) { var c = D.chars[w] || {}; keys[c.dir + '/' + w + '-' + s.sprites[w].ex] = 1; }); }
    function icons(did) { D.diagrams[did].nodes.forEach(function (n) { if (n.icon.indexOf('lane:') !== 0) keys['icon/' + n.icon] = 1; }); }
    items.forEach(function (it) {
      if (it.fr) st(it.fr.st, it.k === 'scene');
      if (it.k === 'epTitle') { keys['panels/' + it.ep.thumb] = 1; if (it.ep.villain) keys['villains/' + it.ep.villain + '-base'] = 1; }
      if (it.k === 'end') { var nx = nextEp(D, it.ep); if (nx) keys['panels/' + nx.thumb] = 1; }
      if (it.k === 'diagram') icons(it.did);
      if (it.k === 'cover' || it.k === 'dgCover') keys['img/cover'] = 1;
      if (it.k === 'closing') keys['panels/finale'] = 1;
      if (it.k === 'cast') D.cast.forEach(function (c) { var ch = D.chars[c.id]; keys['chars/' + c.id + '-' + (ch.ex.indexOf('happy') >= 0 ? 'happy' : 'base')] = 1; });
      if (it.k === 'villains') D.villains.forEach(function (v) { keys['villains/' + v.id + '-base'] = 1; });
      if (it.k === 'contents') D.eps.forEach(function (ep) { keys['panels/' + ep.thumb] = 1; });
    });
    return keys;
  }

  /* ============================================================ shared slide pieces */
  function M(X, key) { return X.deck.media[key] || null; }
  function backdrop(X, s, st) {
    if (st.mode === 'panel' && st.panel) s.cover(M(X, 'panels/' + st.panel), 0, 0, SW, SH, { name: '!!panel-' + st.panel, descr: X.D.panelAlt[st.panel] || 'Comic panel' });
    else if (st.bg) s.cover(M(X, 'bg/' + st.bg), 0, 0, SW, SH, { name: '!!bg-' + st.bg, descr: 'Background' + (st.scene ? ': ' + st.scene : '') });
    s.sp({ x: 0, y: 0, w: SW, h: SH, fillXml: grad([[0, '050A1E', 0], [0.55, '050A1E', 0], [1, '050A1E', 0.45]], 0, true), name: '!!vignette' });
  }
  function spriteBox(c, pos, m) {
    var ar = m.w / m.h, bw, bh, bx, by, ih, iw, left;
    if (c.kind === 'villain') { bw = 0.30 * SW; bh = 0.66 * SH; bx = 0.69 * SW; by = 0.34 * SH; ih = bh; iw = ih * ar; return { x: bx + bw / 2 - iw / 2, y: by + bh * 0.96 - ih, w: iw, h: ih, cx: bx + bw / 2 }; }
    if (c.kind === 'pip') { left = { L: 0.10, C: 0.42, R: 0.76, P: 0.76 }[pos]; if (left == null) left = 0.76; bw = 0.16 * SW; bh = 0.28 * SH; bx = left * SW; by = 0.30 * SH; ih = bh; iw = ih * ar; return { x: bx + bw / 2 - iw / 2, y: by, w: iw, h: ih, cx: bx + bw / 2 }; }
    left = { L: 0.02, C: 0.345, R: 0.67 }[pos]; if (left == null) left = 0.67;
    bw = 0.31 * SW; bh = 0.86 * SH; bx = left * SW; by = 0.14 * SH; ih = 1.52 * bh; iw = ih * ar;
    var ix = bx + bw / 2 - iw / 2, iy = by + 0.03 * bh, x0 = Math.max(ix, bx), y0 = Math.max(iy, by), x1 = Math.min(ix + iw, bx + bw), y1 = Math.min(iy + ih, by + bh);
    return { x: x0, y: y0, w: x1 - x0, h: y1 - y0, crop: { l: (x0 - ix) / iw, t: (y0 - iy) / ih, r: (ix + iw - x1) / iw, b: (iy + ih - y1) / ih }, cx: bx + bw / 2 };
  }
  function offDir(c, pos) { return c.kind === 'villain' || c.kind === 'pip' ? 'R' : pos === 'L' ? 'L' : pos === 'C' ? 'B' : 'R'; }
  function drawSprite(X, s, sp, o) {
    var c = X.D.chars[sp.who] || {}, m = M(X, c.dir + '/' + sp.who + '-' + sp.ex); if (!m) return null;
    var g = spriteBox(c, sp.pos, m), dx = 0, dy = 0;
    if (o.off === 'L') dx = -(g.x + g.w + I(0.3)); else if (o.off === 'R') dx = SW - g.x + I(0.3); else if (o.off === 'B') dy = SH - g.y + I(0.3);
    var id = s.pic({ media: m, x: g.x + dx, y: g.y + dy, w: g.w, h: g.h, crop: g.crop, name: '!!spr-' + sp.who, descr: (c.name || sp.who) + (sp.ex && sp.ex !== 'base' ? ' (' + sp.ex + ')' : ''), lum: o.dim ? -15000 : 0, shadow: { blur: I(0.2), dist: I(0.1), dir: 90, c: '000000', a: 0.35 } });
    if (!o.off && X.motion) {
      if (c.kind === 'pip') s.fx.push({ k: 'path', id: id, t: 0, dur: 1300, path: 'M 0 0 L 0 -0.014 E', loop: 4, rev: 1, ease: 1 });
      else if (c.kind === 'villain') s.fx.push({ k: 'path', id: id, t: 0, dur: 1600, path: 'M 0 0 L 0.004 -0.01 E', loop: 3, rev: 1, ease: 1 });
      else if (o.talk) s.fx.push({ k: 'path', id: id, t: 0, dur: 420, path: 'M 0 0 L 0 -0.012 L 0 0 E', ease: 1 });
    }
    g.id = id; return g;
  }
  function sprites(X, s, st, speaker) {
    var spk = null;
    [X.nextIt, X.prevIt].forEach(function (n) {
      if (!n || n.k !== 'scene' || n.fr.st.mode !== 'scene' || ctxOf(n.fr.st) !== ctxOf(st)) return;
      var other = n.fr.st.sprites;
      Object.keys(other).forEach(function (w) { if (!st.sprites[w]) drawSprite(X, s, { who: w, ex: other[w].ex, pos: other[w].pos }, { off: offDir(X.D.chars[w] || {}, other[w].pos) }); });
    });
    Object.keys(st.sprites).forEach(function (w) {
      var g = drawSprite(X, s, { who: w, ex: st.sprites[w].ex, pos: st.sprites[w].pos }, { dim: speaker && w !== speaker, talk: w === speaker });
      if (w === speaker) spk = g;
    });
    return spk;
  }
  function sceneTag(s, text) {
    if (!text) return null;
    var sz = 13, padX = I(0.13), h = I(0.34), w = Math.min(tw('📍 ' + text, sz, true) + 2 * padX + I(0.08), SW * 0.6), x = SW * 0.016, y = SH * 0.022;
    s.sp({ geom: 'roundRect', adj: { adj: 50000 }, x: x, y: y, w: w, h: h, fill: '0A1022', fa: 0.8, line: { c: 'FFFFFF', a: 0.18, w: 9525 }, name: '!!scene', paras: [P('📍 ' + text, { sz: sz, b: 1, color: 'E5EDFF' })], tx: { anchor: 'ctr', ins: [padX, 0, padX, 0], wrap: false } });
    return { x: x, y: y, w: w, h: h };
  }
  var TAIL = { L: 0.28, C: 0.46, R: 0.62, V: 0.70, P: 0.70 };
  function typing(X, s, id, text) { if (X.motion) s.fx.push({ k: 'appear', id: id, t: 0, lt: clamp(1800 / Math.max(1, text.length), 8, 30) }); }
  function bubble(X, s, st, b, spk, tag) {
    var c = X.D.chars[b.who] || { name: b.who, color: '#334155' }, mode = b.mode || 'say', vill = c.kind === 'villain';
    var pos = b.pos || (st.mode === 'scene' && st.sprites[b.who] ? st.sprites[b.who].pos : (c.side || 'L'));
    var fs = 18, ns = 12, lh = fs * 1.22, bold = mode === 'shout', fill = vill ? '2A1F3D' : mode === 'shout' ? 'FFF7ED' : mode === 'think' ? 'F8FBFF' : 'FFFFFF';
    var padX = I(0.22), padT = I(0.13), padB = I(0.17), nameH = ns * 1.35 * PT, gap = I(0.03);
    var inner = clamp(Math.max(tw(b.v, fs, bold), tw(c.name.toUpperCase(), ns, true) + ns * 0.8 * c.name.length * 100), SW * 0.18 - 2 * padX, SW * 0.47 - 2 * padX);
    var lines = lineCount(b.v, fs, FONT, inner, bold), w = inner + 2 * padX, h = padT + nameH + gap + lines * lh * PT + padB;
    var x = pos === 'C' ? SW * 0.32 : (pos === 'R' || pos === 'P') ? SW * 0.97 - w : pos === 'V' ? SW * 0.98 - w : SW * 0.03, y = SH * 0.032;
    if (tag && x < tag.x + tag.w + I(0.08)) y = tag.y + tag.h + I(0.1);
    var tipX = clamp(spk ? spk.cx : x + w * (TAIL[pos] || 0.3), x + w * 0.12, x + w * 0.88), tipY = y + h + I(0.34);
    s.sp({ geom: 'wedgeRoundRectCallout', adj: { adj1: (tipX - (x + w / 2)) / w * 100000, adj2: (tipY - (y + h / 2)) / h * 100000, adj3: Math.min(16667, I(0.24) / Math.min(w, h) * 100000) },
      x: x, y: y, w: w, h: h, fill: fill,
      line: { c: vill ? 'F3E8FF' : mode === 'shout' ? 'B91C1C' : '111111', w: mode === 'shout' ? PT * 3 : PT * 2.25, dash: mode === 'think' ? 'dash' : null },
      shadow: { blur: 0, dist: I(0.07), dir: 51, c: '000000', a: 0.28 }, name: '!!bubble', descr: 'Speech bubble' });
    s.text(x + padX, y + padT, inner, nameH, [P(c.name, { sz: ns, b: 1, color: vill ? hx(c.color) : inkOn(c.color, fill), caps: 1, spc: 90, font: FONT_D })], { name: '!!bubble-name' });
    var id = s.text(x + padX, y + padT + nameH + gap, inner, lines * lh * PT + I(0.03), [P(b.v, { sz: fs, b: bold, i: mode === 'think', color: vill ? 'F3E8FF' : '111111', sp: lh })], { name: 'Speech ' + (++X.seq) });
    typing(X, s, id, b.v);
  }
  function caption(X, s, text) {
    var fs = 17, lh = fs * 1.25, padX = I(0.17), padY = I(0.11);
    var inner = Math.min(tw(text, fs, true), SW * 0.5 - 2 * padX), lines = lineCount(text, fs, FONT, inner, true), w = inner + 2 * padX, h = lines * lh * PT + 2 * padY, x = SW * 0.016, y = SH * 0.09;
    s.sp({ geom: 'roundRect', adj: { adj: Math.min(16667, I(0.07) / Math.min(w, h) * 100000) }, x: x, y: y, w: w, h: h, fill: 'FFF3C4', line: { c: '2B2210', w: PT * 1.75 }, shadow: { blur: 0, dist: I(0.055), dir: 45, c: '000000', a: 0.35 }, name: '!!cap', descr: 'Narration box' });
    var id = s.text(x + padX, y + padY, inner, lines * lh * PT + I(0.03), [P(text, { sz: fs, b: 1, color: '2B2210', sp: lh })], { name: 'Narration ' + (++X.seq) });
    typing(X, s, id, text);
  }
  function sfxText(X, s, text) {
    var fs = 60, tw0 = tw(text, fs, true, FONT_D), w = tw0 + I(0.5), h = fs * 1.4 * PT, cx = SW / 2, cy = SH * 0.42;
    var bw = tw0 * 1.55 + I(1.2), bh = h * 2.3;
    var bid = s.sp({ geom: 'irregularSeal1', x: cx - bw / 2, y: cy - bh / 2, w: bw, h: bh, rot: -6, fill: 'B91C1C', line: { c: 'FFF7ED', w: PT * 3 }, shadow: { blur: 0, dist: I(0.09), dir: 50, c: '000000', a: 0.35 }, name: 'Sound burst' });
    var id = s.text(cx - w / 2, cy - h / 2, w, h, [P(text, { sz: fs, b: 1, color: 'FFD84D', font: FONT_D, algn: 'ctr', outline: { w: PT * 2, c: '2A1300' } })], { name: 'Sound effect', anchor: 'ctr', rot: -8, wrap: false });
    if (X.motion) s.fx.push({ k: 'zoom', id: bid, t: 0, dur: 300 }, { k: 'zoom', id: id, t: 0, dur: 380 });
  }
  function progressBar(X, s, frac) { s.sp({ x: 0, y: SH - I(0.06), w: Math.max(I(0.05), SW * frac), h: I(0.06), fill: X.accent, name: '!!progress', descr: 'Episode progress' }); }
  function boardFrame(X, s, o) {
    o = o || {};
    var w = o.w || SW - I(1.0), h = o.h || SH - I(0.78), x = (SW - w) / 2, y = o.y != null ? o.y : I(0.42), r = o.r || I(0.2);
    s.sp({ geom: 'roundRect', adj: { adj: r / Math.min(w, h) * 100000 }, x: x, y: y, w: w, h: h, fill: X.accent, name: '!!board-edge', shadow: { blur: I(0.35), dist: I(0.12), dir: 90, c: '030818', a: 0.45 } });
    s.sp({ geom: 'roundRect', adj: { adj: r / Math.min(w, h - I(0.09)) * 100000 }, x: x, y: y + I(0.09), w: w, h: h - I(0.09), fill: PAPER, name: '!!board' });
    return { x: x + I(0.45), y: y + I(0.09) + I(0.3), w: w - I(0.9), h: h - I(0.09) - I(0.3) - I(0.26) };
  }
  function srcRuns(X, keys) {
    var runs = [{ t: (keys.length > 1 ? 'Sources' : 'Source') + ' on Microsoft Learn: ', b: 1, color: '0F172A' }];
    keys.forEach(function (k, i) { var S = X.D.sources[k]; if (!S) return; if (i) runs.push({ t: ' · ' }); runs.push({ t: S[0], link: S[1], color: LINK }); });
    return runs;
  }
  function srcFooter(X, s, keys, x, yBottom, w, rule) { // returns the height used above yBottom
    if (!keys || !keys.length) return 0;
    var fs = 12, lh = 15, runs = srcRuns(X, keys), n = lineCount(runs, fs, FONT, w), h = n * lh * PT;
    s.text(x, yBottom - h, w, h, [P(runs, { sz: fs, color: '475569', sp: lh })], { name: 'Sources' });
    if (rule) s.path([{ t: 'M', p: { x: x, y: yBottom - h - I(0.1) } }, { t: 'L', p: { x: x + w, y: yBottom - h - I(0.1) } }], { line: { c: 'D6CFBD', w: 9525, dash: 'dash' }, name: 'Rule' });
    return h + I(0.2);
  }
  function srcNotes(X, keys) { return (keys || []).map(function (k) { var S = X.D.sources[k]; return S ? '• ' + S[0] + (S[2] ? ' (updated ' + S[2] + ')' : '') + ': ' + S[1] : ''; }).filter(Boolean).join('\n'); }

  /* ============================================================ lesson cards */
  function blockH(X, s, bl, x, y, w, k, draw) {
    var fs = 16 * k, lh = fs * 1.28, n, i, h;
    if (bl.k === 'p') { n = lineCount(bl.r, fs, FONT, w); if (draw) s.text(x, y, w, n * lh * PT, [P(bl.r, { sz: fs, color: BODY, sp: lh, linkColor: LINK })]); return n * lh * PT; }
    if (bl.k === 'h3') { var f3 = 18 * k, l3 = f3 * 1.22; n = lineCount(bl.r, f3, FONT_D, w, true); if (draw) s.text(x, y, w, n * l3 * PT, [P(bl.r, { sz: f3, b: 1, color: '13204A', font: FONT_D, sp: l3 })]); return n * l3 * PT; }
    if (bl.k === 'ul') {
      var ind = I(0.24) * k, tot = 0, paras = [];
      bl.items.forEach(function (it, j) { var nn = lineCount(it, fs, FONT, w - ind); tot += nn * lh * PT + (j ? 4 * k * PT : 0); paras.push(P(it, { sz: fs, color: BODY, sp: lh, bef: j ? 4 * k : 0, bullet: { ch: '•', c: X.accent }, marL: ind, ind: -ind, linkColor: LINK })); });
      if (draw) s.text(x, y, w, tot, paras);
      return tot;
    }
    if (bl.k === 'note') {
      var fn = 15 * k, ln = fn * 1.28, pad = I(0.15) * k; n = lineCount(bl.r, fn, FONT, w - 2 * pad); h = n * ln * PT + 2 * pad;
      if (draw) { s.sp({ geom: 'roundRect', adj: { adj: Math.min(30000, I(0.12) / h * 100000) }, x: x, y: y, w: w, h: h, fill: 'E0F2FE', name: 'Note' }); s.text(x + pad, y + pad, w - 2 * pad, n * ln * PT, [P(bl.r, { sz: fn, color: '0C3A5B', sp: ln, linkColor: LINK })]); }
      return h;
    }
    if (bl.k === 'steps') {
      var items = bl.items, cols = items.length >= 6 ? 2 : 1, gc = I(0.2), cw = (w - (cols - 1) * gc) / cols;
      var ft = 17 * k, fd = 15 * k, lt = ft * 1.22, ld = fd * 1.26, nz = Math.max(12, 13 * k), cd = Math.max(I(0.4) * k, nz * 30000 + 30000), pad2 = I(0.1) * k, tx = cd + I(0.16) * k, tw2 = cw - tx - 2 * pad2;
      var per = Math.ceil(items.length / cols), gr = I(0.08) * k, colH = [0, 0];
      items.forEach(function (it, j) {
        var c = Math.floor(j / per), nt = lineCount(it.t, ft, FONT, tw2, true), nd = lineCount(it.d, fd, FONT, tw2), th = nt * lt * PT + nd * ld * PT, hh = Math.max(cd, th) + 2 * pad2, xx = x + c * (cw + gc), yy = y + colH[c];
        if (draw) {
          s.sp({ geom: 'roundRect', adj: { adj: Math.min(30000, I(0.14) / hh * 100000) }, x: xx, y: yy, w: cw, h: hh, fill: 'FFFFFF', line: { c: 'E6E1D3', w: 9525 }, name: 'Step ' + it.no });
          s.sp({ geom: 'ellipse', x: xx + pad2, y: yy + (hh - cd) / 2, w: cd, h: cd, fill: deep(X.accent), paras: [P(String(it.no), { sz: nz, b: 1, color: 'FFFFFF', font: FONT_D, algn: 'ctr' })], tx: { anchor: 'ctr' }, name: 'Step number ' + it.no });
          s.text(xx + pad2 + tx, yy + pad2, tw2, hh - 2 * pad2, [P(it.t, { sz: ft, b: 1, color: INK, sp: lt }), P(it.d, { sz: fd, color: BODY, sp: ld })], { anchor: 'ctr' });
        }
        colH[c] += hh + gr;
      });
      return Math.max(colH[0], colH[1]) - gr;
    }
    if (bl.k === 'tiles') {
      var nT = bl.items.length, perRow = nT <= 5 ? nT : 3, rows = Math.ceil(nT / perRow), gt = I(0.16) * k, tW = (w - (perRow - 1) * gt) / perRow;
      var fi = 24 * k, ftt = 16 * k, fdd = 14 * k, p2 = I(0.14) * k;
      var hs = bl.items.map(function (it) { return 2 * p2 + fi * 1.2 * PT + 4 * k * PT + lineCount(it.t, ftt, FONT, tW - 2 * p2, true) * ftt * 1.25 * PT + lineCount(it.d, fdd, FONT, tW - 2 * p2) * fdd * 1.3 * PT; });
      var tot2 = 0;
      for (var r = 0; r < rows; r++) {
        var rh = Math.max.apply(null, hs.slice(r * perRow, r * perRow + perRow));
        if (draw) bl.items.slice(r * perRow, r * perRow + perRow).forEach(function (it, j) {
          var xx = x + j * (tW + gt), yy = y + tot2;
          s.sp({ geom: 'roundRect', adj: { adj: Math.min(16667, I(0.14) / Math.min(tW, rh) * 100000) }, x: xx, y: yy, w: tW, h: rh, fill: 'FFFFFF', line: { c: 'E6E1D3', w: 9525 }, name: 'Tile' });
          s.text(xx + p2, yy + p2, tW - 2 * p2, rh - 2 * p2, [P(it.ic, { sz: fi, sp: fi * 1.2 }), P(it.t, { sz: ftt, b: 1, color: INK, sp: ftt * 1.25, bef: 4 * k }), P(it.d, { sz: fdd, color: BODY, sp: fdd * 1.3 })]);
        });
        tot2 += rh + (r < rows - 1 ? gt : 0);
      }
      return tot2;
    }
    if (bl.k === 'duo') {
      var gd = I(0.22) * k, cw2 = (w - gd) / 2, pd = I(0.16) * k;
      var hmax = Math.max.apply(null, bl.cols.map(function (col) { return stackH(X, s, col, 0, 0, cw2 - 2 * pd, k, false) + 2 * pd; }));
      if (draw) bl.cols.forEach(function (col, j) { var xx = x + j * (cw2 + gd); s.sp({ geom: 'roundRect', adj: { adj: Math.min(16667, I(0.14) / Math.min(cw2, hmax) * 100000) }, x: xx, y: y, w: cw2, h: hmax, fill: 'FFFFFF', line: { c: 'E6E1D3', w: 9525 }, name: 'Column' }); stackH(X, s, col, xx + pd, y + pd, cw2 - 2 * pd, k, true); });
      return hmax;
    }
    if (bl.k === 'table') {
      var ncol = bl.head.length, ft2 = 14 * k, lt2 = ft2 * 1.26, mx = I(0.09) * k, my = I(0.05) * k;
      var wts = bl.head.map(function (hd, ci) { var m = tw(plain(hd), ft2, true); bl.rows.forEach(function (row) { m = Math.max(m, Math.min(tw(plain(row[ci] || ''), ft2), I(4.2))); }); return clamp(m, I(0.9), I(4.2)); });
      var nat = wts.map(function (v) { return v + 2 * mx + I(0.05); }), sum = nat.reduce(function (a, b) { return a + b; }, 0), cols2;
      if (sum >= w) cols2 = nat.map(function (v) { return v / sum * w; });
      else { // natural widths, the spare room goes to the columns that carry long text
        var grow = nat.map(function (v, ci) { return wts[ci] >= I(1.6) ? v : 0; }), gs = grow.reduce(function (a, b) { return a + b; }, 0);
        if (!gs) { grow = nat.map(function (v, ci) { return ci === nat.length - 1 ? 1 : 0; }); gs = 1; }
        cols2 = nat.map(function (v, ci) { return v + (w - sum) * grow[ci] / gs; });
      }
      var trs = [bl.head].concat(bl.rows).map(function (row, ri) {
        var cells = []; for (i = 0; i < ncol; i++) cells.push(row[i] || [{ t: '' }]);
        return { h: Math.max.apply(null, cells.map(function (cell, ci) { return lineCount(cell, ft2, FONT, cols2[ci] - 2 * mx, ri === 0) * lt2 * PT; })) + 2 * my + I(0.02),
          cells: cells.map(function (cell) { return { paras: [P(cell, { sz: ft2, b: ri === 0 ? 1 : undefined, color: ri === 0 ? 'FFFFFF' : '1E293B', sp: lt2, linkColor: LINK })], m: [mx, my, mx, my], fill: ri === 0 ? '13204A' : (ri % 2 === 0 ? 'FAF7EF' : 'FFFFFF'), top: ri > 1 ? { c: 'ECE6D6', w: 9525 } : null }; }) };
      });
      h = trs.reduce(function (a, r2) { return a + r2.h; }, 0);
      if (draw) s.table({ x: x, y: y, w: w, cols: cols2, rows: trs, name: 'Table' });
      return h;
    }
    return 0;
  }
  function stackH(X, s, blocks, x, y, w, k, draw) { var t = 0; blocks.forEach(function (bl, i) { if (i) t += I(0.14) * k; t += blockH(X, s, bl, x, y + t, w, k, draw); }); return t; }
  function cardHead(X, s, card, x, y, w, draw, cont) {
    var t = 0;
    if (card.kicker && card.kicker.length) { if (draw) s.text(x, y, w, 12 * 1.35 * PT, [P(card.kicker, { sz: 12, b: 1, color: MUTED, caps: 1, spc: 140 })], { name: '!!card-kick' }); t += 12 * 1.35 * PT + I(0.04); }
    var title = cont ? card.title.concat([{ t: ' (continued)' }]) : card.title, tf = 28, tl = tf * 1.15, n = lineCount(title, tf, FONT_D, w, true);
    if (draw) s.text(x, y + t, w, n * tl * PT, [P(title, { sz: tf, b: 1, color: INK, font: FONT_D, sp: tl })], { name: '!!card-title' });
    return t + n * tl * PT + I(0.16);
  }
  function cardPages(X, card, box) {
    var head = cardHead(X, null, card, 0, 0, box.w, false), src = card.srcs.length ? lineCount(srcRuns(X, card.srcs), 12, FONT, box.w) * 15 * PT + I(0.2) : 0, avail = box.h - head - src;
    var ks = [1, 0.94, 0.88, 0.82, 0.76, 0.7];
    for (var i = 0; i < ks.length; i++) if (stackH(X, null, card.blocks, 0, 0, box.w, ks[i], false) <= avail) return [{ blocks: card.blocks, k: ks[i] }];
    var pages = [{ blocks: [], k: 0.82 }], used = 0;
    card.blocks.forEach(function (bl) {
      var hh = blockH(X, null, bl, 0, 0, box.w, 0.82, false) + I(0.14) * 0.82, pg = pages[pages.length - 1];
      if (pg.blocks.length && used + hh > avail) { pages.push({ blocks: [bl], k: 0.82, cont: true }); used = hh; } else { pg.blocks.push(bl); used += hh; }
    });
    return pages;
  }
  function cardText(card) {
    var out = [];
    function blk(b) {
      if (b.k === 'p' || b.k === 'h3' || b.k === 'note') out.push((b.k === 'note' ? 'Note: ' : '') + plain(b.r));
      else if (b.k === 'ul') b.items.forEach(function (it) { out.push('• ' + plain(it)); });
      else if (b.k === 'steps') b.items.forEach(function (it) { out.push(it.no + '. ' + plain(it.t) + ': ' + plain(it.d)); });
      else if (b.k === 'tiles') b.items.forEach(function (it) { out.push('• ' + plain(it.t) + ': ' + plain(it.d)); });
      else if (b.k === 'duo') b.cols.forEach(function (c) { c.forEach(blk); });
      else if (b.k === 'table') [b.head].concat(b.rows).forEach(function (r) { out.push(r.map(plain).join(' | ')); });
    }
    card.blocks.forEach(blk);
    return out.join('\n');
  }

  /* ============================================================ story slides */
  function storySlide(X, it, speaker, bare) {
    var st = it.fr.st, s = X.deck.slide({ bg: STAGE, ctx: ctxOf(st) });
    backdrop(X, s, st);
    if (bare) return { s: s };
    var spk = st.mode === 'scene' ? sprites(X, s, st, speaker) : null;
    return { s: s, spk: spk, tag: sceneTag(s, st.scene) };
  }
  function sceneSlide(X, it) {
    var b = it.fr.beat, r = storySlide(X, it, b.t === 'say' ? b.who : null), s = r.s, c = X.D.chars[b.who] || {};
    if (b.t === 'say') { bubble(X, s, it.fr.st, b, r.spk, r.tag); s.notes = (c.name || b.who) + (b.mode === 'think' ? ' (thinking)' : b.mode === 'shout' ? ' (shouting)' : '') + ': ' + b.v; s.adv = 1300 + b.v.length * 40; }
    else if (b.t === 'cap') { caption(X, s, b.v); s.notes = 'Narration: ' + b.v; s.adv = 1300 + b.v.length * 36; }
    else { sfxText(X, s, b.v); s.notes = 'Sound effect: ' + b.v; s.adv = 1600; }
    if (it.fr.st.scene && (!X.prevIt || X.prevIt.fr.st.scene !== it.fr.st.scene)) s.notes = 'Scene: ' + it.fr.st.scene + '\n' + s.notes;
    progressBar(X, s, it.prog);
  }
  function boardBox() { var w = SW - I(1.0), h = SH - I(0.78); return { x: (SW - w) / 2 + I(0.45), y: I(0.42) + I(0.39), w: w - I(0.9), h: h - I(0.65) }; }
  function cardSlides(X, it) {
    var card = it.fr.beat.card, pages = cardPages(X, card, boardBox());
    pages.forEach(function (pg, pi) {
      var s = storySlide(X, it, null, true).s;
      s.sp({ x: 0, y: 0, w: SW, h: SH, fill: NIGHT, fa: 0.74, name: '!!scrim' });
      var box = boardFrame(X, s), y = box.y;
      y += cardHead(X, s, card, box.x, y, box.w, true, pg.cont);
      stackH(X, s, pg.blocks, box.x, y, box.w, pg.k, true);
      if (pi === pages.length - 1) srcFooter(X, s, card.srcs, box.x, box.y + box.h, box.w, true);
      progressBar(X, s, it.prog);
      var txt = cardText({ blocks: pg.blocks });
      s.notes = 'Lesson card: ' + plain(card.title) + (pg.cont ? ' (continued)' : '') + '\n' + txt + (pi === pages.length - 1 && card.srcs.length ? '\n\nSources on Microsoft Learn:\n' + srcNotes(X, card.srcs) : '');
      s.adv = Math.min(26000, 5000 + txt.length * 22);
    });
  }
  function quizSlide(X, it, reveal) {
    var b = it.fr.beat, s = storySlide(X, it, null, true).s, L = 'ABCDEF', ok = 0;
    b.opts.forEach(function (o, i) { if (o.ok) ok = i; });
    s.sp({ x: 0, y: 0, w: SW, h: SH, fill: NIGHT, fa: 0.74, name: '!!scrim' });
    var box = boardFrame(X, s), y = box.y;
    s.text(box.x, y, box.w, 12 * 1.35 * PT, [P('Checkpoint · ' + X.ep.n, { sz: 12, b: 1, color: MUTED, caps: 1, spc: 140 })], { name: '!!q-kick' }); y += 12 * 1.35 * PT + I(0.06);
    var qf = 26, ql = qf * 1.18, qn = lineCount(b.q, qf, FONT, box.w, true);
    s.text(box.x, y, box.w, qn * ql * PT, [P(b.q, { sz: qf, b: 1, color: INK, sp: ql })], { name: '!!q-text' }); y += qn * ql * PT + I(0.22);
    var of = 19, ol = of * 1.24, cd = I(0.46), pad = I(0.12), pill = I(1.45), tW = box.w - 2 * pad - cd - I(0.2) - pill;
    b.opts.forEach(function (o, i) {
      var n = lineCount(o.t, of, FONT, tW), hh = Math.max(cd, n * ol * PT) + 2 * pad, right = reveal && i === ok;
      s.sp({ geom: 'roundRect', adj: { adj: Math.min(30000, I(0.12) / hh * 100000) }, x: box.x, y: y, w: box.w, h: hh, fill: right ? 'ECFDF5' : 'FFFFFF', line: { c: right ? '16A34A' : 'D9D2C1', w: PT * 2 }, name: '!!opt-' + i });
      s.sp({ geom: 'ellipse', x: box.x + pad, y: y + (hh - cd) / 2, w: cd, h: cd, fill: right ? '15803D' : 'E2E8F0', paras: [P(L[i], { sz: 14, b: 1, color: right ? 'FFFFFF' : '0F172A', algn: 'ctr' })], tx: { anchor: 'ctr' }, name: '!!optl-' + i });
      s.text(box.x + pad + cd + I(0.2), y + pad, tW, hh - 2 * pad, [P(o.t, { sz: of, color: reveal && !right ? MUTED : '0F172A', sp: ol })], { name: '!!optt-' + i, anchor: 'ctr' });
      if (right) s.sp({ geom: 'roundRect', adj: { adj: 50000 }, x: box.x + box.w - pad - pill + I(0.1), y: y + (hh - I(0.36)) / 2, w: pill - I(0.1), h: I(0.36), fill: '15803D', paras: [P('✔ Correct', { sz: 13, b: 1, color: 'FFFFFF', algn: 'ctr' })], tx: { anchor: 'ctr', wrap: false }, name: 'Correct answer' });
      y += hh + I(0.12);
    });
    if (reveal) {
      var fw = box.w, ff = 17, fl = ff * 1.26, padF = I(0.16), runs = [{ t: 'Answer: ' + L[ok] + '. ', b: 1, color: '14532D' }, { t: b.ex }];
      var fn = lineCount(runs, ff, FONT, fw - 2 * padF), fh = fn * fl * PT + 2 * padF;
      y += I(0.06);
      var fid = s.sp({ geom: 'roundRect', adj: { adj: Math.min(30000, I(0.12) / fh * 100000) }, x: box.x, y: y, w: fw, h: fh, fill: 'DCFCE7', name: 'Answer box' });
      var tid = s.text(box.x + padF, y + padF, fw - 2 * padF, fn * fl * PT, [P(runs, { sz: ff, color: '1E293B', sp: fl })], { name: 'Answer text' });
      if (X.motion) s.fx.push({ k: 'fade', id: fid, t: 0, dur: 400 }, { k: 'fade', id: tid, t: 0, dur: 400 });
      srcFooter(X, s, b.srcs, box.x, box.y + box.h, box.w, true);
      s.notes = 'Answer: ' + L[ok] + '. ' + b.opts[ok].t + '\n' + b.ex + (b.srcs.length ? '\n\nSources on Microsoft Learn:\n' + srcNotes(X, b.srcs) : '');
      s.adv = 3200 + b.ex.length * 32;
    } else {
      s.text(box.x, box.y + box.h - I(0.32), box.w, I(0.32), [P('Think it over, then click to reveal the answer.', { sz: 15, i: 1, color: MUTED })], { name: 'Hint' });
      s.notes = 'Checkpoint question: ' + b.q + '\n' + b.opts.map(function (o, i) { return L[i] + '. ' + o.t; }).join('\n') + '\nAsk the room, then click to reveal the answer.';
      s.adv = 6000;
    }
    progressBar(X, s, it.prog);
  }
  function endSlide(X, it) {
    var s = storySlide(X, it, null, true).s, ep = X.ep, nx = nextEp(X.D, ep);
    s.sp({ x: 0, y: 0, w: SW, h: SH, fill: NIGHT, fa: 0.74, name: '!!scrim' });
    var box = boardFrame(X, s, { w: I(9.2), h: I(5.0), y: I(1.25) }), cx = box.x, cw = box.w, y = box.y + I(0.05);
    s.text(cx, y, cw, I(0.3), [P(ep.n + ' complete', { sz: 13, b: 1, color: MUTED, caps: 1, spc: 140, algn: 'ctr' })], { name: 'End kicker' }); y += I(0.38);
    var tn = lineCount('✔ ' + ep.title, 38, FONT_D, cw, true);
    s.text(cx, y, cw, tn * 44 * PT, [P('✔ ' + ep.title, { sz: 38, b: 1, color: INK, font: FONT_D, algn: 'ctr', sp: 44 })], { name: 'End title' }); y += tn * 44 * PT + I(0.35);
    if (nx) {
      var tw3 = I(2.7), th3 = I(1.8), tx3 = cx + I(0.3);
      s.cover(M(X, 'panels/' + nx.thumb), tx3, y, tw3, th3, { name: 'Next episode art', descr: X.D.panelAlt[nx.thumb] || '' });
      var lx = tx3 + tw3 + I(0.35), lw = cx + cw - lx - I(0.2);
      s.text(lx, y + I(0.1), lw, I(0.3), [P('Up next · ' + nx.n, { sz: 13, b: 1, color: MUTED, caps: 1, spc: 120 })], { name: 'Up next' });
      s.text(lx, y + I(0.45), lw, I(0.9), [P(nx.title, { sz: 26, b: 1, color: INK, font: FONT_D, sp: 30 })], { name: 'Next title' });
      s.text(lx, y + I(1.3), lw, I(0.5), [P(it.whole ? [{ t: 'Read it ▶', jump: nx.id, color: LINK, b: 1 }] : [{ t: nx.sub, color: BODY }], { sz: 15, sp: 19 })], { name: 'Next link' });
    } else s.text(cx, y + I(0.3), cw, I(0.6), [P('🎉 You finished the whole story!', { sz: 26, b: 1, color: INK, algn: 'ctr' })], { name: 'The end' });
    s.notes = ep.n + ' complete: ' + ep.title + '.' + (nx ? '\nUp next: ' + nx.n + ', ' + nx.title + '. ' + nx.sub : '\nThat was the last episode.');
    s.adv = 4500; progressBar(X, s, 1);
  }

  /* ============================================================ title, cover, cast, contents, sources */
  function hero(s, m, side, x0, x1, o) { // picture between x0 and x1, a fade on the inner edge
    if (!m) return;
    s.cover(m, x0, 0, x1 - x0, SH, o);
    var fw = I(3.4);
    if (side === 'right') s.sp({ x: x0 - 1, y: 0, w: fw, h: SH, fillXml: grad([[0, NIGHT, 1], [0.5, NIGHT, 0.85], [1, NIGHT, 0]], 0), name: 'Fade' });
    else s.sp({ x: x1 - fw + 1, y: 0, w: fw, h: SH, fillXml: grad([[0, NIGHT, 0], [0.5, NIGHT, 0.85], [1, NIGHT, 1]], 0), name: 'Fade' });
  }
  function pill(s, x, y, text, fill, color, o) {
    o = o || {}; var sz = o.sz || 12, h = o.h || I(0.34), w = tw(text.toUpperCase(), sz, true) + sz * 0.9 * text.length * 100 + I(0.32);
    s.sp({ geom: 'roundRect', adj: { adj: 50000 }, x: x, y: y, w: w, h: h, fill: fill, fa: o.fa, paras: [P(text, { sz: sz, b: 1, color: color, caps: 1, spc: 90, algn: 'ctr' })], tx: { anchor: 'ctr', wrap: false }, name: o.name || 'Pill' });
    return w;
  }
  function epTitleSlide(X, ep) {
    var s = X.deck.slide({ bg: NIGHT, ctx: 'title:' + ep.id }); X.deck.anchors[ep.id] = s;
    hero(s, M(X, 'panels/' + ep.thumb), 'right', I(5.0), SW, { name: 'Episode art', descr: X.D.panelAlt[ep.thumb] || '' });
    var x = I(0.7), w = I(6.6), topic = (ep.phase && ep.phase !== ep.n ? ep.phase + ' · ' : '') + ep.dur;
    var v = ep.villain && X.D.villains.filter(function (vv) { return vv.id === ep.villain; })[0];
    var tn = lineCount(ep.title, 46, FONT_D, w, true), sn = lineCount(ep.sub, 20, FONT, w);
    var total = I(0.38) + I(0.22) + tn * 52 * PT + I(0.15) + sn * 26 * PT + (v ? I(0.55) + I(1.5) : 0);
    var y = Math.max(I(1.15), (SH - total) / 2);
    s.text(x, I(0.5), w, I(0.3), [P(X.D.title, { sz: 12, b: 1, color: '9FB0D9', caps: 1, spc: 160 })], { name: 'Brand' });
    var pw = pill(s, x, y, ep.n, deep(X.accent), 'FFFFFF', { sz: 13, h: I(0.38), name: 'Episode number' });
    s.text(x + pw + I(0.18), y, w - pw, I(0.38), [P(topic, { sz: 14, b: 1, color: 'CBD5F5', caps: 1, spc: 80 })], { name: 'Topic', anchor: 'ctr' });
    y += I(0.38) + I(0.22);
    s.text(x, y, w, tn * 52 * PT, [P(ep.title, { sz: 46, b: 1, color: 'FFFFFF', font: FONT_D, sp: 52 })], { name: 'Title' }); y += tn * 52 * PT + I(0.15);
    s.text(x, y, w, sn * 26 * PT, [P(ep.sub, { sz: 20, color: 'DCE3F7', sp: 26 })], { name: 'Subtitle' }); y += sn * 26 * PT + I(0.55);
    if (v) {
      s.contain(M(X, 'villains/' + v.id + '-base'), x, y, I(1.2), I(1.5), { name: 'Villain', descr: v.name });
      s.text(x + I(1.4), y + I(0.12), I(4.9), I(1.3), [P([{ t: 'Villain of the episode: ', color: '9FB0D9' }, { t: v.name, color: hx((X.D.chars[v.id] || {}).color || '#ffffff') }], { sz: 14, b: 1, sp: 18 }), P(v.what, { sz: 14, color: 'DCE3F7', sp: 18, bef: 3 }), P('Defeated by: ' + v.beat, { sz: 13, i: 1, color: '9FB0D9', sp: 17, bef: 3 })], { name: 'Villain text' });
    }
    s.notes = ep.n + ' · ' + ep.title + '\n' + ep.sub + '\nTopic: ' + topic + '.' + (v ? '\nVillain: ' + v.name + '. ' + v.what : '') +
      '\n\nHow to present: click to advance. Speech types itself, characters move with Morph, and each diagram step plays its traffic hop by hop. Every line of dialogue is in these notes.';
    s.adv = 5000;
  }
  function coverSlide(X, sub, lead, stats) {
    var D = X.D, s = X.deck.slide({ bg: NIGHT, ctx: 'cover' });
    hero(s, M(X, 'img/cover'), 'right', I(5.6), SW, { name: 'Cover art', descr: D.panelAlt.cover || '' });
    var x = I(0.7), w = I(6.5), y = I(1.15);
    var cw = pill(s, x, y, 'Cloud Adoption Framework', '0EA5E9', '7DD3FC', { fa: 0.25, name: 'CAF' });
    pill(s, x + cw + I(0.15), y, 'Well-Architected Framework', '8B5CF6', 'C4B5FD', { fa: 0.3, name: 'WAF' });
    y += I(0.6);
    s.text(x, y, I(7.2), I(1.0), [P([{ t: 'Land It ' }, { t: 'Well', color: 'FBBF24' }, { t: ', Dewi!' }], { sz: 58, b: 1, color: 'FFFFFF', font: FONT_D, sp: 64 })], { name: 'Title' }); y += I(1.1);
    if (sub) { s.text(x, y, w, I(0.5), [P(sub, { sz: 24, b: 1, color: 'E7ECFF', font: FONT_D, sp: 29 })], { name: 'Subtitle' }); y += I(0.6); }
    var ln = lineCount(lead, 18, FONT, w);
    s.text(x, y, w, ln * 25 * PT, [P(lead, { sz: 18, color: 'E7ECFF', sp: 25 })], { name: 'Lead' }); y += ln * 25 * PT + I(0.25);
    var stn = lineCount(stats, 14, FONT, w);
    s.text(x, y, w, stn * 19 * PT, [P(stats, { sz: 14, color: '9FB0D9', sp: 19 })], { name: 'Stats' }); y += stn * 19 * PT + I(0.14);
    s.text(x, y, w, I(0.3), [P([{ t: D.site.replace(/^https?:\/\//, '').replace(/\/$/, ''), link: D.site, color: '7DD3FC', b: 1 }], { sz: 14 })], { name: 'Website' });
    s.notes = D.title + ' · ' + D.tag + '.\nRead it online: ' + D.site + '\n\nHow to present: click to advance. Speech types itself, characters move with Morph transitions, and each diagram step plays its traffic hop by hop. Speaker notes hold every line, lesson card and source.\n\n' + D.disclaimer;
    s.adv = 6000;
    return s;
  }
  function castSlide(X) {
    var D = X.D, s = X.deck.slide({ bg: NIGHT, ctx: 'cast' });
    s.text(I(0.6), I(0.4), I(12), I(0.3), [P('The cast', { sz: 12, b: 1, color: '9FB0D9', caps: 1, spc: 160 })], { name: 'Kicker' });
    s.text(I(0.6), I(0.7), I(12), I(0.65), [P('Meet the crew', { sz: 34, b: 1, color: 'FFFFFF', font: FONT_D })], { name: 'Title' });
    s.text(I(0.6), I(1.35), I(12), I(0.5), [P('The harbor and its airships are our picture of Azure, so the ideas are easier to remember. The lesson cards and diagrams use the real Azure names.', { sz: 15, color: '9FB0D9', sp: 20 })], { name: 'Intro' });
    var n = D.cast.length, gap = I(0.2), cw = (SW - I(1.2) - (n - 1) * gap) / n, y = I(2.1), ch = SH - y - I(0.4);
    D.cast.forEach(function (c, i) {
      var x = I(0.6) + i * (cw + gap), chd = D.chars[c.id], ex = chd.ex.indexOf('happy') >= 0 ? 'happy' : 'base', m = M(X, 'chars/' + c.id + '-' + ex);
      s.sp({ geom: 'roundRect', adj: { adj: 7000 }, x: x, y: y, w: cw, h: ch, fill: '101A36', line: { c: '22305A', w: 9525 }, name: 'Card ' + c.name });
      if (m) {
        if (chd.kind === 'pip') s.contain(m, x + I(0.2), y + I(0.25), cw - I(0.4), I(2.2), { name: c.name, descr: c.name });
        else { var vis = 0.5, ih = I(2.55), iw = ih * (m.w / (m.h * vis)); s.pic({ media: m, x: x + (cw - iw) / 2, y: y + I(0.15), w: iw, h: ih, crop: { l: 0, t: 0, r: 0, b: 1 - vis }, name: c.name, descr: c.name }); }
      }
      s.text(x + I(0.18), y + I(2.85), cw - I(0.36), I(0.4), [P(c.name, { sz: 20, b: 1, color: hx(chd.color), font: FONT_D })], { name: 'Name' });
      s.text(x + I(0.18), y + I(3.3), cw - I(0.36), ch - I(3.4), [P(c.bio, { sz: 12.5, color: 'DCE3F7', sp: 16.5 })], { name: 'Bio' });
    });
    s.notes = D.cast.map(function (c) { return c.name + ': ' + c.bio; }).join('\n');
    s.adv = 9000;
  }
  function villainsSlide(X) {
    var D = X.D, s = X.deck.slide({ bg: NIGHT, ctx: 'villains' });
    s.text(I(0.6), I(0.4), I(12), I(0.3), [P('Anti-patterns', { sz: 12, b: 1, color: '9FB0D9', caps: 1, spc: 160 })], { name: 'Kicker' });
    s.text(I(0.6), I(0.7), I(12), I(0.65), [P('The villains', { sz: 34, b: 1, color: 'FFFFFF', font: FONT_D })], { name: 'Title' });
    s.text(I(0.6), I(1.35), I(12), I(0.4), [P('Every villain is a real-world mistake. Each one is defeated by a principle from CAF or WAF.', { sz: 15, color: '9FB0D9' })], { name: 'Intro' });
    var cols = 4, gap = I(0.18), cw = (SW - I(1.2) - (cols - 1) * gap) / cols, chh = I(2.45), y0 = I(1.95);
    D.villains.forEach(function (v, i) {
      var r = Math.floor(i / cols), c = i % cols, x = I(0.6) + c * (cw + gap) + (r === 1 ? (cw + gap) / 2 : 0), y = y0 + r * (chh + gap);
      s.sp({ geom: 'roundRect', adj: { adj: 9000 }, x: x, y: y, w: cw, h: chh, fill: '101A36', line: { c: '22305A', w: 9525 }, name: 'Card ' + v.name });
      s.contain(M(X, 'villains/' + v.id + '-base'), x + I(0.12), y + I(0.15), I(1.0), chh - I(0.3), { name: v.name, descr: v.name });
      s.text(x + I(1.22), y + I(0.15), cw - I(1.34), chh - I(0.3), [P(v.name, { sz: 16, b: 1, color: hx((D.chars[v.id] || {}).color || '#fff'), font: FONT_D, sp: 20 }), P(v.what, { sz: 12, color: 'DCE3F7', sp: 15.5, bef: 3 }), P([{ t: 'Defeated by: ', b: 1 }, { t: v.beat }], { sz: 12, color: '9FB0D9', sp: 15.5, bef: 4 })], { name: 'Villain text' });
    });
    s.notes = D.villains.map(function (v) { return v.name + ': ' + v.what + ' Defeated by: ' + v.beat + '.'; }).join('\n');
    s.adv = 9000;
  }
  function contentsSlide(X) {
    var D = X.D, s = X.deck.slide({ bg: NIGHT, ctx: 'contents' });
    s.text(I(0.6), I(0.35), I(12), I(0.6), [P('Contents', { sz: 32, b: 1, color: 'FFFFFF', font: FONT_D })], { name: 'Title' });
    var parts = [D.eps.filter(function (e) { return e.part === 'prologue' || e.part === 'caf'; }), D.eps.filter(function (e) { return e.part === 'waf' || e.part === 'finale'; })];
    var heads = ['Part 1 · The voyage: Cloud Adoption Framework', 'Part 2 · The shipwright\'s code: Well-Architected Framework'], cw = I(5.95);
    parts.forEach(function (list, pi) {
      var x = I(0.6) + pi * (cw + I(0.25)), y = I(1.05);
      s.text(x, y, cw, I(0.3), [P(heads[pi], { sz: 12, b: 1, color: pi ? 'C4B5FD' : '7DD3FC', caps: 1, spc: 100 })], { name: 'Part ' + (pi + 1) }); y += I(0.42);
      list.forEach(function (ep) {
        s.cover(M(X, 'panels/' + ep.thumb), x, y, I(0.8), I(0.53), { name: ep.n + ' art', descr: '' });
        s.text(x + I(0.95), y - I(0.02), cw - I(0.95), I(0.3), [P([{ t: ep.n + ' · ' + ep.title, jump: ep.id, color: 'FFFFFF' }], { sz: 15, b: 1 })], { name: ep.n });
        s.text(x + I(0.95), y + I(0.28), cw - I(0.95), I(0.26), [P(ep.phase + ' · ' + ep.dur, { sz: 12, color: '9FB0D9' })], { name: ep.n + ' topic' });
        y += I(0.6);
      });
    });
    s.notes = 'Contents. Click an episode title to jump to it.\n' + D.eps.map(function (e) { return e.n + ' · ' + e.title + ' (' + e.phase + ', ' + e.dur + ')'; }).join('\n');
    s.adv = 7000;
  }
  function sourcesSlides(X, keys, kicker, title) {
    keys = keys.filter(function (k) { return X.D.sources[k]; }); var list = keys.map(function (k) { return X.D.sources[k]; });
    var colW = I(5.95), gap = I(0.4), nd = I(0.4), tx = I(0.55), tW = colW - tx, top = I(1.55), bottom = SH - I(0.9), pages = [[]], col = 0, y = top;
    function meta(S) { return (S[2] ? 'Updated ' + S[2] + ' · ' : '') + S[1].replace(/^https:\/\//, ''); }
    list.forEach(function (S, i) {
      var h = lineCount(S[0], 15, FONT, tW, true) * 18.5 * PT + lineCount(meta(S), 12, FONT, tW) * 15 * PT;
      if (y + h > bottom && y > top) { col++; y = top; if (col > 1) { pages.push([]); col = 0; } }
      pages[pages.length - 1].push({ S: S, n: i + 1, x: I(0.6) + col * (colW + gap), y: y, h: h });
      y += h + I(0.22);
    });
    pages.forEach(function (pg, p) {
      var s = X.deck.slide({ bg: NIGHT, ctx: 'src:' + kicker + p });
      boardFrame(X, s, { w: SW - I(0.5), h: SH - I(0.44), y: I(0.22) });
      s.text(I(0.6), I(0.42), I(12), I(0.3), [P(kicker, { sz: 12, b: 1, color: MUTED, caps: 1, spc: 140 })], { name: 'Kicker' });
      s.text(I(0.6), I(0.7), I(12), I(0.6), [P(title + (pages.length > 1 ? ' (' + (p + 1) + ' of ' + pages.length + ')' : ''), { sz: 30, b: 1, color: INK, font: FONT_D })], { name: 'Title' });
      pg.forEach(function (it) {
        s.sp({ geom: 'ellipse', x: it.x, y: it.y + 9.25 * PT - nd / 2, w: nd, h: nd, fill: deep(X.accent), paras: [P(String(it.n), { sz: 12, b: 1, color: 'FFFFFF', algn: 'ctr' })], tx: { anchor: 'ctr' }, name: 'Number ' + it.n });
        s.text(it.x + tx, it.y, tW, it.h, [P([{ t: it.S[0], link: it.S[1], color: LINK }], { sz: 15, b: 1, sp: 18.5 }), P(meta(it.S), { sz: 12, color: MUTED, sp: 15 })], { name: 'Source ' + it.n });
      });
      s.text(I(0.6), SH - I(0.7), I(12), I(0.3), [P('All pages were read on ' + X.D.retrieved + '. Learn changes often, so check each page before you present.', { sz: 12, i: 1, color: MUTED })], { name: 'Retrieved' });
      s.notes = title + '\n' + srcNotes(X, pg.map(function (it) { return keys[it.n - 1]; }));
      s.adv = 7000;
    });
  }
  function closingSlide(X) {
    var D = X.D, s = X.deck.slide({ bg: NIGHT, ctx: 'closing' }), x = I(7.55), w = I(5.3), y = I(1.0);
    hero(s, M(X, 'panels/finale'), 'left', 0, I(7.4), { name: 'Finale art', descr: D.panelAlt.finale || '' });
    s.text(x, y, w, I(0.3), [P(D.title, { sz: 13, b: 1, color: '9FB0D9', caps: 1, spc: 160 })], { name: 'Brand' }); y += I(0.45);
    s.text(x, y, w, I(0.9), [P([{ t: 'Land it ' }, { t: 'well', color: 'FBBF24' }, { t: '.' }], { sz: 50, b: 1, color: 'FFFFFF', font: FONT_D })], { name: 'Title' }); y += I(1.0);
    var msg = 'Thanks for reading. Plan the voyage with the Cloud Adoption Framework, and build every workload with the Well-Architected Framework.', mn = lineCount(msg, 19, FONT, w);
    s.text(x, y, w, mn * 25 * PT, [P(msg, { sz: 19, color: 'E7ECFF', sp: 25 })], { name: 'Message' }); y += mn * 25 * PT + I(0.25);
    s.text(x, y, w, I(0.6), [P('Read it online:', { sz: 14, color: 'CBD5F5' }), P([{ t: D.site.replace(/^https?:\/\//, '').replace(/\/$/, ''), link: D.site, color: '7DD3FC', b: 1 }], { sz: 14 })], { name: 'Website' }); y += I(0.75);
    var dn = lineCount(D.disclaimer, 12, FONT, w);
    s.text(x, y, w, dn * 15.5 * PT, [P(D.disclaimer, { sz: 12, color: 'B4C1E3', sp: 15.5 })], { name: 'Disclaimer' });
    s.notes = 'Naming notes (checked ' + D.retrieved + '):\n' + D.notes.map(function (n) { return '• ' + n; }).join('\n') + '\n\n' + D.disclaimer;
    s.adv = 8000;
  }
  function dgContentsSlide(X) {
    var D = X.D, s = X.deck.slide({ bg: NIGHT, ctx: 'dgcontents' });
    s.text(I(0.6), I(0.35), I(12), I(0.6), [P('The diagrams', { sz: 32, b: 1, color: 'FFFFFF', font: FONT_D })], { name: 'Title' });
    var list = D.dgOrder, rows = Math.ceil(list.length / 2), cw = I(5.95);
    list.forEach(function (x, i) {
      var d = D.diagrams[x.id], ep = epById(D, x.ep), c = Math.floor(i / rows), r = i % rows, xx = I(0.6) + c * (cw + I(0.25)), y = I(1.15) + r * I(0.64);
      s.sp({ geom: 'ellipse', x: xx, y: y + I(0.04), w: I(0.36), h: I(0.36), fill: hx(ep ? ep.accent : '#0ea5e9'), paras: [P(String(i + 1), { sz: 12, b: 1, color: 'FFFFFF', algn: 'ctr' })], tx: { anchor: 'ctr' }, name: 'Number' });
      s.text(xx + I(0.5), y, cw - I(0.5), I(0.3), [P([{ t: d.title, jump: 'dg:' + x.id, color: 'FFFFFF' }], { sz: 15, b: 1 })], { name: d.title });
      s.text(xx + I(0.5), y + I(0.29), cw - I(0.5), I(0.26), [P((ep ? ep.n + ' · ' + ep.title + ' · ' : '') + d.steps.length + ' steps', { sz: 12, color: '9FB0D9' })], { name: 'From' });
    });
    s.notes = 'Click a diagram title to jump to it.';
    s.adv = 7000;
  }

  /* ============================================================ animated architecture diagrams */
  function exitPt(N, T) {
    if (!N.hw && !N.ht) return { x: N.x, y: N.y };
    var dx = T.x - N.x, dy = T.y - N.y, pad = 6;
    var tx = dx > 0 ? (N.hw + pad) / dx : (dx < 0 ? -(N.hw + pad) / dx : Infinity), ty = dy > 0 ? (N.hb + pad) / dy : (dy < 0 ? -(N.ht + pad) / dy : Infinity);
    var t = Math.min(tx, ty); if (!(t < 0.95)) t = 0.5;
    return { x: N.x + dx * t, y: N.y + dy * t };
  }
  function roundCmds(p, r) {
    var c = [{ t: 'M', p: p[0] }];
    for (var i = 1; i < p.length - 1; i++) {
      var a = p[i - 1], b = p[i], d = p[i + 1], l1 = Math.hypot(b.x - a.x, b.y - a.y) || 1, l2 = Math.hypot(d.x - b.x, d.y - b.y) || 1, rr = Math.min(r, l1 / 2, l2 / 2);
      c.push({ t: 'L', p: { x: b.x - (b.x - a.x) / l1 * rr, y: b.y - (b.y - a.y) / l1 * rr } }, { t: 'Q', c: { x: b.x, y: b.y }, p: { x: b.x + (d.x - b.x) / l2 * rr, y: b.y + (d.y - b.y) / l2 * rr } });
    }
    c.push({ t: 'L', p: p[p.length - 1] });
    return c;
  }
  function sample(cmds) {
    var out = [], cur = null;
    cmds.forEach(function (c) {
      if (c.t !== 'Q') { cur = c.p; out.push(cur); return; }
      for (var j = 1; j <= 6; j++) { var u = j / 6, v = 1 - u; out.push({ x: v * v * cur.x + 2 * v * u * c.c.x + u * u * c.p.x, y: v * v * cur.y + 2 * v * u * c.c.y + u * u * c.p.y }); }
      cur = c.p;
    });
    return out;
  }
  function plen(poly) { var L = 0; for (var i = 1; i < poly.length; i++) L += Math.hypot(poly[i].x - poly[i - 1].x, poly[i].y - poly[i - 1].y); return L; }
  function pointAt(poly, f) {
    var L = plen(poly) * f, acc = 0;
    for (var i = 1; i < poly.length; i++) { var d = Math.hypot(poly[i].x - poly[i - 1].x, poly[i].y - poly[i - 1].y); if (acc + d >= L) { var u = d ? (L - acc) / d : 0; return { x: poly[i - 1].x + (poly[i].x - poly[i - 1].x) * u, y: poly[i - 1].y + (poly[i].y - poly[i - 1].y) * u }; } acc += d; }
    return poly[poly.length - 1];
  }
  var GEOM = {};
  function dgGeom(d) {
    if (GEOM[d.id]) return GEOM[d.id];
    var nodes = {}, edges = {};
    d.nodes.forEach(function (nd) {
      var cls = nd.cls || '', lane = nd.icon.indexOf('lane:') === 0, sm = cls.indexOf('sm') >= 0, N = { x: nd.x, y: nd.y, nd: nd, lane: lane, sm: sm };
      if (lane) { N.hw = 75; N.ht = 20; N.hb = 20; }
      else { var half = sm ? 26 : 31, lines = String(nd.label).split('\n'), top = cls.indexOf('lt') >= 0; N.half = half; N.ic = sm ? 36 : 44; N.lines = lines; N.top = top; N.hw = half; N.ht = top ? half + 6 + lines.length * 14 : half; N.hb = top ? half : half + 6 + lines.length * 14; }
      nodes[nd.id] = N;
    });
    d.edges.forEach(function (ed) {
      var A = nodes[ed.a], B = typeof ed.b === 'string' ? nodes[ed.b] : { x: ed.b[0], y: ed.b[1], hw: 0, ht: 0, hb: 0 };
      if (!A || !B) return;
      var tree = (ed.cls || '').indexOf('tree') >= 0, pts;
      if (tree) { var y1 = A.y + A.hb - 2, y2 = B.y - B.ht - 4, mid = Math.round((y1 + y2) / 2); pts = [{ x: A.x, y: y1 }, { x: A.x, y: mid }, { x: B.x, y: mid }, { x: B.x, y: y2 }]; }
      else { pts = [{ x: A.x, y: A.y }].concat((ed.via || []).map(function (p) { return { x: p[0], y: p[1] }; })).concat([{ x: B.x, y: B.y }]); pts[0] = exitPt(A, pts[1]); pts[pts.length - 1] = exitPt(B, pts[pts.length - 2]); }
      var cmds = roundCmds(pts, 14);
      edges[ed.id] = { id: ed.id, tree: tree, cls: ed.cls || '', label: ed.label, lp: ed.lp == null ? 0.5 : ed.lp, cmds: cmds, poly: sample(cmds) };
    });
    return (GEOM[d.id] = { nodes: nodes, edges: edges });
  }
  function dgState(d, k) {
    var down = {}, ghost = {};
    d.nodes.forEach(function (n) { if ((n.cls || '').indexOf('ghost') >= 0) ghost[n.id] = 1; });
    for (var i = 0; i <= k && i < d.steps.length; i++) { var s = d.steps[i]; s.down.forEach(function (x) { down[x] = 1; }); s.up.forEach(function (x) { delete down[x]; }); s.show.forEach(function (x) { delete ghost[x]; }); s.hide.forEach(function (x) { ghost[x] = 1; }); }
    return { down: down, ghost: ghost };
  }
  function lab(d, id) { var n = d.nodes.filter(function (x) { return x.id === id; })[0]; return n ? String(n.label).replace(/\n/g, ' ') : id; }
  function diagramSlides(X, it) {
    var d = X.D.diagrams[it.did], G = dgGeom(d), label = it.ep ? it.ep.n + ' · ' + it.ep.title : 'Architecture';
    for (var k = -1; k < d.steps.length; k++) {
      var s = diagramSlide(X, d, G, k, label, (it.standalone ? 'dgs:' : 'dg:') + d.id + ':' + (it.ep ? it.ep.id : ''), it.standalone ? null : it);
      if (k === -1 && it.standalone) X.deck.anchors['dg:' + d.id] = s;
    }
  }
  function diagramSlide(X, d, G, k, label, ctx, it) {
    var s = X.deck.slide({ bg: NIGHT, ctx: ctx });
    if (it && it.fr) { backdrop(X, s, it.fr.st); s.sp({ x: 0, y: 0, w: SW, h: SH, fill: NIGHT, fa: 0.74, name: '!!scrim' }); }
    boardFrame(X, s, { w: SW - I(0.5), h: SH - I(0.44), y: I(0.22) });
    var DGX = I(0.55), DGY = I(1.25), srcW = SW - 2 * DGX, srcH = lineCount(srcRuns(X, d.src), 12, FONT, srcW) * 15 * PT;
    var DGW = Math.min(I(9.6), (SH - I(0.36) - srcH - I(0.2) - DGY) * d.vb[0] / d.vb[1]);
    var sc = DGW / d.vb[0], st = dgState(d, k), step = d.steps[k], col = step ? KIND[step.k] || KIND.req : '0EA5E9';
    function px(v) { return DGX + v * sc; }
    function py(v) { return DGY + v * sc; }
    function cm(c) { return c.map(function (q) { return q.t === 'Q' ? { t: 'Q', c: { x: px(q.c.x), y: py(q.c.y) }, p: { x: px(q.p.x), y: py(q.p.y) } } : { t: q.t, p: { x: px(q.p.x), y: py(q.p.y) } }; }); }
    s.text(DGX, I(0.42), I(9.9), I(0.26), [P('Animated architecture · ' + label, { sz: 12, b: 1, color: MUTED, caps: 1, spc: 140 })], { name: '!!dg-kick' });
    s.text(DGX, I(0.66), I(10), I(0.52), [P(d.title, { sz: 26, b: 1, color: INK, font: FONT_D, sp: 31 })], { name: '!!dg-title' });
    var cw = d.vb[0] * sc, ch = d.vb[1] * sc;
    s.sp({ geom: 'roundRect', adj: { adj: I(0.14) / ch * 100000 }, x: DGX, y: DGY, w: cw, h: ch, fillXml: grad([[0, 'F7FBFF'], [1, 'EEF4FB']], 90), line: { c: 'DBE4F0', w: 9525 }, name: '!!dg-canvas', descr: d.title + '. The steps are described in the speaker notes.' });
    d.zones.forEach(function (z) {
      var zs = ZONE[(z.cls || '').split(' ').filter(function (c) { return ZONE[c]; })[0] || '_'], dn = st.down[z.id];
      s.sp({ geom: 'roundRect', adj: { adj: 14 / Math.min(z.w, z.h) * 100000 }, x: px(z.x), y: py(z.y), w: z.w * sc, h: z.h * sc, fill: dn ? 'FEE2E2' : zs[0], fa: 0.55, line: { c: dn ? 'EF4444' : zs[1], w: 1.5 * sc, dash: zs[2] }, name: '!!z-' + z.id });
      s.text(px(z.x + 12), py(z.y + 19) - 10.5 * PT, tw(z.label, 10.5, true) + I(0.08), 10.5 * 1.3 * PT, [P(z.label, { sz: 10.5, b: 1, color: '475569', glow: { r: 38100, c: 'F8FBFF' } })], { name: '!!zl-' + z.id, wrap: false });
    });
    Object.keys(G.edges).forEach(function (id) {
      var e = G.edges[id];
      s.path(cm(e.cmds), { line: { c: e.tree ? 'B6C2D6' : '94A3B8', a: e.cls.indexOf('faded') >= 0 ? 0.35 : null, w: (e.tree ? 2 : 2.4) * sc, dash: e.cls.indexOf('dash') >= 0 ? 'dash' : null, tail: e.tree ? null : 'triangle', cap: 'rnd' }, name: '!!e-' + id });
    });
    var hops = [], t = 0;
    if (step) {
      step.e.forEach(function (raw) {
        var rev = raw.charAt(0) === '-', e = G.edges[rev ? raw.slice(1) : raw]; if (!e) return;
        hops.push({ e: e, poly: rev ? e.poly.slice().reverse() : e.poly, dur: Math.max(420, plen(e.poly) / 230 * 1000) });
      });
      hops.forEach(function (h) { h.lid = s.path(cm(h.e.cmds), { line: { c: col, w: 4 * sc, tail: h.e.tree ? null : 'triangle', cap: 'rnd' }, name: 'Lit ' + h.e.id }); });
    }
    d.nodes.forEach(function (nd) {
      var N = G.nodes[nd.id], hl = step && step.hl.indexOf(nd.id) >= 0, dn = st.down[nd.id], al = st.ghost[nd.id] ? 0.2 : dn ? 0.55 : 1, glow = hl ? { r: I(0.07), c: col, a: 0.45 } : null;
      if (N.lane) {
        var lc = hx(nd.icon.slice(5));
        s.sp({ geom: 'roundRect', adj: { adj: 45000 }, x: px(nd.x - 75), y: py(nd.y - 20), w: 150 * sc, h: 40 * sc, fill: hl ? 'F0F9FF' : 'FFFFFF', fa: al, line: { c: hl ? col : lc, w: (hl ? 3.5 : 2.5) * sc, a: al }, glow: glow, name: '!!n-' + nd.id });
        s.sp({ geom: 'ellipse', x: px(nd.x - 75 + 11), y: py(nd.y - 7), w: 14 * sc, h: 14 * sc, fill: lc, fa: al, name: '!!nd-' + nd.id });
        s.text(px(nd.x + 9 - 58), py(nd.y - 20), 116 * sc, 40 * sc, [P(nd.label, { sz: 10, b: 1, color: '1E293B', alpha: al, algn: 'ctr' })], { anchor: 'ctr', name: '!!nl-' + nd.id, wrap: false });
        return;
      }
      var half = N.half, ic = N.ic, fs = N.sm ? 9 : 9.5, pitch = 14 * sc;
      s.sp({ geom: 'roundRect', adj: { adj: 13 / (2 * half) * 100000 }, x: px(nd.x - half), y: py(nd.y - half), w: 2 * half * sc, h: 2 * half * sc, fill: hl ? 'F0F9FF' : dn ? 'FEE2E2' : 'FFFFFF', fa: al, line: { c: hl ? col : dn ? 'EF4444' : 'CBD5E1', w: (hl ? 3.5 : 1.5) * sc, a: al }, glow: glow, name: '!!n-' + nd.id });
      s.pic({ media: M(X, 'icon/' + nd.icon), x: px(nd.x - ic / 2), y: py(nd.y - ic / 2), w: ic * sc, h: ic * sc, alpha: al, name: '!!ni-' + nd.id, descr: String(nd.label).replace(/\n/g, ' ') });
      var lines = N.lines, base = N.top ? nd.y - half - 8 - (lines.length - 1) * 14 : nd.y + half + 16;
      s.text(px(nd.x) - I(0.9), py(base) - fs * PT, I(1.8), lines.length * pitch + I(0.03), lines.map(function (l) { return P(l, { sz: fs, b: 1, color: '1E293B', alpha: al, algn: 'ctr', sp: pitch / PT, glow: { r: 38100, c: 'F4F8FD' } }); }), { name: '!!nl-' + nd.id, wrap: false });
    });
    Object.keys(G.edges).forEach(function (id) {
      var e = G.edges[id]; if (!e.label) return;
      var p = pointAt(e.poly, e.lp), fs = 8.5, w = tw(e.label, fs, true) + I(0.12), h = fs * 1.5 * PT;
      s.sp({ geom: 'roundRect', adj: { adj: 30000 }, x: px(p.x) - w / 2, y: py(p.y) - h / 2, w: w, h: h, fill: 'FFFFFF', line: { c: 'CBD5E1', w: 6350 }, paras: [P(e.label, { sz: fs, b: 1, color: '475569', algn: 'ctr' })], tx: { anchor: 'ctr', wrap: false }, name: '!!el-' + id });
    });
    if (step) {
      if (hops.length) {
        // one packet per hop, each with a single motion path: no reliance on how players chain paths
        var r = 6 * sc;
        hops.forEach(function (h, hi) {
          var p0 = h.poly[0];
          h.pid = s.sp({ geom: 'ellipse', x: px(p0.x) - r, y: py(p0.y) - r, w: 2 * r, h: 2 * r, fill: col, line: { c: 'FFFFFF', w: 2 * sc }, glow: { r: I(0.09), c: col, a: 0.55 }, name: 'Packet ' + (hi + 1), descr: 'Traffic packet' });
        });
        var lastPid = hops[hops.length - 1].pid;
        if (X.motion) {
          hops.forEach(function (h, hi) {
            var pts = h.poly, p0 = pts[0], stepN = Math.max(1, Math.floor(pts.length / 24)), path = '';
            for (var i = 0; i < pts.length; i += stepN) path += (i ? ' L ' : 'M ') + f4((pts[i].x - p0.x) * sc / SW) + ' ' + f4((pts[i].y - p0.y) * sc / SH);
            var last = pts[pts.length - 1]; path += ' L ' + f4((last.x - p0.x) * sc / SW) + ' ' + f4((last.y - p0.y) * sc / SH) + ' E';
            s.fx.push({ k: 'fade', id: h.lid, t: R(t), dur: 250 }, { k: 'appear', id: h.pid, t: R(t) }, { k: 'path', id: h.pid, t: R(t), dur: R(h.dur), path: path, ease: 1 });
            t += h.dur;
            if (hi < hops.length - 1) s.fx.push({ k: 'hide', id: h.pid, t: R(t) });
          });
        }
        var lastH = hops[hops.length - 1], end = lastH.poly[lastH.poly.length - 1];
        if (step.block) {
          var xid = s.text(px(end.x + 10) - I(0.15), py(end.y - 10) - I(0.3), I(0.45), I(0.42), [P('✖', { sz: 22, b: 1, color: 'DC2626', algn: 'ctr' })], { name: 'Blocked', anchor: 'ctr', wrap: false });
          if (X.motion) s.fx.push({ k: 'zoom', id: xid, t: R(t), dur: 350 });
          t += 380;
        }
        if (X.motion) s.fx.push({ k: 'out', id: lastPid, t: R(t), dur: 250 });
      }
    }
    // right-hand panel: step badge, dots, caption, legend
    var rx = DGX + cw + I(0.25), rw = SW - I(0.55) - rx, y = DGY;
    var badge = step ? 'Step ' + (k + 1) + ' / ' + d.steps.length + ' · ' + (KIND_NAME[step.k] || step.k) : 'Overview', bw = Math.min(rw, tw(badge.toUpperCase(), 12, true) + 12 * 0.8 * badge.length * 100 + I(0.3));
    s.sp({ geom: 'roundRect', adj: { adj: 50000 }, x: rx, y: y, w: bw, h: I(0.36), fill: step ? deep(col) : INK, paras: [P(badge, { sz: 12, b: 1, color: 'FFFFFF', caps: 1, spc: 80, algn: 'ctr' })], tx: { anchor: 'ctr', wrap: false }, name: '!!dg-badge' });
    y += I(0.5);
    var dd = I(0.12), dg = I(0.07);
    for (var i = 0; i < d.steps.length; i++) s.sp({ geom: 'ellipse', x: rx + i * (dd + dg), y: y, w: dd, h: dd, fill: i <= k ? 'CBD5E1' : 'E2E8F0', name: '!!dot-' + i });
    if (step) s.sp({ geom: 'ellipse', x: rx + k * (dd + dg) - I(0.025), y: y - I(0.025), w: dd + I(0.05), h: dd + I(0.05), fill: col, name: '!!dot-on' });
    y += dd + I(0.3);
    var legendH = LEGEND.length * I(0.27), cap = step ? step.c : 'The whole architecture at a glance. Click to watch the traffic move hop by hop, one step per slide.';
    var cf = 17, maxH = DGY + ch - legendH - I(0.25) - y;
    while (cf > 13 && lineCount(cap, cf, FONT, rw) * cf * 1.24 * PT > maxH) cf -= 0.5;
    var cn = lineCount(cap, cf, FONT, rw);
    s.text(rx, y, rw, cn * cf * 1.24 * PT, [P(cap, { sz: cf, color: '1E293B', sp: cf * 1.24 })], { name: '!!dg-cap' });
    var ly = DGY + ch - legendH;
    LEGEND.forEach(function (kk, j) {
      var yy = ly + j * I(0.27);
      s.sp({ geom: 'ellipse', x: rx, y: yy + I(0.07), w: I(0.13), h: I(0.13), fill: KIND[kk], name: '!!lg-' + kk });
      s.text(rx + I(0.22), yy, rw - I(0.22), I(0.27), [P(KIND_NAME[kk], { sz: 13, color: '475569' })], { name: '!!lgt-' + kk, anchor: 'ctr' });
    });
    srcFooter(X, s, d.src, DGX, SH - I(0.36), srcW, false);
    var hopTxt = step ? step.e.map(function (raw) { var rev = raw.charAt(0) === '-', ed = d.edges.filter(function (e) { return e.id === (rev ? raw.slice(1) : raw); })[0]; if (!ed) return ''; var a = lab(d, ed.a), b = typeof ed.b === 'string' ? lab(d, ed.b) : 'out'; return rev ? b + ' → ' + a : a + ' → ' + b; }).filter(Boolean).join(', ') : '';
    s.notes = d.title + (step ? '\nStep ' + (k + 1) + ' of ' + d.steps.length + ' (' + (KIND_NAME[step.k] || step.k) + '): ' + step.c + (hopTxt ? '\nTraffic: ' + hopTxt + (step.block ? ' (blocked)' : '') : '') :
      '\nOverview. Steps:\n' + d.steps.map(function (x, j) { return (j + 1) + '. ' + x.c; }).join('\n')) + (d.note ? '\n\nNote: ' + d.note : '') + '\n\nSources on Microsoft Learn:\n' + srcNotes(X, d.src);
    s.adv = step ? t + 2600 + step.c.length * 25 : 4000;
    return s;
  }

  /* ============================================================ compose + build */
  var STORY = { scene: 1, card: 1, quiz: 1, answer: 1, end: 1 };
  function compose(deck, items, opts, prog) {
    var X = { deck: deck, D: deck.D, seq: 0, motion: true }, i = 0;
    function near(idx, dir) { var j = idx + dir; while (items[j] && items[j].k === 'section') j += dir; var n = items[j]; return n && STORY[n.k] && n.ep === items[idx].ep ? n : null; }
    function one(it, idx) {
      var D = deck.D;
      X.ep = it.ep || null; X.accent = hx(it.ep ? it.ep.accent : '#0ea5e9'); X.prevIt = near(idx, -1); X.nextIt = near(idx, 1);
      switch (it.k) {
        case 'section': deck.section(it.name); break;
        case 'cover': coverSlide(X, '', 'Contoso is moving to Azure, and Dewi has to lead it. Learn the Cloud Adoption Framework (how a whole organization adopts Azure) and the Well-Architected Framework (how to build each workload well) in an anime story, with animated architecture diagrams and every fact linked to Microsoft Learn.',
          D.eps.length + ' episodes · about ' + D.eps.reduce(function (a, e) { return a + (parseInt(e.dur, 10) || 0); }, 0) + ' minutes · ' + D.dgOrder.length + ' animated diagrams · ' + Object.keys(D.eps.reduce(function (o, e) { e.srcs.forEach(function (k) { o[k] = 1; }); return o; }, {})).length + ' Microsoft Learn sources'); break;
        case 'dgCover': coverSlide(X, 'The animated architecture diagrams', 'Every diagram from the comic, built from editable shapes with the official Azure icons. Each step is a slide, and its traffic flies hop by hop. The speaker notes explain every step and link to Microsoft Learn.', D.dgOrder.length + ' diagrams · ' + D.dgOrder.reduce(function (a, x) { return a + D.diagrams[x.id].steps.length; }, 0) + ' steps'); break;
        case 'cast': castSlide(X); break;
        case 'villains': villainsSlide(X); break;
        case 'contents': contentsSlide(X); break;
        case 'dgContents': dgContentsSlide(X); break;
        case 'epTitle': epTitleSlide(X, it.ep); break;
        case 'scene': sceneSlide(X, it); break;
        case 'card': cardSlides(X, it); break;
        case 'diagram': diagramSlides(X, it); break;
        case 'quiz': quizSlide(X, it, false); break;
        case 'answer': quizSlide(X, it, true); break;
        case 'end': endSlide(X, it); break;
        case 'sources': sourcesSlides(X, it.ep.srcs, it.ep.n + ' · ' + it.ep.title, 'Sources on Microsoft Learn'); break;
        case 'dgSources': sourcesSlides(X, Object.keys(D.dgOrder.reduce(function (o, x) { D.diagrams[x.id].src.forEach(function (k) { o[k] = 1; }); return o; }, {})), 'All diagrams', 'Sources on Microsoft Learn'); break;
        case 'closing': closingSlide(X); break;
      }
    }
    return new Promise(function (res, rej) {
      function tick() {
        try {
          var t0 = Date.now();
          while (i < items.length && Date.now() - t0 < 40) { one(items[i], i); i++; }
          prog(i / items.length);
          if (i < items.length) { setTimeout(tick, 0); return; }
          deck.slides.forEach(function (s, j) {
            var p = deck.slides[j - 1];
            s.trans = p && s.ctx && p.ctx === s.ctx ? 'morph' : 'fade';
            s.adv = opts.auto ? Math.max(1500, s.adv || 4000) + (s.trans === 'morph' ? 1000 : 750) : 0;
          });
          res();
        } catch (e) { rej(e); }
      }
      tick();
    });
  }
  var dataP = null;
  function loadData(o) {
    if (!dataP) dataP = Promise.all([o.data, o.icons].map(function (u) { return fetch(u).then(function (r) { if (!r.ok) throw new Error('Could not load the comic data (' + r.status + ').'); return r.json(); }); }))
      .catch(function (e) { dataP = null; throw e; });
    return dataP;
  }
  function fileName(D, o) {
    if (o.scope === 'all') return 'Land-It-Well-Dewi-Full-Comic.pptx';
    if (o.scope === 'diagrams') return 'Land-It-Well-Dewi-Architecture-Diagrams.pptx';
    var ep = epById(D, o.ep) || D.eps[0]; return 'Land-It-Well-Dewi-' + ep.n.replace(/\s+/g, '') + '-' + slug(ep.title) + '.pptx';
  }
  function build(o, onProgress) {
    var prog = onProgress || function () { };
    prog(0.02, 'Loading the story…');
    return loadData(o).then(function (res) {
      var D = res[0], icons = res[1], items = makePlan(D, o), ep = epById(D, o.ep) || D.eps[0];
      var deck = new Deck(D, { auto: !!o.auto, title: o.scope === 'all' ? D.title + ': the full comic' : o.scope === 'diagrams' ? D.title + ': architecture diagrams' : D.title + ' · ' + ep.n + ': ' + ep.title });
      return loadMedia(deck, icons, mediaKeys(D, items), function (f) { prog(0.05 + f * 0.5, 'Preparing pictures…'); })
        .then(function () { return compose(deck, items, o, function (f) { prog(0.55 + f * 0.2, 'Building slides…'); }); })
        .then(function () { return pack(deck, function (f) { prog(0.75 + f * 0.25, 'Packing the file…'); }); })
        .then(function (blob) { return { blob: blob, name: fileName(D, o), slides: deck.slides.length }; });
    });
  }

  /* ============================================================ dialog */
  var CUR = null;
  function save(blob, name) {
    if (window.navigator && navigator.msSaveOrOpenBlob) { navigator.msSaveOrOpenBlob(blob, name); return; }
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 60000);
  }
  function open(o) {
    if (CUR) CUR.close(true);
    var ov = document.createElement('div'), busy = false, prevFocus = document.activeElement;
    ov.className = 'xp-ov';
    function opt(v, inner) { return '<label class="xp-o"><input type="radio" name="xp-scope" value="' + v + '">' + inner + '</label>'; }
    ov.innerHTML = '<div class="xp" role="dialog" aria-modal="true" aria-labelledby="xp-h"><button class="xp-x" type="button" aria-label="Close">✕</button>' +
      '<span class="kick">Export</span><h2 id="xp-h">Download as PowerPoint</h2>' +
      '<p class="xp-sub">Built right here in your browser, nothing is uploaded. Every beat becomes a slide with Morph transitions and typed speech, plus lesson cards, quizzes, editable diagrams whose packets fly hop by hop, and speaker notes.</p><div class="xp-opts">' +
      (o.ep ? opt('episode', '<b>This episode</b><span data-c="episode">…</span>') : opt('pick', '<b>One episode</b><span data-c="pick">…</span><select aria-label="Episode"></select>')) +
      opt('all', '<b>The whole comic</b><span data-c="all">…</span>') + opt('diagrams', '<b>Only the animated diagrams</b><span data-c="diagrams">…</span>') + '</div>' +
      '<label class="xp-chk"><input type="checkbox" id="xp-auto"><span>Self-running: slides advance on their own, like auto-play. You can still click.</span></label>' +
      '<div class="xp-prog" hidden><div class="xp-bar"><i></i></div><p class="xp-st" aria-live="polite"></p></div>' +
      '<div class="go"><button class="btn" type="button" data-xp="close">Close</button><button class="btn primary" type="button" data-xp="go">⬇ Download .pptx</button></div>' +
      '<p class="xp-note">Best in PowerPoint for Microsoft 365, PowerPoint 2019 or later, or PowerPoint for the web. Older versions show simple fades instead of Morph.</p></div>';
    document.body.appendChild(ov);
    var $ = function (q) { return ov.querySelector(q); }, sel = $('select'), go = $('[data-xp="go"]'), st = $('.xp-st'), bar = $('.xp-bar i');
    var want = o.scope === 'diagrams' ? 'diagrams' : o.scope === 'all' ? 'all' : o.ep ? 'episode' : 'pick';
    (ov.querySelector('input[value="' + want + '"]') || ov.querySelector('input[name="xp-scope"]')).checked = true;
    function close(force) { if (busy && !force) return; ov.remove(); CUR = null; document.removeEventListener('keydown', key); if (!force && prevFocus && prevFocus.focus) try { prevFocus.focus(); } catch (e) { } }
    function key(ev) { if (ev.key === 'Escape') close(); }
    CUR = { close: close };
    document.addEventListener('keydown', key);
    ov.addEventListener('click', function (ev) { if (ev.target === ov || ev.target.closest('.xp-x') || ev.target.closest('[data-xp="close"]')) close(); });
    function status(f, msg, err) { $('.xp-prog').hidden = false; bar.style.width = Math.round(f * 100) + '%'; st.textContent = msg; st.classList.toggle('err', !!err); }
    loadData(o).then(function (res) {
      var D = res[0];
      function count() {
        [['episode', o.ep], ['pick', sel && sel.value], ['all'], ['diagrams']].forEach(function (x) {
          var el = $('[data-c="' + x[0] + '"]'); if (!el) return;
          var n = slideEstimate(D, makePlan(D, { scope: x[0] === 'pick' ? 'episode' : x[0], ep: x[1] })), ep = x[1] && epById(D, x[1]);
          el.textContent = (x[0] === 'episode' && ep ? ep.n + ' · ' + ep.title + ' · ' : x[0] === 'all' ? D.eps.length + ' episodes with cast and contents · ' : x[0] === 'diagrams' ? D.dgOrder.length + ' diagrams, step by step · ' : '') + 'about ' + n + ' slides';
        });
      }
      if (sel) { sel.innerHTML = D.eps.map(function (e) { return '<option value="' + e.id + '">' + xe(e.n + ' · ' + e.title) + '</option>'; }).join(''); sel.addEventListener('change', function () { $('input[value="pick"]').checked = true; count(); }); }
      count();
    }, function (e) { status(0, e.message || 'The comic data could not load.', true); });
    go.addEventListener('click', function () {
      if (busy) return;
      var scope = ov.querySelector('input[name="xp-scope"]:checked').value, ep = scope === 'pick' ? sel.value : o.ep;
      busy = true; go.disabled = true; status(0.01, 'Starting…');
      build({ scope: scope === 'pick' ? 'episode' : scope, ep: ep, auto: $('#xp-auto').checked, data: o.data, icons: o.icons }, status).then(function (r) {
        save(r.blob, r.name); busy = false; go.disabled = false;
        status(1, 'Done: ' + r.name + ' (' + r.slides + ' slides, ' + (r.blob.size / 1048576).toFixed(1) + ' MB). Check your downloads.');
      }, function (e) { busy = false; go.disabled = false; status(0, 'Sorry, the export failed: ' + (e && e.message || e), true); });
    });
    setTimeout(function () { try { go.focus(); } catch (e) { } }, 30);
  }

  LIW.pptx = { open: open, build: build, plan: makePlan };
})();
