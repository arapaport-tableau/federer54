(function () {
  'use strict';

  // Federer 54 — a scrolling story about Roger Federer's career and the 54% of points he won.
  // Data: two embedded file datasources queried live through the Extensions API:
  //   "Federer Matches" (one row per tour-level match, 1998–2021)
  //   "Federer Charted", "Federer Shots", "Federer Serves" (Match Charting Project, shot level)
  // All DOM is built with createElement/textContent; SVG is hand-drawn (no external libraries).

  var root = document.getElementById('app');
  var SVGNS = 'http://www.w3.org/2000/svg';
  var C = {
    ink: '#1F2A24', soft: '#4A564F', muted: '#8C8A80', rule: '#DDD6C4', grass: '#2E6B3F',
    grassSoft: '#A9C4AF', purple: '#4B2E6B', purpleSoft: '#B9A9CC', gold: '#A88F52', goldSoft: '#E4D8B8',
    grey: '#C9C4B6', paper: '#FBF9F4'
  };

  // ------------------------------------------------------------------ helpers
  function extractData(result) { return result && Array.isArray(result.data) ? result.data : []; }

  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined && text !== null) e.textContent = String(text);
    return e;
  }
  function add(parent) { for (var i = 1; i < arguments.length; i++) if (arguments[i]) parent.appendChild(arguments[i]); return parent; }
  function sv(tag, attrs, text) {
    var e = document.createElementNS(SVGNS, tag);
    for (var k in attrs) if (Object.prototype.hasOwnProperty.call(attrs, k)) e.setAttribute(k, attrs[k]);
    if (text !== undefined && text !== null) e.textContent = String(text);
    return e;
  }
  // Phone mode: decided from the width the app is given (and touch), so it works in a phone browser and inside Tableau.
  // Everything phone-specific checks MOB; with MOB false the app draws exactly as on desktop.
  var MOB = false;
  function detectMobile() {
    var w = window.innerWidth || 1200, coarse = false;
    try { coarse = !!(window.matchMedia && window.matchMedia('(pointer: coarse)').matches); } catch (e) { /* ignore */ }
    return w <= 640 || (coarse && w <= 900);
  }
  function svgRoot(w, h) { return sv('svg', { viewBox: '0 0 ' + w + ' ' + h, preserveAspectRatio: 'xMidYMid meet', role: 'img' }); }
  function num(v) { if (v === null || v === undefined || v === '') return null; var n = Number(v); return isFinite(n) ? n : null; }
  function pct(a, b, d) { return b ? (100 * a / b).toFixed(d === undefined ? 1 : d) + '%' : '—'; }
  function fmt(n) { return n === null || n === undefined ? '—' : Math.round(n).toLocaleString('en-US'); }
  function scale(d0, d1, r0, r1) { return function (v) { return r0 + (v - d0) * (r1 - r0) / (d1 - d0); }; }
  function sum(arr, f) { var s = 0; arr.forEach(function (x) { var v = f(x); if (v !== null && v !== undefined) s += v; }); return s; }

  function renderError(msg) {
    root.textContent = '';
    add(root, el('p', 'error', 'Live query unavailable: ' + msg));
  }

  // tooltip
  var tip = el('div', 'tip');
  document.body.appendChild(tip);
  function showTip(evt, title, lines) {
    tip.textContent = '';
    add(tip, el('b', null, title));
    lines.forEach(function (l) { add(tip, el('div', null, l)); });
    var x = evt.clientX + 14, y = evt.clientY + 14;
    var w = 290;
    if (x + w > window.innerWidth) x = evt.clientX - w;
    if (y + 110 > window.innerHeight) y = evt.clientY - 110;
    tip.style.left = x + 'px'; tip.style.top = y + 'px'; tip.style.opacity = '1';
  }
  function hideTip() { tip.style.opacity = '0'; }

  // scroll reveal (falls back to visible when IntersectionObserver is unavailable)
  var io = null;
  var revealCallbacks = [];
  function reveal(node, onIn) {
    node.classList.add('reveal');
    if (onIn) { node.__onIn = onIn; }
    if (!io) { node.classList.add('in'); if (onIn) onIn(); return node; }
    io.observe(node);
    return node;
  }
  if ('IntersectionObserver' in window) {
    io = new IntersectionObserver(function (entries) {
      entries.forEach(function (en) {
        if (en.isIntersecting) {
          en.target.classList.add('in');
          if (en.target.__onIn) { var f = en.target.__onIn; en.target.__onIn = null; f(); }
          io.unobserve(en.target);
        }
      });
    }, { threshold: 0.15 });
  }

  // ------------------------------------------------------------------ data access
  function getDataSources() {
    var wb = tableau.extensions.workbook;
    if (wb && typeof wb.getAllDataSourcesAsync === 'function') return wb.getAllDataSourcesAsync();
    var wc = tableau.extensions.worksheetContent;
    if (wc && wc.worksheet && typeof wc.worksheet.getDataSourcesAsync === 'function') return wc.worksheet.getDataSourcesAsync();
    return Promise.resolve([]);
  }
  function byName(list, name) {
    var f = null;
    (list || []).forEach(function (ds) { if (ds && ds.name === name) f = ds; });
    if (!f) {
      var names = (list || []).map(function (d) { return d && d.name; }).join(', ');
      throw new Error('Datasource "' + name + '" not found. Available: ' + (names || '(none)'));
    }
    return f;
  }

  var MATCH_DIMS = ['match_key', 'match_date', 'season', 'tournament', 'level', 'surface', 'round', 'round_order', 'best_of',
    'opponent', 'opp_country', 'result', 'win', 'score', 'title', 'slam_title', 'defining'];
  var MATCH_MEAS = ['opp_rank', 'fed_rank', 'fed_age', 'minutes', 'aces', 'double_faults', 'serve_pts', 'serve_pts_won', 'first_in',
    'first_won', 'second_won', 'return_pts', 'return_pts_won', 'bp_faced', 'bp_saved', 'bp_chances', 'bp_converted',
    'points_won', 'points_total'];

  // One row per match_key (unique), so SUM() over each measure returns the row value.
  // If the aggregated query is rejected, retry with every field as a plain (disaggregated) column.
  function queryRows(ds, dims, meas, filters) {
    var q1 = { fields: dims.map(function (d) { return { fieldCaption: d }; })
      .concat(meas.map(function (m) { return { fieldCaption: m, function: 'SUM', fieldAlias: m }; })) };
    var q2 = { fields: dims.concat(meas).map(function (d) { return { fieldCaption: d }; }) };
    if (filters) { q1.filters = filters; q2.filters = filters; }
    return ds.queryAsync(q1).then(extractData).catch(function () {
      return ds.queryAsync(q2).then(extractData);
    }).then(function (rows) {
      return rows.map(function (r) {
        var o = {};
        dims.concat(meas).forEach(function (k) {
          var v = r[k];
          if (v === undefined) v = r['SUM(' + k + ')'];
          if (v === undefined) v = r['SUM_' + k];
          if (v && typeof v === 'object' && 'value' in v) v = v.value;
          o[k] = v;
        });
        return o;
      });
    });
  }

  function normMatches(rows) {
    var out = rows.map(function (r) {
      var o = {};
      MATCH_DIMS.forEach(function (k) { o[k] = r[k]; });
      MATCH_MEAS.forEach(function (k) { o[k] = num(r[k]); });
      ['season', 'round_order', 'best_of', 'win', 'title', 'slam_title'].forEach(function (k) { o[k] = num(r[k]); });
      o.match_date = String(o.match_date || '').slice(0, 10);
      o.defining = o.defining ? String(o.defining) : '';
      o.hasPts = o.points_total !== null && o.points_total > 0 && o.points_won !== null;
      o.pp = o.hasPts ? 100 * o.points_won / o.points_total : null;
      return o;
    });
    out.sort(function (a, b) { return a.match_date < b.match_date ? -1 : a.match_date > b.match_date ? 1 : (a.round_order - b.round_order); });
    return out;
  }
  var CH_DIMS = ['chart_id', 'match_key', 'opp_hand', 'best_of', 'final_rule', 'seq', 'set_starts', 'game_starts', 'mp_fed', 'mp_opp', 'bp_fed', 'bp_opp', 'sp_fed', 'sp_opp',
    'codes', 'source', 'spd', 'sw', 'sdp', 'rdp', 'snum', 'dfed', 'dopp', 'rlen'];
  var CH_MEAS = ['n_points', 'fed_points', 'aces', 'df', 'winners', 'ue', 'w_f', 'w_b', 'ue_f', 'ue_b', 'forced_induced', 'net_n', 'net_w',
    'snv_n', 'snv_w', 'sv_n', 'sv_w', 'first_in', 'first_w', 'second_n', 'second_w', 'rt_n', 'rt_w', 'r1_3_n', 'r1_3_w', 'r4_6_n', 'r4_6_w',
    'r7_9_n', 'r7_9_w', 'r10_n', 'r10_w', 'rally_sum', 'trick_n', 'trick_w', 'bh_slice', 'bh_drive', 'avg1_mph', 'max_mph', 'dist_fed_m', 'dist_opp_m', 'slam_cov'];
  var SH_DIMS = ['chart_id', 'shot', 'phase', 'origin', 'target', 'depth', 'outcome', 'miss'];
  var SH_MEAS = ['n', 'won', 'approach'];
  var SV_DIMS = ['chart_id', 'side', 'serve_no', 'direction', 'result'];
  var SV_MEAS = ['n', 'won', 'serve_volley'];

  // Large tables are fetched in chart_id ranges; if range filters are rejected, fall back to one query.
  function queryChunked(ds, dims, meas, maxId, step) {
    var parts = [];
    for (var lo = 1; lo <= maxId; lo += step) parts.push([lo, Math.min(maxId, lo + step - 1)]);
    var filtered = parts.map(function (r) {
      return queryRows(ds, dims, meas, [{ field: { fieldCaption: 'chart_id' }, filterType: 'QUANTITATIVE_NUMERICAL',
        quantitativeFilterType: 'RANGE', min: r[0], max: r[1] }]);
    });
    return Promise.all(filtered).then(function (chunks) {
      var all = [];
      chunks.forEach(function (c) { all = all.concat(c); });
      if (!all.length) throw new Error('empty');
      return all;
    }).catch(function () { return queryRows(ds, dims, meas); });
  }

  function normCharted(rows) {
    return rows.map(function (r) {
      var o = {};
      CH_DIMS.forEach(function (k) { o[k] = r[k] === null || r[k] === undefined ? '' : String(r[k]); });
      CH_MEAS.forEach(function (k) { o[k] = num(r[k]) || 0; });
      o.chart_id = num(r.chart_id); o.best_of = num(r.best_of) || 3;
      return o;
    });
  }
  function normShots(rows) {
    return rows.map(function (r) {
      return { chart_id: num(r.chart_id), shot: String(r.shot), phase: String(r.phase), origin: String(r.origin), target: String(r.target),
        depth: String(r.depth), outcome: String(r.outcome), miss: String(r.miss || 'None'), n: num(r.n) || 0, won: num(r.won) || 0, approach: num(r.approach) || 0 };
    });
  }
  function normServes(rows) {
    return rows.map(function (r) {
      return { chart_id: num(r.chart_id), side: String(r.side), serve_no: num(r.serve_no), direction: String(r.direction), result: String(r.result),
        n: num(r.n) || 0, won: num(r.won) || 0, serve_volley: num(r.serve_volley) || 0 };
    });
  }
  function plist(s) { return String(s || '').split('|').filter(function (x) { return x !== ''; }).map(Number); }

  // Rebuild a point-by-point series (score, server, running diff, match points) from the compact charted row.
  function buildSeries(c) {
    var seq = c.seq, n = seq.length;
    var setSt = plist(c.set_starts), gameSt = plist(c.game_starts), mpF = {}, mpO = {};
    plist(c.mp_fed).forEach(function (i) { mpF[i] = 1; }); plist(c.mp_opp).forEach(function (i) { mpO[i] = 1; });
    var lastSetNo = c.best_of === 5 ? 5 : 3, rule = c.final_rule;
    var out = [], cum = 0, setNo = 0, fg = 0, og = 0;
    var TS = ['0', '15', '30', '40'];
    for (var g = 0; g < gameSt.length; g++) {
      var a = gameSt[g], b = g + 1 < gameSt.length ? gameSt[g + 1] : n;
      if (setSt.indexOf(a) >= 0) { setNo++; fg = 0; og = 0; }
      var last = setNo === lastSetNo;
      var tb = fg === og && ((fg === 6 && (!last || rule === 'R1' || rule === 'RA')) || (fg === 12 && last && rule === 'RT'));
      var sp = 0, rp = 0;
      for (var i = a; i < b; i++) {
        var ch = seq.charAt(i), fedSrv = ch === 'A' || ch === 'B', won = ch === 'A' || ch === 'C';
        var gs;
        if (tb) gs = sp + '-' + rp;
        else if (sp >= 3 && rp >= 3) gs = sp === rp ? '40-40' : (sp > rp ? 'AD-40' : '40-AD');
        else gs = TS[Math.min(sp, 3)] + '-' + TS[Math.min(rp, 3)];
        cum += won ? 1 : -1;
        out.push({ point_no: i + 1, set_no: setNo, fed_games: fg, opp_games: og, game_score: gs, server: fedSrv ? 'Federer' : 'Opponent',
          tiebreak: tb ? 1 : 0, fed_won: won ? 1 : 0, fed_cum_diff: cum, mp_fed: mpF[i] ? 1 : 0, mp_opp: mpO[i] ? 1 : 0 });
        var srvWon = fedSrv === won;
        if (srvWon) sp++; else rp++;
      }
      var lastCh = seq.charAt(b - 1);
      if (lastCh === 'A' || lastCh === 'C') fg++; else og++;
    }
    out.rule = rule; out.bestOf = c.best_of;
    return out;
  }

  // ------------------------------------------------------------------ copy for the defining matches
  var DEFINING = [
    { key: 'D2001W', chip: '2001 · Sampras', title: 'The changing of the guard',
      story: 'A 19-year-old with a ponytail meets the four-time defending champion on Centre Court, the only time they ever played an official match. Federer wins in five sets and ends Sampras’s 31-match winning streak at Wimbledon. He won just 10 more points than Sampras over the whole afternoon.' },
    { key: 'D2003W', chip: '2003 · First major', title: 'The first one',
      story: 'Federer beats Mark Philippoussis in straight sets for his first Grand Slam title. Behind the tidy scoreline, both sets that went to tiebreaks were decided by a handful of points.' },
    { key: 'D2004AO', chip: '2004 · World No. 1', title: 'The summit',
      story: 'He wins the Australian Open over Marat Safin and becomes world No. 1 the next day. He stays there for a record 237 consecutive weeks.' },
    { key: 'D2008W', chip: '2008 · Nadal', title: 'The greatest match ever played?',
      story: 'Down two sets, Federer saves two championship points in the fourth-set tiebreak and forces a fifth set in the falling light. Nadal takes it 9–7 after 4 hours and 48 minutes. Over 413 points the gap between them is five.' },
    { key: 'D2009RG', chip: '2009 · Roland Garros', title: 'The career Grand Slam',
      story: 'With Nadal knocked out early, Federer finally wins in Paris. He completes the career Grand Slam and ties Sampras’s record of 14 majors.' },
    { key: 'D2009W', chip: '2009 · Roddick', title: 'Sixteen–fourteen',
      story: 'A record 15th major, in a fifth set that runs to 16–14. Roddick loses his serve only once in the match, in the 77th and final game.' },
    { key: 'D2012W', chip: '2012 · Murray', title: 'Seven at Wimbledon',
      story: 'He beats Andy Murray for his seventh Wimbledon title and 17th major. The win also returns him to world No. 1, and he soon passes Sampras’s record for total weeks at the top.' },
    { key: 'D2017AO', chip: '2017 · The return', title: 'The comeback nobody expected',
      story: 'After six months out with a knee injury, the 35-year-old arrives as the 17th seed. He trails Nadal by a break in the fifth set, then wins five straight games for his 18th major.' },
    { key: 'D2019W', chip: '2019 · Djokovic', title: 'More points. Fewer sets.',
      story: 'Federer holds two championship points on his own serve at 8–7, 40–15 in the fifth set, and loses the first 12–12 final-set tiebreak in Wimbledon history. He wins more points than Djokovic and still loses the match.' }
  ];

  // ------------------------------------------------------------------ sections
  function renderHero(M) {
    var stat = M.filter(function (m) { return m.hasPts; });
    var pw = sum(stat, function (m) { return m.points_won; }), pt = sum(stat, function (m) { return m.points_total; });
    var W = sum(M, function (m) { return m.win; }), L = M.length - W;
    var p = 100 * pw / pt;

    var s = el('section', 'hero');
    var heroPh = (window.FED_PHOTOS || []).filter(function (ph) { return ph.hero; })[0];
    var top = el('div', 'hero-top'), txt = el('div', 'hero-text');
    add(txt, el('div', 'eyebrow', 'Roger Federer · 1998–2021'));
    var big = el('div', 'big num');
    big.textContent = Math.round(p);
    add(big, el('sup', null, '%'));
    add(txt, big);
    var q = el('blockquote', 'quote');
    q.textContent = '“In the 1,526 singles matches I played in my career, I won almost 80% of those matches. Now, I have a question for all of you: What percentage of points do you think I won in those matches? Only 54%.”';
    add(q, el('span', 'who', 'Roger Federer, Dartmouth commencement, 2024'));
    add(txt, q);
    add(top, txt);
    if (heroPh) {
      s.className = 'hero has-photo';
      var fig = el('figure', 'hero-photo'), im = el('img');
      im.src = heroPh.file; im.alt = heroPh.alt || ''; im.decoding = 'async';
      if (heroPh.focus) im.style.objectPosition = heroPh.focus;
      im.addEventListener('error', function () { fig.style.display = 'none'; s.className = 'hero'; });
      add(fig, im);
      if (heroPh.credit) add(fig, el('figcaption', 'credit', heroPh.credit));
      add(top, fig);
    }
    add(s, top);

    var k = el('div', 'kpis');
    [[fmt(M.length), 'Matches'], [fmt(W) + '–' + fmt(L), 'Won – lost (' + pct(W, M.length) + ')'],
      [fmt(sum(M, function (m) { return m.title; })), 'Titles'], [fmt(sum(M, function (m) { return m.slam_title; })), 'Majors'],
      [p.toFixed(1) + '%', 'Points won']].forEach(function (x, i) {
      var d = el('div', 'kpi' + (i === 4 ? ' hi' : ''));
      add(d, el('div', 'v num', x[0]), el('div', 'l', x[1]));
      add(k, d);
    });
    add(s, k);

    var dots = el('div', 'dots100');
    var spans = [];
    for (var i = 0; i < 100; i++) { var sp = el('span'); spans.push(sp); add(dots, sp); }
    var wrap = el('div');
    add(wrap, dots, el('div', 'dots-cap', 'Out of every 100 points he played, he won ' + Math.round(p) + ' and lost ' + (100 - Math.round(p)) + '.'));
    add(s, reveal(wrap, function () {
      spans.forEach(function (sp, j) { if (j < Math.round(p)) setTimeout(function () { sp.classList.add('on'); }, 14 * j); });
    }));
    add(s, el('div', 'scroll-cue', 'Scroll ↓'));
    return s;
  }

  function chapterHead(numeral, kicker, title, lede) {
    var h = el('div', 'ch-head');
    add(h, el('span', 'eyebrow', numeral + ' · ' + kicker), el('h2', null, title));
    if (lede) add(h, el('p', 'lede', lede));
    return h;
  }

  // I. The Arc — season small multiples sharing the x axis
  function renderArc(M) {
    var seasons = {};
    M.forEach(function (m) {
      var s = seasons[m.season] || (seasons[m.season] = { y: m.season, w: 0, n: 0, pw: 0, pt: 0, slams: 0, titles: 0 });
      s.n++; s.w += m.win; s.slams += m.slam_title; s.titles += m.title;
      if (m.hasPts) { s.pw += m.points_won; s.pt += m.points_total; }
    });
    var S = Object.keys(seasons).map(function (k) { return seasons[k]; }).sort(function (a, b) { return a.y - b.y; });
    S.forEach(function (s) { s.wp = 100 * s.w / s.n; s.pp = s.pt ? 100 * s.pw / s.pt : null; });
    var peak = S.filter(function (s) { return s.n >= 20; }).sort(function (a, b) { return b.wp - a.wp; })[0];
    var early = S.filter(function (s) { return s.y <= 2000 && s.pp; });
    var earlyPP = sum(early, function (s) { return s.pw; }) / sum(early, function (s) { return s.pt; }) * 100;

    var sec = el('section', 'ch');
    add(sec, chapterHead('I', 'The arc', 'A gain of a few points won per hundred turned a promising player into the best in the world.',
      'From 1998 to 2000 Federer won ' + earlyPP.toFixed(1) + '% of his points, about even with his opponents. In ' + peak.y + ' he won ' +
      peak.pp.toFixed(1) + '%, and that was enough to win ' + peak.wp.toFixed(0) + '% of his matches (' + peak.w + '–' + (peak.n - peak.w) + ').'));

    var W = MOB ? 420 : 960, H = 470, ml = MOB ? 40 : 56, mr = MOB ? 12 : 24, top1 = 20, h1 = 170, top2 = 250, h2 = 150;
    var x = scale(S[0].y - 0.5, S[S.length - 1].y + 0.5, ml, W - mr);
    var y1 = scale(30, 100, top1 + h1, top1);
    var y2 = scale(46, 58, top2 + h2, top2);
    var svg = svgRoot(W, H);
    var g = sv('g', { class: 'axis' });
    [30, 50, 75, 100].forEach(function (t) {
      add(g, sv('line', { x1: ml, x2: W - mr, y1: y1(t), y2: y1(t), 'stroke-dasharray': t === 50 ? '2 3' : null }));
      add(g, sv('text', { x: ml - 10, y: y1(t) + 4, 'text-anchor': 'end' }, t + '%'));
    });
    [46, 50, 54, 58].forEach(function (t) {
      add(g, sv('line', { x1: ml, x2: W - mr, y1: y2(t), y2: y2(t), stroke: t === 50 ? C.gold : C.rule, 'stroke-dasharray': t === 50 ? '3 3' : null }));
      add(g, sv('text', { x: ml - 10, y: y2(t) + 4, 'text-anchor': 'end' }, t + '%'));
    });
    add(svg, g);
    add(svg, sv('text', { x: ml, y: top1 - 6, class: 'lbl-ink' }, 'Matches won'));
    add(svg, sv('text', { x: ml, y: top2 - 10, class: 'lbl-ink' }, 'Points won'));
    add(svg, sv('text', { x: x(2009), y: y2(50) - 6, 'text-anchor': 'middle', class: 'callout-sm', fill: C.gold }, 'even: 50% of points'));

    function path(acc, ys) {
      var d = '';
      S.forEach(function (s) { var v = acc(s); if (v === null) return; d += (d ? 'L' : 'M') + x(s.y).toFixed(1) + ',' + ys(v).toFixed(1); });
      return d;
    }
    var p1 = sv('path', { d: path(function (s) { return s.wp; }, y1), fill: 'none', stroke: C.grass, 'stroke-width': 2.2 });
    var p2 = sv('path', { d: path(function (s) { return s.pp; }, y2), fill: 'none', stroke: C.grass, 'stroke-width': 2.2 });
    add(svg, p1, p2);

    S.forEach(function (s) {
      var tipLines = [s.w + '–' + (s.n - s.w) + ' (' + s.wp.toFixed(1) + '% of matches)',
        (s.pp ? s.pp.toFixed(1) + '% of points won' : 'no point data'), s.titles + ' titles' + (s.slams ? ', ' + s.slams + ' major' + (s.slams > 1 ? 's' : '') : '')];
      var isPeak = s.y === peak.y;
      var c1 = sv('circle', { cx: x(s.y), cy: y1(s.wp), r: isPeak ? 5 : 3.2, fill: isPeak ? C.grass : C.paper, stroke: C.grass, 'stroke-width': 1.6 });
      add(svg, c1);
      if (s.pp !== null) add(svg, sv('circle', { cx: x(s.y), cy: y2(s.pp), r: isPeak ? 5 : 3.2, fill: isPeak ? C.grass : C.paper, stroke: C.grass, 'stroke-width': 1.6 }));
      // majors as gold dots under the chart
      for (var i = 0; i < s.slams; i++) add(svg, sv('circle', { cx: x(s.y), cy: H - 46 + i * 9, r: 3.4, fill: C.gold }));
      add(svg, sv('text', { x: x(s.y), y: H - 52 - 6, 'text-anchor': 'middle', class: 'lbl' }, (MOB ? s.y % 4 === 0 : (s.y % 2 === 0 || s.y === S[0].y)) ? String(s.y) : ''));
      var hit = sv('rect', { x: x(s.y) - (MOB ? 8 : 18), y: top1, width: MOB ? 16 : 36, height: H - top1 - 20, fill: 'transparent' });
      hit.addEventListener('mousemove', function (e) { showTip(e, String(s.y), tipLines); });
      hit.addEventListener('mouseleave', hideTip);
      add(svg, hit);
    });
    add(svg, sv('text', { x: ml - 10, y: H - 42, 'text-anchor': 'end', class: 'lbl', fill: C.gold }, 'Majors'));
    add(svg, sv('text', { x: x(peak.y) + 10, y: y1(peak.wp) - 10, class: 'callout' }, peak.y + ': ' + peak.w + '–' + (peak.n - peak.w)));
    add(svg, sv('text', { x: x(peak.y) + 10, y: y2(peak.pp) - 12, class: 'callout' }, peak.pp.toFixed(1) + '% of points'));

    var chart = el('div', 'chart');
    add(chart, svg);
    add(sec, reveal(chart));
    add(sec, el('p', 'note', 'Each gold dot is a major title. Points won is serve points won plus return points won, out of all points played. Hover over a season for its record.'));
    return sec;
  }

  // II. The Margin — mirrored dot histogram (wins above, losses below)
  // Comparison data (derived from TennisMyLife ATP match stats, tour level incl. Davis Cup/Olympics, walkovers excluded).
  var RIVAL_SUMMARY = [{"name": "Roger Federer", "n": 1452, "w": 1190, "med": 55.3, "band": 47.0, "lostMore": 36, "wonFewer": 11, "ptsPct": 54.1, "y0": 1998, "y1": 2021}, {"name": "Rafael Nadal", "n": 1281, "w": 1056, "med": 55.6, "band": 40.4, "lostMore": 18, "wonFewer": 26, "ptsPct": 54.4, "y0": 2002, "y1": 2024}, {"name": "Novak Djokovic", "n": 1375, "w": 1146, "med": 55.4, "band": 44.5, "lostMore": 19, "wonFewer": 18, "ptsPct": 54.4, "y0": 2004, "y1": 2026}, {"name": "Andy Murray", "n": 964, "w": 705, "med": 53.8, "band": 45.4, "lostMore": 18, "wonFewer": 18, "ptsPct": 52.7, "y0": 2005, "y1": 2024}, {"name": "Pete Sampras", "n": 768, "w": 621, "med": 54.6, "band": 51.2, "lostMore": 20, "wonFewer": 19, "ptsPct": 53.6, "y0": 1991, "y1": 2002}, {"name": "Andre Agassi", "n": 842, "w": 643, "med": 54.7, "band": 42.2, "lostMore": 15, "wonFewer": 18, "ptsPct": 53.3, "y0": 1991, "y1": 2006}, {"name": "Jannik Sinner", "n": 446, "w": 359, "med": 55.0, "band": 45.7, "lostMore": 11, "wonFewer": 10, "ptsPct": 53.9, "y0": 2019, "y1": 2026}, {"name": "Carlos Alcaraz", "n": 390, "w": 321, "med": 54.6, "band": 52.1, "lostMore": 4, "wonFewer": 7, "ptsPct": 53.7, "y0": 2020, "y1": 2026}];
  var RIVAL_DOTS = {"Rafael Nadal":[[58.82,1],[64.77,1],[62.89,1],[51.09,1],[55.73,1],[63.85,1],[55.9,1],[65.15,1],[55.74,1],[54.39,1],[53.3,1],[62.75,1],[60.0,1],[56.3,1],[52.44,1],[50.29,1],[59.43,1],[60.19,1],[54.7,1],[54.22,1],[58.27,1],[64.52,1],[55.93,1],[57.72,1],[58.77,1],[55.07,1],[61.17,1],[55.88,1],[53.72,1],[76.92,1],[56.52,1],[53.33,1],[58.88,1],[61.0,1],[51.23,1],[51.6,1],[57.46,1],[62.24,1],[61.84,1],[63.31,1],[65.38,1],[54.55,1],[63.89,1],[61.0,1],[51.38,1],[49.77,1],[53.79,1],[51.16,1],[56.35,1],[54.13,1],[57.14,1],[60.0,1],[62.42,1],[58.54,1],[50.61,1],[53.54,1],[57.14,1],[57.63,1],[53.1,1],[59.22,1],[71.43,1],[52.74,1],[57.26,1],[54.94,1],[60.91,1],[59.09,1],[63.22,1],[48.75,1],[57.73,1],[54.39,1],[61.43,1],[58.39,1],[52.33,1],[54.9,1],[51.06,1],[55.56,1],[55.65,1],[58.0,1],[57.41,1],[57.14,1],[63.92,1],[54.48,1],[54.09,1],[60.0,1],[58.29,1],[63.2,1],[54.24,1],[59.41,1],[57.66,1],[61.54,1],[55.86,1],[61.29,1],[66.32,1],[60.19,1],[59.46,1],[64.39,1],[66.12,1],[62.33,1],[55.79,1],[56.72,1],[57.92,1],[52.24,1],[59.13,1],[55.11,1],[55.63,1],[56.25,1],[59.18,1],[61.43,1],[57.53,1],[57.56,1],[62.38,1],[67.9,1],[54.04,1],[54.04,1],[58.04,1],[58.72,1],[50.83,1],[67.47,1],[55.88,1],[53.96,1],[51.07,1],[54.14,1],[53.39,1],[50.79,1],[54.4,1],[58.93,1],[58.6,1],[62.26,1],[58.38,1],[53.0,1],[77.27,1],[66.67,1],[67.48,1],[58.87,1],[61.25,1],[59.26,1],[55.85,1],[50.13,1],[49.86,1],[52.94,1],[53.98,1],[53.66,1],[55.12,1],[64.13,1],[56.25,1],[52.71,1],[56.39,1],[51.77,1],[61.29,1],[61.62,1],[54.4,1],[52.15,1],[57.84,1],[59.43,1],[56.84,1],[55.4,1],[54.6,1],[57.41,1],[64.63,1],[55.47,1],[56.8,1],[61.7,1],[64.08,1],[55.65,1],[57.73,1],[55.15,1],[62.2,1],[55.07,1],[48.98,1],[56.57,1],[58.64,1],[62.41,1],[52.73,1],[62.63,1],[51.16,1],[54.55,1],[55.56,1],[62.77,1],[58.14,1],[55.67,1],[56.88,1],[58.25,1],[50.0,1],[54.07,1],[60.38,1],[55.43,1],[57.89,1],[52.14,1],[74.55,1],[50.85,1],[51.16,1],[54.61,1],[52.99,1],[56.2,1],[55.0,1],[54.48,1],[62.33,1],[56.35,1],[57.0,1],[55.73,1],[52.3,1],[52.94,1],[49.06,1],[63.54,1],[64.2,1],[57.34,1],[56.25,1],[54.31,1],[51.35,1],[56.2,1],[54.39,1],[50.72,1],[50.96,1],[56.25,1],[53.33,1],[60.42,1],[62.94,1],[57.14,1],[60.25,1],[52.96,1],[53.24,1],[51.05,1],[60.99,1],[52.14,1],[53.25,1],[59.7,1],[55.03,1],[50.0,1],[55.61,1],[56.45,1],[58.48,1],[57.48,1],[55.73,1],[59.62,1],[58.01,1],[59.75,1],[59.06,1],[62.7,1],[55.56,1],[62.77,1],[61.86,1],[57.55,1],[52.02,1],[63.22,1],[57.27,1],[54.32,1],[53.33,1],[53.12,1],[54.26,1],[58.88,1],[58.72,1],[67.86,1],[65.88,1],[69.14,1],[61.26,1],[56.25,1],[56.79,1],[61.84,1],[61.64,1],[57.21,1],[58.23,1],[61.54,1],[56.35,1],[58.59,1],[58.64,1],[53.16,1],[62.5,1],[61.31,1],[57.61,1],[52.9,1],[55.0,1],[55.47,1],[62.11,1],[57.32,1],[58.28,1],[56.12,1],[54.89,1],[57.51,1],[51.91,1],[55.47,1],[53.1,1],[57.66,1],[51.9,1],[50.7,1],[49.07,1],[39.29,1],[33.75,1],[38.1,1],[44.67,1],[67.72,1],[55.67,1],[61.69,1],[52.9,1],[58.47,1],[57.67,1],[59.6,1],[61.54,1],[67.44,1],[59.05,1],[58.33,1],[53.54,1],[58.33,1],[58.54,1],[62.75,1],[62.22,1],[60.64,1],[55.2,1],[67.47,1],[61.47,1],[57.69,1],[53.12,1],[54.74,1],[55.09,1],[61.84,1],[62.0,1],[56.52,1],[55.47,1],[58.78,1],[57.07,1],[59.38,1],[55.5,1],[59.48,1],[58.6,1],[51.91,1],[57.43,1],[53.09,1],[49.14,1],[52.45,1],[57.14,1],[53.92,1],[60.0,1],[50.94,1],[54.07,1],[49.26,1],[61.0,1],[63.04,1],[57.69,1],[67.95,1],[51.94,1],[55.05,1],[51.46,1],[55.96,1],[56.2,1],[51.91,1],[54.09,1],[55.43,1],[64.08,1],[50.97,1],[60.25,1],[55.28,1],[60.19,1],[51.7,1],[56.49,1],[64.04,1],[65.48,1],[53.15,1],[56.88,1],[59.41,1],[56.76,1],[57.41,1],[57.02,1],[64.37,1],[54.69,1],[53.03,1],[58.54,1],[56.0,1],[60.87,1],[55.07,1],[57.81,1],[62.89,1],[55.14,1],[60.38,1],[55.68,1],[57.76,1],[56.36,1],[52.15,1],[52.45,1],[55.46,1],[53.78,1],[55.42,1],[55.77,1],[51.16,1],[60.2,1],[65.59,1],[53.3,1],[57.14,1],[54.41,1],[55.2,1],[55.94,1],[56.16,1],[56.48,1],[52.32,1],[52.2,1],[53.64,1],[52.29,1],[54.14,1],[59.8,1],[48.66,1],[54.39,1],[51.74,1],[53.05,1],[62.5,1],[52.94,1],[56.91,1],[58.65,1],[53.77,1],[54.38,1],[53.65,1],[52.08,1],[53.85,1],[56.8,1],[55.56,1],[59.87,1],[60.45,1],[53.97,1],[52.69,1],[54.97,1],[56.8,1],[60.19,1],[61.54,1],[49.53,1],[56.3,1],[51.91,1],[63.33,1],[68.35,1],[65.06,1],[54.01,1],[63.83,1],[61.7,1],[56.2,1],[58.1,1],[68.29,1],[58.25,1],[56.88,1],[56.25,1],[54.17,1],[51.87,1],[55.22,1],[52.72,1],[61.39,1],[65.94,1],[61.07,1],[59.43,1],[71.55,1],[57.36,1],[65.87,1],[52.85,1],[53.42,1],[52.44,1],[59.65,1],[59.63,1],[55.33,1],[62.24,1],[67.95,1],[58.02,1],[61.0,1],[50.0,1],[57.01,1],[54.24,1],[57.03,1],[57.14,1],[55.83,1],[52.3,1],[57.66,1],[55.44,1],[54.84,1],[58.42,1],[53.95,1],[61.46,1],[53.11,1],[54.92,1],[58.56,1],[61.96,1],[55.41,1],[57.48,1],[52.23,1],[60.19,1],[55.45,1],[56.25,1],[57.41,1],[53.08,1],[69.6,1],[60.4,1],[60.71,1],[54.81,1],[52.11,1],[50.74,1],[55.12,1],[58.6,1],[60.0,1],[57.86,1],[66.67,1],[55.86,1],[56.0,1],[65.56,1],[53.99,1],[72.22,1],[52.17,1],[51.18,1],[53.58,1],[56.29,1],[53.85,1],[62.77,1],[59.46,1],[56.72,1],[55.56,1],[55.74,1],[55.96,1],[59.05,1],[68.42,1],[55.62,1],[58.99,1],[53.73,1],[54.55,1],[55.19,1],[50.96,1],[56.03,1],[66.28,1],[54.29,1],[59.62,1],[59.0,1],[55.37,1],[58.04,1],[55.3,1],[54.96,1],[56.88,1],[60.32,1],[54.39,1],[64.13,1],[53.45,1],[60.19,1],[62.37,1],[64.04,1],[58.54,1],[57.81,1],[58.12,1],[58.77,1],[51.5,1],[62.62,1],[56.76,1],[57.52,1],[52.69,1],[63.33,1],[58.95,1],[61.49,1],[58.9,1],[69.49,1],[61.11,1],[62.96,1],[60.65,1],[62.25,1],[61.07,1],[56.61,1],[56.85,1],[63.04,1],[57.85,1],[56.8,1],[55.85,1],[56.25,1],[53.85,1],[59.73,1],[63.45,1],[58.92,1],[56.67,1],[53.3,1],[56.69,1],[56.03,1],[54.02,1],[58.33,1],[68.0,1],[60.67,1],[53.4,1],[55.49,1],[56.52,1],[51.83,1],[56.67,1],[56.36,1],[61.54,1],[54.46,1],[51.35,1],[55.56,1],[56.6,1],[58.59,1],[61.17,1],[57.25,1],[51.66,1],[64.77,1],[56.49,1],[50.51,1],[57.48,1],[53.24,1],[53.67,1],[66.28,1],[58.88,1],[50.0,1],[55.4,1],[58.49,1],[58.59,1],[63.41,1],[59.6,1],[55.08,1],[55.47,1],[59.57,1],[53.57,1],[62.11,1],[60.2,1],[65.17,1],[49.21,1],[51.16,1],[61.29,1],[62.11,1],[54.9,1],[55.67,1],[52.36,1],[58.33,1],[60.0,1],[52.84,1],[58.14,1],[62.5,1],[53.69,1],[55.75,1],[50.52,1],[63.95,1],[60.0,1],[52.35,1],[51.53,1],[53.85,1],[53.16,1],[57.67,1],[64.62,1],[58.82,1],[54.84,1],[65.6,1],[54.84,1],[54.26,1],[56.6,1],[51.97,1],[51.41,1],[55.26,1],[59.55,1],[55.0,1],[55.97,1],[57.25,1],[52.35,1],[60.23,1],[49.08,1],[50.0,1],[59.43,1],[56.52,1],[57.52,1],[58.88,1],[64.84,1],[56.22,1],[53.75,1],[61.49,1],[52.7,1],[55.26,1],[52.33,1],[52.27,1],[53.7,1],[67.07,1],[53.47,1],[56.36,1],[56.9,1],[54.74,1],[58.41,1],[58.04,1],[70.51,1],[61.22,1],[61.76,1],[55.3,1],[67.09,1],[62.14,1],[56.0,1],[55.93,1],[55.17,1],[58.33,1],[60.75,1],[59.29,1],[52.46,1],[60.78,1],[56.92,1],[59.82,1],[56.56,1],[52.0,1],[57.23,1],[59.35,1],[58.25,1],[57.99,1],[59.01,1],[58.56,1],[53.33,1],[57.28,1],[56.64,1],[56.41,1],[63.5,1],[49.71,1],[54.72,1],[57.07,1],[53.79,1],[65.98,1],[57.26,1],[60.78,1],[58.97,1],[52.98,1],[56.3,1],[51.7,1],[55.42,1],[53.48,1],[58.73,1],[59.01,1],[55.26,1],[50.58,1],[60.64,1],[55.65,1],[57.41,1],[53.01,1],[51.2,1],[55.2,1],[54.49,1],[51.15,1],[51.7,1],[54.95,1],[57.98,1],[55.26,1],[47.65,1],[53.57,1],[57.03,1],[59.62,1],[57.48,1],[56.2,1],[59.8,1],[56.2,1],[64.58,1],[56.82,1],[52.54,1],[58.82,1],[56.03,1],[53.76,1],[55.38,1],[57.06,1],[53.15,1],[64.77,1],[61.96,1],[56.0,1],[58.76,1],[49.29,1],[58.56,1],[59.39,1],[52.53,1],[55.98,1],[57.94,1],[54.27,1],[52.61,1],[57.63,1],[50.21,1],[52.51,1],[52.01,1],[56.15,1],[56.68,1],[54.79,1],[55.21,1],[57.6,1],[56.52,1],[54.86,1],[58.65,1],[56.59,1],[57.23,1],[54.79,1],[53.62,1],[56.04,1],[56.14,1],[57.55,1],[58.72,1],[48.48,1],[55.38,1],[54.21,1],[59.66,1],[55.2,1],[64.23,1],[55.17,1],[60.96,1],[58.33,1],[51.64,1],[52.9,1],[69.41,1],[58.82,1],[61.17,1],[52.13,1],[65.48,1],[59.6,1],[54.68,1],[51.22,1],[56.19,1],[71.79,1],[55.73,1],[55.37,1],[62.79,1],[62.24,1],[61.86,1],[60.61,1],[61.0,1],[53.26,1],[64.89,1],[61.22,1],[62.38,1],[55.8,1],[56.16,1],[53.15,1],[63.04,1],[65.87,1],[57.58,1],[65.87,1],[55.61,1],[63.36,1],[51.87,1],[59.48,1],[58.89,1],[53.33,1],[54.32,1],[64.29,1],[58.93,1],[62.12,1],[53.97,1],[55.73,1],[53.68,1],[56.06,1],[58.77,1],[55.56,1],[56.25,1],[55.77,1],[56.73,1],[50.0,1],[54.02,1],[63.27,1],[52.35,1],[59.8,1],[56.64,1],[57.84,1],[57.14,1],[59.5,1],[53.9,1],[56.7,1],[58.76,1],[63.64,1],[57.38,1],[61.29,1],[69.23,1],[63.74,1],[57.95,1],[61.81,1],[51.5,1],[51.44,1],[54.62,1],[67.42,1],[53.93,1],[53.59,1],[57.0,1],[53.93,1],[51.02,1],[57.76,1],[71.05,1],[73.91,1],[56.48,1],[58.16,1],[64.29,1],[62.22,1],[59.81,1],[55.14,1],[50.49,1],[56.62,1],[55.12,1],[55.93,1],[61.11,1],[53.75,1],[50.3,1],[58.62,1],[62.0,1],[55.56,1],[57.37,1],[52.47,1],[58.7,1],[55.25,1],[63.83,1],[50.22,1],[55.74,1],[56.42,1],[53.23,1],[58.93,1],[52.61,1],[51.85,1],[54.39,1],[52.23,1],[55.2,1],[53.64,1],[55.12,1],[53.33,1],[54.02,1],[56.22,1],[56.32,1],[58.1,1],[57.63,1],[59.49,1],[53.33,1],[62.92,1],[62.37,1],[51.75,1],[55.45,1],[60.0,1],[48.84,1],[59.18,1],[55.26,1],[48.65,1],[55.28,1],[56.74,1],[48.88,1],[64.0,1],[48.91,1],[55.71,1],[80.33,1],[63.38,1],[54.0,1],[55.68,1],[65.85,1],[63.83,1],[56.43,1],[55.19,1],[55.93,1],[54.84,1],[59.38,1],[65.88,1],[52.56,1],[60.19,1],[63.64,1],[58.41,1],[63.44,1],[53.48,1],[52.94,1],[65.85,1],[58.88,1],[60.78,1],[61.86,1],[55.38,1],[62.11,1],[61.06,1],[54.41,1],[54.19,1],[61.22,1],[62.24,1],[59.09,1],[54.52,1],[54.66,1],[68.85,1],[58.56,1],[55.43,1],[54.26,1],[52.38,1],[53.98,1],[53.89,1],[59.21,1],[58.1,1],[53.44,1],[50.92,1],[55.26,1],[54.63,1],[56.03,1],[49.0,1],[55.09,1],[60.53,1],[56.48,1],[60.24,1],[62.22,1],[54.31,1],[57.14,1],[56.6,1],[56.62,1],[56.76,1],[59.5,1],[52.04,1],[56.49,1],[59.63,1],[64.57,1],[52.22,1],[62.32,1],[54.76,1],[60.53,1],[55.86,1],[60.68,1],[55.71,1],[50.57,1],[56.33,1],[53.51,1],[58.27,1],[57.98,1],[60.61,1],[69.44,1],[51.28,1],[58.42,1],[54.76,1],[57.8,1],[56.67,1],[65.48,1],[58.54,1],[57.94,1],[67.09,1],[65.52,1],[67.07,1],[51.54,1],[53.04,1],[66.28,1],[61.9,1],[66.2,1],[60.91,1],[55.19,1],[66.67,1],[63.95,1],[61.46,1],[56.88,1],[52.17,1],[50.26,1],[59.04,1],[60.54,1],[59.03,1],[54.55,1],[59.3,1],[52.3,1],[53.28,1],[60.0,1],[68.83,1],[56.41,1],[55.22,1],[55.46,1],[54.14,1],[64.04,1],[58.41,1],[62.24,1],[58.02,1],[53.92,1],[54.08,1],[65.22,1],[52.38,1],[64.13,1],[54.14,1],[55.13,1],[59.26,1],[56.35,1],[53.68,1],[52.74,1],[56.62,1],[55.77,1],[54.72,1],[56.69,1],[58.42,1],[52.2,1],[53.02,1],[48.49,1],[26.15,0],[39.04,0],[47.5,0],[45.95,0],[39.6,0],[41.24,0],[44.17,0],[42.16,0],[45.82,0],[50.6,0],[40.0,0],[48.78,0],[44.03,0],[49.14,0],[42.28,0],[45.31,0],[48.5,0],[47.24,0],[49.48,0],[46.77,0],[46.28,0],[46.62,0],[51.23,0],[45.0,0],[42.14,0],[50.5,0],[47.93,0],[47.23,0],[44.53,0],[43.24,0],[42.86,0],[38.61,0],[45.99,0],[37.5,0],[46.46,0],[45.83,0],[45.31,0],[48.63,0],[47.92,0],[45.69,0],[46.75,0],[49.31,0],[49.18,0],[48.11,0],[47.83,0],[47.06,0],[44.88,0],[45.74,0],[37.32,0],[52.16,0],[42.2,0],[45.26,0],[51.67,0],[48.15,0],[41.28,0],[48.01,0],[45.26,0],[49.48,0],[45.59,0],[50.23,0],[41.73,0],[44.53,0],[39.45,0],[45.1,0],[40.37,0],[36.79,0],[46.38,0],[49.69,0],[42.01,0],[48.02,0],[52.25,0],[49.7,0],[43.9,0],[46.84,0],[43.43,0],[48.08,0],[41.04,0],[49.25,0],[47.39,0],[44.34,0],[50.26,0],[49.67,0],[40.95,0],[50.84,0],[47.22,0],[48.35,0],[41.41,0],[43.14,0],[51.91,0],[39.6,0],[47.06,0],[51.89,0],[44.24,0],[41.84,0],[48.75,0],[44.9,0],[45.58,0],[46.27,0],[34.48,0],[49.72,0],[48.86,0],[46.72,0],[45.75,0],[48.91,0],[47.37,0],[49.12,0],[50.5,0],[45.64,0],[37.63,0],[50.0,0],[43.7,0],[44.96,0],[40.19,0],[45.08,0],[45.34,0],[37.5,0],[43.96,0],[48.29,0],[46.21,0],[43.7,0],[47.51,0],[42.52,0],[48.21,0],[46.41,0],[48.1,0],[44.85,0],[42.42,0],[44.09,0],[45.26,0],[50.9,0],[51.11,0],[41.07,0],[41.9,0],[46.95,0],[49.01,0],[43.88,0],[47.08,0],[41.51,0],[45.45,0],[44.7,0],[43.7,0],[49.37,0],[44.44,0],[45.06,0],[46.9,0],[45.38,0],[44.08,0],[47.24,0],[48.92,0],[42.42,0],[43.36,0],[46.39,0],[36.26,0],[35.23,0],[46.56,0],[39.08,0],[48.6,0],[44.96,0],[44.68,0],[49.6,0],[45.93,0],[44.06,0],[47.74,0],[44.62,0],[49.65,0],[47.89,0],[46.43,0],[46.76,0],[44.92,0],[47.7,0],[44.64,0],[48.91,0],[43.7,0],[49.64,0],[49.33,0],[47.17,0],[46.38,0],[41.32,0],[47.24,0],[49.54,0],[46.63,0],[39.8,0],[48.61,0],[48.37,0],[47.15,0],[38.51,0],[43.55,0],[36.56,0],[47.89,0],[43.85,0],[51.58,0],[44.38,0],[51.58,0],[46.25,0],[48.95,0],[44.72,0],[47.28,0],[51.14,0],[44.44,0],[43.08,0],[42.11,0],[42.37,0],[46.2,0],[46.83,0],[45.11,0],[46.92,0],[44.83,0],[46.33,0],[52.47,0],[43.2,0],[45.52,0],[40.43,0],[47.06,0],[33.33,0],[47.37,0],[48.88,0],[45.45,0],[49.67,0],[44.6,0],[48.86,0],[45.86,0],[48.28,0],[46.43,0],[47.66,0],[44.24,0]],"Novak Djokovic":[[59.63,1],[62.77,1],[61.04,1],[55.67,1],[58.15,1],[53.15,1],[52.71,1],[59.8,1],[57.39,1],[57.04,1],[65.85,1],[55.32,1],[58.95,1],[58.56,1],[54.9,1],[60.4,1],[52.63,1],[56.52,1],[60.0,1],[63.27,1],[62.5,1],[50.98,1],[68.09,1],[80.0,1],[55.06,1],[61.17,1],[54.23,1],[58.47,1],[53.3,1],[64.23,1],[56.37,1],[55.32,1],[52.47,1],[55.38,1],[54.38,1],[57.14,1],[67.11,1],[55.56,1],[56.14,1],[55.17,1],[53.25,1],[64.44,1],[54.55,1],[57.84,1],[55.05,1],[59.81,1],[57.6,1],[53.55,1],[56.8,1],[53.37,1],[52.9,1],[52.56,1],[53.57,1],[62.26,1],[53.28,1],[57.52,1],[51.85,1],[57.83,1],[54.35,1],[48.57,1],[52.56,1],[59.5,1],[53.63,1],[56.0,1],[56.44,1],[52.36,1],[53.42,1],[55.97,1],[55.39,1],[62.41,1],[63.97,1],[57.71,1],[54.59,1],[54.85,1],[51.64,1],[65.17,1],[62.5,1],[63.83,1],[52.02,1],[57.14,1],[54.48,1],[56.25,1],[62.77,1],[51.64,1],[49.66,1],[59.32,1],[56.68,1],[59.24,1],[59.09,1],[52.6,1],[54.24,1],[54.6,1],[55.15,1],[64.62,1],[63.87,1],[60.71,1],[56.0,1],[53.5,1],[53.58,1],[53.76,1],[52.29,1],[51.85,1],[59.6,1],[54.17,1],[53.85,1],[63.64,1],[54.1,1],[50.32,1],[53.28,1],[51.74,1],[59.86,1],[56.13,1],[55.38,1],[63.16,1],[61.32,1],[53.9,1],[61.54,1],[52.96,1],[55.67,1],[57.04,1],[70.67,1],[59.31,1],[56.76,1],[59.28,1],[58.33,1],[61.27,1],[59.29,1],[55.43,1],[53.38,1],[52.73,1],[60.43,1],[61.07,1],[54.79,1],[62.5,1],[55.67,1],[52.73,1],[52.54,1],[58.42,1],[60.81,1],[54.74,1],[54.98,1],[55.68,1],[51.46,1],[55.95,1],[59.43,1],[53.59,1],[54.02,1],[59.02,1],[60.0,1],[62.92,1],[63.22,1],[58.18,1],[57.14,1],[57.14,1],[67.9,1],[51.97,1],[56.52,1],[59.33,1],[53.05,1],[55.47,1],[51.85,1],[53.52,1],[58.04,1],[57.6,1],[49.2,1],[56.43,1],[50.22,1],[54.92,1],[53.85,1],[56.43,1],[50.27,1],[58.56,1],[60.38,1],[60.0,1],[54.48,1],[56.21,1],[60.78,1],[57.06,1],[55.03,1],[54.49,1],[56.2,1],[64.29,1],[55.46,1],[52.49,1],[57.23,1],[59.14,1],[54.84,1],[56.92,1],[56.39,1],[54.48,1],[52.63,1],[59.0,1],[56.52,1],[55.4,1],[53.61,1],[60.36,1],[52.41,1],[53.21,1],[55.98,1],[55.93,1],[58.74,1],[53.46,1],[55.56,1],[54.73,1],[55.47,1],[54.26,1],[56.76,1],[62.5,1],[57.14,1],[52.16,1],[59.09,1],[53.08,1],[57.5,1],[63.64,1],[51.52,1],[57.76,1],[54.29,1],[61.0,1],[57.66,1],[51.38,1],[56.69,1],[69.05,1],[52.22,1],[51.66,1],[52.72,1],[54.47,1],[61.76,1],[50.0,1],[62.5,1],[53.43,1],[51.0,1],[54.17,1],[60.19,1],[53.8,1],[56.34,1],[55.96,1],[53.07,1],[62.84,1],[58.94,1],[53.95,1],[60.67,1],[58.33,1],[59.52,1],[54.37,1],[54.46,1],[59.35,1],[53.61,1],[55.63,1],[56.28,1],[55.81,1],[54.88,1],[58.56,1],[57.84,1],[56.57,1],[52.99,1],[61.97,1],[55.09,1],[56.14,1],[57.0,1],[55.44,1],[55.28,1],[58.08,1],[52.4,1],[55.76,1],[52.48,1],[60.71,1],[61.39,1],[63.1,1],[51.88,1],[50.96,1],[60.94,1],[61.18,1],[54.44,1],[56.84,1],[55.74,1],[55.72,1],[54.88,1],[56.03,1],[56.86,1],[50.25,1],[51.32,1],[50.87,1],[56.9,1],[51.15,1],[54.71,1],[57.66,1],[60.0,1],[56.91,1],[58.1,1],[54.55,1],[49.44,1],[54.95,1],[55.17,1],[57.66,1],[56.36,1],[58.77,1],[60.95,1],[54.67,1],[53.91,1],[58.93,1],[60.61,1],[58.24,1],[61.07,1],[59.39,1],[54.37,1],[59.38,1],[60.25,1],[57.47,1],[53.82,1],[55.94,1],[54.1,1],[60.0,1],[53.44,1],[58.73,1],[55.65,1],[68.35,1],[61.96,1],[59.09,1],[46.97,1],[59.79,1],[57.58,1],[61.9,1],[64.84,1],[55.69,1],[61.62,1],[53.7,1],[53.23,1],[53.5,1],[57.14,1],[63.23,1],[59.65,1],[53.85,1],[46.85,1],[47.77,1],[55.56,1],[50.0,1],[55.98,1],[54.51,1],[52.12,1],[60.99,1],[47.0,1],[51.96,1],[50.53,1],[55.64,1],[54.31,1],[52.78,1],[54.22,1],[53.41,1],[52.66,1],[53.85,1],[60.4,1],[59.0,1],[54.89,1],[58.18,1],[57.83,1],[54.55,1],[51.72,1],[60.71,1],[56.25,1],[60.78,1],[55.65,1],[61.36,1],[63.37,1],[53.6,1],[56.65,1],[51.93,1],[60.0,1],[56.34,1],[57.58,1],[59.75,1],[54.5,1],[57.45,1],[65.1,1],[60.76,1],[52.07,1],[55.04,1],[59.02,1],[56.62,1],[56.28,1],[59.21,1],[50.78,1],[54.26,1],[53.98,1],[53.53,1],[54.96,1],[60.33,1],[55.32,1],[55.56,1],[51.81,1],[64.44,1],[55.0,1],[51.79,1],[59.89,1],[55.87,1],[57.07,1],[52.99,1],[67.11,1],[65.35,1],[62.68,1],[59.23,1],[56.56,1],[55.28,1],[52.66,1],[59.81,1],[59.43,1],[57.14,1],[52.63,1],[56.52,1],[62.38,1],[63.04,1],[51.17,1],[52.51,1],[62.77,1],[57.43,1],[60.26,1],[61.15,1],[57.06,1],[57.95,1],[62.42,1],[54.31,1],[60.26,1],[58.49,1],[53.85,1],[48.34,1],[58.27,1],[56.45,1],[58.4,1],[57.32,1],[55.68,1],[57.38,1],[56.82,1],[57.89,1],[69.74,1],[54.84,1],[58.33,1],[62.26,1],[59.13,1],[55.56,1],[58.97,1],[64.29,1],[51.64,1],[58.65,1],[65.91,1],[38.05,1],[41.96,1],[42.59,1],[59.76,1],[55.41,1],[55.74,1],[69.14,1],[54.07,1],[51.59,1],[63.74,1],[58.77,1],[57.76,1],[54.49,1],[56.99,1],[55.44,1],[53.21,1],[54.36,1],[61.76,1],[61.0,1],[59.32,1],[55.22,1],[62.25,1],[62.68,1],[55.65,1],[60.4,1],[56.73,1],[50.52,1],[57.47,1],[58.33,1],[56.64,1],[53.94,1],[55.03,1],[51.19,1],[51.23,1],[50.54,1],[54.62,1],[54.98,1],[54.73,1],[58.33,1],[56.77,1],[54.73,1],[58.1,1],[53.62,1],[55.93,1],[62.86,1],[56.52,1],[63.75,1],[58.06,1],[55.73,1],[59.46,1],[53.57,1],[51.19,1],[59.26,1],[57.69,1],[59.8,1],[62.63,1],[60.19,1],[58.71,1],[55.16,1],[53.91,1],[56.52,1],[66.93,1],[55.02,1],[55.68,1],[65.26,1],[55.81,1],[54.07,1],[60.4,1],[56.78,1],[58.33,1],[55.98,1],[60.93,1],[52.82,1],[51.98,1],[63.64,1],[53.06,1],[54.88,1],[56.88,1],[54.21,1],[72.37,1],[63.21,1],[56.52,1],[54.3,1],[56.67,1],[52.88,1],[55.68,1],[55.56,1],[53.7,1],[65.26,1],[61.7,1],[53.93,1],[52.9,1],[61.29,1],[63.64,1],[58.33,1],[64.96,1],[55.61,1],[54.82,1],[56.5,1],[52.56,1],[53.54,1],[55.08,1],[63.04,1],[66.27,1],[54.6,1],[63.33,1],[57.5,1],[53.38,1],[58.18,1],[55.28,1],[57.4,1],[59.71,1],[57.47,1],[55.03,1],[57.14,1],[51.76,1],[59.48,1],[65.06,1],[61.0,1],[56.1,1],[50.48,1],[55.75,1],[55.09,1],[55.41,1],[53.23,1],[57.39,1],[57.14,1],[60.13,1],[55.41,1],[59.77,1],[58.96,1],[53.1,1],[55.08,1],[59.38,1],[58.9,1],[53.55,1],[56.71,1],[54.03,1],[51.75,1],[54.38,1],[63.22,1],[54.21,1],[55.08,1],[54.61,1],[50.33,1],[57.28,1],[52.36,1],[70.64,1],[63.91,1],[53.5,1],[54.17,1],[53.65,1],[68.03,1],[50.34,1],[63.11,1],[64.37,1],[63.22,1],[58.49,1],[59.05,1],[62.5,1],[57.89,1],[55.56,1],[66.28,1],[61.82,1],[56.82,1],[57.14,1],[53.85,1],[54.97,1],[57.39,1],[65.22,1],[55.12,1],[58.59,1],[53.91,1],[59.06,1],[61.18,1],[56.71,1],[63.58,1],[55.47,1],[61.29,1],[54.89,1],[57.14,1],[50.7,1],[62.86,1],[53.88,1],[50.25,1],[56.14,1],[56.78,1],[55.56,1],[60.4,1],[70.51,1],[72.22,1],[55.0,1],[56.2,1],[57.75,1],[51.85,1],[51.49,1],[55.76,1],[58.24,1],[61.38,1],[53.45,1],[62.77,1],[53.77,1],[52.78,1],[61.74,1],[53.1,1],[54.64,1],[53.51,1],[55.14,1],[50.72,1],[50.82,1],[52.97,1],[50.57,1],[62.33,1],[65.67,1],[60.96,1],[57.06,1],[52.63,1],[63.83,1],[57.94,1],[56.91,1],[58.33,1],[65.52,1],[57.03,1],[51.96,1],[58.27,1],[55.0,1],[54.62,1],[57.38,1],[62.5,1],[58.93,1],[66.67,1],[63.44,1],[65.06,1],[59.2,1],[62.11,1],[62.14,1],[57.38,1],[54.74,1],[65.52,1],[59.49,1],[60.98,1],[55.28,1],[52.85,1],[57.23,1],[58.38,1],[55.41,1],[61.68,1],[64.42,1],[59.75,1],[52.86,1],[57.05,1],[54.05,1],[55.34,1],[54.71,1],[54.25,1],[63.37,1],[57.14,1],[60.53,1],[52.17,1],[57.94,1],[52.23,1],[59.63,1],[60.61,1],[63.44,1],[55.64,1],[53.46,1],[55.41,1],[54.68,1],[51.54,1],[50.88,1],[50.22,1],[62.84,1],[53.69,1],[57.84,1],[54.26,1],[57.41,1],[59.62,1],[55.71,1],[57.29,1],[54.04,1],[55.94,1],[58.68,1],[51.9,1],[60.78,1],[56.41,1],[56.71,1],[60.53,1],[59.06,1],[59.63,1],[55.76,1],[59.13,1],[55.08,1],[51.31,1],[56.78,1],[53.44,1],[57.39,1],[59.62,1],[50.6,1],[63.64,1],[56.59,1],[60.4,1],[58.12,1],[52.06,1],[52.58,1],[57.5,1],[52.5,1],[56.3,1],[51.69,1],[50.97,1],[55.62,1],[59.74,1],[52.84,1],[54.17,1],[54.19,1],[52.07,1],[54.93,1],[54.76,1],[56.07,1],[68.75,1],[56.61,1],[57.89,1],[57.3,1],[56.06,1],[54.47,1],[55.94,1],[53.66,1],[54.55,1],[62.12,1],[62.5,1],[56.54,1],[55.38,1],[58.28,1],[64.38,1],[56.77,1],[51.1,1],[57.84,1],[67.46,1],[52.45,1],[61.11,1],[58.09,1],[62.0,1],[53.9,1],[54.81,1],[57.93,1],[57.66,1],[55.94,1],[65.91,1],[65.59,1],[57.84,1],[55.17,1],[54.19,1],[56.41,1],[67.5,1],[56.12,1],[61.11,1],[58.82,1],[53.52,1],[61.59,1],[60.65,1],[52.14,1],[58.0,1],[55.9,1],[58.76,1],[63.95,1],[56.63,1],[54.26,1],[51.63,1],[61.05,1],[50.53,1],[65.88,1],[59.46,1],[69.14,1],[64.54,1],[58.29,1],[63.51,1],[70.49,1],[56.93,1],[50.0,1],[65.91,1],[53.37,1],[66.67,1],[56.56,1],[58.49,1],[63.04,1],[58.93,1],[53.97,1],[57.58,1],[53.44,1],[56.2,1],[58.15,1],[58.41,1],[53.59,1],[54.36,1],[50.26,1],[51.96,1],[54.9,1],[59.46,1],[56.3,1],[63.92,1],[63.04,1],[55.48,1],[61.82,1],[52.23,1],[65.62,1],[58.52,1],[54.13,1],[52.17,1],[56.12,1],[51.34,1],[54.07,1],[52.98,1],[53.85,1],[65.22,1],[56.8,1],[57.94,1],[60.23,1],[52.9,1],[63.04,1],[65.43,1],[54.62,1],[65.12,1],[58.18,1],[57.62,1],[52.24,1],[61.11,1],[50.7,1],[55.93,1],[48.48,1],[52.02,1],[54.69,1],[55.0,1],[48.66,1],[53.77,1],[55.5,1],[50.0,1],[53.71,1],[54.1,1],[54.14,1],[59.01,1],[53.68,1],[53.27,1],[51.66,1],[50.85,1],[61.62,1],[57.41,1],[61.0,1],[52.05,1],[57.58,1],[49.49,1],[60.74,1],[51.97,1],[60.84,1],[53.2,1],[56.35,1],[55.32,1],[56.88,1],[58.1,1],[52.12,1],[55.22,1],[61.39,1],[52.97,1],[51.18,1],[54.81,1],[57.89,1],[55.75,1],[59.14,1],[50.6,1],[51.34,1],[65.12,1],[51.98,1],[61.8,1],[53.93,1],[55.8,1],[53.96,1],[50.2,1],[54.05,1],[55.22,1],[55.02,1],[59.13,1],[56.72,1],[55.74,1],[70.21,1],[50.96,1],[56.57,1],[58.41,1],[53.46,1],[58.56,1],[51.88,1],[59.26,1],[49.32,1],[59.79,1],[62.11,1],[60.82,1],[51.83,1],[53.93,1],[55.45,1],[52.6,1],[50.0,1],[66.67,1],[62.32,1],[68.64,1],[55.75,1],[55.56,1],[53.33,1],[52.3,1],[56.67,1],[54.01,1],[57.6,1],[58.97,1],[56.8,1],[56.67,1],[58.26,1],[54.39,1],[57.98,1],[55.15,1],[57.46,1],[56.3,1],[55.13,1],[57.36,1],[53.25,1],[57.76,1],[52.69,1],[54.49,1],[55.94,1],[57.39,1],[54.04,1],[60.0,1],[54.2,1],[58.38,1],[57.74,1],[62.84,1],[53.14,1],[54.15,1],[57.06,1],[59.87,1],[56.65,1],[57.32,1],[62.24,1],[58.39,1],[58.77,1],[52.99,1],[52.82,1],[58.02,1],[63.64,1],[56.93,1],[65.22,1],[61.11,1],[58.72,1],[69.3,1],[61.9,1],[61.27,1],[58.4,1],[54.59,1],[55.44,1],[57.23,1],[59.13,1],[64.37,1],[58.16,1],[55.4,1],[64.37,1],[60.82,1],[58.77,1],[58.12,1],[51.29,1],[54.14,1],[52.2,1],[55.56,1],[54.27,1],[50.26,1],[51.63,1],[56.91,1],[54.63,1],[63.77,1],[61.33,1],[61.45,1],[57.81,1],[54.4,1],[51.16,1],[52.08,1],[49.5,1],[52.56,1],[52.08,1],[51.3,1],[57.36,1],[52.76,1],[61.76,1],[63.74,1],[58.18,1],[52.76,1],[56.84,1],[53.77,1],[53.67,1],[52.55,1],[54.31,1],[53.7,1],[57.22,1],[57.14,1],[50.64,1],[61.84,1],[53.8,1],[56.67,1],[59.8,1],[55.56,1],[55.73,1],[50.16,1],[54.25,1],[55.56,1],[57.82,1],[57.89,1],[52.41,1],[61.68,1],[58.77,1],[57.58,1],[57.94,1],[60.0,1],[64.04,1],[59.43,1],[58.72,1],[52.47,1],[59.32,1],[51.11,1],[54.87,1],[57.27,1],[58.06,1],[64.12,1],[54.68,1],[58.33,1],[60.78,1],[57.22,1],[51.74,1],[57.61,1],[59.22,1],[51.02,1],[60.5,1],[55.29,1],[58.88,1],[65.82,1],[68.75,1],[64.08,1],[57.14,1],[51.76,1],[53.8,1],[71.43,1],[65.88,1],[57.76,1],[57.14,1],[56.73,1],[53.17,1],[60.2,1],[57.55,1],[53.79,1],[55.46,1],[70.42,1],[54.49,1],[52.9,1],[54.89,1],[64.52,1],[57.26,1],[62.5,1],[53.09,1],[53.08,1],[62.59,1],[59.38,1],[56.46,1],[60.26,1],[63.97,1],[58.49,1],[51.97,1],[59.74,1],[52.57,1],[55.73,1],[53.67,1],[55.91,1],[55.28,1],[62.11,1],[60.26,1],[52.51,1],[60.0,1],[59.05,1],[52.02,1],[55.56,1],[76.67,1],[66.67,1],[59.01,1],[55.78,1],[55.13,1],[52.96,1],[54.48,1],[52.81,1],[63.95,1],[53.42,1],[57.89,1],[54.34,1],[51.31,1],[63.77,1],[54.45,1],[50.42,1],[60.22,1],[46.32,1],[51.29,1],[51.18,1],[59.38,1],[54.49,1],[45.74,0],[45.83,0],[47.45,0],[44.23,0],[48.77,0],[45.45,0],[48.84,0],[44.27,0],[45.99,0],[47.28,0],[51.25,0],[45.78,0],[47.89,0],[50.0,0],[49.07,0],[47.13,0],[51.56,0],[42.08,0],[39.81,0],[43.55,0],[49.15,0],[47.5,0],[49.4,0],[49.21,0],[45.6,0],[47.87,0],[45.24,0],[46.94,0],[43.56,0],[47.01,0],[39.39,0],[42.54,0],[45.4,0],[44.85,0],[51.02,0],[45.74,0],[43.92,0],[48.39,0],[46.58,0],[41.84,0],[46.88,0],[49.32,0],[43.51,0],[49.19,0],[48.73,0],[47.59,0],[45.16,0],[49.7,0],[50.0,0],[51.1,0],[48.97,0],[47.25,0],[43.62,0],[49.43,0],[46.76,0],[51.32,0],[47.53,0],[50.25,0],[48.63,0],[50.45,0],[44.97,0],[48.45,0],[46.95,0],[43.7,0],[49.78,0],[50.0,0],[45.99,0],[46.74,0],[46.09,0],[43.87,0],[50.0,0],[42.61,0],[47.24,0],[46.67,0],[42.42,0],[44.51,0],[43.22,0],[49.5,0],[46.34,0],[49.72,0],[46.24,0],[43.21,0],[48.77,0],[50.0,0],[44.06,0],[49.4,0],[49.54,0],[39.64,0],[48.12,0],[49.25,0],[43.69,0],[47.85,0],[50.78,0],[49.28,0],[46.88,0],[50.7,0],[49.1,0],[46.03,0],[46.72,0],[44.66,0],[41.88,0],[43.36,0],[50.27,0],[53.5,0],[42.55,0],[49.7,0],[44.39,0],[49.42,0],[48.0,0],[49.29,0],[46.81,0],[46.06,0],[48.7,0],[43.57,0],[42.86,0],[51.27,0],[48.47,0],[43.1,0],[47.15,0],[41.84,0],[44.9,0],[50.63,0],[46.31,0],[43.14,0],[50.0,0],[43.48,0],[47.71,0],[46.24,0],[49.83,0],[48.63,0],[47.89,0],[46.67,0],[50.92,0],[46.85,0],[47.68,0],[50.0,0],[42.48,0],[42.86,0],[43.55,0],[43.01,0],[49.75,0],[42.34,0],[49.61,0],[49.49,0],[47.16,0],[45.71,0],[49.48,0],[49.55,0],[45.74,0],[41.99,0],[48.58,0],[45.96,0],[43.1,0],[47.1,0],[40.71,0],[49.78,0],[41.44,0],[49.4,0],[46.21,0],[50.53,0],[40.91,0],[46.4,0],[45.64,0],[39.42,0],[44.8,0],[41.28,0],[46.21,0],[46.52,0],[46.38,0],[48.51,0],[45.04,0],[42.54,0],[47.31,0],[49.73,0],[42.24,0],[42.06,0],[44.04,0],[48.89,0],[51.11,0],[46.02,0],[39.61,0],[40.0,0],[47.8,0],[48.22,0],[44.34,0],[52.47,0],[37.21,0],[46.32,0],[46.85,0],[48.13,0],[47.22,0],[42.74,0],[49.21,0],[52.66,0],[43.17,0],[48.29,0],[47.57,0],[49.52,0],[49.4,0],[45.45,0],[45.7,0],[39.64,0],[49.55,0],[46.38,0],[49.42,0],[45.73,0],[44.21,0],[47.89,0],[45.53,0],[46.67,0],[46.62,0],[47.59,0],[46.27,0],[44.72,0],[41.07,0],[47.57,0],[39.29,0],[47.98,0],[39.42,0],[50.0,0],[29.06,0],[45.99,0],[46.95,0],[49.62,0],[41.88,0],[51.32,0],[47.02,0],[46.78,0],[45.19,0]]};

  function renderMargin(M) {
    var stat = M.filter(function (m) { return m.hasPts; });
    var wonLess = stat.filter(function (m) { return m.win && m.pp < 50; });
    var lostMore = stat.filter(function (m) { return !m.win && m.pp > 50; });
    var median = stat.map(function (m) { return m.pp; }).sort(function (a, b) { return a - b; })[Math.floor(stat.length / 2)];
    var near = stat.filter(function (m) { return m.pp >= 48 && m.pp <= 56; }).length;

    var sec = el('section', 'ch');
    add(sec, chapterHead('II', 'The margin', 'He dominated, but he rarely won by much.',
      'Each dot is one match, placed by the share of points he won. Wins stack above the line and losses below. In his median match he won ' +
      median.toFixed(1) + '% of the points, and ' + Math.round(100 * near / stat.length) + '% of all his matches fell between 48% and 56%. He lost ' + lostMore.length +
      ' matches in which he won more points than his opponent, and won ' + wonLess.length + ' in which he won fewer.'));

    // One dot chart. Picking Nadal or Djokovic moves Federer's dots into that player's career shape.
    var SUM = {}; RIVAL_SUMMARY.forEach(function (s) { SUM[s.name] = s; });
    var lo = 30, hi = 80, MB = MOB ? 2 : 1;   // phones: 2-point bins so the dots stay big enough to tap
    var PL = [
      { name: 'Roger Federer', short: 'Federer', list: stat.map(function (m) { return { pp: m.pp, win: !!m.win, m: m }; }) },
      { name: 'Rafael Nadal', short: 'Nadal', list: RIVAL_DOTS['Rafael Nadal'].map(function (d) { return { pp: d[0], win: !!d[1] }; }) },
      { name: 'Novak Djokovic', short: 'Djokovic', list: RIVAL_DOTS['Novak Djokovic'].map(function (d) { return { pp: d[0], win: !!d[1] }; }) }];
    function binsOf(list) {
      var bins = {};
      list.forEach(function (d) { var b = MB * Math.floor(Math.max(lo, Math.min(hi - 0.001, d.pp)) / MB); var k = b + (d.win ? 'W' : 'L'); (bins[k] || (bins[k] = [])).push(d); });
      return bins;
    }
    var maxW = 0, maxL = 0;
    PL.forEach(function (p) {
      p.bins = binsOf(p.list);
      Object.keys(p.bins).forEach(function (k) { if (k.slice(-1) === 'W') maxW = Math.max(maxW, p.bins[k].length); else maxL = Math.max(maxL, p.bins[k].length); });
    });
    var W = MOB ? 440 : 960, ml = MOB ? 16 : 30, mr = MOB ? 16 : 30;
    var bw = (W - ml - mr) / (hi - lo) * MB;
    var perRow = 2, sp = bw / perRow;
    while (Math.ceil(maxW / perRow) * sp > (MOB ? 360 : 330) && perRow < 6) { perRow++; sp = bw / perRow; }
    var r = Math.max(1.4, sp * (MOB ? 0.44 : 0.38));
    var upH = Math.ceil(maxW / perRow) * sp, dnH = Math.ceil(maxL / perRow) * sp;
    var axisY = upH + 62, H = axisY + dnH + 70;
    var x = scale(lo, hi, ml, W - mr);

    // every match gets a slot in its bin stack; wins and losses each sorted by share of points
    PL.forEach(function (p) {
      p.W = []; p.L = [];
      Object.keys(p.bins).forEach(function (k) {
        var b = Number(k.slice(0, -1)), won = k.slice(-1) === 'W';
        p.bins[k].sort(function (a, c) { return a.pp - c.pp; }).forEach(function (d, i) {
          var col = i % perRow, row = Math.floor(i / perRow);
          d.x = x(b) + sp * (col + 0.5);
          d.y = won ? axisY - 6 - row * sp - sp / 2 : axisY + 6 + row * sp + sp / 2;
          d.hl = (won && d.pp < 50) || (!won && d.pp > 50);
          (won ? p.W : p.L).push(d);
        });
      });
      p.W.sort(function (a, c) { return a.pp - c.pp; }); p.L.sort(function (a, c) { return a.pp - c.pp; });
      var pp = p.list.map(function (d) { return d.pp; }).sort(function (a, c) { return a - c; });
      p.med = pp[Math.floor(pp.length / 2)];
      p.against = p.list.filter(function (d) { return d.hl; }).length;
    });

    var svg = svgRoot(W, H);
    add(svg, sv('line', { x1: x(50), x2: x(50), y1: 18, y2: H - 30, stroke: C.gold, 'stroke-dasharray': '3 3' }));
    add(svg, sv('text', { x: x(50) - 6, y: 14, 'text-anchor': 'end', class: 'callout-sm', fill: C.gold }, '← won fewer points'));
    add(svg, sv('text', { x: x(50) + 6, y: 14, class: 'callout-sm', fill: C.gold }, 'won more points →'));
    add(svg, sv('line', { x1: ml, x2: W - mr, y1: axisY, y2: axisY, stroke: C.ink, 'stroke-width': 0.8 }));
    for (var t = lo; t <= hi; t += 5) add(svg, sv('line', { x1: x(t), x2: x(t), y1: axisY - 3, y2: axisY + 3, stroke: C.ink, 'stroke-width': 0.8 }));
    for (var t2 = lo; t2 <= hi; t2 += 5) add(svg, sv('text', { x: x(t2), y: axisY + dnH + 30, 'text-anchor': 'middle', class: 'lbl' }, t2 + '%'));
    add(svg, sv('text', { x: ml, y: axisY - 12, class: 'lbl-ink', fill: C.grass }, 'WINS'));
    add(svg, sv('text', { x: ml, y: axisY + 22, class: 'lbl-ink', fill: C.purple }, 'LOSSES'));
    var who = sv('text', { x: ml, y: 58, class: 'mar-who' }, 'Federer');
    var whoSub = sv('text', { x: ml, y: 80, class: 'callout-sm' }, '');
    add(svg, who, whoSub);
    var fedMedLine = sv('line', { x1: x(PL[0].med), x2: x(PL[0].med), y1: 32, y2: axisY, stroke: C.grass, 'stroke-width': 1, 'stroke-dasharray': '2 3', opacity: 0 });
    var medLine = sv('line', { x1: x(PL[0].med), x2: x(PL[0].med), y1: 32, y2: axisY, stroke: C.ink, 'stroke-width': 0.6, opacity: 0.5 });
    var medLab = sv('text', { x: x(PL[0].med) + 8, y: 42, class: 'callout' }, '');
    add(svg, fedMedLine, medLine, medLab);

    // a fixed pool of circles (Federer's matches); for another player each circle moves to the same rank in his career
    var curName = 'Roger Federer';
    var pool = [];
    var dotsG = sv('g', {});
    ['W', 'L'].forEach(function (grp) {
      PL[0][grp].forEach(function (d, i) {
        var c = sv('circle', { cx: d.x.toFixed(1), cy: d.y.toFixed(1), r: r.toFixed(2), fill: d.hl ? C.gold : (grp === 'W' ? C.grass : C.purple), opacity: d.hl ? 1 : 0.82 });
        var o = { c: c, grp: grp, i: i, n: PL[0][grp].length, x: d.x, y: d.y, d: d };
        c.addEventListener('mousemove', function (e) {
          var q = o.d;
          if (q.m) showTip(e, q.m.tournament + ' ' + q.m.season + ' · ' + q.m.round, [(q.m.win ? 'W' : 'L') + ' vs ' + q.m.opponent + '  ' + q.m.score,
            q.pp.toFixed(1) + '% of points (' + fmt(q.m.points_won) + ' of ' + fmt(q.m.points_total) + ')']);
          else showTip(e, curName.split(' ').slice(-1)[0] + ' ' + (q.win ? 'win' : 'loss'), [q.pp.toFixed(1) + '% of points']);
        });
        c.addEventListener('mouseleave', hideTip);
        pool.push(o); add(dotsG, c);
      });
    });
    add(svg, dotsG);
    // phones: dots are too small to tap one by one, so each column is one tap target with its counts
    var binHits = null;
    if (MOB) {
      binHits = sv('g', {}); add(svg, binHits);
      var drawHits = function (p) {
        binHits.textContent = '';
        for (var b0 = lo; b0 < hi; b0 += MB) (function (b0) {
          var w = p.bins[b0 + 'W'] || [], l = p.bins[b0 + 'L'] || [];
          if (!w.length && !l.length) return;
          var against = w.concat(l).filter(function (d) { return d.hl; }).length;
          var hr = sv('rect', { x: x(b0), y: 20, width: bw, height: H - 40, fill: 'transparent' });
          hr.addEventListener('mousemove', function (e) {
            showTip(e, p.short + ' · ' + b0 + '–' + (b0 + MB) + '% of points', [w.length + ' win' + (w.length === 1 ? '' : 's') + ', ' + l.length + ' loss' + (l.length === 1 ? '' : 'es'),
              against ? against + ' went against the points' : 'none went against the points']);
          });
          hr.addEventListener('mouseleave', hideTip);
          add(binHits, hr);
        })(b0);
      };
    }

    var anim = null;
    function ease(t) { return t < 0.5 ? 4 * t * t * t : 1 - Math.pow(-2 * t + 2, 3) / 2; }
    function show(p, instant) {
      curName = p.name;
      if (binHits) drawHits(p);
      var fed = p === PL[0];
      pool.forEach(function (o) {
        var arr = p[o.grp];
        var d = arr[Math.min(arr.length - 1, Math.floor(o.i * arr.length / o.n))];
        o.fx = o.x; o.fy = o.y; o.tx = d.x; o.ty = d.y; o.d = d;
        o.delay = (d.x - ml) / (W - ml - mr) * 0.25;
        o.col = d.hl ? C.gold : (o.grp === 'W' ? C.grass : C.purple); o.op = d.hl ? 1 : 0.82;
      });
      var m0 = parseFloat(medLine.getAttribute('x1')), m1 = x(p.med);
      medLab.textContent = (fed ? 'Median match: ' : p.short + '’s median match: ') + p.med.toFixed(1) + '% of points';
      who.textContent = p.short;
      whoSub.textContent = fmt(p.list.length) + ' matches · ' + p.against + ' went against the points';
      fedMedLine.setAttribute('opacity', fed ? 0 : 0.8);
      if (anim) cancelAnimationFrame(anim);
      var t0 = null, dur = instant || reducedMotion() ? 1 : MOTION.shift;
      function step(ts) {
        if (t0 === null) t0 = ts;
        var g = Math.min(1, (ts - t0) / dur);
        pool.forEach(function (o) {
          var k = ease(Math.max(0, Math.min(1, (g * 1.25 - o.delay))));
          o.x = o.fx + (o.tx - o.fx) * k; o.y = o.fy + (o.ty - o.fy) * k;
          o.c.setAttribute('cx', o.x.toFixed(1)); o.c.setAttribute('cy', o.y.toFixed(1));
          if (k > 0.5 && o.c.getAttribute('fill') !== o.col) { o.c.setAttribute('fill', o.col); o.c.setAttribute('opacity', o.op); }
        });
        var km = ease(g), mx = m0 + (m1 - m0) * km;
        medLine.setAttribute('x1', mx); medLine.setAttribute('x2', mx); medLab.setAttribute('x', mx + 8);
        if (g < 1) anim = requestAnimationFrame(step); else anim = null;
      }
      anim = requestAnimationFrame(step);
      btns.forEach(function (b) { b[1].className = 'seg' + (b[0] === p ? ' on' : ''); });
      lgAgainst.textContent = 'Result went against the points (' + p.against + ' matches)';
      var s = SUM[p.name];
      cap.textContent = fed ? 'Pick Nadal or Djokovic to move these dots into his career. Each dot keeps its place in the order, so you can watch where the shape changes.'
        : p.short + ', ' + s.y0 + '–' + s.y1 + ': median match ' + p.med.toFixed(1) + '% of points (Federer ' + PL[0].med.toFixed(1) + '%, dashed green line). Lost ' + s.lostMore +
          ' matches with more points (Federer ' + SUM['Roger Federer'].lostMore + ') and won ' + s.wonFewer + ' with fewer (Federer ' + SUM['Roger Federer'].wonFewer + ').';
    }

    var chart = el('div', 'chart');
    var bar = el('div', 'rp-viewbar mar-bar');
    add(bar, el('span', 'rp-vl', 'Player'));
    var btns = PL.map(function (p) {
      var b = el('button', 'seg', p.short);
      b.addEventListener('click', function () { if (curName !== p.name) show(p); });
      add(bar, b); return [p, b];
    });
    var lg = el('div', 'legend'), lgAgainst = null;
    [[C.grass, 'Win'], [C.purple, 'Loss'], [C.gold, '']].forEach(function (x2, k) {
      var s = el('span'); var i = el('i'); i.style.background = x2[0]; add(s, i);
      var tn = document.createTextNode(x2[1]); s.appendChild(tn); if (k === 2) lgAgainst = tn; add(lg, s);
    });
    var cap = el('p', 'mar-cap');
    add(chart, bar, svg, lg, cap);
    show(PL[0], true);
    add(sec, reveal(chart));
    add(sec, el('p', 'note', stat.length + ' of ' + M.length + ' Federer matches include serve stats. Most of the matches without them are Davis Cup and Olympic ties. Hover over any dot to see the match.'));

    // ---- when the points and the result disagreed
    var rows = RIVAL_SUMMARY.map(function (s) { var o = {}; for (var k in s) o[k] = s[k]; o.net = s.wonFewer - s.lostMore; return o; })
      .sort(function (a, b) { return b.net - a.net || a.name.localeCompare(b.name); });
    var F = SUM['Roger Federer'], N = SUM['Rafael Nadal'], D = SUM['Novak Djokovic'];
    function sg(v) { return (v > 0 ? '+' : v < 0 ? '−' : '') + Math.abs(v); }
    var others = rows.filter(function (o) { return o.name !== 'Roger Federer'; });
    var oLo = Math.min.apply(null, others.map(function (o) { return o.net; })), oHi = Math.max.apply(null, others.map(function (o) { return o.net; }));
    var sub = el('div', 'mar-sub');
    add(sub, el('h3', null, 'All three dominated points the same way. Where they part is what happened when the points and the result disagreed.'));
    var gap = Math.max(Math.abs(F.med - N.med), Math.abs(F.med - D.med));
    add(sub, el('p', 'lede', 'Federer’s median match (' + F.med.toFixed(1) + '% of points) is within ' + gap.toFixed(1) + ' percentage points of Djokovic’s (' + D.med.toFixed(1) + '%) and Nadal’s (' + N.med.toFixed(1) +
      '%). But he lost ' + F.lostMore + ' matches in which he won more points and won only ' + F.wonFewer + ' in which he won fewer: a net of ' + sg(F.wonFewer - F.lostMore) +
      '. Nadal finished ' + sg(N.wonFewer - N.lostMore) + ' and Djokovic ' + sg(D.wonFewer - D.lostMore) + '. Every other great player below lands between ' + sg(oLo) + ' and ' + sg(oHi) + '.'));
    add(sec, sub);

    var LW = MOB ? 440 : 960, lab = MOB ? 118 : 230, netW = MOB ? 44 : 80, rowH = 44, top = 46;
    var LH = top + rows.length * rowH + 30;
    var cx0 = lab + (LW - lab - netW) / 2, half = (LW - lab - netW) / 2 - 30;
    var maxC = Math.max.apply(null, rows.map(function (o) { return Math.max(o.lostMore, o.wonFewer); }));
    var bx = function (v) { return v / maxC * half; };
    var lsvg = svgRoot(LW, LH);
    add(lsvg, sv('text', { x: cx0 - 8, y: 18, 'text-anchor': 'end', class: 'lbl-ink', fill: C.purple }, MOB ? '← Lost, more pts' : '← Lost despite winning more points'));
    add(lsvg, sv('text', { x: cx0 + 8, y: 18, class: 'lbl-ink', fill: C.grass }, MOB ? 'Won, fewer pts →' : 'Won despite winning fewer points →'));
    add(lsvg, sv('text', { x: LW - 4, y: 18, 'text-anchor': 'end', class: 'lbl-ink' }, 'Net'));
    rows.forEach(function (o, i) {
      var y0 = top + i * rowH, cy = y0 + rowH / 2, fed = o.name === 'Roger Federer';
      var g = sv('g', {});
      if (fed) add(g, sv('rect', { x: 0, y: y0 + 2, width: LW, height: rowH - 4, fill: C.goldSoft, opacity: 0.45, rx: 3 }));
      add(g, sv('text', { x: 10, y: cy - 2, class: fed ? 'lbl-serif' : 'lbl-ink', 'font-weight': fed ? 600 : 400 }, o.name));
      add(g, sv('text', { x: 10, y: cy + 13, class: 'lbl' }, fmt(o.n) + ' matches' + (MOB ? '' : ' · ' + o.y0 + '–' + o.y1)));
      add(g, sv('rect', { x: cx0 - bx(o.lostMore), y: cy - 9, width: bx(o.lostMore), height: 18, fill: C.purple, opacity: fed ? 1 : 0.7 }));
      add(g, sv('rect', { x: cx0, y: cy - 9, width: bx(o.wonFewer), height: 18, fill: C.grass, opacity: fed ? 1 : 0.7 }));
      add(g, sv('text', { x: cx0 - bx(o.lostMore) - 6, y: cy + 4, 'text-anchor': 'end', class: 'lbl-ink' }, o.lostMore));
      add(g, sv('text', { x: cx0 + bx(o.wonFewer) + 6, y: cy + 4, class: 'lbl-ink' }, o.wonFewer));
      add(g, sv('text', { x: LW - 4, y: cy + 5, 'text-anchor': 'end', class: 'mar-net', fill: o.net < -5 ? C.gold : C.ink }, sg(o.net)));
      g.addEventListener('mousemove', function (e) {
        showTip(e, o.name + ' · ' + o.y0 + '–' + o.y1, [fmt(o.n) + ' matches, won ' + (100 * o.w / o.n).toFixed(1) + '%',
          'Median match: ' + o.med.toFixed(1) + '% of points · career: ' + o.ptsPct.toFixed(1) + '%',
          'Lost with more points: ' + o.lostMore + ' · won with fewer: ' + o.wonFewer,
          'Per 100 matches: ' + (100 * o.lostMore / o.n).toFixed(1) + ' vs ' + (100 * o.wonFewer / o.n).toFixed(1)]);
      });
      g.addEventListener('mouseleave', hideTip);
      add(lsvg, g);
    });
    add(lsvg, sv('line', { x1: cx0, x2: cx0, y1: top - 4, y2: LH - 24, stroke: C.ink, 'stroke-width': 0.8 }));
    var lchart = el('div', 'chart');
    add(lchart, lsvg);
    add(sec, reveal(lchart));
    add(sec, el('p', 'note', 'Every player is counted the same way: tour-level matches including Davis Cup and the Olympics, walkovers excluded, matches with serve stats only (recorded from 1991). ' +
      'Matches with exactly 50% of points are not counted on either side. Active players’ careers run to the latest data available.'));
    return sec;
  }

  // III. Leverage — rivals scatter + S-curve of match win rate by points-won bucket
  function renderRivals(M, CH, SH) {
    var stat = M.filter(function (m) { return m.hasPts; });
    // S-curve from all matches: win rate per 1-point bucket
    var buckets = {};
    stat.forEach(function (m) { var b = Math.round(m.pp); var o = buckets[b] || (buckets[b] = { b: b, n: 0, w: 0 }); o.n++; o.w += m.win; });
    var curve = Object.keys(buckets).map(function (k) { return buckets[k]; }).filter(function (o) { return o.n >= 8; }).sort(function (a, b) { return a.b - b.b; });
    var at = function (v) { var o = buckets[v]; return o ? Math.round(100 * o.w / o.n) : null; };
    var at50 = at(50), at54 = at(54);

    var opp = {};
    M.forEach(function (m) {
      var o = opp[m.opponent] || (opp[m.opponent] = { name: m.opponent, n: 0, w: 0, pw: 0, pt: 0 });
      o.n++; o.w += m.win; if (m.hasPts) { o.pw += m.points_won; o.pt += m.points_total; }
    });
    var R = Object.keys(opp).map(function (k) { return opp[k]; }).filter(function (o) { return o.n >= 8 && o.pt > 0; });
    R.forEach(function (o) { o.wp = 100 * o.w / o.n; o.pp = 100 * o.pw / o.pt; });
    var focus = { 'Rafael Nadal': 1, 'Novak Djokovic': 1, 'Andy Murray': 1 };
    var named = { 'Rafael Nadal': 1, 'Novak Djokovic': 1, 'Andy Murray': 1, 'Andy Roddick': 1, 'Lleyton Hewitt': 1, 'Andre Agassi': 1,
      'Stan Wawrinka': 1, 'Juan Martin del Potro': 1, 'David Nalbandian': 1, 'Tomas Berdych': 1 };
    if (MOB) named = { 'Rafael Nadal': 1, 'Novak Djokovic': 1, 'Andy Murray': 1, 'Andy Roddick': 1 };   // phones: label the main rivals only
    var nadal = opp['Rafael Nadal'], djok = opp['Novak Djokovic'], rod = opp['Andy Roddick'];

    var sec = el('section', 'ch');
    add(sec, chapterHead('III', 'Leverage', 'A small edge in points becomes a large edge in matches.',
      'Across his career, matches in which he won 50% of the points went his way ' + at50 + '% of the time. At 54% he won ' + at54 +
      '% of them. Against Nadal (' + nadal.w + '–' + (nadal.n - nadal.w) + ') and Djokovic (' + djok.w + '–' + (djok.n - djok.w) +
      ') he won about half the points: ' + nadal.pp.toFixed(1) + '% and ' + djok.pp.toFixed(1) + '%. Against Roddick he won ' + rod.pp.toFixed(1) +
      '%, and the record was ' + rod.w + '–' + (rod.n - rod.w) + '.'));

    var W = MOB ? 440 : 960, H = MOB ? 500 : 560, ml = MOB ? 40 : 60, mr = MOB ? 14 : 30, mt = 44, mb = 50;
    var X0 = 42, X1 = 62;
    var xs = scale(X0, X1, ml, W - mr), ys = scale(0, 100, H - mb, mt);
    var svg = svgRoot(W, H);
    var g = sv('g', { class: 'axis' });
    [0, 25, 50, 75, 100].forEach(function (t) {
      add(g, sv('line', { x1: ml, x2: W - mr, y1: ys(t), y2: ys(t) }));
      add(g, sv('text', { x: ml - 10, y: ys(t) + 4, 'text-anchor': 'end' }, t + '%'));
    });
    for (var t0 = X0; t0 <= X1; t0 += 2) add(g, sv('text', { x: xs(t0), y: H - mb + 22, 'text-anchor': 'middle' }, t0 + '%'));
    add(svg, g);
    add(svg, sv('line', { x1: xs(50), x2: xs(50), y1: mt, y2: H - mb, stroke: C.gold, 'stroke-dasharray': '3 3' }));
    add(svg, sv('text', { x: (ml + W - mr) / 2, y: H - 8, 'text-anchor': 'middle', class: 'lbl-ink' }, 'Share of points won'));
    add(svg, sv('text', { x: ml, y: 16, class: 'lbl-ink' }, 'Matches won'));

    // S-curve
    var d = '';
    curve.forEach(function (o) { if (o.b < X0 || o.b > X1) return; d += (d ? 'L' : 'M') + xs(o.b).toFixed(1) + ',' + ys(100 * o.w / o.n).toFixed(1); });
    var cp = sv('path', { d: d, fill: 'none', stroke: C.grassSoft, 'stroke-width': 7, 'stroke-linecap': 'round', 'stroke-linejoin': 'round', opacity: 0.8 });
    add(svg, cp);
    add(svg, sv('text', { x: xs(X0) + 6, y: ys(14), class: 'callout-sm', fill: C.grass }, 'Green curve: how often he won the match'));
    add(svg, sv('text', { x: xs(X0) + 6, y: ys(14) + 15, class: 'callout-sm', fill: C.grass }, 'at each share of points (all matches)'));
    if (at54 !== null) {
      add(svg, sv('circle', { cx: xs(54), cy: ys(at54), r: 6, fill: C.grass }));
      add(svg, sv('line', { x1: xs(54), x2: xs(55.4), y1: ys(at54), y2: ys(36), stroke: C.grass, 'stroke-width': 0.8 }));
      if (MOB) { add(svg, sv('text', { x: xs(55.4) + 4, y: ys(36) + 4, class: 'callout' }, '54% of points →')); add(svg, sv('text', { x: xs(55.4) + 4, y: ys(36) + 22, class: 'callout' }, 'won ' + at54 + '% of matches')); }
      else add(svg, sv('text', { x: xs(55.4) + 4, y: ys(36) + 4, class: 'callout' }, '54% of points → won ' + at54 + '% of matches'));
    }
    // opponent tiers on the same curve
    tierPoints(M).forEach(function (t) {
      if (!t.n) return;
      var tx = xs(Math.max(X0, Math.min(X1, t.pp))), ty = ys(t.wp);
      add(svg, sv('rect', { x: tx - 5, y: ty - 5, width: 10, height: 10, fill: C.gold, stroke: C.paper, 'stroke-width': 1.2, transform: 'rotate(45 ' + tx + ' ' + ty + ')' }));
    });
    if (at50 !== null) {
      add(svg, sv('circle', { cx: xs(50), cy: ys(at50), r: 5, fill: C.gold }));
      add(svg, sv('text', { x: xs(50) - 16, y: ys(at50) - 14, 'text-anchor': 'end', class: 'callout' }, '50% of points → won ' + at50 + '%'));
    }

    var boxes = [];
    function free(bx) {
      return boxes.every(function (b) { return bx.x1 < b.x0 || bx.x0 > b.x1 || bx.y1 < b.y0 || bx.y0 > b.y1; });
    }
    R.forEach(function (o) { boxes.push({ x0: xs(o.pp) - 2, x1: xs(o.pp) + 2, y0: ys(o.wp) - 2, y1: ys(o.wp) + 2 }); });
    var labels = [];

    R.sort(function (a, b) { return a.n - b.n; }).forEach(function (o) {
      var isF = focus[o.name], isN = named[o.name];
      var rad = 3 + Math.sqrt(o.n) * 1.25;
      var c = sv('circle', { cx: xs(o.pp), cy: ys(o.wp), r: rad.toFixed(1), fill: isF ? C.purple : (isN ? C.ink : C.grey),
        'fill-opacity': isF ? 0.9 : (isN ? 0.75 : 0.6), stroke: C.paper, 'stroke-width': 1 });
      c.addEventListener('mousemove', function (e) {
        showTip(e, o.name, [o.w + '–' + (o.n - o.w) + ' head-to-head (' + o.wp.toFixed(0) + '%)', o.pp.toFixed(1) + '% of points won']);
      });
      c.addEventListener('mouseleave', hideTip);
      add(svg, c);
      if (isN) labels.push({ o: o, rad: rad, isF: isF });
    });
    // place labels after all circles, nudging to avoid overlaps
    labels.sort(function (a, b) { return b.o.n - a.o.n; }).forEach(function (L) {
      var o = L.o, txt = o.name.split(' ').slice(-1)[0] + ' ' + o.w + '–' + (o.n - o.w);
      var w = txt.length * 6.6, cx = xs(o.pp), cy = ys(o.wp);
      var tries = [[1, 0], [-1, 0], [1, -14], [1, 14], [-1, -14], [-1, 14], [1, -28], [1, 28], [-1, 28], [-1, -28]];
      for (var i = 0; i < tries.length; i++) {
        var side = tries[i][0], dy = tries[i][1];
        var x0 = side > 0 ? cx + L.rad + 4 : cx - L.rad - 4 - w;
        var bx = { x0: x0, x1: x0 + w, y0: cy + dy - 9, y1: cy + dy + 4 };
        if (free(bx) || i === tries.length - 1) {
          boxes.push(bx);
          if (dy) add(svg, sv('line', { x1: cx + side * L.rad, y1: cy, x2: side > 0 ? x0 : x0 + w, y2: cy + dy - 3, stroke: C.muted, 'stroke-width': 0.6 }));
          add(svg, sv('text', { x: x0, y: cy + dy + 1, class: 'lbl-ink', fill: L.isF ? C.purple : C.ink }, txt));
          break;
        }
      }
    });

    var chart = el('div', 'chart');
    add(chart, svg);
    add(sec, reveal(chart));
    add(sec, el('p', 'note', 'Each circle is an opponent he played at least 8 times, sized by the number of matches. Gold diamonds group his opponents by ranking. The green curve shows how often he won a match at each share of points won, across all of his matches. Hover over a circle for the head-to-head.'));
    var tierRow = el('div', 'tiers'); add(tierRow, el('span', 'tl', '\u25C6 By opponent ranking'));
    tierPoints(M).forEach(function (t) { var cp = el('span', 'tc'); add(cp, el('b', null, t.lab), el('span', null, t.pp.toFixed(1) + '% of points \u2192 ' + Math.round(t.wp) + '% of matches')); add(tierRow, cp); });
    add(sec, tierRow);
    if (CH && SH) add(sec, reveal(nadalPanel(M, CH, SH)));
    return sec;
  }

  // IV. Surfaces
  function renderSurfaces(M) {
    var S = {};
    M.forEach(function (m) {
      var s = S[m.surface] || (S[m.surface] = { s: m.surface, n: 0, w: 0, pw: 0, pt: 0, t: 0 });
      s.n++; s.w += m.win; s.t += m.title; if (m.hasPts) { s.pw += m.points_won; s.pt += m.points_total; }
    });
    var L = Object.keys(S).map(function (k) { return S[k]; }).filter(function (s) { return s.n >= 10; }).sort(function (a, b) { return b.w / b.n - a.w / a.n; });
    var grass = S.Grass, clay = S.Clay;
    var sec = el('section', 'ch');
    add(sec, chapterHead('IV', 'Surfaces', 'Grass was home. On clay, a slightly smaller edge in points meant many fewer wins.',
      'On grass he won ' + pct(grass.w, grass.n, 1) + ' of his matches with ' + pct(grass.pw, grass.pt, 1) + ' of points. On clay he won ' +
      pct(clay.w, clay.n, 1) + ' with ' + pct(clay.pw, clay.pt, 1) + '. The gap in points between the two is under ' +
      Math.ceil(100 * grass.pw / grass.pt - 100 * clay.pw / clay.pt) + ' per hundred.'));
    var box = el('div', 'chart');
    var hd = el('div', 'surf-head');
    ['Surface', 'Record', 'Matches won', 'Points won'].forEach(function (t) { add(hd, el('span', null, t)); });
    add(box, hd);
    var bars = [];
    L.forEach(function (s) {
      var row = el('div', 'surf');
      var bar = el('div', 'bar'); var b = el('b'); add(bar, b);
      if (s.s === 'Clay') b.style.background = C.gold;
      if (s.s === 'Carpet') b.style.background = C.muted;
      bars.push([b, 100 * s.w / s.n]);
      var pw = el('div', 'pw num', pct(s.pw, s.pt, 1)); add(pw, el('small', null, 'of points'));
      var nm = el('div', 'nm', s.s);
      add(nm, el('small', null, pct(s.w, s.n, 1) + ' of matches'));
      var wl = el('div', 'wl num', s.w + '–' + (s.n - s.w) + ' · ' + pct(s.w, s.n, 1) + ' · ' + s.t + ' titles');
      add(row, nm, wl, bar, pw);
      add(box, row);
    });
    add(sec, reveal(box, function () { bars.forEach(function (x) { setTimeout(function () { x[0].style.width = x[1].toFixed(1) + '%'; }, 120); }); }));
    return sec;
  }

  // ================================================================== animated redraws
  // morph(boxes, build): snapshot what is drawn in each box, run build() (which redraws them), then tween every
  // element that exists in both versions from its old state to its new one. Elements are matched by data-k when
  // present, otherwise by their position in the tree. New elements fade in.
  var MORPH_NUM = ['x', 'y', 'width', 'height', 'cx', 'cy', 'r', 'rx', 'x1', 'x2', 'y1', 'y2', 'opacity', 'fill-opacity', 'stroke-opacity', 'stroke-width'];
  var NUM_RE = /-?\d[\d,]*(?:\.\d+)?/g;
  // one set of motion settings for the whole app (CSS transitions mirror these: swap-in .5s, camera .6s)
  var MOTION = { morph: 750, shift: 1100, pop: 420, dwell: 380 };
  var MORPH_MS = MOTION.morph;
  function whenVisible(node, fn) {
    if (!('IntersectionObserver' in window)) { fn(); return; }
    var ob = new IntersectionObserver(function (es) { es.forEach(function (e) { if (e.isIntersecting) { ob.disconnect(); fn(); } }); }, { threshold: 0.35 });
    ob.observe(node);
  }
  function reducedMotion() { try { return !!(window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches); } catch (e) { return false; } }
  function easeIO(t) { return t < 0.5 ? 4 * t * t * t : 1 - Math.pow(-2 * t + 2, 3) / 2; }
  function parseCol(v) {
    if (!v) return null;
    var m = /^#([0-9a-f]{6})$/i.exec(v);
    if (m) return [parseInt(m[1].slice(0, 2), 16), parseInt(m[1].slice(2, 4), 16), parseInt(m[1].slice(4, 6), 16)];
    m = /^#([0-9a-f]{3})$/i.exec(v);
    if (m) return [0, 1, 2].map(function (i) { return parseInt(m[1][i] + m[1][i], 16); });
    m = /^rgba?\(([^)]+)\)$/i.exec(v);
    if (m) { var p = m[1].split(',').map(parseFloat); return p.length >= 3 ? p.slice(0, 3) : null; }
    return null;
  }
  function numSplit(s) {
    var nums = [], skel = String(s).replace(NUM_RE, function (n) { nums.push(n); return '\u0001'; });
    return { skel: skel, nums: nums };
  }
  function numJob(a, b) {
    if (a === b) return null;
    var A = numSplit(a), B = numSplit(b);
    if (A.skel !== B.skel || A.nums.length !== B.nums.length || !A.nums.length) return null;
    return { skel: B.skel, from: A.nums.map(function (n) { return parseFloat(n.replace(/,/g, '')); }),
      to: B.nums.map(function (n) { return parseFloat(n.replace(/,/g, '')); }),
      dec: B.nums.map(function (n) { var i = n.indexOf('.'); return i < 0 ? 0 : n.length - i - 1; }),
      com: B.nums.map(function (n) { return n.indexOf(',') >= 0; }) };
  }
  function numAt(j, k) {
    var i = 0;
    return j.skel.replace(/\u0001/g, function () {
      var v = j.from[i] + (j.to[i] - j.from[i]) * k, s = v.toFixed(j.dec[i]);
      if (j.com[i]) s = Number(s).toLocaleString('en-US', { minimumFractionDigits: j.dec[i], maximumFractionDigits: j.dec[i] });
      i++; return s;
    });
  }
  function mkKeys(root) {
    var out = [['@', root]];
    (function walk(node, base) {
      var cnt = {};
      for (var i = 0; i < node.children.length; i++) {
        var k = node.children[i], tg = k.tagName;
        cnt[tg] = (cnt[tg] || 0) + 1;
        var dk = k.getAttribute('data-k');
        var key = dk ? base + '#' + dk : base + '/' + tg + cnt[tg];
        out.push([key, k]); walk(k, key);
      }
    })(root, '');
    return out;
  }
  function snap(e) {
    var s = { a: {} };
    MORPH_NUM.forEach(function (n) { var v = e.getAttribute(n); if (v !== null && v !== '' && !isNaN(+v)) s.a[n] = +v; });
    ['d', 'points', 'transform'].forEach(function (n) { var v = e.getAttribute(n); if (v) s[n] = v; });
    ['fill', 'stroke'].forEach(function (n) { var v = e.getAttribute(n); if (parseCol(v)) s[n] = v; });
    if (!e.children.length) s.t = e.textContent;
    if (e.style && /px$/.test(e.style.width || '')) s.w = parseFloat(e.style.width);
    return s;
  }
  // skip(e): elements the caller animates itself (they are left out of the fade-in)
  function morph(boxes, build, dur, skip) {
    dur = dur || MORPH_MS;
    if (reducedMotion()) { build(); return; }
    if (boxes[0].__morph) cancelAnimationFrame(boxes[0].__morph);
    var old = {};
    boxes.forEach(function (b, bi) { mkKeys(b).forEach(function (p) { old[bi + p[0]] = snap(p[1]); }); });
    var hadOld = Object.keys(old).length > boxes.length;
    build();
    if (!hadOld) return;
    var jobs = [], fresh = [];
    boxes.forEach(function (b, bi) {
      mkKeys(b).forEach(function (p) {
        var e = p[1], o = old[bi + p[0]];
        if (e.parentNode && e.parentNode.__fresh && p[0] !== '@') { e.__fresh = 1; return; }
        if (!o) { e.__fresh = 1; if (!(skip && skip(e))) fresh.push(e); return; }
        var n = snap(e), j = { e: e, a: [], s: [], c: [], t: null, w: null };
        for (var k in n.a) if (k in o.a && o.a[k] !== n.a[k]) j.a.push([k, o.a[k], n.a[k]]);
        ['d', 'points', 'transform'].forEach(function (k) { if (o[k] && n[k]) { var q = numJob(o[k], n[k]); if (q) j.s.push([k, q]); } });
        ['fill', 'stroke'].forEach(function (k) { if (o[k] && n[k] && o[k] !== n[k]) j.c.push([k, parseCol(o[k]), parseCol(n[k]), n[k]]); });
        if (o.t !== undefined && n.t !== undefined) j.t = numJob(o.t, n.t);
        if (o.w !== undefined && n.w !== undefined && o.w !== n.w) j.w = [o.w, n.w];
        if (j.a.length || j.s.length || j.c.length || j.t || j.w) jobs.push(j);
      });
    });
    function apply(k) {
      jobs.forEach(function (j) {
        j.a.forEach(function (a) { j.e.setAttribute(a[0], (a[1] + (a[2] - a[1]) * k).toFixed(2)); });
        j.s.forEach(function (s) { j.e.setAttribute(s[0], numAt(s[1], k)); });
        j.c.forEach(function (c) {
          if (k >= 1) { j.e.setAttribute(c[0], c[3]); return; }
          j.e.setAttribute(c[0], 'rgb(' + [0, 1, 2].map(function (i) { return Math.round(c[1][i] + (c[2][i] - c[1][i]) * k); }).join(',') + ')');
        });
        if (j.t) j.e.textContent = numAt(j.t, k);
        if (j.w) j.e.style.width = (j.w[0] + (j.w[1] - j.w[0]) * k).toFixed(1) + 'px';
      });
      fresh.forEach(function (e) { if (e.style) e.style.opacity = k >= 1 ? '' : String(Math.max(0, (k - 0.25) / 0.75)); });
    }
    apply(0);
    var t0 = null;
    function step(ts) {
      if (t0 === null) t0 = ts;
      var g = Math.min(1, (ts - t0) / dur);
      apply(g >= 1 ? 1 : easeIO(g));
      boxes[0].__morph = g < 1 ? requestAnimationFrame(step) : 0;
    }
    boxes[0].__morph = requestAnimationFrame(step);
  }
  // Replay an element's entrance (fade and rise) — for content that is replaced rather than morphed.
  function swapIn(node, cls) {
    node.classList.remove(cls || 'swap-in'); void node.offsetWidth; node.classList.add(cls || 'swap-in');
  }


  // ================================================================== loading screen
  // A rally on a small court while the data loads. The scoreboard is the progress bar: each datasource that arrives
  // wins a point (15, 30, 40, game). Facts rotate underneath. It sits over the page and fades out once the story is drawn.
  var LOADER = null;
  function makeLoader() {
    var ov = el('div', 'ld'), card = el('div', 'ld-card');
    var court = '<svg class="ld-court" viewBox="0 0 240 120" aria-hidden="true">' +
      '<rect x="0" y="0" width="240" height="120" rx="8" fill="#2B5E3B"/>' +
      '<rect x="12" y="12" width="216" height="96" fill="#377550"/>' +
      '<g stroke="rgba(247,244,236,.9)" stroke-width="1.4" fill="none">' +
      '<rect x="12" y="12" width="216" height="96"/><line x1="12" y1="24" x2="228" y2="24"/><line x1="12" y1="96" x2="228" y2="96"/>' +
      '<line x1="68" y1="24" x2="68" y2="96"/><line x1="172" y1="24" x2="172" y2="96"/><line x1="68" y1="60" x2="172" y2="60"/></g>' +
      '<line x1="120" y1="6" x2="120" y2="114" stroke="#F7F4EC" stroke-width="2.6"/>' +
      '<g class="ld-p ld-fed"><circle r="7.5" fill="#F7F4EC" stroke="#2E6B3F" stroke-width="1.6"/><text y="2.6" text-anchor="middle" font-size="7" font-weight="700" fill="#2E6B3F" font-family="-apple-system, Helvetica, Arial, sans-serif">RF</text></g>' +
      '<g class="ld-p ld-opp"><circle r="7.5" fill="#4B2E6B" stroke="#F7F4EC" stroke-width="1.6"/></g>' +
      '<g class="ld-ballx"><circle class="ld-shadow" r="3" fill="#000" opacity=".28"/><g class="ld-bally"><circle r="3.6" fill="#E6EE7A" stroke="#9CA43A" stroke-width=".6"/></g></g>' +
      '</svg>';
    card.innerHTML = court;
    var sb = el('div', 'ld-score');
    add(sb, el('span', 'ld-lab', 'Loading the match'));
    var big = el('div', 'ld-big num', '0 – 0'); add(sb, big);
    add(card, sb);
    var steps = el('ul', 'ld-steps');
    var STEP = [['Career matches', 'matches'], ['Match charts', 'charted matches'], ['Shot-by-shot data', 'shots'], ['Serve data', 'serves']];
    var lis = STEP.map(function (s) { var li = el('li'); add(li, el('i'), el('span', null, s[0])); add(steps, li); return li; });
    add(card, steps);
    var FACTS = ['He won almost 80% of his matches, but only 54% of the points.',
      'Volunteers for the Match Charting Project logged every shot of 704 of his matches.',
      'In the 2019 Wimbledon final he won more points than Djokovic, and still lost.',
      'He finished his career with 103 titles and 20 major titles.',
      'Small edges, applied again and again: that is the whole story.'];
    var fact = el('p', 'ld-fact', FACTS[0]); add(card, fact);
    add(ov, card);
    document.body.appendChild(ov);
    var fi = 0, timer = setInterval(function () { fi = (fi + 1) % FACTS.length; fact.textContent = FACTS[fi]; swapIn(fact, 'swap-in'); }, 3200);
    var t0 = Date.now(), done = 0, SC = ['0 – 0', '15 – 0', '30 – 0', '40 – 0', 'Game'];
    function score() { big.textContent = SC[Math.min(4, done)]; swapIn(big, 'ld-pop'); if (done >= 4) big.classList.add('game'); }
    return {
      // mark one datasource as loaded; count is shown next to its label
      track: function (k, promise, countFn) {
        return promise.then(function (rows) {
          var n = countFn ? countFn(rows) : (rows && rows.length) || 0;
          lis[k].classList.add('ok');
          if (n) lis[k].lastChild.textContent = fmt(n) + ' ' + STEP[k][1];
          done++; score();
          return rows;
        });
      },
      finish: function () {
        clearInterval(timer);
        done = 4; score();
        var wait = Math.max(450, 1300 - (Date.now() - t0));
        setTimeout(function () { ov.classList.add('out'); setTimeout(function () { if (ov.parentNode) ov.parentNode.removeChild(ov); }, 700); }, wait);
      },
      fail: function () { clearInterval(timer); if (ov.parentNode) ov.parentNode.removeChild(ov); }
    };
  }


  // ================================================================== courts
  // Court in metres, centred on the net. x: −5.485..5.485 (doubles), y: −11.885 (opponent baseline) .. 11.885 (Federer).
  // Lanes: L / M / R are screen thirds of the singles court. Federer is at the bottom, so his forehand side is screen-right.
  var LANE_X = { L: -2.75, M: 0, R: 2.75 };
  var COURT = { surround: '#2B5E3B', inner: '#377550', line: 'rgba(247,244,236,.92)' };
  var gradSeq = 0;

  function drawCourt(s, half) {
    var mx = 1.6, myTop = half ? 1.8 : 2.8, myBot = half ? 0.9 : 2.4;
    var yTop = -11.885 - myTop, yBot = half ? 0.9 : 11.885 + myBot;
    var W = (10.97 + 2 * mx) * s, H = (yBot - yTop) * s;
    var X = function (x) { return (x + 5.485 + mx) * s; };
    var Y = function (y) { return (y - yTop) * s; };
    var svg = svgRoot(W, H);
    add(svg, sv('rect', { x: 0, y: 0, width: W, height: H, rx: 6, fill: COURT.surround }));
    var y0 = -11.885, y1 = half ? 0 : 11.885;
    add(svg, sv('rect', { x: X(-5.485), y: Y(y0), width: 10.97 * s, height: (y1 - y0) * s, fill: COURT.inner }));
    var lw = Math.max(1, 0.06 * s);
    function L(x1, yA, x2, yB, w) { add(svg, sv('line', { x1: X(x1), y1: Y(yA), x2: X(x2), y2: Y(yB), stroke: COURT.line, 'stroke-width': w || lw })); }
    L(-5.485, y0, 5.485, y0); if (!half) L(-5.485, y1, 5.485, y1);
    L(-5.485, y0, -5.485, y1); L(5.485, y0, 5.485, y1);
    L(-4.115, y0, -4.115, y1); L(4.115, y0, 4.115, y1);
    L(-4.115, -6.4, 4.115, -6.4); if (!half) L(-4.115, 6.4, 4.115, 6.4);
    L(0, -6.4, 0, half ? 0 : 6.4);
    L(0, y0, 0, y0 + 0.25); if (!half) L(0, y1, 0, y1 - 0.25);
    add(svg, sv('line', { x1: X(-6.0), y1: Y(0), x2: X(6.0), y2: Y(0), stroke: '#F7F4EC', 'stroke-width': lw * 2.2, opacity: 0.95 }));
    return { svg: svg, X: X, Y: Y, s: s, W: W, H: H };
  }

  // Name a shot pattern from its origin lane (Federer's side) and target lane (opponent's side).
  function patternName(shot, o, t) {
    var fh = shot === 'Forehand';
    var bh = shot === 'Backhand' || shot === 'Slice';
    if (o === '?' || t === '?') return null;
    if (t === 'M') return 'Down the middle';
    if (o === 'M') return t === 'L' ? 'From the centre, to the forehand corner' : 'From the centre, to the backhand corner';
    if (o === t) {
      if (fh && o === 'L') return 'Inside-in';
      if (bh && o === 'R') return 'Inside-in';
      return 'Down the line';
    }
    if (fh && o === 'L') return 'Inside-out';
    if (bh && o === 'R') return 'Inside-out';
    return 'Crosscourt';
  }
  function originLabel(o, phase, shot) {
    if (phase === 'Return') return { L: 'from the ad court', M: 'off a serve down the T', R: 'from the deuce court' }[o] || '';
    return { L: 'from his backhand corner', M: 'from the centre', R: 'from his forehand corner' }[o] || '';
  }
  function targetLabel(t) { return { L: 'to the forehand corner', M: 'down the middle', R: 'to the backhand corner' }[t] || ''; }

  function aggPairs(rows) {
    var P = {}, tot = 0;
    rows.forEach(function (r) {
      if (r.origin === '?' || r.target === '?') return;
      var k = r.origin + r.target;
      var o = P[k] || (P[k] = { o: r.origin, t: r.target, n: 0, win: 0, ue: 0, fe: 0, won: 0 });
      o.n += r.n; o.won += r.won; tot += r.n;
      if (r.outcome === 'Winner') o.win += r.n; else if (r.outcome === 'Unforced error') o.ue += r.n; else if (r.outcome === 'Forced error') o.fe += r.n;
    });
    var list = Object.keys(P).map(function (k) { var o = P[k]; o.share = tot ? o.n / tot : 0; return o; });
    return { list: list, total: tot };
  }

  // Flow map: one ribbon per origin→target pair, width by share. Optional highlight pattern name.
  function flowCourt(rows, shot, phase, opts) {
    opts = opts || {};
    var c = drawCourt(opts.scale || 18, false);
    var A = aggPairs(rows);
    var max = Math.max.apply(null, A.list.map(function (o) { return o.share; }).concat([0.0001]));
    var gid = 'fg' + (++gradSeq);
    var defs = sv('defs', {});
    add(c.svg, defs);
    var ribbons = [];
    A.list.sort(function (a, b) { return a.share - b.share; }).forEach(function (o) {
      var name = patternName(shot, o.o, o.t);
      var x0 = c.X(LANE_X[o.o]), y0 = c.Y(11.2), x1 = c.X(LANE_X[o.t]), y1 = c.Y(-10.4);
      var bow = o.o === o.t ? (o.o === 'L' ? -1 : 1) * 0.5 * c.s : 0;
      var d = 'M' + x0 + ',' + y0 + ' Q' + ((x0 + x1) / 2 + bow) + ',' + c.Y(0) + ' ' + x1 + ',' + y1;
      var hl = opts.highlight && name === opts.highlight;
      var w = 1.2 + (opts.maxWidth || 13) * o.share / max;
      var p = sv('path', { d: d, fill: 'none', stroke: hl ? C.gold : '#F7F4EC', 'stroke-width': w.toFixed(1), 'stroke-linecap': 'round',
        opacity: hl ? 0.95 : (0.22 + 0.6 * o.share / max).toFixed(2) });
      p.__base = p.getAttribute('opacity');
      if (!opts.noHover) {
        p.addEventListener('mousemove', function (e) {
          ribbons.forEach(function (r) { r.setAttribute('opacity', r === p ? 1 : 0.12); });
          showTip(e, (name || 'Pattern') + (phase === 'Return' ? ' return' : ''), [
            originLabel(o.o, phase, shot) + ', ' + targetLabel(o.t),
            fmt(o.n) + ' shots · ' + (100 * o.share).toFixed(1) + '% of ' + (phase === 'Return' ? 'returns' : shot.toLowerCase() + 's'),
            'Winners ' + pct(o.win, o.n) + ' · Unforced errors ' + pct(o.ue, o.n) + ' · Points won ' + pct(o.won, o.n)]);
        });
        p.addEventListener('mouseleave', function () { hideTip(); ribbons.forEach(function (r) { r.setAttribute('opacity', r.__base); }); });
      }
      ribbons.push(p);
      add(c.svg, p);
    });
    // target shares
    if (!opts.noLabels) {
      var tgt = { L: 0, M: 0, R: 0 }, org = { L: 0, M: 0, R: 0 };
      A.list.forEach(function (o) { tgt[o.t] += o.n; org[o.o] += o.n; });
      ['L', 'M', 'R'].forEach(function (k) {
        if (!A.total) return;
        add(c.svg, sv('circle', { cx: c.X(LANE_X[k]), cy: c.Y(-10.4), r: 3 + 9 * Math.sqrt(tgt[k] / A.total), fill: C.gold, stroke: '#F7F4EC', 'stroke-width': 1 }));
        add(c.svg, sv('text', { x: c.X(LANE_X[k]), y: c.Y(-12.3), 'text-anchor': 'middle', fill: '#F7F4EC', 'font-size': 11, 'font-family': 'var(--sans)' },
          Math.round(100 * tgt[k] / A.total) + '%'));
        add(c.svg, sv('circle', { cx: c.X(LANE_X[k]), cy: c.Y(11.2), r: 2.5 + 6 * Math.sqrt(org[k] / A.total), fill: '#F7F4EC', opacity: 0.85 }));
      });
      add(c.svg, sv('text', { x: c.W / 2, y: c.H - 6, 'text-anchor': 'middle', fill: '#F7F4EC', opacity: 0.75, 'font-size': 10, 'letter-spacing': '2', 'font-family': 'var(--sans)' }, 'FEDERER'));
      add(c.svg, sv('text', { x: c.W / 2, y: 14, 'text-anchor': 'middle', fill: '#F7F4EC', opacity: 0.75, 'font-size': 10, 'letter-spacing': '2', 'font-family': 'var(--sans)' }, 'OPPONENT'));
    }
    return { svg: c.svg, agg: A };
  }

  // Serve map: opponent half, deuce box (screen-left) and ad box (screen-right), thirds = wide / body / T.
  function serveCourt(rows, serveNo, opts) {
    opts = opts || {};
    var c = drawCourt(opts.scale || 15, true);
    var ZX = { Deuce: { Wide: [-4.115, -2.743], Body: [-2.743, -1.372], T: [-1.372, 0] }, Ad: { T: [0, 1.372], Body: [1.372, 2.743], Wide: [2.743, 4.115] } };
    var Z = {}, inTot = { Deuce: 0, Ad: 0 }, all = { Deuce: 0, Ad: 0 }, faults = { Deuce: 0, Ad: 0 };
    rows.forEach(function (r) {
      if (r.serve_no !== serveNo) return;
      all[r.side] = (all[r.side] || 0) + r.n;
      if (r.result === 'Fault' || r.result === 'Double fault') { faults[r.side] += r.n; return; }
      if (!ZX[r.side] || !ZX[r.side][r.direction]) return;
      var k = r.side + r.direction, z = Z[k] || (Z[k] = { n: 0, ace: 0, unret: 0, won: 0, snv: 0 });
      z.n += r.n; z.won += r.won; z.snv += r.serve_volley; inTot[r.side] += r.n;
      if (r.result === 'Ace') z.ace += r.n; if (r.result === 'Unreturned') z.unret += r.n;
    });
    var best = null;
    ['Deuce', 'Ad'].forEach(function (side) {
      ['Wide', 'Body', 'T'].forEach(function (dir) {
        var z = Z[side + dir] || { n: 0, ace: 0, unret: 0, won: 0 }, share = inTot[side] ? z.n / inTot[side] : 0;
        var xr = ZX[side][dir];
        var rect = sv('rect', { x: c.X(xr[0]) + 1, y: c.Y(-6.4) + 1, width: (xr[1] - xr[0]) * c.s - 2, height: 6.4 * c.s - 2,
          fill: C.gold, opacity: (0.08 + 0.85 * share).toFixed(2) });
        rect.addEventListener('mousemove', function (e) {
          showTip(e, side + ' court · ' + (dir === 'T' ? 'down the T' : dir === 'Wide' ? 'out wide' : 'into the body'), [
            fmt(z.n) + ' serves in (' + (100 * share).toFixed(1) + '% of ' + (serveNo === 1 ? 'first' : 'second') + ' serves in on this side)',
            'Aces ' + pct(z.ace, z.n) + ' · Unreturned ' + pct(z.ace + z.unret, z.n), 'Points won ' + pct(z.won, z.n)]);
        });
        rect.addEventListener('mouseleave', hideTip);
        add(c.svg, rect);
        var cx = c.X((xr[0] + xr[1]) / 2);
        add(c.svg, sv('text', { x: cx, y: c.Y(-3.6), 'text-anchor': 'middle', fill: share > 0.45 ? C.ink : '#F7F4EC', 'font-size': Math.round(c.s * 0.5), 'font-family': 'var(--serif)', 'pointer-events': 'none' }, Math.round(100 * share) + '%'));
        add(c.svg, sv('text', { x: cx, y: c.Y(-2.3), 'text-anchor': 'middle', fill: share > 0.45 ? C.ink : '#F7F4EC', opacity: 0.85, 'font-size': Math.round(c.s * 0.34), 'font-family': 'var(--sans)', 'pointer-events': 'none' }, dir === 'T' ? 'T' : dir.toLowerCase()));
        if (!best || share > best.share) best = { side: side, dir: dir, share: share };
      });
      add(c.svg, sv('text', { x: c.X(side === 'Deuce' ? -2.05 : 2.05), y: c.Y(-7.0), 'text-anchor': 'middle', fill: '#F7F4EC', opacity: 0.75, 'font-size': Math.max(9, Math.round(c.s * 0.34)), 'letter-spacing': '1', 'font-family': 'var(--sans)' }, side.toUpperCase()));
    });
    return { svg: c.svg, zones: Z, inTot: inTot, all: all, faults: faults };
  }

  function miniPattern(o, t, color) {
    var c = drawCourt(4.2, false);
    var x0 = c.X(LANE_X[o]), y0 = c.Y(11.2), x1 = c.X(LANE_X[t]), y1 = c.Y(-10.4);
    add(c.svg, sv('path', { d: 'M' + x0 + ',' + y0 + ' Q' + ((x0 + x1) / 2) + ',' + c.Y(0) + ' ' + x1 + ',' + y1, fill: 'none', stroke: color || C.gold, 'stroke-width': 3, 'stroke-linecap': 'round' }));
    add(c.svg, sv('circle', { cx: x1, cy: y1, r: 3.2, fill: color || C.gold }));
    return c.svg;
  }

  // ================================================================== V. The shots
  var ERAS = { 'All': [0, 9999], '1998–2003': [1998, 2003], '2004–07': [2004, 2007], '2008–13': [2008, 2013], '2014–21': [2014, 2021] };
  var OPPS = ['All', 'Rafael Nadal', 'Novak Djokovic', 'Andy Murray', 'Everyone else'];

  function renderShots(M, CHall, SH, SV) {
    var CH = CHall.filter(function (c) { return c.source !== 'Slam'; });
    var byKey = {}; M.forEach(function (m) { byKey[m.match_key] = m; });
    var meta = {}; CH.forEach(function (c) { meta[c.chart_id] = { c: c, m: byKey[c.match_key] }; });
    var F = { surface: 'All', opp: 'All', era: 'All', result: 'All' };
    var shotSel = 'Forehand', serveSel = 1;

    var sec = el('section', 'ch');
    var totalPts = sum(CH, function (c) { return c.n_points; });
    add(sec, chapterHead('V', 'Shot by shot', 'The 54% was built one shot at a time: the forehand, the backhand, the serve.',
      'Volunteers for the Match Charting Project have logged every shot of ' + fmt(CH.length) + ' Federer matches, ' + fmt(totalPts) +
      ' points in all. Each court shows where his shots came from (the bottom) and where they went (the top). Use the filters to compare surfaces, rivals and eras.'));

    var filters = el('div', 'filters');
    function group(label, key, values, names) {
      var g = el('div', 'fgroup'); add(g, el('span', 'fl', label));
      values.forEach(function (v, i) {
        var b = el('button', 'pill' + (F[key] === v ? ' on' : ''), names ? names[i] : v);
        b.addEventListener('click', function () { setFilter(key, v); });
        add(g, b); (PILLS[key] || (PILLS[key] = {}))[v] = b;
      });
      return g;
    }
    var PILLS = {};
    function setFilter(key, v) {
      F[key] = v;
      Object.keys(PILLS[key]).forEach(function (k) { PILLS[key][k].classList.toggle('on', k === String(v)); });
      update();
    }
    // matches (charted or not) that fit a set of filters
    function fits(m, f, charted) {
      if (f.surface !== 'All' && m.surface !== f.surface) return false;
      if (f.result !== 'All' && m.result !== f.result) return false;
      var er = ERAS[f.era]; if (m.season < er[0] || m.season > er[1]) return false;
      if (f.opp !== 'All') { if (f.opp === 'Everyone else') { if (OPPS.indexOf(m.opponent) > 0) return false; } else if (m.opponent !== f.opp) return false; }
      return true;
    }
    var chartedKeys = {}; CH.forEach(function (c) { chartedKeys[c.match_key] = 1; });
    function countCharted(f) { return M.filter(function (m) { return chartedKeys[m.match_key] && fits(m, f); }).length; }
    // explain an empty selection in tennis terms, and suggest the nearest filters that do have data
    function emptyReason() {
      var all = M.filter(function (m) { return fits(m, F); });
      var who = F.opp !== 'All' && F.opp !== 'Everyone else' ? F.opp.split(' ').slice(-1)[0] : null;
      var surf = F.surface !== 'All' ? F.surface.toLowerCase() : null, era = F.era !== 'All' ? F.era : null;
      var res = F.result === 'W' ? 'wins' : F.result === 'L' ? 'losses' : null;
      if (all.length) return 'Federer played ' + fmt(all.length) + ' match' + (all.length === 1 ? '' : 'es') + ' that fit these filters, but none of them were charted shot by shot by the Match Charting Project.';
      if (who) {
        var vs = M.filter(function (m) { return m.opponent === F.opp; }).sort(function (a, b) { return String(a.match_date).localeCompare(String(b.match_date)); });
        var first = vs[0], last = vs[vs.length - 1], er = ERAS[F.era];
        if (era && first && er[1] < first.season) return 'Federer and ' + who + ' hadn’t played each other yet. Their first match was in ' + first.season + ', at ' + first.tournament + '.';
        if (era && last && er[0] > last.season) return 'Federer and ' + who + ' had stopped meeting by then. Their last match was in ' + last.season + ', at ' + last.tournament + '.';
        var inEra = vs.filter(function (m) { return fits(m, { surface: 'All', result: 'All', era: F.era, opp: F.opp }); });
        if (surf && !inEra.filter(function (m) { return m.surface === F.surface; }).length) return 'Federer and ' + who + ' played ' + inEra.length + ' time' + (inEra.length === 1 ? '' : 's') + (era ? ' in ' + era : '') + ', but never on ' + surf + '.';
        if (res) return 'Federer had no ' + res + ' against ' + who + (surf ? ' on ' + surf : '') + (era ? ' in ' + era : '') + '.';
      }
      if (surf && era) return 'Federer played no matches on ' + surf + (who ? ' against ' + who : '') + ' in ' + era + (res ? ' that ended in ' + res : '') + '.';
      return 'Federer has no matches that fit all of these filters.';
    }
    function suggestions() {
      var out = [];
      ['era', 'surface', 'opp', 'result'].forEach(function (k) {
        if (F[k] === 'All') return;
        var best = null, keys = Object.keys(PILLS[k]), cur = keys.indexOf(String(F[k]));
        // for eras, the nearest era with data (the next one first); otherwise the value with the most matches
        var order = k === 'era' ? keys.slice(cur + 1).concat(keys.slice(1, cur).reverse()) : keys;
        order.forEach(function (v) {
          if (v === F[k] || v === 'All') return;
          var f = {}; for (var q in F) f[q] = F[q]; f[k] = v;
          var n = countCharted(f); if (!n) return;
          if (!best || (k !== 'era' && n > best.n)) best = { k: k, v: v, n: n };
        });
        if (best) out.push(best);
      });
      return out.slice(0, 3);
    }
    add(filters, group('Surface', 'surface', ['All', 'Grass', 'Hard', 'Clay']),
      group('Opponent', 'opp', OPPS, ['All', 'Nadal', 'Djokovic', 'Murray', 'Everyone else']),
      group('Era', 'era', Object.keys(ERAS)),
      group('Result', 'result', ['All', 'W', 'L'], ['All', 'Wins', 'Losses']));
    add(sec, filters);
    var summary = el('div', 'fsummary'); add(sec, summary);
    var empty = el('div', 'fempty'); empty.style.display = 'none'; add(sec, empty);

    // Panel A: shot map
    var pA = el('div', 'panel');
    add(pA, el('h3', null, 'Where the shots went'));
    var subA = el('p', 'sub'); add(pA, subA);
    var tabs = el('div', 'tabs');
    ['Forehand', 'Backhand', 'Slice', 'Return', 'Volley'].forEach(function (s) {
      var b = el('button', 'pill' + (s === shotSel ? ' on' : ''), s);
      b.addEventListener('click', function () { shotSel = s; Array.prototype.forEach.call(tabs.children, function (x) { x.classList.remove('on'); }); b.classList.add('on'); drawA(); });
      add(tabs, b);
    });
    add(pA, tabs);
    var viewSel = 'Landing';
    var vtabs = el('div', 'tabs');
    ['Landing', 'Patterns'].forEach(function (v) {
      var b = el('button', 'pill' + (v === viewSel ? ' on' : ''), v === 'Landing' ? 'Where they landed' : 'Patterns (from \u2192 to)');
      b.addEventListener('click', function () { viewSel = v; Array.prototype.forEach.call(vtabs.children, function (x) { x.classList.remove('on'); }); b.classList.add('on'); drawA(); });
      add(vtabs, b);
    });
    add(pA, vtabs);
    var gridA = el('div', 'courtgrid'); var courtA = el('div', 'court'); var tableA = el('div');
    add(gridA, courtA, tableA); add(pA, gridA);
    add(sec, reveal(pA));

    // Panel B: serve map + rally length
    var pB = el('div', 'panel');
    add(pB, el('h3', null, 'The serve, and how long the points lasted'));
    var subB = el('p', 'sub'); add(pB, subB);
    var tabsB = el('div', 'tabs');
    [1, 2].forEach(function (s) {
      var b = el('button', 'pill' + (s === serveSel ? ' on' : ''), s === 1 ? 'First serve' : 'Second serve');
      b.addEventListener('click', function () { serveSel = s; Array.prototype.forEach.call(tabsB.children, function (x) { x.classList.remove('on'); }); b.classList.add('on'); drawB(); });
      add(tabsB, b);
    });
    add(pB, tabsB);
    var twoB = el('div', 'two'); var courtB = el('div', 'court'); var rallyB = el('div');
    add(twoB, courtB, rallyB); add(pB, twoB);
    add(sec, reveal(pB));

    // Panel C: signature cards
    var pC = el('div', 'panel');
    add(pC, el('h3', null, 'The signatures'));
    add(pC, el('p', 'sub', 'The shots people remember, and what they were worth. These cards follow the filters above too.'));
    var cards = el('div', 'cards'); add(pC, cards);
    add(sec, reveal(pC));
    add(sec, el('p', 'note', 'Shot data is parsed from the Match Charting Project’s raw shot codes, so counts can differ slightly from Tennis Abstract’s match reports. Direction is recorded in thirds of the court. A shot’s starting position is taken from where the previous ball landed, and return depth is recorded only for returns. Hover over any ribbon or zone for details.'));
    add(sec, el('p', 'note', chartBias(M, CHall)));

    var ids = {}, sel = { sh: [], sv: [], ch: [] };
    function update() { morph([summary, subA, courtA, tableA, subB, courtB, rallyB, cards], update0); }
    function drawA() { morph([subA, courtA, tableA], drawA0); }
    function drawB() { morph([subB, courtB, rallyB], drawB0); }
    function update0() {
      ids = {}; var nM = 0, nP = 0;
      CH.forEach(function (c) {
        var m = meta[c.chart_id].m; if (!m) return;
        if (F.surface !== 'All' && m.surface !== F.surface) return;
        if (F.result !== 'All' && m.result !== F.result) return;
        var er = ERAS[F.era]; if (m.season < er[0] || m.season > er[1]) return;
        if (F.opp !== 'All') {
          if (F.opp === 'Everyone else') { if (OPPS.indexOf(m.opponent) > 0) return; }
          else if (m.opponent !== F.opp) return;
        }
        ids[c.chart_id] = 1; nM++; nP += c.n_points;
      });
      sel.sh = SH.filter(function (r) { return ids[r.chart_id]; });
      sel.sv = SV.filter(function (r) { return ids[r.chart_id]; });
      sel.ch = CH.filter(function (c) { return ids[c.chart_id]; });
      var won = sum(sel.ch, function (c) { return c.fed_points; });
      summary.textContent = nM ? ('Showing ' + fmt(nM) + ' charted matches · ' + fmt(nP) + ' points · Federer won ' + pct(won, nP) + ' of them.') : '';
      [pA, pB, pC].forEach(function (p) { p.style.display = nM ? '' : 'none'; });
      empty.style.display = nM ? 'none' : '';
      if (!nM) {
        empty.textContent = '';
        add(empty, el('div', 'fe-k', 'Nothing to show'), el('p', 'fe-t', emptyReason()));
        var sg = suggestions();
        if (sg.length) {
          var row = el('div', 'fe-s'); add(row, el('span', 'fl', 'Try'));
          var NAME = { era: 'Era', surface: 'Surface', opp: 'Opponent', result: 'Result' };
          sg.forEach(function (o) {
            var lab = o.v === 'All' ? 'any ' + NAME[o.k].toLowerCase() : (o.k === 'opp' ? (o.v === 'Everyone else' ? 'Everyone else' : o.v.split(' ').slice(-1)[0]) : o.k === 'result' ? (o.v === 'W' ? 'wins' : 'losses') : o.v);
            var b = el('button', 'pill', lab + ' · ' + o.n + ' match' + (o.n === 1 ? '' : 'es'));
            b.addEventListener('click', function () { setFilter(o.k, o.v); });
            add(row, b);
          });
          add(empty, row);
        }
        swapIn(empty, 'swap-in');
        return;
      }
      drawA0(); drawB0(); drawC0();
    }

    function drawA0() {
      courtA.textContent = ''; tableA.textContent = '';
      var rows = sel.sh.filter(function (r) {
        if (shotSel === 'Return') return r.phase === 'Return';
        if (shotSel === 'Volley') return r.phase === 'Rally' && (r.shot === 'Volley' || r.shot === 'Overhead');
        return r.phase === 'Rally' && r.shot === shotSel;
      });
      // Data-driven highlight: the pattern with the highest winner rate among those he used at least 5% of the time.
      var pk = shotSel === 'Return' ? 'Return' : shotSel;
      var pre = aggPairs(rows), byName = {};
      pre.list.forEach(function (o) {
        var nm = patternName(pk, o.o, o.t); if (!nm) return;
        var g = byName[nm] || (byName[nm] = { name: nm, n: 0, win: 0 }); g.n += o.n; g.win += o.win;
      });
      var best = null;
      Object.keys(byName).forEach(function (k) {
        var g = byName[k];
        if (!pre.total || g.n / pre.total < 0.05 || g.n < 30) return;
        if (!best || g.win / g.n > best.win / best.n) best = g;
      });
      var hl = best ? best.name : null;
      var fc = flowCourt(rows, shotSel === 'Return' ? 'Return' : shotSel, shotSel === 'Return' ? 'Return' : 'Rally', { highlight: hl });
      if (viewSel === 'Landing') {
        var LET = { Forehand: 'f', Backhand: 'b', Slice: 's', Volley: 'v', Overhead: 'o', 'Drop shot': 'u', Other: 'q' };
        var DIRK = { L: '1', M: '2', R: '3' }, DEPK = { Short: '7', Mid: '8', Deep: '9' };
        var items = rows.map(function (r) {
          return { w: r.n, shot: { letter: LET[r.shot] || 'q', dir: DIRK[r.target] || '0', depth: DEPK[r.depth] || '', isReturn: r.phase === 'Return',
            outcome: r.outcome, miss: r.miss === 'None' ? '' : r.miss } };
        });
        add(courtA, dotCourt(items, { scale: 18, cap: MOB ? 900 : 1800, half: MOB, seed: 3, title: 'Federer \u2191', noLabels: true }));
        add(courtA, dotLegend());
      } else add(courtA, fc.svg);
      var A = fc.agg;
      var win = sum(rows, function (r) { return r.outcome === 'Winner' ? r.n : 0; }), ue = sum(rows, function (r) { return r.outcome === 'Unforced error' ? r.n : 0; });
      var all = sum(rows, function (r) { return r.n; });
      var noun = shotSel === 'Return' ? 'returns' : shotSel.toLowerCase() + 's';
      subA.textContent = all ? (fmt(all) + ' ' + (shotSel === 'Return' ? 'returns' : shotSel.toLowerCase() + 's') + ': ' + pct(win, all) + ' were winners and ' + pct(ue, all) +
        ' were unforced errors. ' + (viewSel === 'Landing' ? 'Each dot is a shot, placed at random within the zone it was recorded in. Misses appear in the net, past the sideline or past the baseline.' : 'Ribbon width shows how often he chose each direction.' + (best ? ' Gold marks ' + (best.name === 'Down the middle' ? 'shots down the middle' : 'the ' + best.name.toLowerCase() + ' ' + (shotSel === 'Return' ? 'return' : shotSel.toLowerCase())) + ': ' +
          pct(best.n, pre.total, 0) + ' of his ' + noun + ', but ' + pct(best.win, best.n) + ' of them were winners, the highest rate of any pattern he used at least 5% of the time (' +
          pct(win, all) + ' across all ' + noun + ').' : ''))) : '';
      // pattern table, grouped by name
      var G = {};
      A.list.forEach(function (o) {
        var nm = patternName(shotSel === 'Return' ? 'Return' : shotSel, o.o, o.t);
        var k = shotSel === 'Return' ? nm + '|' + o.o : nm;
        var g = G[k] || (G[k] = { name: nm, o: o.o, n: 0, win: 0, ue: 0, won: 0 });
        g.n += o.n; g.win += o.win; g.ue += o.ue; g.won += o.won;
      });
      var L = Object.keys(G).map(function (k) { return G[k]; }).sort(function (a, b) { return b.n - a.n; });
      var t = el('table', 'pat'); var hr = el('tr');
      ['Pattern', 'Share', 'Winners', 'Errors', 'Pts won'].forEach(function (h) { add(hr, el('th', null, h)); });
      add(t, hr);
      var maxN = L.length ? L[0].n : 1;
      L.forEach(function (g) {
        var tr = el('tr', g.name === hl ? 'hl' : null); tr.setAttribute('data-k', 'p:' + g.name + '|' + g.o);
        var td = el('td'); add(td, el('span', 'pn', g.name + (shotSel === 'Return' ? ' · ' + (g.o === 'L' ? 'ad' : g.o === 'R' ? 'deuce' : 'T serve') : '')));
        add(td, el('span', 'pd', fmt(g.n) + ' shots' + (g.name === hl ? ' \u00b7 highest winner rate' : '')));
        var tdS = el('td'); var bar = el('span', 'mbar'); bar.style.width = Math.max(2, 60 * g.n / maxN) + 'px'; add(tdS, bar); tdS.appendChild(document.createTextNode(pct(g.n, A.total, 0)));
        add(tr, td, tdS, el('td', null, pct(g.win, g.n)), el('td', null, pct(g.ue, g.n)), el('td', null, pct(g.won, g.n)));
        add(t, tr);
      });
      add(tableA, t);
      if (shotSel === 'Return') {
        var D = { Short: 0, Mid: 0, Deep: 0 }, dt = 0;
        rows.forEach(function (r) { if (D[r.depth] !== undefined) { D[r.depth] += r.n; dt += r.n; } });
        if (dt) add(tableA, el('p', 'note', 'Return depth: ' + pct(D.Deep, dt, 0) + ' deep (near the baseline), ' + pct(D.Mid, dt, 0) + ' mid-court, ' + pct(D.Short, dt, 0) + ' short (inside the service line).'));
      }
    }

    function drawB0() {
      courtB.textContent = ''; rallyB.textContent = '';
      var sc = serveCourt(sel.sv, serveSel);
      add(courtB, el('div', 'mini-title', (serveSel === 1 ? 'First' : 'Second') + ' serves in play, by target'));
      add(courtB, sc.svg);
      var zw = function (side, dir) { var z = sc.zones[side + dir]; return z && sc.inTot[side] ? z.n / sc.inTot[side] : 0; };
      var aces = 0, inn = 0, won = 0;
      Object.keys(sc.zones).forEach(function (k) { aces += sc.zones[k].ace; inn += sc.zones[k].n; won += sc.zones[k].won; });
      var attempts = sc.all.Deuce + sc.all.Ad, faults = sc.faults.Deuce + sc.faults.Ad;
      subB.textContent = attempts ? ((serveSel === 1 ? 'First' : 'Second') + ' serves: ' + pct(attempts - faults, attempts) + ' went in. When they did, ' + pct(aces, inn) +
        ' were aces and he won ' + pct(won, inn) + ' of the points. On the ad side he served wide ' + pct(zw('Ad', 'Wide') * 100, 100, 0) + ' of the time, and on the deuce side he went down the T ' +
        pct(zw('Deuce', 'T') * 100, 100, 0) + ' of the time.') : '';
      // rally length dot plot
      var B = [['1–3 shots', 'r1_3'], ['4–6', 'r4_6'], ['7–9', 'r7_9'], ['10+', 'r10']];
      var totN = sum(sel.ch, function (c) { return c.n_points; });
      add(rallyB, el('div', 'mini-title', 'Points won by rally length'));
      var W = 460, H = 250, ml = 92, mr = 20, mt = 16, rowH = 50;
      var x = scale(40, 65, ml, W - mr);
      var svg = svgRoot(W, H);
      add(svg, sv('line', { x1: x(50), x2: x(50), y1: mt - 6, y2: mt + rowH * 4, stroke: C.gold, 'stroke-dasharray': '3 3' }));
      [40, 45, 50, 55, 60, 65].forEach(function (t) { add(svg, sv('text', { x: x(t), y: mt + rowH * 4 + 16, 'text-anchor': 'middle', class: 'lbl' }, t + '%')); });
      var best = null;
      B.forEach(function (b, i) {
        var n = sum(sel.ch, function (c) { return c[b[1] + '_n']; }), w = sum(sel.ch, function (c) { return c[b[1] + '_w']; });
        var p = n ? 100 * w / n : null, cy = mt + rowH * i + rowH / 2;
        add(svg, sv('text', { x: 0, y: cy - 2, class: 'lbl-serif' }, b[0]));
        add(svg, sv('text', { x: 0, y: cy + 14, class: 'lbl' }, n ? pct(n, totN, 0) + ' of points' : ''));
        add(svg, sv('line', { x1: ml, x2: W - mr, y1: cy, y2: cy, stroke: C.rule }));
        if (p !== null) {
          var pc = Math.max(40, Math.min(65, p));
          add(svg, sv('line', { x1: x(50), x2: x(pc), y1: cy, y2: cy, stroke: p >= 50 ? C.grass : C.purple, 'stroke-width': 3 }));
          add(svg, sv('circle', { cx: x(pc), cy: cy, r: 7, fill: p >= 50 ? C.grass : C.purple }));
          add(svg, sv('text', { x: x(pc) + (p >= 50 ? 12 : -12), y: cy + 4, 'text-anchor': p >= 50 ? 'start' : 'end', class: 'lbl-ink' }, p.toFixed(1) + '%'));
          if (!best || p > best.p) best = { p: p, l: b[0] };
        }
      });
      add(svg, sv('text', { x: x(50), y: mt - 8, 'text-anchor': 'middle', class: 'callout-sm', fill: C.gold }, 'even'));
      add(rallyB, svg);
      if (best) add(rallyB, el('p', 'note', 'Most points were over within three shots. His edge was largest in ' + best.l.replace(' shots', '') + '-shot rallies (' + best.p.toFixed(1) + '%). Rally length counts every shot that landed in, including the serve.'));
    }

    function drawC0() {
      cards.textContent = '';
      var rows = sel.sh;
      function patt(shot, o, t) { var n = 0, w = 0, won = 0; rows.forEach(function (r) { if (r.phase === 'Rally' && r.shot === shot && r.origin === o && r.target === t) { n += r.n; won += r.won; if (r.outcome === 'Winner') w += r.n; } }); return { n: n, w: w, won: won }; }
      var io = patt('Forehand', 'L', 'R'), dtl = patt('Backhand', 'L', 'L');
      var fhW = sum(rows, function (r) { return r.phase === 'Rally' && r.shot === 'Forehand' && r.outcome === 'Winner' ? r.n : 0; });
      var bhW = sum(rows, function (r) { return r.phase === 'Rally' && r.shot === 'Backhand' && r.outcome === 'Winner' ? r.n : 0; });
      var ch = sel.ch;
      var snvN = sum(ch, function (c) { return c.snv_n; }), snvW = sum(ch, function (c) { return c.snv_w; }), svN = sum(ch, function (c) { return c.sv_n; });
      var netN = sum(ch, function (c) { return c.net_n; }), netW = sum(ch, function (c) { return c.net_w; });
      var sl = sum(ch, function (c) { return c.bh_slice; }), dr = sum(ch, function (c) { return c.bh_drive; });
      var tN = sum(ch, function (c) { return c.trick_n; }), tW = sum(ch, function (c) { return c.trick_w; });
      var aces = sum(ch, function (c) { return c.aces; });
      var D = [
        ['The inside-out forehand', io.n ? pct(io.w, io.n) : '—', 'winners. He ran around his backhand and hit the forehand into the opponent’s backhand corner ' + fmt(io.n) + ' times, which produced ' + pct(io.w, fhW, 0) + ' of his rally forehand winners. He won ' + pct(io.won, io.n, 0) + ' of those points.', miniPattern('L', 'R')],
        ['The one-handed backhand down the line', dtl.n ? pct(dtl.w, dtl.n) : '—', 'winners, from ' + fmt(dtl.n) + ' attempts. That is ' + pct(dtl.w, bhW, 0) + ' of his rally backhand winners, from a shot he chose only ' + pct(dtl.n, sum(rows, function (r) { return r.phase === 'Rally' && r.shot === 'Backhand' ? r.n : 0; }), 0) + ' of the time.', miniPattern('L', 'L')],
        ['The slice', pct(sl, sl + dr, 0), 'of his rally backhands were sliced. He used the slice to reset a rally, change its pace or approach the net.', null],
        ['Serve and volley', snvN ? pct(snvW, snvN, 0) : '—', 'of points won when he served and volleyed (' + fmt(snvN) + ' points, ' + pct(snvN, svN, 1) + ' of his service points).', null],
        ['At the net', netN ? pct(netW, netN, 0) : '—', 'of points won when he came forward (' + fmt(netN) + ' points), against ' + pct(sum(ch, function (c) { return c.fed_points; }), sum(ch, function (c) { return c.n_points; }), 0) + ' overall.', null],
        ['The trick shot', fmt(tN), 'tweeners, behind-the-back shots and other trick shots charted. He won ' + fmt(tW) + ' of those points, since most were last-ditch scrambles. He also hit ' + fmt(aces) + ' aces.', null]
      ];
      D.forEach(function (d) {
        var c = el('div', 'card'); var txt = el('div');
        add(txt, el('div', 'k', d[0]), el('div', 'v num', d[1]), el('div', 'd', d[2]));
        var mini = el('div', 'mini'); if (d[3]) add(mini, d[3]);
        add(c, txt, mini); add(cards, c);
      });
    }

    update();
    return sec;
  }

  // ================================================================== VI. Match explorer
  function renderExplorer(M, CH, SH, SV) {
    var nCharted = CH.filter(function (c) { return c.source !== 'Slam'; }).length, nSlam = CH.length - nCharted;
    var byKey = {}; M.forEach(function (m) { byKey[m.match_key] = m; });
    var chartByKey = {}; CH.forEach(function (c) { chartByKey[c.match_key] = c; });
    var defByKey = {}; M.forEach(function (m) { if (m.defining) defByKey[m.defining] = m; });
    var shBy = {}; SH.forEach(function (r) { (shBy[r.chart_id] || (shBy[r.chart_id] = [])).push(r); });
    var svBy = {}; SV.forEach(function (r) { (svBy[r.chart_id] || (svBy[r.chart_id] = [])).push(r); });

    var sec = el('section', 'ch');
    add(sec, chapterHead('VIII', 'The matches', 'His biggest matches were decided by a few points either way.',
      'Start with one of the career-defining matches, or pick any of ' + fmt(CH.length) + ' matches with point-by-point data (' + fmt(nCharted) + ' charted shot by shot, plus ' + fmt(nSlam) + ' Grand Slam matches with serve speeds). Then replay the key points. The line tracks his running total of points won minus points lost: above the line he is ahead on points, below it he is behind.'));
    var bar = el('div', 'explorer-bar');
    var chips = el('div', 'chips');
    var selEl = el('select', 'matchsel');
    var opt0 = el('option', null, 'Any match with point-by-point data…'); opt0.value = ''; add(selEl, opt0);
    var list = CH.map(function (c) { return byKey[c.match_key]; }).filter(Boolean).sort(function (a, b) { return a.match_date < b.match_date ? 1 : -1; });
    var groups = {};
    list.forEach(function (m) {
      var g = groups[m.season];
      if (!g) { g = groups[m.season] = el('optgroup'); g.label = String(m.season); add(selEl, g); }
      var cc = chartByKey[m.match_key]; var o = el('option', null, m.tournament + ' ' + roundName(m.round) + ' · ' + m.opponent + ' (' + m.result + ')' + (cc && cc.source === 'Slam' ? ' · points only' : ''));
      o.value = m.match_key; add(g, o);
    });
    add(bar, chips, selEl);
    add(sec, bar);
    var panel = el('div', 'match');
    add(sec, reveal(panel));
    var chipEls = [];
    DEFINING.forEach(function (d, i) {
      var m = defByKey[d.key]; if (!m) return;
      var c = el('button', 'chip', d.chip);
      c.addEventListener('click', function () { selEl.value = ''; show(m, d, i); });
      chipEls[i] = c; add(chips, c);
    });
    selEl.addEventListener('change', function () {
      var m = byKey[selEl.value]; if (!m) return;
      var di = -1; DEFINING.forEach(function (d, i) { if (d.key === m.defining) di = i; });
      show(m, di >= 0 ? DEFINING[di] : null, di);
    });

    function show(m, d, di) {
      chipEls.forEach(function (c, j) { if (c) c.className = 'chip' + (j === di ? ' on' + (m.win ? '' : ' loss') : ''); });
      var c = chartByKey[m.match_key];
      panel.textContent = ''; swapIn(panel, 'swap-in');
      add(panel, el('div', 'meta', m.tournament + ' ' + m.season + ' · ' + roundName(m.round) + ' · ' + m.surface));
      var h = el('h3', null, d ? d.title : (m.result === 'W' ? 'Federer d. ' : 'Lost to ') + m.opponent);
      add(h, el('span', 'res ' + (m.win ? 'W' : 'L'), m.win ? 'WON' : 'LOST'));
      add(panel, h);
      add(panel, el('div', 'score num', 'vs ' + m.opponent + '   ' + fedScore(m)));
      if (d) add(panel, el('p', 'story', d.story));
      var ms = el('div', 'mstats auto');
      var shotData = c && c.source !== 'Slam';
      var fp = c ? c.fed_points : m.points_won, tp = c ? c.n_points : m.points_total;
      var stats = [[tp ? fmt(fp) + ' \u2013 ' + fmt(tp - fp) : '\u2014', 'Points'], [tp ? pct(fp, tp) : '\u2014', 'Share of points']];
      if (shotData) stats.push([fmt(c.winners + c.aces) + ' / ' + fmt(c.ue + c.df), 'Winners / unforced errors']);
      stats.push([shotData ? fmt(c.aces) + ' / ' + fmt(c.df) : (m.aces !== null ? fmt(m.aces) + ' / ' + fmt(m.double_faults) : '\u2014'), 'Aces / double faults']);
      if (shotData && c.net_n) stats.push([fmt(c.net_w) + ' / ' + fmt(c.net_n), 'Net points won']);
      if (shotData) stats.push([(c.rally_sum / c.n_points).toFixed(1), 'Avg rally (shots)']);
      if (c && c.avg1_mph) stats.push([fmt(c.avg1_mph) + ' mph', 'Avg first serve']);
      if (c && c.max_mph) stats.push([fmt(c.max_mph) + ' mph', 'Fastest serve']);
      if (c && c.dist_fed_m) stats.push([(c.dist_fed_m / 1000).toFixed(2) + ' km', 'Distance run (opp. ' + (c.dist_opp_m / 1000).toFixed(2) + ')']);
      if (!shotData && m.minutes) stats.push([Math.floor(m.minutes / 60) + 'h ' + ('0' + (m.minutes % 60)).slice(-2) + 'm', 'Duration']);
      stats.forEach(function (x) { var s = el('div', 'mstat'); add(s, el('div', 'v num', x[0]), el('div', 'l', x[1])); add(ms, s); });
      add(panel, ms);
      var chDiv = el('div', 'chart');
      if (c) add(chDiv, momentum(buildSeries(c), m)); else add(chDiv, setBars(m));
      add(panel, chDiv);
      if (shotData) {
        var MS = matchShots(c);
        attachSlam(MS, c);
        add(panel, el('div', 'mini-title rp-title', 'Replay the key points'));
        var rp = replayer(m, c, MS); panel.__rp = rp; add(panel, rp);
        add(panel, el('div', 'mini-title rp-title', 'Where the shots landed'));
        add(panel, matchCourts(m, c, function (pi) {
          try { rp.scrollIntoView({ behavior: 'smooth', block: 'center' }); } catch (e) { rp.scrollIntoView(); }
          if (rp.__playIndex) setTimeout(function () { rp.__playIndex(pi); }, 350);
        }));
        if (MS.speeds.length) add(panel, speedStrip(MS));
      } else if (c) {
        var MS2 = { series: buildSeries(c), points: [] , speeds: [] };
        attachSlam(MS2, c);
        if (MS2.speeds.length) add(panel, speedStrip(MS2));
        add(panel, el('p', 'note', 'This match was not charted shot by shot, so there is no shot map or replay. The point-by-point line and serve speeds come from the Grand Slam\u2019s own tracking data.'));
      }
      var momSvg = panel.querySelector('.mom-view svg');
      var ln = momSvg && momSvg.__reveal ? null : panel.querySelector('path.momline');
      if (momSvg && momSvg.__reveal) momSvg.__reveal(2200);
      if (ln) {
        try {
          var Ltot = ln.getTotalLength();
          if (Ltot) { ln.style.strokeDasharray = Ltot; ln.style.strokeDashoffset = Ltot; ln.getBoundingClientRect();
            ln.style.transition = 'stroke-dashoffset 2.2s ease'; requestAnimationFrame(function () { ln.style.strokeDashoffset = '0'; }); }
        } catch (e) { /* decorative */ }
      }
    }
    var start = DEFINING.length - 1;
    show(defByKey[DEFINING[start].key], DEFINING[start], start);
    // jump straight to a match and replay one point (used by the Pressure chapter)
    sec.__open = function (matchKey, pointIndex) {
      var m = byKey[matchKey]; if (!m) return;
      var di = -1; DEFINING.forEach(function (d, i) { if (d.key === m.defining) di = i; });
      selEl.value = matchKey; show(m, di >= 0 ? DEFINING[di] : null, di);
      panel.scrollIntoView({ behavior: 'smooth', block: 'start' });
      if (panel.__rp && panel.__rp.__playIndex) setTimeout(function () { panel.__rp.__playIndex(pointIndex); }, 450);
    };
    add(sec, el('p', 'note', 'Shot data: Tennis Abstract Match Charting Project. Serve speed, distance run and Slam-only point data: Grand Slam point-by-point data compiled by Jeff Sackmann (IBM SlamTracker / Infosys), 2011\u20132021. Speeds are radar readings of the serve that started the point, shown in mph (converted from km/h where only km/h was published).'));
    return sec;
  }

  // ================================================================== photo interludes
  // Full-bleed photo bands between chapters. Photos are supplied by the app's owner, packaged under content/img/
  // and listed in src/photos.js (window.FED_PHOTOS). Each caption pulls its numbers from the data.
  var bands = [];
  function photoText(key, ctx) {
    var SH = ctx.SH, CH = ctx.CH, M = ctx.M;
    function patt(shot, o, t) { var n = 0, w = 0; SH.forEach(function (r) { if (r.phase === 'Rally' && r.shot === shot && (!o || r.origin === o) && (!t || r.target === t)) { n += r.n; if (r.outcome === 'Winner') w += r.n; } }); return { n: n, w: w }; }
    if (key === 'backhand') {
      var d = patt('Backhand', 'L', 'L'), all = patt('Backhand');
      return { kicker: 'The one-handed backhand', line: 'Hit down the line, ' + pct(d.w, d.n) + ' of them were winners. He chose that shot only ' + pct(d.n, all.n, 0) + ' of the time, yet it produced ' + pct(d.w, all.w, 0) + ' of his backhand winners.' };
    }
    if (key === 'forehand') {
      var io = patt('Forehand', 'L', 'R'), fa = patt('Forehand');
      return { kicker: 'The forehand', line: fmt(fa.w) + ' charted rally forehand winners. Running around the backhand to hit inside-out, ' + pct(io.w, io.n) + ' of those forehands were winners.' };
    }
    if (key === 'serve') {
      var aces = sum(CH, function (c) { return c.aces; }), sv = sum(CH, function (c) { return c.sv_n; }), sw = sum(CH, function (c) { return c.sv_w; });
      return { kicker: 'The serve', line: fmt(aces) + ' aces in the charted matches. He won ' + pct(sw, sv) + ' of the points he served.' };
    }
    if (key === 'volley') {
      var nn = sum(CH, function (c) { return c.net_n; }), nw = sum(CH, function (c) { return c.net_w; });
      return { kicker: 'At the net', line: 'He won ' + pct(nw, nn, 0) + ' of the ' + fmt(nn) + ' points where he came forward.' };
    }
    if (key === 'slice') {
      var sl = sum(CH, function (c) { return c.bh_slice; }), dr = sum(CH, function (c) { return c.bh_drive; });
      return { kicker: 'The slice', line: pct(sl, sl + dr, 0) + ' of his rally backhands were sliced. He used it to change the pace, buy time and set up the net.' };
    }
    if (key === 'usopen') {
      var uo = M.filter(function (m) { return m.tournament === 'US Open'; });
      var yrs = uo.filter(function (m) { return m.title; }).map(function (m) { return m.season; });
      var w = sum(uo, function (m) { return m.win; });
      var run = yrs.length > 1 && yrs[yrs.length - 1] - yrs[0] === yrs.length - 1;
      return { kicker: 'Flushing Meadows', line: yrs.length + (run ? ' straight US Open titles, ' + yrs[0] + '–' + yrs[yrs.length - 1] : ' US Open titles') + '. His record in New York: ' + w + '–' + (uo.length - w) + '.' };
    }
    if (key === 'pressure') {
      var bf = sum(M, function (m) { return m.bp_faced; }), bs = sum(M, function (m) { return m.bp_saved; });
      return { kicker: 'Under pressure', line: 'He faced ' + fmt(bf) + ' break points on his serve and saved ' + pct(bs, bf, 0) + ' of them.' };
    }
    var t = sum(M, function (m) { return m.title; });
    return { kicker: 'Champion', line: fmt(t) + ' titles. ' + fmt(sum(M, function (m) { return m.slam_title; })) + ' majors.' };
  }
  function interlude(photo, ctx) {
    var band = el('figure', 'band' + (photo.side === 'right' ? ' right' : '') + (photo.layout === 'split' ? ' split' : '') + (photo.cap === 'top' ? ' top' : ''));
    band.setAttribute('data-key', photo.key);
    var img = el('img');
    img.src = photo.file; img.alt = photo.alt || ''; img.decoding = 'async';
    if (photo.focus) img.style.objectPosition = photo.focus;
    img.addEventListener('error', function () { band.style.display = 'none'; });
    var cap = el('figcaption');
    var tx = photoText(photo.key, ctx);
    add(cap, el('span', 'eyebrow', tx.kicker), el('span', 'bline', tx.line));
    if (photo.credit) add(cap, el('span', 'credit', photo.credit));
    add(band, img, cap);
    bands.push(band);
    return band;
  }
  function parallax() {
    var vh = window.innerHeight || 800;
    bands.forEach(function (b) {
      var r = b.getBoundingClientRect();
      if (r.bottom < -200 || r.top > vh + 200) return;
      var prog = (r.top + r.height / 2 - vh / 2) / vh;      // −0.5 .. 0.5 while on screen
      var img = b.firstChild;
      if (!b.classList.contains('split')) img.style.transform = 'translate3d(0,' + (prog * -12).toFixed(2) + '%,0) scale(1.08)';
      b.classList.toggle('in', r.top < vh * 0.8);
    });
  }
  var ticking = false;
  function onScroll() { if (!ticking) { ticking = true; requestAnimationFrame(function () { ticking = false; parallax(); }); } }
  window.addEventListener('scroll', onScroll, { passive: true });
  document.addEventListener('scroll', onScroll, { passive: true, capture: true });
  window.addEventListener('resize', onScroll);

  // ================================================================== shot-code parser (Match Charting Project notation)
  var SHOT_INFO = {
    f: ['Forehand', 'FH'], b: ['Backhand', 'BH'], r: ['Forehand slice', 'FH'], s: ['Backhand slice', 'BH'], v: ['Forehand volley', 'FH'], z: ['Backhand volley', 'BH'],
    o: ['Overhead', 'FH'], p: ['Backhand overhead', 'BH'], u: ['Forehand drop shot', 'FH'], y: ['Backhand drop shot', 'BH'], l: ['Forehand lob', 'FH'],
    m: ['Backhand lob', 'BH'], h: ['Forehand half-volley', 'FH'], i: ['Backhand half-volley', 'BH'], j: ['Forehand swinging volley', 'FH'],
    k: ['Backhand swinging volley', 'BH'], t: ['Trick shot', 'FH'], q: ['Unknown shot', 'FH']
  };
  var NET_LETTERS = 'vzopjkhi';
  var MISS = { n: 'Net', w: 'Wide', d: 'Deep', x: 'Wide & deep', '!': 'Shank', e: 'Unknown' };
  var SERVE_DIR = { '4': 'Wide', '5': 'Body', '6': 'T', '0': 'Unknown' };
  function parseServeCode(s) {
    s = s.replace(/^c+/, '');
    var m = /^([0456])(\+)?;?/.exec(s);
    if (!m) return null;
    var rest = s.slice(m[0].length), fault = null, out = null;
    if (/^[nwdxge!V]/.test(rest)) fault = { n: 'Net', w: 'Wide', d: 'Long', x: 'Wide & long', g: 'Foot fault', e: 'Fault', '!': 'Shank', V: 'Time violation' }[rest.charAt(0)];
    else if (rest.charAt(0) === '*') out = 'Ace';
    else if (rest.charAt(0) === '#') out = 'Unreturned';
    return { dir: m[1], snv: !!m[2], fault: fault, out: out, rest: (fault || out) ? '' : rest };
  }
  var RALLY_RE = /([fbrsvzopuylmhijktq])([+\-=;^]*)(\d?)(\d?)([+\-=;^]*)([nwdx!e]?)([*#@]?)/y;
  function parseRally(s) {
    var out = [], pos = 0;
    s = s.replace(/;/g, '');
    while (pos < s.length) {
      if (s.charAt(pos) === 'C') { pos++; continue; }
      RALLY_RE.lastIndex = pos;
      var m = RALLY_RE.exec(s);
      if (!m || m[0].length === 0) break;
      out.push({ letter: m[1], mods: m[2] + m[5], dir: m[3], depth: m[4], err: m[6], term: m[7] });
      pos = RALLY_RE.lastIndex;
    }
    return out;
  }
  // Returns a structured point: serves (incl. a first-serve fault), shots with hitter, and how it ended.
  function parsePoint(code, fedServing) {
    var parts = String(code || '').split('~'), c1 = parts[0] || '', c2 = parts[1] || '';
    var P = { serves: [], shots: [], end: null, serveNo: c2 ? 2 : 1, rally: 0, fedServing: fedServing };
    if (/^[SRPQ]/.test(c1) || (!c1 && !c2)) { P.end = 'Unrecorded'; return P; }
    if (c2) { var f = parseServeCode(c1); if (f) P.serves.push({ no: 1, dir: f.dir, snv: f.snv, fault: f.fault || 'Fault' }); }
    var sv = parseServeCode(c2 || c1);
    if (!sv) { P.end = 'Unrecorded'; return P; }
    P.serves.push({ no: P.serveNo, dir: sv.dir, snv: sv.snv, fault: sv.fault, out: sv.out });
    if (sv.fault) { P.end = 'Double fault'; P.rally = 0; return P; }
    if (sv.out) { P.end = sv.out; P.rally = 1; return P; }
    var R = parseRally(sv.rest);
    R.forEach(function (r, k) {
      var info = SHOT_INFO[r.letter] || SHOT_INFO.q;
      r.name = info[0]; r.side = info[1];
      r.hitterFed = (k % 2 === 0) !== fedServing;     // k = 0 is the return
      r.isReturn = k === 0;
      r.atNet = NET_LETTERS.indexOf(r.letter) >= 0 || r.mods.indexOf('-') >= 0;
      r.approach = r.mods.indexOf('+') >= 0;
      r.outcome = r.term === '*' ? 'Winner' : r.term === '@' ? 'Unforced error' : (r.term === '#' || (r.err && k === R.length - 1)) ? 'Forced error' : 'In play';
      r.miss = (r.outcome === 'Unforced error' || r.outcome === 'Forced error') ? (MISS[r.err] || '') : '';
      P.shots.push(r);
    });
    var last = R[R.length - 1];
    P.rally = 1 + R.length - (last && last.outcome !== 'Winner' && last.outcome !== 'In play' ? 1 : 0);
    P.end = last ? last.outcome : 'In play';
    return P;
  }
  function chartCodes(c) { var a = String(c.codes || '').split('|'); return a.slice(1, a.length - 1); }

  // ================================================================== landing geometry (metres; Federer at the bottom, y > 0)
  function rng(seed) { var t = seed >>> 0; return function () { t += 0x6D2B79F5; var r = Math.imul(t ^ (t >>> 15), 1 | t); r ^= r + Math.imul(r ^ (r >>> 7), 61 | r); return ((r ^ (r >>> 14)) >>> 0) / 4294967296; }; }
  var LX_OPP = { '1': -2.75, '2': 0, '3': 2.75 };    // landing lane on the opponent half (screen x)
  var LX_FED = { '1': 2.75, '2': 0, '3': -2.75 };    // landing lane on Federer's half
  function depthBand(shotLetter, depth, isReturn) {
    if (isReturn && depth === '7') return [1.8, 6.2];
    if (isReturn && depth === '8') return [6.4, 9.4];
    if (isReturn && depth === '9') return [9.4, 11.6];
    if ('uy'.indexOf(shotLetter) >= 0) return [1.0, 5.4];
    if ('lm'.indexOf(shotLetter) >= 0) return [9.2, 11.7];
    if (NET_LETTERS.indexOf(shotLetter) >= 0) return [3.0, 10.5];
    return [6.6, 11.6];
  }
  // Place one landing for a shot hit toward the opponent (top half). Returns {x, y, inCourt}.
  function landingXY(s, R) {
    var lane = LX_OPP[s.dir], hasLane = lane !== undefined;
    // soft (triangular) scatter so the three lanes blend instead of forming hard blocks
    var x = hasLane ? lane + (R() + R() - 1) * (s.dir === '2' ? 1.9 : 2.1) : (R() - 0.5) * 7.6;
    if (x > 4.0) x = 4.0 - R() * 0.3; if (x < -4.0) x = -4.0 + R() * 0.3;
    var band = depthBand(s.letter, s.depth, s.isReturn), mid = (band[0] + band[1]) / 2, half = (band[1] - band[0]) / 2;
    var y = -Math.min(11.8, Math.max(0.6, mid + (R() + R() - 1) * half * 1.25));
    if (s.miss === 'Net') { y = 0.15 + R() * 0.55; }
    else if (s.miss === 'Wide' || s.miss === 'Wide & deep') { var sgn = hasLane && s.dir !== '2' ? Math.sign(lane) : (R() < 0.5 ? -1 : 1); x = sgn * (4.35 + R() * 0.9); if (s.miss === 'Wide & deep') y = -(12.1 + R() * 1.0); }
    else if (s.miss === 'Deep') { y = -(12.1 + R() * 1.0); }
    return { x: x, y: y };
  }

  // Dot court: each landing is a dot placed at random inside its recorded zone. `items` = [{shot, w}] (w = weight for sampling).
  function dotCourt(items, opts) {
    opts = opts || {};
    var c = drawCourt(opts.scale || 16, !!opts.half);
    var R = rng(opts.seed || 7);
    var total = 0; items.forEach(function (it) { total += it.w || 1; });
    var cap = opts.cap || 1600, k = total > cap ? cap / total : 1;
    var lanes = { L: 0, M: 0, R: 0 }, inPlayTot = 0;
    var g = sv('g', {});
    var order = { 'In play': 0, 'Forced error': 1, 'Unforced error': 2, 'Winner': 3 };
    items.slice().sort(function (a, b) { return order[a.shot.outcome] - order[b.shot.outcome]; }).forEach(function (it) {
      var w = it.w || 1, nDots = Math.floor(w * k) + (R() < (w * k) % 1 ? 1 : 0);
      var s = it.shot;
      if (!s.miss && LX_OPP[s.dir] !== undefined) { lanes[{ '1': 'L', '2': 'M', '3': 'R' }[s.dir]] += w; inPlayTot += w; }
      for (var d = 0; d < nDots; d++) {
        var p = landingXY(s, R);
        var col = s.outcome === 'Winner' ? C.gold : s.outcome === 'Unforced error' ? '#C8B3E6' : s.outcome === 'Forced error' ? '#9C8DB2' : '#F7F4EC';
        var r = s.outcome === 'Winner' ? 3.1 : s.miss ? 2.4 : 1.9;
        add(g, sv('circle', { cx: c.X(p.x).toFixed(1), cy: c.Y(p.y).toFixed(1), r: (r * (opts.dotScale || 1)).toFixed(2), fill: col,
          opacity: s.outcome === 'In play' ? 0.55 : 0.92, stroke: s.outcome === 'Winner' ? '#5E4E22' : 'none', 'stroke-width': 0.4 }));
      }
    });
    add(c.svg, g);
    if (!opts.noLabels && inPlayTot) {
      ['L', 'M', 'R'].forEach(function (l) {
        add(c.svg, sv('text', { x: c.X(LANE_X[l]), y: c.Y(-12.35), 'text-anchor': 'middle', fill: '#F7F4EC', 'font-size': opts.small ? 9 : 11, 'font-family': 'var(--sans)' }, Math.round(100 * lanes[l] / inPlayTot) + '%'));
      });
    }
    if (opts.title) add(c.svg, sv('text', { x: c.W / 2, y: c.H - 7, 'text-anchor': 'middle', fill: '#F7F4EC', opacity: 0.85, 'font-size': opts.small ? 9 : 10, 'letter-spacing': '1.5', 'font-family': 'var(--sans)' }, opts.title.toUpperCase()));
    return c.svg;
  }
  function dotLegend() {
    var lg = el('div', 'legend');
    [[C.gold, 'Winner'], ['#F7F4EC', 'In play'], ['#C8B3E6', 'Unforced error'], ['#9C8DB2', 'Forced error']].forEach(function (x) {
      var s = el('span'); var i = el('i'); i.style.background = x[0]; i.style.border = '1px solid #2B5E3B'; add(s, i); s.appendChild(document.createTextNode(x[1])); add(lg, s);
    });
    return lg;
  }
  // Serve dots for one match: in-play serves inside their box third; faults drawn in the net, wide or long.
  function serveDots(serves, opts) {
    opts = opts || {};
    var c = drawCourt(opts.scale || 14, true);
    var R = rng(11);
    var ZX = { Deuce: { '4': [-4.0, -2.8], '5': [-2.7, -1.45], '6': [-1.35, -0.15] }, Ad: { '6': [0.15, 1.35], '5': [1.45, 2.7], '4': [2.8, 4.0] } };
    var g = sv('g', {});
    serves.forEach(function (s) {
      var zx = ZX[s.side][s.dir] || ZX[s.side]['5'];
      var x = zx[0] + R() * (zx[1] - zx[0]), y = -(0.6 + R() * 5.5);
      var col = s.out === 'Ace' ? C.gold : s.out === 'Unreturned' ? '#E4D8B8' : s.fault ? '#C8B3E6' : '#F7F4EC';
      if (s.fault === 'Net') y = 0.15 + R() * 0.5;
      else if (s.fault === 'Long' || s.fault === 'Wide & long') y = -(6.6 + R() * 1.2);
      if (s.fault === 'Wide' || s.fault === 'Wide & long') x = (s.side === 'Deuce' ? -1 : 1) * (s.dir === '6' ? -0.25 - R() * 0.5 : 4.3 + R() * 0.8);
      add(g, sv('circle', { cx: c.X(x).toFixed(1), cy: c.Y(y).toFixed(1), r: s.out === 'Ace' ? 3.4 : 2.4, fill: col, opacity: s.fault ? 0.8 : 0.9, stroke: s.out === 'Ace' ? '#5E4E22' : 'none', 'stroke-width': 0.4 }));
    });
    add(c.svg, g);
    add(c.svg, sv('text', { x: c.X(-2.05), y: c.Y(-7.4), 'text-anchor': 'middle', fill: '#F7F4EC', opacity: 0.75, 'font-size': 9, 'letter-spacing': '1', 'font-family': 'var(--sans)' }, 'DEUCE'));
    add(c.svg, sv('text', { x: c.X(2.05), y: c.Y(-7.4), 'text-anchor': 'middle', fill: '#F7F4EC', opacity: 0.75, 'font-size': 9, 'letter-spacing': '1', 'font-family': 'var(--sans)' }, 'AD'));
    if (opts.title) add(c.svg, sv('text', { x: c.W / 2, y: c.H - 5, 'text-anchor': 'middle', fill: '#F7F4EC', opacity: 0.85, 'font-size': 9, 'letter-spacing': '1.5', 'font-family': 'var(--sans)' }, opts.title.toUpperCase()));
    return c.svg;
  }

  // All of Federer's shots and serves in one charted match, from the raw codes.
  function matchShots(c) {
    var series = buildSeries(c), codes = chartCodes(c);
    var shots = [], serves = [], points = [];
    series.forEach(function (p, i) {
      var P = parsePoint(codes[i], p.server === 'Federer');
      var score = p.game_score.split('-');
      var TV = { '0': 0, '15': 1, '30': 2, '40': 3, 'AD': 4 };
      var a = p.tiebreak ? +score[0] : TV[score[0]], b = p.tiebreak ? +score[1] : TV[score[1]];
      var side = ((a || 0) + (b || 0)) % 2 === 0 ? 'Deuce' : 'Ad';
      P.side = side; P.meta = p; points.push(P);
      if (P.fedServing) P.serves.forEach(function (s) { serves.push({ side: side, dir: s.dir, fault: s.fault, out: s.out, no: s.no }); });
      P.shots.forEach(function (s) { if (s.hitterFed) shots.push(s); });
    });
    return { series: series, points: points, shots: shots, serves: serves };
  }

  // ================================================================== key moments + replay
  function keyMoments(c, MS, oppLast) {
    var L = [];
    function lab(p) { return 'Set ' + p.set_no + ' · ' + p.fed_games + '–' + p.opp_games + (p.tiebreak ? ' (tiebreak)' : '') + ' · ' + p.game_score.replace('-', '–'); }
    function push(list, kind, txt) {
      plist(list).forEach(function (i) {
        var p = MS.series[i]; if (!p) return;
        L.push({ i: i, kind: kind, label: txt(p), sub: lab(p) + ' · ' + (p.server === 'Federer' ? 'Federer' : oppLast) + ' serving' });
      });
    }
    push(c.mp_fed, 'Match point', function (p) { return 'Federer match point — ' + (p.fed_won ? 'converted' : 'saved by ' + oppLast); });
    push(c.mp_opp, 'Match point', function (p) { return oppLast + ' match point — ' + (p.fed_won ? 'saved by Federer' : 'converted'); });
    push(c.sp_fed, 'Set point', function (p) { return 'Federer set point — ' + (p.fed_won ? 'won the set' : 'saved'); });
    push(c.sp_opp, 'Set point', function (p) { return oppLast + ' set point — ' + (p.fed_won ? 'saved by Federer' : 'lost the set'); });
    push(c.bp_fed, 'Break point', function (p) { return 'Federer break point — ' + (p.fed_won ? 'broke' : 'held by ' + oppLast); });
    push(c.bp_opp, 'Break point', function (p) { return oppLast + ' break point — ' + (p.fed_won ? 'saved' : 'broken'); });
    var longest = MS.points.map(function (P, i) { return { i: i, r: P.rally }; }).sort(function (a, b) { return b.r - a.r; }).slice(0, 3);
    longest.forEach(function (o) { var p = MS.series[o.i]; L.push({ i: o.i, r: o.r, kind: 'Longest rallies', label: o.r + '-shot rally — ' + (p.fed_won ? 'won by Federer' : 'won by ' + oppLast), sub: lab(p) }); });
    var seen = {}, out = [];
    L.forEach(function (m) { var k = m.kind + m.i; if (!seen[k]) { seen[k] = 1; out.push(m); } });
    var rank = { 'Match point': 0, 'Set point': 1, 'Break point': 2, 'Longest rallies': 3 };
    return out.sort(function (a, b) { return rank[a.kind] - rank[b.kind] || (a.kind === 'Longest rallies' ? b.r - a.r : a.i - b.i); });
  }

  // Position of a shot's start and end on the court (metres). Approximate by design: zones come from the chart.
  function shotPath(P, k, prevEnd, R) {
    var s = P.shots[k], fed = s.hitterFed;
    var start = prevEnd ? { x: prevEnd.x, y: prevEnd.y } : { x: 0, y: fed ? 11.5 : -11.5 };
    if (!s.atNet && Math.abs(start.y) < 9.5) start.y = (fed ? 1 : -1) * Math.max(Math.abs(start.y) + 1.2, 9.6);
    if (s.atNet) start.y = (fed ? 1 : -1) * (2.2 + R() * 1.6);
    var lane = fed ? LX_OPP[s.dir] : LX_FED[s.dir];
    var band = depthBand(s.letter, s.depth, s.isReturn);
    var end = { x: lane !== undefined ? lane + (R() - 0.5) * 1.4 : (R() - 0.5) * 5, y: (fed ? -1 : 1) * (band[0] + R() * (band[1] - band[0])) };
    if (s.miss === 'Net') end = { x: end.x, y: (fed ? 1 : -1) * 0.3 };
    else if (s.miss === 'Wide' || s.miss === 'Wide & deep') { end.x = (end.x >= 0 ? 1 : -1) * (4.6 + R() * 0.6); if (s.miss === 'Wide & deep') end.y = (fed ? -1 : 1) * 12.6; }
    else if (s.miss === 'Deep') end.y = (fed ? -1 : 1) * (12.4 + R() * 0.6);
    return { start: start, end: end };
  }
  function servePath(P, serve, fedSrv) {
    var deuce = P.side === 'Deuce';
    var sx = (fedSrv ? 1 : -1) * (deuce ? 1 : -1) * 0.9, sy = fedSrv ? 12.1 : -12.1;
    var zone = { '4': 3.4, '5': 2.0, '6': 0.45, '0': 2.0 }[serve.dir] || 2.0;
    if (!serve.fault && P.sw) zone = { W: 3.6, BW: 2.9, B: 2.1, BC: 1.3, C: 0.45 }[P.sw] || zone;
    var tx = (fedSrv ? -1 : 1) * (deuce ? 1 : -1) * zone, ty = (fedSrv ? -1 : 1) * (P.sdp === 'CTL' ? 6.0 : P.sdp === 'NCTL' ? 4.6 : 5.6);
    if (serve.fault === 'Net') ty = (fedSrv ? 1 : -1) * 0.3;
    else if (serve.fault === 'Long' || serve.fault === 'Wide & long') ty = (fedSrv ? -1 : 1) * 7.3;
    if (serve.fault === 'Wide' || serve.fault === 'Wide & long') tx = (tx >= 0 ? 1 : -1) * (serve.dir === '6' ? -0.4 : 4.5);
    return { start: { x: sx, y: sy }, end: { x: tx, y: ty } };
  }

  // Approximate average ball speeds (m/s) by shot letter, used only to pace rally shots relative to one another.
  // These are estimates, not measurements, and are never shown as numbers.
  var PACE = { f: 29, b: 26, r: 19, s: 19, v: 22, z: 21, o: 34, p: 30, u: 11, y: 11, l: 13, m: 13, h: 18, i: 18, j: 27, k: 25, t: 18, q: 22 };
  var RETURN_PACE = { f: 24, b: 23, r: 18, s: 18 };
  var MPH_TO_MS = 0.44704;

  function replayer(m, c, MS) {
    var oppLast = m.opponent.split(' ').slice(-1)[0];
    var oppLefty = String(c.opp_hand || '').trim() === 'L';
    var wrap = el('div', 'replay');
    var left = el('div', 'rp-left'), right = el('div', 'rp-right');
    add(wrap, left, right);
    // ---- stage: the court drawing plus layers, in either the overhead view or the perspective "TV view"
    var viewBar = el('div', 'rp-viewbar');
    var stageBox = el('div', 'rp-stage');
    add(left, viewBar, stageBox);
    var view = 'tv', S = null, capTop = null;
    function playerG(fill, label) {
      var g = sv('g', {});
      var foot = sv('ellipse', { rx: 6, ry: 2.4, fill: 'rgba(0,0,0,.3)' });
      var body = sv('line', { stroke: fill, 'stroke-width': 3.2, 'stroke-linecap': 'round' });
      var ring = sv('circle', { r: 9, fill: 'none', stroke: C.gold, 'stroke-width': 2.2, opacity: 0 });
      var head = sv('circle', { r: 8.5, fill: fill, stroke: '#F7F4EC', 'stroke-width': 1.5 });
      var lab = sv('text', { 'text-anchor': 'middle', fill: '#F7F4EC', 'font-size': 8.5, 'font-family': 'var(--sans)' }, label);
      var tag = sv('text', { 'text-anchor': 'middle', fill: '#F7F4EC', 'font-size': 9, 'font-weight': 600, 'font-family': 'var(--sans)', opacity: 0 });
      var clock = sv('text', { opacity: 0 });
      add(g, foot, body, ring, head, lab, tag, clock);
      g.__foot = foot; g.__body = body; g.__ring = ring; g.__head = head; g.__lab = lab; g.__tag = tag; g.__clock = clock;
      return g;
    }
    // initials from first and last name (Andy Roddick -> AR), like Federer's RF
    var oppInit = (function (nm) { var w = nm.split(/\s+/).filter(Boolean); return ((w[0] || '').charAt(0) + (w.length > 1 ? w[w.length - 1].charAt(0) : (w[0] || '').charAt(1))).toUpperCase(); })(m.opponent);
    var fedDot = playerG(C.ink, 'RF'), oppDot = playerG(C.purple, oppInit);
    var tail = sv('g', {});
    for (var ti = 0; ti < 4; ti++) add(tail, sv('line', { stroke: '#E3F06A', 'stroke-linecap': 'round', opacity: 0 }));
    var shadow = sv('ellipse', { rx: 4, ry: 2, fill: 'rgba(0,0,0,.35)', opacity: 0 });
    var ball = sv('circle', { r: 4.2, fill: '#E3F06A', stroke: '#6E7A1E', 'stroke-width': 0.8, opacity: 0 });
    var pill = sv('g', { opacity: 0 }), pillBg = sv('rect', { rx: 8, ry: 8, height: 17, fill: 'rgba(20,28,24,.82)', stroke: 'rgba(247,244,236,.35)', 'stroke-width': 0.6 }),
      pillTx = sv('text', { x: 7, y: 12.5, fill: '#F7F4EC', 'font-size': 11.5, 'font-family': 'var(--sans)', 'font-weight': 600 });
    add(pill, pillBg, pillTx);
    var fx = sv('g', {});
    function speedText(st) { return st.kind === 'serve' ? (st.mph && !st.fault ? st.mph + ' mph' : null) : '≈' + st.estMph + ' mph'; }

    function makeStage(v) {
      var st0 = { view: v };
      if (v === 'overhead') {
        var c0 = drawCourt(17, false);
        st0.svg = c0.svg; st0.W = c0.W; st0.H = c0.H;
        st0.proj = function (x, y, z) { return [c0.X(x), c0.Y(y) - (z || 0) * c0.s * 0.55, c0.s]; };
      } else {
        // a fixed broadcast camera high behind Federer's baseline, looking down the court
        var cam = { x: 0, y: 11.885 + 14, z: 24 }, tgt = { x: 0, y: 0, z: 0 };   // high behind the baseline, like the main broadcast camera
        var fx0 = tgt.x - cam.x, fy0 = tgt.y - cam.y, fz0 = tgt.z - cam.z, fl = Math.sqrt(fx0 * fx0 + fy0 * fy0 + fz0 * fz0);
        var f = [fx0 / fl, fy0 / fl, fz0 / fl], rgt = [1, 0, 0], up = [0, f[2], -f[1]];
        var raw = function (x, y, z) {
          var d = [x - cam.x, y - cam.y, (z || 0) - cam.z], depth = d[0] * f[0] + d[1] * f[1] + d[2] * f[2];
          return [(d[0] * rgt[0]) / depth, (d[1] * up[1] + d[2] * up[2]) / depth, depth];
        };
        var W = 360, top = 62, pad = 26, bot = 18;   // generous margins so the far player and caption never collide
        var pts = [[-7.4, -14.8], [7.4, -14.8], [-7.4, 14.0], [7.4, 14.0]].map(function (q) { return raw(q[0], q[1], 0); });
        var minX = Math.min.apply(null, pts.map(function (q) { return q[0]; })), maxX = Math.max.apply(null, pts.map(function (q) { return q[0]; }));
        var minY = Math.min.apply(null, pts.map(function (q) { return q[1]; })), maxY = Math.max.apply(null, pts.map(function (q) { return q[1]; }));
        var sc = (W - 2 * pad) / (maxX - minX), H = Math.round(top + (maxY - minY) * sc + bot);
        var offX = (W - (maxX - minX) * sc) / 2;
        st0.W = W; st0.H = H;
        st0.proj = function (x, y, z) { var r = raw(x, y, z); return [offX + (r[0] - minX) * sc, top + (maxY - r[1]) * sc, sc / r[2]]; };
        var P = st0.proj;
        var svg0 = svgRoot(W, H);
        add(svg0, sv('rect', { x: 0, y: 0, width: W, height: H, rx: 6, fill: COURT.surround }));   // same frame green as the overhead court
        function poly(q, attrs) { attrs.points = q.map(function (p) { var s2 = P(p[0], p[1], p[2] || 0); return s2[0].toFixed(1) + ',' + s2[1].toFixed(1); }).join(' '); add(svg0, sv('polygon', attrs)); }
        poly([[-7.4, -14.8], [7.4, -14.8], [7.4, 14.0], [-7.4, 14.0]], { fill: COURT.surround });
        poly([[-5.485, -11.885], [5.485, -11.885], [5.485, 11.885], [-5.485, 11.885]], { fill: COURT.inner });
        function ln(x1, y1, x2, y2, w) { var a = P(x1, y1, 0), b = P(x2, y2, 0); add(svg0, sv('line', { x1: a[0], y1: a[1], x2: b[0], y2: b[1], stroke: COURT.line, 'stroke-width': w || Math.max(1, 0.06 * 17) })); }
        ln(-5.485, -11.885, 5.485, -11.885); ln(-5.485, 11.885, 5.485, 11.885);
        ln(-5.485, -11.885, -5.485, 11.885); ln(5.485, -11.885, 5.485, 11.885);
        ln(-4.115, -11.885, -4.115, 11.885); ln(4.115, -11.885, 4.115, 11.885);
        ln(-4.115, -6.4, 4.115, -6.4); ln(-4.115, 6.4, 4.115, 6.4); ln(0, -6.4, 0, 6.4);
        ln(0, -11.885, 0, -11.6); ln(0, 11.885, 0, 11.6);
        st0.svg = svg0;
        // the net is drawn between the far and near layers so far-side players and balls pass behind it
        var net = sv('g', {});
        var nb1 = P(-6.4, 0, 0), nb2 = P(6.4, 0, 0), nt1 = P(-6.4, 0, 1.07), nt2 = P(6.4, 0, 1.07), ntc = P(0, 0, 0.914);
        add(net, sv('polygon', { points: [nb1, nb2, nt2, ntc, nt1].map(function (q) { return q[0].toFixed(1) + ',' + q[1].toFixed(1); }).join(' '), fill: 'rgba(20,30,24,.38)', stroke: 'none' }));
        add(net, sv('polyline', { points: [nt1, ntc, nt2].map(function (q) { return q[0].toFixed(1) + ',' + q[1].toFixed(1); }).join(' '), fill: 'none', stroke: '#F7F4EC', 'stroke-width': 2.2 }));
        add(net, sv('line', { x1: nb1[0], y1: nb1[1], x2: nt1[0], y2: nt1[1], stroke: '#D8D2C2', 'stroke-width': 2 }), sv('line', { x1: nb2[0], y1: nb2[1], x2: nt2[0], y2: nt2[1], stroke: '#D8D2C2', 'stroke-width': 2 }));
        st0.net = net;
      }
      st0.trails = sv('g', {}); st0.marks = sv('g', {}); st0.far = sv('g', {}); st0.near = sv('g', {});
      add(st0.svg, st0.trails, st0.marks, st0.far);
      if (st0.net) add(st0.svg, st0.net);
      add(st0.svg, st0.near, fx, pill);
      st0.cap = sv('text', { x: st0.W / 2, y: v === 'tv' ? 30 : 20, 'text-anchor': 'middle', fill: '#F7F4EC', 'font-family': 'var(--sans)', 'font-size': 11, opacity: 0.92 });
      add(st0.svg, st0.cap);
      add(st0.near, fedDot); add(st0.far, oppDot); add(st0.near, shadow, tail, ball);
      return st0;
    }
    function setView(v) {
      var oldCap = capTop;
      // camera move: the old view tilts away as the new one tilts in
      var ghost = null;
      if (S && S.svg && S.svg.parentNode === stageBox) {
        ghost = S.svg.cloneNode(true); ghost.setAttribute('class', 'rp-ghost ' + (v === 'tv' ? 'to-tv' : 'to-top'));
      }
      view = v; stageBox.textContent = '';
      S = makeStage(v); capTop = S.cap;
      if (oldCap) { capTop.textContent = oldCap.textContent; capTop.setAttribute('font-family', oldCap.getAttribute('font-family')); capTop.setAttribute('font-size', oldCap.getAttribute('font-size')); }
      add(stageBox, S.svg);
      stageBox.classList.toggle('top', v !== 'tv');
      if (ghost) {
        add(stageBox, ghost);
        swapIn(S.svg, v === 'tv' ? 'cam-in-tv' : 'cam-in-top');
        setTimeout(function () { if (ghost.parentNode) ghost.parentNode.removeChild(ghost); }, 700);
      }
      bTV.classList.toggle('on', v === 'tv'); bTop.classList.toggle('on', v !== 'tv');
      if (T) { T.curK = -2; render(T.now); } else { placeP(fedDot, { x: 0.8, y: 12.4 }); placeP(oppDot, { x: -0.8, y: -12.4 }); }
    }
    // which layer an actor belongs to (far side of the net is drawn behind it in the TV view)
    function layer(node, y) { var g = (S.view === 'tv' && y < 0) ? S.far : S.near; if (node.parentNode !== g) g.appendChild(node); }
    var bTV = el('button', 'seg on', 'TV view'), bTop = el('button', 'seg', 'Overhead view');
    add(viewBar, el('span', 'rp-vl', 'Camera'), bTV, bTop);
    var prog = el('div', 'rp-prog'), progFill = el('div', 'rp-prog-fill'); add(prog, progFill); add(left, prog);
    var reactBox = el('div', 'rp-react');
    var rx = el('div', 'rp-rx'), rxHead = el('div', 'rp-rx-head'), rxLab = el('span', 'rp-rx-lab', 'Time to react'), rxVal = el('span', 'rp-rx-val', '');
    var rxTrack = el('div', 'rp-rx-track'), rxFill = el('div', 'rp-rx-fill');
    add(rxHead, rxLab, rxVal); add(rxTrack, rxFill); add(rx, rxHead, rxTrack); add(left, rx);
    var ctr = el('div', 'rp-ctr');
    var bPrev = el('button', 'pill', '⏮'), bPlay = el('button', 'pill on', '❚❚ Pause'), bNext = el('button', 'pill', '⏭');
    var bReal = el('button', 'pill on', 'Real time'), bSlow = el('button', 'pill', 'Slow motion');
    bPrev.title = 'Previous shot'; bNext.title = 'Next shot';
    var counter = el('span', 'rp-count', '');
    add(ctr, bPrev, bPlay, bNext, el('span', 'rp-sep'), bReal, bSlow, counter); add(left, ctr);

    var board = el('div', 'rp-board'), stakes = el('div', 'rp-stakes'), info = el('div', 'rp-info'), log = el('ol', 'rp-log');
    var momentsBox = el('div', 'rp-moments');
    var hud = el('div', 'rp-hud');
    var hudSpeed = el('div', 'rp-tile'), hudReact = el('div', 'rp-tile');
    add(hudSpeed, el('div', 'rp-tv', '\u2014'), el('div', 'rp-tl', 'Ball speed'));
    add(hudReact, el('div', 'rp-tv', '\u2014'), el('div', 'rp-tl', 'Time to react'));
    add(hud, hudSpeed, hudReact);
    add(right, el('div', 'mini-title', 'Key moments'), momentsBox, board, stakes, info, hud, el('div', 'mini-title', 'Shot by shot'), log);

    var moments = keyMoments(c, MS, oppLast);
    var sets = fedSets(m);
    var current = null, anim = null, tempo = 1, paused = false, T = null;
    function placeP(g, p) {
      var gr = S.proj(p.x, p.y, 0), tv = S.view === 'tv';
      var hd = tv ? S.proj(p.x, p.y, 1.75) : gr;
      var k = gr[2], r = tv ? Math.max(7, Math.min(11, 0.62 * k)) : 8.5;
      g.__foot.setAttribute('cx', gr[0].toFixed(1)); g.__foot.setAttribute('cy', gr[1].toFixed(1));
      g.__foot.setAttribute('rx', (tv ? Math.max(3, 0.35 * k) : 0).toFixed(1)); g.__foot.setAttribute('ry', (tv ? Math.max(1.2, 0.14 * k) : 0).toFixed(1));
      g.__body.setAttribute('x1', gr[0].toFixed(1)); g.__body.setAttribute('y1', gr[1].toFixed(1)); g.__body.setAttribute('x2', hd[0].toFixed(1)); g.__body.setAttribute('y2', hd[1].toFixed(1));
      g.__body.setAttribute('stroke-width', tv ? Math.max(2, 0.16 * k).toFixed(1) : 0);
      g.__head.setAttribute('cx', hd[0].toFixed(1)); g.__head.setAttribute('cy', hd[1].toFixed(1)); g.__head.setAttribute('r', r.toFixed(1));
      g.__ring.setAttribute('cx', hd[0].toFixed(1)); g.__ring.setAttribute('cy', hd[1].toFixed(1));
      g.__lab.setAttribute('x', hd[0].toFixed(1)); g.__lab.setAttribute('y', (hd[1] + r * 0.38).toFixed(1)); g.__lab.setAttribute('font-size', (r * 0.98).toFixed(1));
      var below = !tv && g === oppDot, side = tv && g === oppDot;
      g.__tag.setAttribute('text-anchor', side ? 'start' : 'middle');
      g.__tag.setAttribute('x', (side ? hd[0] + r + 5 : hd[0]).toFixed(1)); g.__tag.setAttribute('y', (below ? hd[1] + r + 12 : side ? hd[1] + 3 : hd[1] - r - 5).toFixed(1));
      g.__r = r;
      layer(g, p.y);
    }

    function renderMoments() {
      momentsBox.textContent = '';
      if (!moments.length) { add(momentsBox, el('p', 'note', 'No key points were found in this chart.')); return; }
      var groups = {};
      moments.forEach(function (mo) { (groups[mo.kind] || (groups[mo.kind] = [])).push(mo); });
      Object.keys(groups).forEach(function (kd) {
        var g = el('div', 'rp-group'); add(g, el('div', 'rp-kind', kd + ' (' + groups[kd].length + ')'));
        groups[kd].slice(0, kd === 'Break point' ? 12 : 8).forEach(function (mo) {
          var b = el('button', 'rp-item' + (current === mo ? ' on' : ''));
          add(b, el('span', 'rp-l', mo.label), el('span', 'rp-s', mo.sub));
          b.addEventListener('click', function () { play(mo); });
          add(g, b);
        });
        add(momentsBox, g);
      });
    }
    function scoreboard(p, P, note) {
      board.textContent = '';
      var t = el('table', 'sb');
      [['Federer', true], [m.opponent, false]].forEach(function (row) {
        var tr = el('tr');
        var srv = p && ((p.server === 'Federer') === row[1]);
        add(tr, el('td', 'sb-name', (srv ? '● ' : '') + row[0]));
        var nSets = p ? p.set_no - 1 : sets.length;
        for (var k = 0; k < nSets; k++) {
          var g = (sets[k] || '').replace(/\(.*\)/, '').split('-');
          add(tr, el('td', 'sb-set', row[1] ? g[0] : g[1]));
        }
        if (p) {
          add(tr, el('td', 'sb-cur', String(row[1] ? p.fed_games : p.opp_games)));
          var gs = p.game_score.split('-'); var fs = p.server === 'Federer';
          add(tr, el('td', 'sb-pts', row[1] ? (fs ? gs[0] : gs[1]) : (fs ? gs[1] : gs[0])));
        }
        add(t, tr);
      });
      add(board, t);
      if (note) add(board, el('div', 'sb-note', note));
    }
    function stakesLine(mo, p) {
      var i = mo.i, opp = oppLast, has = function (list) { return plist(list).indexOf(i) >= 0; };
      if (has(c.mp_fed)) return 'At stake: if Federer wins this point, he wins the match.';
      if (has(c.mp_opp)) return 'At stake: if ' + opp + ' wins this point, he wins the match.';
      if (has(c.sp_fed)) return 'At stake: if Federer wins this point, he wins set ' + p.set_no + '.';
      if (has(c.sp_opp)) return 'At stake: if ' + opp + ' wins this point, he wins set ' + p.set_no + '.';
      if (has(c.bp_fed)) return 'At stake: if Federer wins this point, he breaks ' + opp + '’s serve.';
      if (has(c.bp_opp)) return 'At stake: if ' + opp + ' wins this point, he breaks Federer’s serve.';
      return 'A ' + MS.points[i].rally + '-shot rally, one of the longest of the match.';
    }
    function describe(s0) {
      var who = s0.hitterFed ? 'Federer' : oppLast;
      var dir = s0.hitterFed ? { '1': 'to the forehand corner', '2': 'down the middle', '3': 'to the backhand corner' }[s0.dir] : { '1': 'to Federer’s forehand', '2': 'down the middle', '3': 'to Federer’s backhand' }[s0.dir];
      var dep = s0.isReturn ? { '7': ', short', '8': ', mid-court', '9': ', deep' }[s0.depth] || '' : '';
      var end = s0.outcome === 'Winner' ? ' — winner' : s0.outcome === 'In play' ? '' : ' — ' + s0.outcome.toLowerCase() + (s0.miss ? ' (' + s0.miss.toLowerCase() + ')' : '');
      return who + ': ' + (s0.isReturn ? s0.name.toLowerCase() + ' return' : s0.name.toLowerCase()) + (s0.approach ? ' approach' : '') + (dir ? ' ' + dir : '') + dep + end;
    }
    // Which side of the ball the player stands on at contact (screen x offset). Federer is right-handed.
    function sideOffset(fed, shot) {
      if (!shot) return 0;
      var fh = shot.side === 'FH', righty = fed ? true : !oppLefty;
      var ballOnScreenRight = (fh === righty) === fed;   // a right-hander's forehand is on his right; Federer faces up the screen
      return ballOnScreenRight ? -0.55 : 0.55;           // the player stands on the other side of the ball
    }

    function build(mo) {
      var p = MS.series[mo.i], P = MS.points[mo.i];
      var R = rng(1000 + mo.i), steps = [];
      var sgn = function (fed) { return fed ? 1 : -1; };
      var clamp = function (v, lo, hi) { return Math.max(lo, Math.min(hi, v)); };
      var measured = !!P.speed;
      P.serves.forEach(function (sv0) {
        var path = servePath(P, sv0, P.fedServing);
        var mph = measured && !sv0.fault ? P.speed : (sv0.no === 1 ? 116 : 94);   // unmeasured or fault serves use a typical pace
        steps.push({ A: path.start, B: path.end, fed: P.fedServing, kind: 'serve', h: 0.5, fault: !!sv0.fault, ace: sv0.out === 'Ace', unret: sv0.out === 'Unreturned', snv: sv0.snv,
          v: mph * MPH_TO_MS * 0.78, v2: mph * MPH_TO_MS * 0.55, mph: measured && !sv0.fault ? P.speed : null, tag: 'SERVE',
          short: (P.fedServing ? 'Federer' : oppLast) + ' · ' + (sv0.no === 1 ? '1st' : '2nd') + ' serve' + (sv0.fault ? ' · fault' : '') + (measured && !sv0.fault ? ' · ' + P.speed + ' mph' : ''),
          text: (P.fedServing ? 'Federer' : oppLast) + ': ' + (sv0.no === 1 ? 'first' : 'second') + ' serve ' + ({ '4': 'out wide', '5': 'into the body', '6': 'down the T' }[sv0.dir] || '') + (sv0.snv ? ', serve and volley' : '') +
            (sv0.fault ? ' — fault (' + sv0.fault.toLowerCase() + ')' : sv0.out ? ' — ' + sv0.out.toLowerCase() : '') + (measured && !sv0.fault ? ' · ' + P.speed + ' mph' : '') });
      });
      var prevB = null;
      P.shots.forEach(function (s0, k) {
        var path = shotPath(P, k, prevB, R);
        var v = (s0.isReturn ? (RETURN_PACE[s0.letter] || PACE[s0.letter]) : PACE[s0.letter]) || 22;
        steps.push({ B: path.end, fed: s0.hitterFed, shot: s0, kind: s0.atNet ? 'volley' : ('lm'.indexOf(s0.letter) >= 0 ? 'lob' : 'drive'),
          h: 'lm'.indexOf(s0.letter) >= 0 ? 2.8 : s0.atNet ? 0.5 : 1.2, winner: s0.outcome === 'Winner', error: s0.outcome !== 'Winner' && s0.outcome !== 'In play',
          v: v, v2: v * 0.7, tag: s0.atNet ? (s0.side === 'FH' ? 'FH VOLLEY' : 'BH VOLLEY') : (s0.side === 'FH' ? 'FOREHAND' : 'BACKHAND'),
          short: (s0.hitterFed ? 'Federer' : oppLast) + ' ' + (s0.isReturn ? 'return' : s0.name.toLowerCase()), text: describe(s0) });
        prevB = path.end;
      });
      for (var k = 0; k < steps.length; k++) {
        var st = steps[k], nx = steps[k + 1];
        if (!st.A) st.A = steps[k - 1].C || steps[k - 1].B;
        var recvFed = !st.fed;
        if (st.fault) { st.C = null; continue; }
        if (nx && nx.kind !== 'serve') {
          if (nx.shot && nx.shot.atNet) {
            var yN = sgn(recvFed) * (2.4 + R() * 1.4), t = (yN - st.A.y) / ((st.B.y - st.A.y) || 1);
            st.C = { x: clamp(st.A.x + (st.B.x - st.A.x) * clamp(t, 0, 1.3), -5, 5), y: yN }; st.noBounce = true;
          } else {
            var dx = st.B.x - st.A.x, by = Math.abs(st.B.y), cy = by < 6.5 ? by + 1.4 : clamp(by + 2.6, 11.0, 13.4);
            st.C = { x: clamp(st.B.x + dx * 0.22, -5.4, 5.4), y: sgn(recvFed) * cy };
          }
          st.nextShot = nx.shot;
        } else if (st.winner || st.ace) {
          var ux = st.B.x - st.A.x, uy = st.B.y - st.A.y, L = Math.sqrt(ux * ux + uy * uy) || 1;
          st.C = { x: clamp(st.B.x + ux / L * 4.5, -6.6, 6.6), y: clamp(st.B.y + uy / L * 4.5, -14.2, 14.2) }; st.past = true;
        } else if (st.unret) {
          st.C = { x: clamp(st.B.x * 1.25, -5.4, 5.4), y: sgn(recvFed) * 12.2 };
        } else st.C = null;
      }
      // deterministic player positions per step, so the replay can be paused and stepped through
      var deuce = P.side === 'Deuce';
      function startPos() {
        var srvFed = P.fedServing, sx = (srvFed ? 1 : -1) * (deuce ? 1 : -1) * 1.0;
        var server = { x: sx, y: (srvFed ? 1 : -1) * 12.6 }, recv = { x: -sx * 2.8, y: (srvFed ? -1 : 1) * 12.4 };
        return srvFed ? { fed: server, opp: recv } : { fed: recv, opp: server };
      }
      var dist = function (a, b) { var dx = b.x - a.x, dy = b.y - a.y; return Math.sqrt(dx * dx + dy * dy); };
      var pos = startPos(), t0 = 600;                           // the toss
      steps.forEach(function (st, k) {
        if (st.kind === 'serve' && k > 0) pos = startPos();
        st.from = { fed: { x: pos.fed.x, y: pos.fed.y }, opp: { x: pos.opp.x, y: pos.opp.y } };
        var hitter = st.fed ? 'fed' : 'opp', recv = st.fed ? 'opp' : 'fed';
        var sh = st.shot, toNet = (st.kind === 'serve' && st.snv) || (sh && (sh.approach || sh.atNet));
        st.to = { fed: { x: pos.fed.x, y: pos.fed.y }, opp: { x: pos.opp.x, y: pos.opp.y } };
        if (!st.fault) st.to[hitter] = toNet ? { x: st.A.x * 0.5, y: (st.fed ? 1 : -1) * 3.6 } : { x: st.A.x * 0.3, y: (st.fed ? 1 : -1) * 12.4 };
        if (st.C) {
          var off = st.nextShot ? sideOffset(!st.fed, st.nextShot) : 0;
          st.to[recv] = st.past ? { x: pos[recv].x + (st.C.x - pos[recv].x) * 0.45, y: pos[recv].y + (st.C.y - pos[recv].y) * 0.45 } : { x: st.C.x + off, y: st.C.y + (st.fed ? -0.3 : 0.3) };
        }
        st.D1 = 1000 * (st.noBounce ? dist(st.A, st.C) : dist(st.A, st.B)) / st.v;
        st.D2 = st.C && !st.noBounce ? 1000 * dist(st.B, st.C) / st.v2 : 0;
        // time the receiver has from the opponent's contact to his own contact (none for faults, winners and aces)
        st.react = st.C && !st.past && !st.fault ? (st.D1 + (st.noBounce ? 0 : st.D2)) / 1000 : null;
        st.measured = st.kind === 'serve' && !!st.mph;
        st.estMph = st.kind === 'serve' ? (st.mph || null) : Math.round(st.v / 0.78 * 2.23694);   // launch speed implied by the average pace
        st.t0 = t0; st.k = k;
        t0 += st.D1 + st.D2 + (st.fault ? 900 : 40);
        pos = st.to;
      });
      return { p: p, P: P, steps: steps, end: t0 + 200 };
    }

    // Two-row chart for the point: ball speed per shot (hitter's colour) and time the receiver had to react.
    function drawReact(T) {
      reactBox.textContent = '';
      var steps = T.steps, n = steps.length;
      add(reactBox, el('div', 'mini-title rp-title2', 'Shot by shot: speed and time to react'));
      var W = 360, ml = 34, rowH = 78, gap = 26, mt = 14, H = mt + rowH * 2 + gap + 18;
      var bw = (W - ml - 4) / n;
      var svgR = svgRoot(W, H);
      var spMax = 140, rxMax = Math.max(1.6, Math.max.apply(null, steps.map(function (st) { return st.react || 0; })) * 1.1);
      var ys = scale(0, spMax, mt + rowH, mt), yr = scale(0, rxMax, mt + rowH * 2 + gap, mt + rowH + gap);
      add(svgR, sv('text', { x: 0, y: mt - 4, class: 'lbl-ink' }, 'Ball speed'));
      add(svgR, sv('text', { x: 0, y: mt + rowH + gap - 4, class: 'lbl-ink' }, 'Time to react'));
      [0, 50, 100].forEach(function (t) { add(svgR, sv('line', { x1: ml, x2: W, y1: ys(t), y2: ys(t), stroke: C.rule })); add(svgR, sv('text', { x: ml - 4, y: ys(t) + 3, 'text-anchor': 'end', class: 'lbl' }, t)); });
      add(svgR, sv('text', { x: 62, y: mt - 4, class: 'lbl' }, '(mph)'));
      [0, 0.5, 1, 1.5].forEach(function (t) { if (t > rxMax) return; add(svgR, sv('line', { x1: ml, x2: W, y1: yr(t), y2: yr(t), stroke: C.rule })); add(svgR, sv('text', { x: ml - 4, y: yr(t) + 3, 'text-anchor': 'end', class: 'lbl' }, t + 's')); });
      var sumF = 0, nF = 0, sumO = 0, nO = 0;
      steps.forEach(function (st, k) {
        var x0 = (ml + 2 + k * bw).toFixed(1), w0 = Math.max(1.5, bw - 1.5).toFixed(1);
        var spd = st.kind === 'serve' ? (st.fault ? null : st.mph) : st.estMph, hitCol = st.fed ? C.grass : C.purple;
        if (spd) {
          var r1 = sv('rect', { class: 'rb', 'data-k': k, x: x0, y: ys(Math.min(spMax, spd)).toFixed(1), width: w0, height: (ys(0) - ys(Math.min(spMax, spd))).toFixed(1),
            fill: hitCol, opacity: 0.7, 'fill-opacity': st.kind === 'serve' ? 1 : 0.55, stroke: hitCol, 'stroke-width': 0.6 });
          r1.addEventListener('mousemove', function (e) { showTip(e, 'Shot ' + (k + 1) + ': ' + (st.kind === 'serve' ? spd + ' mph' : '\u2248' + spd + ' mph'), [(st.fed ? 'Federer' : oppLast) + ' hitting', st.kind === 'serve' ? 'Radar serve speed' : 'Typical speed for this shot type (estimate)']); });
          r1.addEventListener('mouseleave', hideTip); r1.addEventListener('click', function () { seekStep(k); });
          add(svgR, r1);
        }
        if (st.react) {
          var recvFed = !st.fed, col = recvFed ? C.grass : C.purple;
          if (recvFed) { sumF += st.react; nF++; } else { sumO += st.react; nO++; }
          var r2 = sv('rect', { class: 'rb', 'data-k': k, x: x0, y: yr(st.react).toFixed(1), width: w0, height: (yr(0) - yr(st.react)).toFixed(1),
            fill: col, opacity: 0.7, 'fill-opacity': st.measured ? 1 : 0.55, stroke: col, 'stroke-width': 0.6 });
          r2.addEventListener('mousemove', function (e) { showTip(e, 'Shot ' + (k + 1) + ': ' + st.react.toFixed(2) + ' s to react', [(recvFed ? 'Federer' : oppLast) + ' receiving', st.measured ? 'From the radar serve speed' : 'Estimated from shot type']); });
          r2.addEventListener('mouseleave', hideTip); r2.addEventListener('click', function () { seekStep(k); });
          add(svgR, r2);
        }
      });
      add(svgR, sv('text', { x: ml, y: H - 3, class: 'lbl' }, 'Shot 1'));
      add(svgR, sv('text', { x: W, y: H - 3, 'text-anchor': 'end', class: 'lbl' }, 'Shot ' + n));
      add(reactBox, svgR);
      var lg = el('div', 'legend');
      var a1 = el('span'); var i1 = el('i'); i1.style.background = C.grass; add(a1, i1); a1.appendChild(document.createTextNode('Federer' + (nF ? ' \u00b7 ' + (sumF / nF).toFixed(2) + ' s avg to react' : ''))); add(lg, a1);
      var a2 = el('span'); var i2 = el('i'); i2.style.background = C.purple; add(a2, i2); a2.appendChild(document.createTextNode(oppLast + (nO ? ' \u00b7 ' + (sumO / nO).toFixed(2) + ' s avg to react' : ''))); add(lg, a2);
      add(reactBox, lg);
      add(reactBox, el('p', 'note', 'Solid bars are measured (radar serve speed). Lighter bars are estimates from the shot type. Click a bar to jump to that shot.'));
    }
    function play(mo) {
      current = mo; renderMoments();
      T = build(mo);
      var p = T.p, P = T.P;
      scoreboard(p, P);
      stakes.textContent = stakesLine(mo, p);
      info.textContent = '';
      var lines = [mo.label, (P.serveNo === 1 ? 'First' : 'Second') + ' serve · ' + (P.end === 'In play' ? '' : P.end + ' · ') + 'rally of ' + P.rally + (P.rally === 1 ? ' shot' : ' shots')];
      if (P.speed) lines.splice(1, 0, 'Serve speed ' + P.speed + ' mph (radar)');
      lines.forEach(function (l, k) { add(info, el('div', k === 0 ? 'rp-head' : 'rp-line', l)); });
      log.textContent = '';
      T.items = T.steps.map(function (st) {
        var spd = st.kind === 'serve' ? '' : ' \u00b7 \u2248' + st.estMph + ' mph';
        var extra = spd + (st.react ? ' \u00b7 ' + st.react.toFixed(2) + ' s to react' + (st.measured ? '' : ' (est.)') : '');
        var li = el('li', null, st.text + extra); add(log, li); return li;
      });
      drawReact(T);
      T.now = 0; T.curK = -2; T.done = false;
      capTop.setAttribute('font-family', 'var(--sans)'); capTop.setAttribute('font-size', 11);
      setPaused(false);
      start();
    }
    function setPaused(v) { paused = v; bPlay.textContent = v ? '▶ Play' : '❚❚ Pause'; }
    function seekStep(k) {
      if (!T) return;
      k = Math.max(0, Math.min(T.steps.length - 1, k));
      T.now = T.steps[k].t0; T.curK = -2; T.done = false;
      capTop.setAttribute('font-family', 'var(--sans)'); capTop.setAttribute('font-size', 11);
      scoreboard(T.p, T.P);
      render(T.now); setPaused(true);
    }
    function start() {
      if (anim) cancelAnimationFrame(anim);
      var last = null;
      function frame(ts) {
        if (last === null) last = ts;
        var dt = Math.min(100, ts - last); last = ts;
        if (!paused && !T.done) T.now += dt / tempo;
        render(T.now);
        anim = requestAnimationFrame(frame);
      }
      anim = requestAnimationFrame(frame);
    }
    var ease = function (t) { return t < 0.5 ? 2 * t * t : 1 - Math.pow(-2 * t + 2, 2) / 2; };
    var lerp = function (a, b, t) { return { x: a.x + (b.x - a.x) * t, y: a.y + (b.y - a.y) * t }; };
    // ball height (metres) along a shot: contact height, net clearance and the kick after the bounce
    function zModel(st) {
      var L0 = st.shot ? st.shot.letter : '';
      var zA = st.kind === 'serve' ? 2.8 : (st.shot && st.shot.atNet ? 1.1 : 1.0);
      var apex1 = st.kind === 'serve' ? 0.15 : 'lm'.indexOf(L0) >= 0 ? 5.0 : 'uy'.indexOf(L0) >= 0 ? 1.2 : 'rs'.indexOf(L0) >= 0 ? 0.75 : 'op'.indexOf(L0) >= 0 ? 0.1 : (st.shot && st.shot.atNet) ? 0.3 : 1.35;
      var zB = Math.abs(st.B.y) < 0.6 ? 0.45 : 0;   // balls that end in the net stop at net height
      var apex2 = 'lm'.indexOf(L0) >= 0 ? 1.5 : 'rsuy'.indexOf(L0) >= 0 ? 0.35 : st.kind === 'serve' ? 0.5 : 0.7;
      var zC = st.past ? 0.4 : 1.0;
      return { zA: zA, apex1: apex1, zB: zB, apex2: apex2, zC: zC };
    }
    function ballAt(st, u) {
      var Z = st.__z || (st.__z = zModel(st));
      if (u <= st.D1 || !st.D2) {
        var a = Math.max(0, Math.min(1, u / st.D1)), tgt = st.noBounce ? st.C : st.B, zEnd = st.noBounce ? 1.1 : Z.zB, apex = st.noBounce ? 0.6 : Z.apex1;
        return { x: st.A.x + (tgt.x - st.A.x) * a, y: st.A.y + (tgt.y - st.A.y) * a, z: Z.zA * (1 - a) + zEnd * a + 4 * apex * a * (1 - a), v: st.v, seg: 1, a: a };
      }
      var b = Math.max(0, Math.min(1, (u - st.D1) / st.D2));
      return { x: st.B.x + (st.C.x - st.B.x) * b, y: st.B.y + (st.C.y - st.B.y) * b, z: Z.zC * b + 4 * Z.apex2 * b * (1 - b), v: st.v2, seg: 2, a: b };
    }
    function rebuildMarks(k, u) {
      S.marks.textContent = ''; S.trails.textContent = '';
      var tv = S.view === 'tv';
      for (var j = Math.max(0, k - 3); j <= k; j++) {
        var st = T.steps[j]; if (!st) continue;
        var bright = j === k, tot = st.D1 + st.D2, pts = [];
        for (var q = 0; q <= 24; q++) { var bp = ballAt(st, tot * q / 24); var pr = S.proj(bp.x, bp.y, bp.z); pts.push(pr[0].toFixed(1) + ',' + pr[1].toFixed(1)); }
        add(S.trails, sv('polyline', { points: pts.join(' '), fill: 'none', stroke: st.fed ? '#F7F4EC' : '#E3D6F5', 'stroke-width': bright ? 1.4 : 1,
          opacity: bright ? 0.7 : 0.1 + 0.06 * (j - k + 3), 'stroke-linejoin': 'round' }));
        if (!st.noBounce && (j < k || u > st.D1)) {
          var gold = st.winner || st.ace, bad = st.error || st.fault, bpj = S.proj(st.B.x, st.B.y, 0), kk = bpj[2];
          var rr = gold ? Math.max(3.5, 0.32 * kk) : Math.max(2, 0.16 * kk);
          add(S.marks, sv('ellipse', { cx: bpj[0].toFixed(1), cy: bpj[1].toFixed(1), rx: rr.toFixed(1), ry: (tv ? rr * 0.42 : rr).toFixed(1), fill: gold ? C.gold : 'none',
            stroke: gold ? '#5E4E22' : bad ? '#C8B3E6' : '#F7F4EC', 'stroke-width': 1.1, opacity: bright || gold ? 0.95 : 0.25 + 0.15 * (j - k + 3) }));
        }
      }
    }
    function render(now) {
      if (!T) return;
      var steps = T.steps, st = null;
      for (var i = steps.length - 1; i >= 0; i--) if (now >= steps[i].t0) { st = steps[i]; break; }
      var waiting = !st;
      if (!st) st = steps[0];
      var u = waiting ? 0 : now - st.t0, tot = st.D1 + st.D2;
      if (st.k !== T.curK) {
        T.curK = st.k;
        T.items.forEach(function (li, j) { li.className = j === st.k ? 'on' : (j < st.k ? 'done' : ''); });
        var li = T.items[st.k]; if (li) log.scrollTop = Math.max(0, li.offsetTop - log.offsetTop - 70);
        capTop.textContent = (st.k + 1) + ' / ' + steps.length + ' · ' + st.short;
        counter.textContent = 'Shot ' + (st.k + 1) + ' of ' + steps.length + (tempo === 1 ? ' · real time' : ' · 1/' + tempo + ' speed');
        T.markedBounce = false; fx.textContent = ''; T.mphAt = undefined;
        rebuildMarks(st.k, u);
        var og = st.fed ? oppDot : fedDot;
        og.__tag.setAttribute('opacity', 0); og.__ring.setAttribute('opacity', 0);
        (st.fed ? fedDot : oppDot).__tag.textContent = st.tag;
        fedDot.__clock.setAttribute('opacity', 0); oppDot.__clock.setAttribute('opacity', 0);
        if (st.react) {
          rxLab.textContent = (st.fed ? oppLast : 'Federer') + ' has ' + st.react.toFixed(2) + ' s to react' + (st.measured ? '' : ' (est.)');
          rxFill.style.background = st.fed ? C.purple : C.grass; rxFill.style.width = '100%';
        } else { rxLab.textContent = 'Time to react'; rxFill.style.width = '0%'; }
        var sp = st.kind === 'serve' ? (st.fault ? null : (st.mph ? st.mph + ' mph' : null)) : '\u2248' + st.estMph + ' mph';
        hudSpeed.firstChild.textContent = sp || '\u2014';
        hudSpeed.lastChild.textContent = st.kind === 'serve' ? (st.fault ? 'Fault \u00b7 speed not recorded' : st.mph ? 'Serve speed \u00b7 radar' : 'Serve speed \u00b7 not recorded') : 'Ball speed \u00b7 estimated from shot type';
        hudReact.firstChild.textContent = st.react ? st.react.toFixed(2) + ' s' : (st.past ? 'No chance' : '\u2014');
        hudReact.lastChild.textContent = st.react ? 'Time for ' + (st.fed ? oppLast : 'Federer') + ' to react \u00b7 ' + (st.measured ? 'from serve speed' : 'estimated') : 'Time to react';
        Array.prototype.forEach.call(reactBox.querySelectorAll('rect.rb'), function (r) { r.setAttribute('opacity', +r.getAttribute('data-k') === st.k ? 1 : 0.7); r.setAttribute('stroke', +r.getAttribute('data-k') === st.k ? C.ink : r.getAttribute('fill')); r.setAttribute('stroke-width', +r.getAttribute('data-k') === st.k ? 1.2 : 0.6); });
      }
      if (!T.markedBounce && u > st.D1) {
        T.markedBounce = true; rebuildMarks(st.k, u);
      }
      // hitter ring pulse + stroke tag
      var hg = st.fed ? fedDot : oppDot, since = waiting ? -1 : u;
      hg.__ring.setAttribute('r', ((hg.__r || 8.5) + Math.min(1, Math.max(0, since) / 350) * 9).toFixed(1)); hg.__ring.setAttribute('opacity', since < 0 ? 0 : Math.max(0, 0.9 - since / 400).toFixed(2));
      hg.__tag.setAttribute('opacity', since < 0 ? 0 : since < 900 ? 1 : Math.max(0, 1 - (since - 900) / 300).toFixed(2));
      // reaction bar under the court: drains from the opponent's contact to the receiver's contact
      if (st.react && !waiting) {
        var leftS = Math.max(0, st.react - u / 1000);
        rxVal.textContent = leftS > 0 ? leftS.toFixed(2) + ' s left' : 'contact';
        rxFill.style.width = (100 * leftS / st.react).toFixed(1) + '%';
      } else if (!waiting) { rxFill.style.width = '0%'; rxVal.textContent = st.past ? 'no chance to reach it' : st.fault ? 'fault' : ''; }
      // players
      var e = ease(Math.min(1, Math.max(0, u / Math.max(tot, 1))));
      placeP(fedDot, lerp(st.from.fed, st.to.fed, e)); placeP(oppDot, lerp(st.from.opp, st.to.opp, e));
      // ball, with real height in metres
      var bp, show = true, vNow = 0;
      if (waiting) { bp = { x: st.A.x, y: st.A.y, z: 1.0 + 2.2 * Math.sin(Math.PI * Math.min(1, now / 600)) }; }
      else if (u <= tot) { bp = ballAt(st, u); vNow = bp.v; }
      else {
        var endP = st.C || st.B; bp = { x: endP.x, y: endP.y, z: st.C && !st.past ? 1.0 : 0 };
        show = st.C ? (st.past || st.unret ? u < tot + 250 : true) : u < st.D1 + 250;
      }
      var sp0 = S.proj(bp.x, bp.y, bp.z), gp0 = S.proj(bp.x, bp.y, 0), kb = sp0[2];
      layer(ball, bp.y); layer(shadow, bp.y); layer(tail, bp.y);
      var bxS = sp0[0], byS = sp0[1];
      var br = S.view === 'tv' ? Math.max(2.6, Math.min(6, 0.2 * kb)) : 4 + bp.z * 0.9;
      shadow.setAttribute('cx', gp0[0].toFixed(1)); shadow.setAttribute('cy', gp0[1].toFixed(1)); shadow.setAttribute('rx', (br * 1.05).toFixed(1)); shadow.setAttribute('ry', (S.view === 'tv' ? br * 0.42 : br * 0.5).toFixed(1));
      shadow.setAttribute('opacity', show ? Math.max(0.25, 0.8 - bp.z * 0.08).toFixed(2) : 0);
      ball.setAttribute('cx', bxS.toFixed(1)); ball.setAttribute('cy', byS.toFixed(1)); ball.setAttribute('r', br.toFixed(2)); ball.setAttribute('opacity', show ? 1 : 0);
      var ptxt = speedText(st), inFlight = !waiting && show && u <= tot + 150;
      if (ptxt && inFlight) {
        if (pillTx.textContent !== ptxt) { pillTx.textContent = ptxt; pillBg.setAttribute('width', (ptxt.length * 6.6 + 12).toFixed(0)); }
        var flip = bxS > S.W - 80;
        pill.setAttribute('transform', 'translate(' + (flip ? bxS - (ptxt.length * 6.6 + 24) : bxS + 11).toFixed(1) + ',' + (byS - 22).toFixed(1) + ')');
        pill.setAttribute('opacity', 1);
        pillTx.setAttribute('font-style', st.kind === 'serve' ? 'normal' : 'italic');
      } else pill.setAttribute('opacity', 0);
      // pace tail: the distance the ball covers in ~70 ms of real time, drawn back along its on-screen direction
      var prevB = (!waiting && u <= tot) ? ballAt(st, Math.max(0, u - 30)) : bp, pp = S.proj(prevB.x, prevB.y, prevB.z);
      var dxs = bxS - pp[0], dys = byS - pp[1], Ls = Math.sqrt(dxs * dxs + dys * dys) || 1, len = vNow * 0.07 * kb;
      Array.prototype.forEach.call(tail.childNodes, function (ln0, j) {
        var f0 = j / 4, f1 = (j + 1) / 4;
        ln0.setAttribute('x1', (bxS - dxs / Ls * len * f0).toFixed(1)); ln0.setAttribute('y1', (byS - dys / Ls * len * f0).toFixed(1));
        ln0.setAttribute('x2', (bxS - dxs / Ls * len * f1).toFixed(1)); ln0.setAttribute('y2', (byS - dys / Ls * len * f1).toFixed(1));
        ln0.setAttribute('stroke-width', (br * 1.5 - j * br * 0.32).toFixed(1)); ln0.setAttribute('opacity', show && vNow && Ls > 0.2 ? (0.55 - j * 0.13).toFixed(2) : 0);
      });
      progFill.style.width = Math.min(100, 100 * now / T.end).toFixed(1) + '%';
      if (T.done || now >= T.end) { pill.setAttribute('opacity', 0); ball.setAttribute('opacity', 0); shadow.setAttribute('opacity', 0); Array.prototype.forEach.call(tail.childNodes, function (ln) { ln.setAttribute('opacity', 0); }); }
      if (now >= T.end && !T.done) {
        T.done = true;
        capTop.textContent = (T.p.fed_won ? 'Point Federer' : 'Point ' + oppLast) + (T.P.end && T.P.end !== 'In play' ? ' · ' + T.P.end.toLowerCase() : '');
        capTop.setAttribute('font-family', 'var(--serif)'); capTop.setAttribute('font-size', 15);
        T.items.forEach(function (li2) { li2.className = 'done'; });
        var nxt = MS.series[current.i + 1];
        if (nxt) scoreboard(nxt, MS.points[current.i + 1], 'Score after the point');
        else scoreboard(null, null, m.win ? 'Game, set and match Federer' : 'Game, set and match ' + oppLast);
        setPaused(true); bPlay.textContent = '↺ Replay';
      }
    }
    bPlay.addEventListener('click', function () {
      if (!T) return;
      if (T.done) { play(current); return; }
      setPaused(!paused);
    });
    bPrev.addEventListener('click', function () { if (T) seekStep(T.curK - 1); });
    bNext.addEventListener('click', function () { if (T) seekStep(T.curK + 1); });
    function setTempo(v) { tempo = v; bReal.classList.toggle('on', v === 1); bSlow.classList.toggle('on', v !== 1); if (T && T.steps[T.curK]) counter.textContent = 'Shot ' + (T.curK + 1) + ' of ' + T.steps.length + (v === 1 ? ' · real time' : ' · 1/' + v + ' speed'); }
    bReal.addEventListener('click', function () { setTempo(1); });
    bSlow.addEventListener('click', function () { setTempo(3); });
    bTV.addEventListener('click', function () { if (view !== 'tv') setView('tv'); });
    bTop.addEventListener('click', function () { if (view !== 'overhead') setView('overhead'); });
    setView('tv');
    setTempo(1);
    renderMoments();
    var first = moments.filter(function (mo) { return mo.kind === 'Match point'; })[0] || moments[0];
    var autoT = first ? setTimeout(function () { play(first); }, 400) : null;
    wrap.__playIndex = function (i) {
      if (autoT) clearTimeout(autoT);
      var mo = moments.filter(function (x) { return x.i === i; })[0] || { i: i, kind: 'Match point', label: 'Match point', sub: '' };
      play(mo);
    };
    add(left, reactBox);
    add(right, el('p', 'note', 'Serve speeds are radar readings from the Grand Slam data, shown in mph. Rally shots have no measured speed: their speeds (shown with \u2248, in italics) are typical values for each shot type, not measurements, so every forehand of a given kind shows the same figure. Time to react is the time from one player\u2019s contact to the other\u2019s. For serves it is modelled from the radar speed and court geometry, and for rally shots it is estimated. Court positions are approximate: the chart records direction in thirds of the court, and depth only for returns.'));
    return wrap;
  }

  // ================================================================== Grand Slam tracking data (serve speed in mph, width, depth, distance)
  function slist(v) { var a = String(v || '').split('|'); return a.length > 2 ? a.slice(1, a.length - 1) : []; }
  function attachSlam(MS, c) {
    var spd = slist(c.spd), sw = slist(c.sw), sdp = slist(c.sdp), sn = slist(c.snum), dF = slist(c.dfed);
    MS.speeds = [];
    MS.series.forEach(function (p, i) {
      var P = MS.points[i];
      var k = spd[i] ? +spd[i] : 0;
      if (P) { if (k) P.speed = k; if (sw[i]) P.sw = sw[i]; if (sdp[i]) P.sdp = sdp[i]; if (dF[i]) P.dist = +dF[i]; }
      if (k && p.server === 'Federer') {
        var no = P ? P.serveNo : (sn[i] === '2' ? 2 : 1);
        MS.speeds.push({ i: i, mph: k, no: no, won: p.fed_won, p: p });
      }
    });
    return MS;
  }
  function speedStrip(MS) {
    var box = el('div', 'speed');
    add(box, el('div', 'mini-title rp-title', 'Every Federer serve, by speed'));
    var W = MOB ? 440 : 960, H = 170, ml = MOB ? 84 : 110, mr = MOB ? 28 : 24, rows = { 1: 52, 2: 112 };
    var ks = MS.speeds.map(function (s) { return s.mph; });
    var lo = Math.max(60, Math.floor(Math.min.apply(null, ks) / 10) * 10 - 10), hi = Math.min(160, Math.ceil(Math.max.apply(null, ks) / 10) * 10 + 10);
    var x = scale(lo, hi, ml, W - mr);
    var svg = svgRoot(W, H);
    for (var t = lo; t <= hi; t += 10) {
      add(svg, sv('line', { x1: x(t), x2: x(t), y1: 26, y2: H - 30, stroke: C.rule, 'stroke-dasharray': t % 50 ? '1 4' : null }));
      if (t % 20 === 0) add(svg, sv('text', { x: x(t), y: H - 12, 'text-anchor': 'middle', class: 'lbl' }, t + ' mph'));
    }
    var R = rng(5);
    [1, 2].forEach(function (no) {
      var L = MS.speeds.filter(function (s) { return s.no === no; });
      add(svg, sv('text', { x: 0, y: rows[no] + 4, class: 'lbl-serif' }, no === 1 ? 'First serve' : 'Second serve'));
      if (!L.length) return;
      var sorted = L.map(function (s) { return s.mph; }).sort(function (a, b) { return a - b; }), med = sorted[Math.floor(sorted.length / 2)];
      var wonP = L.filter(function (s) { return s.won; }).length;
      add(svg, sv('text', { x: 0, y: rows[no] + 20, class: 'lbl' }, 'median ' + med + ' \u00b7 won ' + Math.round(100 * wonP / L.length) + '%'));
      L.forEach(function (s) {
        var cdot = sv('circle', { cx: x(s.mph) + (R() - 0.5) * 3, cy: rows[no] + (R() - 0.5) * 22, r: 3.6, fill: s.won ? C.grass : C.purple, opacity: 0.75 });
        cdot.addEventListener('mousemove', function (e) { showTip(e, s.mph + ' mph \u00b7 ' + (no === 1 ? 'first' : 'second') + ' serve', ['Set ' + s.p.set_no + ', ' + s.p.fed_games + '\u2013' + s.p.opp_games + ', ' + s.p.game_score, s.won ? 'Federer won the point' : 'Federer lost the point']); });
        cdot.addEventListener('mouseleave', hideTip);
        add(svg, cdot);
      });
      add(svg, sv('line', { x1: x(med), x2: x(med), y1: rows[no] - 18, y2: rows[no] + 18, stroke: C.ink, 'stroke-width': 1.5 }));
    });
    add(box, svg);
    var lg = el('div', 'legend');
    [[C.grass, 'Federer won the point'], [C.purple, 'Federer lost the point']].forEach(function (q) { var sp = el('span'); var i = el('i'); i.style.background = q[0]; add(sp, i); sp.appendChild(document.createTextNode(q[1])); add(lg, sp); });
    add(box, lg);
    return box;
  }

  // ================================================================== match courts: where the shots landed, both players
  // Honest zones: rally shots are recorded by direction only (thirds), returns by direction and depth, serves by box and
  // direction. In-play shots are shown as shaded zones; winners and errors as dots you can click to replay the point.
  var MC_LANES = { '1': [-4.115, -1.372], '2': [-1.372, 1.372], '3': [1.372, 4.115] };          // opponent half, screen x
  var MC_DEPTH = { '7': [-6.4, -0.4], '8': [-9.4, -6.4], '9': [-11.885, -9.4] };                 // opponent half, y
  var MC_SV = { Deuce: { '4': [-4.115, -2.75], '5': [-2.75, -1.37], '6': [-1.37, 0] }, Ad: { '6': [0, 1.37], '5': [1.37, 2.75], '4': [2.75, 4.115] } };
  var MC_PAD = 14;
  var MC_DIRWORD = { '1': 'to the forehand side', '2': 'down the middle', '3': 'to the backhand side' };
  var MC_SVWORD = { '4': 'wide', '5': 'into the body', '6': 'down the T' };
  var MC_FAM = [
    { key: 'fh', name: 'Forehands', f: function (x) { return !x.isReturn && (x.letter === 'f' || x.letter === 'r'); } },
    { key: 'bh', name: 'Backhands', f: function (x) { return !x.isReturn && (x.letter === 'b' || x.letter === 's'); } },
    { key: 'rt', name: 'Returns', f: function (x) { return x.isReturn; }, depth: true },
    { key: 'nt', name: 'At the net', f: function (x) { return !x.isReturn && x.atNet; } }];

  function mcCollect(c) {
    var series = buildSeries(c), codes = chartCodes(c), out = { fed: [], opp: [], fsv: [], osv: [], series: series };
    var TV = { '0': 0, '15': 1, '30': 2, '40': 3, 'AD': 4 };
    series.forEach(function (p, i) {
      var P = parsePoint(codes[i], p.server === 'Federer');
      var sc = p.game_score.split('-'), a = p.tiebreak ? +sc[0] : TV[sc[0]], b = p.tiebreak ? +sc[1] : TV[sc[1]];
      var side = ((a || 0) + (b || 0)) % 2 === 0 ? 'Deuce' : 'Ad';
      P.serves.forEach(function (s, k) { (P.fedServing ? out.fsv : out.osv).push({ pi: i, k: k, side: side, dir: s.dir, fault: s.fault, out: s.out, no: s.no }); });
      P.shots.forEach(function (s, k) { var o = { pi: i, k: k }; for (var q in s) o[q] = s[q]; (s.hitterFed ? out.fed : out.opp).push(o); });
    });
    return out;
  }


  // One big court, one shot type at a time, a sentence that says what it shows. Federer by default; the opponent
  // and the errors are optional layers.
  function matchCourts(m, c, onPlay) {
    var D = mcCollect(c), series = D.series, n = series.length;
    var opp = m.opponent.split(' ').slice(-1)[0], lefty = c.opp_hand === 'L';
    var big = {};
    ['bp_fed', 'bp_opp', 'sp_fed', 'sp_opp', 'mp_fed', 'mp_opp'].forEach(function (k) { plist(c[k]).forEach(function (i) { big[i] = 1; }); });
    series.forEach(function (p, i) { if (p.tiebreak) big[i] = 1; });
    var nSets = n ? series[n - 1].set_no : 0;
    // ball speed: serves have radar readings (Grand Slam data, where available); rally shots only a typical pace by shot type
    var SPD = slist(c.spd).map(function (v) { return +v || 0; });
    function estMph(s) { var v = (s.isReturn ? (RETURN_PACE[s.letter] || PACE[s.letter]) : PACE[s.letter]) || 22; return Math.round(v / 0.78 * 2.23694); }
    var st = { shot: 'fh', pts: 'all', compare: false, errors: false };
    var SHOTS = [['fh', 'Forehands'], ['bh', 'Backhands'], ['rt', 'Returns'], ['nt', 'At the net'], ['sv', 'Serves']];
    var FAM = {}; MC_FAM.forEach(function (F) { FAM[F.key] = F; });

    // which side of the receiver a direction code points to (codes are written for a right-handed receiver)
    function sideOf(d, recvLefty) { if (d === '2') return 'middle'; var fh = d === '1'; if (recvLefty) fh = !fh; return fh ? 'forehand side' : 'backhand side'; }
    function dirPhrase(d, recv, recvLefty) { return d === '2' ? 'down the middle' : 'to ' + recv + '’s ' + sideOf(d, recvLefty); }
    function pc(a, b) { return b ? Math.round(100 * a / b) + '%' : '—'; }
    function keep(o) {
      if (st.pts === 'all') return true;
      if (st.pts === 'big') return !!big[o.pi];
      return series[o.pi] && 's' + series[o.pi].set_no === st.pts;
    }
    function ptLabel(pi) {
      var p = series[pi];
      return 'Set ' + p.set_no + ' · ' + p.fed_games + '–' + p.opp_games + (p.tiebreak ? ' (tiebreak)' : '') + ' · ' + p.game_score.replace('-', '–') + ' · ' + (p.server === 'Federer' ? 'Federer' : opp) + ' serving';
    }

    // ---- layout
    var wrap = el('div', 'mc');
    var top = el('div', 'mc-top');
    var tabs = el('div', 'tabs mc-tabs');
    var tabBtns = SHOTS.map(function (s) { var b = el('button', 'pill' + (s[0] === st.shot ? ' on' : ''), s[1]); add(tabs, b); return [s[0], b]; });
    var g2 = el('div', 'rp-viewbar'); add(g2, el('span', 'rp-vl', 'Points'));
    var ptBtns = [['all', 'All']]; for (var si = 1; si <= nSets; si++) ptBtns.push(['s' + si, 'Set ' + si]); ptBtns.push(['big', 'Big points']);
    var pb = ptBtns.map(function (o) { var b = el('button', 'seg' + (o[0] === 'all' ? ' on' : ''), o[1]); add(g2, b); return [o[0], b]; });
    add(top, tabs, g2); add(wrap, top);
    var body = el('div', 'mc-body');
    var courtBox = el('div', 'mc-court'), side = el('div', 'mc-side-col');
    var head = el('p', 'mc-head'), table = el('div', 'mc-table');
    var layers = el('div', 'mc-layers');
    var bCmp = el('button', 'pill', 'Compare with ' + opp), bErr = el('button', 'pill', 'Show errors');
    add(layers, bCmp, bErr);
    var key = el('div', 'mc-key');
    var how = el('p', 'mc-how', 'Shading shows how often he hit each zone; the darker the zone, the more shots went there. Click a dot to replay that point. Rally shots are charted by direction only, so their zones run the length of the court.');
    add(side, head, table, layers, key, how);
    add(body, courtBox, side); add(wrap, body);

    // ---- drawing helpers
    function dot(g, Cg, o, x, y, mirror, kind, desc, speed) {
      var X = mirror ? -x : x, Y = mirror ? -y : y, a;
      if (kind === 'win') a = { r: 5, fill: C.gold, stroke: '#5E4E22', 'stroke-width': 0.8 };
      else if (kind === 'ue') a = { r: 4.2, fill: '#C8B3E6', stroke: 'none' };
      else a = { r: 3.9, fill: 'none', stroke: '#EFE9F7', 'stroke-width': 1.6 };
      a.cx = Cg.X(X).toFixed(1); a.cy = Cg.Y(Y).toFixed(1); a.class = 'mc-dot'; a['data-k'] = 'd' + (mirror ? 'o' : 'f') + o.pi + '_' + o.k;
      var d = sv('circle', a);
      d.addEventListener('mousemove', function (e) { showTip(e, desc, (speed ? [speed] : []).concat([ptLabel(o.pi), (series[o.pi].fed_won ? 'Federer' : opp) + ' won the point', 'Click to replay this point'])); });
      d.addEventListener('mouseleave', hideTip);
      d.addEventListener('click', function () { hideTip(); onPlay(o.pi); });
      add(g, d);
    }
    function zone(g, Cg, x0, x1, y0, y1, frac, mirror, k, label, lx, ly, cls, tip) {
      var r = sv('rect', { x: Cg.X(x0).toFixed(1), y: Cg.Y(y0).toFixed(1), width: ((x1 - x0) * Cg.s).toFixed(1), height: ((y1 - y0) * Cg.s).toFixed(1),
        fill: mirror ? '#B9A5D8' : '#F7F4EC', 'fill-opacity': (0.05 + 0.55 * frac).toFixed(3), 'data-k': 'z' + (mirror ? 'o' : 'f') + k });
      r.addEventListener('mousemove', function (e) { showTip(e, tip[0], tip.slice(1)); });
      r.addEventListener('mouseleave', hideTip);
      add(g, r);
      if (label) add(g, sv('text', { x: Cg.X(lx).toFixed(1), y: Cg.Y(ly).toFixed(1), 'text-anchor': 'middle', class: cls + (mirror ? ' o' : ''), 'data-k': 'l' + (mirror ? 'o' : 'f') + k }, label));
    }

    function rallyLayer(gz, gd, Cg, F, list, mirror, R, hitter, recv, recvLefty) {
      var inn = list.filter(function (s) { return (s.outcome === 'In play' || s.outcome === 'Winner') && !s.miss && MC_LANES[s.dir]; });
      var tot = inn.length, Z = {}, mx = 0;
      inn.forEach(function (s) { var k = s.dir + (F.depth ? (MC_DEPTH[s.depth] ? s.depth : '8') : ''); var z = Z[k] || (Z[k] = { n: 0, w: 0 }); z.n++; if (s.outcome === 'Winner') z.w++; mx = Math.max(mx, z.n); });
      ['1', '2', '3'].forEach(function (d) {
        (F.depth ? ['7', '8', '9'] : ['']).forEach(function (dp) {
          var k = d + dp, z = Z[k] || { n: 0, w: 0 }, lx = MC_LANES[d], ly = F.depth ? MC_DEPTH[dp] : [-11.885, 0];
          var x0 = mirror ? -lx[1] : lx[0], x1 = mirror ? -lx[0] : lx[1], y0 = mirror ? -ly[1] : ly[0], y1 = mirror ? -ly[0] : ly[1];
          var share = tot ? z.n / tot : 0, lab = tot ? Math.round(100 * share) + '%' : '';
          var ty = F.depth ? (y0 + y1) / 2 + 0.45 : (mirror ? 3.4 : -2.6);
          var dw = { '7': 'short', '8': 'mid-court', '9': 'deep' }[dp];
          zone(gz, Cg, x0, x1, y0, y1, mx ? z.n / mx : 0, mirror, k, lab, (x0 + x1) / 2, ty, F.depth ? 'mc-lab-s' : 'mc-lab',
            [hitter + ' · ' + F.name.toLowerCase() + ' ' + dirPhrase(d, recv, recvLefty) + (dw ? ', ' + dw : ''), Math.round(100 * share) + '% of his in-court ' + F.name.toLowerCase() + ' (' + z.n + ' of ' + tot + ')', z.w + ' winner' + (z.w === 1 ? '' : 's')]);
        });
      });
      list.forEach(function (s) {
        var kind = s.outcome === 'Winner' ? 'win' : s.outcome === 'Unforced error' ? 'ue' : s.outcome === 'Forced error' ? 'fe' : null;
        if (!kind || (kind !== 'win' && !st.errors)) return;
        var p = landingXY(s, rng(1009 * (s.pi + 1) + 31 * (s.k + 1) + (mirror ? 7 : 0))), what = s.isReturn ? 'return' : (SHOT_INFO[s.letter] || SHOT_INFO.q)[0].toLowerCase();
        dot(gd, Cg, s, p.x, p.y, mirror, kind, hitter + ' ' + what + ' ' + (MC_LANES[s.dir] ? dirPhrase(s.dir, recv, recvLefty) + ' · ' : '') + s.outcome.toLowerCase() + (s.miss ? ' (' + s.miss.toLowerCase() + ')' : ''),
          'Ball speed ≈' + estMph(s) + ' mph (typical for this shot, not measured)');
      });
      return { Z: Z, tot: tot };
    }
    function serveLayer(gz, gd, Cg, list, mirror, R, hitter) {
      var out = {};
      ['Deuce', 'Ad'].forEach(function (sd) {
        var inn = list.filter(function (s) { return s.side === sd && !s.fault && MC_SV[sd][s.dir]; }), tot = inn.length, Z = {}, mx = 0;
        inn.forEach(function (s) { var z = Z[s.dir] || (Z[s.dir] = { n: 0, a: 0, sp: 0, sn: 0 }); z.n++; if (s.out === 'Ace') z.a++; if (SPD[s.pi]) { z.sp += SPD[s.pi]; z.sn++; } mx = Math.max(mx, z.n); });
        ['4', '5', '6'].forEach(function (d) {
          var z = Z[d] || { n: 0, a: 0, sp: 0, sn: 0 }, lx = MC_SV[sd][d];
          var x0 = mirror ? -lx[1] : lx[0], x1 = mirror ? -lx[0] : lx[1], y0 = mirror ? 0 : -6.4, y1 = mirror ? 6.4 : 0;
          zone(gz, Cg, x0, x1, y0, y1, mx ? z.n / mx : 0, mirror, sd + d, tot ? Math.round(100 * z.n / tot) + '%' : '', (x0 + x1) / 2, mirror ? 7.9 : -7.4, 'mc-lab-s',
            [hitter + ' · ' + sd.toLowerCase() + ' court, ' + MC_SVWORD[d], (tot ? Math.round(100 * z.n / tot) : 0) + '% of his serves in on that side (' + z.n + ' of ' + tot + ')', z.a + ' ace' + (z.a === 1 ? '' : 's')].concat(z.sn ? ['Average speed ' + Math.round(z.sp / z.sn) + ' mph (radar)'] : []));
        });
        add(gz, sv('text', { x: Cg.X((sd === 'Deuce') !== mirror ? -2.06 : 2.06).toFixed(1), y: Cg.Y(mirror ? 9.6 : -9.3).toFixed(1), 'text-anchor': 'middle', class: 'mc-boxlab' }, sd.toUpperCase()));
        out[sd] = { Z: Z, tot: tot };
      });
      list.forEach(function (s) {
        var kind = s.out === 'Ace' ? 'win' : (s.fault && s.no === 2 && st.errors) ? 'ue' : null;
        if (!kind) return;
        var R = rng(1009 * (s.pi + 1) + 31 * (s.k + 1) + (mirror ? 7 : 0));
        var lx = MC_SV[s.side][s.dir] || MC_SV[s.side]['5'], x = lx[0] + R() * (lx[1] - lx[0]), y = -(0.8 + R() * 5.0);
        if (kind === 'ue') {
          if (s.fault === 'Net') y = 0.25 + R() * 0.5; else if (s.fault === 'Long' || s.fault === 'Wide & long') y = -(6.7 + R() * 1.0);
          if (s.fault === 'Wide' || s.fault === 'Wide & long') x = (s.side === 'Deuce' ? -1 : 1) * (4.4 + R() * 0.7);
        }
        dot(gd, Cg, { pi: s.pi, k: 's' + s.k }, x, y, mirror, kind, hitter + ' ' + (kind === 'win' ? 'ace ' + (MC_SVWORD[s.dir] || '') : 'double fault' + (s.fault ? ' (' + s.fault.toLowerCase() + ')' : '')) + ' · ' + s.side.toLowerCase() + ' court',
          kind === 'win' ? (SPD[s.pi] ? 'Serve speed ' + SPD[s.pi] + ' mph (radar)' : 'Serve speed not recorded for this match') : null);
      });
      return out;
    }

    // ---- entrance: zones fill in lane by lane, percentages count up, winners and errors land from the net
    function landDots(dots, netY, delay0) {
      if (!dots.length || reducedMotion()) return;
      var spread = Math.min(700, dots.length * 22), t0 = null;
      dots.forEach(function (d, i) {
        d.__to = [+d.getAttribute('cy'), +d.getAttribute('r')]; d.__delay = (delay0 || 0) + spread * i / dots.length;
        d.setAttribute('cy', netY); d.setAttribute('r', (d.__to[1] * 0.35).toFixed(2)); d.style.opacity = '0';
      });
      function step(ts) {
        if (t0 === null) t0 = ts;
        var t = ts - t0, live = false;
        dots.forEach(function (d) {
          var g = Math.max(0, Math.min(1, (t - d.__delay) / 520)), k = 1 - Math.pow(1 - g, 3);
          if (g < 1) live = true;
          d.setAttribute('cy', (netY + (d.__to[0] - netY) * k).toFixed(1));
          d.setAttribute('r', (d.__to[1] * (0.35 + 0.65 * k)).toFixed(2));
          d.style.opacity = g >= 1 ? '' : String(Math.min(1, g * 2.5));
        });
        if (live) requestAnimationFrame(step);
      }
      requestAnimationFrame(step);
    }
    function enter(svg) {
      if (reducedMotion()) return;
      var zs = [].slice.call(svg.querySelectorAll('rect[data-k]')), labs = [].slice.call(svg.querySelectorAll('.mc-lab, .mc-lab-s'));
      zs.forEach(function (z, i) { z.__to = +z.getAttribute('fill-opacity'); z.__delay = 90 * (i % 9); z.setAttribute('fill-opacity', 0); });
      labs.forEach(function (l) { l.__txt = l.textContent; l.__job = numJob(l.__txt.replace(/\d+/, '0'), l.__txt); l.style.opacity = '0'; });
      var t0 = null;
      function step(ts) {
        if (t0 === null) t0 = ts;
        var t = ts - t0, g = Math.min(1, t / 700), k = easeIO(g);
        zs.forEach(function (z) { var q = Math.max(0, Math.min(1, (t - z.__delay) / 520)); z.setAttribute('fill-opacity', (z.__to * easeIO(q)).toFixed(3)); });
        labs.forEach(function (l) { l.style.opacity = g >= 1 ? '' : String(k); l.textContent = g >= 1 || !l.__job ? l.__txt : numAt(l.__job, k); });
        if (g < 1 || t < 700 + 90 * 8) requestAnimationFrame(step);
      }
      requestAnimationFrame(step);
      landDots([].slice.call(svg.querySelectorAll('.mc-dot')), svg.__netY, 380);
    }

    // ---- the takeaway sentence
    function topLane(res) {
      var L = { '1': 0, '2': 0, '3': 0 }, W = { '1': 0, '2': 0, '3': 0 }, w = 0;
      Object.keys(res.Z).forEach(function (k) { L[k[0]] += res.Z[k].n; W[k[0]] += res.Z[k].w; w += res.Z[k].w; });
      var best = ['1', '2', '3'].sort(function (a, b) { return L[b] - L[a]; })[0], bw = ['1', '2', '3'].sort(function (a, b) { return W[b] - W[a]; })[0];
      return { d: best, share: res.tot ? L[best] / res.tot : 0, w: w, wd: bw, wn: W[bw] };
    }
    function sentence(F, fl, ol, rf, ro) {
      var noun = F.key === 'nt' ? 'shots at the net' : F.name.toLowerCase(), t;
      if (F.key === 'sv') {
        var parts = [];
        ['Deuce', 'Ad'].forEach(function (sd) {
          var r = rf[sd]; if (!r.tot) return;
          var b = ['4', '5', '6'].sort(function (a, x) { return (r.Z[x] ? r.Z[x].n : 0) - (r.Z[a] ? r.Z[a].n : 0); })[0];
          parts.push('on the ' + sd.toLowerCase() + ' side he went ' + MC_SVWORD[b] + ' ' + pc(r.Z[b].n, r.tot) + ' of the time');
        });
        var aces = fl.filter(function (s) { return s.out === 'Ace'; }).length;
        t = parts.length ? 'Federer hit ' + (aces ? aces + ' ace' + (aces === 1 ? '' : 's') : 'no aces') + '; ' + parts.join(', and ') + '.' : 'No Federer serves in these points.';
        if (st.compare) { var oa = ol.filter(function (s) { return s.out === 'Ace'; }).length; t += ' ' + opp + ' hit ' + (oa ? oa + ' ace' + (oa === 1 ? '' : 's') : 'no aces') + '.'; }
        return t;
      }
      if (!rf.tot) return 'No Federer ' + noun + ' in these points.';
      var a = topLane(rf);
      if (F.key === 'rt') {
        var inP = fl.filter(function (s) { return (s.outcome === 'In play' || s.outcome === 'Winner') && !s.miss; }).length;
        var deep = fl.filter(function (s) { return s.depth === '9' && !s.miss; }).length;
        t = 'Federer put ' + pc(inP, fl.length) + ' of his returns in play, and ' + pc(deep, inP) + ' of those landed deep. Most went ' + dirPhrase(a.d, opp, lefty) + ' (' + pc(a.share * 100, 100) + ').';
      } else {
        t = 'Federer sent ' + pc(a.share * 100, 100) + ' of his ' + noun + ' ' + dirPhrase(a.d, opp, lefty) + '.';
        if (a.w >= 3 && a.wn) t += ' ' + (a.wn === a.w ? 'All ' + a.w + ' of his' : a.wn + ' of his ' + a.w) + ' ' + (F.key === 'nt' ? 'winners at the net' : noun.replace(/s$/, '') + ' winners') + ' went ' + dirPhrase(a.wd, opp, lefty) + '.';
        else if (a.w) t += ' He hit ' + a.w + ' winner' + (a.w === 1 ? '' : 's') + '.';
        else t += ' None of them were winners.';
      }
      if (F.key === 'nt') {
        var pts = {}; fl.forEach(function (s) { pts[s.pi] = 1; }); var ks = Object.keys(pts), won = ks.filter(function (i) { return series[i].fed_won; }).length;
        t += ' He won ' + won + ' of the ' + ks.length + ' points in which he came forward.';
      }
      if (st.compare && ro && ro.tot) { var b2 = topLane(ro); t += ' ' + opp + ' went ' + dirPhrase(b2.d, 'Federer', false) + ' most (' + pc(b2.share * 100, 100) + ').'; }
      return t;
    }

    // ---- the comparison table for the selected shot
    function rows(F, fl, ol) {
      function cnt(L, f) { return L.filter(f).length; }
      function shot(L) {
        return [String(L.length), pc(cnt(L, function (s) { return s.outcome === 'Winner'; }), L.length), pc(cnt(L, function (s) { return s.outcome === 'Unforced error'; }), L.length),
          pc(cnt(L, function (s) { return s.outcome === 'Forced error'; }), L.length)];
      }
      if (F.key === 'sv') {
        var s = function (L) {
          var f1 = L.filter(function (x) { return x.no === 1; });
          return [pc(cnt(f1, function (x) { return !x.fault; }), f1.length), String(cnt(L, function (x) { return x.out === 'Ace'; })), String(cnt(L, function (x) { return x.no === 2 && x.fault; }))];
        };
        var A = s(fl), B = s(ol);
        return [['First serves in', A[0], B[0]], ['Aces', A[1], B[1]], ['Double faults', A[2], B[2]]];
      }
      if (F.key === 'rt') {
        var r = function (L) {
          var inP = cnt(L, function (x) { return (x.outcome === 'In play' || x.outcome === 'Winner') && !x.miss; });
          return [String(L.length), pc(inP, L.length), pc(cnt(L, function (x) { return x.depth === '9' && !x.miss; }), inP), pc(cnt(L, function (x) { return x.outcome === 'Winner'; }), L.length)];
        };
        var A2 = r(fl), B2 = r(ol);
        return [['Returns', A2[0], B2[0]], ['In play', A2[1], B2[1]], ['Landed deep', A2[2], B2[2]], ['Winners', A2[3], B2[3]]];
      }
      var A3 = shot(fl), B3 = shot(ol);
      return [[F.key === 'nt' ? 'Shots at the net' : F.name, A3[0], B3[0]], ['Winners', A3[1], B3[1]], ['Unforced errors', A3[2], B3[2]], ['Forced errors', A3[3], B3[3]]];
    }
    function drawTable(R) {
      table.textContent = '';
      var t = el('table', 'mc-stats'), hr = el('tr');
      add(hr, el('th', null, ''), el('th', null, 'Federer')); if (st.compare) add(hr, el('th', null, opp));
      add(t, hr);
      R.forEach(function (r) { var tr = el('tr'); add(tr, el('td', null, r[0]), el('td', 'num', r[1])); if (st.compare) add(tr, el('td', 'num', r[2])); add(t, tr); });
      add(table, t);
    }
    function drawKey() {
      key.textContent = '';
      var items = [['dot', C.gold, st.shot === 'sv' ? 'Ace' : 'Winner']];
      if (st.errors) items.push(['dot', '#C8B3E6', st.shot === 'sv' ? 'Double fault' : 'Unforced error']);
      if (st.errors && st.shot !== 'sv') items.push(['ring', '#9C8DB2', 'Forced error']);
      if (st.compare) items.push(['zone', '#F7F4EC', 'Federer (top)'], ['zone', '#B9A5D8', opp + ' (bottom)']);
      items.forEach(function (x) { var s = el('span'); var i = el('i', 'mc-sw ' + x[0]); if (x[0] === 'ring') i.style.borderColor = x[1]; else i.style.background = x[1]; add(s, i); s.appendChild(document.createTextNode(x[2])); add(key, s); });
    }

    function build() {
      var R = rng(31), F = FAM[st.shot] || { key: 'sv', name: 'Serves' };
      courtBox.textContent = ''; courtBox.className = 'mc-court' + (st.compare ? ' full' : '');
      var Cg = drawCourt(20, !st.compare), svg = Cg.svg, gz = sv('g', {}), gd = sv('g', {});
      add(svg, gz, gd); svg.__netY = Cg.Y(0);
      var fl, ol, rf, ro = null;
      if (st.shot === 'sv') {
        fl = D.fsv.filter(keep); ol = D.osv.filter(keep);
        rf = serveLayer(gz, gd, Cg, fl, false, R, 'Federer');
        if (st.compare) serveLayer(gz, gd, Cg, ol, true, R, opp);
      } else {
        fl = D.fed.filter(keep).filter(F.f); ol = D.opp.filter(keep).filter(F.f);
        rf = rallyLayer(gz, gd, Cg, F, fl, false, R, 'Federer', opp, lefty);
        if (st.compare) ro = rallyLayer(gz, gd, Cg, F, ol, true, R, opp, 'Federer', false);
      }
      add(svg, sv('text', { x: Cg.W / 2, y: 16, 'text-anchor': 'middle', class: 'mc-edge' }, (st.shot === 'sv' ? 'Federer’s serves' : 'Federer’s shots') + ' land here ↑'));
      if (st.compare) add(svg, sv('text', { x: Cg.W / 2, y: Cg.H - 8, 'text-anchor': 'middle', class: 'mc-edge' }, opp + '’s land here ↓'));
      add(courtBox, svg);
      var nPts = 0; for (var i = 0; i < n; i++) if (keep({ pi: i })) nPts++;
      head.textContent = sentence(F, fl, ol, rf, ro);
      drawTable(rows(F, fl, ol));
      drawKey();
      how.setAttribute('data-n', nPts);
      how.textContent = 'Shading shows how often he hit each zone; the darker the zone, the more shots went there. Click a dot to replay that point. ' +
        (st.shot === 'sv' ? 'Serves are charted by box and direction (wide, body or T); the percentages above each box are shares of serves that went in on that side.'
          : st.shot === 'rt' ? 'Returns are charted by direction and depth (short, mid-court or deep).'
          : 'Rally shots are charted by direction only, so their zones run the length of the court.');
    }
    build();
    var firstSvg = courtBox.querySelector('svg');
    if (!reducedMotion()) {
      [].slice.call(firstSvg.querySelectorAll('rect[data-k], .mc-lab, .mc-lab-s, .mc-dot')).forEach(function (e) { e.style.visibility = 'hidden'; });
      whenVisible(courtBox, function () {
        [].slice.call(firstSvg.querySelectorAll('[style]')).forEach(function (e) { e.style.visibility = ''; });
        enter(firstSvg);
      });
    }
    // a redraw that keeps what stays (zones and numbers tween) and lands the new dots
    function redraw() {
      var oldK = {}; [].slice.call(courtBox.querySelectorAll('.mc-dot')).forEach(function (d) { oldK[d.getAttribute('data-k')] = 1; });
      var h0 = head.textContent;
      morph([head, courtBox, table], build, null, function (e) { return e.classList && e.classList.contains('mc-dot'); });
      var svg = courtBox.querySelector('svg');
      landDots([].slice.call(svg.querySelectorAll('.mc-dot')).filter(function (d) { return !oldK[d.getAttribute('data-k')]; }), svg.__netY, 120);
      if (head.textContent !== h0 && !numJob(h0, head.textContent)) swapIn(head, 'swap-in');
    }

    // ---- interactions: filters and shot types morph; adding or removing the opponent re-lays the court
    tabBtns.forEach(function (o) {
      o[1].addEventListener('click', function () {
        if (st.shot === o[0]) return;
        st.shot = o[0]; tabBtns.forEach(function (q) { q[1].className = 'pill' + (q[0] === st.shot ? ' on' : ''); });
        redraw();
      });
    });
    pb.forEach(function (o) {
      o[1].addEventListener('click', function () {
        if (st.pts === o[0]) return;
        st.pts = o[0]; pb.forEach(function (q) { q[1].className = 'seg' + (q[0] === st.pts ? ' on' : ''); });
        redraw();
      });
    });
    bErr.addEventListener('click', function () { st.errors = !st.errors; bErr.className = 'pill' + (st.errors ? ' on' : ''); redraw(); });
    bCmp.addEventListener('click', function () { st.compare = !st.compare; bCmp.className = 'pill' + (st.compare ? ' on' : ''); build(); enter(courtBox.querySelector('svg')); swapIn(table, 'swap-in'); swapIn(head, 'swap-in'); });
    return wrap;
  }


  // ================================================================== shared insight helpers
  function statRows(M) { return M.filter(function (m) { return m.hasPts && m.serve_pts && m.return_pts; }); }
  function svRt(L) {
    var sp = sum(L, function (m) { return m.serve_pts; }), sw = sum(L, function (m) { return m.serve_pts_won; });
    var rp = sum(L, function (m) { return m.return_pts; }), rw = sum(L, function (m) { return m.return_pts_won; });
    return { sv: sp ? sw / sp : null, rt: rp ? rw / rp : null, n: L.length };
  }
  // sets, tiebreaks and deciders from the score string (winner-first in the source)
  function setTally(M) {
    var T = { setsW: 0, setsL: 0, tbW: 0, tbL: 0, decW: 0, decL: 0, fiveW: 0, fiveL: 0 };
    M.forEach(function (m) {
      var sc = String(m.score || ''); if (/W\/O|RET|DEF|ABD/i.test(sc)) return;
      var sets = sc.split(' ').filter(function (x) { return /^\d+-\d+/.test(x); });
      sets.forEach(function (x) {
        var g = /^(\d+)-(\d+)/.exec(x), a = +g[1], b = +g[2], fed = m.win ? a > b : b > a;
        if (fed) T.setsW++; else T.setsL++;
        if ((a === 7 && b === 6) || (a === 6 && b === 7) || /\(/.test(x)) { if (fed) T.tbW++; else T.tbL++; }
      });
      var maxSets = m.best_of === 5 ? 5 : 3;
      if (sets.length === maxSets) { if (m.win) T.decW++; else T.decL++; if (maxSets === 5) { if (m.win) T.fiveW++; else T.fiveL++; } }
    });
    return T;
  }

  // ================================================================== VI. The Evolution
  function renderEvolution(M, CH) {
    var S = statRows(M);
    var sec = el('section', 'ch');
    // by age (whole years with enough matches)
    var byAge = {};
    S.forEach(function (m) { if (!m.fed_age) return; var a = Math.floor(m.fed_age); (byAge[a] || (byAge[a] = [])).push(m); });
    var ages = Object.keys(byAge).map(Number).filter(function (a) { return byAge[a].length >= 20; }).sort(function (a, b) { return a - b; });
    var A = ages.map(function (a) { var r = svRt(byAge[a]); r.age = a; return r; });
    var young = svRt(S.filter(function (m) { return m.fed_age < 21; })), old = svRt(S.filter(function (m) { return m.fed_age >= 35; }));
    var peakRt = A.slice().sort(function (a, b) { return b.rt - a.rt; })[0];
    add(sec, chapterHead('VI', 'The evolution', 'As the years went on, the serve did more and more of the work.',
      'As a teenager he won ' + pct(young.sv * 100, 100) + ' of his service points. After turning 35 he won ' + pct(old.sv * 100, 100) +
      ', higher than in his peak years. His return peaked at age ' + peakRt.age + ' (' + pct(peakRt.rt * 100, 100) + ' of return points won) and then slowly declined. The late-career Federer won by shortening points and holding serve.'));

    // panel 1: serve and return by age, two stacked panels sharing the age axis
    var W = MOB ? 440 : 960, H = 380, ml = MOB ? 42 : 60, mr = MOB ? 12 : 30, t1 = 30, h1 = 130, t2 = 220, h2 = 120;
    var x = scale(ages[0] - 0.5, ages[ages.length - 1] + 0.5, ml, W - mr);
    var ySv = scale(0.60, 0.74, t1 + h1, t1), yRt = scale(0.34, 0.44, t2 + h2, t2);
    var svg = svgRoot(W, H);
    [0.60, 0.65, 0.70].forEach(function (t) { add(svg, sv('line', { x1: ml, x2: W - mr, y1: ySv(t), y2: ySv(t), stroke: C.rule })); add(svg, sv('text', { x: ml - 8, y: ySv(t) + 4, 'text-anchor': 'end', class: 'lbl' }, Math.round(t * 100) + '%')); });
    [0.35, 0.40].forEach(function (t) { add(svg, sv('line', { x1: ml, x2: W - mr, y1: yRt(t), y2: yRt(t), stroke: C.rule })); add(svg, sv('text', { x: ml - 8, y: yRt(t) + 4, 'text-anchor': 'end', class: 'lbl' }, Math.round(t * 100) + '%')); });
    add(svg, sv('text', { x: ml, y: t1 - 10, class: 'lbl-ink' }, 'Service points won'));
    add(svg, sv('text', { x: ml, y: t2 - 10, class: 'lbl-ink' }, 'Return points won'));
    function line(acc, y, col) {
      var d = ''; A.forEach(function (r) { d += (d ? 'L' : 'M') + x(r.age).toFixed(1) + ',' + y(acc(r)).toFixed(1); });
      add(svg, sv('path', { d: d, fill: 'none', stroke: col, 'stroke-width': 2.4 }));
      A.forEach(function (r) {
        var c0 = sv('circle', { cx: x(r.age), cy: y(acc(r)), r: 3.6, fill: C.paper, stroke: col, 'stroke-width': 1.8 });
        c0.addEventListener('mousemove', function (e) { showTip(e, 'Age ' + r.age, ['Serve points won ' + pct(r.sv * 100, 100), 'Return points won ' + pct(r.rt * 100, 100), r.n + ' matches']); });
        c0.addEventListener('mouseleave', hideTip); add(svg, c0);
      });
    }
    line(function (r) { return r.sv; }, ySv, C.grass); line(function (r) { return r.rt; }, yRt, C.purple);
    A.forEach(function (r) { if (r.age % (MOB ? 4 : 2) === 0) add(svg, sv('text', { x: x(r.age), y: H - 8, 'text-anchor': 'middle', class: 'lbl' }, 'age ' + r.age)); });
    var c1 = el('div', 'chart'); add(c1, svg); add(sec, reveal(c1));
    add(sec, el('p', 'note', 'Each dot is one year of age with at least 20 matches with serve stats. The two panels use different scales because service points are won far more often than return points.'));

    // panel 2: dominance ratio by season
    var seasons = {};
    S.forEach(function (m) { (seasons[m.season] || (seasons[m.season] = [])).push(m); });
    var Ys = Object.keys(seasons).map(Number).sort(function (a, b) { return a - b; }).filter(function (y) { return seasons[y].length >= 10; });
    var D = Ys.map(function (yr) {
      var r = svRt(seasons[yr]), L = M.filter(function (m) { return m.season === yr; });
      var w = sum(L, function (m) { return m.win; }); return { y: yr, dr: r.rt / (1 - r.sv), win: w / L.length, n: L.length, w: w, l: L.length - w };
    });
    var best = D.slice().sort(function (a, b) { return b.dr - a.dr; })[0], y06 = D.filter(function (d) { return d.y === 2006; })[0];
    var p2 = el('div', 'panel');
    add(p2, el('h3', null, 'His late-career points were as dominant as his peak. The results weren’t.'));
    function span(a0, b0) {
      var L = D.filter(function (d) { return d.y >= a0 && d.y <= b0; }), Ls = S.filter(function (m) { return m.season >= a0 && m.season <= b0; }), r = svRt(Ls);
      var Lm = M.filter(function (m) { return m.season >= a0 && m.season <= b0; }), w = sum(Lm, function (m) { return m.win; });
      return { dr: r.rt / (1 - r.sv), win: w / Lm.length, w: w, l: Lm.length - w };
    }
    var peak = span(2004, 2007), late = span(2014, 2019);
    add(p2, el('p', 'sub', 'Dominance ratio is the share of return points he won divided by the share his opponents won on their return. Above 1 means he was out-pointing his opponents. From 2014 to 2019 his ratio was ' + late.dr.toFixed(2) +
      ', essentially the same as at his 2004\u201307 peak (' + peak.dr.toFixed(2) + '). Yet he won ' + pct(late.win * 100, 100, 0) + ' of his matches (' + late.w + '\u2013' + late.l + ') rather than ' + pct(peak.win * 100, 100, 0) +
      ' (' + peak.w + '\u2013' + peak.l + '). The same points bought fewer wins, which suggests the difference came on the big points. That is what the next chapter looks at.'));
    var W2 = MOB ? 440 : 960, H2 = 250, mt2 = 20, mb2 = 34;
    var x2 = scale(Ys[0] - 0.5, Ys[Ys.length - 1] + 0.5, ml, W2 - mr), y2 = scale(0.9, 1.5, H2 - mb2, mt2);
    var svg2 = svgRoot(W2, H2);
    [1.0, 1.2, 1.4].forEach(function (t) { add(svg2, sv('line', { x1: ml, x2: W2 - mr, y1: y2(t), y2: y2(t), stroke: t === 1 ? C.gold : C.rule, 'stroke-dasharray': t === 1 ? '3 3' : null })); add(svg2, sv('text', { x: ml - 8, y: y2(t) + 4, 'text-anchor': 'end', class: 'lbl' }, t.toFixed(1))); });
    var dd = ''; D.forEach(function (d) { dd += (dd ? 'L' : 'M') + x2(d.y).toFixed(1) + ',' + y2(d.dr).toFixed(1); });
    add(svg2, sv('path', { d: dd, fill: 'none', stroke: C.grass, 'stroke-width': 2.4 }));
    D.forEach(function (d) {
      var hl = d === best || d === y06;
      var c0 = sv('circle', { cx: x2(d.y), cy: y2(d.dr), r: hl ? 5.5 : 3.4, fill: hl ? C.gold : C.paper, stroke: C.grass, 'stroke-width': 1.8 });
      c0.addEventListener('mousemove', function (e) { showTip(e, String(d.y), ['Dominance ratio ' + d.dr.toFixed(2), 'Matches won ' + pct(d.win * 100, 100, 0) + ' (' + d.n + ' matches)']); });
      c0.addEventListener('mouseleave', hideTip); add(svg2, c0);
      if (hl) add(svg2, sv('text', { x: x2(d.y), y: y2(d.dr) - 12, 'text-anchor': 'middle', class: 'callout' }, d.y + ': ' + d.dr.toFixed(2) + ' · won ' + pct(d.win * 100, 100, 0)));
      if (d.y % (MOB ? 4 : 2) === 0) add(svg2, sv('text', { x: x2(d.y), y: H2 - 10, 'text-anchor': 'middle', class: 'lbl' }, String(d.y)));
    });
    add(svg2, sv('text', { x: W2 - mr, y: y2(1.0) - 6, 'text-anchor': 'end', class: 'callout-sm', fill: C.gold }, 'even with opponents'));
    var c2 = el('div', 'chart'); add(c2, svg2); add(p2, c2); add(sec, reveal(p2));

    // panel 3: style eras from the charted matches
    var CHm = CH.filter(function (c) { return c.source !== 'Slam'; });
    var mk = {}; M.forEach(function (m) { mk[m.match_key] = m; });
    var ERA = [['1998–2003', 1998, 2003, 'The all-court prodigy'], ['2004–07', 2004, 2007, 'The peak'], ['2008–13', 2008, 2013, 'The baseline years'],
      ['2014–16', 2014, 2016, 'Edberg as coach, bigger racquet'], ['2017–21', 2017, 2021, 'The renaissance']];
    var rows = ERA.map(function (e) {
      var L = CHm.filter(function (c) { var m = mk[c.match_key]; return m && m.season >= e[1] && m.season <= e[2]; });
      var f = function (k) { return sum(L, function (c) { return c[k]; }); };
      return { e: e, n: L.length, snv: f('snv_n') / f('sv_n'), snvW: f('snv_w') / f('snv_n'), net: f('net_n') / f('n_points'), netW: f('net_w') / f('net_n'),
        slice: f('bh_slice') / (f('bh_slice') + f('bh_drive')), rally: f('rally_sum') / f('n_points'), short: f('r1_3_n') / f('n_points') };
    });
    var p3 = el('div', 'panel');
    var r0 = rows[0], r2 = rows[2], r4 = rows[4];
    add(p3, el('h3', null, 'He changed his game three times, and each change worked.'));
    add(p3, el('p', 'sub', 'Serve-and-volley went from ' + pct(r0.snv * 100, 100, 0) + ' of his service points as a young player to ' + pct(r2.snv * 100, 100, 1) +
      ' at the height of the baseline era. It then came back after 2014, and in his final years he won ' + pct(r4.snvW * 100, 100, 0) + ' of those points. His average rally fell from ' + r2.rally.toFixed(1) + ' shots to ' + r4.rally.toFixed(1) + ' as he looked to finish points sooner.'));
    var tb = el('table', 'pat evo'); var hr = el('tr');
    ['Era', 'Serve & volley', 'S&V won', 'At the net', 'Net won', 'Slice share', 'Avg rally', '≤ 3 shots'].forEach(function (h) { add(hr, el('th', null, h)); });
    add(tb, hr);
    function cellBar(v, maxv, col) { var td = el('td'); var b = el('span', 'mbar'); b.style.width = Math.max(2, 56 * v / maxv) + 'px'; if (col) b.style.background = col; add(td, b); td.appendChild(document.createTextNode(pct(v * 100, 100, 0))); return td; }
    var mx = function (k) { return Math.max.apply(null, rows.map(function (r) { return r[k]; })); };
    rows.forEach(function (r) {
      var tr = el('tr'); var td = el('td'); add(td, el('span', 'pn', r.e[0]), el('span', 'pd', r.e[3] + ' · ' + r.n + ' charted matches')); add(tr, td);
      add(tr, cellBar(r.snv, mx('snv')), el('td', null, pct(r.snvW * 100, 100, 0)), cellBar(r.net, mx('net'), C.gold), el('td', null, pct(r.netW * 100, 100, 0)),
        cellBar(r.slice, mx('slice'), C.purple), el('td', null, r.rally.toFixed(1)), el('td', null, pct(r.short * 100, 100, 0)));
      add(tb, tr);
    });
    add(p3, tb);
    add(p3, el('p', 'note', 'Style figures come from the Match Charting Project. The charted matches lean towards big matches and top opponents, so they describe his game in important matches better than in a typical first-round match.'));
    add(sec, reveal(p3));
    return sec;
  }

  // ================================================================== VII. Pressure
  function renderPressure(M, CH, openMatch) {
    var S = statRows(M);
    var sec = el('section', 'ch');
    var base = svRt(S);
    var bpS = sum(S, function (m) { return m.bp_saved; }) / sum(S, function (m) { return m.bp_faced; });
    var bpC = sum(S, function (m) { return m.bp_converted; }) / sum(S, function (m) { return m.bp_chances; });
    var RIV = ['Rafael Nadal', 'Novak Djokovic', 'Andy Murray', 'Juan Martin del Potro', 'Stan Wawrinka', 'Lleyton Hewitt', 'Andy Roddick', 'Tomas Berdych'];
    var R = [{ name: 'Everyone', L: S }].concat(RIV.map(function (o) { return { name: o, L: S.filter(function (m) { return m.opponent === o; }) }; })).map(function (r) {
      var b = svRt(r.L), bf = sum(r.L, function (m) { return m.bp_faced; }), bc = sum(r.L, function (m) { return m.bp_chances; });
      r.sv = b.sv; r.rt = b.rt; r.bpS = sum(r.L, function (m) { return m.bp_saved; }) / bf; r.bpC = sum(r.L, function (m) { return m.bp_converted; }) / bc;
      r.dS = r.bpS - r.sv; r.dC = r.bpC - r.rt; r.bf = bf; r.bc = bc; return r;
    });
    var nad = R[1], djo = R[2];
    add(sec, chapterHead('VII', 'Pressure', 'Against his greatest rivals, his edge got smaller on the biggest points.',
      'Overall he saved ' + pct(bpS * 100, 100) + ' of break points against a ' + pct(base.sv * 100, 100) + ' baseline on serve, and converted ' + pct(bpC * 100, 100) + ' of his own against a ' + pct(base.rt * 100, 100) +
      ' baseline on return. Against Nadal his save rate fell ' + Math.abs(nad.dS * 100).toFixed(1) + ' points below his usual serve level, and against Djokovic ' + Math.abs(djo.dS * 100).toFixed(1) + ' points below.'));

    // panel 1: pressure delta, diverging dots
    var p1 = el('div', 'panel');
    add(p1, el('h3', null, 'The pressure gap, opponent by opponent'));
    add(p1, el('p', 'sub', 'Each dot shows how much better (right) or worse (left) he did on break points than on ordinary points against that opponent. On serve, it compares break points saved with service points won. On return, it compares break points converted with return points won.'));
    var W = MOB ? 440 : 960, rowH = 30, H = 46 + R.length * rowH, ml = MOB ? 104 : 210, mid = MOB ? 12 : 20, half = (W - ml - mid) / 2;
    var xs = scale(-0.08, 0.08, ml, ml + half), xr = scale(-0.08, 0.08, ml + half + mid, W);
    var svg = svgRoot(W, H);
    add(svg, sv('text', { x: ml + half / 2, y: 14, 'text-anchor': 'middle', class: 'lbl-ink' }, MOB ? 'Saving BPs' : 'On serve: saving break points'));
    add(svg, sv('text', { x: ml + half + mid + half / 2, y: 14, 'text-anchor': 'middle', class: 'lbl-ink' }, MOB ? 'Converting BPs' : 'On return: converting break points'));
    [xs, xr].forEach(function (xx) {
      add(svg, sv('line', { x1: xx(0), x2: xx(0), y1: 24, y2: H - 10, stroke: C.gold, 'stroke-dasharray': '3 3' }));
      [-0.05, 0.05].forEach(function (t) { add(svg, sv('text', { x: xx(t), y: H - 2, 'text-anchor': 'middle', class: 'lbl' }, (t > 0 ? '+' : '') + (t * 100) + ' pts')); });
    });
    R.forEach(function (r, i) {
      var cy = 36 + i * rowH;
      add(svg, sv('text', { x: 0, y: cy + 4, class: i === 0 ? 'lbl-serif' : 'lbl-ink' }, r.name === 'Everyone' ? (MOB ? 'All' : 'All opponents') : (MOB ? r.name.split(' ').slice(-1)[0] : r.name)));
      add(svg, sv('line', { x1: ml, x2: W, y1: cy, y2: cy, stroke: C.rule, opacity: 0.6 }));
      [[xs, r.dS, r.bpS, r.sv, r.bf, 'saved'], [xr, r.dC, r.bpC, r.rt, r.bc, 'converted']].forEach(function (q) {
        var v = Math.max(-0.08, Math.min(0.08, q[1])), col = q[1] >= 0 ? C.grass : C.purple;
        add(svg, sv('line', { x1: q[0](0), x2: q[0](v), y1: cy, y2: cy, stroke: col, 'stroke-width': 3, opacity: 0.6 }));
        var c0 = sv('circle', { cx: q[0](v), cy: cy, r: 6, fill: col });
        c0.addEventListener('mousemove', function (e) { showTip(e, r.name === 'Everyone' ? 'All opponents' : r.name, ['Break points ' + q[5] + ': ' + pct(q[2] * 100, 100) + ' of ' + fmt(q[4]), 'Usual level: ' + pct(q[3] * 100, 100), 'Difference: ' + (q[1] >= 0 ? '+' : '') + (q[1] * 100).toFixed(1) + ' points']); });
        c0.addEventListener('mouseleave', hideTip); add(svg, c0);
        // phones: the value sits beside the zero line, on the side away from the dot, so the two halves never collide
        if (MOB) add(svg, sv('text', { x: q[0](0) + (q[1] >= 0 ? -6 : 6), y: cy + 4, 'text-anchor': q[1] >= 0 ? 'end' : 'start', class: 'lbl' }, (q[1] >= 0 ? '+' : '') + (q[1] * 100).toFixed(1)));
        else add(svg, sv('text', { x: q[0](v) + (q[1] >= 0 ? 10 : -10), y: cy + 4, 'text-anchor': q[1] >= 0 ? 'start' : 'end', class: 'lbl' }, (q[1] >= 0 ? '+' : '') + (q[1] * 100).toFixed(1)));
      });
    });
    var c1 = el('div', 'chart'); add(c1, svg); add(p1, c1); add(sec, reveal(p1));

    // panel 2: the format ladder
    var T = setTally(M), W0 = sum(M, function (m) { return m.win; });
    var fin = S.filter(function (m) { return m.level === 'Grand Slam' && m.round === 'F'; });
    var finAll = M.filter(function (m) { return m.level === 'Grand Slam' && m.round === 'F'; }), finW = sum(finAll, function (m) { return m.win; });
    var L = [['Matches', W0, M.length - W0], ['Sets', T.setsW, T.setsL], ['Deciding sets', T.decW, T.decL], ['Tiebreaks', T.tbW, T.tbL], ['Five-set matches', T.fiveW, T.fiveL], ['Grand Slam finals', finW, finAll.length - finW]];
    var p2 = el('div', 'panel');
    add(p2, el('h3', null, 'The closer the contest, the more his edge shrank toward 54%'));
    add(p2, el('p', 'sub', 'Over a whole match his small edge in points compounds into a large edge in results. In a tiebreak or a deciding set there are fewer points for it to work with, so the win rate falls toward a coin flip. In ' + finAll.length + ' Grand Slam finals he won ' +
      pct(sum(fin, function (m) { return m.points_won; }), sum(fin, function (m) { return m.points_total; })) + ' of the points.'));
    var box = el('div', 'ladder');
    L.sort(function (a, b) { return b[1] / (b[1] + b[2]) - a[1] / (a[1] + a[2]); });
    L.forEach(function (r) {
      var row = el('div', 'lrow'), n = r[1] + r[2], p = n ? r[1] / n : 0;
      var nm = el('div', 'lnm', r[0]); var bar = el('div', 'lbar'); var b = el('b'); b.style.width = (100 * p).toFixed(1) + '%'; add(bar, b);
      var tick = el('i'); tick.style.left = '50%'; add(bar, tick);
      add(row, nm, bar, el('div', 'lval num', pct(p * 100, 100, 1)), el('div', 'lwl num', fmt(r[1]) + '–' + fmt(r[2])));
      add(box, row);
    });
    add(p2, box);
    add(p2, el('p', 'note', 'Sets, tiebreaks and deciding sets are counted from the match scores (walkovers and retirements excluded). A deciding set is the final possible set of a match: the third in best-of-three, the fifth in best-of-five.'));
    add(sec, reveal(p2));

    // panel 3: so close, the charted losses in which he held match point
    var mk = {}; M.forEach(function (m) { mk[m.match_key] = m; });
    var close = CH.filter(function (c) { var m = mk[c.match_key]; return m && !m.win && plist(c.mp_fed).length; }).map(function (c) { return { c: c, m: mk[c.match_key], n: plist(c.mp_fed).length }; })
      .sort(function (a, b) { return a.m.match_date < b.m.match_date ? 1 : -1; });
    var p3 = el('div', 'panel');
    add(p3, el('h3', null, 'So close: ' + close.length + ' losses in which he held match point'));
    add(p3, el('p', 'sub', 'These are the point-by-point matches in which Federer was one point from winning and still lost. Click any of them to replay his first match point.'));
    var grid = el('div', 'close-grid');
    close.forEach(function (o) {
      var b = el('button', 'close-card');
      add(b, el('span', 'cc-y', String(o.m.season) + ' · ' + o.m.tournament), el('span', 'cc-o', o.m.opponent), el('span', 'cc-s', roundName(o.m.round) + ' · ' + fedScore(o.m)),
        el('span', 'cc-n', o.n + ' match point' + (o.n > 1 ? 's' : '') + ' held'));
      b.addEventListener('click', function () { openMatch(o.m.match_key, plist(o.c.mp_fed)[0]); });
      add(grid, b);
    });
    add(p3, grid);
    add(p3, el('p', 'note', 'Only matches with point-by-point data are included (charted matches plus Grand Slam tracking data), so the real career total is higher.'));
    add(sec, reveal(p3));
    return sec;
  }

  // ================================================================== additions to existing chapters
  // Opponent tiers placed on the leverage curve (called from renderRivals).
  function tierPoints(M) {
    var S = statRows(M);
    var T = [['Top 10', 1, 10], ['11–20', 11, 20], ['21–50', 21, 50], ['51–100', 51, 100], ['100+', 101, 99999]];
    return T.map(function (t) {
      var L = S.filter(function (m) { return m.opp_rank >= t[1] && m.opp_rank <= t[2]; });
      return { lab: t[0], n: L.length, pp: 100 * sum(L, function (m) { return m.points_won; }) / sum(L, function (m) { return m.points_total; }), wp: 100 * sum(L, function (m) { return m.win; }) / L.length };
    });
  }
  // The backhand against Nadal, before and after 2017 (charted rally backhands).
  function nadalBackhand(M, CH, SH) {
    var mk = {}; M.forEach(function (m) { mk[m.match_key] = m; });
    var cm = {}; CH.forEach(function (c) { cm[c.chart_id] = mk[c.match_key]; });
    function agg(f) {
      var n = 0, w = 0, ue = 0, sl = 0;
      SH.forEach(function (r) { var m = cm[r.chart_id]; if (!m || !f(m) || r.phase !== 'Rally') return;
        if (r.shot === 'Backhand') { n += r.n; if (r.outcome === 'Winner') w += r.n; if (r.outcome === 'Unforced error') ue += r.n; }
        if (r.shot === 'Slice') sl += r.n; });
      return { n: n, w: n ? w / n : 0, ue: n ? ue / n : 0, slice: (n + sl) ? sl / (n + sl) : 0 };
    }
    return [['vs Nadal, 2004–16', agg(function (m) { return m.opponent === 'Rafael Nadal' && m.season <= 2016; })],
      ['vs Nadal, 2017–19', agg(function (m) { return m.opponent === 'Rafael Nadal' && m.season >= 2017; })],
      ['vs everyone else', agg(function (m) { return m.opponent !== 'Rafael Nadal'; })]];
  }
  function nadalPanel(M, CH, SH) {
    var D = nadalBackhand(M, CH, SH);
    var p = el('div', 'panel');
    add(p, el('h3', null, 'The backhand that finally solved Nadal'));
    add(p, el('p', 'sub', 'For years Nadal’s high, heavy forehand pinned Federer’s one-handed backhand: just ' + pct(D[0][1].w * 100, 100) + ' of his backhand drives against Nadal were winners, against ' + pct(D[2][1].w * 100, 100) +
      ' versus everyone else. From 2017, with the bigger racquet he had adopted in 2014 and a more aggressive, earlier backhand, that rose to ' + pct(D[1][1].w * 100, 100) + ', and he beat Nadal in all four of their meetings that year.'));
    var W = MOB ? 440 : 960, H = 150, ml = MOB ? 130 : 200, rowH = 38;
    var x = scale(0, 0.08, ml, W - (MOB ? 70 : 140)), svg = svgRoot(W, H);
    D.forEach(function (d, i) {
      var cy = 24 + i * rowH, col = i === 1 ? C.gold : i === 0 ? C.purple : C.grey;
      add(svg, sv('text', { x: 0, y: cy + 5, class: 'lbl-ink' }, d[0]));
      add(svg, sv('rect', { x: ml, y: cy - 9, width: Math.max(2, x(d[1].w) - ml), height: 18, rx: 3, fill: col }));
      add(svg, sv('text', { x: x(d[1].w) + 8, y: cy + 5, class: 'lbl-ink' }, pct(d[1].w * 100, 100) + ' winners · ' + fmt(d[1].n) + ' backhands · slice ' + pct(d[1].slice * 100, 100, 0)));
    });
    add(p, svg);
    add(p, el('p', 'note', 'These are rally backhand drives (slices excluded) from charted matches. The 2017–19 sample is small (' + fmt(D[1][1].n) + ' backhands), so treat the jump as indicative rather than precise.'));
    return p;
  }
  // Which matches get charted (bias note for chapter V).
  function chartBias(M, CH) {
    var keys = {}; CH.forEach(function (c) { if (c.source !== 'Slam') keys[c.match_key] = 1; });
    var all = M.length, ch = M.filter(function (m) { return keys[m.match_key]; });
    var slamAll = M.filter(function (m) { return m.level === 'Grand Slam'; }).length / all, slamCh = ch.filter(function (m) { return m.level === 'Grand Slam'; }).length / ch.length;
    var t10All = M.filter(function (m) { return m.opp_rank && m.opp_rank <= 10; }).length / all, t10Ch = ch.filter(function (m) { return m.opp_rank && m.opp_rank <= 10; }).length / ch.length;
    return 'A note on the sample: volunteers chart the matches people want to see. Grand Slams are ' + pct(slamAll * 100, 100, 0) + ' of his career matches but ' + pct(slamCh * 100, 100, 0) +
      ' of the charted ones, and top-10 opponents are ' + pct(t10All * 100, 100, 0) + ' of his career but ' + pct(t10Ch * 100, 100, 0) + ' of the charted matches. These shot patterns describe his game in big matches more than in a typical one.';
  }

  // ================================================================== sticky chapter list + touch tooltips
  function chapterNav(sections) {
    var nav = el('nav', 'chnav'), items = [];
    sections.forEach(function (s) {
      var b = el('button', 'chn', s[0]); b.title = s[1];
      b.addEventListener('click', function () { s[2].scrollIntoView({ behavior: 'smooth', block: 'start' }); });
      items.push([b, s[2]]); add(nav, b);
    });
    document.body.appendChild(nav);
    function upd() {
      var vh = window.innerHeight || 800, cur = -1;
      items.forEach(function (it, i) { if (it[1].getBoundingClientRect().top < vh * 0.4) cur = i; });
      nav.classList.toggle('show', cur >= 0);
      items.forEach(function (it, i) { it[0].classList.toggle('on', i === cur); });
    }
    window.addEventListener('scroll', upd, { passive: true }); document.addEventListener('scroll', upd, { passive: true, capture: true });
    upd();
  }
  var tipTimer = null;
  document.addEventListener('touchstart', function (e) {
    var t = e.touches && e.touches[0]; if (!t) return;
    var ev; try { ev = new MouseEvent('mousemove', { clientX: t.clientX, clientY: t.clientY, bubbles: true }); } catch (err) { return; }
    e.target.dispatchEvent(ev);
    clearTimeout(tipTimer); tipTimer = setTimeout(hideTip, 2600);
  }, { passive: true });

  // Source scores are written winner-first; restate them from Federer's side.
  function fedSets(m) {
    return String(m.score || '').split(' ').filter(function (s) { return /^\d/.test(s); }).map(function (s) {
      if (m.win) return s;
      var mt = s.match(/^(\d+)-(\d+)(\(.*\))?$/);
      return mt ? mt[2] + '-' + mt[1] + (mt[3] || '') : s;
    });
  }
  function fedScore(m) { return fedSets(m).join(' ').replace(/-/g, '–'); }

  function roundName(r) { return { F: 'Final', SF: 'Semifinal', QF: 'Quarterfinal', R16: 'Fourth round', R32: 'Third round' }[r] || r; }

  // ---------------------------------------------------------------- match story: chance to win vs points lead
  var MOM_N = 0;
  function pctTxt(v) { var x = v * 100; return ((x > 98 && x < 100) || (x > 0 && x < 2) ? x.toFixed(1) : String(Math.round(x))) + '%'; }
  function lastName(m) { return m.opponent.split(' ').slice(-1)[0]; }

  // Chance Federer wins the match before every point, from the score alone. Both players are given the same
  // chance of winning a service point (the server's share of points in this match), so the line moves only
  // with the scoreboard. imp[i] = chance if he wins point i minus chance if he loses it.
  function winProb(series, m) {
    var n = series.length, rule = series.rule || 'R1';
    var need = +(series.bestOf || m.best_of) === 5 ? 3 : 2;
    var sw = 0; series.forEach(function (q) { if ((q.server === 'Federer') === !!q.fed_won) sw++; });
    var p = Math.min(0.85, Math.max(0.5, sw / n)), memo = {};
    function pw(fs) { return fs ? p : 1 - p; }
    function game(fs, a, b) {
      if (a >= 4 && a - b >= 2) return fs ? 1 : 0;
      if (b >= 4 && b - a >= 2) return fs ? 0 : 1;
      if (a >= 3 && b >= 3 && a === b) { var d = p * p / (p * p + (1 - p) * (1 - p)); return fs ? d : 1 - d; }
      var k = 'g' + (+fs) + '_' + a + '_' + b; if (k in memo) return memo[k];
      return (memo[k] = p * game(fs, a + 1, b) + (1 - p) * game(fs, a, b + 1));
    }
    function tb(ff, f, o, T) {
      if (f >= T && f - o >= 2) return 1;
      if (o >= T && o - f >= 2) return 0;
      if (f >= T - 1 && o >= T - 1 && f === o) return 0.5; // equal serve rates: level late in a tiebreak is a coin flip
      var k = 't' + (+ff) + '_' + f + '_' + o + '_' + T; if (k in memo) return memo[k];
      var t = f + o, fedSrv = ((t + 1) >> 1) % 2 === 0 ? ff : !ff, q = pw(fedSrv);
      return (memo[k] = q * tb(ff, f + 1, o, T) + (1 - q) * tb(ff, f, o + 1, T));
    }
    function kindFor(done) {
      if (done !== 2 * need - 2) return { at: 6, T: 7, id: 'n' };
      if (rule === 'R0') return { at: 0, T: 7, id: 'a' };
      if (rule === 'RT') return { at: 12, T: 7, id: 't' };
      if (rule === 'RA') return { at: 6, T: 10, id: 'x' };
      return { at: 6, T: 7, id: 'n' };
    }
    function setp(fs, fg, og, K) {
      if (fg >= 6 && fg - og >= 2) return 1;
      if (og >= 6 && og - fg >= 2) return 0;
      if (K.at && fg === K.at + 1) return 1;
      if (K.at && og === K.at + 1) return 0;
      if (K.at && fg === K.at && og === K.at) return tb(fs, 0, 0, K.T);
      if (!K.at && fg >= 6 && fg === og) { var a = game(true, 0, 0), b = game(false, 0, 0); return a * b / (a * b + (1 - a) * (1 - b)); }
      var k = 's' + (+fs) + '_' + fg + '_' + og + '_' + K.id; if (k in memo) return memo[k];
      var g = game(fs, 0, 0);
      return (memo[k] = g * setp(!fs, fg + 1, og, K) + (1 - g) * setp(!fs, fg, og + 1, K));
    }
    function match(a, b) {
      if (a >= need) return 1; if (b >= need) return 0;
      var k = 'm' + a + '_' + b; if (k in memo) return memo[k];
      var q = setp(true, 0, 0, kindFor(a + b));
      return (memo[k] = q * match(a + 1, b) + (1 - q) * match(a, b + 1));
    }
    function at(q, fedSrv, K, A, B, F, O, FF, FS, OS) {
      var sp;
      if (q.tiebreak) sp = tb(FF, F, O, K.T);
      else {
        var gp = game(fedSrv, A, B);
        sp = gp * setp(!fedSrv, q.fed_games + 1, q.opp_games, K) + (1 - gp) * setp(!fedSrv, q.fed_games, q.opp_games + 1, K);
      }
      return sp * match(FS + 1, OS) + (1 - sp) * match(FS, OS + 1);
    }
    var W = [], imp = [], games = [], fs = 0, os = 0, a = 0, b = 0, f = 0, o = 0, ff = true, g = null;
    for (var i = 0; i < n; i++) {
      var q = series[i], fedSrv = q.server === 'Federer', prev = series[i - 1];
      var newSet = i > 0 && q.set_no !== prev.set_no;
      if (newSet) { if (prev.fed_won) fs++; else os++; }
      if (i === 0 || newSet || q.fed_games !== prev.fed_games || q.opp_games !== prev.opp_games) {
        a = b = f = o = 0; ff = fedSrv; g = { a: i, b: i, tb: !!q.tiebreak }; games.push(g);
      }
      g.b = i;
      var K = kindFor(fs + os);
      W.push(at(q, fedSrv, K, a, b, f, o, ff, fs, os));
      var win = q.tiebreak ? at(q, fedSrv, K, a, b, f + 1, o, ff, fs, os) : at(q, fedSrv, K, fedSrv ? a + 1 : a, fedSrv ? b : b + 1, f, o, ff, fs, os);
      var lose = q.tiebreak ? at(q, fedSrv, K, a, b, f, o + 1, ff, fs, os) : at(q, fedSrv, K, fedSrv ? a : a + 1, fedSrv ? b + 1 : b, f, o, ff, fs, os);
      imp.push(win - lose);
      if (q.tiebreak) { if (q.fed_won) f++; else o++; }
      else if (fedSrv === !!q.fed_won) a++; else b++;
    }
    W.push(m.win ? 1 : 0);
    games.forEach(function (gm) {
      gm.start = W[gm.a]; gm.end = W[gm.b + 1]; gm.swing = gm.end - gm.start;
      var seg = W.slice(gm.a, gm.b + 2); gm.peak = Math.max.apply(null, seg); gm.low = Math.min.apply(null, seg);
    });
    return { W: W, imp: imp, games: games, p: p };
  }

  function setTallies(series, m) {
    var sets = fedSets(m), S = [];
    series.forEach(function (p, i) {
      if (i === 0 || p.set_no !== series[i - 1].set_no) S.push({ no: p.set_no, i0: i, w: 0, l: 0, tw: 0, tl: 0, tb0: -1 });
      var s = S[S.length - 1];
      s.i1 = i + 1;
      if (p.fed_won) s.w++; else s.l++;
      if (p.tiebreak) { if (s.tb0 < 0) s.tb0 = i; if (p.fed_won) s.tw++; else s.tl++; }
    });
    var fs = 0, os = 0;
    S.forEach(function (s, k) {
      var g = String(sets[k] || '').replace(/\(.*\)/, '').split('-').map(Number);
      s.fedWon = g.length === 2 && !isNaN(g[0]) ? g[0] > g[1] : !!series[s.i1 - 1].fed_won;
      if (s.fedWon) fs++; else os++;
      s.sets = fs + '–' + os;
      s.label = sets[k] ? sets[k].replace('-', '–') : '';
      s.odd = s.w !== s.l && (s.w > s.l) !== s.fedWon;
    });
    S.fs = fs; S.os = os;
    return S;
  }

  // Shared frame: set scoreboard strip, bands, tiebreak hatching, per-set points row, result tag.
  function momFrame(series, m, S, uid, dom, endVal) {
    var F = MOB ? { W: 400, H: 420, ml: 32, mr: 12, mt: 66, mb: 112 } : { W: 960, H: 410, ml: 44, mr: 132, mt: 66, mb: 78 }, n = series.length, opp = lastName(m);
    F.x = scale(0, n, F.ml, F.W - F.mr); F.y = scale(dom[0], dom[1], F.H - F.mb, F.mt);
    var x = F.x, H = F.H, mt = F.mt, mb = F.mb, mr = F.mr;
    var svg = F.svg = svgRoot(F.W, H);
    var defs = F.defs = sv('defs', {});
    var hatch = sv('pattern', { id: uid + 'h', width: 6, height: 6, patternUnits: 'userSpaceOnUse', patternTransform: 'rotate(45)' });
    add(hatch, sv('line', { x1: 0, y1: 0, x2: 0, y2: 6, stroke: C.gold, 'stroke-width': 2, opacity: 0.35 }));
    add(defs, hatch); add(svg, defs);
    function fits(t, w) { return t.length * 6.4 + 14 < w; }
    F.setEls = [];
    S.forEach(function (s, k) {
      var x0 = x(s.i0), x1 = x(s.i1), bw = x1 - x0, cx = (x0 + x1) / 2, col = s.fedWon ? C.grass : C.purple;
      if (k % 2 === 1) add(svg, sv('rect', { x: x0, y: mt - 8, width: bw, height: H - mb - mt + 8, fill: '#F1ECDF' }));
      if (s.tb0 >= 0) add(svg, sv('rect', { x: x(s.tb0), y: mt - 8, width: x1 - x(s.tb0), height: H - mb - mt + 8, fill: 'url(#' + uid + 'h)' }));
      var E = { s: s, bar: sv('g', {}) };
      add(E.bar, sv('rect', { x: x0 + 1, y: 8, width: Math.max(0, bw - 2), height: 24, rx: 2, fill: col }));
      var who = s.fedWon ? 'Federer' : opp;
      var t = ['Set ' + s.no + ' · ' + s.label + ' · ' + who, 'Set ' + s.no + ' · ' + s.label, s.label].filter(function (o) { return fits(o, bw); })[0] || '';
      if (t) add(E.bar, sv('text', { x: cx, y: 24, 'text-anchor': 'middle', class: 'mom-set' }, t));
      add(svg, E.bar);
      E.sets = sv('text', { x: cx, y: 48, 'text-anchor': 'middle', class: 'lbl' }, (fits('sets ' + s.sets, bw) ? 'sets ' : '') + s.sets);
      E.suffix = fits(s.w + '–' + s.l + ' pts', bw) ? ' pts' : '';
      E.ptFinal = s.w + '–' + s.l + E.suffix + (s.odd ? ' ◆' : '');
      E.pt = sv('text', { x: cx, y: H - mb + 20, 'text-anchor': 'middle', class: 'lbl-ink', fill: s.odd ? C.gold : C.ink, 'font-weight': s.odd ? 600 : 400 }, E.ptFinal);
      add(svg, E.sets, E.pt);
      if (s.tb0 >= 0) { E.tb = sv('text', { x: cx, y: H - mb + 36, 'text-anchor': 'middle', class: 'lbl' }, 'TB ' + s.tw + '–' + s.tl); add(svg, E.tb); }
      F.setEls.push(E);
    });
    // ---- draw-on: the match plays out left to right at a steady pace per point
    //  - the line is clipped at the pen; the fill fades in just behind it (gradient mask)
    //  - match-point dots and turning-point markers pop in whole when the pen reaches them
    //  - the pen pauses briefly on each turning point
    //  - the set strip lights up set by set; per-set points count up; sets score and tiebreaks appear as each set ends
    //  - the result tag appears only once the line reaches the end
    //  - hover is limited to what has been drawn; nothing runs until the chart is on screen
    F.plot = sv('g', {}); F.fillG = sv('g', {}); F.lineG = sv('g', {}); add(F.plot, F.fillG, F.lineG);
    F.popG = sv('g', {}); F.pops = []; F.dwells = []; F.drawP = n;
    F.pen = sv('circle', { r: 3.6, fill: C.ink, stroke: C.paper, 'stroke-width': 1.2, opacity: 0, 'pointer-events': 'none' });
    var lineClip = sv('clipPath', { id: uid + 'lc' }), lineRect = sv('rect', { x: 0, y: 0, width: F.W, height: H });
    add(lineClip, lineRect);
    var grad = sv('linearGradient', { id: uid + 'fg', gradientUnits: 'userSpaceOnUse', x1: 0, x2: 1, y1: 0, y2: 0 });
    add(grad, sv('stop', { offset: 0, 'stop-color': '#fff' }), sv('stop', { offset: 1, 'stop-color': '#000' }));
    var fmask = sv('mask', { id: uid + 'fm', maskUnits: 'userSpaceOnUse', x: 0, y: 0, width: F.W, height: H });
    add(fmask, sv('rect', { x: 0, y: 0, width: F.W, height: H, fill: 'url(#' + uid + 'fg)' }));
    add(defs, lineClip, grad, fmask);
    var revealRun = 0;
    function popAt(o, k) {   // k: 0..1 → scale with a small overshoot
      var e = k >= 1 ? 1 : 1 + 2.2 * Math.pow(k - 1, 3) + 1.2 * Math.pow(k - 1, 2);
      o.g.setAttribute('opacity', Math.min(1, k * 1.6).toFixed(2));
      o.g.setAttribute('transform', k >= 1 ? '' : 'translate(' + o.cx + ',' + o.cy + ') scale(' + Math.max(0.01, e).toFixed(3) + ') translate(' + (-o.cx) + ',' + (-o.cy) + ')');
    }
    function setState(E, pi, done) {
      var s = E.s, cur = pi > s.i0 && pi < s.i1, fin = done || pi >= s.i1;
      E.bar.setAttribute('opacity', fin || cur ? 1 : 0.22);
      E.sets.setAttribute('opacity', fin ? 1 : 0);
      if (E.tb) E.tb.setAttribute('opacity', fin ? 1 : 0);
      if (fin) { E.pt.textContent = E.ptFinal; E.pt.setAttribute('opacity', 1); }
      else if (cur && F.cum) {
        var w = F.cum[pi] - F.cum[s.i0], l = (pi - s.i0) - w;
        E.pt.textContent = w + '–' + l + E.suffix; E.pt.setAttribute('opacity', 1);
      } else E.pt.setAttribute('opacity', 0);
    }
    function finish() {
      F.drawP = n;
      F.lineG.removeAttribute('clip-path'); F.fillG.removeAttribute('mask');
      F.pen.setAttribute('opacity', 0);
      F.pops.forEach(function (o) { popAt(o, 1); });
      F.setEls.forEach(function (E) { setState(E, n, true); });
      if (F.tagG) { F.tagG.setAttribute('opacity', 1); F.tagG.setAttribute('transform', ''); }
      if (svg.__turn) svg.__turn(-1);
    }
    function timeline() {
      var per = Math.max(1800, Math.min(3200, 1300 + n * 3.2)) / n, dwell = MOTION.dwell, segs = [], t = 0, p = 0;
      F.dwells.slice().sort(function (a, b) { return a - b; }).forEach(function (d) {
        if (d <= p || d >= n) return;
        segs.push([t, t + (d - p) * per, p, d]); t += (d - p) * per;
        segs.push([t, t + dwell, d, d]); t += dwell; p = d;
      });
      segs.push([t, t + (n - p) * per, p, n]); t += (n - p) * per;
      return { segs: segs, T: t };
    }
    function prep() {
      F.drawP = 0;
      F.lineG.setAttribute('clip-path', 'url(#' + uid + 'lc)'); F.fillG.setAttribute('mask', 'url(#' + uid + 'fm)');
      lineRect.setAttribute('width', F.ml); grad.setAttribute('x1', F.ml - 40); grad.setAttribute('x2', F.ml);
      F.pops.forEach(function (o) { o.g.setAttribute('opacity', 0); o.shown = 0; });
      F.setEls.forEach(function (E) { setState(E, 0, false); });
      if (F.tagG) F.tagG.setAttribute('opacity', 0);
      if (svg.__turn) svg.__turn(-2);
    }
    function run(id) {
      var TL = timeline(), t0 = null, j = 0, si = 0, pts = F.pts || [], endT = null;
      F.pen.setAttribute('opacity', pts.length ? 1 : 0);
      function step(ts) { try { step0(ts); } catch (err) { if (window.console) console.error('reveal', err); finish(); } }
      function step0(ts) {
        if (id !== revealRun) return;
        if (t0 === null) t0 = ts;
        var t = ts - t0;
        while (si < TL.segs.length - 1 && t > TL.segs[si][1]) si++;
        var sg = TL.segs[si], f = sg[1] > sg[0] ? Math.max(0, Math.min(1, (t - sg[0]) / (sg[1] - sg[0]))) : 1;
        var p = Math.min(n, sg[2] + (sg[3] - sg[2]) * f), dwelling = sg[2] === sg[3] && t < sg[1];
        F.drawP = p;
        var cx = x(p);
        lineRect.setAttribute('width', cx + 1);
        grad.setAttribute('x1', (cx - 34).toFixed(1)); grad.setAttribute('x2', (cx + 1).toFixed(1));
        if (pts.length) {
          while (j < pts.length - 2 && pts[j + 1][0] < cx) j++;
          var a = pts[j], b = pts[Math.min(pts.length - 1, j + 1)], q = b[0] > a[0] ? Math.max(0, Math.min(1, (cx - a[0]) / (b[0] - a[0]))) : 0;
          F.pen.setAttribute('cx', (a[0] + (b[0] - a[0]) * q).toFixed(1)); F.pen.setAttribute('cy', (a[1] + (b[1] - a[1]) * q).toFixed(1));
          F.pen.setAttribute('r', dwelling ? 5 : 3.6);
        }
        F.pops.forEach(function (o) {
          if (!o.shown && p >= o.at - 0.01) { o.shown = ts; if (o.turn !== undefined && svg.__turn) svg.__turn(o.turn); }
          if (o.shown) popAt(o, Math.min(1, (ts - o.shown) / MOTION.pop));
        });
        var pi = Math.floor(p + 1e-6);
        F.setEls.forEach(function (E) { setState(E, pi, false); });
        if (p >= n) {
          if (endT === null) { endT = ts; F.pen.setAttribute('opacity', 0); F.lineG.removeAttribute('clip-path'); F.fillG.removeAttribute('mask'); }
          var k = Math.min(1, (ts - endT) / MOTION.pop);
          if (F.tagG) { F.tagG.setAttribute('opacity', k.toFixed(2)); F.tagG.setAttribute('transform', 'translate(' + ((1 - easeIO(k)) * 14).toFixed(1) + ',0)'); }
          var popsDone = F.pops.every(function (o) { return o.shown && ts - o.shown >= MOTION.pop; });
          if (k >= 1 && popsDone) { finish(); return; }
        }
        requestAnimationFrame(step);
      }
      requestAnimationFrame(step);
    }
    // opts.now: start immediately (the chart is already on screen); otherwise wait until it scrolls into view
    svg.__reveal = function (opts) {
      var id = ++revealRun;
      if (reducedMotion()) { finish(); return; }
      prep();
      if (opts && opts.now) { run(id); return; }
      whenVisible(svg, function () { if (id === revealRun) run(id); });
    };
    F.tag = function (sub) {
      var won = !!m.win, yy = F.y(endVal);
      if (MOB) {   // phones: the result sits on its own line under the chart instead of in a right-hand margin
        var gm = F.tagG = sv('g', {}), ty2 = H - mb + 50, tw = (won ? 'Federer' : opp).length * 6.2 + 46;
        add(gm, sv('rect', { x: F.ml, y: ty2, width: tw, height: 20, rx: 2, fill: won ? C.grass : C.purple }));
        add(gm, sv('text', { x: F.ml + tw / 2, y: ty2 + 14, 'text-anchor': 'middle', class: 'mom-set mom-set-s' }, (won ? 'Federer' : opp).toUpperCase() + ' WINS'));
        var st2 = sv('text', { x: F.ml + tw + 8, y: ty2 + 14, class: 'lbl-ink' });
        sub.forEach(function (q, k) { add(st2, sv('tspan', { fill: q[1] || (k ? C.muted : C.ink) }, (k ? ' · ' : '') + q[0])); });
        add(gm, st2); add(svg, gm);
        return;
      }
      var ty = Math.max(mt + 4, Math.min(H - mb - 60, yy - 12)), tx = x(n) + 12;
      var g = F.tagG = sv('g', {});
      add(g, sv('line', { x1: x(n) + 4, x2: tx - 2, y1: yy, y2: ty + 11, stroke: C.muted, 'stroke-width': 0.8 }));
      add(g, sv('rect', { x: tx, y: ty, width: mr - 16, height: 22, rx: 2, fill: won ? C.grass : C.purple }));
      add(g, sv('text', { x: tx + (mr - 16) / 2, y: ty + 15, 'text-anchor': 'middle', class: MOB ? 'mom-set mom-set-s' : 'mom-set' }, (won ? 'Federer' : opp).toUpperCase() + ' WINS'));
      sub.forEach(function (s, k) { add(g, sv('text', { x: tx, y: ty + 40 + k * 16, class: k ? 'lbl' : 'lbl-ink', fill: s[1] || (k ? C.muted : C.ink) }, s[0])); });
      add(svg, g);
    };
    F.hover = function (fn) {
      var guide = sv('line', { y1: mt, y2: H - mb, stroke: C.ink, 'stroke-width': 0.5, opacity: 0 });
      add(svg, guide);
      var hit = sv('rect', { x: F.ml, y: mt, width: F.W - F.ml - mr, height: H - mb - mt, fill: 'transparent' });
      hit.addEventListener('mousemove', function (e) {
        var box = svg.getBoundingClientRect();
        var vx = (e.clientX - box.left) * F.W / box.width;
        var i = Math.max(0, Math.min(n - 1, Math.round((vx - F.ml) / (F.W - F.ml - mr) * n) - 1));
        if (F.drawP < n && i + 1 > F.drawP) { hideTip(); guide.setAttribute('opacity', 0); return; }
        var gx = fn.gx ? fn.gx(i) : x(i + 1);
        guide.setAttribute('x1', gx); guide.setAttribute('x2', gx); guide.setAttribute('opacity', 0.4);
        var r = fn(i); showTip(e, r[0], r[1]);
      });
      hit.addEventListener('mouseleave', function () { hideTip(); guide.setAttribute('opacity', 0); });
      if (MOB) hit.addEventListener('touchmove', function (e) { var t = e.touches && e.touches[0]; if (!t) return; hit.dispatchEvent(new MouseEvent('mousemove', { clientX: t.clientX, clientY: t.clientY, bubbles: true })); }, { passive: true });
      add(svg, hit);
    };
    F.legend = function (items) {
      if (MOB) {   // phones: the key wraps onto several lines below the chart
        var lines = [[]], len = 0;
        items.forEach(function (it, i) {
          var t = it[0];
          if (it[1] && i + 1 < items.length) { if (len + t.length + items[i + 1][0].length > 46 && lines[lines.length - 1].length) { lines.push([]); len = 0; } }
          else if (!it[1] && len + t.length > 46 && lines[lines.length - 1].length && !(i && items[i - 1][1])) { lines.push([]); len = 0; }
          lines[lines.length - 1].push(it); len += t.length;
        });
        lines.forEach(function (ln, k) {
          var leg2 = sv('text', { x: F.ml, y: H - 8 + k * 15, class: 'callout-sm' });
          ln.forEach(function (it) { add(leg2, sv('tspan', it[1] ? { fill: it[1] } : {}, it[0].replace(/\s+$/, ' '))); });
          add(svg, leg2);
        });
        svg.setAttribute('viewBox', '0 0 ' + F.W + ' ' + (H + (lines.length - 1) * 15));
        return;
      }
      var leg = sv('text', { x: F.ml, y: H - 8, class: 'callout-sm' });
      items.forEach(function (it) { if (it[1]) add(leg, sv('tspan', { fill: it[1] }, it[0])); else add(leg, sv('tspan', {}, it[0])); });
      add(svg, leg);
    };
    return F;
  }

  function mpDots(F, series, m, yOf, xOf) {
    var opp = lastName(m), fedMP = 0, oppMP = 0;
    series.forEach(function (p, i) {
      if (!(p.mp_fed || p.mp_opp)) return;
      var cx = xOf(i), cy = F.y(yOf(i));
      var c = sv('circle', { cx: cx, cy: cy, r: 5.5, fill: p.mp_fed ? C.gold : C.purple, stroke: C.paper, 'stroke-width': 1.5 });
      if (p.mp_fed) fedMP++; else oppMP++;
      c.addEventListener('mousemove', function (e) {
        showTip(e, (p.mp_fed ? 'Federer' : opp) + ' match point', ['Set ' + p.set_no + ', ' + p.fed_games + '–' + p.opp_games +
          (p.tiebreak ? ' tiebreak' : '') + ', ' + p.game_score + ' (' + p.server + ' serving)', (p.fed_won ? 'Federer' : m.opponent) + ' won the point']);
      });
      c.addEventListener('mouseleave', hideTip);
      var pg = sv('g', {}); add(pg, c); add(F.popG, pg);
      F.pops.push({ at: F.atOf(i), g: pg, cx: cx, cy: cy });
    });
    return [fedMP, oppMP];
  }
  function mpLegend(mp, m) {
    var out = [], opp = lastName(m);
    if (mp[0]) out.push(['● ', C.gold], ['Federer match point' + (mp[0] > 1 ? 's' : '') + ' (' + mp[0] + ')   ']);
    if (mp[1]) out.push(['● ', C.purple], [opp + ' match point' + (mp[1] > 1 ? 's' : '') + ' (' + mp[1] + ')   ']);
    return out;
  }

  function pointsChart(series, m, S, uid) {
    var n = series.length, vals = series.map(function (p) { return p.fed_cum_diff; });
    var lo = Math.min(0, Math.min.apply(null, vals)) - 3, hi = Math.max(0, Math.max.apply(null, vals)) + 3, last = vals[n - 1];
    var F = momFrame(series, m, S, uid, [lo, hi], last), x = F.x, y = F.y, svg = F.svg;
    var clipA = sv('clipPath', { id: uid + 'a' }); add(clipA, sv('rect', { x: 0, y: 0, width: F.W, height: y(0) }));
    var clipB = sv('clipPath', { id: uid + 'b' }); add(clipB, sv('rect', { x: 0, y: y(0), width: F.W, height: F.H }));
    add(F.defs, clipA, clipB);
    add(svg, sv('line', { x1: F.ml, x2: F.W - F.mr, y1: y(0), y2: y(0), stroke: C.ink, 'stroke-width': 0.8 }));
    add(svg, sv('text', { x: F.ml - 8, y: y(0) + 4, 'text-anchor': 'end', class: 'lbl' }, 'even'));
    var tick = Math.max(5, Math.round((hi - lo) / 4 / 5) * 5);
    for (var tk = Math.ceil(lo / tick) * tick; tk <= hi; tk += tick) {
      if (tk === 0) continue;
      add(svg, sv('text', { x: F.ml - 8, y: y(tk) + 4, 'text-anchor': 'end', class: 'lbl' }, (tk > 0 ? '+' : '') + tk));
      add(svg, sv('line', { x1: F.ml, x2: F.W - F.mr, y1: y(tk), y2: y(tk), stroke: C.rule, 'stroke-dasharray': '2 4' }));
    }
    var area = 'M' + x(0) + ',' + y(0), line = 'M' + x(0) + ',' + y(0);
    vals.forEach(function (v, i) { var px = x(i + 1).toFixed(1), py = y(v).toFixed(1); area += 'L' + px + ',' + py; line += 'L' + px + ',' + py; });
    area += 'L' + x(n) + ',' + y(0) + 'Z';
    add(svg, F.plot, F.popG);
    F.pts = [[x(0), y(0)]].concat(vals.map(function (v, i) { return [x(i + 1), y(v)]; }));
    F.atOf = function (i) { return i + 1; };
    F.cum = [0]; series.forEach(function (q, i) { F.cum.push(F.cum[i] + (q.fed_won ? 1 : 0)); });
    add(F.fillG, sv('path', { d: area, fill: C.grassSoft, opacity: 0.5, 'clip-path': 'url(#' + uid + 'a)' }));
    add(F.fillG, sv('path', { d: area, fill: C.purpleSoft, opacity: 0.6, 'clip-path': 'url(#' + uid + 'b)' }));
    add(F.lineG, sv('path', { d: line, fill: 'none', stroke: C.ink, 'stroke-width': 1.3, 'stroke-linejoin': 'round', class: 'momline' }));
    var mp = mpDots(F, series, m, function (i) { return vals[i]; }, function (i) { return x(i + 1); });
    add(svg, F.pen);
    var won = !!m.win;
    F.tag([['Federer ' + (last > 0 ? '+' : '') + last + ' on points', (last > 0) !== won && last !== 0 ? C.gold : C.ink], ['sets ' + S.fs + '–' + S.os]]);
    var leg = [['Line: total points won minus lost   ']].concat(mpLegend(mp, m));
    if (S.some(function (s) { return s.tb0 >= 0; })) leg.push(['▨ tiebreak   ']);
    if (S.some(function (s) { return s.odd; })) leg.push(['◆ ', C.gold], ['set lost despite more points']);
    F.legend(leg);
    F.hover(function (i) {
      var p = series[i];
      return ['Point ' + p.point_no + ' · Set ' + p.set_no, ['Games ' + p.fed_games + '–' + p.opp_games + (p.tiebreak ? ' (tiebreak)' : '') + ', ' + p.game_score + ', ' + p.server + ' serving',
        (p.fed_won ? 'Federer won it' : lastName(m) + ' won it') + ' · running: ' + (p.fed_cum_diff > 0 ? '+' : '') + p.fed_cum_diff]];
    });
    return svg;
  }

  function describeGame(gm, series, m, WP) {
    var opp = lastName(m), q = series[gm.a], z = series[gm.b], n = series.length;
    var setEnds = gm.b + 1 === n || series[gm.b + 1].set_no !== z.set_no, matchEnds = gm.b + 1 === n;
    var fedWonLast = !!z.fed_won, winner = fedWonLast ? 'Federer' : opp;
    var head, bits = [];
    if (gm.tb) {
      var tw = 0, tl = 0; for (var i = gm.a; i <= gm.b; i++) { if (series[i].fed_won) tw++; else tl++; }
      head = 'Set ' + q.set_no + ' tiebreak at ' + q.fed_games + '–' + q.opp_games;
      bits.push((fedWonLast ? 'Federer wins it ' : 'Federer loses it ') + tw + '–' + tl);
    } else {
      var srv = q.server === 'Federer' ? 'Federer' : opp, srvWon = (q.server === 'Federer') === fedWonLast;
      head = 'Set ' + q.set_no + ' · ' + q.fed_games + '–' + q.opp_games + ' · ' + srv + ' serving';
      var mpF = 0, mpO = 0;
      for (var j = gm.a; j <= gm.b; j++) { if (series[j].mp_fed && !series[j].fed_won) mpF++; if (series[j].mp_opp && series[j].fed_won) mpO++; }
      if (mpF) bits.push(opp + ' saves ' + (mpF > 1 ? mpF + ' match points' : 'a match point'));
      if (mpO) bits.push('Federer saves ' + (mpO > 1 ? mpO + ' match points' : 'a match point'));
      bits.push(srvWon ? srv + ' holds' : winner + ' breaks');
    }
    var tail = gm.tb ? (matchEnds ? ', and ' + winner + ' wins the match' : ', and the set') : (matchEnds ? ' to win the match' : setEnds ? ' to take the set' : '');
    var txt = bits.join(', ') + tail + '. Chance to win: ' + pctTxt(gm.start) + ' → ' + pctTxt(gm.end);
    if (gm.swing < 0 && gm.peak - Math.max(gm.start, gm.end) > 0.04) txt += ' (it was ' + pctTxt(gm.peak) + ' during the game)';
    if (gm.swing > 0 && Math.min(gm.start, gm.end) - gm.low > 0.04) txt += ' (it fell to ' + pctTxt(gm.low) + ' during the game)';
    return { head: head, txt: txt + '.' };
  }

  function wpChart(series, m, S, uid, WP, turns) {
    var n = series.length, W = WP.W;
    var F = momFrame(series, m, S, uid + 'w', [0, 1], W[n]), x = F.x, y = F.y, svg = F.svg;
    var clipA = sv('clipPath', { id: uid + 'wa' }); add(clipA, sv('rect', { x: 0, y: 0, width: F.W, height: y(0.5) }));
    var clipB = sv('clipPath', { id: uid + 'wb' }); add(clipB, sv('rect', { x: 0, y: y(0.5), width: F.W, height: F.H }));
    add(F.defs, clipA, clipB);
    [0, 0.25, 0.75, 1].forEach(function (t) {
      add(svg, sv('line', { x1: F.ml, x2: F.W - F.mr, y1: y(t), y2: y(t), stroke: C.rule, 'stroke-dasharray': '2 4' }));
      add(svg, sv('text', { x: F.ml - 8, y: y(t) + 4, 'text-anchor': 'end', class: 'lbl' }, Math.round(t * 100) + '%'));
    });
    add(svg, sv('line', { x1: F.ml, x2: F.W - F.mr, y1: y(0.5), y2: y(0.5), stroke: C.ink, 'stroke-width': 0.8 }));
    add(svg, sv('text', { x: F.ml - 8, y: y(0.5) + 4, 'text-anchor': 'end', class: 'lbl' }, '50%'));
    var area = 'M' + x(0) + ',' + y(0.5), line = '';
    W.forEach(function (v, i) { var px = x(i).toFixed(1), py = y(v).toFixed(1); area += 'L' + px + ',' + py; line += (i ? 'L' : 'M') + px + ',' + py; });
    area += 'L' + x(n) + ',' + y(0.5) + 'Z';
    add(svg, F.plot, F.popG);
    F.pts = W.map(function (v, i) { return [x(i), y(v)]; });
    F.atOf = function (i) { return i; };
    F.cum = [0]; series.forEach(function (q, i) { F.cum.push(F.cum[i] + (q.fed_won ? 1 : 0)); });
    add(F.fillG, sv('path', { d: area, fill: C.grassSoft, opacity: 0.5, 'clip-path': 'url(#' + uid + 'wa)' }));
    add(F.fillG, sv('path', { d: area, fill: C.purpleSoft, opacity: 0.6, 'clip-path': 'url(#' + uid + 'wb)' }));
    add(F.lineG, sv('path', { d: line, fill: 'none', stroke: C.ink, 'stroke-width': 1.3, 'stroke-linejoin': 'round', class: 'momline' }));
    var mp = mpDots(F, series, m, function (i) { return W[i]; }, function (i) { return x(i); });
    // numbered turning points, placed at the end of each game, nudged above (gain) or below (loss) the line
    var placed = [];
    turns.forEach(function (gm, k) {
      var px = x(gm.b + 1), py = y(gm.end), up = gm.swing > 0;
      var my = Math.max(F.mt + 10, Math.min(F.H - F.mb - 10, py + (up ? -24 : 24)));
      placed.forEach(function (pp) { if (Math.abs(pp[0] - px) < 20 && Math.abs(pp[1] - my) < 20) my += up ? -22 : 22; });
      my = Math.max(F.mt + 10, Math.min(F.H - F.mb - 10, my));
      placed.push([px, my]);
      var tg = sv('g', {});
      add(tg, sv('line', { x1: px, y1: py, x2: px, y2: my, stroke: C.muted, 'stroke-width': 0.8 }));
      add(tg, sv('circle', { cx: px, cy: my, r: 9, fill: C.paper, stroke: up ? C.grass : C.purple, 'stroke-width': 1.6 }));
      add(tg, sv('text', { x: px, y: my + 4, 'text-anchor': 'middle', class: 'mom-num' }, k + 1));
      add(F.popG, tg);
      F.pops.push({ at: gm.b + 1, g: tg, cx: px, cy: my, turn: k });
      F.dwells.push(gm.b + 1);
    });
    add(svg, F.pen);
    var fp = sum(series, function (p) { return p.fed_won; });
    F.tag([[fp + '–' + (n - fp) + ' on points'], ['sets ' + S.fs + '–' + S.os]]);
    var leg = [['Line: Federer’s chance to win, from the score   ']].concat(mpLegend(mp, m));
    if (turns.length) leg.push(['① turning points (below)']);
    F.legend(leg);
    var hov = function (i) {
      var p = series[i];
      return ['Point ' + p.point_no + ' · Set ' + p.set_no, ['Games ' + p.fed_games + '–' + p.opp_games + (p.tiebreak ? ' (tiebreak)' : '') + ', ' + p.game_score + ', ' + p.server + ' serving',
        'Chance to win before it: ' + pctTxt(W[i]) + ' · after: ' + pctTxt(W[i + 1]),
        'At stake on this point: ' + Math.round(Math.abs(WP.imp[i]) * 100) + ' percentage points',
        (p.fed_won ? 'Federer won it' : lastName(m) + ' won it')]];
    };
    hov.gx = function (i) { return x(i); };
    F.hover(hov);
    return svg;
  }

  function momentum(series, m) {
    var n = series.length, uid = 'mom' + (++MOM_N), opp = lastName(m), won = !!m.win;
    var S = setTallies(series, m), WP = winProb(series, m), W = WP.W;
    var vals = series.map(function (p) { return p.fed_cum_diff; });
    var fp = sum(series, function (p) { return p.fed_won; }), op = n - fp;
    var aheadPts = vals.filter(function (v) { return v > 0; }).length / n;
    var fav = W.slice(0, n).filter(function (v) { return v > 0.5; }).length / n;
    var N = Math.max(5, Math.round(n * 0.1));
    var order = series.map(function (p, i) { return i; }).sort(function (a, b) { return Math.abs(WP.imp[b]) - Math.abs(WP.imp[a]); });
    var topWon = order.slice(0, N).filter(function (i) { return series[i].fed_won; }).length;
    var restWon = fp - topWon;

    var wrap = el('div', 'mom');
    // 1. two scoreboards
    var duo = el('div', 'mom-duo');
    function card(q, a, b, flag) {
      var c = el('div', 'duo'); add(c, el('div', 'duo-q', q));
      var row = el('div', 'duo-row');
      [a, b].forEach(function (v, k) {
        var cell = el('div', 'duo-cell' + (k === 1 && flag ? ' hi' : ''));
        add(cell, el('div', 'v num', v[0]), el('div', 'l', v[1])); add(row, cell);
        if (!k) add(row, el('div', 'duo-vs', 'vs'));
      });
      add(c, row); return c;
    }
    add(duo, card('Who was ahead?', [Math.round(aheadPts * 100) + '%', 'of the match ahead on total points'],
      [Math.round(fav * 100) + '%', 'of the match the favourite to win, going by the score'], (aheadPts > 0.5) !== (fav > 0.5) || Math.abs(aheadPts - fav) > 0.2));
    add(duo, card('Which points did he win?', [(100 * fp / n).toFixed(1) + '%', 'of all ' + n + ' points'],
      [Math.round(100 * topWon / N) + '%', 'of the ' + N + ' biggest points (' + topWon + ' of ' + N + ')'], Math.abs(topWon / N - fp / n) >= 0.05));
    add(wrap, duo);

    // 2. view switch
    var bar = el('div', 'rp-viewbar mom-bar');
    add(bar, el('span', 'rp-vl', 'Chart'));
    var b1 = el('button', 'seg on', 'Chance to win'), b2 = el('button', 'seg', 'Points lead');
    add(bar, b1, b2); add(wrap, bar);
    var lede = el('p', 'mom-lede', 'The line is Federer’s chance to win the match at each point, going only by the score. Above 50% (green) he was the favourite; below it (purple), ' + opp + ' was. Switch to Points lead to see total points won instead.');
    add(wrap, lede);
    b1.addEventListener('click', function () { lede.style.display = ''; });
    b2.addEventListener('click', function () { lede.style.display = 'none'; });

    // 3. turning points: the four games that moved his chance the most
    var turns = WP.games.slice().sort(function (a, b) { return Math.abs(b.swing) - Math.abs(a.swing); }).slice(0, 4)
      .filter(function (g) { return Math.abs(g.swing) >= 0.05; }).sort(function (a, b) { return a.a - b.a; });
    var wpBox = el('div', 'mom-view'), ptBox = el('div', 'mom-view');
    add(wpBox, wpChart(series, m, S, uid, WP, turns));
    if (turns.length) {
      var ol = el('ol', 'mom-turns');
      turns.forEach(function (gm, k) {
        var d = describeGame(gm, series, m, WP), li = el('li');
        add(li, el('span', 'tn ' + (gm.swing > 0 ? 'up' : 'down'), k + 1));
        var body = el('div'); add(body, el('b', null, d.head), el('span', null, d.txt)); add(li, body); add(ol, li);
      });
      add(wpBox, ol);
      // the list follows the draw: each turning point lights up when its marker appears (-2 = start, -1 = done)
      var wpSvg = wpBox.querySelector('svg');
      wpSvg.__turn = function (k) {
        if (k === -2) { ol.classList.add('drawing'); Array.prototype.forEach.call(ol.children, function (li) { li.classList.remove('lit'); }); }
        else if (k === -1) ol.classList.remove('drawing');
        else if (ol.children[k]) ol.children[k].classList.add('lit');
      };
    }
    add(ptBox, pointsChart(series, m, S, uid)); ptBox.style.display = 'none';
    add(wrap, wpBox, ptBox);
    // switching charts: the line itself morphs from one measure to the other; everything else fades in
    // switching charts redraws the new one left to right, like the first load
    function swapChart(showB, hideB, on, off) {
      if (showB.style.display !== 'none') return;
      on.className = 'seg on'; off.className = 'seg';
      hideB.style.display = 'none'; showB.style.display = '';
      var svgN = showB.querySelector('svg');
      var ln = showB.querySelector('path.momline'); if (ln) { ln.style.strokeDasharray = ''; ln.style.strokeDashoffset = ''; ln.style.transition = ''; }
      if (svgN && svgN.__reveal) svgN.__reveal({ now: true });
      var list = showB.querySelector('.mom-turns'); if (list) swapIn(list, 'swap-in');
    }
    b1.addEventListener('click', function () { swapChart(wpBox, ptBox, b1, b2); });
    b2.addEventListener('click', function () { swapChart(ptBox, wpBox, b2, b1); });

    // 4. the reading, in one sentence
    var peakI = 0, lowI = 0;
    for (var i = 0; i < n; i++) { if (W[i] > W[peakI]) peakI = i; if (W[i] < W[lowI]) lowI = i; }
    function where(i) { var p = series[i]; return 'set ' + p.set_no + ' at ' + p.fed_games + '–' + p.opp_games + (p.tiebreak ? ' in the tiebreak' : ', ' + p.game_score.replace('-', '–')); }
    var tbs = S.filter(function (s) { return s.tb0 >= 0; }), tbW = tbs.filter(function (s) { return s.fedWon; }).length;
    var tw = sum(tbs, function (s) { return s.tw; }), tl = sum(tbs, function (s) { return s.tl; });
    var NUM = ['no', 'one', 'two', 'three', 'four', 'five'], txt;
    if (fp > op && !won) {
      txt = 'Federer won more points (' + fp + '–' + op + ') and was ahead on points for ' + Math.round(aheadPts * 100) + '% of the match, but the score only favoured him ' +
        Math.round(fav * 100) + '% of the time. His chance to win peaked at ' + pctTxt(W[peakI]) + ' in ' + where(peakI) + '.';
      if (tbs.length && tbW < tbs.length) txt += ' ' + opp + ' won ' + (tbW === 0 ? (tbs.length === 1 ? 'the only tiebreak' : 'all ' + NUM[tbs.length] + ' tiebreaks') : (tbs.length - tbW) + ' of ' + tbs.length + ' tiebreaks') +
        ' (' + tl + '–' + tw + ' on tiebreak points)';
      txt += (tbs.length && tbW < tbs.length ? ' and ' : ' ' + opp + ' won ') + Math.round(100 * (N - topWon) / N) + '% of the biggest points.';
    } else if (fp < op && won) {
      txt = 'Federer won fewer points (' + fp + '–' + op + ') but won ' + Math.round(100 * topWon / N) + '% of the biggest points, and the match.' +
        ' His chance to win fell as low as ' + pctTxt(W[lowI]) + ' in ' + where(lowI) + '.';
    } else if (won) {
      txt = 'Federer won ' + fp + '–' + op + ' on points and ' + Math.round(100 * topWon / N) + '% of the biggest points.' +
        (W[lowI] < 0.35 ? ' His chance to win fell as low as ' + pctTxt(W[lowI]) + ' in ' + where(lowI) + ' before he turned it around.' : ' The score favoured him ' + Math.round(fav * 100) + '% of the time.');
    } else {
      txt = opp + ' won ' + op + '–' + fp + ' on points, and Federer won ' + Math.round(100 * topWon / N) + '% of the biggest points.' +
        (W[peakI] > 0.65 ? ' His chance to win peaked at ' + pctTxt(W[peakI]) + ' in ' + where(peakI) + '.' : '');
    }
    var odd = S.filter(function (s) { return s.odd; });
    if (odd.length) txt += ' ' + odd.map(function (s) { return 'Set ' + s.no; }).join(' and ') + ' went to the player who won fewer points in ' + (odd.length > 1 ? 'them' : 'it') + '.';
    add(wrap, el('p', 'momnote' + ((fp > op) !== won && fp !== op ? ' twist' : ''), txt));
    add(wrap, el('p', 'note mom-model', 'How to read the chance to win: it is worked out from the score alone, giving both players the same chance of winning a point on serve (' +
      Math.round(WP.p * 100) + '%, the server’s share of points in this match). It shows how close each player was to winning at every moment; it is a simplified model, not a betting line. ' +
      'The biggest points are the ' + N + ' (10%) where winning or losing the point moved that chance the most.'));
    return wrap;
  }

  function setBars(m) {
    var sets = fedSets(m);
    var W = MOB ? 440 : 960, H = 170, ml = MOB ? 80 : 110, mr = MOB ? 14 : 30;
    var svg = svgRoot(W, H);
    var bw = (W - ml - mr) / sets.length;
    add(svg, sv('text', { x: 0, y: 58, class: 'lbl-serif' }, 'Federer'));
    add(svg, sv('text', { x: 0, y: 118, class: 'lbl-serif' }, m.opponent.split(' ').slice(-1)[0]));
    sets.forEach(function (s, i) {
      var g = s.replace(/\(.*\)/, '').split('-').map(Number);
      var cx = ml + bw * i + bw / 2;
      var fw = g[0] > g[1];
      add(svg, sv('text', { x: cx, y: 18, 'text-anchor': 'middle', class: 'lbl' }, 'Set ' + (i + 1)));
      add(svg, sv('rect', { x: cx - 26, y: 34, width: 52, height: 36, rx: 3, fill: fw ? C.grass : C.paper, stroke: C.rule }));
      add(svg, sv('text', { x: cx, y: 59, 'text-anchor': 'middle', class: 'lbl-serif', fill: fw ? '#fff' : C.ink }, g[0]));
      add(svg, sv('rect', { x: cx - 26, y: 94, width: 52, height: 36, rx: 3, fill: fw ? C.paper : C.purple, stroke: C.rule }));
      add(svg, sv('text', { x: cx, y: 119, 'text-anchor': 'middle', class: 'lbl-serif', fill: fw ? C.ink : '#fff' }, g[1]));
    });
    add(svg, sv('text', { x: ml, y: H - 8, class: 'callout-sm' }, 'Point-by-point data isn’t available for this final, so the sets are shown instead.'));
    return svg;
  }

  // VI. Epilogue
  function renderEpilogue(M) {
    var stat = M.filter(function (m) { return m.hasPts; });
    var pw = sum(stat, function (m) { return m.points_won; }), pt = sum(stat, function (m) { return m.points_total; });
    var sec = el('section', 'epi');
    add(sec, el('span', 'eyebrow', 'IX · Epilogue'));
    add(sec, el('h2', null, 'He lost ' + fmt(pt - pw) + ' points. He kept playing the next one.'));
    var nSeasons = M[M.length - 1].season - M[0].season + 1;
    add(sec, el('p', 'lede', 'Across ' + nSeasons + ' seasons, in the ' + fmt(stat.length) + ' matches with full stats, he won ' + fmt(pw) + ' points and lost ' + fmt(pt - pw) +
      '. The career came from that small edge, applied again and again, and from being able to let a lost point go.'));
    var split = el('div', 'split');
    var w = el('div', 'w num', fmt(pw)), l = el('div', 'l num', fmt(pt - pw));
    w.style.flex = String(pw); l.style.flex = String(pt - pw);
    add(split, w, l);
    add(sec, reveal(split));
    var cap = el('div', 'split-cap');
    add(cap, el('span', null, 'Points won · ' + pct(pw, pt, 1)), el('span', null, 'Points lost · ' + pct(pt - pw, pt, 1)));
    add(sec, cap);
    var q = el('blockquote', 'quote');
    q.style.margin = '56px auto 0'; q.style.textAlign = 'left';
    q.textContent = '“In other words: Even top-ranked tennis players win barely over half the points they play.”';
    add(q, el('span', 'who', 'Roger Federer'));
    add(sec, q);
    return sec;
  }

  function renderFooter(M) {
    var f = el('footer');
    add(f, el('div', null, 'Match data: TennisMyLife ATP match database (tour-level singles, ' + fmt(M.length) +
      ' matches, walkovers excluded). Point-by-point data: Tennis Abstract Match Charting Project by Jeff Sackmann and contributors, CC BY-NC-SA 4.0.'));
    add(f, el('div', null, 'This is a fixed snapshot of the data embedded in the workbook (Federer retired in 2022, so it is complete), queried through the Tableau Extensions API. Made for non-commercial, educational use. Not affiliated with Roger Federer or any of his sponsors.'));
    return f;
  }

  // ------------------------------------------------------------------ boot
  function boot() {
    if (!root) return;
    root.textContent = '';
    MOB = detectMobile(); document.documentElement.classList.toggle('is-mobile', MOB);
    LOADER = makeLoader();   // the rally plays while the data loads
    if (!window.tableau || !tableau.extensions) { renderError('Extensions API not loaded'); return; }
    // draw (or redraw) the whole story; a redraw happens when the app crosses the phone breakpoint
    function renderAll(res) {
      var oldNav = document.querySelector('.chnav'); if (oldNav && oldNav.parentNode) oldNav.parentNode.removeChild(oldNav);
      bands.length = 0;
      var M = res[0], CH = res[1], SH = res[2], SV = res[3];
        if (!M.length) throw new Error('Federer Matches returned no rows');
        root.textContent = '';
        var hero = renderHero(M), arc = renderArc(M), margin = renderMargin(M), riv = renderRivals(M, CH, SH), surf = renderSurfaces(M);
        var shots = renderShots(M, CH, SH, SV), evo = renderEvolution(M, CH), exp = renderExplorer(M, CH, SH, SV), epi = renderEpilogue(M);
        var pres = renderPressure(M, CH, function (mk, i) { exp.__open(mk, i); });
        var ctx = { M: M, CH: CH, SH: SH };
        var order = [['hero', hero], ['arc', arc], ['margin', margin], ['rivals', riv], ['surf', surf], ['shots', shots], ['evo', evo], ['pres', pres], ['exp', exp]];
        var after = {};
        (window.FED_PHOTOS || []).filter(function (ph) { return !ph.hero; }).forEach(function (ph) { (after[ph.after || 'exp'] || (after[ph.after || 'exp'] = [])).push(interlude(ph, ctx)); });
        var seqNodes = [];
        order.forEach(function (o) { seqNodes.push(o[1]); (after[o[0]] || []).forEach(function (bnd) { seqNodes.push(bnd); }); });
        seqNodes.push(epi, renderFooter(M));
        seqNodes.forEach(function (nd) { add(root, nd); });
        parallax();
        chapterNav([['54%', 'The opening', hero], ['I', 'The arc', arc], ['II', 'The margin', margin], ['III', 'Leverage', riv], ['IV', 'Surfaces', surf], ['V', 'Shot by shot', shots],
          ['VI', 'The evolution', evo], ['VII', 'Pressure', pres], ['VIII', 'The matches', exp], ['IX', 'Epilogue', epi]]);
    }
    var LASTRES = null, rzT = null;
    window.addEventListener('resize', function () {
      clearTimeout(rzT);
      rzT = setTimeout(function () {
        var m = detectMobile(); if (m === MOB || !LASTRES) return;
        MOB = m; document.documentElement.classList.toggle('is-mobile', MOB);
        var y = window.scrollY; renderAll(LASTRES); window.scrollTo(0, y);
      }, 300);
    });
    tableau.extensions.initializeAsync().then(function () {
      return getDataSources();
    }).then(function (list) {
      var dsM = byName(list, 'Federer Matches'), dsC = byName(list, 'Federer Charted');
      var dsS = byName(list, 'Federer Shots'), dsV = byName(list, 'Federer Serves');
      function sumN(rows) { var t = 0; (rows || []).forEach(function (x) { t += num(x.n) || 0; }); return t; }
      return Promise.all([LOADER.track(0, queryRows(dsM, MATCH_DIMS, MATCH_MEAS)), LOADER.track(1, queryRows(dsC, CH_DIMS, CH_MEAS))]).then(function (r) {
        var CH = normCharted(r[1]);
        var maxId = Math.max.apply(null, CH.map(function (c) { return c.chart_id; }).concat([1]));
        return Promise.all([LOADER.track(2, queryChunked(dsS, SH_DIMS, SH_MEAS, maxId, 120), sumN), LOADER.track(3, queryChunked(dsV, SV_DIMS, SV_MEAS, maxId, 240), sumN)]).then(function (r2) {
          return [normMatches(r[0]), CH, normShots(r2[0]), normServes(r2[1])];
        });
      });
    }).then(function (res) {
      LASTRES = res; renderAll(res);
      LOADER.finish();
    }).catch(function (err) {
      if (LOADER) LOADER.fail();
      renderError((err && (err.message || err.errorCode)) || String(err));
    });
  }

  boot();
})();
