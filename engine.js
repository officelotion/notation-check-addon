/*
 * 表記チェック Word版 ― 中核処理 (engine.js)  v0.2
 * Wordに依存しない純粋な処理。アドインでもNodeのテストでも同じものを使う。
 *
 * 段落モデル: { runs: [ { text: "…", boxed: true|false } ] }
 *   boxed = 文字の枠(罫線)またはグレーの網かけがついている
 *
 * v0.2の変更
 *  - 変更の記録を「元の文字の位置」で持つ(pieces)。Wordへは、元の段落の文字列だけを基準に適用できる
 *    (変更履歴の「削除された文字」が残っていても、探す位置がずれない)
 *  - 記号のない電話番号に【電】を付ける
 *  - 枠・網かけの判定を、見えるものだけにした(isBoxedProps)
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.NotationEngine = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // 凡例の元の文字 → 変換後
  var LABEL_MAP = {
    '対': '対', '時': '時', '場': '場', '師': '師', '費': '費', '定': '定', '保': '保',
    '持': '持', '申': '申', '問': '問', '講': '師',
    '℡': '電', '☎': '電',
    'Ｆ': 'FAX', 'F': 'FAX', 'fax': 'FAX', 'FAX': 'FAX', 'ＦＡＸ': 'FAX',
    'Ｅ': 'e', 'E': 'e'
  };

  // 記号のない電話番号: 03（1234）1234 / （123）1234 の形。直前がすでに【電】【FAX】【e】、またはFAXの語のときは付けない
  var TEL_RE = /(?<!【電】|【FAX】|【e】|FAX|ＦＡＸ|ファクス|ファックス|[0-9])(?:[0-9]{2,4})?（[0-9]{2,5}）[0-9]{3,4}(?![0-9])/g;

  function joined(p) { return p.runs.map(function (r) { return r.text; }).join(''); }

  // str が、text の pos より前に何番目か(重ならない数え方。Wordの検索と同じ)
  function countBefore(text, str, pos) {
    var n = 0, i = text.indexOf(str);
    while (i !== -1 && i < pos) { n++; i = text.indexOf(str, i + str.length); }
    return n;
  }
  function nthIndex(text, str, n) {
    var i = text.indexOf(str);
    while (i !== -1 && n > 0) { n--; i = text.indexOf(str, i + str.length); }
    return i;
  }

  // 枠・網かけが「見える」ものだけを、凡例の目印にする
  //  bdr = {val} または null / shd = {val, fill} または null
  function isBoxedProps(bdr, shd, hl) {
    if (hl && /^(lightGray|darkGray)$/i.test(hl)) return true;   // 蛍光ペンのグレー
    if (bdr && !/^(nil|none)$/i.test(bdr.val || '')) return true;
    if (shd) {
      var v = shd.val || '', f = shd.fill || '';
      if (v && !/^(nil|clear)$/i.test(v)) return true;
      if (f && !/^(auto|FFFFFF)$/i.test(f)) return true;
    }
    return false;
  }

  // ---- 元の文字の位置を覚える(pieces) ----
  // pieces: 段落の現在の文字を、元の文字の範囲[os,oe)と対応づけた断片の並び。edited=trueは、いずれかのルールで変わった断片
  function initPieces(p) {
    var t = joined(p);
    p.orig = t;
    p.pieces = [{ text: t, os: 0, oe: t.length, edited: false, box: false }];
    // 「記号(絵文字)」のように、Wordの中で特別な形(symEx)になっている文字の位置。Wordの検索では見つけられない
    p.opaque = {};
    var pos0 = 0;
    p.runs.forEach(function (r) { if (r.sym) r.sym.forEach(function (o) { p.opaque[pos0 + o] = true; }); pos0 += r.text.length; });
  }
  function pieceEdit(p, s, e, t, box) {
    var before = [], after = [], mid = [], pre = '', post = '', pos = 0;
    p.pieces.forEach(function (pc) {
      var ps = pos, pe = pos + pc.text.length; pos = pe;
      if (pe <= s) { before.push(pc); return; }
      if (ps >= e) { after.push(pc); return; }
      if (!pc.edited) { // 変えていない断片は、必要な分だけ切り出す
        var a = Math.max(s, ps) - ps, b = Math.min(e, pe) - ps;
        if (a > 0) before.push({ text: pc.text.slice(0, a), os: pc.os, oe: pc.os + a, edited: false, box: false });
        mid.push({ os: pc.os + a, oe: pc.os + b, box: false });
        if (b < pc.text.length) after.push({ text: pc.text.slice(b), os: pc.os + b, oe: pc.oe, edited: false, box: false });
      } else { // すでに変えた断片は、ひとまとまりとして取り込む
        if (ps < s) pre = pc.text.slice(0, s - ps);
        if (pe > e) post = pc.text.slice(e - ps);
        mid.push({ os: pc.os, oe: pc.oe, box: pc.box });
      }
    });
    if (!mid.length) return;
    var os = Math.min.apply(null, mid.map(function (m) { return m.os; }));
    var oe = Math.max.apply(null, mid.map(function (m) { return m.oe; }));
    var bx = !!box || mid.some(function (m) { return m.box; });
    p.pieces = before.concat([{ text: pre + t + post, os: os, oe: oe, edited: true, box: bx }], after);
  }

  // 段落の中の正規表現の一致を、書式(run)を保ったまま置き換える
  function editPara(p, re, fn) {
    var text = joined(p), edits = [], m;
    re.lastIndex = 0;
    while ((m = re.exec(text)) !== null) {
      if (m[0].length === 0) { re.lastIndex++; continue; }
      edits.push({ s: m.index, e: m.index + m[0].length, t: fn(m), before: m[0], occ: countBefore(text, m[0], m.index) });
    }
    // 位置のずれを避けて後ろから適用し、記録は先頭からの順
    for (var k = edits.length - 1; k >= 0; k--) {
      applyEdit(p, edits[k]);
      pieceEdit(p, edits[k].s, edits[k].e, edits[k].t, false);
    }
    return edits;
  }

  function editAt(p, s, e, t) {
    var text = joined(p), before = text.slice(s, e), ed = { s: s, e: e, t: t, before: before, occ: countBefore(text, before, s) };
    applyEdit(p, ed);
    pieceEdit(p, s, e, t, false);
    return ed;
  }

  function applyEdit(p, ed) {
    var pos = 0, first = true;
    for (var i = 0; i < p.runs.length; i++) {
      var r = p.runs[i], rs = pos, re = pos + r.text.length;
      pos = re;
      var a = Math.max(ed.s, rs), b = Math.min(ed.e, re);
      if (a >= b) continue;
      r.text = r.text.slice(0, a - rs) + (first ? ed.t : '') + r.text.slice(b - rs);
      first = false;
    }
  }

  function toHalfDigits(s) { return s.replace(/[０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); }); }

  /* 自動修正 (1段落)  戻り値: 変更の一覧(ルールごとの記録) */
  function fixParagraph(p, opts) {
    var changes = [];
    initPieces(p);
    function log(rule, edits) { edits.forEach(function (e) { changes.push({ rule: rule, before: e.before, after: e.t, occ: e.occ }); }); }

    // 0) 空白1: 本文先頭の全角スペースを削除(小見出しは対象外。fixAllが opts.stripLead を渡す)
    if (opts && opts.stripLead) log('空白1', editPara(p, /^\u3000+(?=[^\u3000\s])/g, function () { return ''; }));

    // 1) 凡例: 枠・網かけつきの1文字 → 【●】(書式は外す)
    var pos = 0, shift = 0, snapshot = joined(p);
    p.runs.forEach(function (r) {
      var start = pos; pos += r.text.length;
      if (!r.boxed) return;
      var key = r.text.trim();
      if (Object.prototype.hasOwnProperty.call(LABEL_MAP, key)) {
        var after = '【' + LABEL_MAP[key] + '】', old = r.text;
        changes.push({ rule: '凡例', before: old, after: after, occ: countBefore(snapshot, old, start), box: true });
        pieceEdit(p, start + shift, start + shift + old.length, after, true);
        shift += after.length - old.length;
        r.text = after; r.boxed = false;
      }
    });
    // 2) 凡例: 書式なしの ℡ ☎ → 【電】
    log('凡例(電)', editPara(p, /[℡☎]/g, function () { return '【電】'; }));
    // 2a) 凡例1c: 半角の角カッコ [時] → 【時】、行頭のコロン形式 時： → 【時】(コロンも消す)
    log('凡例1c([ ])', editPara(p, /[\[［]([対時場師費定保持申問電])[\]］]/g, function (m) { return '【' + m[1] + '】'; }));
    log('凡例1c(コロン)', editPara(p, /(?:^|(?<=\v))([対時場師費定保持申問])[：:]/g, function (m) { return '【' + m[1] + '】'; }));
    // 2b) 凡例1b: 書式のない行頭の「問」(あとに内線・電話・コロンが続くとき) → 【問】
    log('凡例1b(問)', editPara(p, /^問(?![題い])(?=[：:]|[^\s、。】]{1,40}?(?:内線|【電】|電話|[0-9０-９]{2,4}[-－ー（(][0-9０-９]))/g, function () { return '【問】'; }));
    // 2b2) 凡例1b: 書式のない凡例(時+日付を起点に、場→師→費→定→保→持→申→問の順) → 【●】、「F＋電話番号」→【FAX】
    var plain = findPlainLabels(p);
    if (plain.length) {
      var ptext = joined(p), pedits = plain.map(function (h) {
        var b = ptext.charAt(h.index), t = '【' + h.label + '】';
        return { s: h.index, e: h.index + 1, t: t, before: b, occ: countBefore(ptext, b, h.index) };
      }).sort(function (a, b) { return a.s - b.s; });
      for (var k = pedits.length - 1; k >= 0; k--) {
        applyEdit(p, pedits[k]);
        pieceEdit(p, pedits[k].s, pedits[k].e, pedits[k].t, false);
      }
      pedits.forEach(function (e) { changes.push({ rule: '凡例1b', before: e.before, after: e.t, occ: e.occ }); });
    }
    // 2c) ●2: ●小見出しのあとの文章では、【凡例】の直前の「。」を削除
    if (/^●/.test(joined(p))) log('●2', editPara(p, /。(?=【[対時場師費定保持申問電]】)/g, function () { return ''; }));
    // 2c2) ▽月日 → ▷月日(担当スタッフの確認で決定)
    log('▽', editPara(p, /▽(?=[0-9０-９]+月)/g, function () { return '▷'; }));
    // 2d) ●1: ●小見出しと、そのあとの▶の間に、全角スペースを入れる
    if (/^●/.test(joined(p))) log('●1', editPara(p, /(?<=^●[^▶\u3000\s]{1,40})▶/g, function () { return '\u3000▶'; }));
    // 3) M01 全角数字 → 半角数字
    log('M01', editPara(p, /[０-９]+/g, function (m) { return toHalfDigits(m[0]); }));
    // 4) M03 半角カッコ → 全角カッコ
    log('M03', editPara(p, /[()]/g, function (m) { return m[0] === '(' ? '（' : '）'; }));
    // 5) 「（代表）」を削除
    log('13(代表)', editPara(p, /（代表）/g, function () { return ''; }));
    // 6) 「（内線123）」 → 「・内線123」
    log('内線(括弧)', editPara(p, /（内線([0-9]+)）/g, function (m) { return '・内線' + m[1]; }));
    // 7) ルール05: 「内線」の前に全角「・」(直前が漢字・かな・英字のとき)
    log('05', editPara(p, /(?<=[\u3400-\u9FFF\u3040-\u30FA\u30FC-\u30FFA-Za-z])内線/g, function () { return '・内線'; }));
    // 8) 記号のない電話番号に【電】
    log('凡例(電・記号なし)', editPara(p, TEL_RE, function (m) { return '【電】' + m[0]; }));
    return changes;
  }

  // 08: 「℡内線」はチェックのみ(置換しない)
  function findParagraph(p) {
    var out = [], text = joined(p), re = /(?:[℡☎]|【電】)内線/g, m;
    while ((m = re.exec(text)) !== null) out.push({ rule: '08', text: m[0], index: m.index });
    return out;
  }


  // ---- 書式のない凡例の候補(チェックのみ。置換はしない) ----
  // 「時＋日付」を起点に、そのあとの 場→師→費→定→保→持→申→問 を、順番どおりに探す。2つ以上そろったときだけ候補にする
  var PLAIN_STEPS = [
    { l: '場', re: /場(?=[^0-9０-９\s、。】])/g, weak: true, prev: /[分）)午時日正曜回まで頃半前後]/ },
    { l: '師', re: /師(?=[^\s、。】])/g, weak: true },
    { l: '費', re: /費(?=[0-9０-９]|無料|不要)/g },
    { l: '定', re: /定(?=[0-9０-９]+(?:人|名|組|家族|チーム|団体|世帯))/g },
    { l: '保', re: /保(?=[0-9０-９]+人|あり|なし|無|有)/g },
    { l: '持', re: /持(?=[^\s、。】])/g, weak: true },
    { l: '申', re: /申(?=[0-9０-９]+月|電話|電子|郵送|窓口|直接|はがき|ハガキ|メール|FAX|ＦＡＸ|ホームページ|往復|当日|不要|要|ご)/g },
    { l: '問', re: /問(?=[^\s、。】]{2,40}?(?:【電】|☎|℡|内線|[0-9０-９]{2,4}[-－ー（(][0-9０-９]|電話))/g }
  ];
  function ctxOf(t, i, len) { return (t.slice(Math.max(0, i - 6), i) + '[' + t.slice(i, i + len) + ']' + t.slice(i + len, i + len + 6)).replace(/\n/g, '↵'); }
  function findPlainLabels(p) {
    var t = joined(p), out = [], re = /時(?=[0-9０-９]{1,2}月)/g, m, a = -1;
    while ((m = re.exec(t)) !== null) { if (t[m.index - 1] !== '【') { a = m.index; break; } }
    if (a < 0) return out;
    var hits = [{ l: '時', i: a, weak: false }], pos = a;
    PLAIN_STEPS.forEach(function (s) {
      var r = new RegExp(s.re.source, 'g'), mm;
      r.lastIndex = pos + 1;
      while ((mm = r.exec(t)) !== null) {
        if (s.prev && !s.prev.test(t[mm.index - 1] || '')) continue;
        hits.push({ l: s.l, i: mm.index, weak: !!s.weak }); pos = mm.index; break;
      }
    });
    // 「F」のあとに電話番号
    var rf = /(?<![A-Za-zＡ-Ｚａ-ｚ])[FＦ](?=[0-9０-９]{2,4}[-－ー（(])/g, mf;
    while ((mf = rf.exec(t)) !== null) hits.push({ l: 'FAX', i: mf.index, weak: false, f: true });
    if (hits.filter(function (h) { return !h.f; }).length < 2) return out;   // 起点以外に1つもなければ候補にしない
    hits.forEach(function (h) {
      out.push({ rule: '凡例1b候補', text: ctxOf(t, h.i, h.f ? 1 : 1) + ' → 【' + h.l + '】' + (h.weak ? '(要確認)' : ''), index: h.i, label: h.l, weak: h.weak, note: '(置換しません。見分け方の確認用)' });
    });
    return out;
  }

  // ---- Ω: 小見出し(最初の行)のあとに、Ωだけの段落を入れる ----
  // 文書全体の処理。戻り値: null(何もしない) / {kind:'insert', before:n, heading:h}(n番目の段落の前に新しい段落) / {kind:'fill', para:n, heading:h}(空の段落nをΩにする)
  function tOf(p) { return joined(p).replace(/[\s\u3000\v]+/g, ''); }
  function headingIndex(paras) {   // 小見出し=最初の行(連載タイトルなら、その次の行)。なければ -1
    var h = -1, i;
    for (i = 0; i < paras.length; i++) if (tOf(paras[i]) && !isDeletableLine(paras[i])) { h = i; break; }
    if (h < 0) return -1;
    var SKIPH = function (p) { var x = tOf(p); return /第[0-9０-９]+回/.test(x) || /^[0-9０-９]{4}[\/／.][0-9０-９]{1,2}[\/／.][0-9０-９]{1,2}号?$/.test(x); };   // 連載タイトル・号数の行
    if (SKIPH(paras[h])) {
      var h2 = -1;
      for (i = h + 1; i < paras.length; i++) if (tOf(paras[i]) && !isDeletableLine(paras[i])) { h2 = i; break; }
      return h2;
    }
    return h;
  }

  // ---- 構2: 小見出し+凡例だけで本文がない文書。Ωは入れず、「●小見出し　【時】…」の1つの段落にする ----
  // 戻り値: null / {kind:'inline', heading:h, idx}(小見出しと凡例が同じ段落) / {kind:'join', heading:h, legend:n, idx}(凡例が次の段落)
  function kou2Info(paras) {
    var h = headingIndex(paras);
    if (h < 0) return null;
    var reL = /【(?:時|対)】/, ht = joined(paras[h]), m = reL.exec(ht), i;
    if (m) {
      var pre = ht.slice(0, m.index);
      if (m.index > 0 && m.index <= 60 && !/[。！？★]/.test(pre)) return { kind: 'inline', heading: h, idx: m.index };
      return null;
    }
    for (i = h + 1; i < paras.length; i++) if (tOf(paras[i]) && !isDeletableLine(paras[i])) break;
    if (i >= paras.length) return null;
    var t = joined(paras[i]); m = reL.exec(t);
    if (m && m.index <= 3 && !/[。！？]/.test(t.slice(0, m.index))) return { kind: 'join', heading: h, legend: i, idx: m.index };
    return null;
  }
  function planOmega(paras) {
    function t(p) { return tOf(p); }
    var h = headingIndex(paras), i;
    if (h < 0) return null;
    if (kou2Info(paras)) return null;     // 構2: Ωは入れない
    if (t(paras[h]) === 'Ω') return null;
    var ht = joined(paras[h]), sm = /★(?=[^\s\u3000★])/.exec(ht);
    if (sm) return { kind: 'star', para: h, idx: sm.index + 1, occ: (ht.slice(0, sm.index).match(/★/g) || []).length, heading: h };   // 小見出し★本文 → ★で改行してΩ(★は残す)
    var n = h + 1;
    if (n >= paras.length) return null;
    if (t(paras[n]) === 'Ω') return null;
    if (!t(paras[n])) {
      var m = n + 1; while (m < paras.length && !t(paras[m])) m++;
      if (m >= paras.length) return null;       // 本文がない
      if (t(paras[m]) === 'Ω') return null;     // 空行のあとにΩがすでにある
      return { kind: 'fill', para: n, heading: h };
    }
    return { kind: 'insert', before: n, heading: h };
  }
  // 予定の段落の文字の並び(expected)に、Ωを反映する
  function applyOmegaToTexts(texts, plan) {
    var a = texts.slice();
    if (!plan) return a;
    if (plan.kind === 'fill') a[plan.para] = 'Ω';
    else if (plan.kind === 'star') { var st = a[plan.para]; a.splice(plan.para, 1, st.slice(0, plan.idx), 'Ω', st.slice(plan.idx)); }
    else a.splice(plan.before, 0, 'Ω');
    return a;
  }
  // ---- 本文2: 凡例の前(本文の最後)に「。」を付けて、段落区切りで改行 ----
  // 戻り値: [ {para, idx, label, addPeriod} (段落の途中から凡例が始まる: idx の前で段落を分ける) | {para, end:true} (前の段落の末尾に「。」) ]
  function needPeriod(c) { return !/[\s\u3000。！？!?」』）)］\]、，,.．・：:；;…―—\-】]/.test(c); }
  function planBodyBreak(paras) {
    var h = headingIndex(paras), out = [], i;
    if (h < 0 || kou2Info(paras)) return out;     // 構2は、凡例の前で段落を分けない
    for (i = h + 1; i < paras.length; i++) {
      var t = joined(paras[i]), m = /【(?:対|時)】/.exec(t);
      if (!m || tOf(paras[i]) === 'Ω') continue;
      if (m.index > 0) {
        var before = t.slice(0, m.index).replace(/[\s\u3000\v]+$/, '');
        if (!before) break;
        out.push({ para: i, idx: m.index, label: m[0], addPeriod: needPeriod(before.charAt(before.length - 1)) });
      } else {
        var j = i - 1;
        if (j > h && tOf(paras[j]) && tOf(paras[j]) !== 'Ω' && !/^[【●]/.test(joined(paras[j]))) {
          var pt = joined(paras[j]).replace(/[\s\u3000\v]+$/, '');
          if (needPeriod(pt.charAt(pt.length - 1))) out.push({ para: j, end: true });
        }
      }
      break;   // 最初の凡例のかたまりだけ
    }
    return out;
  }

  // ---- 本文の最後が「？」「！」で終わるときは、黄色でマークする(「。」で終わるのが原則のイレギュラー) ----
  // 戻り値: [ {para, idx, ch} ] para番目の段落の idx 文字目(本文の最後の文字)
  function planEndMark(paras) {
    var h = headingIndex(paras), out = [], i;
    if (h < 0 || kou2Info(paras)) return out;
    for (i = h + 1; i < paras.length; i++) {
      var t = joined(paras[i]), m = /【(?:対|時)】/.exec(t);
      if (!m || tOf(paras[i]) === 'Ω') continue;
      var para = -1, txt = '';
      if (m.index > 0) { para = i; txt = t.slice(0, m.index); }
      else { var j = i - 1; if (j > h && tOf(paras[j]) && tOf(paras[j]) !== 'Ω' && !/^[【●]/.test(joined(paras[j]))) { para = j; txt = joined(paras[j]); } }
      if (para >= 0) {
        var end = txt.replace(/[\s\u3000\v]+$/, '').length, c = txt.charAt(end - 1);
        if (/[？！?!]/.test(c)) out.push({ para: para, idx: end - 1, ch: c });
      }
      break;
    }
    return out;
  }
  function applyBreakToTexts(texts, plans) {
    var a = texts.slice();
    plans.slice().sort(function (x, y) { return y.para - x.para; }).forEach(function (pl) {
      if (pl.end) { a[pl.para] = a[pl.para].replace(/[\s\u3000\v]+$/, '') + '。'; return; }
      var t = a[pl.para];
      var before = t.slice(0, pl.idx).replace(/[\s\u3000\v]+$/, '');
      a.splice(pl.para, 1, before + (pl.addPeriod ? '。' : ''), t.slice(pl.idx));
    });
    return a;
  }

  // ---- ルール13(不要な行の削除)・空白2(空行の削除)・本文1(凡例の段落をつなぐ) ----
  var DEL_RES = [
    /^[0-9０-９]{4}[\/／.][0-9０-９]{1,2}[\/／.][0-9０-９]{1,2}号?(?:原稿案)?$/,   // 「2026/10/25号」「2026/10/25号 原稿案」
    /^講師(?:名)?読み/,                                                           // 講師読み・講師名読み
    /^[0-9０-９]+文字$/, /^原稿案/, /^文字数[：:]/,
    /^「広報たちかわ」の横書き/, /^凡例[：:]/, /^【?公開する会議日程】?$/                 // 構2: 原稿書式の説明・凡例の説明・見出し
  ];
  function isDeletableLine(p) { var x = tOf(p); return !!x && DEL_RES.some(function (re) { return re.test(x); }); }
  function planStructure(paras) {
    var del = {}, keep = [], i;
    for (i = 0; i < paras.length; i++) { if (isDeletableLine(paras[i])) del[i] = true; else keep.push(i); }
    var sub = keep.map(function (k) { return paras[k]; }), om = planOmega(sub);
    var fillOrig = (om && om.kind === 'fill') ? keep[om.para] : -1, omO = null;
    if (om) { omO = {}; for (var k in om) omO[k] = om[k]; ['para', 'before', 'heading'].forEach(function (key) { if (om[key] !== undefined) omO[key] = keep[om[key]]; }); }
    var emptyDel = [];
    keep.forEach(function (oi) { var p = paras[oi]; if (!tOf(p) && !p.hasObj && oi !== fillOrig) emptyDel.push(oi); });
    var lines = Object.keys(del).map(Number);
    return { lines: lines, empties: emptyDel, del: lines.concat(emptyDel).sort(function (a, b) { return a - b; }), omega: omO };
  }
  // 凡例のかたまり(【時】【対】で始まる段落)に、すき間なく続く段落は、1つの段落につなぐ。つなぐ段落(j)の番号を返す
  function planMerge(paras) {
    var h = headingIndex(paras), out = [], s = -1, i;
    if (h < 0) return out;
    var k = kou2Info(paras);
    if (k && k.kind === 'join') { for (i = h + 1; i <= k.legend; i++) out.push(i); s = k.legend; }   // 小見出しと凡例の段落をつなぐ(間の空行も)
    else if (k) s = h;
    else for (i = h + 1; i < paras.length; i++) if (/^【(?:時|対)】/.test(joined(paras[i]))) { s = i; break; }
    if (s < 0) return out;
    for (i = s + 1; i < paras.length; i++) {
      var x = tOf(paras[i]);
      if (!x || /^https?:/i.test(x) || isDeletableLine(paras[i])) break;
      out.push(i);
    }
    return out;
  }
  function applyMergeToTexts(texts, js) {
    if (!js.length) return texts.slice();
    var a = texts.slice(), s = js[0] - 1;
    while (s > 0 && !a[s]) s--;
    for (var k = 0; k < js.length; k++) a[s] += a[js[k]];
    return a.filter(function (t, i) { return js.indexOf(i) < 0; });
  }
  // 構造の変更(不要な行の削除・★で改行・Ω・本文2・本文1)を、文字の並びの上で順に行った結果(確認用の予定)。paras は段落、または文字の配列
  function simulateStructure(input) {
    var paras = input.map(function (x) { return typeof x === 'string' ? { runs: x ? [{ text: x, boxed: false }] : [] } : x; });
    var mk = function (t) { return { runs: t ? [{ text: t, boxed: false }] : [] }; };
    var rest = paras.filter(function (p) { return !isDeletableLine(p); });
    var a = applyOmegaToTexts(rest.map(joined), planOmega(rest));
    a = applyBreakToTexts(a, planBodyBreak(a.map(mk)));
    return applyMergeToTexts(a, planMerge(a.map(mk)));
  }

  function fixAll(paras) {
    var all = [];
    var hd = headingIndex(paras);
    paras.forEach(function (p, i) { fixParagraph(p, { stripLead: hd >= 0 && i > hd }).forEach(function (c) { c.para = i; all.push(c); }); });
    // 構2: 「●」と、小見出しと凡例のあいだの全角スペース(変換後の【時】【対】を目印にする)
    var keep = [];
    paras.forEach(function (p, i) { if (!isDeletableLine(p)) keep.push(i); });
    var k = kou2Info(keep.map(function (i) { return paras[i]; }));
    if (k) {
      var hi = keep[k.heading], hp = paras[hi], ht = joined(hp), ed;
      function rec(pi, e) { all.push({ rule: '構2', before: e.before, after: e.t, occ: e.occ, para: pi }); }
      if (k.kind === 'inline') {
        var pc = ht.charAt(k.idx - 1);
        if (pc && !/[\u3000\s]/.test(pc)) { ed = editAt(hp, k.idx - 1, k.idx, pc + '\u3000'); rec(hi, ed); }
      } else {
        // 全角スペースは、凡例(枠つきの文字を置き換えたもの)ではなく、小見出しの最後の文字につける(枠が移らないように)
        if (k.idx === 0 && ht && !/[\u3000\s]$/.test(ht)) { ed = editAt(hp, ht.length - 1, ht.length, ht.charAt(ht.length - 1) + '\u3000'); rec(hi, ed); }
      }
      if (!/^●/.test(joined(hp))) { ed = editAt(hp, 0, 1, '●' + joined(hp).charAt(0)); rec(hi, ed); }
    }
    return all;
  }
  function findAll(paras) {
    var all = [];
    paras.forEach(function (p, i) { findParagraph(p).forEach(function (c) { c.para = i; all.push(c); }); });
    return all;
  }

  // Wordへ適用する変更(元の段落の文字列が基準)。fixAllのあとに呼ぶ
  //  { para, os, oe, before, after, box, occ, ambiguous }
  function collectEdits(paras) {
    var out = [];
    paras.forEach(function (p, i) {
      if (!p.pieces) return;
      p.pieces.forEach(function (pc) {
        if (!pc.edited) return;
        var before = p.orig.slice(pc.os, pc.oe);
        if (before === pc.text && !pc.box) return;
        var occ = countBefore(p.orig, before, pc.os);
        var ed = { para: i, os: pc.os, oe: pc.oe, before: before, after: pc.text, box: pc.box,
                   occ: occ, ambiguous: nthIndex(p.orig, before, occ) !== pc.os };
        // 特別な文字(symEx)を含む変更は、前後の普通の文字を目印にして、その間を置き換える
        var op = p.opaque || {}, hasOp = false, allOp = true;
        for (var q = pc.os; q < pc.oe; q++) { if (op[q]) hasOp = true; else allOp = false; }
        if (hasOp) {
          var b = '', k2 = pc.os - 1, a = '', k3 = pc.oe;
          while (k2 >= 0 && !op[k2] && b.length < 10 && !/[\u0000-\u001f]/.test(p.orig[k2])) { b = p.orig[k2] + b; k2--; }
          while (k3 < p.orig.length && !op[k3] && a.length < 10 && !/[\u0000-\u001f]/.test(p.orig[k3])) { a += p.orig[k3]; k3++; }
          ed.opaque = true; ed.unsupported = !allOp; ed.ctxB = b; ed.ctxA = a;
          ed.ctxBocc = b ? countBefore(p.orig, b, pc.os - b.length) : 0;
          ed.ctxAocc = a ? countBefore(p.orig, a, pc.oe) : 0;
          ed.ambiguous = !!((b && nthIndex(p.orig, b, ed.ctxBocc) !== pc.os - b.length) || (a && nthIndex(p.orig, a, ed.ctxAocc) !== pc.oe));
        }
        out.push(ed);
      });
    });
    return out;
  }

  // 変更のしかた: 前に足すだけ / 後ろに足すだけ / 削除 / 置換 (変更履歴を最小にするため)
  function classifyEdit(before, after, box) {
    if (after === '') return { mode: 'delete', text: '' };
    if (!box && after.length > before.length) {
      if (after.slice(after.length - before.length) === before) return { mode: 'insertStart', text: after.slice(0, after.length - before.length) };
      if (after.slice(0, before.length) === before) return { mode: 'insertEnd', text: after.slice(before.length) };
    }
    return { mode: 'replace', text: after };
  }

  return {
    LABEL_MAP: LABEL_MAP, joined: joined, editPara: editPara, fixParagraph: fixParagraph,
    fixAll: fixAll, findAll: findAll, findParagraph: findParagraph, findPlainLabels: findPlainLabels,
    simulateStructure: simulateStructure, planStructure: planStructure, planMerge: planMerge, headingIndex: headingIndex, isDeletableLine: isDeletableLine, planBodyBreak: planBodyBreak, applyBreakToTexts: applyBreakToTexts, planOmega: planOmega, applyOmegaToTexts: applyOmegaToTexts,
    kou2Info: kou2Info, planEndMark: planEndMark, collectEdits: collectEdits, classifyEdit: classifyEdit, isBoxedProps: isBoxedProps,
    countBefore: countBefore, nthIndex: nthIndex
  };
});
