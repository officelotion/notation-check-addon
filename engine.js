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
  function isBoxedProps(bdr, shd) {
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
  function fixParagraph(p) {
    var changes = [];
    initPieces(p);
    function log(rule, edits) { edits.forEach(function (e) { changes.push({ rule: rule, before: e.before, after: e.t, occ: e.occ }); }); }

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
  function planOmega(paras) {
    function t(p) { return joined(p).replace(/[\s\u3000\v]+/g, ''); }
    var h = -1, i;
    for (i = 0; i < paras.length; i++) if (t(paras[i])) { h = i; break; }
    if (h < 0) return null;
    if (/第[0-9０-９]+回/.test(t(paras[h]))) {   // 連載タイトルは残して、次の行を小見出しとする
      var h2 = -1;
      for (i = h + 1; i < paras.length; i++) if (t(paras[i])) { h2 = i; break; }
      if (h2 < 0) return null;
      h = h2;
    }
    if (t(paras[h]) === 'Ω') return null;
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
    if (plan.kind === 'fill') a[plan.para] = 'Ω'; else a.splice(plan.before, 0, 'Ω');
    return a;
  }

  function fixAll(paras) {
    var all = [];
    paras.forEach(function (p, i) { fixParagraph(p).forEach(function (c) { c.para = i; all.push(c); }); });
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
    planOmega: planOmega, applyOmegaToTexts: applyOmegaToTexts,
    collectEdits: collectEdits, classifyEdit: classifyEdit, isBoxedProps: isBoxedProps,
    countBefore: countBefore, nthIndex: nthIndex
  };
});
