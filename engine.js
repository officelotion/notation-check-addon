/*
 * 表記チェック Word版 ― 中核処理 (engine.js)  v0.1
 * Wordに依存しない純粋な処理。アドインでもNodeのテストでも同じものを使う。
 *
 * 段落モデル: { runs: [ { text: "…", boxed: true|false } ] }
 *   boxed = 文字の枠(罫線)またはグレーの網かけがついている
 *
 * 変更の記録: { rule, before, after, occ, box }
 *   occ = その段落の中で、同じ文字列(before)が何番目か(0から)。Wordで探すときに使う
 *   box = 枠・網かけを外す必要がある変更
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.NotationEngine = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // 凡例の元の文字 → 変換後
  var LABEL_MAP = {
    '対': '対', '時': '時', '場': '場', '師': '師', '費': '費', '定': '定', '保': '保',
    '持': '持', '申': '申', '問': '問',
    '℡': '電', '☎': '電',
    'Ｆ': 'FAX', 'F': 'FAX', 'fax': 'FAX', 'FAX': 'FAX', 'ＦＡＸ': 'FAX',
    'Ｅ': 'e', 'E': 'e'
  };

  function joined(p) { return p.runs.map(function (r) { return r.text; }).join(''); }

  function countBefore(text, str, pos) {
    var n = 0, i = text.indexOf(str);
    while (i !== -1 && i < pos) { n++; i = text.indexOf(str, i + 1); }
    return n;
  }

  // 段落の中の正規表現の一致を、書式(run)を保ったまま置き換える
  function editPara(p, re, fn) {
    var text = joined(p), edits = [], m;
    re.lastIndex = 0;
    while ((m = re.exec(text)) !== null) {
      if (m[0].length === 0) { re.lastIndex++; continue; }
      edits.push({ s: m.index, e: m.index + m[0].length, t: fn(m), before: m[0], occ: countBefore(text, m[0], m.index) });
    }
    // 順番どおり(先頭から)に適用するため、位置のずれを避けて後ろから適用し、記録は先頭からの順
    for (var k = edits.length - 1; k >= 0; k--) applyEdit(p, edits[k]);
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

  /*
   * 自動修正 (1段落)  戻り値: 変更の一覧
   * 「後ろから適用」のため、同じ規則の中の変更は、Wordへは後ろから順に適用する(adapter側)。
   */
  function fixParagraph(p) {
    var changes = [];
    function log(rule, edits) { edits.forEach(function (e) { changes.push({ rule: rule, before: e.before, after: e.t, occ: e.occ }); }); }

    // 1) 凡例: 枠・網かけつきの1文字 → 【●】(書式は外す)
    var pos = 0, snapshot = joined(p);
    p.runs.forEach(function (r) {
      var start = pos; pos += r.text.length;
      if (!r.boxed) return;
      var key = r.text.trim();
      if (Object.prototype.hasOwnProperty.call(LABEL_MAP, key)) {
        var after = '【' + LABEL_MAP[key] + '】';
        changes.push({ rule: '凡例', before: r.text, after: after, occ: countBefore(snapshot, r.text, start), box: true });
        r.text = after; r.boxed = false;
      }
    });
    // 2) 凡例: 書式なしの ℡ ☎ → 【電】
    log('凡例(電)', editPara(p, /[℡☎]/g, function () { return '【電】'; }));
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
    return changes;
  }

  // 08: 「℡内線」はチェックのみ(置換しない)
  function findParagraph(p) {
    var out = [], text = joined(p), re = /(?:[℡☎]|【電】)内線/g, m;
    while ((m = re.exec(text)) !== null) out.push({ rule: '08', text: m[0], index: m.index });
    return out;
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

  return { LABEL_MAP: LABEL_MAP, joined: joined, editPara: editPara, fixParagraph: fixParagraph, fixAll: fixAll, findAll: findAll, findParagraph: findParagraph };
});
