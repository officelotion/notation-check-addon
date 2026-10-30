/* 表記チェック Word版 ― アドインの画面とWordの接続 (taskpane.js)  v0.3.2(テスト版) */
/* global Office, Word, NotationEngine */
(function () {
  'use strict';
  var W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var LS_KEY = 'notationcheck_lastFixAt';
  var $ = function (id) { return document.getElementById(id); };
  var logEl, statusEl;
  var BUTTONS = ['btnCheck', 'btnFix', 'btnUndoSel', 'btnUndoAll', 'btnInfo', 'btnCopy'];

  function log(msg) { logEl.textContent += msg + '\n'; logEl.scrollTop = logEl.scrollHeight; }
  function status(msg) { statusEl.textContent = msg; }
  function busy(b) { BUTTONS.forEach(function (id) { if ($(id)) $(id).disabled = b; }); }
  function show(s) { return '「' + String(s).replace(/\v/g, '↵') + '」'; }
  // 照合用: 見えない文字(異体字セレクタ・幅なし文字など)を除き、全角半角の違いをそろえる
  function normA(s) { return String(s || '').replace(/[\u0000-\u001f\u007f\ufe00-\ufe0f\u200b-\u200d\u2060\ufeff\u00ad]/g, '').replace(/\u00a0/g, ' ').normalize('NFKC'); }
  var WILD = '\uE000';
  function markedText(p) { // 特別な文字(symEx)の位置を、ワイルドカードにした解析側の文字列
    var out = '';
    p.runs.forEach(function (r) {
      for (var i = 0; i < r.text.length; i++) out += (r.sym && r.sym.indexOf(i) >= 0) ? WILD : r.text[i];
    });
    return out;
  }
  function wildEq(a, b) { // a の WILD は、b の0文字または1文字に一致してよい
    if (a.indexOf(WILD) < 0) return a === b;
    var cur = { 0: true };
    for (var i = 0; i < a.length; i++) {
      var next = {};
      Object.keys(cur).forEach(function (js) {
        var j = +js;
        if (a[i] === WILD) { next[j] = true; if (j < b.length) next[j + 1] = true; }
        else if (j < b.length && a[i] === b[j]) next[j + 1] = true;
      });
      cur = next;
    }
    return !!cur[b.length];
  }
  function cp(c) { return c === undefined ? '(なし)' : 'U+' + c.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0'); }
  function firstDiff(a, b) { var n = Math.min(a.length, b.length), i = 0; while (i < n && a[i] === b[i]) i++; return i; }
  function norm(s) { return String(s || '').replace(/[\u0000-\u001f\u007f]/g, ''); }
  function short(s) { s = norm(s); return s.length > 40 ? s.slice(0, 40) + '…' : s; }
  function errText(e) { return (e && (e.code ? e.code + ': ' : '') + (e.message || e)) || '不明'; }

  // ---- 文書(OOXML) → 段落モデル(engine.js用) ----
  function buildModel(ooxml, noMerge) {   // noMerge: 段落記号が削除されていても、Wordの段落と同じく別々に数える(最初の読み込み用)
    var dom = new DOMParser().parseFromString(ooxml, 'application/xml');
    var body = dom.getElementsByTagNameNS(W_NS, 'body')[0];
    if (!body) throw new Error('文書の中身(w:body)が見つかりません');
    var paras = [];
    var SKIP = { del: 1, moveFrom: 1, Fallback: 1, txbxContent: 1, sectPr: 1, pPr: 1, drawing: 1, pict: 1, AlternateContent: 1 };

    function attr(el, name) { return el.getAttributeNS(W_NS, name) || el.getAttribute('w:' + name) || ''; }
    function symChar(ac) { // mc:AlternateContent の中の w16se:symEx(例: ☎ の絵文字)
      var se = ac.getElementsByTagNameNS('*', 'symEx')[0];
      if (!se) return '';
      var hex = '';
      for (var k = 0; k < se.attributes.length; k++) { if (se.attributes[k].localName === 'char') hex = se.attributes[k].value; }
      var cpv = parseInt(hex, 16);
      return cpv ? String.fromCodePoint(cpv) : '';
    }
    function runOf(r) {
      var text = '', bdr = null, shd = null, hl = '', sym = [];
      for (var c = r.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        var n = c.localName;
        if (n === 'rPr') {
          for (var q = c.firstChild; q; q = q.nextSibling) {
            if (q.nodeType !== 1) continue;
            if (q.localName === 'bdr') bdr = { val: attr(q, 'val') };
            else if (q.localName === 'shd') shd = { val: attr(q, 'val'), fill: attr(q, 'fill') };
            else if (q.localName === 'highlight') hl = attr(q, 'val');
          }
        } else if (n === 't') text += c.textContent;
        else if (n === 'tab') text += '\t';
        else if (n === 'br' || n === 'cr') text += '\v';
        else if (n === 'AlternateContent') { var sc = symChar(c); if (sc) { sym.push(text.length); text += sc; } }
      }
      return { text: text, boxed: NotationEngine.isBoxedProps(bdr, shd, hl), sym: sym.length ? sym : null };
    }
    function collect(el, p) {
      for (var c = el.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1 || SKIP[c.localName]) continue;
        if (c.localName === 'r') { var r = runOf(c); if (r.text) p.runs.push(r); }
        else if (c.localName !== 'p') collect(c, p);
      }
    }
    var carry = false;
    function markDeleted(pEl) {
      for (var a = pEl.firstChild; a; a = a.nextSibling) {
        if (a.nodeType === 1 && a.localName === 'pPr') {
          for (var b = a.firstChild; b; b = b.nextSibling) {
            if (b.nodeType === 1 && b.localName === 'rPr') {
              for (var q = b.firstChild; q; q = q.nextSibling) if (q.nodeType === 1 && q.localName === 'del') return true;
            }
          }
        }
      }
      return false;
    }
    function walk(node) {
      for (var c = node.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1 || SKIP[c.localName]) continue;
        if (c.localName === 'p') {
          var p;
          if (carry) { p = paras[paras.length - 1]; p.wcount++; } else { p = { runs: [], wcount: 1 }; paras.push(p); }
          var n0 = p.runs.length;
          collect(c, p);
          if (c.getElementsByTagNameNS(W_NS, 'drawing').length || c.getElementsByTagNameNS(W_NS, 'pict').length) p.hasObj = true;
          // 段落記号が変更履歴で削除されている(=次の段落とつながる)。ただし、段落ぜんたいの削除(中身がない)は、別の段落のまま数える
          carry = !noMerge && markDeleted(c) && p.runs.length > 0;
        }
        else walk(c);
      }
    }
    walk(body);
    return paras;
  }

  // 文書を読み込む(解析した段落 model と、Wordの段落 wp)
  async function loadModel(ctx, noMerge) {
    var body = ctx.document.body;
    var ox = body.getOoxml();
    var ps = body.paragraphs;
    ps.load('items/text');
    await ctx.sync();
    return { model: buildModel(ox.value, noMerge), wp: ps.items.slice() };
  }
  // 解析した段落とWordの段落を、文字で照合して対応づける。
  // 片方にだけある「空の段落」(getOoxmlが末尾に足すことがある)は、読み飛ばす。文字が食い違えば、そこで止める
  async function readDocument(ctx) {
    var L = await loadModel(ctx, true), model = L.model, wp = L.wp;
    var m2 = [], w2 = [], skipped = [], bad = null, i = 0, j = 0;
    while (i < model.length || j < wp.length) {
      var mt = i < model.length ? normA(markedText(model[i])) : null;
      var wt = j < wp.length ? normA(wp[j].text) : null;
      if (mt !== null && wt !== null && wildEq(mt, wt)) { m2.push(model[i]); w2.push(wp[j]); i++; j++; }
      else if (mt === '') { skipped.push('解析' + (i + 1) + '番目(空)'); i++; }
      else if (wt === '') { skipped.push('Word' + (j + 1) + '番目(空)'); j++; }
      else { bad = { i: i, j: j, mt: mt, wt: wt }; break; }
    }
    return { model: m2, raw: model, wordParas: w2, skipped: skipped, bad: bad, nModel: model.length, nWord: wp.length, aligned: !bad };
  }
  function explainMisalign(d) {
    var b = d.bad, k = firstDiff(b.mt, b.wt);
    function around(t) { return show(t.slice(Math.max(0, k - 8), k + 12)); }
    return '段落の照合が、ここで止まりました(Word ' + d.nWord + ' 段落 / 解析 ' + d.nModel + ' 段落)。' +
      '\n  食い違いの場所: 段落の先頭から ' + (k + 1) + ' 文字め' +
      '\n  Word ' + (b.j + 1) + '番目: ' + around(b.wt) + '  文字コード ' + cp(b.wt[k]) +
      '\n  解析 ' + (b.i + 1) + '番目: ' + around(b.mt) + '  文字コード ' + cp(b.mt[k]) +
      '\n(すでに変更履歴がある文書では、削除された文字のために食い違うことがあります)';
  }

  function summarize(changes) {
    var by = {};
    changes.forEach(function (c) { by[c.rule] = (by[c.rule] || 0) + 1; });
    var lines = Object.keys(by).map(function (k) { return '  ' + k + ' : ' + by[k] + '件'; });
    return lines.length ? lines.join('\n') : '  (直す場所はありません)';
  }
  function listChanges(changes, finds) {
    changes.slice(0, 80).forEach(function (c) {
      log('  段落' + (c.para + 1) + ' [' + c.rule + '] ' + show(c.before) + ' → ' + show(c.after));
    });
    if (changes.length > 80) log('  …ほか ' + (changes.length - 80) + ' 件');
    finds.forEach(function (f) { log('  段落' + (f.para + 1) + ' [' + f.rule + ' チェックのみ] ' + f.text + (f.note || '(自動では直しません)')); });
  }

  // ---- STEP 1: チェック ----
  async function doCheck() {
    busy(true); logEl.textContent = ''; status('チェック中…');
    try {
      await Word.run(async function (ctx) {
        var d = await readDocument(ctx);
        if (!d.aligned) log('注意: ' + explainMisalign(d) + '\n「自動修正」は実行できません。\n');
        else if (d.skipped.length) log('(空の段落の差を読み飛ばしました: ' + d.skipped.join(', ') + ')');
        var cm = d.aligned ? d.model : d.raw;   // 照合が止まったときも、チェックは解析した全段落で行う
        var finds = NotationEngine.findAll(cm);
        var changes = NotationEngine.fixAll(cm);
        var om = NotationEngine.planStructure(cm).omega;
        status('チェック完了: 直す場所 ' + (changes.length + (om ? 1 : 0)) + ' 件 / 確認だけ ' + finds.length + ' 件');
        log('【直す場所の内訳】\n' + summarize(changes) + '\n\n【一覧】');
        listChanges(changes, finds);
        NotationEngine.planBodyBreak(cm).forEach(function (pl) {
          log('  段落' + (pl.para + 1) + ' [本文2] ' + (pl.end ? '最後に「。」を付けます' : '凡例の前に、' + (pl.addPeriod ? '「。」と' : '') + '改行(段落区切り)を入れます'));
        });
        var sp2 = NotationEngine.planStructure(cm);
        if (sp2.lines.length) cm0Log(cm, sp2.lines);
        if (sp2.empties.length) log('  [空白2] 空行の段落 ' + sp2.empties.length + ' 件を削除します');
        var mg2 = NotationEngine.planMerge(cm);
        if (mg2.length) log('  [本文1] 凡例のあとに続く段落(' + mg2.length + ' 件)を、1つの段落につなぎます');
        if (om) log('  [Ω] 小見出し(段落' + (om.heading + 1) + ')' + (om.kind === 'star' ? 'の★のあとで改行して、Ωだけの段落を入れます(★は残します)' : 'のあとに、Ωだけの段落を' + (om.kind === 'fill' ? '空行から作ります' : '入れます')));
      });
    } catch (e) { status('エラー'); log('エラー: ' + errText(e)); }
    busy(false);
  }

  // ---- 枠・網かけを外す ----
  // 1回目: Wordの命令で外す(Macでは Font.shading が無いなど、使えない命令がある)
  // 2回目: 外れずに残った【●】を、「枠・網かけなし」の文字として、OOXMLで入れ直す
  async function trySet(ctx, fn) {
    try { fn(); await ctx.sync(); return ''; }
    catch (e) { return errText(e).slice(0, 80); }
  }
  async function clearBoxApi(ctx, rng) {
    var r = {};
    // 枠(Font.borders の各辺を、見えなくする)
    r.border = '';
    try {
      var bs = rng.font.borders; bs.load('items'); await ctx.sync();
      bs.items.forEach(function (b) { b.visible = false; });
      await ctx.sync();
    } catch (e) { r.border = errText(e).slice(0, 80); }
    // 網かけ(Range.shading。模様を「なし」にして、背景を白にする)
    var t1 = await trySet(ctx, function () { rng.shading.texture = 'TextureNone'; });
    if (t1) { var t2 = await trySet(ctx, function () { rng.shading.texture = 'None'; }); r.texture = t2 ? t1 : ''; } else r.texture = '';
    r.color = await trySet(ctx, function () { rng.shading.backgroundPatternColor = '#FFFFFF'; });
    return r;
  }

  var LABEL_RE = /^(?:【.{1,4}】|[（）])$/;   // 凡例の【●】と、枠が間違ってはみ出した（ ）
  // 枠・網かけが残っている【●】を探す(段落の番号と、同じ文字列の何番目か)
  function boxedLabelRuns(model) {
    var out = [];
    model.forEach(function (p, pi) {
      var text = NotationEngine.joined(p), pos = 0;
      p.runs.forEach(function (r) {
        var start = pos; pos += r.text.length;
        var t = r.text.trim();
        if (r.boxed && LABEL_RE.test(t)) out.push({ para: pi, text: t, occ: NotationEngine.countBefore(text, t, start) });
      });
    });
    return out;
  }
  // 解析した段落 mi に対応する、Wordの段落(つながった段落は、最初のもの)を探す。位置で対応づけ、文字が合わなければ、空でない段落の順で対応づける
  function subseq(a, b) { var i = 0; for (var j = 0; j < b.length && i < a.length; j++) if (a[i] === b[j]) i++; return i === a.length; }
  function wordParaOf(L, mi) {
    var k = 0, i;
    for (i = 0; i < mi; i++) k += (L.model[i].wcount || 1);
    var cnt = L.model[mi].wcount || 1, txt = '';
    for (i = 0; i < cnt && k + i < L.wp.length; i++) txt += L.wp[k + i].text;
    if (k < L.wp.length && subseq(normA(NotationEngine.joined(L.model[mi])), normA(txt))) return { first: L.wp[k], last: L.wp[Math.min(k + cnt - 1, L.wp.length - 1)], idx: k, cnt: cnt };
    var mI = nonEmptyIdx(L.model, NotationEngine.joined), wI = nonEmptyIdx(L.wp, function (p) { return p.text; });
    var rank = mI.indexOf(mi);
    if (rank >= 0 && wI[rank] !== undefined) return { first: L.wp[wI[rank]], last: L.wp[wI[rank]], idx: wI[rank], cnt: 1 };
    return null;
  }
  function nonEmptyIdx(arr, getText) {
    var m = [];
    arr.forEach(function (x, i) { if (norm(getText(x))) m.push(i); });
    return m;
  }
  function ooxmlPlainRun(text) {
    var esc = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    return '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
      '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>' +
      '</pkg:xmlData></pkg:part>' +
      '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>' +
      '<w:document xmlns:w="' + W_NS + '"><w:body><w:p><w:r><w:rPr><w:highlight w:val="none"/><w:bdr w:val="nil"/><w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:rPr>' +
      '<w:t xml:space="preserve">' + esc + '</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>';
  }
  // 残った枠・網かけを、OOXMLで入れ直して外す。戻り値: {found, done, ng}
  async function fixRemainingBoxes(ctx) {
    var L = await loadModel(ctx);
    var runs = boxedLabelRuns(L.model);
    var res = { found: runs.length, done: 0, ng: 0 };
    if (!runs.length) return res;
    var work = runs.map(function (r) {
      var wf = wordParaOf(L, r.para);
      return { para: wf ? wf.first : null, text: r.text, occ: r.occ, pi: r.para };
    });
    var searches = work.map(function (w) {
      if (!w.para) return null;
      var rs = w.para.search(w.text, { matchCase: true });
      rs.load('items/text');
      return rs;
    });
    await ctx.sync();
    for (var i = work.length - 1; i >= 0; i--) {
      var w = work[i];
      var cands = searches[i] ? searches[i].items.filter(function (r) { return r.text === w.text; }) : [];
      var t = cands[w.occ];
      if (!t) { res.ng++; log('  × ' + show(w.text) + '(解析' + (w.pi + 1) + '番目の段落)の場所を特定できません'); continue; }
      try { t.insertOoxml(ooxmlPlainRun(w.text), 'Replace'); await ctx.sync(); res.done++; }
      catch (e) { res.ng++; log('  × ' + show(w.text) + ' をOOXMLで入れ直せません: ' + errText(e)); }
    }
    return res;
  }

  // ---- STEP 2: 置換(変更履歴つき) ----
  // 手順: ①段落ごとに、変える場所を、変更前に全部探す(元の文字が基準なので、変更履歴の「削除された文字」に影響されない)
  //       ②探した場所を、後ろから順に変える ③最後に、文書を読み直して、予定どおりの文字になったか確かめる
  function cm0Log(cm, lines) { lines.forEach(function (i) { log('  段落' + (i + 1) + ' [13 削除] 「' + short(NotationEngine.joined(cm[i])) + '」の行を削除します'); }); }
  async function doFix() {
    busy(true); logEl.textContent = ''; status('置換中…');
    var okN = 0, ngN = 0, boxStat = { n: 0, border: {}, texture: {}, color: {} }, boxFix = null, boxLeft0 = 0;
    var prevMode = null, modeChanged = false;
    try {
      await Word.run(async function (ctx) {
        var d = await readDocument(ctx);
        if (!d.aligned) { status('中止'); log('置換を中止しました。\n' + explainMisalign(d) + '\n文書の画像・テキストボックス・既存の変更履歴などが原因かもしれません。この文書のまま、作者に知らせてください。'); return; }
        var changes = NotationEngine.fixAll(d.model);
        var edits = NotationEngine.collectEdits(d.model);
        var expected = d.model.map(NotationEngine.joined);
        var sp = NotationEngine.planStructure(d.model);   // 不要な行・空行の削除、Ω(元の段落の番号で)
        var omega = sp.omega;
        var bplan = NotationEngine.planBodyBreak(d.model);
        var star = !!(omega && omega.kind === 'star');
        expected = NotationEngine.simulateStructure(d.model);
        edits = edits.filter(function (e) { return sp.del.indexOf(e.para) < 0; });   // 削除する段落の中の変更は、しない
        if (!edits.length && !omega && !sp.del.length && !bplan.length && !NotationEngine.planMerge(d.model).length) { status('直す場所はありませんでした'); return; }

        ctx.document.load('changeTrackingMode');
        await ctx.sync();
        prevMode = ctx.document.changeTrackingMode;
        ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        await ctx.sync();
        modeChanged = true;
        try { localStorage.setItem(LS_KEY, String(Date.now())); } catch (e) { /* 保存できなくても続ける */ }

        // Ω: 小見出しのあとに、Ωだけの段落を入れる(空行があれば、Ωにする)。ほかの変更より先に行う
        if (omega && !star) {
          try {
            if (omega.kind === 'fill') d.wordParas[omega.para].insertText('Ω', 'Start');
            else d.wordParas[omega.before].insertParagraph('Ω', 'Before');
            await ctx.sync();
            okN++;
            log('  Ω: 小見出し(段落' + (omega.heading + 1) + ')のあとに、Ωだけの段落を' + (omega.kind === 'fill' ? '空行から作りました' : '入れました'));
          } catch (err) { ngN++; log('  × Ω の段落を入れられません: ' + errText(err)); }
        }

        // 不要な行(ルール13)と空行(空白2)の段落を、削除する(変更履歴つき)
        if (sp.del.length) {
          try {
            sp.del.slice().sort(function (a, b) { return b - a; }).forEach(function (i) { d.wordParas[i].delete(); });
            await ctx.sync();
            okN += sp.del.length;
            log('  削除: 不要な行 ' + sp.lines.length + ' 件・空行 ' + sp.empties.length + ' 件の段落を削除しました');
          } catch (err) { ngN++; log('  × 段落の削除に失敗: ' + errText(err)); }
        }

        var byPara = {}, order = [];
        edits.forEach(function (e) { if (!byPara[e.para]) { byPara[e.para] = []; order.push(e.para); } byPara[e.para].push(e); });

        for (var oi = 0; oi < order.length; oi++) {
          var pi = order[oi], list = byPara[pi], para = d.wordParas[pi];
          // ① 探す(変更前)
          var searches = list.map(function (e) {
            if (e.opaque) { // 特別な文字: 前後の普通の文字を探す
              if (e.ambiguous || e.unsupported) return null;
              var sb = e.ctxB ? para.search(e.ctxB, { matchCase: true }) : null;
              var sa = e.ctxA ? para.search(e.ctxA, { matchCase: true }) : null;
              if (sb) sb.load('items/text');
              if (sa) sa.load('items/text');
              return { opq: true, b: sb, a: sa };
            }
            if (e.ambiguous || e.before.length > 250 || /[\u0000-\u001f]/.test(e.before)) return null;
            var rs = para.search(e.before, { matchCase: true });
            rs.load('items/text');
            return rs;
          });
          await ctx.sync();
          var targets = list.map(function (e, i) {
            var sr = searches[i];
            if (!sr) return null;
            if (sr.opq) {
              var rB = sr.b ? sr.b.items.filter(function (r) { return r.text === e.ctxB; })[e.ctxBocc] : null;
              var rA = sr.a ? sr.a.items.filter(function (r) { return r.text === e.ctxA; })[e.ctxAocc] : null;
              if ((e.ctxB && !rB) || (e.ctxA && !rA)) return null;
              var st = rB ? rB.getRange('End') : para.getRange('Content').getRange('Start');
              var en = rA ? rA.getRange('Start') : para.getRange('Content').getRange('End');
              var between = st.expandTo(en);
              between.load('text');
              return { between: between, opq: true };
            }
            var cands = sr.items.filter(function (r) { return r.text === e.before; });
            return cands[e.occ] || null;
          });
          if (targets.some(function (t) { return t && t.opq; })) {
            await ctx.sync();   // 間の文字数を確かめる(大きく食い違うときは、触らない)
            targets = targets.map(function (t) { return (t && t.opq) ? (t.between.text.length <= 3 ? t.between : null) : t; });
          }
          // ② 後ろから変える
          var idx = list.map(function (e, i) { return i; }).sort(function (a, b) { return list[b].os - list[a].os; });
          for (var k = 0; k < idx.length; k++) {
            var e = list[idx[k]], target = targets[idx[k]];
            if (!target) { ngN++; log('  × 段落' + (pi + 1) + ' ' + show(e.before) + ' の場所を特定できません' + (e.ambiguous ? '(同じ文字が重なっている)' : '')); continue; }
            var cls = NotationEngine.classifyEdit(e.before, e.after, e.box);
            try {
              var nr = null;
              if (cls.mode === 'delete') target.delete();
              else if (cls.mode === 'insertStart') target.insertText(cls.text, 'Start');
              else if (cls.mode === 'insertEnd') target.insertText(cls.text, 'End');
              else nr = target.insertText(cls.text, 'Replace');
              await ctx.sync();
              if (e.box && nr) {
                var res = await clearBoxApi(ctx, nr);
                boxStat.n++;
                ['border', 'texture', 'color'].forEach(function (key) {
                  var v = res[key] || '○';
                  boxStat[key][v] = (boxStat[key][v] || 0) + 1;
                });
              }
              okN++;
            } catch (err) { ngN++; log('  × 段落' + (pi + 1) + ' ' + show(e.before) + ' の変更に失敗: ' + errText(err)); }
          }
        }

        // ②b 枠・網かけが残っていないか確かめ、残っていればOOXMLで入れ直す
        var Lb = await loadModel(ctx);
        boxLeft0 = boxedLabelRuns(Lb.model).length;
        if (boxLeft0) boxFix = await fixRemainingBoxes(ctx);

        // ②c 本文2: 凡例の前に「。」と段落区切り(Enter)を入れる。段落の数が変わるので、最後に行う
        {
          var Lx = await loadModel(ctx);
          if (star) {   // 小見出し★本文 → ★で改行して、Ωだけの段落を入れる(★は残す)
            try {
              var hi = NotationEngine.headingIndex(Lx.model), hwf = hi >= 0 ? wordParaOf(Lx, hi) : null, hw = hwf ? hwf.first : null;
              if (!hw) throw new Error('小見出しの場所を特定できません');
              var srs = hw.search('★', { matchCase: true }); srs.load('items'); await ctx.sync();
              if (srs.items.length <= omega.occ) throw new Error('★の位置が見つかりません');
              srs.items[omega.occ].insertText('\nΩ\n', 'After'); await ctx.sync();
              okN++; log('  ★: 小見出し(段落' + (omega.heading + 1) + ')の★のあとで改行し、Ωだけの段落を入れました(★は残します)');
              Lx = await loadModel(ctx);
            } catch (err) { ngN++; log('  × ★での改行に失敗: ' + errText(err)); }
          }
          var bp = NotationEngine.planBodyBreak(Lx.model);
          bp.sort(function (x, y) { return y.para - x.para; });
          for (var bi = 0; bi < bp.length; bi++) {
            var pl = bp[bi], wfx = wordParaOf(Lx, pl.para), wpar = wfx ? (pl.end ? wfx.last : wfx.first) : null;
            if (!wpar) { ngN++; log('  × 本文2: 段落' + (pl.para + 1) + ' の場所を特定できません'); continue; }
            try {
              if (pl.end) { wpar.insertText('。', 'End'); await ctx.sync(); }
              else {
                var rs = wpar.search(pl.label, { matchCase: true }); rs.load('items'); await ctx.sync();
                if (!rs.items.length) throw new Error('凡例の位置が見つかりません');
                rs.items[0].insertText((pl.addPeriod ? '。' : '') + '\n', 'Start'); await ctx.sync();
              }
              okN++;
              log('  本文2: 段落' + (pl.para + 1) + (pl.end ? ' の最後に「。」を付けました' : ' の凡例の前に、' + (pl.addPeriod ? '「。」と' : '') + '改行を入れました'));
            } catch (err) { ngN++; log('  × 本文2: 段落' + (pl.para + 1) + ' の変更に失敗: ' + errText(err)); }
          }
        }

        // ②d 本文1: 凡例のかたまりに、すき間なく続く段落を、1つの段落につなぐ(段落記号を削除)
        for (var mg = 0; mg < 30; mg++) {
          var Lm = await loadModel(ctx), js = NotationEngine.planMerge(Lm.model);
          if (!js.length) break;
          try {
            var wb = wordParaOf(Lm, js[0]), wa = wordParaOf(Lm, js[0] - 1);
            if (!wb || !wa) throw new Error('段落の場所を特定できません');
            wa.last.getRange('Content').getRange('End').expandTo(wb.first.getRange('Start')).delete();
            await ctx.sync();
            var Ln = await loadModel(ctx), js2 = NotationEngine.planMerge(Ln.model);
            if (js2.length >= js.length) throw new Error('段落をつなげませんでした');
            okN++; if (mg === 0) log('  本文1: 凡例のあとに続く段落(' + js.length + ' 件)を、1つの段落につなぎます');
          } catch (err) { ngN++; log('  × 段落をつなげません: ' + errText(err)); break; }
        }

        // ③ 確かめる(空でない段落の並びを、予定と比べる)
        var L2 = await loadModel(ctx);
        var act = L2.model.map(function (p) { return norm(NotationEngine.joined(p)); }).filter(function (t) { return t; });
        var exp = expected.map(norm).filter(function (t) { return t; });
        var bad = [];
        if (act.length !== exp.length) bad.push('段落の数が変わりました(' + exp.length + ' → ' + act.length + ')');
        else act.forEach(function (t, i) {
          var ta = normA(t), ea = normA(exp[i]);
          if (ta === ea) return;
          var d = firstDiff(ea, ta), a0 = Math.max(0, d - 8);
          bad.push('段落' + (i + 1) + '(空を除く) 差は' + (d + 1) + '文字目付近 予定=' + show(ea.slice(a0, d + 12)) + ' / 実際=' + show(ta.slice(a0, d + 12)) + ' (予定' + cp(ea[d]) + ' / 実際' + cp(ta[d]) + ' ・文字数 ' + ea.length + ' / ' + ta.length + ')');
        });

        status('置換完了: ' + okN + ' 箇所' + (ngN ? ' / 失敗 ' + ngN + ' 箇所' : '') + (bad.length ? ' / 確認で不一致 ' + bad.length : ' / 確認OK'));
        log('\n変更履歴つきで直しました。「校閲」タブの変更履歴で確認できます。');
        log('\n【確認】予定どおりの文字になったか: ' + (bad.length ? '不一致 ' + bad.length + ' 段落' : 'すべて一致'));
        bad.slice(0, 5).forEach(function (b) { log('  ' + b); });
        var left = boxedLabelRuns(L2.model).length;
        log('\n【枠・網かけ】 外す対象 ' + boxStat.n + ' 箇所');
        log('  Wordの命令で外せなかった数: ' + boxLeft0 + ' 箇所' + (boxLeft0 ? '(下の2回目で入れ直し)' : ''));
        if (boxFix) log('  OOXMLで入れ直した数: ' + boxFix.done + ' 箇所' + (boxFix.ng ? ' / 失敗 ' + boxFix.ng : ''));
        log('  最終的に枠・網かけが残った数: ' + left + ' 箇所' + (left ? '  ← 手で外す必要があります' : '  ← OK'));
        if (boxStat.n && boxLeft0) {
          log('  (参考) 1回目の命令の結果:');
          [['border', '枠'], ['texture', '網かけの模様'], ['color', '網かけの色']].forEach(function (p) {
            log('    ' + p[1] + ': ' + Object.keys(boxStat[p[0]]).map(function (v) { return v + ' ×' + boxStat[p[0]][v]; }).join(' / '));
          });
        }
      });
    } catch (e) { status('エラー'); log('エラー: ' + errText(e) + (e.debugInfo ? '\n' + JSON.stringify(e.debugInfo) : '')); }
    // 変更履歴の設定を、実行前に戻す
    if (modeChanged) {
      try { await Word.run(async function (ctx) { ctx.document.changeTrackingMode = prevMode; await ctx.sync(); }); }
      catch (e) { log('(変更履歴の設定を元に戻せませんでした: ' + errText(e) + ')'); }
    }
    busy(false);
  }

  // ---- STEP 3: 選択範囲の変更を元に戻す ----
  // 「対」→「【対】」は、削除と追加の2つの変更になる。選択の端に接する側が漏れると、文字が消えたように見えるため、
  // 選択範囲に「重なる」「隣り合う」変更を、段落ごとに集めて、まとめて戻す
  async function doUndo(all) {
    busy(true); logEl.textContent = ''; status('元に戻し中…');
    try {
      await Word.run(async function (ctx) {
        var sel = ctx.document.getSelection();
        sel.load('isEmpty');
        await ctx.sync();
        if (sel.isEmpty) { status('範囲を選択してください'); log('元に戻したい範囲を、ドラッグで選択してから押してください。'); return; }
        function isParaMark(tc) {
          var tx = String(tc.text || '');
          return (tc.type === 'Deleted' && /^[\r\n\u000b\u0007]+$/.test(tx)) || (tc.type === 'Added' && /^Ω[\r\n\u0007]*$/.test(tx));
        }
        var since = 0;
        try { since = parseInt(localStorage.getItem(LS_KEY) || '0', 10) || 0; } catch (e) { /* なし */ }
        // 選択に触れた段落の変更履歴を集める(Wordが変更を結合・分割することがあるため、戻したあとにもう一度集め直す)
        async function collect() {
          var ps = sel.paragraphs; ps.load('items');
          await ctx.sync();
          // 段落の先頭に追加された文字(例: 【時】)は、その段落の変更履歴として返らないことがあるため、
          // 前の段落〜あとの段落までの広い範囲から集める
          var first = ps.items[0], last = ps.items[ps.items.length - 1];
          var prevP = first.getPreviousOrNullObject(), nextP = last.getNextOrNullObject();
          prevP.load('isNullObject'); nextP.load('isNullObject');
          await ctx.sync();
          var wide = (prevP.isNullObject ? first : prevP).getRange('Start').expandTo((nextP.isNullObject ? last : nextP).getRange('End'));
          var tcs = wide.getTrackedChanges(); tcs.load('items/date,items/type,items/text');
          await ctx.sync();
          var lists = [tcs];
          var cand = [];
          lists.forEach(function (l) { l.items.forEach(function (tc) { cand.push({ tc: tc, rel: tc.getRange().compareLocationWith(sel) }); }); });
          await ctx.sync();
          var hit = cand.filter(function (c) { return ['Before', 'After', 'Unrelated'].indexOf(c.rel.value) < 0 || isParaMark(c.tc); });
          var mine = (all || !since) ? hit : hit.filter(function (c) { return c.tc.date && c.tc.date.getTime() >= since - 61000; });
          return { hit: hit, mine: mine };
        }
        function show(c) { var t = String(c.tc.text || '').replace(/[\r\n\u000b\u0007]/g, '¶'); if (t.length > 12) t = t.slice(0, 12) + '…'; return c.tc.type + '「' + t + '」' + c.rel.value; }
        var r = await collect();
        if (!r.hit.length) { status('選択範囲に変更履歴はありません'); return; }
        if (!r.mine.length) {
          status('このアドインの変更は、選択範囲にありません');
          log('選択範囲にある変更履歴 ' + r.hit.length + ' 件は、アドインの修正より前のものです。残しました。\nすべて戻したいときは「選択範囲のすべての変更履歴を戻す」を使ってください。');
          return;
        }
        var other = r.hit.length - r.mine.length, total = 0, kinds = {}, rels = {}, pass = 0;
        log('【戻す前の変更履歴(種類「文字」位置)】\n' + r.mine.map(show).join('\n'));
        while (r.mine.length && pass < 6) {
          pass++;
          r.mine.forEach(function (c) { kinds[c.tc.type] = (kinds[c.tc.type] || 0) + 1; rels[c.rel.value] = (rels[c.rel.value] || 0) + 1; });
          // 後ろから戻す(前の変更の位置が、あとの変更でずれないように)
          for (var i = r.mine.length - 1; i >= 0; i--) { r.mine[i].tc.reject(); }
          await ctx.sync();
          total += r.mine.length;
          var prev = r.mine.length;
          r = await collect();
          if (r.mine.length) log('(' + pass + '回目のあとに、まだ ' + r.mine.length + ' 件残っていたので、もう一度戻します)');
          if (r.mine.length >= prev && pass >= 3) break;
        }
        status('選択範囲の変更を ' + total + ' 件、元に戻しました');
        log('内訳: ' + Object.keys(kinds).map(function (k) { return k + ' ' + kinds[k]; }).join(' / '));
        log('選択範囲との位置関係: ' + Object.keys(rels).map(function (k) { return k + ' ' + rels[k]; }).join(' / '));
        if (r.mine.length) log('【戻せなかった変更 ' + r.mine.length + ' 件】\n' + r.mine.map(show).join('\n'));
        else log('戻したあとの確認: 選択範囲にアドインの変更は残っていません');
        if (other) log('アドインの修正より前の変更履歴 ' + other + ' 件は、残しました。');
      });
    } catch (e) { status('エラー'); log('エラー: ' + errText(e) + '\n(選択範囲の取り消しに対応していないWordかもしれません。「テスト情報」で WordApi 1.6 を確認してください)'); }
    busy(false);
  }

  // ---- テスト情報 ----
  function doInfo() {
    logEl.textContent = '';
    var r = Office.context.requirements;
    function sup(set, v) { try { return r.isSetSupported(set, v) ? '○' : '×'; } catch (e) { return '?'; } }
    log('アドイン v0.8.1');
    log('プラットフォーム: ' + (Office.context.platform || '?'));
    try { var dg = Office.context.diagnostics; if (dg) log('Word: ' + dg.version + ' / ' + dg.platform); } catch (e) { /* なし */ }
    log('対応状況(○=使える)');
    [['WordApi', '1.3'], ['WordApi', '1.4'], ['WordApi', '1.6'], ['WordApiDesktop', '1.1'], ['WordApiDesktop', '1.2']].forEach(function (x) {
      log('  ' + x[0] + ' ' + x[1] + ' : ' + sup(x[0], x[1]));
    });
    var ls = '○'; try { localStorage.setItem('nc_t', '1'); localStorage.removeItem('nc_t'); } catch (e) { ls = '×'; }
    log('  保存領域(元に戻すの判定に使用) : ' + ls);
    log('\nWordApi 1.4 = 変更履歴 / 1.6 = 選択範囲の元に戻す / Desktop 1.1 = 枠・網かけ');
    status('テスト情報を表示しました');
  }

  // ---- ログをコピー ----
  function doCopy() {
    var ta = document.createElement('textarea');
    ta.value = logEl.textContent; ta.style.position = 'fixed'; ta.style.opacity = '0';
    document.body.appendChild(ta); ta.select();
    var okc = false;
    try { okc = document.execCommand('copy'); } catch (e) { okc = false; }
    document.body.removeChild(ta);
    status(okc ? 'ログをコピーしました' : 'コピーできませんでした。ログを選択して、手でコピーしてください');
  }

  Office.onReady(function (info) {
    logEl = $('log'); statusEl = $('status');
    if (info.host !== Office.HostType.Word) { status('このアドインはWordで使います'); return; }
    $('btnCheck').onclick = doCheck;
    $('btnFix').onclick = doFix;
    $('btnUndoSel').onclick = function () { doUndo(false); };
    $('btnUndoAll').onclick = function () { doUndo(true); };
    $('btnInfo').onclick = doInfo;
    $('btnCopy').onclick = doCopy;
    status('準備できました。まず「チェックする」を押してください');
  });
})();
