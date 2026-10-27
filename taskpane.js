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
  function buildModel(ooxml) {
    var dom = new DOMParser().parseFromString(ooxml, 'application/xml');
    var body = dom.getElementsByTagNameNS(W_NS, 'body')[0];
    if (!body) throw new Error('文書の中身(w:body)が見つかりません');
    var paras = [];
    var SKIP = { Fallback: 1, txbxContent: 1, sectPr: 1, pPr: 1, drawing: 1, pict: 1, AlternateContent: 1 };

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
      var text = '', bdr = null, shd = null, sym = [];
      for (var c = r.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        var n = c.localName;
        if (n === 'rPr') {
          for (var q = c.firstChild; q; q = q.nextSibling) {
            if (q.nodeType !== 1) continue;
            if (q.localName === 'bdr') bdr = { val: attr(q, 'val') };
            else if (q.localName === 'shd') shd = { val: attr(q, 'val'), fill: attr(q, 'fill') };
          }
        } else if (n === 't') text += c.textContent;
        else if (n === 'tab') text += '\t';
        else if (n === 'br' || n === 'cr') text += '\v';
        else if (n === 'AlternateContent') { var sc = symChar(c); if (sc) { sym.push(text.length); text += sc; } }
      }
      return { text: text, boxed: NotationEngine.isBoxedProps(bdr, shd), sym: sym.length ? sym : null };
    }
    function collect(el, p) {
      for (var c = el.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1 || SKIP[c.localName]) continue;
        if (c.localName === 'r') { var r = runOf(c); if (r.text) p.runs.push(r); }
        else if (c.localName !== 'p') collect(c, p);
      }
    }
    function walk(node) {
      for (var c = node.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1 || SKIP[c.localName]) continue;
        if (c.localName === 'p') { var p = { runs: [] }; paras.push(p); collect(c, p); }
        else walk(c);
      }
    }
    walk(body);
    return paras;
  }

  // 文書を読み込む(解析した段落 model と、Wordの段落 wp)
  async function loadModel(ctx) {
    var body = ctx.document.body;
    var ox = body.getOoxml();
    var ps = body.paragraphs;
    ps.load('items/text');
    await ctx.sync();
    return { model: buildModel(ox.value), wp: ps.items.slice() };
  }
  // 解析した段落とWordの段落を、文字で照合して対応づける。
  // 片方にだけある「空の段落」(getOoxmlが末尾に足すことがある)は、読み飛ばす。文字が食い違えば、そこで止める
  async function readDocument(ctx) {
    var L = await loadModel(ctx), model = L.model, wp = L.wp;
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
        status('チェック完了: 直す場所 ' + changes.length + ' 件 / 確認だけ ' + finds.length + ' 件');
        log('【直す場所の内訳】\n' + summarize(changes) + '\n\n【一覧】');
        listChanges(changes, finds);
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

  var LABEL_RE = /^【.{1,4}】$/;
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
      '<w:document xmlns:w="' + W_NS + '"><w:body><w:p><w:r><w:rPr><w:bdr w:val="nil"/><w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:rPr>' +
      '<w:t xml:space="preserve">' + esc + '</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>';
  }
  // 残った枠・網かけを、OOXMLで入れ直して外す。戻り値: {found, done, ng}
  async function fixRemainingBoxes(ctx) {
    var L = await loadModel(ctx);
    var runs = boxedLabelRuns(L.model);
    var res = { found: runs.length, done: 0, ng: 0 };
    if (!runs.length) return res;
    var mIdx = nonEmptyIdx(L.model, NotationEngine.joined);
    var wIdx = nonEmptyIdx(L.wp, function (p) { return p.text; });
    var work = runs.map(function (r) {
      var rank = mIdx.indexOf(r.para);
      var wp = (rank >= 0 && wIdx[rank] !== undefined) ? L.wp[wIdx[rank]] : null;
      return { para: wp, text: r.text, occ: r.occ, pi: r.para };
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
        if (!edits.length) { status('直す場所はありませんでした'); return; }

        ctx.document.load('changeTrackingMode');
        await ctx.sync();
        prevMode = ctx.document.changeTrackingMode;
        ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        await ctx.sync();
        modeChanged = true;
        try { localStorage.setItem(LS_KEY, String(Date.now())); } catch (e) { /* 保存できなくても続ける */ }

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
        var paras = sel.paragraphs; paras.load('items');
        await ctx.sync();
        if (sel.isEmpty) { status('範囲を選択してください'); log('元に戻したい範囲を、ドラッグで選択してから押してください。'); return; }
        var lists = paras.items.map(function (p) { var t = p.getRange('Whole').getTrackedChanges(); t.load('items/date,items/type'); return t; });
        await ctx.sync();
        var cand = [];
        lists.forEach(function (l) { l.items.forEach(function (tc) { cand.push({ tc: tc, rel: tc.getRange().compareLocationWith(sel) }); }); });
        await ctx.sync();
        var hit = cand.filter(function (c) { return ['Before', 'After', 'Unrelated'].indexOf(c.rel.value) < 0; });
        if (!hit.length) { status('選択範囲に変更履歴はありません'); return; }
        var since = 0;
        try { since = parseInt(localStorage.getItem(LS_KEY) || '0', 10) || 0; } catch (e) { /* なし */ }
        var mine = (all || !since) ? hit : hit.filter(function (c) { return c.tc.date && c.tc.date.getTime() >= since - 61000; });
        var other = hit.length - mine.length;
        if (!mine.length) {
          status('このアドインの変更は、選択範囲にありません');
          log('選択範囲にある変更履歴 ' + hit.length + ' 件は、アドインの修正より前のものです。残しました。\nすべて戻したいときは「選択範囲のすべての変更履歴を戻す」を使ってください。');
          return;
        }
        var kinds = {}, rels = {};
        mine.forEach(function (c) { kinds[c.tc.type] = (kinds[c.tc.type] || 0) + 1; rels[c.rel.value] = (rels[c.rel.value] || 0) + 1; });
        mine.forEach(function (c) { c.tc.reject(); });
        await ctx.sync();
        status('選択範囲の変更を ' + mine.length + ' 件、元に戻しました');
        log('内訳: ' + Object.keys(kinds).map(function (k) { return k + ' ' + kinds[k]; }).join(' / '));
        log('選択範囲との位置関係: ' + Object.keys(rels).map(function (k) { return k + ' ' + rels[k]; }).join(' / '));
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
    log('アドイン v0.3.2');
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
