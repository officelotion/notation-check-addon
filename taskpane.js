/* 表記チェック Word版 ― アドインの画面とWordの接続 (taskpane.js)  v0.2.1(テスト版) */
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
    function runOf(r) {
      var text = '', bdr = null, shd = null;
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
      }
      return { text: text, boxed: NotationEngine.isBoxedProps(bdr, shd) };
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
      var mt = i < model.length ? norm(NotationEngine.joined(model[i])) : null;
      var wt = j < wp.length ? norm(wp[j].text) : null;
      if (mt !== null && wt !== null && mt === wt) { m2.push(model[i]); w2.push(wp[j]); i++; j++; }
      else if (mt === '') { skipped.push('解析' + (i + 1) + '番目(空)'); i++; }
      else if (wt === '') { skipped.push('Word' + (j + 1) + '番目(空)'); j++; }
      else { bad = { i: i, j: j, mt: mt, wt: wt }; break; }
    }
    return { model: m2, wordParas: w2, skipped: skipped, bad: bad, nModel: model.length, nWord: wp.length, aligned: !bad };
  }
  function explainMisalign(d) {
    var b = d.bad;
    return '段落の照合が、ここで止まりました(Word ' + d.nWord + ' 段落 / 解析 ' + d.nModel + ' 段落)。' +
      '\n  Word ' + (b.j + 1) + '番目 = ' + show(short(b.wt)) + '\n  解析 ' + (b.i + 1) + '番目 = ' + show(short(b.mt)) +
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
    finds.forEach(function (f) { log('  段落' + (f.para + 1) + ' [08 チェックのみ] ' + f.text + '(自動では直しません)'); });
  }

  // ---- STEP 1: チェック ----
  async function doCheck() {
    busy(true); logEl.textContent = ''; status('チェック中…');
    try {
      await Word.run(async function (ctx) {
        var d = await readDocument(ctx);
        if (!d.aligned) log('注意: ' + explainMisalign(d) + '\n「自動修正」は実行できません。\n');
        else if (d.skipped.length) log('(空の段落の差を読み飛ばしました: ' + d.skipped.join(', ') + ')');
        var finds = NotationEngine.findAll(d.model);
        var changes = NotationEngine.fixAll(d.model);
        status('チェック完了: 直す場所 ' + changes.length + ' 件 / 確認だけ ' + finds.length + ' 件');
        log('【直す場所の内訳】\n' + summarize(changes) + '\n\n【一覧】');
        listChanges(changes, finds);
      });
    } catch (e) { status('エラー'); log('エラー: ' + errText(e)); }
    busy(false);
  }

  // ---- 枠・網かけを外す(Wordの種類によって使える命令が違うため、結果を記録する) ----
  async function trySet(ctx, fn) {
    try { fn(); await ctx.sync(); return ''; }
    catch (e) { return errText(e).slice(0, 80); }
  }
  async function clearBox(ctx, rng) {
    var r = {};
    r.border = await trySet(ctx, function () { rng.font.borders.outsideBorderType = 'None'; });
    // 網かけの模様(15%など)は、色を白にしても残るため、模様を「なし」にする
    var t1 = await trySet(ctx, function () { rng.font.shading.texture = 'None'; });
    if (t1) {
      var t2 = await trySet(ctx, function () { rng.font.shading.texture = 'TextureNone'; });
      r.texture = t2 ? ('None: ' + t1 + ' / TextureNone: ' + t2) : '';
    } else r.texture = '';
    r.color = await trySet(ctx, function () { rng.font.shading.backgroundPatternColor = '#FFFFFF'; });
    return r;
  }

  // ---- STEP 2: 置換(変更履歴つき) ----
  // 手順: ①段落ごとに、変える場所を、変更前に全部探す(元の文字が基準なので、変更履歴の「削除された文字」に影響されない)
  //       ②探した場所を、後ろから順に変える ③最後に、文書を読み直して、予定どおりの文字になったか確かめる
  async function doFix() {
    busy(true); logEl.textContent = ''; status('置換中…');
    var okN = 0, ngN = 0, boxStat = { n: 0, border: {}, texture: {}, color: {} };
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
            if (e.ambiguous || e.before.length > 250 || /[\u0000-\u001f]/.test(e.before)) return null;
            var rs = para.search(e.before, { matchCase: true });
            rs.load('items/text');
            return rs;
          });
          await ctx.sync();
          var targets = list.map(function (e, i) {
            if (!searches[i]) return null;
            var cands = searches[i].items.filter(function (r) { return r.text === e.before; });
            return cands[e.occ] || null;
          });
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
                var res = await clearBox(ctx, nr);
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

        // ③ 確かめる(空でない段落の並びを、予定と比べる)
        var L2 = await loadModel(ctx);
        var act = L2.model.map(function (p) { return norm(NotationEngine.joined(p)); }).filter(function (t) { return t; });
        var exp = expected.map(norm).filter(function (t) { return t; });
        var bad = [];
        if (act.length !== exp.length) bad.push('段落の数が変わりました(' + exp.length + ' → ' + act.length + ')');
        else act.forEach(function (t, i) { if (t !== exp[i]) bad.push('段落' + (i + 1) + '(空を除く) 予定=' + show(short(exp[i])) + ' / 実際=' + show(short(t))); });

        status('置換完了: ' + okN + ' 箇所' + (ngN ? ' / 失敗 ' + ngN + ' 箇所' : '') + (bad.length ? ' / 確認で不一致 ' + bad.length : ' / 確認OK'));
        log('\n変更履歴つきで直しました。「校閲」タブの変更履歴で確認できます。');
        log('\n【確認】予定どおりの文字になったか: ' + (bad.length ? '不一致 ' + bad.length + ' 段落' : 'すべて一致'));
        bad.slice(0, 5).forEach(function (b) { log('  ' + b); });
        if (boxStat.n) {
          log('\n【枠・網かけを外す処理】 ' + boxStat.n + ' 箇所');
          [['border', '枠'], ['texture', '網かけの模様'], ['color', '網かけの色']].forEach(function (p) {
            log('  ' + p[1] + ': ' + Object.keys(boxStat[p[0]]).map(function (v) { return v + ' ×' + boxStat[p[0]][v]; }).join(' / '));
          });
          log('  (○以外は、使えなかった命令です。この文字列をそのまま作者に知らせてください)');
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
  async function doUndo(all) {
    busy(true); logEl.textContent = ''; status('元に戻し中…');
    try {
      await Word.run(async function (ctx) {
        var sel = ctx.document.getSelection();
        sel.load('isEmpty');
        var tcs = sel.getTrackedChanges();
        tcs.load('items/date,items/type,items/author');
        await ctx.sync();
        if (sel.isEmpty) { status('範囲を選択してください'); log('元に戻したい範囲を、ドラッグで選択してから押してください。'); return; }
        var since = 0;
        try { since = parseInt(localStorage.getItem(LS_KEY) || '0', 10) || 0; } catch (e) { /* なし */ }
        var items = tcs.items;
        if (!items.length) { status('選択範囲に変更履歴はありません'); return; }
        var mine = (all || !since) ? items : items.filter(function (tc) { return tc.date && tc.date.getTime() >= since - 3000; });
        var other = items.length - mine.length;
        if (!mine.length) {
          status('このアドインの変更は、選択範囲にありません');
          log('選択範囲にある変更履歴 ' + items.length + ' 件は、アドインの修正より前のものです。残しました。\nすべて戻したいときは「選択範囲のすべての変更履歴を戻す」を使ってください。');
          return;
        }
        var kinds = {};
        mine.forEach(function (tc) { kinds[tc.type] = (kinds[tc.type] || 0) + 1; });
        mine.forEach(function (tc) { tc.reject(); });
        await ctx.sync();
        status('選択範囲の変更を ' + mine.length + ' 件、元に戻しました');
        log('内訳: ' + Object.keys(kinds).map(function (k) { return k + ' ' + kinds[k]; }).join(' / '));
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
    log('アドイン v0.2.1');
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
