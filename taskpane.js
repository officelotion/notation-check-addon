/* 表記チェック Word版 ― アドインの画面とWordの接続 (taskpane.js)  v0.1(テスト版) */
/* global Office, Word, NotationEngine */
(function () {
  'use strict';
  var W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var $ = function (id) { return document.getElementById(id); };
  var logEl, statusEl;

  function log(msg) { logEl.textContent += msg + '\n'; logEl.scrollTop = logEl.scrollHeight; }
  function status(msg) { statusEl.textContent = msg; }
  function busy(b) { ['btnCheck', 'btnFix', 'btnUndoSel', 'btnInfo'].forEach(function (id) { $(id).disabled = b; }); }

  // ---- 文書(OOXML) → 段落モデル(engine.js用) ----
  function buildModel(ooxml) {
    var dom = new DOMParser().parseFromString(ooxml, 'application/xml');
    var body = dom.getElementsByTagNameNS(W_NS, 'body')[0];
    if (!body) throw new Error('文書の中身(w:body)が見つかりません');
    var paras = [];
    var SKIP = { Fallback: 1, txbxContent: 1, sectPr: 1, pPr: 1, drawing: 1, pict: 1, AlternateContent: 1 };

    function runOf(r) {
      var text = '', boxed = false;
      for (var c = r.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        var n = c.localName;
        if (n === 'rPr') {
          for (var q = c.firstChild; q; q = q.nextSibling) {
            if (q.nodeType === 1 && (q.localName === 'bdr' || q.localName === 'shd')) boxed = true;
          }
        } else if (n === 't') text += c.textContent;
        else if (n === 'tab') text += '\t';
        else if (n === 'br' || n === 'cr') text += '\v';
      }
      return { text: text, boxed: boxed };
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

  async function readDocument(ctx) {
    var body = ctx.document.body;
    var ox = body.getOoxml();
    var ps = body.paragraphs;
    ps.load('items');
    await ctx.sync();
    var model = buildModel(ox.value);
    return { model: model, wordParas: ps.items, aligned: model.length === ps.items.length };
  }

  function summarize(changes, finds) {
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
  function show(s) { return '「' + String(s).replace(/\v/g, '↵') + '」'; }

  // ---- STEP 1: チェック ----
  async function doCheck() {
    busy(true); logEl.textContent = ''; status('チェック中…');
    try {
      await Word.run(async function (ctx) {
        var d = await readDocument(ctx);
        if (!d.aligned) log('注意: 段落の数が一致しません(文書 ' + d.wordParas.length + ' / 解析 ' + d.model.length + ')。「置換」は実行できません。');
        var finds = NotationEngine.findAll(d.model);
        var changes = NotationEngine.fixAll(d.model);
        status('チェック完了: 直す場所 ' + changes.length + ' 件 / 確認だけ ' + finds.length + ' 件');
        log('【直す場所の内訳】\n' + summarize(changes, finds) + '\n\n【一覧】');
        listChanges(changes, finds);
      });
    } catch (e) { status('エラー'); log('エラー: ' + e.message); }
    busy(false);
  }

  // ---- STEP 2: 置換(変更履歴つき) ----
  async function tryClearBox(ctx, range) {
    // 枠・網かけを外す(Wordの種類によって使える命令が違うため、順に試す)
    var tried = [];
    try {
      range.font.borders.outsideBorderType = 'None';
      await ctx.sync(); tried.push('枠');
    } catch (e) { tried.push('枠×(' + (e.code || e.message) + ')'); }
    try {
      range.font.shading.backgroundPatternColor = '#FFFFFF';
      await ctx.sync(); tried.push('網かけ');
    } catch (e) { tried.push('網かけ×(' + (e.code || e.message) + ')'); }
    return tried;
  }

  async function doFix() {
    busy(true); logEl.textContent = ''; status('置換中…');
    var ok = 0, ng = 0, boxMsg = null;
    try {
      await Word.run(async function (ctx) {
        var d = await readDocument(ctx);
        if (!d.aligned) { status('中止'); log('段落の数が一致しないため、置換を中止しました。(文書 ' + d.wordParas.length + ' / 解析 ' + d.model.length + ')\n文書の画像・テキストボックスなどが原因かもしれません。この文書のまま、作者に知らせてください。'); return; }
        var changes = NotationEngine.fixAll(d.model);
        if (!changes.length) { status('直す場所はありませんでした'); return; }

        ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        await ctx.sync();

        // 同じ段落・同じルールの変更は、後ろから適用する(位置がずれないように)
        var groups = [], cur = null;
        changes.forEach(function (c) {
          if (cur && cur.para === c.para && cur.rule === c.rule) cur.items.push(c);
          else { cur = { para: c.para, rule: c.rule, items: [c] }; groups.push(cur); }
        });
        for (var g = 0; g < groups.length; g++) {
          var items = groups[g].items.slice().reverse();
          for (var k = 0; k < items.length; k++) {
            var c = items[k], para = d.wordParas[c.para];
            var rs = para.search(c.before, { matchCase: true });
            rs.load('items/text');
            await ctx.sync();
            var cands = rs.items.filter(function (r) { return r.text === c.before; });
            var target = cands[c.occ];
            if (!target) { ng++; log('  × 段落' + (c.para + 1) + ' [' + c.rule + '] ' + show(c.before) + ' が見つかりません'); continue; }
            target.insertText(c.after, 'Replace');
            await ctx.sync();
            if (c.box) {
              var inserted = target; // 置換後の範囲
              var res = await tryClearBox(ctx, inserted);
              if (boxMsg === null) boxMsg = res.join(' / ');
            }
            ok++;
          }
        }
        status('置換完了: ' + ok + ' 件' + (ng ? ' / 見つからず ' + ng + ' 件' : ''));
        log('\n変更履歴つきで直しました。「校閲」タブの変更履歴で確認できます。');
        if (boxMsg) log('枠・網かけを外す処理の結果: ' + boxMsg + '\n(×がある場合は、枠が残ることがあります。作者に知らせてください)');
      });
    } catch (e) { status('エラー'); log('エラー: ' + e.message + (e.debugInfo ? '\n' + JSON.stringify(e.debugInfo) : '')); }
    busy(false);
  }

  // ---- STEP 3: 選択範囲の変更を元に戻す ----
  async function doUndoSel() {
    busy(true); logEl.textContent = ''; status('元に戻し中…');
    try {
      await Word.run(async function (ctx) {
        var sel = ctx.document.getSelection();
        var tcs = sel.getTrackedChanges();
        tcs.load('items');
        await ctx.sync();
        var n = tcs.items.length;
        if (!n) { status('選択範囲に変更履歴はありません'); return; }
        tcs.rejectAll();
        await ctx.sync();
        status('選択範囲の変更を ' + n + ' 件、元に戻しました');
      });
    } catch (e) { status('エラー'); log('エラー: ' + e.message + '\n(選択範囲の取り消しに対応していないWordかもしれません)'); }
    busy(false);
  }

  // ---- テスト情報 ----
  function doInfo() {
    logEl.textContent = '';
    var r = Office.context.requirements;
    function sup(set, v) { try { return r.isSetSupported(set, v) ? '○' : '×'; } catch (e) { return '?'; } }
    log('プラットフォーム: ' + (Office.context.platform || '?'));
    try { log('Word: ' + (Office.context.diagnostics ? Office.context.diagnostics.version : '?')); } catch (e) {}
    log('対応状況(○=使える)');
    [['WordApi', '1.3'], ['WordApi', '1.4'], ['WordApi', '1.6'], ['WordApiDesktop', '1.1'], ['WordApiDesktop', '1.2']].forEach(function (x) {
      log('  ' + x[0] + ' ' + x[1] + ' : ' + sup(x[0], x[1]));
    });
    log('\nWordApi 1.4 = 変更履歴 / 1.6 = 選択範囲の元に戻す / Desktop 1.1 = 枠・網かけ');
    status('テスト情報を表示しました');
  }

  Office.onReady(function (info) {
    logEl = $('log'); statusEl = $('status');
    if (info.host !== Office.HostType.Word) { status('このアドインはWordで使います'); return; }
    $('btnCheck').onclick = doCheck;
    $('btnFix').onclick = doFix;
    $('btnUndoSel').onclick = doUndoSel;
    $('btnInfo').onclick = doInfo;
    status('準備できました。まず「チェックする」を押してください');
  });
})();
