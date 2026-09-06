/*******************************************************************
 *  ALL INVESTMENTS  —  mobile app API  (Google Apps Script)
 *
 *  HOW TO USE (no coding needed):
 *   1. In your "All Investments" Google Sheet:  Extensions ▸ Apps Script
 *   2. Click  +  next to "Files" ▸ Script, name it  API  and paste this file.
 *   3. Set your passcode once — see the note below the header.
 *   4. Deploy ▸ New deployment ▸ type "Web app"
 *        Execute as:            Me
 *        Who has access:        Anyone
 *      Copy the Web app URL — you will paste it into the phone app.
 *
 *  Every time you change this file you must  Deploy ▸ Manage deployments ▸
 *  ✎ (edit) ▸ Version: New version ▸ Deploy  for the change to go live.
 *******************************************************************/

/*  YOUR PASSCODE IS NOT IN THIS FILE.
 *
 *  It lives in the project's stored settings, so pasting a new version of
 *  this file can never wipe it. To set or change it, either:
 *
 *    a) run  SET_PASSCODE('your-passcode')  once from the editor, or
 *    b) Project Settings (gear, left sidebar) ▸ Script Properties ▸
 *       Add script property ▸ name  PASSCODE  ▸ value  your-passcode
 *
 *  Run  SHOW_PASSCODE_STATUS()  any time to check one is set.
 */
function getPasscode_() {
  var v = PropertiesService.getScriptProperties().getProperty('PASSCODE');
  return v ? String(v).trim() : '';
}

// Convenience: run this once from the editor to set it.
function SET_PASSCODE(value) {
  if (!value || String(value).trim().length < 4) {
    throw new Error('Pass a passcode of at least 4 characters, e.g. SET_PASSCODE("tiger-4471")');
  }
  PropertiesService.getScriptProperties().setProperty('PASSCODE', String(value).trim());
  Logger.log('Passcode saved. It will survive future updates to this file.');
}

// Tells you whether a passcode is set, without printing it.
function SHOW_PASSCODE_STATUS() {
  var p = getPasscode_();
  Logger.log(p ? 'A passcode is set (' + p.length + ' characters). Nothing to do.'
               : 'No passcode set yet. Run SET_PASSCODE("something") once.');
}

// Sheet / range configuration.  Only touch this if you rename tabs.
var CFG = {
  netWorth:      'Net Worth',
  breakdown:     'Breakdown',
  buckets:       'Buckets',
  bucketLedger:  'Bucket Ledger',
  daily:         'Daily Change',
  monthly:       'Monthly Gains',
  transactions:  'Transactions',
  masterList:    'Master List',
  fxRates:       'FX Rates',
  platforms:     ['INDMoney','Kite','Vested','Coin','Groww','Binance','Coinswitch','CoinDCX'],
  categories:    ['US Stocks','Stocks','Mutual Fund','Gold/Silver','Crypto','Charges'],
  usdCategories: ['US Stocks'],             // categories whose Amount is in USD
  txnHistory:    400                        // how many recent transactions to send
};

// Bumped with every change. The app shows it, so we can always tell whether the
// deployed version is the one you last pasted in — saving the file is not enough,
// the deployment has to be republished as a New version.
var SCRIPT_BUILD = '2026-09-06f';

/* ============================ ENTRY POINTS ============================ */

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    auth_(p.key);
    var action = p.action || 'all';
    var out;
    if (action === 'all')        out = getAll_();
    else if (action === 'txns')  out = { transactions: getTransactions_(Number(p.limit) || CFG.txnHistory) };
    else if (action === 'ping')  out = { ok: true };
    else throw new Error('Unknown action: ' + action);
    return json_({ ok: true, data: out, scriptBuild: SCRIPT_BUILD, serverTime: new Date().toISOString() });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err), scriptBuild: SCRIPT_BUILD });
  }
}

function doPost(e) {
  var lock = null;
  var params = (e && e.parameter) || {};
  var wantsText = String(params.format || '').toLowerCase() === 'text';
  try {
    var raw = (e && e.postData && e.postData.contents) || '';
    var body = null;
    try { body = JSON.parse(raw); } catch (parseErr) { body = null; }

    // The app sends JSON. The iOS Shortcut can't easily build JSON around OCR
    // text (newlines and quotes would break it), so it instead puts key and
    // action in the query string and posts the raw text as the body.
    if (!body || typeof body !== 'object') {
      body = { key: params.key, action: params.action, data: { text: raw } };
    }
    if (!body.action && params.action) body.action = params.action;
    if (!body.key && params.key) body.key = params.key;

    auth_(body.key);

    // Read-only actions: no lock needed, and they must not block writes.
    if (body.action === 'parseText') {
      var pt = tryParse_(body.data && body.data.text);
      return wantsText ? text_(describeParse_(pt)) : json_({ ok: true, data: pt });
    }
    if (body.action === 'parseImage') {
      var d = body.data || {};
      // One contract note can need several screenshots — the charges are at the
      // top and do not stay on screen as you scroll. They are read in order and
      // treated as one note; overlapping trades are counted once.
      var shots = d.images && d.images.length ? d.images : [{ image: d.image, mimeType: d.mimeType }];
      if (shots.length > 8) throw new Error('That is more than 8 screenshots — send them in two goes.');
      var texts = [];
      for (var si = 0; si < shots.length; si++) {
        texts.push(ocrImage_(shots[si].image, shots[si].mimeType));
      }
      var pi = tryParse_(texts.join('\n'));
      return wantsText ? text_(describeParse_(pi)) : json_({ ok: true, data: pi });
    }

    lock = LockService.getScriptLock();
    lock.waitLock(20000);
    var result;
    if (body.action === 'addTxn')         result = addTransaction_(body.data || {});
    else if (body.action === 'editTxn')   result = editTransaction_(body.data || {});
    else if (body.action === 'deleteTxn') result = deleteTransaction_(body.data || {});
    else if (body.action === 'addFromText') {
      // One-shot for the Shortcut: read the text, then write the row(s).
      // A Kite contract note can hold several trades; an INDmoney screen holds one.
      var parsed = parseAny_(body.data && body.data.text);
      result = saveParsed_(parsed, params.bucket || '');
    }
    else throw new Error('Unknown action: ' + body.action);

    // Sheets applies queued operations here, so this is where a deferred
    // failure appears — possibly one that has nothing to do with the data,
    // long after the row itself landed. Check the row before calling a save
    // that worked a failure.
    try {
      SpreadsheetApp.flush();
    } catch (flushErr) {
      if (!rowLooksSaved_(body.action, result, body.data || {})) throw flushErr;
      result.warning = String(flushErr.message || flushErr);
    }
    return wantsText ? text_(result.summary || 'Saved.')
                     : json_({ ok: true, data: result, scriptBuild: SCRIPT_BUILD });
  } catch (err) {
    var msg = String(err.message || err);
    return wantsText ? text_('Could not log it.\n' + msg)
                     : json_({ ok: false, error: msg, scriptBuild: SCRIPT_BUILD });
  } finally {
    try { if (lock) lock.releaseLock(); } catch (_) {}
  }
}

// Writes whatever the parser produced — one row for an INDmoney screen, one
// per equity trade for a Kite note — and returns a summary for the Shortcut.
function saveParsed_(parsed, bucket) {
  if (parsed.kind === 'kiteNote') {
    if (!parsed.trades.length) {
      throw new Error('Nothing to save' +
        (parsed.skipped.length ? ' — the note only had F&O trades.' : '.'));
    }
    var sh = sheet_(CFG.transactions);
    var rows = [], names = [], already = [];
    for (var i = 0; i < parsed.trades.length; i++) {
      var tr = parsed.trades[i];
      // Sharing a second screenshot of the same note must not re-save the
      // trades the first one already wrote.
      if (existingRowFor_(sh, tr)) { already.push(tr.ticker); continue; }
      tr.bucket = bucket;
      var r = addTransaction_(tr);
      rows.push(r.row);
      names.push(tr.ticker);
    }
    var msg = rows.length
      ? 'Saved ' + rows.length + ' row' + (rows.length > 1 ? 's' : '') +
        ' (' + names.join(', ') + ') at ' + rows.join(', ') + '.'
      : 'Nothing new to save.';
    if (already.length) {
      msg += '\nSkipped ' + already.join(', ') + ' — already in the sheet.';
    }
    return { row: rows.length ? rows[0] : 0, rows: rows, skippedExisting: already,
             summary: describeParse_(parsed) + '\n' + msg };
  }

  if (parsed.amount == null) {
    throw new Error('Could not read the amount' +
      (parsed.order === 'Sell' ? ' (no "Credit to Wallet" row)' : '') +
      '. Add this one in the app instead.');
  }
  parsed.bucket = bucket;
  var one = addTransaction_(parsed);
  one.summary = describeParse_(parsed) + '\nSaved to row ' + one.row + '.';
  return one;
}

// A short human-readable summary, for the Shortcut's "Show Result".
function describeParse_(p) {
  if (p && p.kind === 'kiteNote') return describeKiteNote_(p);
  if (!p || p.parseError) {
    return 'Could not read that screenshot.\n' + ((p && p.parseError) || '');
  }
  // Prices carry 3 decimals on INDmoney — rounding them here would make the
  // preview disagree with what actually gets written.
  var money = function (v, dp) {
    if (v == null) return '—';
    var s = Number(v).toFixed(dp == null ? 2 : dp);
    return '$' + (dp ? s.replace(/(\.\d*?)0+$/, '$1').replace(/\.$/, '') : s);
  };
  var lines = [
    p.ticker + '  ·  ' + String(p.order).toUpperCase(),
    p.volume + ' @ ' + money(p.price, 4),
    (p.order === 'Sell' ? 'Credited ' : 'Charged ') + money(p.amount),
    p.date || 'no date'
  ];
  (p.warnings || []).forEach(function (w) { lines.push('! ' + w); });
  return lines.join('\n');
}

function describeKiteNote_(p) {
  var rupee = function (v) { return v == null ? '—' : '\u20b9' + Number(v).toFixed(2); };
  var lines = [];
  for (var i = 0; i < p.trades.length; i++) {
    var t = p.trades[i];
    lines.push(t.ticker + '  \u00b7  ' + t.order.toUpperCase());
    lines.push('  ' + t.volume + ' @ ' + rupee(t.price) +
               '  =  ' + rupee(t.amount) +
               '   fees ' + rupee(t.fees) +
               (t.dp != null && t.dp !== '' ? ' + DP ' + rupee(t.dp) : ''));
  }
  if (!p.trades.length) lines.push('No equity trades on this note.');
  (p.warnings || []).forEach(function (w) { lines.push('! ' + w); });
  return lines.join('\n');
}

/* ============================== READERS ============================== */

function getAll_() {
  return {
    netWorth:     getNetWorth_(),
    breakdown:    getBreakdown_(),
    buckets:      getBuckets_(),
    daily:        getDaily_(),
    monthly:      getMonthly_(),
    transactions: getTransactions_(CFG.txnHistory),
    lists:        getLists_()
  };
}

// Net Worth tab: the asset table (A:L) plus headline cells found by label,
// so the app keeps working if rows move around.
function getNetWorth_() {
  var sh = sheet_(CFG.netWorth);
  var vals = sh.getRange(1, 1, 30, 30).getValues();     // A1:AD30 (labels live here)

  var assets = [];
  for (var r = 1; r < vals.length; r++) {
    var label = String(vals[r][0] || '').trim();
    if (!label) continue;
    if (/^Total/i.test(label)) {
      assets.push(rowObj_(vals[r], label, true));
      break;
    }
    assets.push(rowObj_(vals[r], label, false));
  }
  function rowObj_(row, label, isTotal) {
    return {
      name: label, isTotal: isTotal,
      invested: num_(row[1]), dayChange: num_(row[2]), dayChangePct: num_(row[3]),
      value: num_(row[4]), currReturns: num_(row[5]), currPnlPct: num_(row[6]),
      realized: num_(row[7]), totalReturns: num_(row[8]), totalPnlPct: num_(row[9]),
      xirr: num_(row[10]), allocation: num_(row[11])
    };
  }

  // Headline numbers by label lookup
  var head = {};
  var nw = findLabel_(vals, /^Networth/i);              // N17 style: label in col N, value O, change Q, pct R
  if (nw) {
    head.netWorth      = num_(vals[nw.r][nw.c + 1]);
    head.netWorthDay   = num_(vals[nw.r][nw.c + 3]);
    head.netWorthDayPct= num_(vals[nw.r][nw.c + 4]);
  }
  var inv = findLabel_(vals, /^Investment Value/i);     // col A label, value in E
  if (inv) head.investmentValue = num_(vals[inv.r][4]);
  var mi = findLabel_(vals, /^Money In/i);
  if (mi) head.moneyIn = num_(vals[mi.r][4]);
  var bf = findLabel_(vals, /^Blocked Funds/i);
  if (bf) head.blockedFunds = num_(vals[bf.r][4]);
  var va = findLabel_(vals, /^Virtual Available/i);
  if (va) head.virtualAvailable = num_(vals[va.r][4]);
  var rb = findLabel_(vals, /^Running Balance/i);
  if (rb) head.runningBalance = num_(vals[rb.r][rb.c + 1]);
  var ind = findLabel_(vals, /^INDmoney \$/i);
  if (ind) { head.indWalletLabel = String(vals[ind.r][0]); head.indWalletInr = num_(vals[ind.r][4]); }
  var kw = findLabel_(vals, /^Kite Wallet/i);
  if (kw) head.kiteWallet = num_(vals[kw.r][4]);
  var fx = findLabel_(vals, /^INDmoney \$/i);
  if (fx) head.usdInr = num_(vals[fx.r][3]);

  // Bank balances: header "Bank"/"Amount" somewhere in row 1 (col N/O)
  var banks = [];
  var bh = findLabel_([vals[0]], /^Bank$/i);
  if (bh) {
    for (var i = 1; i < vals.length; i++) {
      var b = String(vals[i][bh.c] || '').trim(), a = vals[i][bh.c + 1];
      if (!b) break;
      if (/Running Balance|Networth/i.test(b)) continue;
      banks.push({ name: b.replace(/^🔒\s*/, ''), locked: /^🔒/.test(b), amount: num_(a) });
    }
  }

  // Monthly history tables: "Month"/"Net worth" and "Month"/"Investment"
  var monthlyNW = twoColSeries_(vals, /^Net worth$/i), monthlyInv = twoColSeries_(vals, /^Investment$/i);

  return { assets: assets, head: head, banks: banks, monthlyNetWorth: monthlyNW, monthlyInvestment: monthlyInv };
}

function twoColSeries_(vals, valueHeaderRe) {
  var h = findLabel_([vals[0]], valueHeaderRe);
  if (!h) return [];
  var out = [];
  for (var i = 1; i < vals.length; i++) {
    var m = vals[i][h.c - 1], v = vals[i][h.c];
    if (m === '' || m === null || m === undefined) break;
    out.push({ month: String(m), value: num_(v) });
  }
  return out;
}

function getBreakdown_() {
  var sh = sheet_(CFG.breakdown);
  var vals = sh.getDataRange().getValues();
  var rows = [];
  for (var r = 1; r < vals.length; r++) {
    if (!vals[r][0]) continue;
    rows.push({ type: String(vals[r][0]), allocation: num_(vals[r][1]), value: num_(vals[r][2]),
      invested: num_(vals[r][3]), realized: num_(vals[r][4]), currPnl: num_(vals[r][5]),
      currPnlPct: num_(vals[r][6]), totalPnl: num_(vals[r][7]) });
  }
  return rows;
}

function getBuckets_() {
  var sh = sheet_(CFG.buckets);
  var vals = sh.getRange(1, 1, Math.max(2, sh.getLastRow()), 13).getValues();  // A:M
  var buckets = [];
  for (var r = 1; r < vals.length; r++) {
    if (!vals[r][0]) continue;
    if (isBucketListEnd_(vals[r][0])) break;          // the summary/reconciliation blocks below the list
    buckets.push({ name: String(vals[r][0]), targetPct: num_(vals[r][1]), notes: String(vals[r][2] || ''),
      excluded: vals[r][3] === true, allocation: num_(vals[r][5]), value: num_(vals[r][6]),
      invested: num_(vals[r][7]), investedLiveFx: num_(vals[r][8]), realized: num_(vals[r][9]),
      currPnl: num_(vals[r][10]), currPnlPct: num_(vals[r][11]), totalPnl: num_(vals[r][12]) });
  }
  // Per-holding ledger (Bucket Ledger A:J)
  var holdings = [];
  var ls = sheetOrNull_(CFG.bucketLedger);
  if (ls && ls.getLastRow() > 1) {
    var lv = ls.getRange(2, 1, ls.getLastRow() - 1, 10).getValues();
    lv.forEach(function (row) {
      if (!row[0]) return;
      holdings.push({ ticker: String(row[0]), platform: String(row[1] || ''), bucket: String(row[2] || ''),
        category: String(row[3] || ''), qty: num_(row[4]), invested: num_(row[5]), realized: num_(row[6]),
        value: num_(row[7]), investedNative: num_(row[8]), excluded: row[9] === true });
    });
  }
  return { buckets: buckets, holdings: holdings };
}

function getDaily_() {
  var sh = sheet_(CFG.daily);
  var last = sh.getLastRow();
  var tz = tz_();
  var a = sh.getRange(2, 1, last - 1, 7).getValues();          // A:G  date,value,in/out, _, date,value,payin
  var b = sh.getRange(2, 62, last - 1, 6).getValues();         // BJ:BO date,invested,currVal,currRet,realized,totRet
  var out = [];
  for (var i = 0; i < a.length; i++) {
    if (!(a[i][0] instanceof Date)) continue;
    out.push({ date: fmtDate_(a[i][0], tz), netWorth: num_(a[i][1]), cashFlow: num_(a[i][2]),
      investment: num_(a[i][5]), payInOut: num_(a[i][6]),
      invested: num_(b[i] && b[i][1]), currValue: num_(b[i] && b[i][2]), currReturns: num_(b[i] && b[i][3]),
      realized: num_(b[i] && b[i][4]), totalReturns: num_(b[i] && b[i][5]) });
  }
  return out;
}

function getMonthly_() {
  var sh = sheet_(CFG.monthly);
  var vals = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 12).getValues();
  var tz = tz_(), out = [];
  vals.forEach(function (r) {
    if (!(r[0] instanceof Date)) return;
    out.push({ monthEnd: fmtDate_(r[0], tz), month: String(r[1]), opening: num_(r[2]), closing: num_(r[3]),
      netContribution: num_(r[4]), gain: num_(r[5]), realized: num_(r[6]), unrealized: num_(r[7]),
      returnPct: num_(r[8]), source: String(r[9] || '') });
  });
  return out;
}

// Transactions: every row that has a date and is a real transaction
// (the "... VALUE" pseudo-rows used for XIRR are excluded).  Each row carries
// its sheet row number so the app can edit it in place.
function getTransactions_(limit) {
  var sh = sheet_(CFG.transactions);
  var last = sh.getLastRow();
  var vals = sh.getRange(2, 1, last - 1, 17).getValues();      // A:Q
  var tz = tz_(), out = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    if (!(r[1] instanceof Date)) continue;
    if (/\bVALUE$/i.test(String(r[2] || ''))) continue;
    out.push({ row: i + 2, sno: num_(r[0]), date: fmtDate_(r[1], tz), ticker: String(r[2] || ''),
      category: String(r[3] || ''), order: String(r[4] || ''), platform: String(r[5] || ''),
      amount: num_(r[6]), notional: num_(r[7]), volume: num_(r[8]), price: num_(r[9]),
      dp: num_(r[10]), fees: num_(r[11]), remark: String(r[12] || ''), inr: num_(r[13]),
      fx: num_(r[15]), bucket: String(r[16] || '') });
  }
  out = out.slice(Math.max(0, out.length - limit)).reverse();      // newest row first
  out.sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : b.row - a.row; });
  return out;
}

function getLists_() {
  var tickers = [];
  var ml = sheetOrNull_(CFG.masterList);
  if (ml) {
    ml.getRange(2, 2, Math.max(1, ml.getLastRow() - 1), 4).getValues().forEach(function (r) {
      var t = String(r[0] || '').trim();
      if (t && !/\bVALUE$/i.test(t)) tickers.push({ ticker: t, category: String(r[3] || '') });
    });
  }
  var buckets = [];
  var bs = sheetOrNull_(CFG.buckets);
  if (bs) {
    var bv = bs.getRange(2, 1, Math.max(1, bs.getLastRow() - 1), 1).getValues();
    for (var i = 0; i < bv.length; i++) {
      if (!bv[i][0]) continue;
      if (isBucketListEnd_(bv[i][0])) break;
      buckets.push(String(bv[i][0]));
    }
  }
  return { tickers: tickers, buckets: buckets, platforms: CFG.platforms, categories: CFG.categories, orders: ['Buy', 'Sell'] };
}

/* ============================== WRITERS ============================== */

// Adds a transaction just below the last real transaction (before the
// "... VALUE" rows), reusing an empty pre-formatted row if one exists.
function addTransaction_(d) {
  validateTxn_(d);
  var sh = sheet_(CFG.transactions);
  var last = sh.getLastRow();
  var colB = sh.getRange(1, 2, last, 2).getValues();          // B:C
  var lastReal = 1, firstValue = 0;
  for (var i = 1; i < colB.length; i++) {
    var isDate = colB[i][0] instanceof Date, isValue = /\bVALUE$/i.test(String(colB[i][1] || ''));
    if (isDate && !isValue) lastReal = i + 1;
    if (isDate && isValue && !firstValue) firstValue = i + 1;
  }
  var target = lastReal + 1;
  var blankBetween = !firstValue || target < firstValue;
  if (!blankBetween) {                    // no spare row – insert one before the VALUE rows
    sh.insertRowBefore(firstValue);
    target = firstValue;
  } else if (sh.getRange(target, 2).getValue() !== '' && sh.getRange(target, 2).getValue() !== null) {
    sh.insertRowAfter(lastReal);
    target = lastReal + 1;
  }
  var prevSno = num_(sh.getRange(lastReal, 1).getValue()) || 0;
  writeTxnRow_(sh, target, d, prevSno + 1);
  // Renumbering is tidy-up after the row already exists. On a Sheets Table the
  // SNo. column may refuse a write; that must not report the save as failed.
  try { renumberBelow_(sh, target); } catch (renumErr) {}
  return { row: target, sno: prevSno + 1 };
}

function editTransaction_(d) {
  var row = Number(d.row);
  if (!row || row < 2) throw new Error('Missing row number');
  validateTxn_(d);
  var sh = sheet_(CFG.transactions);
  var cur = sh.getRange(row, 1, 1, 3).getValues()[0];
  if (!(cur[1] instanceof Date)) throw new Error('Row ' + row + ' is not a transaction any more — refresh and try again.');
  if (/\bVALUE$/i.test(String(cur[2] || ''))) throw new Error('That row is a system row and cannot be edited.');
  writeTxnRow_(sh, row, d, num_(cur[0]));
  return { row: row };
}

function deleteTransaction_(d) {
  var row = Number(d.row);
  if (!row || row < 2) throw new Error('Missing row number');
  var sh = sheet_(CFG.transactions);
  var cur = sh.getRange(row, 1, 1, 3).getValues()[0];
  if (/\bVALUE$/i.test(String(cur[2] || ''))) throw new Error('That row is a system row and cannot be deleted.');
  if (String(d.ticker || '') && String(cur[2]) !== String(d.ticker)) throw new Error('Row changed since you loaded it — refresh and try again.');
  sh.deleteRow(row);
  try { renumberBelow_(sh, row); } catch (renumErr) {}
  return { row: row };
}

// Reads the row back and says whether the transaction is actually in the sheet.
// Used to tell a real write failure apart from a cosmetic one that surfaced late.
function rowLooksSaved_(action, result, d) {
  if (action === 'deleteTxn') return true;
  if (!result || !result.row) return false;
  try {
    var sh = sheet_(CFG.transactions);
    var cells = sh.getRange(result.row, 2, 1, 2).getValues()[0];   // B date, C ticker
    if (!(cells[0] instanceof Date)) return false;
    var want = String((d && d.ticker) || (result && result.ticker) || '').trim().toUpperCase();
    if (want && String(cells[1]).trim().toUpperCase() !== want) return false;
    return true;
  } catch (readErr) {
    return false;      // can't confirm it, so don't claim it
  }
}

// Is this exact trade already in the sheet? Two screenshots of one contract
// note, shared separately, would otherwise write the overlapping rows twice —
// and a double-counted trade quietly corrupts the XIRR columns.
// Matched on date + ticker + side + quantity + price + platform.
function existingRowFor_(sh, d) {
  var last = sh.getLastRow();
  var span = Math.min(200, last - 1);
  if (span <= 0) return 0;
  var start = last - span + 1;
  var vals = sh.getRange(start, 2, span, 9).getValues();       // B..J
  var want = String(d.ticker || '').trim().toUpperCase();
  var vol  = numOrBlank_(d.volume), price = numOrBlank_(d.price);
  if (!want || vol === '' || price === '') return 0;

  var near = function (a, b) {
    if (a === '' || b === '') return false;
    return Math.abs(Number(a) - Number(b)) <= Math.max(1e-6, Math.abs(Number(b)) * 1e-6);
  };
  for (var i = vals.length - 1; i >= 0; i--) {
    var r = vals[i];
    if (!(r[0] instanceof Date)) continue;
    if (Utilities.formatDate(r[0], tz_(), 'yyyy-MM-dd') !== d.date) continue;
    if (String(r[1]).trim().toUpperCase() !== want) continue;
    if (String(r[3]) !== String(d.order)) continue;
    if (String(r[4] || '').toLowerCase() !== String(d.platform || '').toLowerCase()) continue;
    if (!near(r[7], vol) || !near(r[8], price)) continue;
    return start + i;
  }
  return 0;
}

function validateTxn_(d) {
  if (!d.date || !/^\d{4}-\d{2}-\d{2}$/.test(d.date)) throw new Error('Date must be YYYY-MM-DD');
  if (!String(d.ticker || '').trim()) throw new Error('Ticker is required');
  if (['Buy', 'Sell'].indexOf(d.order) < 0) throw new Error('Order must be Buy or Sell');
  if (!d.category) throw new Error('Category is required');
  var hasAmount = d.amount !== '' && d.amount !== null && d.amount !== undefined;
  var hasVP = d.volume !== '' && d.volume != null && d.price !== '' && d.price != null;
  if (!hasAmount && !hasVP) throw new Error('Enter an Amount, or Volume and Unit Price');
}

// Writes one transaction row A:Q mirroring how the sheet is filled by hand:
//   G Amount        – typed value, or =I*J when left blank
//   H Notional      – typed value, or =I*J
//   L Fees          – typed value, or the INDmoney GOOGLEFINANCE formula
//   N/O/P           – ± INR value, its negative, and the FX rate of that date
function writeTxnRow_(sh, row, d, sno) {
  var parts = d.date.split('-');
  var date = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  var isUsd = CFG.usdCategories.indexOf(d.category) >= 0;
  var fx = fxRateFor_(date);

  var vol = numOrBlank_(d.volume), price = numOrBlank_(d.price);
  var amountCell   = numOrBlank_(d.amount);
  if (amountCell === '' && vol !== '' && price !== '') amountCell = '=I' + row + '*J' + row;
  var notionalCell = numOrBlank_(d.notional);
  if (notionalCell === '') notionalCell = '=I' + row + '*J' + row;
  // Fees behave differently per platform, so proximity is the wrong signal:
  //   you typed one        -> write exactly that, whatever the platform
  //   INDmoney, none typed -> derived from G-H, so use the formula
  //   any other platform   -> you enter it by hand, so leave the cell empty
  //                           rather than inheriting a neighbour's formula
  var feesCell = numOrBlank_(d.fees);
  var feesFrom = 0;
  if (feesCell === '' && /^INDmoney$/i.test(d.platform || '')) {
    // Prefer the formula an existing INDmoney row already uses, so a change you
    // make to it carries forward. Only fall back to our own copy if none exists.
    feesFrom = lastFeeFormulaRow_(sh, row, d.platform);
    if (!feesFrom) {
      feesCell = '=IFERROR(IF(F' + row + '="INDmoney", IF(E' + row + '="sell",-1,1)*(G' + row + '-H' + row + ')*GOOGLEFINANCE("CURRENCY:USDINR"),""),"")';
    }
  }

  // INR value of the trade for XIRR/SIP columns
  var amountNum = numOrBlank_(d.amount);
  if (amountNum === '' && vol !== '' && price !== '') amountNum = vol * price;
  var sign = d.order === 'Buy' ? -1 : 1;
  var inr = amountNum === '' ? '' : sign * amountNum * (isUsd ? fx : 1);

  var values = [
    sno, date, String(d.ticker).trim().toUpperCase(), d.category, d.order, d.platform || '',
    amountCell, notionalCell, vol, price, numOrBlank_(d.dp), feesCell, d.remark || '',
    inr, inr === '' ? '' : -inr, fx, d.bucket || ''
  ];

  // Columns filled by an ARRAYFORMULA (SNo., For XIRR, For SIP) are left empty —
  // writing a literal into one makes Sheets complain and breaks the formula for
  // the whole column.
  var skip = {};
  var generated = generatedColumns_(sh);
  for (var g in generated) if (generated.hasOwnProperty(g)) skip[g] = true;
  if (feesFrom) skip[12] = true;          // the Fees formula is copied in below

  // Values first, so any copied formula calculates against them.
  writeSkipping_(sh, row, values, skip);

  if (feesFrom && !generated[12]) {
    try {
      sh.getRange(feesFrom, 12)
        .copyTo(sh.getRange(row, 12), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
    } catch (copyErr) {
      // If the copy is refused, write our own rather than leaving Fees empty.
      sh.getRange(row, 12).setValue(
        '=IFERROR(IF(F' + row + '="INDmoney", IF(E' + row + '="sell",-1,1)*(G' + row + '-H' + row + ')*GOOGLEFINANCE("CURRENCY:USDINR"),""),"")');
    }
  }

  // NOTHING may be formatted here. The Transactions range is a Sheets Table,
  // and a typed column rejects formatting with "You can't set the number
  // format of cells in a typed column" — and because Apps Script queues these
  // operations and only reports them when it flushes, a try/catch around the
  // call does not catch it. A typed date column formats itself anyway.
}

// The nearest row above that is the SAME platform and holds a formula in Fees.
// Matching the platform is the point: a Kite row's hand-typed fee must never
// become an INDmoney row's formula, or the other way round.
function lastFeeFormulaRow_(sh, beforeRow, platform) {
  var span = Math.min(400, beforeRow - 2);
  if (span <= 0) return 0;
  var start = beforeRow - span;
  var plats = sh.getRange(start, 6, span, 1).getValues();      // F Platform
  var forms = sh.getRange(start, 12, span, 1).getFormulas();   // L Fees
  var want = String(platform || '').toLowerCase();
  for (var i = span - 1; i >= 0; i--) {
    if (forms[i][0] && String(plats[i][0] || '').toLowerCase() === want) return start + i;
  }
  return 0;
}

// Writes the row in contiguous runs, stepping over generated columns.
function writeSkipping_(sh, row, values, generated) {
  var start = -1;
  for (var c = 0; c <= values.length; c++) {
    var skip = (c === values.length) || generated[c + 1];
    if (skip) {
      if (start >= 0) {
        var chunk = values.slice(start, c);
        sh.getRange(row, start + 1, 1, chunk.length).setValues([chunk]);
        start = -1;
      }
    } else if (start < 0) {
      start = c;
    }
  }
}

// Which columns are produced by a formula rather than typed in. Detected by
// looking for an ARRAYFORMULA in the top rows of each column, so it adapts if
// you convert another column later — nothing here is hard-coded to A, N or O.
function generatedColumns_(sh) {
  if (sh.__generated) return sh.__generated;
  var out = {};
  try {
    var f = sh.getRange(1, 1, 3, 20).getFormulas();
    for (var r = 0; r < f.length; r++) {
      for (var c = 0; c < f[r].length; c++) {
        if (/^=\s*ARRAYFORMULA/i.test(String(f[r][c] || ''))) out[c + 1] = true;
      }
    }
  } catch (e) { /* if this fails, write everything as before */ }
  sh.__generated = out;
  return out;
}

function renumberBelow_(sh, fromRow) {
  // Nothing to do when SNo. is generated by a formula — it renumbers itself.
  if (generatedColumns_(sh)[1]) return;
  var last = sh.getLastRow();
  if (fromRow > last) return;
  var rng = sh.getRange(fromRow, 1, last - fromRow + 1, 2);
  var vals = rng.getValues();
  var prev = fromRow > 2 ? num_(sh.getRange(fromRow - 1, 1).getValue()) : 0;
  var col = [];
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][1] instanceof Date) { prev += 1; col.push([prev]); }
    else col.push([vals[i][0]]);
  }
  sh.getRange(fromRow, 1, col.length, 1).setValues(col);
}

// FX rate (USD→INR) for a trade date, from the FX Rates tab.
//
// The tab has no row for days the market was shut, so an exact-match lookup
// would miss weekends and holidays. This takes the most recent rate ON OR
// BEFORE the trade date — the same approximate match the sheet's own columns
// use, which is why a 04-Sep trade carries the 03-Sep rate.
function fxRateFor_(date) {
  var sh = sheetOrNull_(CFG.fxRates);
  if (!sh) return '';
  var vals = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 2).getValues();
  var want = Utilities.formatDate(date, tz_(), 'yyyy-MM-dd');
  var bestDate = null, bestRate = '', earliest = null, earliestRate = '';

  for (var i = 0; i < vals.length; i++) {
    var d = vals[i][0], v = vals[i][1];
    if (!(d instanceof Date) || typeof v !== 'number') continue;
    var key = Utilities.formatDate(d, tz_(), 'yyyy-MM-dd');
    if (earliest === null || key < earliest) { earliest = key; earliestRate = v; }
    if (key <= want && (bestDate === null || key > bestDate)) { bestDate = key; bestRate = v; }
  }
  // A trade dated before the tab starts falls back to the earliest known rate.
  return bestDate !== null ? bestRate : earliestRate;
}

/* ============================== HELPERS ============================== */

function auth_(key) {
  var pass = getPasscode_();
  if (!pass) {
    throw new Error('No passcode set. In the Apps Script editor run SET_PASSCODE("your-passcode") once, then deploy a new version.');
  }
  if (String(key || '') !== pass) throw new Error('Wrong passcode');
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
// Plain text, for the Shortcut — JSON in a "Show Result" box is unreadable.
function text_(s) {
  return ContentService.createTextOutput(String(s)).setMimeType(ContentService.MimeType.TEXT);
}
function sheet_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Tab not found: ' + name);
  return sh;
}
function sheetOrNull_(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
function tz_() { return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'Asia/Kolkata'; }
function fmtDate_(d, tz) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); }
function num_(v) { return (typeof v === 'number' && isFinite(v)) ? v : (v === '' || v == null ? null : (isNaN(Number(v)) ? null : Number(v))); }
function numOrBlank_(v) { if (v === '' || v === null || v === undefined) return ''; var n = Number(v); return isNaN(n) ? '' : n; }
function isBucketListEnd_(v) { return /^(TOTAL|RECONCILIATION|Platform)\b/i.test(String(v)); }
function findLabel_(vals, re) {
  for (var r = 0; r < vals.length; r++) for (var c = 0; c < vals[r].length; c++) {
    if (typeof vals[r][c] === 'string' && re.test(vals[r][c].trim())) return { r: r, c: c };
  }
  return null;
}

// Run this once from the editor (▶) to check everything reads correctly.
function TEST_readAll() {
  var d = getAll_();
  Logger.log('Net worth: ' + d.netWorth.head.netWorth);
  Logger.log('Assets: ' + d.netWorth.assets.length + ', transactions: ' + d.transactions.length + ', daily points: ' + d.daily.length);
}

/* ==========================================================================
 *  SCREENSHOT CAPTURE  —  INDmoney order confirmation screen
 *
 *  Two ways in, one parser:
 *    parseText   the iOS Shortcut reads the text on your phone and posts it
 *    parseImage  the app posts the image; Drive OCR reads it here
 *
 *  Neither writes anything. They return the parsed fields; the app shows
 *  them to you and you press Save, which goes through addTxn as usual.
 *
 *  parseImage needs the Drive service switched on once:
 *    Apps Script editor ▸ Services (+) ▸ Drive API ▸ Add
 * ========================================================================== */

var MONTHS = {jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,sept:9,oct:10,nov:11,dec:12};

// The rows of the INDmoney Order Details block, in the order they appear.
var SHOT_LABELS = [
  { key: 'ticker',    re: /^stock\s*ticker\s*:?$/i },
  { key: 'orderValue',re: /^order\s*value\s*:?$/i },
  { key: 'credit',    re: /^credit\s*to\s*wallet\s*:?$/i },
  { key: 'quantity',  re: /^quantity\s*:?$/i },
  { key: 'price',     re: /^avg\.?\s*price\s*:?$/i },
  { key: 'orderType', re: /^order\s*type\s*:?$/i },
  { key: 'placedOn',  re: /^placed\s*on\s*:?$/i }
];

function parseIndmoneyShot_(raw) {
  var text = String(raw || '').replace(/ /g, ' ');
  var flat = text.replace(/[ \t]+/g, ' ');
  var warn = [];

  // Pass 1 — each value sits right after its label.
  var f = {
    ticker:     grab_(flat, /Stock\s*Ticker\s*:?\s*([A-Z][A-Z0-9.\-]{0,6})\b/),
    orderValue: grab_(flat, /Order\s*Value\s*:?\s*\$?\s*([0-9][0-9,]*\.?[0-9]*)/i),
    credit:     grab_(flat, /Credit\s*to\s*Wallet\s*:?\s*\$?\s*([0-9][0-9,]*\.?[0-9]*)/i),
    quantity:   grab_(flat, /Quantity\s*:?\s*([0-9]*\.?[0-9]+)/i),
    price:      grab_(flat, /Avg\.?\s*Price\s*:?\s*\$?\s*([0-9][0-9,]*\.?[0-9]*)/i),
    orderType:  grab_(flat, /Order\s*Type\s*:?\s*(BUY|SELL)\b/i)
  };

  // Pass 2 — some OCR engines read a two-column table down the labels first
  // and then down the values. Fill anything pass 1 missed from that shape.
  if (!f.ticker || !f.quantity || !f.price) {
    var col = columnwise_(text);
    var usedCol = false;
    for (var k in col) {
      if (col.hasOwnProperty(k) && !f[k]) { f[k] = col[k]; usedCol = true; }
    }
    if (usedCol) warn.push('read as a two-column layout — check the values line up');
  }

  // --- side ---
  var side = f.orderType ? f.orderType.toUpperCase() : null;
  if (!side) {
    var mHead = /\b(Buy|Sell)\s+Order\s+Successful/i.exec(flat);
    if (mHead) { side = mHead[1].toUpperCase(); warn.push('side read from the heading, not the Order Type row'); }
  }
  if (!side) throw new Error('Could not find BUY/SELL. Is this an INDmoney order screen?');

  if (!f.ticker)   throw new Error('Could not read the Stock Ticker.');
  if (!f.quantity) throw new Error('Could not read the Quantity.');
  if (!f.price)    throw new Error('Could not read the Avg. Price.');

  var ticker = f.ticker.toUpperCase();
  var qty    = Number(f.quantity);
  var price  = Number(String(f.price).replace(/,/g, ''));
  if (!(qty > 0))   throw new Error('Quantity read as ' + f.quantity + ', which cannot be right.');
  if (!(price > 0)) throw new Error('Avg. Price read as ' + f.price + ', which cannot be right.');

  // --- the cash figure (column G) ---
  // BUY : the big headline = what was charged
  // SELL: "Credit to Wallet"  = what was received
  var gross0 = qty * price;
  var cash = null;
  if (side === 'SELL' && f.credit) {
    cash = Number(String(f.credit).replace(/,/g, ''));
  } else if (side === 'BUY') {
    var mHeadline = /(?:Buy|Sell)\s+Order\s+Successful\s*\$\s*([0-9][0-9,]*\.[0-9]{2})(?![0-9])/i.exec(flat);
    if (mHeadline) cash = Number(mHeadline[1].replace(/,/g, ''));
  }

  // Positional matching is at the mercy of how the OCR ordered the text, so
  // fall back to finding it by value: the cash figure always sits within a
  // fraction of a percent of quantity x price — above it on a buy (you paid
  // the fee) and below it on a sell (the fee came out). Nothing else on the
  // screen sits in that band, which makes this a safer signal than position.
  if (cash == null) {
    var cands = moneyCandidates_(flat).filter(function (v) {
      if (Math.abs(v - gross0) <= 0.02) return false;          // that is the order value
      var r = v / gross0;
      return r > 0.9 && r < 1.1;                               // within 10%
    });
    if (side === 'BUY') cands = cands.filter(function (v) { return v > gross0; });
    else                cands = cands.filter(function (v) { return v < gross0; });

    if (cands.length) {
      // closest to the order value wins — the fee is small
      cands.sort(function (a, b) { return Math.abs(a - gross0) - Math.abs(b - gross0); });
      cash = cands[0];
      warn.push('the ' + (side === 'BUY' ? 'amount charged' : 'amount credited') +
                ' was matched by value, not by its label — worth a check');
    }
  }

  if (cash == null) {
    warn.push(side === 'SELL'
      ? 'no "Credit to Wallet" found — enter the amount credited yourself'
      : 'could not read the amount charged — enter it yourself');
  }

  // --- date: "7:01 PM, 04 Sep 2026" or "04 Sep 2026" ---
  var date = null;
  var mDate = /\b(\d{1,2})\s+([A-Za-z]{3,9})\.?\s+(\d{4})\b/.exec(flat);
  if (mDate) {
    var w = mDate[2].toLowerCase();
    var mon = MONTHS[w.slice(0, 4)] || MONTHS[w.slice(0, 3)];
    if (mon) date = mDate[3] + '-' + pad2_(mon) + '-' + pad2_(Number(mDate[1]));
  }
  if (!date) warn.push('no date found — defaulted to today');

  var gross = qty * price;
  if (cash != null) {
    var fee = side === 'SELL' ? gross - cash : cash - gross;
    if (Math.abs(fee) < 0.005) {
      warn.push('the fee comes out as zero — the cash amount was probably read from the wrong row');
    } else if (fee < -0.5 || fee > Math.max(2, gross * 0.05)) {
      warn.push('the fee looks wrong ($' + fee.toFixed(2) + ') — check the amounts');
    }
  }

  return {
    platform: 'INDMoney',
    category: 'US Stocks',
    order: side === 'BUY' ? 'Buy' : 'Sell',
    ticker: ticker,
    volume: qty,
    price: price,
    amount: cash,
    gross: round_(gross),
    date: date,
    warnings: warn
  };
}

function grab_(s, re) { var m = re.exec(s); return m ? m[1] : null; }

// Every money-looking figure in the text. Prefers $-prefixed ones but also
// accepts bare two-decimal numbers, since OCR sometimes drops the symbol.
function moneyCandidates_(s) {
  var out = [], m;
  var withSign = /\$\s*([0-9][0-9,]*\.[0-9]{2})(?![0-9])/g;
  while ((m = withSign.exec(s)) !== null) out.push(Number(m[1].replace(/,/g, '')));
  var bare = /(?:^|[^\w.$])([0-9][0-9,]*\.[0-9]{2})(?![0-9])/g;
  while ((m = bare.exec(s)) !== null) {
    var v = Number(m[1].replace(/,/g, ''));
    if (out.indexOf(v) < 0) out.push(v);
  }
  return out;
}

// Handles the "all labels, then all values" shape. Returns only what it is
// confident about: it requires at least three labels in an unbroken run.
function columnwise_(text) {
  var lines = String(text).split(/\r?\n/).map(function (l) { return l.trim(); })
                          .filter(function (l) { return l.length; });
  var order = [], lastIdx = -1, run = 0, bestRun = 0, bestOrder = [], bestEnd = -1;

  for (var i = 0; i < lines.length; i++) {
    var key = labelOf_(lines[i]);
    if (key) {
      if (lastIdx === i - 1 || lastIdx === -1) run++; else { run = 1; order = []; }
      order.push(key); lastIdx = i;
      if (run > bestRun) { bestRun = run; bestOrder = order.slice(); bestEnd = i; }
    }
  }
  if (bestRun < 3) return {};

  // The values follow the label run — but not always cleanly. Real OCR output
  // slots stray interface text in between (a "Buy More" button, for one), so
  // pairing them off position-by-position shifts everything after the intruder.
  // Instead each label claims the next line that LOOKS like its kind of value.
  var SHAPES = {
    ticker:     /^([A-Z][A-Z0-9.\-]{0,6})$/,
    quantity:   /^([0-9]*\.?[0-9]+)$/,
    price:      /^\$?\s*([0-9][0-9,]*\.?[0-9]*)$/,
    orderValue: /^\$?\s*([0-9][0-9,]*\.?[0-9]*)$/,
    credit:     /^\$?\s*([0-9][0-9,]*\.?[0-9]*)$/,
    orderType:  /\b(BUY|SELL)\b/i
  };
  var LOOKAHEAD = 4;   // don't let a missing value swallow the next one

  var vals = lines.slice(bestEnd + 1).filter(function (l) { return !labelOf_(l); });
  var out = {}, at = 0;
  for (var j = 0; j < bestOrder.length; j++) {
    var k = bestOrder[j], shape = SHAPES[k];
    if (!shape) continue;
    for (var scan = at; scan < vals.length && scan < at + LOOKAHEAD; scan++) {
      var m = shape.exec(vals[scan]);
      if (m) { out[k] = m[1]; at = scan + 1; break; }
    }
  }
  return out;
}

function labelOf_(line) {
  for (var i = 0; i < SHOT_LABELS.length; i++) {
    if (SHOT_LABELS[i].re.test(line)) return SHOT_LABELS[i].key;
  }
  return null;
}

function pad2_(n) { return (n < 10 ? '0' : '') + n; }
function round_(n) { return Math.round(n * 1e6) / 1e6; }

// Converts an image to text using Drive's OCR, then deletes the temp file.
// Works with either version of the Drive advanced service.
/* =================== KITE — VIRTUAL CONTRACT NOTE ==================== */
//
// The Kite note is a different animal from an INDmoney order screen: it can
// carry several trades, and it states each trade's share of the charges on
// its own line. What it never states is the trade DATE (only clock times) or
// the DP charge, so those are filled in by the caller.
//
// The hard part is the rupee sign. Real OCR of this screen renders ₹ as
// "%", "7", "2", "$" or nothing at all, so ₹0.42 arrives as "20.42" — a
// leading digit that is not a digit. Position can't tell them apart, but
// arithmetic can: the per-trade charges must add up to the stated Total.

var KITE_EQUITY_EXCH  = ['NSE', 'BSE'];
var KITE_DERIV_EXCH   = ['NFO', 'BFO', 'CDS', 'MCX', 'BCD'];

function looksLikeKiteNote_(raw) {
  var t = String(raw || '');
  if (/contract\s*note/i.test(t)) return true;
  if (/\bBrokerage\b/i.test(t) && /\bSTT\b/i.test(t)) return true;
  // A screenshot taken further down the note has neither heading nor charges,
  // so fall back to the shape of a trade block itself.
  return /Qty\.?\s*[0-9][0-9,.]*\s*Avg\.?\s*[0-9]/i.test(t) &&
         /\b(NSE|BSE|NFO|BFO|CDS|MCX|BCD)\b/.test(t);
}

// One money token -> the readings it could plausibly be.
//
//   "$25.71" -> [25.71]          a currency character survived, so every
//   "%26.13" -> [26.13]          digit after it is a real digit
//   "20.42"  -> [20.42, 0.42]    nothing in front, so the leading digit may
//   "70.40"  -> [70.40, 0.40]    itself be the rupee sign misread
function rupeeReadings_(tok) {
  var t = String(tok || '').trim();
  var digits = t.replace(/[^0-9.]/g, '');
  if (!/^[0-9]*\.?[0-9]+$/.test(digits)) return [];
  var n = Number(digits);
  if (!isFinite(n)) return [];
  var out = [n];

  // Did a currency glyph come through as something other than a digit?
  var lead = /^[^0-9]/.test(t);
  if (!lead) {
    // Bare token: the first digit could be a mangled rupee sign, but only
    // when a digit sits in front of the decimal point.
    var m = /^([0-9])([0-9]*\.[0-9]+)$/.exec(digits);
    if (m) {
      var alt = Number(m[2]);
      if (isFinite(alt) && out.indexOf(alt) < 0) out.push(alt);
    }
  }
  return out;
}

// Picks one reading per trade so that the fees sum to the note's Total.
// Falls back to the literal reading, with a warning, when it can't be pinned.
function reconcileFees_(trades, total, warn) {
  var open = [];
  for (var i = 0; i < trades.length; i++) {
    if (trades[i].feeReadings && trades[i].feeReadings.length) open.push(i);
  }
  if (!open.length) return;

  var pick = function (idx, val) { trades[idx].fees = val; };

  if (total == null) {
    // No Total in shot. Anything with only one possible reading is settled;
    // the rest are left for feeByHistory_ to judge.
    for (var a = 0; a < open.length; a++) {
      var only = trades[open[a]].feeReadings;
      if (only.length === 1) pick(open[a], only[0]);
    }
    return;
  }
  if (open.length > 12) {
    for (var b0 = 0; b0 < open.length; b0++) pick(open[b0], trades[open[b0]].feeReadings[0]);
    return;
  }

  // Every combination of readings; there are at most two per trade.
  var solutions = [];
  var walk = function (at, chosen, sum) {
    if (solutions.length > 1) return;                       // ambiguous already
    if (at === open.length) {
      if (Math.abs(sum - total) <= 0.02) solutions.push(chosen.slice());
      return;
    }
    var reads = trades[open[at]].feeReadings;
    for (var r = 0; r < reads.length; r++) {
      if (sum + reads[r] > total + 0.02) continue;          // prune
      chosen.push(reads[r]);
      walk(at + 1, chosen, sum + reads[r]);
      chosen.pop();
    }
  };
  walk(0, [], 0);

  if (solutions.length === 1) {
    for (var b = 0; b < open.length; b++) pick(open[b], solutions[0][b]);
    return;
  }

  // No single answer. Most often this is a screenshot showing part of a longer
  // note: the Total covers trades that are not on screen, so it can never
  // balance. Leave these for feeByHistory_ rather than writing a reading that
  // is probably the mangled one.
  warn.push(solutions.length
    ? 'the charges could be read more than one way — check them against \u20b9' + total
    : 'the \u20b9' + total + ' total covers more trades than this screenshot shows');
}

// The individual charge rows above the Total, as candidate readings.
var KITE_CHARGE_LABEL = /^(Brokerage|STT|CTT|Stamp\s*duty|Exchange\s*turnover\s*charge|SEBI\s*turnover\s*charge|Transaction\s*charge|Clearing\s*charge|IPFT|GST)\b/i;

function chargeRows_(lines) {
  var out = [];
  for (var i = 0; i < lines.length; i++) {
    if (/^Total\b/i.test(lines[i])) break;
    if (!KITE_CHARGE_LABEL.test(lines[i])) continue;
    var m = /(\S+)\s*$/.exec(lines[i]);
    if (!m) continue;
    var reads = rupeeReadings_(m[1]);
    if (reads.length) out.push(reads);
  }
  return out;
}

// Picks the Total reading that the charge rows actually add up to.
function resolveTotal_(totalReadings, components) {
  if (!totalReadings.length) return null;
  if (totalReadings.length === 1) return totalReadings[0];
  if (!components.length || components.length > 10) return null;

  var hits = {};
  var walk = function (at, sum) {
    if (at === components.length) {
      for (var k = 0; k < totalReadings.length; k++) {
        if (Math.abs(sum - totalReadings[k]) <= 0.02) hits[totalReadings[k]] = true;
      }
      return;
    }
    for (var r = 0; r < components[at].length; r++) walk(at + 1, sum + components[at][r]);
  };
  walk(0, 0);

  var found = Object.keys(hits);
  return found.length === 1 ? Number(found[0]) : null;
}

function parseKiteNote_(raw) {
  var text  = String(raw || '').replace(/ /g, ' ');
  var lines = text.split(/\r?\n/).map(function (l) { return l.trim(); })
                  .filter(function (l) { return l.length; });
  var warn = [];

  // --- the Total of all charges, used to disambiguate the per-trade figures ---
  // The Total is itself a money token, so it can read two ways too. The rows
  // above it settle the matter: brokerage + STT + ... + GST must equal it, and
  // only one reading of the Total ever satisfies that.
  var totalReadings = [];
  for (var t = 0; t < lines.length; t++) {
    var mt = /^Total\b\s*(\S+)\s*$/i.exec(lines[t]);
    if (mt) { totalReadings = rupeeReadings_(mt[1]); break; }
  }
  var total = resolveTotal_(totalReadings, chargeRows_(lines));

  // --- the trade lines ---
  // "Qty. 15 Avg. 236.0  11:53:16 NFO BUY"  followed by
  // "BANKNIFTY 03rd AUG 45600 CE  ₹25.71"
  var trades = [], skipped = [];
  var qtyRe = /Qty\.?\s*([0-9][0-9,]*(?:\.[0-9]+)?)\s*Avg\.?\s*([0-9][0-9,]*(?:\.[0-9]+)?)/i;
  var exchRe = /\b(NSE|BSE|NFO|BFO|CDS|MCX|BCD)\b/i;
  var sideRe = /\b(BUY|SELL)\b/i;
  var timeRe = /\b(\d{1,2}:\d{2}:\d{2})\b/;
  // A symbol on its own: not a side, not an exchange, not a time or a number.
  var STOP = { BUY:1, SELL:1, NSE:1, BSE:1, NFO:1, BFO:1, CDS:1, MCX:1, BCD:1, QTY:1, AVG:1 };

  for (var i = 0; i < lines.length; i++) {
    var mq = qtyRe.exec(lines[i]);
    if (!mq) continue;

    // Everything from this "Qty. / Avg." line up to the next one is a single
    // trade. Which of those lines holds the side, the exchange, the symbol and
    // the charge depends on the OCR engine — Google Drive, the phone's own
    // reader and desktop OCR all order them differently — so the block is
    // searched as a whole rather than line by line in an assumed order.
    var block = [lines[i]];
    for (var j = i + 1; j < lines.length && block.length < 8; j++) {
      if (qtyRe.test(lines[j])) break;
      if (/^(Read more|The data is|Total\b)/i.test(lines[j])) break;
      block.push(lines[j]);
    }
    var body = block.slice(1);
    var whole = block.join('\n');

    var mEx   = exchRe.exec(whole);
    var mSide = sideRe.exec(whole);
    var mTime = timeRe.exec(whole);

    // Symbol: the first token in the block that reads like one.
    var name = '';
    for (var b = 0; b < body.length && !name; b++) {
      var toks = body[b].split(/\s+/);
      for (var t = 0; t < toks.length; t++) {
        var tok = toks[t].replace(/[^A-Za-z0-9&.\-]/g, '');
        if (!/^[A-Z][A-Z0-9&.\-]{1,19}$/.test(tok)) continue;
        if (STOP[tok.toUpperCase()]) continue;
        name = tok; break;
      }
    }

    // Charge: the last two-decimal money figure in the block, wherever it sits.
    // The quantity and average price are on the head line, which is excluded,
    // and a clock time can never match this shape.
    var feeTok = '';
    for (var c = 0; c < body.length; c++) {
      var cand = body[c].split(/\s+/);
      for (var k = 0; k < cand.length; k++) {
        if (/^[^0-9]{0,2}[0-9][0-9,]*\.[0-9]{2}$/.test(cand[k])) feeTok = cand[k];
      }
    }

    var exch = mEx ? mEx[1].toUpperCase() : '';
    var side = mSide ? mSide[1].toUpperCase() : '';
    var rec = {
      ticker:   name,
      exchange: exch,
      order:    side === 'SELL' ? 'Sell' : 'Buy',
      volume:   Number(String(mq[1]).replace(/,/g, '')),
      price:    Number(String(mq[2]).replace(/,/g, '')),
      feeReadings: rupeeReadings_(feeTok),
      sideRead: !!side,
      at: mTime ? mTime[1] : ''
    };

    // Equity only, by your choice: F&O has never appeared in the sheet, so
    // there is no category or ticker format to write it into. It is still
    // parsed, because its share of the charges is part of the Total and the
    // equity figures cannot be reconciled without it.
    // The exchange column does not always survive OCR, but a derivative names
    // itself: an option carries CE/PE and an expiry, a future carries FUT.
    var derivName = /\b(CE|PE|FUT)\b/.test(whole) && /\b(\d{1,2}\s*(?:st|nd|rd|th)?|JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\b/i.test(whole);
    if (KITE_DERIV_EXCH.indexOf(exch) >= 0 || derivName) rec.drop = 'F&O';
    else if (!name) rec.drop = 'no symbol';
    else if (!(rec.volume > 0) || !(rec.price > 0)) rec.drop = 'unreadable figures';

    trades.push(rec);
  }

  if (!trades.length) {
    throw new Error('No trades found. Is this the Kite contract note screen?');
  }

  // Several screenshots of one note will overlap, so the same trade appears
  // more than once. A trade is the same trade if the symbol, size, price and
  // clock time match.
  var seen = {}, unique = [], dupes = 0;
  for (var u = 0; u < trades.length; u++) {
    var tu = trades[u];
    var key = [tu.ticker, tu.volume, tu.price, tu.order, tu.at].join('|');
    if (seen[key]) { dupes++; continue; }
    seen[key] = true; unique.push(tu);
  }
  trades = unique;
  if (dupes) warn.push('the screenshots overlapped — ' + dupes +
    ' repeated trade' + (dupes > 1 ? 's were' : ' was') + ' counted once');

  // Reconciled across every line on the note, dropped ones included.
  reconcileFees_(trades, total, warn);
  feeByHistory_(trades, warn);

  var keep = [];
  for (var d = 0; d < trades.length; d++) {
    var dt = trades[d];
    if (!dt.drop) {
      if (!dt.sideRead) warn.push(dt.ticker + ': no BUY/SELL found — assumed Buy');
      keep.push(dt); continue;
    }
    if (dt.drop === 'F&O') skipped.push(dt.ticker || 'an F&O trade');
    else if (dt.drop === 'no symbol') warn.push('a trade line had no readable symbol and was left out');
    else warn.push((dt.ticker || 'a trade') + ': quantity or price did not read — left out');
  }
  trades = keep;

  for (var k = 0; k < trades.length; k++) {
    var tr = trades[k];
    delete tr.feeReadings;
    delete tr.drop;
    delete tr.sideRead;
    delete tr.at;
    tr.platform = 'Kite';
    tr.category = 'Stocks';
    // Kite rows in the sheet keep fees out of Amount: G = H = qty x price.
    tr.amount = round_(tr.volume * tr.price);
    if (tr.fees == null) warn.push(tr.ticker + ': charges did not read — enter them yourself');
  }

  if (skipped.length) {
    warn.push('skipped ' + skipped.join(', ') + ' — F&O is not logged in your sheet');
  }
  return { platform: 'Kite', kind: 'kiteNote', total: total, trades: trades,
           skipped: skipped, warnings: warn };
}

// What your Kite charges normally come to, as a share of the trade value.
// Read from your own rows, so it follows your actual costs rather than my
// assumptions about Zerodha's rate card.
function kiteFeeBand_(sh, order) {
  var last = sh.getLastRow();
  var span = Math.min(600, last - 1);
  if (span <= 0) return null;
  var vals = sh.getRange(last - span + 1, 4, span, 9).getValues();   // D..L
  var ratios = [];
  for (var i = 0; i < vals.length; i++) {
    var ord = String(vals[i][1]), plat = String(vals[i][2]);
    var amt = Number(vals[i][3]), fee = Number(vals[i][8]);
    if (!/^Kite$/i.test(plat) || ord !== order) continue;
    if (!(amt > 0) || !(fee > 0)) continue;
    ratios.push(fee / amt);
  }
  if (ratios.length < 20) return null;              // too little to judge by
  ratios.sort(function (a, b) { return a - b; });
  return ratios[Math.floor(ratios.length * 0.9)];   // the 90th percentile
}

// Last resort when the Total is not in the screenshot: keep the reading that
// looks like a charge you actually pay, and reject one that does not.
function feeByHistory_(trades, warn) {
  var sh, matched = 0, unsure = 0;
  for (var i = 0; i < trades.length; i++) {
    var t = trades[i];
    if (t.fees != null || !t.feeReadings || t.feeReadings.length < 2) continue;
    var value = t.volume * t.price;
    if (!(value > 0)) continue;
    if (!sh) sh = sheet_(CFG.transactions);
    var p90 = kiteFeeBand_(sh, t.order);
    if (!p90) continue;
    var cap = p90 * 3;                               // generous, but 40x is not
    var fits = [];
    for (var r = 0; r < t.feeReadings.length; r++) {
      if (t.feeReadings[r] / value <= cap) fits.push(t.feeReadings[r]);
    }
    if (fits.length === 1) { t.fees = fits[0]; matched++; }
    else unsure++;
  }
  if (matched) warn.push((matched > 1 ? 'the charges were' : 'the charge was') +
    ' matched against what you normally pay, not against a Total — worth a glance');
  if (unsure) warn.push('could not tell what ' + unsure +
    ' of the charges should be — check them before saving');
}

// The DP charge is billed separately and never appears on the note, so it is
// carried forward from the last Kite sell that had one.
function lastDpCharge_(sh) {
  var last = sh.getLastRow();
  var span = Math.min(400, last - 1);
  if (span <= 0) return '';
  var start = last - span + 1;
  var vals = sh.getRange(start, 5, span, 7).getValues();     // E..K
  for (var i = vals.length - 1; i >= 0; i--) {
    var order = String(vals[i][0]), plat = String(vals[i][1]), dp = vals[i][6];
    if (/^Kite$/i.test(plat) && /^Sell$/i.test(order) && dp !== '' && dp != null) {
      return Number(dp);
    }
  }
  return '';
}

function ocrImage_(base64, mimeType) {
  if (!base64) throw new Error('No image received.');
  if (typeof Drive === 'undefined') {
    throw new Error('Turn on the Drive service first: Apps Script editor > Services (+) > Drive API > Add.');
  }
  var bytes = Utilities.base64Decode(String(base64).replace(/^data:[^,]*,/, ''));
  var blob = Utilities.newBlob(bytes, mimeType || 'image/png', 'shot');
  var meta = { title: 'ocr-temp', name: 'ocr-temp', mimeType: 'application/vnd.google-apps.document' };
  var opts = { ocr: true, ocrLanguage: 'en' };

  var file;
  if (Drive.Files.create) file = Drive.Files.create(meta, blob, opts);   // v3
  else                    file = Drive.Files.insert(meta, blob, opts);   // v2
  var id = file.id;

  try {
    var text = DocumentApp.openById(id).getBody().getText();
    if (!text || !text.trim()) throw new Error('Nothing readable in that image.');
    return text;
  } finally {
    try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
  }
}

// Which broker a screenshot came from is the script's problem, not the
// Shortcut's — it just posts whatever text the OCR produced.
function parseAny_(text) {
  if (looksLikeKiteNote_(text)) {
    var k = parseKiteNote_(text);
    // The note has no date and no DP charge; fill both in the way your sheet does.
    var today = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
    var dp = null;
    for (var i = 0; i < k.trades.length; i++) {
      var tr = k.trades[i];
      tr.date = today;
      if (tr.order === 'Sell') {
        if (dp === null) dp = lastDpCharge_(sheet_(CFG.transactions));
        tr.dp = dp;
      }
    }
    // Flag anything the sheet already holds, so a second screenshot of the
    // same note does not offer them for saving again.
    if (k.trades.length) {
      var shx = sheet_(CFG.transactions);
      var dupes = 0;
      for (var q = 0; q < k.trades.length; q++) {
        if (existingRowFor_(shx, k.trades[q])) { k.trades[q].existing = true; dupes++; }
      }
      if (dupes) k.warnings.push(dupes + ' of these ' +
        (dupes > 1 ? 'are' : 'is') + ' already in the sheet and will not be added again');
    }
    return k;
  }
  return parseIndmoneyShot_(text);
}

// Always hands back what the OCR saw, so a failure can be diagnosed instead
// of just refused. The caller decides how to present it.
function tryParse_(text) {
  var out;
  try { out = parseAny_(text); }
  catch (err) { out = { parseError: String(err.message || err) }; }
  out.ocrText = String(text || '').slice(0, 6000);
  return out;
}

/* Run this from the editor to check the parser without taking a screenshot.
 * The sample is real Google Drive OCR output from an AVGO buy — note the
 * "Buy More" button text sitting between the labels and their values, which
 * is exactly the shape that used to break it. */
function TEST_parser() {
  var sample = [
    '11:24', '71', 'Л',
    'Buy Order Successful', '$100.00', 'Broadcom Inc.', '7:01 PM, 04 Sep 2026',
    'Congrats MD! Order successful, you are now a shareholder.',
    'Order Details',
    'Stock Ticker', 'Order Value', 'Quantity', 'Avg. Price', 'Order Type',
    'Buy More',
    'AVGO', '$99.70', '0.280301837', '$355.688', 'BUY, Market'
  ].join('\n');
  var r = parseIndmoneyShot_(sample);
  Logger.log(JSON.stringify(r, null, 2));
  if (r.ticker !== 'AVGO' || r.volume !== 0.280301837 || r.price !== 355.688 || r.amount !== 100) {
    throw new Error('Parser is not reading the sample correctly — send this log to Claude.');
  }
  Logger.log('Parser OK.');
}
