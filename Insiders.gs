/*************************************************************
 * Insiders.gs — insider (Form 4) transaction history + signals (v1)
 * -----------------------------------------------------------
 * DATA SHAPE IS DIFFERENT FROM Fundamentals.gs: Fundamentals is
 * one-row-per-ticker, updated in place, with staleness tracked on
 * that same row's own Last_Updated column. Insiders is one-row-
 * PER-TRANSACTION in Raw_Insiders (many rows per ticker, pure
 * append — a filed transaction is a historical fact that never
 * gets "updated" once written). Staleness/resumability is
 * therefore tracked in a separate, small checkpoint sheet
 * (Raw_InsidersCheckpoint — just Ticker + Last_Checked), NOT on
 * the Raw_Insiders data rows themselves.
 *
 * RESUMABILITY: same mechanism as Fundamentals.gs — the entry
 * point computes which tickers are missing a fresh Last_Checked
 * stamp (within INSIDERS_STALENESS_DAYS) in Raw_InsidersCheckpoint,
 * and only processes those.
 *   - An interrupted run resumes almost exactly where it left off
 *     when you call the exact same function again — no re-work.
 *   - A normal weekly refresh naturally reprocesses everything
 *     once last week's checkpoint ages past the staleness threshold.
 *   - Pass { forceAll: true } to bypass freshness and recheck
 *     every ticker regardless.
 *
 * HONEST SCOPING NOTE: Finnhub's free insider-transactions endpoint
 * does NOT include insider seniority/title (CEO vs Director vs
 * other), so transaction value is normalized against the ticker's
 * own MarketCap instead (from Raw_Universe) — a $500K purchase
 * means very different things for a $50M company vs a $500B one.
 * Also note: there's no reliable way to distinguish a routine
 * pre-scheduled 10b5-1 sale from a discretionary one with this
 * data — cluster detection (multiple DISTINCT insiders selling
 * within the window) is the best available proxy, imperfect but
 * real signal.
 *
 * RETENTION: every fetch is filtered down to transactions within
 * the last INSIDERS_RETENTION_DAYS (370 — wider than a typical
 * "keep 200 days" choice), wide enough to support a true trailing-
 * 12-month Buy Signal calculation.
 *
 * DEDUP: Raw_Insiders is pure append, so each incoming transaction
 * is checked against a composite key (Ticker|Name|Transaction_Date|
 * Transaction_Code) built from the rows already on the sheet —
 * re-running over a ticker whose transactions were already recorded
 * never creates duplicate rows.
 *
 * The pending-ticker list lives in a hidden scratch sheet
 * (Raw_InsidersScratch), not Script Properties — same reason as
 * every other scratch sheet in this project: a list of thousands
 * of tickers would blow past Properties' ~9KB per-value limit.
 *
 * One call per ticker:
 *   Finnhub /stock/insider-transactions -> raw Form 4 transactions
 *
 * RUN ORDER: run Universe.gs first so Raw_Universe has tickers and
 * MarketCap to work from; this still runs fine without it, Buy_
 * Signal_Score just comes back blank ("") for tickers with no
 * MarketCap to normalize against.
 *
 * ENTRY POINTS:
 *   runInsidersRefresh()                    — gap-fill/resume, full universe
 *   runInsidersRefresh({ forceAll: true })  — recheck everyone regardless
 *   runInsidersRefresh_Test()               — synchronous, first 20 pending tickers
 *   runInsidersRefresh_Test(true)           — same, but via chunked triggers
 *************************************************************/
function runInsidersRefresh(options) {
  const mode = (options && options.mode) || MODE.PROD;
  const chunk = (options && typeof options.chunk !== "undefined") ? options.chunk : (mode === MODE.PROD);
  const forceAll = (options && options.forceAll) || false;
  _ensureInsidersHeaders();
  _ensureInsiderSignalsHeaders();
  _ensureInsidersCheckpointHeaders();
  const uniSheet = _getOrCreateUniverseSheet();
  const uniLast = uniSheet.getLastRow();
  if (uniLast < 2) {
    SpreadsheetApp.getActive().toast("Insiders: Raw_Universe is empty — run Universe first.");
    return;
  }
  const allTickers = uniSheet.getRange(2, 1, uniLast - 1, 1).getValues().flat()
    .map(t => String(t).trim().toUpperCase()).filter(Boolean);
  const checkpointSheet = _getOrCreateInsidersCheckpointSheet();
  const cpLast = checkpointSheet.getLastRow();
  const freshSet = new Set();
  if (cpLast >= 2 && !forceAll) {
    const existing = checkpointSheet.getRange(2, 1, cpLast - 1, INSIDERS_CHECKPOINT_HEADERS.length).getValues();
    const staleCutoff = Date.now() - (INSIDERS_STALENESS_DAYS * 24 * 60 * 60 * 1000);
    existing.forEach(r => {
      const ticker = String(r[0] || "").trim().toUpperCase();
      const lastChecked = r[1]; // Last_Checked column
      if (ticker && lastChecked instanceof Date && lastChecked.getTime() >= staleCutoff) {
        freshSet.add(ticker);
      }
    });
  }
  let pending = allTickers.filter(t => !freshSet.has(t));
  if (mode === MODE.TEST) {
    pending = pending.slice(0, DEFAULT_TEST_LIMIT);
  }
  if (pending.length === 0) {
    SpreadsheetApp.getActive().toast(`Insiders: everything already fresh (checked within ${INSIDERS_STALENESS_DAYS} days). Nothing to do.`);
    return;
  }
  _writeInsidersPendingScratch(pending);
  if (chunk) {
    startChunkProcess("Insiders", "chunkWorker_Insiders", "finalize_Insiders", pending.length, 100);
  } else {
    let idx = 0;
    while (idx < pending.length) idx = chunkWorker_Insiders(idx, 100);
    finalize_Insiders();
  }
  SpreadsheetApp.getActive().toast(
    `Insiders refresh started (${mode}). ${pending.length} pending, ${allTickers.length - pending.length} already fresh (skipped).`
  );
}
function runInsidersRefresh_Test(chunk) {
  return runInsidersRefresh({ mode: MODE.TEST, chunk: (chunk === true) });
}
/*************************************************************
 * Chunk worker — reads its slice from the PENDING scratch list
 * (not directly from Raw_Universe by position — same reasoning
 * as chunkWorker_Fundamentals: the pending set can skip around,
 * so a contiguous row-range read against Raw_Universe wouldn't
 * line up with it).
 *************************************************************/
function chunkWorker_Insiders(startIndex, chunkSize) {
  const pendingSheet = _getInsidersPendingScratchSheet();
  const total = Math.max(pendingSheet.getLastRow(), 0);
  if (startIndex >= total) return total;
  const endIndex = Math.min(startIndex + chunkSize, total);
  const numRows = endIndex - startIndex;
  const startRowNum = startIndex + 1; // no header row in this scratch sheet
  const tickers = pendingSheet.getRange(startRowNum, 1, numRows, 1).getValues().flat();

  const insidersSheet = _getOrCreateInsidersSheet();
  const signalsSheet = _getOrCreateInsiderSignalsSheet();
  const checkpointSheet = _getOrCreateInsidersCheckpointSheet();
  const uniSheet = _getOrCreateUniverseSheet();

  // a) Existing composite dedup keys from Raw_Insiders (first 8 columns:
  // Ticker, Name, Transaction_Date, Filing_Date, Change, Shares,
  // Transaction_Price, Transaction_Code — only the 1st/2nd/3rd/8th of
  // those feed the key itself).
  const existingKeys = new Set();
  const iLast = insidersSheet.getLastRow();
  if (iLast >= 2) {
    const existingData = insidersSheet.getRange(2, 1, iLast - 1, 8).getValues();
    existingData.forEach(r => {
      const rTicker = String(r[0] || "");
      const rName = String(r[1] || "");
      const rDate = r[2];
      const rCode = String(r[7] || "");
      const key = rTicker.trim().toUpperCase() + "|" + rName.trim() + "|" +
        rDate.toString().substring(0, 10) + "|" + rCode.trim();
      existingKeys.add(key);
    });
  }

  // b) Map Ticker -> row number in Raw_InsiderSignals (existing-row-
  // update-or-append pattern, same spirit as existingRowByTicker in
  // Fundamentals.gs).
  const signalsRowByTicker = new Map();
  const sLast = signalsSheet.getLastRow();
  if (sLast >= 2) {
    const existingSignalTickers = signalsSheet.getRange(2, 1, sLast - 1, 1).getValues().flat();
    existingSignalTickers.forEach((t, i) => signalsRowByTicker.set(String(t).trim().toUpperCase(), i + 2));
  }

  // c) Map Ticker -> row number in Raw_InsidersCheckpoint (same
  // update-or-append pattern).
  const checkpointRowByTicker = new Map();
  const cpLast = checkpointSheet.getLastRow();
  if (cpLast >= 2) {
    const existingCheckpointTickers = checkpointSheet.getRange(2, 1, cpLast - 1, 1).getValues().flat();
    existingCheckpointTickers.forEach((t, i) => checkpointRowByTicker.set(String(t).trim().toUpperCase(), i + 2));
  }

  // d) Map Ticker -> MarketCap, read from Raw_Universe (Ticker = col 1,
  // MarketCap = col 6 / index 5 zero-indexed).
  const marketCapByTicker = new Map();
  const uniLast = uniSheet.getLastRow();
  if (uniLast >= 2) {
    const uniData = uniSheet.getRange(2, 1, uniLast - 1, 6).getValues();
    uniData.forEach(r => {
      const uTicker = String(r[0] || "").trim().toUpperCase();
      if (uTicker) marketCapByTicker.set(uTicker, r[5]);
    });
  }

  const newInsiderRows = [];
  const newSignalRows = [];
  const newCheckpointRows = [];

  const oneYearCutoffMs = Date.now() - (365 * 24 * 60 * 60 * 1000);
  const ninetyDayCutoffMs = Date.now() - (90 * 24 * 60 * 60 * 1000);

  for (let i = 0; i < tickers.length; i++) {
    const ticker = String(tickers[i] || "").trim().toUpperCase();
    if (!ticker) continue;

    const transactions = _fetchInsiderTransactions(ticker);
    Utilities.sleep(150);

    const filteredTransactions = transactions.filter(tx => {
      const d = new Date(tx.transactionDate);
      if (isNaN(d.getTime())) return false;
      return (Date.now() - d.getTime()) <= INSIDERS_RETENTION_DAYS * 24 * 60 * 60 * 1000;
    });

    for (const tx of filteredTransactions) {
      const txName = String(tx.name || "");
      const txDate = tx.transactionDate || "";
      const txCode = String(tx.transactionCode || "");
      const key = ticker.trim().toUpperCase() + "|" + txName.trim() + "|" +
        txDate.toString().substring(0, 10) + "|" + txCode.trim();
      if (existingKeys.has(key)) continue;
      existingKeys.add(key);
      newInsiderRows.push([
        ticker,
        tx.name || "",
        tx.transactionDate || "",
        tx.filingDate || "",
        tx.change || "",
        tx.share || "",
        tx.transactionPrice || "",
        tx.transactionCode || "",
        TRANSACTION_CODES[tx.transactionCode] || tx.transactionCode || "",
        tx.isDerivative || false,
        new Date()
      ]);
    }

    const marketCap = marketCapByTicker.get(ticker);
    const marketCapIsValid = (typeof marketCap === "number") && isFinite(marketCap) && marketCap > 0;

    // Buy Signal (trailing 12 months)
    const purchases = filteredTransactions.filter(tx => {
      if (tx.transactionCode !== "P") return false;
      const d = new Date(tx.transactionDate);
      if (isNaN(d.getTime())) return false;
      return d.getTime() >= oneYearCutoffMs;
    });
    const netValueBought12M = purchases.reduce(
      (sum, tx) => sum + (Number(tx.share) || 0) * (Number(tx.transactionPrice) || 0), 0
    );
    let buySignalScore;
    if (!marketCapIsValid) {
      buySignalScore = ""; // genuinely unknown — no denominator to judge against
    } else {
      const ratio = netValueBought12M / marketCap;
      buySignalScore = Math.max(0, Math.min(100, ratio * 20000));
      buySignalScore = Math.round(buySignalScore * 10) / 10;
    }

    // Sell Urgency (trailing 90 days)
    const sales = filteredTransactions.filter(tx => {
      if (tx.transactionCode !== "S") return false;
      const d = new Date(tx.transactionDate);
      if (isNaN(d.getTime())) return false;
      return d.getTime() >= ninetyDayCutoffMs;
    });
    let sellUrgencyScore, netValueSold90D, distinctSellers90D;
    if (sales.length === 0) {
      sellUrgencyScore = 0;
      netValueSold90D = 0;
      distinctSellers90D = 0;
    } else {
      netValueSold90D = sales.reduce(
        (sum, tx) => sum + (Number(tx.share) || 0) * (Number(tx.transactionPrice) || 0), 0
      );
      const sellerSet = new Set(sales.map(tx => String(tx.name || "").trim()));
      distinctSellers90D = sellerSet.size;
      const valueScore = !marketCapIsValid ? 0 : Math.max(0, Math.min(100, (netValueSold90D / marketCap) * 20000));
      const clusterBonus = distinctSellers90D >= 3 ? 30 : (distinctSellers90D === 2 ? 15 : 0);
      sellUrgencyScore = Math.max(0, Math.min(100, valueScore + clusterBonus));
      sellUrgencyScore = Math.round(sellUrgencyScore * 10) / 10;
    }

    // Raw_InsiderSignals row (existing-row-update-or-append)
    const signalRow = [ticker, buySignalScore, sellUrgencyScore, netValueBought12M, netValueSold90D, distinctSellers90D, new Date()];
    const existingSignalRow = signalsRowByTicker.get(ticker);
    if (existingSignalRow) {
      signalsSheet.getRange(existingSignalRow, 1, 1, INSIDER_SIGNALS_HEADERS.length).setValues([signalRow]);
    } else {
      newSignalRows.push(signalRow);
    }

    // Raw_InsidersCheckpoint row (existing-row-update-or-append)
    const checkpointRow = [ticker, new Date()];
    const existingCheckpointRow = checkpointRowByTicker.get(ticker);
    if (existingCheckpointRow) {
      checkpointSheet.getRange(existingCheckpointRow, 1, 1, INSIDERS_CHECKPOINT_HEADERS.length).setValues([checkpointRow]);
    } else {
      newCheckpointRows.push(checkpointRow);
    }
  }

  if (newInsiderRows.length > 0) {
    const startRow = insidersSheet.getLastRow() + 1;
    insidersSheet.getRange(startRow, 1, newInsiderRows.length, INSIDERS_HEADERS.length).setValues(newInsiderRows);
  }
  if (newSignalRows.length > 0) {
    const startRow = signalsSheet.getLastRow() + 1;
    signalsSheet.getRange(startRow, 1, newSignalRows.length, INSIDER_SIGNALS_HEADERS.length).setValues(newSignalRows);
  }
  if (newCheckpointRows.length > 0) {
    const startRow = checkpointSheet.getLastRow() + 1;
    checkpointSheet.getRange(startRow, 1, newCheckpointRows.length, INSIDERS_CHECKPOINT_HEADERS.length).setValues(newCheckpointRows);
  }

  return endIndex;
}
function finalize_Insiders() {
  const insidersSheet = _getOrCreateInsidersSheet();
  formatSheet(insidersSheet, INSIDERS_HEADERS.length);
  const signalsSheet = _getOrCreateInsiderSignalsSheet();
  formatSheet(signalsSheet, INSIDER_SIGNALS_HEADERS.length);
  SpreadsheetApp.getActive().toast("Insiders complete: all pending tickers processed.");
}
/*************************************************************
 * Finnhub insider-transactions call (Finnhub-only, no Finviz
 * involved here).
 *************************************************************/
function _fetchInsiderTransactions(ticker) {
  const token = _getFinnhubToken();
  const url = `https://finnhub.io/api/v1/stock/insider-transactions?symbol=${encodeURIComponent(ticker)}&token=${token}`;
  try {
    const res = safeFetch(url, {}, 3);
    const json = JSON.parse(res.getContentText());
    return Array.isArray(json.data) ? json.data : [];
  } catch (e) {
    Logger.log(`_fetchInsiderTransactions: error for ${ticker}: ${e}`);
    return [];
  }
}
/*************************************************************
 * Sheet helpers
 *************************************************************/
function _getOrCreateInsidersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.INSIDERS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.INSIDERS);
  return sheet;
}
function _ensureInsidersHeaders() {
  const sheet = _getOrCreateInsidersSheet();
  const headerRow = sheet.getRange(1, 1, 1, INSIDERS_HEADERS.length).getValues()[0];
  const isEmpty = headerRow.every(v => v === "" || v === null);
  if (isEmpty) {
    sheet.getRange(1, 1, 1, INSIDERS_HEADERS.length).setValues([INSIDERS_HEADERS]);
  }
}
function _getOrCreateInsiderSignalsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.INSIDER_SIGNALS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.INSIDER_SIGNALS);
  return sheet;
}
function _ensureInsiderSignalsHeaders() {
  const sheet = _getOrCreateInsiderSignalsSheet();
  const headerRow = sheet.getRange(1, 1, 1, INSIDER_SIGNALS_HEADERS.length).getValues()[0];
  const isEmpty = headerRow.every(v => v === "" || v === null);
  if (isEmpty) {
    sheet.getRange(1, 1, 1, INSIDER_SIGNALS_HEADERS.length).setValues([INSIDER_SIGNALS_HEADERS]);
  }
}
function _getOrCreateInsidersCheckpointSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.INSIDERS_CHECKPOINT);
  if (!sheet) sheet = ss.insertSheet(SHEETS.INSIDERS_CHECKPOINT);
  return sheet;
}
function _ensureInsidersCheckpointHeaders() {
  const sheet = _getOrCreateInsidersCheckpointSheet();
  const headerRow = sheet.getRange(1, 1, 1, INSIDERS_CHECKPOINT_HEADERS.length).getValues()[0];
  const isEmpty = headerRow.every(v => v === "" || v === null);
  if (isEmpty) {
    sheet.getRange(1, 1, 1, INSIDERS_CHECKPOINT_HEADERS.length).setValues([INSIDERS_CHECKPOINT_HEADERS]);
  }
}
function _writeInsidersPendingScratch(tickers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Raw_InsidersScratch");
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet("Raw_InsidersScratch");
  sheet.hideSheet();
  if (tickers.length > 0) {
    sheet.getRange(1, 1, tickers.length, 1).setValues(tickers.map(t => [t]));
  }
}
function _getInsidersPendingScratchSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Raw_InsidersScratch");
  if (!sheet) sheet = ss.insertSheet("Raw_InsidersScratch");
  return sheet;
}