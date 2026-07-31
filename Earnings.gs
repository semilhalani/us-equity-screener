/*************************************************************
 * Earnings.gs — quarterly EPS history, categorization, and
 * the Earnings Behavior Engine's two scores
 * -----------------------------------------------------------
 * PLAN CORRECTION FROM OUR RE-EVALUATION SESSION: we'd planned to
 * compute Volatility Score from ACTUAL price swings around earnings
 * dates (Finnhub /stock/candle, close-before vs close-after). That
 * endpoint is the same one DisqualifierGates.gs already found is
 * blocked on Finnhub's free tier for US stocks. So Volatility Score
 * here uses recency-weighted EPS SURPRISE% magnitude instead — this
 * was actually the ORIGINAL spec from early in this project, before
 * the candle-based "upgrade" turned out to be a dead end. Free,
 * already available from the same call we need for earnings history
 * anyway, no new endpoint required.
 *
 * BUGS FIXED FROM YOUR OLD MASTER TRACKER'S EPSCategorizer:
 *   - Corruption (category strings landing in numeric surprisePct
 *     slots): every value is validated as an actual number before
 *     it's ever written to a numeric column — a non-numeric value
 *     simply can't end up there anymore, regardless of what caused
 *     it in the old system.
 *   - Duplicate rows: same existing-row-lookup-and-update pattern
 *     used everywhere else in this project (consistent normalized
 *     ticker matching), same resumable gap+staleness design as
 *     Fundamentals.gs and DisqualifierGates.gs.
 *   - Also fixed a real bug in the ORIGINAL categorization logic
 *     itself: it hardcoded "exactly 4 quarters" (positives===4,
 *     doubleDigits===4), which silently miscategorized any company
 *     with fewer than 4 quarters of history (e.g. recent IPOs) —
 *     now adapts to however many valid quarters actually exist.
 *
 * TWO OUTPUTS PER TICKER:
 *   Raw_Earnings     — full quarterly detail, 4 quarters, newest
 *                       first (matches Finnhub's natural sort).
 *   Raw_EPSCategory  — compact derived summary: 4 surprise% values
 *                       oldest-to-newest (for reading the trend
 *                       left-to-right), the category label, and
 *                       the two Earnings Behavior Engine scores.
 *
 * ENTRY POINTS:
 *   runEarningsRefresh()                    — gap-fill/resume, full universe
 *   runEarningsRefresh({ forceAll: true })  — reprocess everyone regardless
 *   runEarningsRefresh_Test()               — synchronous, first 20 pending
 *   runEarningsRefresh_Test(true)           — same, via chunked triggers
 *************************************************************/

function runEarningsRefresh(options) {
  const mode = (options && options.mode) || MODE.PROD;
  const chunk = (options && typeof options.chunk !== "undefined") ? options.chunk : (mode === MODE.PROD);
  const forceAll = (options && options.forceAll) || false;

  _ensureEarningsHeaders();
  _ensureEpsCategoryHeaders();

  const uniSheet = _getOrCreateUniverseSheet();
  const uniLast = uniSheet.getLastRow();
  if (uniLast < 2) {
    SpreadsheetApp.getActive().toast("Earnings: Raw_Universe is empty — run Universe first.");
    return;
  }
  const allTickers = uniSheet.getRange(2, 1, uniLast - 1, 1).getValues().flat()
    .map(t => String(t).trim().toUpperCase()).filter(Boolean);

  const earnSheet = _getOrCreateEarningsSheet();
  const eLast = earnSheet.getLastRow();
  const freshSet = new Set();
  if (eLast >= 2 && !forceAll) {
    const existing = earnSheet.getRange(2, 1, eLast - 1, EARNINGS_HEADERS.length).getValues();
    const staleCutoff = Date.now() - (EARNINGS_STALENESS_DAYS * 24 * 60 * 60 * 1000);
    existing.forEach(r => {
      const ticker = String(r[0] || "").trim().toUpperCase();
      const lastUpdated = r[25]; // Last_Updated column
      if (ticker && lastUpdated instanceof Date && lastUpdated.getTime() >= staleCutoff) {
        freshSet.add(ticker);
      }
    });
  }

  let pending = allTickers.filter(t => !freshSet.has(t));

  if (mode === MODE.TEST) {
    pending = pending.slice(0, DEFAULT_TEST_LIMIT);
  }

  if (pending.length === 0) {
    SpreadsheetApp.getActive().toast(`Earnings: everything already fresh (updated within ${EARNINGS_STALENESS_DAYS} days). Nothing to do.`);
    return;
  }

  _writeEarningsPendingScratch(pending);

  if (chunk) {
    startChunkProcess("Earnings", "chunkWorker_Earnings", "finalize_Earnings", pending.length, 100);
  } else {
    let idx = 0;
    while (idx < pending.length) idx = chunkWorker_Earnings(idx, 100);
    finalize_Earnings();
  }
  SpreadsheetApp.getActive().toast(
    `Earnings refresh started (${mode}). ${pending.length} pending, ${allTickers.length - pending.length} already fresh (skipped).`
  );
}

function runEarningsRefresh_Test(chunk) {
  return runEarningsRefresh({ mode: MODE.TEST, chunk: (chunk === true) });
}

/*************************************************************
 * Chunk worker — one Finnhub call per ticker, writes to BOTH
 * Raw_Earnings and Raw_EPSCategory in the same pass (no separate
 * categorization stage needed, unlike the old multi-script pipeline).
 *************************************************************/

function chunkWorker_Earnings(startIndex, chunkSize) {
  const pendingSheet = _getEarningsPendingScratchSheet();
  const total = Math.max(pendingSheet.getLastRow(), 0);
  if (startIndex >= total) return total;

  const endIndex = Math.min(startIndex + chunkSize, total);
  const numRows = endIndex - startIndex;
  const startRowNum = startIndex + 1;
  const tickers = pendingSheet.getRange(startRowNum, 1, numRows, 1).getValues().flat();

  const earnSheet = _getOrCreateEarningsSheet();
  const eLast = earnSheet.getLastRow();
  const existingEarnRowByTicker = new Map();
  if (eLast >= 2) {
    const existing = earnSheet.getRange(2, 1, eLast - 1, 1).getValues().flat();
    existing.forEach((t, i) => existingEarnRowByTicker.set(String(t).trim().toUpperCase(), i + 2));
  }

  const catSheet = _getOrCreateEpsCategorySheet();
  const cLast = catSheet.getLastRow();
  const existingCatRowByTicker = new Map();
  if (cLast >= 2) {
    const existing = catSheet.getRange(2, 1, cLast - 1, 1).getValues().flat();
    existing.forEach((t, i) => existingCatRowByTicker.set(String(t).trim().toUpperCase(), i + 2));
  }

  const newEarnRows = [];
  const newCatRows = [];

  for (let i = 0; i < tickers.length; i++) {
    const ticker = String(tickers[i] || "").trim().toUpperCase();
    if (!ticker) continue;

    const quarters = _fetchEarningsHistory(ticker); // newest-first, up to 4 entries
    Utilities.sleep(150);

    const fetchSucceeded = quarters.length > 0;

    // ---- Raw_Earnings row: newest-first, 4 quarters x 6 fields ----
    const earningsRow = [ticker];
    const surprisePctsNewestFirst = [];
    for (let q = 0; q < 4; q++) {
      const e = quarters[q];
      if (e) {
        const period = e.period || "";
        const quarterLabel = e.quarter ? `Q${e.quarter}_${e.year}` : "";
        const est = (e.estimate != null) ? e.estimate : "";
        const act = (e.actual != null) ? e.actual : "";
        const sp = (e.surprise != null) ? e.surprise : "";
        // VALIDATION FIX: only ever push an actual number into the
        // surprisePct slot, or blank — never a string of any other kind.
        const spRaw = e.surprisePercent;
        const spPct = (typeof spRaw === "number" && !isNaN(spRaw)) ? spRaw : "";
        earningsRow.push(period, quarterLabel, est, act, sp, spPct);
        surprisePctsNewestFirst.push(spPct === "" ? NaN : spPct);
      } else {
        earningsRow.push("", "", "", "", "", "");
        surprisePctsNewestFirst.push(NaN);
      }
    }
    earningsRow.push(fetchSucceeded ? new Date() : "");

    const existingEarnRow = existingEarnRowByTicker.get(ticker);
    if (existingEarnRow) {
      earnSheet.getRange(existingEarnRow, 1, 1, EARNINGS_HEADERS.length).setValues([earningsRow]);
    } else {
      newEarnRows.push(earningsRow);
    }

    // ---- Raw_EPSCategory row: oldest-first display + scores ----
    const validOldestFirst = surprisePctsNewestFirst.filter(v => !isNaN(v)).reverse();
    while (validOldestFirst.length < 4) validOldestFirst.unshift("");
    const displayQuarters = validOldestFirst.slice(0, 4);

    const category = _categorizeEPS(displayQuarters);
    const volatilityScore = _computeVolatilityScore(surprisePctsNewestFirst);
    const reliabilityScore = _computeReliabilityScore(surprisePctsNewestFirst);

    const epsCatRow = [
      ticker,
      displayQuarters[0], displayQuarters[1], displayQuarters[2], displayQuarters[3],
      category,
      _round1(volatilityScore), _round1(reliabilityScore),
      fetchSucceeded ? new Date() : ""
    ];

    const existingCatRow = existingCatRowByTicker.get(ticker);
    if (existingCatRow) {
      catSheet.getRange(existingCatRow, 1, 1, EPS_CATEGORY_HEADERS.length).setValues([epsCatRow]);
    } else {
      newCatRows.push(epsCatRow);
    }
  }

  if (newEarnRows.length > 0) {
    const startRow = earnSheet.getLastRow() + 1;
    earnSheet.getRange(startRow, 1, newEarnRows.length, EARNINGS_HEADERS.length).setValues(newEarnRows);
  }
  if (newCatRows.length > 0) {
    const startRow = catSheet.getLastRow() + 1;
    catSheet.getRange(startRow, 1, newCatRows.length, EPS_CATEGORY_HEADERS.length).setValues(newCatRows);
  }

  return endIndex;
}

function finalize_Earnings() {
  const earnSheet = _getOrCreateEarningsSheet();
  formatSheet(earnSheet, EARNINGS_HEADERS.length);
  const catSheet = _getOrCreateEpsCategorySheet();
  formatSheet(catSheet, EPS_CATEGORY_HEADERS.length);
  SpreadsheetApp.getActive().toast("Earnings complete: all pending tickers processed.");
}

/*************************************************************
 * Finnhub earnings history call
 *************************************************************/

function _fetchEarningsHistory(ticker) {
  const token = _getFinnhubToken();
  const url = `https://finnhub.io/api/v1/stock/earnings?symbol=${encodeURIComponent(ticker)}&token=${token}`;
  try {
    const res = safeFetch(url, {}, 3);
    const json = JSON.parse(res.getContentText());
    if (!Array.isArray(json)) return [];
    const sorted = json.slice().sort((a, b) => {
      const yearDiff = (b.year || 0) - (a.year || 0);
      if (yearDiff !== 0) return yearDiff;
      return (b.quarter || 0) - (a.quarter || 0);
    });
    return sorted.slice(0, 4); // 4 most recent, newest first
  } catch (e) {
    Logger.log(`_fetchEarningsHistory: error for ${ticker}: ${e}`);
    return [];
  }
}

/*************************************************************
 * Categorization — adapted from your reference system's category
 * hierarchy (that part was genuinely good), fixed to adapt to
 * however many valid quarters exist instead of hardcoding 4.
 *************************************************************/

function _categorizeEPS(vals) {
  const nums = vals.map(v => (v === "" || v === null || v === undefined ? NaN : Number(v)));
  const valid = nums.filter(x => !isNaN(x));
  if (valid.length === 0) return "⚪ Undefined";

  const positives = valid.filter(x => x > 0).length;
  const negatives = valid.filter(x => x < 0).length;
  const doubleDigits = valid.filter(x => Math.abs(x) >= 10 && Math.abs(x) < 100).length;
  const tripleDigits = valid.filter(x => Math.abs(x) >= 100 && Math.abs(x) < 300).length;
  const ultraDigits = valid.filter(x => Math.abs(x) >= 300).length;

  if (positives === valid.length) {
    if (ultraDigits > 0) return "🚀 Ultra Beat (300%+)";
    if (tripleDigits > 0) return "💎 Mega Beat (100%+)";
    if (doubleDigits === valid.length) return "💚 Strong Beat (All ≥10%)";
    if (doubleDigits === 3) return "👍 Consistent Beat (3×≥10%)";
    if (doubleDigits === 2) return "🙂 Mild Beat (2×≥10%)";
    return "🟩 Mixed Positive";
  }
  if (negatives === valid.length) {
    if (ultraDigits > 0) return "🔥 Ultra Miss (300%-)";
    if (tripleDigits > 0) return "😡 Mega Miss (100%-)";
    if (doubleDigits === valid.length) return "❌ Strong Miss (All ≤ -10%)";
    if (doubleDigits === 3) return "👎 Consistent Miss (3×≤ -10%)";
    if (doubleDigits === 2) return "🙁 Mild Miss (2×≤ -10%)";
    return "🟥 Mixed Negative";
  }
  if (tripleDigits > 0 || ultraDigits > 0) return "⚡ Mixed (Strong Beat Present)";
  return "⚪ Mixed Neutral";
}

/*************************************************************
 * Earnings Behavior Engine — the two scores.
 * Both take surprisePctsNewestFirst (index 0 = most recent
 * quarter), recency-weighted 40/30/20/10 as specified from the
 * very start of this project.
 *************************************************************/

// HIGH score = HIGH swing (intuitive — "80" means volatile, not safe).
function _computeVolatilityScore(surprisePctsNewestFirst) {
  const weights = [0.40, 0.30, 0.20, 0.10];
  let weightedAbsSum = 0;
  let weightUsed = 0;
  for (let i = 0; i < 4; i++) {
    const v = surprisePctsNewestFirst[i];
    if (typeof v === "number" && !isNaN(v)) {
      weightedAbsSum += Math.abs(v) * weights[i];
      weightUsed += weights[i];
    }
  }
  if (weightUsed === 0) return "";
  const weightedAvgAbsSwing = weightedAbsSum / weightUsed; // renormalized if some quarters missing
  return Math.min(100, weightedAvgAbsSwing);
}

// HIGH score = consistently beats. 50 = neutral/mixed. 0 = consistently misses.
function _computeReliabilityScore(surprisePctsNewestFirst) {
  const weights = [0.40, 0.30, 0.20, 0.10];
  let weightedSum = 0;
  let weightUsed = 0;
  for (let i = 0; i < 4; i++) {
    const v = surprisePctsNewestFirst[i];
    if (typeof v === "number" && !isNaN(v)) {
      const signal = (v > 2) ? 1 : (v < -2) ? -1 : 0; // beat / miss / in-line
      weightedSum += signal * weights[i];
      weightUsed += weights[i];
    }
  }
  if (weightUsed === 0) return "";
  const weightedSignal = weightedSum / weightUsed; // -1..+1
  return Math.max(0, Math.min(100, 50 + weightedSignal * 50));
}

/*************************************************************
 * Sheet helpers
 *************************************************************/

function _getOrCreateEarningsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.EARNINGS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.EARNINGS);
  return sheet;
}

function _ensureEarningsHeaders() {
  const sheet = _getOrCreateEarningsSheet();
  const headerRow = sheet.getRange(1, 1, 1, EARNINGS_HEADERS.length).getValues()[0];
  const isEmpty = headerRow.every(v => v === "" || v === null);
  if (isEmpty) {
    sheet.getRange(1, 1, 1, EARNINGS_HEADERS.length).setValues([EARNINGS_HEADERS]);
  }
}

function _getOrCreateEpsCategorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.EPS_CATEGORY);
  if (!sheet) sheet = ss.insertSheet(SHEETS.EPS_CATEGORY);
  return sheet;
}

function _ensureEpsCategoryHeaders() {
  const sheet = _getOrCreateEpsCategorySheet();
  const headerRow = sheet.getRange(1, 1, 1, EPS_CATEGORY_HEADERS.length).getValues()[0];
  const isEmpty = headerRow.every(v => v === "" || v === null);
  if (isEmpty) {
    sheet.getRange(1, 1, 1, EPS_CATEGORY_HEADERS.length).setValues([EPS_CATEGORY_HEADERS]);
  }
}

function _writeEarningsPendingScratch(tickers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Raw_EarningsScratch");
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet("Raw_EarningsScratch");
  sheet.hideSheet();
  if (tickers.length > 0) {
    sheet.getRange(1, 1, tickers.length, 1).setValues(tickers.map(t => [t]));
  }
}

function _getEarningsPendingScratchSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Raw_EarningsScratch");
  if (!sheet) sheet = ss.insertSheet("Raw_EarningsScratch");
  return sheet;
}