/*************************************************************
 * DisqualifierGates.gs — hard pass/fail gates, pre-scoring (v4)
 * -----------------------------------------------------------
 * v4 CHANGE (two agreed fixes, combined since both touch row shape):
 *
 * 1. Dilution's free "accumulate our own Shares_Outstanding snapshots"
 *    method (Raw_SharesHistory) is retired entirely. It only would have
 *    started paying off ~11 months in, and in the meantime grew by
 *    ~5,336 rows on every single Universe run, forever — a real
 *    contributor to hitting the workbook's 10M-cell ceiling. The
 *    financials-reported bootstrap (_computeDilutionFromReports) already
 *    gives validated, correct real answers immediately (confirmed: VLO
 *    -5.1% YoY buyback correctly detected as OK, ANVS +59.8% YoY dilution
 *    correctly flagged, EVGO +136.6% correctly disqualified) — it's now
 *    the ONLY dilution method, always called, for every tier, no waiting.
 *    In practice this changes nothing about API call volume: since
 *    Raw_SharesHistory never had 11 months of data to begin with, every
 *    dilution check was already falling through to this same bootstrap.
 *
 * 2. Packed sentence-style status columns (e.g. "Diluting (136.6% YoY,
 *    bootstrapped)") are replaced with clean columns per gate: a real
 *    number, a pass/fail boolean, and a method flag explaining what
 *    produced (or didn't produce) that number. The old _Notes columns
 *    are gone from this raw sheet entirely — plain-English explanations
 *    belong in a future presentation sheet built FROM these clean
 *    numbers, not duplicated as stored text here.
 *
 * WHAT'S IMPLEMENTED TODAY (all 5 planned gates, in some form):
 *   - Liquidity: Avg Volume + Price from Finviz's quote-page snapshot.
 *   - Listing age: computed in Universe.gs (Flag_NewListing).
 *   - Dilution: financials-reported bootstrap, every tier, always.
 *   - Zero-revenue-quarters: Moonshot tier only, flag-only (doesn't
 *     disqualify alone — see reasoning in chunkWorker below).
 *   - Cash runway: Moonshot tier only, disqualifies on its own.
 *
 * ENTRY POINTS:
 *   runDisqualifierGatesRefresh()                   — gap-fill/resume
 *   runDisqualifierGatesRefresh({ forceAll: true })  — reprocess everyone
 *   runDisqualifierGatesRefresh_Test()               — synchronous, first 20 pending
 *   runDisqualifierGatesRefresh_Test(true)           — same, via chunked triggers
 *************************************************************/
function runDisqualifierGatesRefresh(options) {
  const mode = (options && options.mode) || MODE.PROD;
  const chunk = (options && typeof options.chunk !== "undefined") ? options.chunk : (mode === MODE.PROD);
  const forceAll = (options && options.forceAll) || false;
  _ensureDisqualifierGatesHeaders();
  const uniSheet = _getOrCreateUniverseSheet();
  const uniLast = uniSheet.getLastRow();
  if (uniLast < 2) {
    SpreadsheetApp.getActive().toast("DisqualifierGates: Raw_Universe is empty — run Universe first.");
    return;
  }
  const allTickers = uniSheet.getRange(2, 1, uniLast - 1, 1).getValues().flat()
    .map(t => String(t).trim().toUpperCase()).filter(Boolean);
  const gatesSheet = _getOrCreateDisqualifierGatesSheet();
  const gLast = gatesSheet.getLastRow();
  const freshSet = new Set();
  if (gLast >= 2 && !forceAll) {
    const existing = gatesSheet.getRange(2, 1, gLast - 1, DISQUALIFIER_GATES_HEADERS.length).getValues();
    const staleCutoff = Date.now() - (DISQUALIFIER_GATES_STALENESS_DAYS * 24 * 60 * 60 * 1000);
    existing.forEach(r => {
      const ticker = String(r[0] || "").trim().toUpperCase();
      // ---- DisqualifierGates.gs v4: Last_Updated is now the last of 18 clean columns (index 17) ----
      const lastUpdated = r[17]; // Last_Updated column
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
    SpreadsheetApp.getActive().toast(`DisqualifierGates: everything already fresh (updated within ${DISQUALIFIER_GATES_STALENESS_DAYS} days). Nothing to do.`);
    return;
  }
  _writeGatesPendingScratch(pending);
  const gateChunkSize = 80;
  if (chunk) {
    startChunkProcess("DisqualifierGates", "chunkWorker_DisqualifierGates", "finalize_DisqualifierGates", pending.length, gateChunkSize);
  } else {
    let idx = 0;
    while (idx < pending.length) idx = chunkWorker_DisqualifierGates(idx, gateChunkSize);
    finalize_DisqualifierGates();
  }
  SpreadsheetApp.getActive().toast(
    `DisqualifierGates refresh started (${mode}). ${pending.length} pending, ${allTickers.length - pending.length} already fresh (skipped).`
  );
}
function runDisqualifierGatesRefresh_Test(chunk) {
  return runDisqualifierGatesRefresh({ mode: MODE.TEST, chunk: (chunk === true) });
}
/*************************************************************
 * Chunk worker
 *************************************************************/
function chunkWorker_DisqualifierGates(startIndex, chunkSize) {
  const pendingSheet = _getGatesPendingScratchSheet();
  const total = Math.max(pendingSheet.getLastRow(), 0);
  if (startIndex >= total) return total;
  const endIndex = Math.min(startIndex + chunkSize, total);
  const numRows = endIndex - startIndex;
  const startRowNum = startIndex + 1;
  const tickers = pendingSheet.getRange(startRowNum, 1, numRows, 1).getValues().flat();
  const uniSheet = _getOrCreateUniverseSheet();
  const uniLast = uniSheet.getLastRow();
  const uniInfoByTicker = new Map();
  if (uniLast >= 2) {
    const uniData = uniSheet.getRange(2, 1, uniLast - 1, UNIVERSE_HEADERS.length).getValues();
    uniData.forEach(r => {
      const t = String(r[0] || "").trim().toUpperCase();
      if (t) uniInfoByTicker.set(t, { tier: r[7], flagNewListing: r[12] });
    });
  }
  const gatesSheet = _getOrCreateDisqualifierGatesSheet();
  const gLast = gatesSheet.getLastRow();
  const existingRowByTicker = new Map();
  if (gLast >= 2) {
    const existingTickers = gatesSheet.getRange(2, 1, gLast - 1, 1).getValues().flat();
    existingTickers.forEach((t, i) => existingRowByTicker.set(String(t).trim().toUpperCase(), i + 2));
  }
  const newRows = [];
  for (let i = 0; i < tickers.length; i++) {
    const ticker = String(tickers[i] || "").trim().toUpperCase();
    if (!ticker) continue;
    const info = uniInfoByTicker.get(ticker) || {};
    const tier = info.tier || "";
    const flagNewListing = info.flagNewListing;
    const snap = _scrapeFinvizLiquiditySnapshot(ticker);
    Utilities.sleep(250);
    const avgDollarVolume = (typeof snap.avgVolumeShares === "number" && typeof snap.price === "number")
      ? snap.avgVolumeShares * snap.price
      : "";
    const liquidityThreshold = (tier === TIER.MOONSHOT) ? DISQUALIFIERS.MIN_AVG_DOLLAR_VOLUME_MOONSHOT : DISQUALIFIERS.MIN_AVG_DOLLAR_VOLUME;
    const liquidityPass = (typeof avgDollarVolume === "number") ? (avgDollarVolume >= liquidityThreshold) : "";
    const listingAgePass = (flagNewListing === true) ? false : true;
    // financials-reported is fetched AT MOST ONCE per ticker per run. Every
    // tier needs it now for Dilution; Moonshot tickers also need it for
    // Zero-Revenue/Cash-Runway — shared via ensureReports() either way,
    // never double-fetched for the same ticker in the same run.
    let reports = null;
    let reportsFetched = false;
    const ensureReports = () => {
      if (!reportsFetched) {
        reports = _fetchFinancialsReported(ticker);
        reportsFetched = true;
        Utilities.sleep(150);
      }
      return reports;
    };
    // DILUTION GATE — all tiers, financials-reported bootstrap, always.
    const dilution = _computeDilutionFromReports(ensureReports());
    // ZERO-REVENUE-QUARTERS + CASH-RUNWAY — Moonshot tier only.
    let zeroRevenue = { count: "", quartersConsidered: "", pass: "", method: "Not_Applicable" };
    let cashRunway = { months: "", pass: "", method: "Not_Applicable" };
    if (tier === TIER.MOONSHOT) {
      const financials = _computeZeroRevenueAndCashRunwayFromReports(ensureReports());
      zeroRevenue = {
        count: financials.zeroRevenueCount,
        quartersConsidered: financials.zeroRevenueQuartersConsidered,
        pass: financials.zeroRevenuePass,
        method: financials.zeroRevenueMethod
      };
      cashRunway = {
        months: financials.cashRunwayMonths,
        pass: financials.cashRunwayPass,
        method: financials.cashRunwayMethod
      };
    }
    const reasons = [];
    if (liquidityPass === false) reasons.push("Low_Liquidity");
    if (listingAgePass === false) reasons.push("Too_New");
    if (dilution.pass === false) reasons.push("Excessive_Dilution");
    // Zero-Revenue-Quarters is intentionally flag-only, not disqualifying,
    // on its own: pre-revenue is often the entire premise of a legitimate
    // Moonshot bet. It only counts toward disqualification when paired
    // with a failing cash runway (no revenue AND no cash cushion signals
    // imminent forced/dilutive financing).
    if (cashRunway.pass === false) {
      reasons.push(zeroRevenue.pass === false ? "Zero_Revenue_And_Cash_Runway_Risk" : "Cash_Runway_Too_Short");
    }
    const overallDisqualified = (liquidityPass === false) || (listingAgePass === false) ||
      (dilution.pass === false) || (cashRunway.pass === false);
    const fetchSucceeded = (typeof snap.avgVolumeShares === "number" || typeof snap.price === "number");
    const row = [
      ticker,
      tier,
      avgDollarVolume,
      liquidityPass,
      listingAgePass,
      dilution.yoyPct,
      dilution.pass,
      dilution.method,
      zeroRevenue.count,
      zeroRevenue.quartersConsidered,
      zeroRevenue.pass,
      zeroRevenue.method,
      cashRunway.months,
      cashRunway.pass,
      cashRunway.method,
      overallDisqualified,
      reasons.join(", "),
      fetchSucceeded ? new Date() : ""
    ];
    const existingRow = existingRowByTicker.get(ticker);
    if (existingRow) {
      gatesSheet.getRange(existingRow, 1, 1, DISQUALIFIER_GATES_HEADERS.length).setValues([row]);
    } else {
      newRows.push(row);
    }
  }
  if (newRows.length > 0) {
    const startRow = gatesSheet.getLastRow() + 1;
    gatesSheet.getRange(startRow, 1, newRows.length, DISQUALIFIER_GATES_HEADERS.length).setValues(newRows);
  }
  return endIndex;
}
function finalize_DisqualifierGates() {
  const sheet = _getOrCreateDisqualifierGatesSheet();
  formatSheet(sheet, DISQUALIFIER_GATES_HEADERS.length);
  SpreadsheetApp.getActive().toast("DisqualifierGates complete: all pending tickers evaluated.");
}
/*************************************************************
 * DILUTION GATE
 * financials-reported bootstrap only — the free own-history
 * snapshot method (Raw_SharesHistory) has been retired entirely.
 *************************************************************/
function _computeDilutionFromReports(reports) {
  const insufficient = { yoyPct: "", pass: "", method: "Insufficient_History" };
  if (!reports || reports.length === 0) return insufficient;
  const sorted = reports.slice().sort((a, b) => new Date(b.endDate) - new Date(a.endDate));
  // Real test data (VLO, ANVS) showed neither tags a clean numeric "shares
  // outstanding" concept on the balance sheet — VLO only mentions the count
  // inside a text label, never as its own value. The income statement's
  // weighted-average-shares figure (every company needs it to compute EPS)
  // is far more universally present as a real number, so that's the primary
  // source, with the old balance-sheet concepts kept as a fallback for
  // whichever company DOES tag it that way.
  const sharesPatterns = [
    "weightedaveragenumberofsharesoutstandingbasic",
    "weightedaveragenumberofdilutedsharesoutstanding",
    "commonstocksharesoutstanding",
    "commonstocksharesissued"
  ];
  function findShares(rep) {
    const ic = rep.report && rep.report.ic;
    const bs = rep.report && rep.report.bs;
    const fromIc = _extractConceptValue(ic, sharesPatterns);
    if (fromIc !== null) return fromIc;
    return _extractConceptValue(bs, sharesPatterns);
  }
  let latest = null;
  for (const rep of sorted) {
    const shares = findShares(rep);
    if (shares !== null && shares > 0) { latest = { date: new Date(rep.endDate), shares }; break; }
  }
  if (!latest) return insufficient;
  let baseline = null;
  for (const rep of sorted) {
    const repDate = new Date(rep.endDate);
    const gapDays = (latest.date.getTime() - repDate.getTime()) / (1000 * 60 * 60 * 24);
    if (gapDays < DILUTION_MIN_SNAPSHOT_GAP_DAYS) continue;
    const shares = findShares(rep);
    if (shares !== null && shares > 0) { baseline = { date: repDate, shares }; break; }
  }
  if (!baseline) return insufficient;
  const yoyGrowth = (latest.shares - baseline.shares) / baseline.shares;
  const pass = yoyGrowth <= DISQUALIFIERS.MAX_SHARES_OUTSTANDING_YOY_GROWTH;
  const yoyPct = Number((yoyGrowth * 100).toFixed(1));
  return { yoyPct, pass, method: "Bootstrapped" };
}
/*************************************************************
 * ZERO-REVENUE-QUARTERS + CASH-RUNWAY — Moonshot tier only.
 *************************************************************/
function _fetchFinancialsReported(ticker) {
  const token = _getFinnhubToken();
  const url = `https://finnhub.io/api/v1/stock/financials-reported?symbol=${encodeURIComponent(ticker)}&freq=quarterly&token=${token}`;
  try {
    const res = safeFetch(url, {}, 3);
    const json = JSON.parse(res.getContentText());
    return Array.isArray(json.data) ? json.data : [];
  } catch (e) {
    Logger.log(`_fetchFinancialsReported: error for ${ticker}: ${e}`);
    return [];
  }
}
function _extractConceptValue(items, patterns) {
  if (!Array.isArray(items)) return null;
  for (const it of items) {
    const concept = String(it.concept || "").toLowerCase();
    const label = String(it.label || "").toLowerCase();
    if (patterns.some(p => concept.includes(p) || label.includes(p))) {
      const v = Number(it.value);
      if (!isNaN(v)) return v;
    }
  }
  return null;
}
function _computeZeroRevenueAndCashRunwayFromReports(reports) {
  const insufficientZeroRevenue = { zeroRevenueCount: "", zeroRevenueQuartersConsidered: "", zeroRevenuePass: "", zeroRevenueMethod: "Insufficient_Data" };
  const insufficientCashRunway = { cashRunwayMonths: "", cashRunwayPass: "", cashRunwayMethod: "Insufficient_Data" };
  if (!reports || reports.length === 0) {
    return Object.assign({}, insufficientZeroRevenue, insufficientCashRunway);
  }
  const sorted = reports.slice().sort((a, b) => new Date(b.endDate) - new Date(a.endDate));
  const recentQuarters = sorted.slice(0, 8);
  const revenuePatterns = ["revenuefromcontractwithcustomer", "revenues", "totalrevenue", "salesrevenuenet"];
  let quartersConsidered = 0;
  let zeroRevenueCount = 0;
  recentQuarters.forEach(rep => {
    const ic = rep.report && rep.report.ic;
    if (!Array.isArray(ic) || ic.length === 0) return;
    const rev = _extractConceptValue(ic, revenuePatterns);
    quartersConsidered++;
    if (rev === null || Math.abs(rev) < 1) zeroRevenueCount++;
  });
  const zeroRevenueResult = (quartersConsidered === 0)
    ? insufficientZeroRevenue
    : {
        zeroRevenueCount,
        zeroRevenueQuartersConsidered: quartersConsidered,
        zeroRevenuePass: zeroRevenueCount < DISQUALIFIERS.MAX_ZERO_REVENUE_QUARTERS,
        zeroRevenueMethod: "Computed"
      };
  const latest = sorted[0];
  const bs = latest.report && latest.report.bs;
  const cf = latest.report && latest.report.cf;
  const cashPatterns = ["cashandcashequivalentsatcarryingvalue", "cashandcashequivalents"];
  const opCashFlowPatterns = ["netcashprovidedbyusedinoperatingactivities", "netcashprovidedbyoperatingactivities"];
  const cash = _extractConceptValue(bs, cashPatterns);
  const opCashFlowQuarterly = _extractConceptValue(cf, opCashFlowPatterns);
  let cashRunwayResult;
  if (cash === null || opCashFlowQuarterly === null) {
    cashRunwayResult = insufficientCashRunway;
  } else if (opCashFlowQuarterly >= 0) {
    cashRunwayResult = { cashRunwayMonths: "", cashRunwayPass: true, cashRunwayMethod: "Cash_Flow_Positive" };
  } else {
    const monthlyBurn = Math.abs(opCashFlowQuarterly) / 3;
    const runwayMonths = monthlyBurn > 0 ? (cash / monthlyBurn) : Infinity;
    const cashRunwayPass = runwayMonths >= DISQUALIFIERS.MIN_CASH_RUNWAY_MONTHS_MOONSHOT;
    cashRunwayResult = {
      cashRunwayMonths: isFinite(runwayMonths) ? Number(runwayMonths.toFixed(1)) : "",
      cashRunwayPass,
      cashRunwayMethod: "Burn_Rate_Computed"
    };
  }
  return Object.assign({}, zeroRevenueResult, cashRunwayResult);
}
// Convenience wrapper — kept so you can call this by ticker directly from
// the Apps Script editor for a quick check, without pre-fetching reports.
function _computeZeroRevenueAndCashRunway(ticker) {
  return _computeZeroRevenueAndCashRunwayFromReports(_fetchFinancialsReported(ticker));
}
/*************************************************************
 * DIAGNOSTIC — run on 1-2 real tickers to sanity-check both the
 * Moonshot gates AND the dilution bootstrap before a full run.
 *************************************************************/
function testFinancialsReportedParsing(ticker) {
  ticker = String(ticker || "").trim().toUpperCase();
  if (!ticker) { Logger.log("testFinancialsReportedParsing: pass a ticker, e.g. testFinancialsReportedParsing('SOME_TICKER')"); return; }
  const reports = _fetchFinancialsReported(ticker);
  Logger.log(`${ticker}: fetched ${reports.length} report period(s).`);
  if (reports.length === 0) {
    Logger.log(`${ticker}: no reports returned at all.`);
    return;
  }
  const sorted = reports.slice().sort((a, b) => new Date(b.endDate) - new Date(a.endDate));
  const latest = sorted[0];
  Logger.log(`${ticker}: most recent period ${latest.startDate} → ${latest.endDate} (form ${latest.form}).`);
  const ic = latest.report && latest.report.ic;
  const bs = latest.report && latest.report.bs;
  const cf = latest.report && latest.report.cf;
  Logger.log(`${ticker}: ic has ${ic ? ic.length : 0} line items, bs has ${bs ? bs.length : 0}, cf has ${cf ? cf.length : 0}.`);
  const result = _computeZeroRevenueAndCashRunwayFromReports(reports);
  Logger.log(`${ticker}: Zero_Revenue_Method=${result.zeroRevenueMethod}, Count=${result.zeroRevenueCount}/${result.zeroRevenueQuartersConsidered}, Pass=${result.zeroRevenuePass}`);
  Logger.log(`${ticker}: Cash_Runway_Method=${result.cashRunwayMethod}, Months=${result.cashRunwayMonths}, Pass=${result.cashRunwayPass}`);
  const dilutionFallback = _computeDilutionFromReports(reports);
  if (dilutionFallback.method === "Bootstrapped") {
    Logger.log(`${ticker}: Dilution = ${dilutionFallback.yoyPct}% YoY, Pass=${dilutionFallback.pass} (bootstrapped)`);
  } else {
    Logger.log(`${ticker}: Dilution bootstrap found nothing usable (method=${dilutionFallback.method}) — dumping bs[] concepts from the most recent AND oldest report to see what's actually there.`);
    const latestBs = sorted[0].report && sorted[0].report.bs;
    const oldestBs = sorted[sorted.length - 1].report && sorted[sorted.length - 1].report.bs;
    if (latestBs) Logger.log(`${ticker}: latest bs[] concepts: ${latestBs.map(x => x.concept).join(", ")}`);
    if (oldestBs) Logger.log(`${ticker}: oldest bs[] concepts: ${oldestBs.map(x => x.concept).join(", ")}`);
  }
  if (result.zeroRevenueMethod === "Insufficient_Data" && ic) {
    Logger.log(`${ticker}: revenue pattern miss — full ic[] concepts: ${ic.map(x => x.concept).join(", ")}`);
  }
  if (result.cashRunwayMethod === "Insufficient_Data") {
    if (bs) Logger.log(`${ticker}: cash pattern check — full bs[] concepts: ${bs.map(x => x.concept).join(", ")}`);
    if (cf) Logger.log(`${ticker}: op-cash-flow pattern check — full cf[] concepts: ${cf.map(x => x.concept).join(", ")}`);
  }
}
/*************************************************************
 * Finviz snapshot scrape — Avg Volume + Price
 *************************************************************/
function _scrapeFinvizLiquiditySnapshot(ticker) {
  const finvizSymbol = ticker.replace(/\./g, "-");
  const url = `https://finviz.com/stock?t=${encodeURIComponent(finvizSymbol)}`;
  let html = "";
  try {
    const res = safeFetch(url, { headers: { "User-Agent": "Mozilla/5.0", "Accept": "text/html" } }, 3);
    html = res.getContentText();
  } catch (e) {
    Logger.log(`_scrapeFinvizLiquiditySnapshot: fetch failed for ${ticker}: ${e}`);
    return { avgVolumeShares: "", price: "" };
  }
  const avgVolMatch = html.match(/Avg Volume[\s\S]*?<b>([^<]+)<\/b>/i);
  const avgVolumeShares = avgVolMatch ? _parseFinvizAbbreviatedNumber(htmlDecode(avgVolMatch[1])) : "";
  if (!avgVolMatch) {
    const hasLabel = html.includes("Avg Volume");
    Logger.log(`_scrapeFinvizLiquiditySnapshot: Avg Volume not matched for ${ticker}. Label text present: ${hasLabel}. HTML length=${html.length}.`);
  }
  const priceMatch = html.match(/Price[\s\S]*?<b>([^<]+)<\/b>/i);
  const priceRaw = priceMatch ? parseFloat(String(htmlDecode(priceMatch[1])).replace(/,/g, "")) : NaN;
  return { avgVolumeShares, price: isNaN(priceRaw) ? "" : priceRaw };
}
function _parseFinvizAbbreviatedNumber(s) {
  if (!s) return "";
  const cleaned = String(s).trim().replace(/,/g, "");
  const m = cleaned.match(/^(-?[\d.]+)\s*([KMB])?$/i);
  if (!m) {
    const plain = parseFloat(cleaned);
    return isNaN(plain) ? "" : plain;
  }
  const num = parseFloat(m[1]);
  if (isNaN(num)) return "";
  const suffix = (m[2] || "").toUpperCase();
  const mult = suffix === "B" ? 1e9 : suffix === "M" ? 1e6 : suffix === "K" ? 1e3 : 1;
  return num * mult;
}
/*************************************************************
 * Sheet helpers
 *************************************************************/
function _getOrCreateDisqualifierGatesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.DISQUALIFIER_GATES);
  if (!sheet) sheet = ss.insertSheet(SHEETS.DISQUALIFIER_GATES);
  return sheet;
}
function _ensureDisqualifierGatesHeaders() {
  const sheet = _getOrCreateDisqualifierGatesSheet();
  const headerRow = sheet.getRange(1, 1, 1, DISQUALIFIER_GATES_HEADERS.length).getValues()[0];
  const isEmpty = headerRow.every(v => v === "" || v === null);
  if (isEmpty) {
    sheet.getRange(1, 1, 1, DISQUALIFIER_GATES_HEADERS.length).setValues([DISQUALIFIER_GATES_HEADERS]);
  }
}
function _writeGatesPendingScratch(tickers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Raw_GatesScratch");
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet("Raw_GatesScratch");
  sheet.hideSheet();
  if (tickers.length > 0) {
    sheet.getRange(1, 1, tickers.length, 1).setValues(tickers.map(t => [t]));
  }
}
function _getGatesPendingScratchSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Raw_GatesScratch");
  if (!sheet) sheet = ss.insertSheet("Raw_GatesScratch");
  return sheet;
}