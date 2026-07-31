/*************************************************************
 * Fundamentals.gs — core financial metrics per ticker (v2)
 * -----------------------------------------------------------
 * WHAT CHANGED FROM v1: resumability. The entry point no longer
 * blindly processes "all tickers from position 0" every time —
 * it computes which tickers are MISSING from Raw_Fundamentals or
 * STALE (Last_Updated older than FUNDAMENTALS_STALENESS_DAYS),
 * and only processes those. This means:
 *   - An interrupted run resumes almost exactly where it left off
 *     when you call the exact same function again — no re-work.
 *   - A normal weekly refresh naturally reprocesses everything
 *     once last week's data ages past the staleness threshold —
 *     same mechanism serves both needs.
 *   - Pass { forceAll: true } to bypass freshness and reprocess
 *     every ticker regardless (e.g. after changing this script).
 *
 * HONEST FRESHNESS: Last_Updated is only stamped when the fetch
 * actually returned something real (non-blank Sector/Industry OR
 * a non-empty Finnhub metric object). A ticker that came back
 * fully blank (e.g. Finviz blocked us that run) stays "stale" and
 * gets retried automatically next time — it can never get stuck
 * looking falsely up-to-date with garbage data.
 *
 * The pending-ticker list lives in a hidden scratch sheet
 * (Raw_FundamentalsScratch), not Script Properties — same reason
 * as Universe.gs's scratch sheet: a list of thousands of tickers
 * would blow past Properties' ~9KB per-value limit.
 *
 * Two calls per ticker (unchanged from v1):
 *   1. Finviz quote page scrape -> Sector, Industry, Earnings_Date,
 *      Earnings_Session
 *   2. Finnhub /stock/metric?metric=all -> the financial ratios
 *
 * HONEST LABELING NOTE (unchanged): ROCE_proxy_ROIC is ROIC, not
 * true ROCE — Finnhub's free tier doesn't expose real ROCE.
 *
 * RUN ORDER MATTERS: run RefData.gs FIRST for Industry_PE/Sector_PE
 * to populate; this still runs fine without it, just blank there.
 *
 * ENTRY POINTS:
 *   runFundamentalsRefresh()                    — gap-fill/resume, full universe
 *   runFundamentalsRefresh({ forceAll: true })   — reprocess everyone regardless
 *   runFundamentalsRefresh_Test()                — synchronous, first 20 pending tickers
 *   runFundamentalsRefresh_Test(true)            — same, but via chunked triggers
 *************************************************************/

function runFundamentalsRefresh(options) {
  const mode = (options && options.mode) || MODE.PROD;
  const chunk = (options && typeof options.chunk !== "undefined") ? options.chunk : (mode === MODE.PROD);
  const forceAll = (options && options.forceAll) || false;

  _ensureFundamentalsHeaders();

  const uniSheet = _getOrCreateUniverseSheet();
  const uniLast = uniSheet.getLastRow();
  if (uniLast < 2) {
    SpreadsheetApp.getActive().toast("Fundamentals: Raw_Universe is empty — run Universe first.");
    return;
  }
  const allTickers = uniSheet.getRange(2, 1, uniLast - 1, 1).getValues().flat()
    .map(t => String(t).trim().toUpperCase()).filter(Boolean);

  const fundSheet = _getOrCreateFundamentalsSheet();
  const fLast = fundSheet.getLastRow();
  const freshSet = new Set();
  if (fLast >= 2 && !forceAll) {
    const existing = fundSheet.getRange(2, 1, fLast - 1, FUNDAMENTALS_HEADERS.length).getValues();
    const staleCutoff = Date.now() - (FUNDAMENTALS_STALENESS_DAYS * 24 * 60 * 60 * 1000);
    existing.forEach(r => {
      const ticker = String(r[0] || "").trim().toUpperCase();
      const lastUpdated = r[29]; // Last_Updated column
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
    SpreadsheetApp.getActive().toast(`Fundamentals: everything already fresh (updated within ${FUNDAMENTALS_STALENESS_DAYS} days). Nothing to do.`);
    return;
  }

  _writeFundamentalsPendingScratch(pending);

  if (chunk) {
    startChunkProcess("Fundamentals", "chunkWorker_Fundamentals", "finalize_Fundamentals", pending.length, DEFAULT_CHUNK_SIZE);
  } else {
    let idx = 0;
    while (idx < pending.length) idx = chunkWorker_Fundamentals(idx, DEFAULT_CHUNK_SIZE);
    finalize_Fundamentals();
  }
  SpreadsheetApp.getActive().toast(
    `Fundamentals refresh started (${mode}). ${pending.length} pending, ${allTickers.length - pending.length} already fresh (skipped).`
  );
}

function runFundamentalsRefresh_Test(chunk) {
  return runFundamentalsRefresh({ mode: MODE.TEST, chunk: (chunk === true) });
}


/*************************************************************
 * Chunk worker — reads its slice from the PENDING scratch list
 * (not directly from Raw_Universe by position — the pending set
 * can skip around, so a contiguous row-range read wouldn't work).
 *************************************************************/

function chunkWorker_Fundamentals(startIndex, chunkSize) {
  const pendingSheet = _getFundamentalsPendingScratchSheet();
  const total = Math.max(pendingSheet.getLastRow(), 0);
  if (startIndex >= total) return total;

  const endIndex = Math.min(startIndex + chunkSize, total);
  const numRows = endIndex - startIndex;
  const startRowNum = startIndex + 1; // no header row in this scratch sheet
  const tickers = pendingSheet.getRange(startRowNum, 1, numRows, 1).getValues().flat();

  const fundSheet = _getOrCreateFundamentalsSheet();
  const industryMap = _loadRefBenchmarkMap(SHEETS.REF_INDUSTRIES);
  const sectorMap = _loadRefBenchmarkMap(SHEETS.REF_SECTORS);

  const fLast = fundSheet.getLastRow();
  const existingRowByTicker = new Map();
  if (fLast >= 2) {
    const existingTickers = fundSheet.getRange(2, 1, fLast - 1, 1).getValues().flat();
    existingTickers.forEach((t, i) => existingRowByTicker.set(String(t).trim().toUpperCase(), i + 2));
  }

  const newRows = [];

  for (let i = 0; i < tickers.length; i++) {
    const ticker = String(tickers[i] || "").trim().toUpperCase();
    if (!ticker) continue;

    const details = _scrapeFinvizTickerDetails(ticker);
    Utilities.sleep(250);

    const fh = _fetchFinnhubMetrics(ticker);
    Utilities.sleep(150);

    const metric = fh.metric;
    const series = fh.series;

    const industryBench = industryMap.get(details.industry) || {};
    const sectorBench = sectorMap.get(details.sector) || {};

    // DIAGNOSTIC: your test showed Industry_Sales_5Y/Sector_Sales_5Y/etc coming
    // back blank while Sector/Industry themselves scraped fine — which points to
    // a naming-convention mismatch between what this scrape returns and what's
    // stored as "Name" in Raw_RefIndustries/Raw_RefSectors (different spacing,
    // capitalization, or dash character), not a scraping failure. This logs the
    // exact scraped value on a miss so we can see the mismatch directly.
    if (details.industry && Object.keys(industryBench).length === 0) {
      Logger.log(`Fundamentals: no Industry benchmark match for "${details.industry}" (ticker ${ticker})`);
    }
    if (details.sector && Object.keys(sectorBench).length === 0) {
      Logger.log(`Fundamentals: no Sector benchmark match for "${details.sector}" (ticker ${ticker})`);
    }

    // Only stamp Last_Updated if we actually got something real from BOTH
    // sources — a fully blank result (e.g. Finviz blocked us this run)
    // stays "stale" so it gets retried next time, instead of looking
    // falsely up-to-date.
    // FIXED: this used to be an OR across all three checks, which meant a
    // ticker where Finviz succeeded but Finnhub failed (or vice versa) got
    // marked fresh anyway — leaving whichever source failed permanently
    // blank (most scoring inputs come from Finnhub) until forced reprocess.
    // Now both sources have to contribute something for the row to count
    // as a real fetch.
    const finvizOk = (details.sector !== "" || details.industry !== "");
    const finnhubOk = Object.keys(metric).length > 0;
    const fetchSucceeded = finvizOk && finnhubOk;

    const row = [
      ticker,
      details.sector,
      details.industry,
      metric["52WeekLow"] || "",
      metric["52WeekHigh"] || "",
      metric.peAnnual || "",
      industryBench.pe || "",
      sectorBench.pe || "",
      metric.epsAnnual || "",
      metric.bookValuePerShareAnnual || "",
      _latestSeriesValue(series.roe),
      _seriesValueYearsAgo(series.roe, 3),
      metric.netProfitMarginAnnual != null ? metric.netProfitMarginAnnual / 100 : "",
      _latestSeriesValue(series.roic),
      metric.revenueGrowth3Y != null ? metric.revenueGrowth3Y / 100 : "",
      metric.revenueGrowth5Y != null ? metric.revenueGrowth5Y / 100 : "",
      industryBench.salesPast5Y || "",
      sectorBench.salesPast5Y || "",
      metric.epsGrowth3Y != null ? metric.epsGrowth3Y / 100 : "",
      metric.epsGrowth5Y != null ? metric.epsGrowth5Y / 100 : "",
      industryBench.epsPast5Y || "",
      sectorBench.epsPast5Y || "",
      metric["longTermDebt/equityAnnual"] || metric["totalDebt/totalEquityAnnual"] || "",
      _latestSeriesValue(series.longtermDebtTotalAsset),
      _seriesValueYearsAgo(series.longtermDebtTotalAsset, 3),
      metric.dividendPerShareTTM || "",
      metric.dividendYieldIndicatedAnnual != null ? metric.dividendYieldIndicatedAnnual / 100 : "",
      details.earningsDate,
      details.earningsSession,
      fetchSucceeded ? new Date() : ""
    ];

    const existingRow = existingRowByTicker.get(ticker);
    if (existingRow) {
      fundSheet.getRange(existingRow, 1, 1, FUNDAMENTALS_HEADERS.length).setValues([row]);
    } else {
      newRows.push(row);
    }
  }

  if (newRows.length > 0) {
    const startRow = fundSheet.getLastRow() + 1;
    fundSheet.getRange(startRow, 1, newRows.length, FUNDAMENTALS_HEADERS.length).setValues(newRows);
  }

  return endIndex;
}

function finalize_Fundamentals() {
  const sheet = _getOrCreateFundamentalsSheet();
  formatSheet(sheet, FUNDAMENTALS_HEADERS.length);
  SpreadsheetApp.getActive().toast("Fundamentals complete: all pending tickers processed.");
}

/*************************************************************
 * Finviz per-ticker scrape — Sector, Industry, Earnings date/session.
 *************************************************************/

function _scrapeFinvizTickerDetails(ticker) {
  const finvizSymbol = ticker.replace(/\./g, "-");
  //const url = `https://finviz.com/quote?t=${encodeURIComponent(finvizSymbol)}`;
  const url = `https://finviz.com/stock?t=${encodeURIComponent(finvizSymbol)}`; //finviz updated url endpoint
  let html = "";
  try {
    const res = safeFetch(url, { headers: { "User-Agent": "Mozilla/5.0", "Accept": "text/html" } }, 3);
    html = res.getContentText();
  } catch (e) {
    Logger.log(`_scrapeFinvizTickerDetails: fetch failed for ${ticker}: ${e}`);
    return { sector: "", industry: "", earningsDate: "", earningsSession: "" };
  }

  // Generalized defensively (matches with or without .ashx) — your Fundamentals
  // test run suggests these were actually still working, but this costs nothing
  // and guards against Finviz changing this link format too.
  
  //const sectorMatch = html.match(/<a href="[^"]*\?v=111&f=sec_[^"]+"[^>]*>([^<]+)<\/a>/i);
  //const sector = sectorMatch ? htmlDecode(sectorMatch[1]) : "";
  //changed regex as above one was not working
  var sectorMatch = html.match(
    /href="screener\?v=111(?:&|&amp;)f=sec_[^"]*"[^>]*>([^<]+)</i
  );
  var sector = sectorMatch ? htmlDecode(sectorMatch[1]) : "";

  //const industryMatch = html.match(/<a href="[^"]*\?v=111&f=ind_[^"]+"[^>]*>([^<]+)<\/a>/i);
  //const industry = industryMatch ? htmlDecode(industryMatch[1]) : "";
  //changed regex as above one was not working
  var industryMatch = html.match(
    /href="screener\?v=111(?:&|&amp;)f=ind_[^"]*"[\s\S]*?<span[^>]*>([^<]+)<\/span>/i
  );
  var industry = industryMatch ? htmlDecode(industryMatch[1]) : "";

  // FIXED — this exactly matched the pattern you found broken (quote.ashx?t=
  // hardcoded), generalized the same way your working screener fix did.
  /*const earningsRegex = new RegExp(
    '<a[^>]*href="[^"]*\\?t=' + finvizSymbol + '[^"]*ty=ea"[^>]*>\\s*<b><small[^>]*>([^<]+)<\\/small>',
    'i'
  );
  const earningsMatch = html.match(earningsRegex);*/
  //changed regex as above one was not working
  const earningsMatch = html.match(/<a[^>]*>Earnings<\/a>[\s\S]*?<small[^>]*>([^<]+)<\/small>/i);

  let earningsDate = "";
  let earningsSession = "";
  if (earningsMatch && earningsMatch[1]) {
    const parts = earningsMatch[1].trim().split(/\s+/);
    if (parts.length === 3) {
      earningsDate = `${parts[0]} ${parts[1]}`;
      earningsSession = parts[2];
    } else if (parts.length === 2) {
      earningsDate = `${parts[0]} ${parts[1]}`;
    }
  }

  return { sector, industry, earningsDate, earningsSession };
}

/*************************************************************
 * Finnhub metrics call
 *************************************************************/

function _fetchFinnhubMetrics(ticker) {
  const token = _getFinnhubToken();
  const url = `https://finnhub.io/api/v1/stock/metric?symbol=${encodeURIComponent(ticker)}&metric=all&token=${token}`;
  try {
    const res = safeFetch(url, {}, 3);
    const json = JSON.parse(res.getContentText());
    return {
      metric: json.metric || {},
      series: (json.series && json.series.annual) ? json.series.annual : {}
    };
  } catch (e) {
    Logger.log(`_fetchFinnhubMetrics: error for ${ticker}: ${e}`);
    return { metric: {}, series: {} };
  }
}

function _latestSeriesValue(arr) {
  return (arr && arr.length > 0) ? arr[0].v : "";
}
function _seriesValueYearsAgo(arr, years) {
  return (arr && arr.length > years) ? arr[years].v : "";
}

/*************************************************************
 * Benchmark lookup maps (Finviz-sourced, blended reference)
 *************************************************************/

function _loadRefBenchmarkMap(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const map = new Map();
  if (!sheet) return map;
  const last = sheet.getLastRow();
  if (last < 2) return map;
  const data = sheet.getRange(2, 1, last - 1, REF_HEADERS.length).getValues();
  data.forEach(r => {
    const name = String(r[1] || "").trim();
    if (!name) return;
    map.set(name, {
      pe: r[3],
      epsPast5Y: r[10],
      salesPast5Y: r[12]
    });
  });
  return map;
}

/*************************************************************
 * Sheet helpers
 *************************************************************/

function _getOrCreateFundamentalsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.FUNDAMENTALS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.FUNDAMENTALS);
  return sheet;
}

function _ensureFundamentalsHeaders() {
  const sheet = _getOrCreateFundamentalsSheet();
  const headerRow = sheet.getRange(1, 1, 1, FUNDAMENTALS_HEADERS.length).getValues()[0];
  const isEmpty = headerRow.every(v => v === "" || v === null);
  if (isEmpty) {
    sheet.getRange(1, 1, 1, FUNDAMENTALS_HEADERS.length).setValues([FUNDAMENTALS_HEADERS]);
  }
}

function _writeFundamentalsPendingScratch(tickers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Raw_FundamentalsScratch");
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet("Raw_FundamentalsScratch");
  sheet.hideSheet();
  if (tickers.length > 0) {
    sheet.getRange(1, 1, tickers.length, 1).setValues(tickers.map(t => [t]));
  }
}

function _getFundamentalsPendingScratchSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Raw_FundamentalsScratch");
  if (!sheet) sheet = ss.insertSheet("Raw_FundamentalsScratch");
  return sheet;
}