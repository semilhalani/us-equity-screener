/*************************************************************
 * Universe.gs — builds and tiers the master ticker list
 * -----------------------------------------------------------
 * STAGE A — index & cap-bucket membership (Finviz scrape)
 *   Runs one filter (of 9) per chunk tick, writes raw ticker
 *   hits to a hidden scratch sheet. Chunked because some
 *   cap buckets (micro/nano) can span 100+ screener pages —
 *   safer to chunk than to risk the 6-minute execution cap.
 *
 * STAGE B — per-ticker enrichment (Finnhub profile2)
 *   Fills MarketCap (precise $) + IPO date + Listing_Age_Days
 *   for every ticker written in Stage A. Chunked, 100/tick.
 *
 * ENTRY POINTS:
 *   runUniverseRefresh()          — full production run
 *   runUniverseRefresh_Test()     — fast, SP500-only, 20 tickers
 *   runUniverseRefresh_Test(false) — same, but fully synchronous
 *                                     (no triggers, done in one run)
 *************************************************************/
function runUniverseRefresh(options) {
  const mode = (options && options.mode) || MODE.PROD;
  const limit = (options && options.limit) || null;
  const chunk = (options && typeof options.chunk !== "undefined") ? options.chunk : (mode === MODE.PROD);
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROP.PIPELINE_MODE, mode);
  props.setProperty(PROP.UNIVERSE_CHUNK_PREF, chunk ? "true" : "false");
  if (limit) props.setProperty(PROP.TEST_LIMIT, String(limit));
  else props.deleteProperty(PROP.TEST_LIMIT);
  _resetScratchSheet();
  // TEST mode only scrapes SP500 — fast enough to run synchronously end-to-end.
  const activeFilters = (mode === MODE.TEST)
    ? UNIVERSE_FILTERS.filter(f => f.label === "SP500")
    : UNIVERSE_FILTERS;
  props.setProperty(PROP.UNIVERSE_ACTIVE_FILTERS, JSON.stringify(activeFilters));
  if (chunk) {
    startChunkProcess("UniverseIndexCap", "chunkWorker_UniverseIndexCap", "finalize_UniverseIndexCap", activeFilters.length, 1);
  } else {
    let idx = 0;
    while (idx < activeFilters.length) idx = chunkWorker_UniverseIndexCap(idx, 1);
    finalize_UniverseIndexCap();
  }
  SpreadsheetApp.getActive().toast(`Universe refresh started (${mode}).`);
}
function runUniverseRefresh_Test(chunk) {
  return runUniverseRefresh({ mode: MODE.TEST, chunk: (chunk === true), limit: DEFAULT_TEST_LIMIT });
}
/*************************************************************
 * STAGE A — chunked Finviz index/cap scrape
 * One filter (of 9) processed per chunk tick.
 *************************************************************/
function chunkWorker_UniverseIndexCap(startIndex, chunkSize) {
  const props = PropertiesService.getScriptProperties();
  const activeFilters = JSON.parse(props.getProperty(PROP.UNIVERSE_ACTIVE_FILTERS) || "[]");
  const item = activeFilters[startIndex];
  if (!item) return startIndex + 1;
  const tickers = _scrapeFinvizScreenerFilter(item.filter);
  const scratch = _getOrCreateScratchSheet();
  const rows = tickers.map(t => [t, item.type, item.label]);
  if (rows.length > 0) {
    const startRow = scratch.getLastRow() + 1;
    scratch.getRange(startRow, 1, rows.length, 3).setValues(rows);
  }
  Logger.log(`UniverseIndexCap: ${item.label} -> ${tickers.length} tickers`);
  return startIndex + 1;
}
function finalize_UniverseIndexCap() {
  _buildUniverseFromScratch();
  PropertiesService.getScriptProperties().deleteProperty(PROP.UNIVERSE_ACTIVE_FILTERS);
  _startUniverseEnrichStageB();
}
function _scrapeFinvizScreenerFilter(filter) {
  const urlBase = "https://finviz.com/screener?v=111&f=";
  let all = [];
  let page = 1;
  let more = true;
  while (more) {
    const start = (page - 1) * 20 + 1;
    const url = `${urlBase}${filter}&r=${start}`;
    let html = "";
    try {
      const res = safeFetch(url, { headers: { "User-Agent": "Mozilla/5.0" } }, 4);
      html = res.getContentText();
    } catch (e) {
      Logger.log(`_scrapeFinvizScreenerFilter: error on ${filter} page ${page}: ${e}`);
      break;
    }
    const regex = /<a\s+href="[^"]*\?t=([^"&]+)[^"]*"[^>]*>/gi; //replaced regex temporarily
    // const regex = /<a\s+href="quote\.ashx\?t=([^"&]+)[^"]*"[^>]*>/gi;
    const found = [...html.matchAll(regex)].map(m => m[1].trim().toUpperCase());
    const unique = [...new Set(found)];
    if (unique.length === 0) {
      // DIAGNOSTIC: log what actually came back so we can tell a genuine
      // block apart from Finviz changing their HTML structure, instead of
      // guessing at a regex fix with no evidence.
      Logger.log(`_scrapeFinvizScreenerFilter: 0 matches for ${filter} page ${page}. HTTP length=${html.length}. First 500 chars: ${html.substring(0, 500)}`);
      more = false;
    } else {
      unique.forEach(t => { if (!all.includes(t)) all.push(t); });
      page++;
      if (unique.length < 20) more = false; // last page reached
    }
    Utilities.sleep(250); // polite delay
  }
  Logger.log(`_scrapeFinvizScreenerFilter(${filter}): ${all.length} tickers total`);
  return all;
}
/*************************************************************
 * Merge scratch results + Finnhub symbol list -> Raw_Universe
 *************************************************************/
function _buildUniverseFromScratch() {
  const scratch = _getOrCreateScratchSheet();
  const last = scratch.getLastRow();
  const indexMap = new Map();
  const capMap = new Map();
  if (last >= 1) {
    const data = scratch.getRange(1, 1, last, 3).getValues();
    data.forEach(r => {
      const ticker = String(r[0] || "").trim().toUpperCase();
      const type = r[1];
      const label = r[2];
      if (!ticker) return;
      if (type === "index") {
        const arr = indexMap.get(ticker) || [];
        if (!arr.includes(label)) arr.push(label);
        indexMap.set(ticker, arr);
      } else if (type === "cap") {
        capMap.set(ticker, label);
      }
    });
  }
  Logger.log(`_buildUniverseFromScratch: indexMap=${indexMap.size}, capMap=${capMap.size}`);
  const allSymbols = _fetchAllUSSymbols();
  let commonStocks = allSymbols.filter(s => ["Common Stock", "ADR", "EQS"].includes(s.type));
  // OTC / EXCHANGE FILTER: whitelist approach — see Constants.gs for why.
  // Best-effort until verified against a real run: logs an excluded-ticker
  // count per MIC code so you can confirm nothing legitimate got dropped.
  if (!INCLUDE_OTC_TICKERS) {
    const beforeCount = commonStocks.length;
    const excludedMicCounts = new Map();
    commonStocks = commonStocks.filter(s => {
      const mic = String(s.mic || "").trim().toUpperCase();
      const keep = WHITELISTED_EXCHANGE_MICS.includes(mic);
      if (!keep) excludedMicCounts.set(mic, (excludedMicCounts.get(mic) || 0) + 1);
      return keep;
    });
    Logger.log(`Universe: OTC/exchange filter removed ${beforeCount - commonStocks.length} of ${beforeCount} tickers. ` +
      `Excluded MIC breakdown: ${JSON.stringify(Object.fromEntries(excludedMicCounts))}. ` +
      `If a real exchange shows up here, add its MIC code to WHITELISTED_EXCHANGE_MICS in Constants.gs.`);
  }
  const props = PropertiesService.getScriptProperties();
  const mode = props.getProperty(PROP.PIPELINE_MODE) || MODE.PROD;
  const limitRaw = props.getProperty(PROP.TEST_LIMIT);
  if (mode === MODE.TEST && limitRaw) {
    // Sample tickers we KNOW are in the scraped index/cap groups FIRST, so a
    // test run actually demonstrates tiering working (visible SP500 membership,
    // Tier1_Core assignment) instead of a near-random slice of Finnhub's raw
    // list order, which skews toward obscure OTC names.
    const known = commonStocks.filter(s => {
      const t = String(s.symbol || "").trim().toUpperCase();
      return indexMap.has(t) || capMap.has(t);
    });
    const rest = commonStocks.filter(s => {
      const t = String(s.symbol || "").trim().toUpperCase();
      return !(indexMap.has(t) || capMap.has(t));
    });
    commonStocks = known.concat(rest).slice(0, Number(limitRaw));
  }
  const rows = commonStocks.map(s => {
    const ticker = String(s.symbol || "").trim().toUpperCase();
    const indices = indexMap.get(ticker) || [];
    const capBucket = capMap.get(ticker) || "";
    const tier = _assignTier(indices, capBucket);
    const type = s.type || "";
    const isADR = (type === "ADR");
    return [
      ticker, s.description || "", (s.mic || ""), type, isADR,
      "", capBucket, tier, indices.join(","), "",
      "", "", "", new Date(), ""
    ];
  });
  const sheet = _getOrCreateUniverseSheet();
  sheet.clearContents();
  sheet.getRange(1, 1, 1, UNIVERSE_HEADERS.length).setValues([UNIVERSE_HEADERS]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, UNIVERSE_HEADERS.length).setValues(rows);
  }
  formatSheet(sheet, UNIVERSE_HEADERS.length);
  SpreadsheetApp.getActive().toast(`Universe Stage A done: ${rows.length} tickers tiered. Starting Stage B…`);
}
function _fetchAllUSSymbols() {
  const token = _getFinnhubToken();
  const url = `https://finnhub.io/api/v1/stock/symbol?exchange=US&token=${token}`;
  const res = safeFetch(url, {}, 4);
  const json = JSON.parse(res.getContentText());
  return Array.isArray(json) ? json : [];
}
function _assignTier(indices, capBucket) {
  if (indices && indices.length > 0) return TIER.CORE;
  if (capBucket === "Mega" || capBucket === "Large" || capBucket === "Mid") return TIER.LARGE_MID;
  if (capBucket === "Small") return TIER.SMALL;
  if (capBucket === "Micro" || capBucket === "Nano") return TIER.MOONSHOT;
  return TIER.SMALL; // unknown bucket defaults to Small rather than silently dropping the ticker
}
/*************************************************************
 * STAGE B — chunked Finnhub profile2 enrichment
 * 100 tickers per tick — reads its slice straight from the
 * sheet (no stored payload), so there's no size ceiling here.
 *************************************************************/
function _startUniverseEnrichStageB() {
  const props = PropertiesService.getScriptProperties();
  const chunk = props.getProperty(PROP.UNIVERSE_CHUNK_PREF) !== "false";
  const sheet = _getOrCreateUniverseSheet();
  const total = Math.max(sheet.getLastRow() - 1, 0);
  if (total === 0) {
    Logger.log("_startUniverseEnrichStageB: nothing to enrich.");
    return;
  }
  if (chunk) {
    startChunkProcess("UniverseEnrich", "chunkWorker_UniverseEnrich", "finalize_UniverseEnrich", total, 100);
  } else {
    let idx = 0;
    while (idx < total) idx = chunkWorker_UniverseEnrich(idx, 20);
    finalize_UniverseEnrich();
  }
}
function chunkWorker_UniverseEnrich(startIndex, chunkSize) {
  const sheet = _getOrCreateUniverseSheet();
  const total = Math.max(sheet.getLastRow() - 1, 0);
  if (startIndex >= total) return total;
  const endIndex = Math.min(startIndex + chunkSize, total);
  const numRows = endIndex - startIndex;
  const startRowNum = startIndex + 2; // +2: header row + 0-based -> 1-based
  const tickers = sheet.getRange(startRowNum, 1, numRows, 1).getValues().flat();
  for (let i = 0; i < tickers.length; i++) {
    const ticker = String(tickers[i] || "").trim().toUpperCase();
    const rowNum = startRowNum + i;
    if (!ticker) continue;
    try {
      const token = _getFinnhubToken();
      const url = `https://finnhub.io/api/v1/stock/profile2?symbol=${encodeURIComponent(ticker)}&token=${token}`;
      const res = safeFetch(url, {}, 3);
      const profile = JSON.parse(res.getContentText());
      const marketCap = (profile.marketCapitalization != null) ? profile.marketCapitalization * 1e6 : "";
      const ipoDate = profile.ipo || "";
      let listingAgeDays = "";
      if (ipoDate) {
        const ipo = new Date(ipoDate);
        if (!isNaN(ipo.getTime())) {
          listingAgeDays = Math.floor((Date.now() - ipo.getTime()) / (1000 * 60 * 60 * 24));
        }
      }
      const flagNewListing = (listingAgeDays !== "" && listingAgeDays < DISQUALIFIERS.MIN_LISTING_AGE_DAYS);
      sheet.getRange(rowNum, 6).setValue(marketCap);        // MarketCap
      sheet.getRange(rowNum, 11).setValue(ipoDate);          // IPO_Date
      sheet.getRange(rowNum, 12).setValue(listingAgeDays);   // Listing_Age_Days
      sheet.getRange(rowNum, 13).setValue(flagNewListing);   // Flag_NewListing
      sheet.getRange(rowNum, 14).setValue(new Date());       // Last_Updated
      const sharesOutstanding = profile.shareOutstanding != null ? profile.shareOutstanding * 1e6 : "";
      // Current snapshot only — no more history log. Dilution is computed
      // from SEC filings via financials-reported in DisqualifierGates.gs
      // now, so there's no need to accumulate our own time series here.
      sheet.getRange(rowNum, 15).setValue(sharesOutstanding); // Shares_Outstanding
    } catch (e) {
      Logger.log(`chunkWorker_UniverseEnrich: error for ${ticker}: ${e}`);
    }
    Utilities.sleep(150);
  }
  return endIndex;
}
function finalize_UniverseEnrich() {
  const sheet = _getOrCreateUniverseSheet();
  formatSheet(sheet, UNIVERSE_HEADERS.length);
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP.UNIVERSE_CHUNK_PREF);
  props.deleteProperty(PROP.TEST_LIMIT);
  props.deleteProperty(PROP.PIPELINE_MODE);
  SpreadsheetApp.getActive().toast("Universe complete: market cap + IPO dates filled for all tickers.");
}
/*************************************************************
 * Sheet helpers
 *************************************************************/
function _getOrCreateUniverseSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.UNIVERSE);
  if (!sheet) sheet = ss.insertSheet(SHEETS.UNIVERSE);
  return sheet;
}
function _resetScratchSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.UNIVERSE_SCRATCH);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(SHEETS.UNIVERSE_SCRATCH);
  sheet.hideSheet();
}
function _getOrCreateScratchSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.UNIVERSE_SCRATCH);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.UNIVERSE_SCRATCH);
    sheet.hideSheet();
  }
  return sheet;
}