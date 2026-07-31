/*************************************************************
 * Constants.gs — US Stock System v1 (consolidated, final)
 * -----------------------------------------------------------
 * Central configuration shared by every script in the project.
 * Nothing in here fetches data — it's pure config + small,
 * generic utilities used everywhere (safeFetch, formatSheet).
 *
 * This replaces every previous "addition block" version — all
 * the pieces that were scattered across 5 separate paste-in
 * blocks now live here once, in one place, so there's nothing
 * left to accidentally overwrite or miss during copy-paste.
 *************************************************************/
// ---- SCRIPT PROPERTY KEYS ----
// NOTE: We deliberately never store large arrays (e.g. ticker
// lists) under any of these keys — Script Properties have a
// ~9KB per-value limit. Large lists live in sheet ranges instead;
// workers read their slice of work directly from a sheet using
// (startIndex, chunkSize). See ChunkEngine.gs for details.
const PROP = {
  PIPELINE_MODE: "PIPELINE_MODE",
  MASTER_STAGE: "MASTER_STAGE",
  CHUNK_MODULE: "CHUNK_MODULE",
  CHUNK_WORKER: "CHUNK_WORKER",
  CHUNK_FINALIZE: "CHUNK_FINALIZE",
  CHUNK_INDEX: "CHUNK_INDEX",
  CHUNK_TOTAL: "CHUNK_TOTAL",
  CHUNK_SIZE: "CHUNK_SIZE",
  TEST_LIMIT: "TEST_LIMIT",
  FINNHUB_KEY_INDEX: "FINNHUB_KEY_INDEX",
  UNIVERSE_ACTIVE_FILTERS: "UNIVERSE_ACTIVE_FILTERS", // tiny (max 9 items) — safe exception
  UNIVERSE_CHUNK_PREF: "UNIVERSE_CHUNK_PREF"
};
const MODE = { PROD: "PRODUCTION", TEST: "TEST" };
const TIMEZONE = "America/New_York"; // US market timezone
const DEFAULT_CHUNK_SIZE = 40;
const DEFAULT_TEST_LIMIT = 20;
// ---- SHEET NAMES ----
const SHEETS = {
  BUY_LIST: "📋 BUY LIST",
  MOONSHOT: "🎯 MOONSHOT WATCHLIST",
  EARNINGS_ALERTS: "⚠️ EARNINGS ALERTS",
  POST_EARNINGS: "🔄 POST-EARNINGS MOVES",
  INSIDER_FEED: "🔍 INSIDER FEED",
  TRACK_RECORD: "📊 TRACK RECORD",
  UNIVERSE: "Raw_Universe",
  UNIVERSE_SCRATCH: "Raw_UniverseScratch", // auto-created/managed by Universe.gs, not manual
  REF_INDUSTRIES: "Raw_RefIndustries",
  REF_SECTORS: "Raw_RefSectors",
  TIER_BENCHMARKS: "Raw_TierBenchmarks",
  FUNDAMENTALS: "Raw_Fundamentals",
  EARNINGS: "Raw_Earnings",
  EPS_CATEGORY: "Raw_EPSCategory",
  INSIDERS: "Raw_Insiders",
  INSIDERS_CHECKPOINT: "Raw_InsidersCheckpoint", // FIXED — was only a to-do comment below, never actually added; Insiders.gs was calling getSheetByName(undefined)
  INSIDER_SIGNALS: "Raw_InsiderSignals",         // FIXED — same bug as above
  CALENDAR: "Raw_Calendar",
  PRICES: "Raw_Prices",
  SCORE_HISTORY: "Raw_ScoreHistory",
  SCORES: "Scores",
  DISQUALIFIER_GATES: "Raw_DisqualifierGates" // was missing — DisqualifierGates.gs needs this
};
// ---- DILUTION GATE CONFIG ----
// Raw_SharesHistory and its free/no-API-cost snapshot method have been
// retired entirely — it would only have started paying off ~11 months in,
// and in the meantime grew by ~5,336 rows on every single Universe run,
// forever. The financials-reported bootstrap in DisqualifierGates.gs is now
// the ONLY dilution method, for every tier, from day one.
// Two report dates (from Finnhub's financials-reported / SEC filings) need
// to be at least this many days apart to count as a real YoY comparison —
// a little under 365 so a company that files a few days early or late each
// cycle still counts, without accepting anything close to a quarter apart.
const DILUTION_MIN_SNAPSHOT_GAP_DAYS = 330;
// ---- OTC / EXCHANGE INCLUSION FLAG ----
// Most OTC / pink-sheet tickers don't have usable Finviz pages (frequent
// 404s burning retry budget for nothing) and can't currently be traded via
// Trading212 anyway. Set to true later if that changes. Filtering here
// shrinks Raw_Universe itself, which cuts API calls proportionally across
// EVERY downstream script (Fundamentals, Earnings, DisqualifierGates,
// Insiders), not just Universe.gs.
const INCLUDE_OTC_TICKERS = false;
// Whitelist approach, not a blacklist of OTC codes — safer, since Finnhub's
// exact OTC MIC codes aren't verifiable without a live test call. Anything
// NOT in this list gets excluded when INCLUDE_OTC_TICKERS is false. Universe.gs
// logs excluded-ticker MIC counts on every run — check that log after your
// first real run and add any major exchange code that got wrongly excluded.
const WHITELISTED_EXCHANGE_MICS = [
  "XNYS", // NYSE
  "XNAS", // Nasdaq
  "ARCX", // NYSE Arca
  "XASE", // NYSE American
  "BATS", "BATY", // Cboe BZX/BYX
  "IEXG"  // IEX
];
// ---- TIERS ----
const TIER = {
  CORE: "Tier1_Core",        // SP500 / NASDAQ100 / DJ30 constituents
  LARGE_MID: "Tier2_LargeMid",
  SMALL: "Tier3_Small",
  MOONSHOT: "Tier4_Moonshot" // Micro + Nano cap, scored completely differently
};
// Kept for reference / future fallback use.
const CAP_THRESHOLDS = {
  MEGA: 200e9,
  LARGE: 10e9,
  MID: 2e9,
  SMALL: 300e6,
  MICRO: 50e6,
  NANO: 0
};
// ---- DISQUALIFIER THRESHOLDS ----
const DISQUALIFIERS = {
  MIN_LISTING_AGE_DAYS: 180,
  MIN_AVG_DOLLAR_VOLUME: 100000,
  MIN_AVG_DOLLAR_VOLUME_MOONSHOT: 50000,
  MAX_ZERO_REVENUE_QUARTERS: 8,
  MAX_SHARES_OUTSTANDING_YOY_GROWTH: 0.30,
  MIN_CASH_RUNWAY_MONTHS_MOONSHOT: 6
};
// Groups with fewer tickers than this get Low_Sample_Flag = true — was missing.
const MIN_BENCHMARK_SAMPLE_SIZE = 5;
// ---- FINNHUB API KEYS ---- (your 8 real keys, unchanged)
const FINNHUB_KEYS = [
  "YOUR_FINNHUB_KEY_1",
  "YOUR_FINNHUB_KEY_2",
  "YOUR_FINNHUB_KEY_3",
  "YOUR_FINNHUB_KEY_4",
  "YOUR_FINNHUB_KEY_5",
  "YOUR_FINNHUB_KEY_6",
  "YOUR_FINNHUB_KEY_7",
  "YOUR_FINNHUB_KEY_8",
  "YOUR_FINNHUB_KEY_9",
  "YOUR_FINNHUB_KEY_10"
  // ...add as many as you use
];
function _getFinnhubToken() {
  const props = PropertiesService.getScriptProperties();
  let idx = parseInt(props.getProperty(PROP.FINNHUB_KEY_INDEX) || "0", 10);
  const token = FINNHUB_KEYS[idx % FINNHUB_KEYS.length];
  props.setProperty(PROP.FINNHUB_KEY_INDEX, String(idx + 1));
  return token;
}
// ---- UNIVERSE-SPECIFIC CONFIG ----
const UNIVERSE_HEADERS = [
  "Ticker", "Name", "Exchange", "Type", "Is_ADR",
  "MarketCap", "MarketCap_Bucket", "Tier",
  "Index_Membership", "Index_Weightage",
  "IPO_Date", "Listing_Age_Days", "Flag_NewListing", "Last_Updated", "Shares_Outstanding"
];
// The 9 Finviz screener filters that establish index membership
// and market-cap bucket.
const UNIVERSE_FILTERS = [
  { type: "index", label: "SP500", filter: "idx_sp500" },
  { type: "index", label: "NASDAQ100", filter: "idx_ndx" },
  { type: "index", label: "DJ30", filter: "idx_dji" },
  { type: "cap", label: "Mega", filter: "cap_mega" },
  { type: "cap", label: "Large", filter: "cap_large" },
  { type: "cap", label: "Mid", filter: "cap_mid" },
  { type: "cap", label: "Small", filter: "cap_small" },
  { type: "cap", label: "Micro", filter: "cap_micro" },
  { type: "cap", label: "Nano", filter: "cap_nano" }
];
// ---- REFDATA-SPECIFIC CONFIG ----
// Column order confirmed against your own "Sam" file's VLOOKUP formulas.
const REF_HEADERS = [
  "No", "Name", "Market_Cap", "PE", "Fwd_PE", "PEG", "PS", "PB", "PC",
  "PFCF", "EPS_Past5Y", "EPS_Next5Y", "Sales_Past5Y", "Change", "Volume"
];
// ---- FUNDAMENTALS-SPECIFIC CONFIG ----
// THIS WAS THE MISSING ONE — accidentally overwritten by a duplicate
// UNIVERSE_HEADERS during copy-paste. Percentages/ratios are stored as
// true decimals (0.025 = 2.5%), matching the surprisePct convention
// used elsewhere, so every ScoringEngine calculation stays consistent.
const FUNDAMENTALS_HEADERS = [
  "Ticker", "Sector", "Industry", "Week52_Low", "Week52_High",
  "PE", "Industry_PE", "Sector_PE",
  "EPS", "Book_Value", "ROE", "ROE_3yr", "Net_Profit_Margin",
  "ROCE_proxy_ROIC",
  "Revenue_Growth_3Y", "Revenue_Growth_5Y", "Industry_Sales_5Y", "Sector_Sales_5Y",
  "EPS_Growth_3Y", "EPS_Growth_5Y", "Industry_EPS_5Y", "Sector_EPS_5Y",
  "Debt_Equity", "Debt_Asset_Now", "Debt_Asset_3Y_Ago",
  "Dividend_Per_Share", "Dividend_Yield_Pct",
  "Earnings_Date", "Earnings_Session", "Last_Updated"
];
// ---- TIER BENCHMARKS CONFIG ---- (was entirely missing)
const TIER_BENCHMARKS_HEADERS = [
  "Group_Type", "Group_Name", "Tier", "Sample_Count",
  "Median_PE", "Median_ROE", "Median_Net_Profit_Margin",
  "Median_Revenue_Growth_3Y", "Median_EPS_Growth_3Y", "Median_Debt_Equity",
  "Low_Sample_Flag", "Last_Updated"
];
// ---- DISQUALIFIER GATES CONFIG ----
// v4: packed sentence-style status columns (e.g. "Diluting (136.6% YoY,
// bootstrapped)") replaced with clean columns per gate — a real number, a
// pass/fail boolean, and a method flag explaining what did (or didn't)
// produce that number. The old _Notes columns (plain-English sentences)
// are gone entirely from this raw sheet; that kind of presentation text
// belongs in a future presentation sheet built FROM these clean numbers,
// not duplicated as stored text here.
//
// Method flag values actually possible today:
//   Dilution_Method:    "Bootstrapped" | "Insufficient_History"
//   Zero_Revenue_Method: "Computed" | "Not_Applicable" (non-Moonshot tier)
//                        | "Insufficient_Data"
//   Cash_Runway_Method: "Cash_Flow_Positive" | "Burn_Rate_Computed"
//                        | "Not_Applicable" (non-Moonshot tier)
//                        | "Insufficient_Data"
const DISQUALIFIER_GATES_HEADERS = [
  "Ticker", "Tier", "Avg_Dollar_Volume", "Liquidity_Pass", "Listing_Age_Pass",
  "Dilution_YoY_Pct", "Dilution_Pass", "Dilution_Method",
  "Zero_Revenue_Count", "Zero_Revenue_Quarters_Considered", "Zero_Revenue_Pass", "Zero_Revenue_Method",
  "Cash_Runway_Months", "Cash_Runway_Pass", "Cash_Runway_Method",
  "Overall_Disqualified", "Disqualify_Reasons", "Last_Updated"
];
// ---- RESUMABILITY STALENESS THRESHOLDS ----
const FUNDAMENTALS_STALENESS_DAYS = 6;
const DISQUALIFIER_GATES_STALENESS_DAYS = 6;
const EARNINGS_STALENESS_DAYS = 6;
// ---- EARNINGS CONFIG ----
const EARNINGS_HEADERS = [
  "Ticker",
  "Period_1", "Quarter_1", "Estimate_1", "Actual_1", "Surprise_1", "SurprisePct_1",
  "Period_2", "Quarter_2", "Estimate_2", "Actual_2", "Surprise_2", "SurprisePct_2",
  "Period_3", "Quarter_3", "Estimate_3", "Actual_3", "Surprise_3", "SurprisePct_3",
  "Period_4", "Quarter_4", "Estimate_4", "Actual_4", "Surprise_4", "SurprisePct_4",
  "Last_Updated"
];
const EPS_CATEGORY_HEADERS = [
  "Ticker", "Q1_SurprisePct", "Q2_SurprisePct", "Q3_SurprisePct", "Q4_SurprisePct",
  "Category", "Volatility_Score", "Reliability_Score", "Last_Updated"
];
const EPS_CATEGORY_ORDER = [
  "🚀 Ultra Beat (300%+)",
  "💎 Mega Beat (100%+)",
  "💚 Strong Beat (All ≥10%)",
  "👍 Consistent Beat (3×≥10%)",
  "🙂 Mild Beat (2×≥10%)",
  "🟩 Mixed Positive",
  "⚡ Mixed (Strong Beat Present)",
  "⚪ Mixed Neutral",
  "🔥 Ultra Miss (300%-)",
  "😡 Mega Miss (100%-)",
  "❌ Strong Miss (All ≤ -10%)",
  "👎 Consistent Miss (3×≤ -10%)",
  "🙁 Mild Miss (2×≤ -10%)",
  "🟥 Mixed Negative",
  "⚪ Undefined"
];
// ---- SCORING CONFIG ----
// v3: same clean-data/presentation-layer split already applied to
// DisqualifierGates.gs. Emoji+sentence fields replaced with short codes;
// Overall_Disqualified inverted to Qualified (TRUE = good, natural filter
// direction); Action's packed sentence split into Verdict (single enum)
// + Caution_Flags (comma-separated codes). Nothing here is ever blank when
// a real answer exists — "NONE" / "NOT_YET_EVALUATED" / "UNKNOWN" sentinels
// replace ambiguous empty cells, same honest-degradation principle used
// everywhere else in this project. Presentation sheets (Buy List, Moonshot
// Watchlist, etc.) are responsible for turning these codes into whatever
// human sentence fits their own context — never duplicated as stored text
// here.
//
// Possible values:
//   Earnings_Archetype_Code: "STABLE_COMPOUNDER" | "VOLATILITY_HARVEST"
//                             | "HIGH_RISK" | "LOW_CONVICTION" | "" (n/a)
//   Qualified: TRUE | FALSE | "" (gates never ran for this ticker yet)
//   Disqualify_Reason_Codes: "NONE" | comma-separated codes from
//                             DisqualifierGates | "NOT_YET_EVALUATED"
//   Verdict: "STRONG_CANDIDATE" | "WATCH" | "WEAK" | "AVOID_WEAK_FUNDAMENTALS"
//            | "AVOID_DISQUALIFIED" | "MOONSHOT_WATCH" | "SPECULATIVE_WATCH"
//            | "AVOID_WEAK_MOONSHOT" | "INSUFFICIENT_DATA"
//   Caution_Flags: "NONE" | comma-separated codes, e.g. "INCONSISTENT_EARNINGS,
//                  HEAVY_INSIDER_SELLING" | "NO_INSIDER_DATA"
const SCORES_HEADERS = [
  "Ticker", "Name", "Tier", "Sector", "Industry",
  "Quality_Score", "Value_Score", "Growth_Score", "Fundamentals_Score", "Moonshot_Score",
  "Volatility_Score", "Reliability_Score", "Earnings_Archetype_Code",
  "Qualified", "Disqualify_Reason_Codes",
  "Verdict", "Caution_Flags", "Last_Updated"
];
/*************************************************************
 * safeFetch — retrying HTTP fetch with realistic headers.
 *************************************************************/
function safeFetch(url, options = {}, retries = 4) {
  const defaultHeaders = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9',
    'Connection': 'keep-alive',
    'Referer': 'https://finviz.com/'
  };
  const opt = Object.assign({}, options);
  opt.headers = Object.assign({}, defaultHeaders, options.headers || {});
  opt.muteHttpExceptions = true;
  opt.followRedirects = true;
  opt.method = opt.method || 'get';
  let attempt = 0;
  let delay = 250;
  while (attempt < retries) {
    try {
      const res = UrlFetchApp.fetch(url, opt);
      const code = res.getResponseCode();
      if (code >= 200 && code < 300) return res;
      // 404/410 mean "this page doesn't exist" — permanent, not transient.
      // Retrying wastes ~1.75s per ticker on every obscure/OTC name Finviz
      // simply has no page for (common at full-universe scale) — fail fast
      // instead of burning the whole retry budget on something that will
      // never succeed no matter how many times we ask.
      if (code === 404 || code === 410) {
        throw new Error(`safeFetch: HTTP ${code} (not found) for ${url}`);
      }
      Logger.log(`safeFetch non-OK HTTP ${code} for ${url} (attempt ${attempt + 1})`);
    } catch (e) {
      Logger.log(`safeFetch exception for ${url} attempt ${attempt + 1}: ${e}`);
      if (attempt === retries - 1) throw e;
      if (String(e).includes('HTTP 404') || String(e).includes('HTTP 410')) throw e;
    }
    Utilities.sleep(delay);
    attempt++;
    delay = Math.min(delay * 2, 5000);
  }
  throw new Error('safeFetch: exceeded retries for ' + url);
}
/*************************************************************
 * htmlDecode — minimal HTML entity decoding for scraped text.
 *************************************************************/
function htmlDecode(s) {
  if (s == null) return s;
  return String(s)
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .trim();
}
/*************************************************************
 * formatSheet — consistent formatting applied after every write.
 *************************************************************/
function formatSheet(sheet, colCount, options = {}) {
  if (!sheet) return;
  const {
    boldHeader = true,
    freezeHeader = true,
    freezeTickerColumn = true,
    autoSize = true
  } = options;
  if (boldHeader) sheet.getRange(1, 1, 1, colCount).setFontWeight("bold");
  if (freezeHeader) sheet.setFrozenRows(1);
  if (freezeTickerColumn) sheet.setFrozenColumns(1);
  if (autoSize) {
    for (let c = 1; c <= colCount; c++) {
      try { sheet.autoResizeColumn(c); } catch (e) { /* ignore column-not-found edge cases */ }
    }
  }
}
// ---- INSIDERS-SPECIFIC CONFIG ----
// Same weekly-refresh cadence convention as every other script.
const INSIDERS_STALENESS_DAYS = 6;
// Wider than a typical "keep 200 days" choice — this needs to be wide
// enough to support a true trailing-12-month Buy Signal calculation,
// so use 370, not 200 or 365.
const INSIDERS_RETENTION_DAYS = 370;
const TRANSACTION_CODES = {
  P: "Purchase (Open market/private)",
  S: "Sale (Open market/private)",
  V: "Voluntarily reported early",
  A: "Grant/Award/Acquisition (Rule 16b-3(d))",
  D: "Disposition to issuer (Rule 16b-3(e))",
  F: "Payment of exercise price or tax liability",
  I: "Discretionary transaction (Rule 16b-3(f))",
  M: "Exercise/conversion of derivative (Rule 16b-3)",
  C: "Conversion of derivative security",
  E: "Expiration of short derivative position",
  H: "Expiration/cancellation of long derivative position",
  O: "Exercise of out-of-the-money derivative security",
  X: "Exercise of in/at-the-money derivative security",
  G: "Bona fide gift",
  L: "Small acquisition (Rule 16a-6)",
  W: "Acquisition/disposition by will/laws of descent",
  Z: "Deposit/withdrawal from voting trust",
  J: "Other acquisition/disposition",
  K: "Equity swap transaction",
  U: "Disposition via tender/change of control"
};
const INSIDERS_HEADERS = [
  "Ticker", "Name", "Transaction_Date", "Filing_Date", "Change",
  "Shares", "Transaction_Price", "Transaction_Code",
  "Transaction_Code_Description", "Is_Derivative", "Last_Updated"
];
const INSIDERS_CHECKPOINT_HEADERS = ["Ticker", "Last_Checked"];
const INSIDER_SIGNALS_HEADERS = [
  "Ticker", "Buy_Signal_Score", "Sell_Urgency_Score",
  "Net_Value_Bought_12M", "Net_Value_Sold_90D",
  "Distinct_Sellers_90D", "Last_Updated"
];
// ---- PRESENTATION LAYER CONFIG ----
// All four sheets below are pure computation from data already in the
// core pipeline — no API calls, no chunking needed, full rebuild every
// run (same synchronous pattern as RefData.gs/TierBenchmarks.gs).
const BUY_LIST_HEADERS = [
  "Ticker", "Name", "Tier", "Sector", "Fundamentals_Score", "Value_Score",
  "Current_Price", "Week52_Low", "Week52_High", "Pct_Of_52W_Range", "Entry_Timing",
  "Verdict", "Caution_Flags", "Last_Updated"
];
const MOONSHOT_WATCHLIST_HEADERS = [
  "Ticker", "Name", "Sector", "Industry", "Moonshot_Score", "Caution_Flags", "Last_Updated"
];
const EARNINGS_ALERTS_HEADERS = [
  "Ticker", "Name", "Earnings_Date", "Earnings_Session", "Days_Until",
  "Volatility_Score", "Reliability_Score", "Earnings_Archetype_Code", "Last_Updated"
];
// How many days ahead counts as "upcoming" for the alerts sheet.
const EARNINGS_ALERT_WINDOW_DAYS = 14;
const INSIDER_FEED_HEADERS = [
  "Ticker", "Name", "Insider_Name", "Transaction_Date", "Transaction_Code", "Transaction_Type",
  "Shares", "Value", "Pct_Of_Market_Cap", "Cluster_Flag", "Last_Updated"
];
// Only open-market Purchase(P)/Sale(S) codes are economically meaningful
// for this feed — same scope Insiders.gs already uses for Buy_Signal/
// Sell_Urgency, grants/exercises/gifts are noise for this purpose.
const INSIDER_FEED_WINDOW_DAYS = 30;
const INSIDER_FEED_MIN_VALUE = 25000; // materiality floor, in dollars