/*************************************************************
 * TierBenchmarks.gs — our own tier-aware benchmarks
 * -----------------------------------------------------------
 * Finviz's industry/sector averages (Raw_RefIndustries/Sectors)
 * blend every market cap together. A mega-cap and a nano-cap in
 * the same industry don't really belong on the same PE benchmark
 * — this script computes our OWN median PE/ROE/margin/growth/D-E
 * per (Industry-or-Sector, Tier) combination instead, using the
 * real data we've already fetched into Raw_Fundamentals.
 *
 * MEDIAN, not mean — a handful of extreme outliers (a nano-cap
 * with 5,000% growth off a tiny base) barely move a median, but
 * would badly distort a mean. This is the main outlier defense
 * for the benchmark itself; percentile-scoring of INDIVIDUAL
 * stocks against these benchmarks (which needs its own outlier
 * capping) is a ScoringEngine.gs concern for later.
 *
 * PE is filtered to POSITIVE values only before taking the median
 * — a negative PE means an unprofitable company, which isn't a
 * meaningful data point for "what does a fairly-priced profitable
 * peer trade at." ROE/margin/growth/D-E keep negative values,
 * since those ARE meaningful signal for those metrics.
 *
 * SAMPLE SIZE HONESTY: any (Group, Tier) combination with fewer
 * than MIN_BENCHMARK_SAMPLE_SIZE tickers gets Low_Sample_Flag =
 * true. Both Industry-level AND Sector-level (broader, more
 * data) benchmarks are computed side by side — ScoringEngine.gs
 * decides later which to trust for a given stock, falling back
 * to Sector when Industry's sample is too thin. This script's
 * job is just to compute both honestly, not to pick one.
 *
 * ENTRY POINT: runTierBenchmarksRefresh()
 *************************************************************/

function runTierBenchmarksRefresh() {
  const uniSheet = _getOrCreateUniverseSheet();
  const fundSheet = _getOrCreateFundamentalsSheet();

  const uniLast = uniSheet.getLastRow();
  const fundLast = fundSheet.getLastRow();

  if (uniLast < 2 || fundLast < 2) {
    SpreadsheetApp.getActive().toast("TierBenchmarks: need both Raw_Universe and Raw_Fundamentals populated first.");
    Logger.log("TierBenchmarks: aborting — Universe or Fundamentals sheet is empty.");
    return;
  }

  // Ticker -> Tier map from Raw_Universe
  const uniData = uniSheet.getRange(2, 1, uniLast - 1, UNIVERSE_HEADERS.length).getValues();
  const tierByTicker = new Map();
  uniData.forEach(r => {
    const ticker = String(r[0] || "").trim().toUpperCase();
    const tier = r[7]; // column H = Tier
    if (ticker) tierByTicker.set(ticker, tier);
  });

  const fundData = fundSheet.getRange(2, 1, fundLast - 1, FUNDAMENTALS_HEADERS.length).getValues();

  if (fundData.length < uniData.length * 0.5) {
    Logger.log(`TierBenchmarks: WARNING — Raw_Fundamentals has ${fundData.length} rows vs Raw_Universe's ${uniData.length}. ` +
               `Benchmarks will compute from whatever exists, but consider running the full runFundamentalsRefresh() first for reliable coverage.`);
  }

  const groups = new Map(); // key: "Type|Name|Tier" -> accumulator

  function _getGroup(type, name, tier) {
    if (!name || !tier) return null;
    const key = `${type}|${name}|${tier}`;
    if (!groups.has(key)) {
      groups.set(key, {
        type, name, tier,
        pe: [], roe: [], npm: [], revGrowth3Y: [], epsGrowth3Y: [], debtEquity: []
      });
    }
    return groups.get(key);
  }

  fundData.forEach(r => {
    const ticker = String(r[0] || "").trim().toUpperCase();
    const sector = r[1];
    const industry = r[2];
    const pe = r[5];            // PE
    const roe = r[10];          // ROE
    const npm = r[12];          // Net_Profit_Margin
    const revGrowth3Y = r[14];  // Revenue_Growth_3Y
    const epsGrowth3Y = r[18];  // EPS_Growth_3Y
    const debtEquity = r[22];   // Debt_Equity

    const tier = tierByTicker.get(ticker);
    if (!tier) return;

    const indGroup = _getGroup("Industry", industry, tier);
    if (indGroup) {
      indGroup.pe.push(pe); indGroup.roe.push(roe); indGroup.npm.push(npm);
      indGroup.revGrowth3Y.push(revGrowth3Y); indGroup.epsGrowth3Y.push(epsGrowth3Y); indGroup.debtEquity.push(debtEquity);
    }
    const secGroup = _getGroup("Sector", sector, tier);
    if (secGroup) {
      secGroup.pe.push(pe); secGroup.roe.push(roe); secGroup.npm.push(npm);
      secGroup.revGrowth3Y.push(revGrowth3Y); secGroup.epsGrowth3Y.push(epsGrowth3Y); secGroup.debtEquity.push(debtEquity);
    }
  });

  const outRows = [];
  groups.forEach(g => {
    const sampleCount = g.pe.length;
    outRows.push([
      g.type,
      g.name,
      g.tier,
      sampleCount,
      _median(_extractPositiveNumeric(g.pe)),
      _median(_extractNumeric(g.roe)),
      _median(_extractNumeric(g.npm)),
      _median(_extractNumeric(g.revGrowth3Y)),
      _median(_extractNumeric(g.epsGrowth3Y)),
      _median(_extractNumeric(g.debtEquity)),
      sampleCount < MIN_BENCHMARK_SAMPLE_SIZE,
      new Date()
    ]);
  });

  const sheet = _getOrCreateTierBenchmarksSheet();
  sheet.clearContents();
  sheet.getRange(1, 1, 1, TIER_BENCHMARKS_HEADERS.length).setValues([TIER_BENCHMARKS_HEADERS]);
  if (outRows.length > 0) {
    sheet.getRange(2, 1, outRows.length, TIER_BENCHMARKS_HEADERS.length).setValues(outRows);
  }
  formatSheet(sheet, TIER_BENCHMARKS_HEADERS.length);

  SpreadsheetApp.getActive().toast(`TierBenchmarks: computed ${outRows.length} group benchmarks from ${fundData.length} tickers.`);
}

/*************************************************************
 * Numeric helpers
 *************************************************************/

function _extractNumeric(values) {
  return values.filter(v => typeof v === "number" && !isNaN(v) && isFinite(v));
}
function _extractPositiveNumeric(values) {
  return _extractNumeric(values).filter(v => v > 0);
}
function _median(nums) {
  if (!nums || nums.length === 0) return "";
  const sorted = nums.slice().sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return (sorted.length % 2 === 0) ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
}

/*************************************************************
 * Sheet helper
 *************************************************************/

function _getOrCreateTierBenchmarksSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.TIER_BENCHMARKS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.TIER_BENCHMARKS);
  return sheet;
}