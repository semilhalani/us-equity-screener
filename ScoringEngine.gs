/*************************************************************
 * ScoringEngine.gs — Fundamentals Engine + Earnings Behavior (v3)
 * -----------------------------------------------------------
 * v3 CHANGE: applies the same clean-data/presentation-layer split
 * already used in DisqualifierGates.gs. Every previously-packed
 * emoji/sentence field is now a short, filterable code, and nothing
 * is ever left ambiguously blank:
 *   - Earnings_Archetype ("🎯 Volatility Harvesting Candidate") ->
 *     Earnings_Archetype_Code ("VOLATILITY_HARVEST")
 *   - Overall_Disqualified (TRUE = bad) -> Qualified (TRUE = good) —
 *     inverted so "give me the qualified ones" is the natural filter,
 *     and now genuinely blank ("") rather than defaulting to TRUE when
 *     DisqualifierGates has never actually run for a ticker — a ticker
 *     that was never evaluated is NOT the same as one that passed.
 *   - Disqualify_Reasons (blank when fine) -> Disqualify_Reason_Codes,
 *     explicit "NONE" when qualified, "NOT_YET_EVALUATED" when gates
 *     never ran, never a blank cell with an ambiguous meaning.
 *   - Action (one long sentence mixing verdict + reasoning + caveats)
 *     -> Verdict (single short enum) + Caution_Flags (comma-separated
 *     codes, explicit "NONE" when clean).
 * Presentation sheets (Buy List, Moonshot Watchlist, etc.) turn these
 * codes into whatever human sentence fits their own context — plain
 * English never gets stored redundantly as text in this raw sheet.
 *
 * EVERYTHING ELSE UNCHANGED FROM v2: Quality/Value/Growth/Moonshot
 * score formulas, tier-aware benchmarks, DisqualifierGates as a hard
 * block, Earnings Behavior Engine inputs. This is an output-shape
 * change only — no scoring methodology changed.
 *
 * ENTRY POINTS:
 *   runScoringEngineRefresh()              — full universe, writes Scores
 *   runScoringEngineRefresh_Test(limit)    — first N tickers only (default
 *                                            20), writes to Scores_TEST
 *                                            instead — never touches the
 *                                            real Scores sheet.
 *************************************************************/
function runScoringEngineRefresh(options) {
  const testLimit = (options && options.testLimit) || null;
  const uniSheet = _getOrCreateUniverseSheet();
  const fundSheet = _getOrCreateFundamentalsSheet();
  const benchSheet = _getOrCreateTierBenchmarksSheet();
  const gatesSheet = _getOrCreateDisqualifierGatesSheet();
  const catSheet = _getOrCreateEpsCategorySheet();
  const signalsSheet = _getOrCreateInsiderSignalsSheet();
  const uniLast = uniSheet.getLastRow();
  const fundLast = fundSheet.getLastRow();
  if (uniLast < 2 || fundLast < 2) {
    SpreadsheetApp.getActive().toast("ScoringEngine: need Raw_Universe and Raw_Fundamentals populated first.");
    Logger.log("ScoringEngine: aborting — Universe or Fundamentals sheet is empty.");
    return;
  }
  // ---- Ticker -> {name, tier} from Raw_Universe ----
  const uniData = uniSheet.getRange(2, 1, uniLast - 1, UNIVERSE_HEADERS.length).getValues();
  const uniByTicker = new Map();
  uniData.forEach(r => {
    const ticker = String(r[0] || "").trim().toUpperCase();
    if (ticker) uniByTicker.set(ticker, { name: r[1], tier: r[7] });
  });
  const fundData = testLimit
    ? fundSheet.getRange(2, 1, fundLast - 1, FUNDAMENTALS_HEADERS.length).getValues().slice(0, testLimit)
    : fundSheet.getRange(2, 1, fundLast - 1, FUNDAMENTALS_HEADERS.length).getValues();
  // ---- Benchmark map keyed "Type|Name|Tier" ----
  const benchLast = benchSheet.getLastRow();
  const benchByKey = new Map();
  if (benchLast >= 2) {
    const benchData = benchSheet.getRange(2, 1, benchLast - 1, TIER_BENCHMARKS_HEADERS.length).getValues();
    benchData.forEach(r => {
      const key = `${r[0]}|${r[1]}|${r[2]}`;
      benchByKey.set(key, {
        medianPE: r[4], medianROE: r[5], medianNPM: r[6],
        medianRevGrowth3Y: r[7], medianEPSGrowth3Y: r[8], medianDebtEquity: r[9],
        lowSample: r[10]
      });
    });
  } else {
    Logger.log("ScoringEngine: Raw_TierBenchmarks is empty — every score will fall back to neutral (50) benchmark comparisons. Run runTierBenchmarksRefresh() first for meaningful scores.");
  }
  // ---- Disqualification map keyed Ticker ----
  // NOTE: stores the full gate row, not just disqualified/reasons — so we
  // can tell "ticker has a gate row and passed" apart from "ticker has no
  // gate row at all" (DisqualifierGates never ran for it yet). Those are
  // NOT the same situation and shouldn't both silently read as "qualified".
  const gatesLast = gatesSheet.getLastRow();
  const gatesByTicker = new Map();
  if (gatesLast >= 2) {
    const gatesData = gatesSheet.getRange(2, 1, gatesLast - 1, DISQUALIFIER_GATES_HEADERS.length).getValues();
    gatesData.forEach(r => {
      const ticker = String(r[0] || "").trim().toUpperCase();
      // Overall_Disqualified = index 15, Disqualify_Reasons = index 16 in
      // the current (v4) Raw_DisqualifierGates layout.
      gatesByTicker.set(ticker, { disqualified: r[15], reasons: r[16] });
    });
  } else {
    Logger.log("ScoringEngine: Raw_DisqualifierGates is empty — every ticker will show Qualified as blank (not yet evaluated) until you run runDisqualifierGatesRefresh().");
  }
  // ---- Earnings Behavior map keyed Ticker ----
  const catLast = catSheet.getLastRow();
  const epsByTicker = new Map();
  if (catLast >= 2) {
    const catData = catSheet.getRange(2, 1, catLast - 1, EPS_CATEGORY_HEADERS.length).getValues();
    catData.forEach(r => {
      const ticker = String(r[0] || "").trim().toUpperCase();
      epsByTicker.set(ticker, { volatility: r[6], reliability: r[7], category: r[5] });
    });
  } else {
    Logger.log("ScoringEngine: Raw_EPSCategory is empty — Volatility_Score/Reliability_Score/Earnings_Archetype_Code will be blank until you run runEarningsRefresh().");
  }
  // ---- Insider Signal map keyed Ticker ----
  const signalsLast = signalsSheet.getLastRow();
  const insiderByTicker = new Map();
  if (signalsLast >= 2) {
    const signalsData = signalsSheet.getRange(2, 1, signalsLast - 1, INSIDER_SIGNALS_HEADERS.length).getValues();
    signalsData.forEach(r => {
      const ticker = String(r[0] || "").trim().toUpperCase();
      insiderByTicker.set(ticker, { buySignal: r[1], sellUrgency: r[2] });
    });
  } else {
    Logger.log("ScoringEngine: Raw_InsiderSignals is empty — Caution_Flags will note insider data as not yet available until you run runInsidersRefresh().");
  }
  const outRows = [];
  fundData.forEach(r => {
    const ticker = String(r[0] || "").trim().toUpperCase();
    if (!ticker) return;
    const sector = r[1];
    const industry = r[2];
    const pe = r[5];
    const roe = r[10];
    const roe3yr = r[11];
    const npm = r[12];
    const revGrowth3Y = r[14];
    const epsGrowth3Y = r[18];
    const debtEquity = r[22];
    const debtAssetNow = r[23];
    const debtAsset3yAgo = r[24];
    const divYieldPct = r[26];
    const uni = uniByTicker.get(ticker) || {};
    const tier = uni.tier || "";
    const name = uni.name || "";
    // ---- Qualified / Disqualify_Reason_Codes — never ambiguously blank ----
    const gateEntry = gatesByTicker.get(ticker);
    const hasGateData = !!gateEntry;
    const disqualified = hasGateData && gateEntry.disqualified === true;
    let qualified, disqualifyReasonCodes;
    if (!hasGateData) {
      qualified = "";                              // genuinely unknown — gates never ran
      disqualifyReasonCodes = "NOT_YET_EVALUATED";
    } else if (disqualified) {
      qualified = false;
      disqualifyReasonCodes = gateEntry.reasons || "UNKNOWN";
    } else {
      qualified = true;
      disqualifyReasonCodes = "NONE";
    }
    const eps = epsByTicker.get(ticker) || {};
    const volatilityScore = (typeof eps.volatility === "number") ? eps.volatility : "";
    const reliabilityScore = (typeof eps.reliability === "number") ? eps.reliability : "";
    const bench = _getBenchmarkForTicker(industry, sector, tier, benchByKey);
    let qualityScore = "", valueScore = "", growthScore = "", fundamentalsScore = "", moonshotScore = "";
    if (tier === TIER.MOONSHOT) {
      moonshotScore = _computeMoonshotScore(revGrowth3Y, npm, debtAssetNow, debtAsset3yAgo, bench);
    } else {
      qualityScore = _computeQualityScore(roe, npm, roe3yr, bench);
      valueScore = _computeValueScore(pe, debtEquity, divYieldPct, tier, bench);
      growthScore = _computeGrowthScore(revGrowth3Y, epsGrowth3Y, bench);
      fundamentalsScore = _computeFundamentalsComposite(qualityScore, valueScore, growthScore, tier);
    }
    const archetypeCode = _determineEarningsArchetypeCode(qualityScore, volatilityScore);
    const insider = insiderByTicker.get(ticker) || {};
    const buySignal = (typeof insider.buySignal === "number") ? insider.buySignal : null;
    const sellUrgency = (typeof insider.sellUrgency === "number") ? insider.sellUrgency : null;
    const { verdict, cautionFlags } = _determineVerdictAndCautions(
      disqualified, fundamentalsScore, moonshotScore, tier, reliabilityScore, archetypeCode, buySignal, sellUrgency
    );
    outRows.push([
      ticker, name, tier, sector, industry,
      _round1(qualityScore), _round1(valueScore), _round1(growthScore), _round1(fundamentalsScore), _round1(moonshotScore),
      _round1(volatilityScore), _round1(reliabilityScore), archetypeCode,
      qualified, disqualifyReasonCodes,
      verdict, cautionFlags,
      new Date()
    ]);
  });
  const sheet = testLimit ? _getOrCreateScoresTestSheet() : _getOrCreateScoresSheet();
  sheet.clearContents();
  sheet.getRange(1, 1, 1, SCORES_HEADERS.length).setValues([SCORES_HEADERS]);
  if (outRows.length > 0) {
    sheet.getRange(2, 1, outRows.length, SCORES_HEADERS.length).setValues(outRows);
  }
  formatSheet(sheet, SCORES_HEADERS.length);
  SpreadsheetApp.getActive().toast(
    testLimit
      ? `ScoringEngine TEST: scored ${outRows.length} ticker(s) into Scores_TEST — real Scores sheet untouched.`
      : `ScoringEngine: scored ${outRows.length} tickers.`
  );
}
function runScoringEngineRefresh_Test(limit) {
  return runScoringEngineRefresh({ testLimit: (typeof limit === "number" ? limit : DEFAULT_TEST_LIMIT) });
}
function _getOrCreateScoresTestSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Scores_TEST");
  if (!sheet) sheet = ss.insertSheet("Scores_TEST");
  return sheet;
}
/*************************************************************
 * Benchmark lookup — unchanged from v1
 *************************************************************/
function _getBenchmarkForTicker(industry, sector, tier, benchByKey) {
  const indKey = `Industry|${industry}|${tier}`;
  const secKey = `Sector|${sector}|${tier}`;
  const indBench = benchByKey.get(indKey);
  if (indBench && !indBench.lowSample) return indBench;
  const secBench = benchByKey.get(secKey);
  if (secBench) return secBench;
  return indBench || null;
}
/*************************************************************
 * Scoring primitives — unchanged from v1
 *************************************************************/
function _scoreVsBenchmark(stockValue, benchmarkMedian, higherIsBetter) {
  if (typeof stockValue !== "number" || isNaN(stockValue)) return "";
  if (typeof benchmarkMedian !== "number" || isNaN(benchmarkMedian) || benchmarkMedian <= 0) return 50;
  const ratio = stockValue / benchmarkMedian;
  const raw = higherIsBetter ? (50 * ratio) : (50 * (2 - ratio));
  return Math.max(0, Math.min(100, raw));
}
function _weightedAverage(componentsWithWeights) {
  const valid = componentsWithWeights.filter(([v]) => typeof v === "number" && !isNaN(v));
  if (valid.length === 0) return "";
  const totalWeight = valid.reduce((sum, [, w]) => sum + w, 0);
  if (totalWeight === 0) return "";
  const weightedSum = valid.reduce((sum, [v, w]) => sum + v * w, 0);
  return weightedSum / totalWeight;
}
function _round1(v) {
  return (typeof v === "number" && !isNaN(v)) ? Math.round(v * 10) / 10 : v;
}
/*************************************************************
 * Quality / Value / Growth — Tier 1/2/3 — unchanged from v1
 *************************************************************/
function _computeQualityScore(roe, npm, roe3yr, bench) {
  const roeScore = _scoreVsBenchmark(roe, bench ? bench.medianROE : null, true);
  const npmScore = _scoreVsBenchmark(npm, bench ? bench.medianNPM : null, true);
  let stabilityScore = "";
  if (typeof roe === "number" && typeof roe3yr === "number" && !isNaN(roe) && !isNaN(roe3yr)) {
    stabilityScore = Math.max(0, Math.min(100, 100 - Math.abs(roe - roe3yr) * 100));
  }
  return _weightedAverage([
    [roeScore, 0.45],
    [npmScore, 0.35],
    [stabilityScore, 0.20]
  ]);
}
function _computeValueScore(pe, debtEquity, divYieldPct, tier, bench) {
  const peScore = (typeof pe === "number" && pe > 0) ? _scoreVsBenchmark(pe, bench ? bench.medianPE : null, false) : "";
  const deScore = _scoreVsBenchmark(debtEquity, bench ? bench.medianDebtEquity : null, false);
  if (tier === TIER.CORE || tier === TIER.LARGE_MID) {
    const divScore = (typeof divYieldPct === "number" && !isNaN(divYieldPct))
      ? Math.max(0, Math.min(100, (divYieldPct / 0.05) * 100))
      : "";
    return _weightedAverage([
      [peScore, 0.55],
      [deScore, 0.25],
      [divScore, 0.20]
    ]);
  }
  return _weightedAverage([
    [peScore, 0.65],
    [deScore, 0.35]
  ]);
}
function _computeGrowthScore(revGrowth3Y, epsGrowth3Y, bench) {
  const revScore = _scoreVsBenchmark(revGrowth3Y, bench ? bench.medianRevGrowth3Y : null, true);
  const epsScore = _scoreVsBenchmark(epsGrowth3Y, bench ? bench.medianEPSGrowth3Y : null, true);
  return _weightedAverage([
    [revScore, 0.55],
    [epsScore, 0.45]
  ]);
}
function _computeFundamentalsComposite(qualityScore, valueScore, growthScore, tier) {
  let weights;
  if (tier === TIER.CORE) weights = [0.40, 0.35, 0.25];
  else if (tier === TIER.LARGE_MID) weights = [0.35, 0.30, 0.35];
  else weights = [0.30, 0.20, 0.50];
  return _weightedAverage([
    [qualityScore, weights[0]],
    [valueScore, weights[1]],
    [growthScore, weights[2]]
  ]);
}
/*************************************************************
 * Moonshot Score — Tier 4 — unchanged from v1
 *************************************************************/
function _computeMoonshotScore(revGrowth3Y, npm, debtAssetNow, debtAsset3yAgo, bench) {
  const revScore = _scoreVsBenchmark(revGrowth3Y, bench ? bench.medianRevGrowth3Y : null, true);
  const npmScore = _scoreVsBenchmark(npm, bench ? bench.medianNPM : null, true);
  let debtTrendScore = "";
  if (typeof debtAssetNow === "number" && typeof debtAsset3yAgo === "number" && !isNaN(debtAssetNow) && !isNaN(debtAsset3yAgo)) {
    const change = debtAssetNow - debtAsset3yAgo;
    debtTrendScore = Math.max(0, Math.min(100, 50 - change * 200));
  }
  return _weightedAverage([
    [revScore, 0.50],
    [npmScore, 0.30],
    [debtTrendScore, 0.20]
  ]);
}
/*************************************************************
 * Earnings Archetype — v3: returns a short code, not emoji+text.
 * Crosses Quality with Volatility instead of treating volatility as
 * a flat penalty. Volatility_Score is HIGH = volatile (intuitive
 * direction) — unchanged thresholds from v2, only the output values
 * changed.
 *************************************************************/
function _determineEarningsArchetypeCode(qualityScore, volatilityScore) {
  if (typeof qualityScore !== "number" || typeof volatilityScore !== "number") return "";
  const highQuality = qualityScore >= 55;
  const highSwing = volatilityScore >= 40;
  if (highQuality && highSwing) return "VOLATILITY_HARVEST";
  if (!highQuality && highSwing) return "HIGH_RISK";
  if (highQuality && !highSwing) return "STABLE_COMPOUNDER";
  return "LOW_CONVICTION";
}
/*************************************************************
 * Verdict + Caution_Flags determination — v3 replacement for the old
 * _determineAction(). Same decision logic as v2, but returns a short
 * enum (Verdict) and a set of short codes (Caution_Flags) instead of
 * building one long sentence. Disqualification is still a hard block,
 * checked first, regardless of everything else. Thresholds (insider
 * caveat >=60, archetype cutoffs inside _determineEarningsArchetypeCode)
 * are starting points, not yet tuned against real outcomes — revisit
 * once Track Record / backtest data exists.
 *************************************************************/
function _determineVerdictAndCautions(disqualified, fundamentalsScore, moonshotScore, tier, reliabilityScore, archetypeCode, buySignal, sellUrgency) {
  const cautions = [];
  if (buySignal === null && sellUrgency === null) {
    cautions.push("NO_INSIDER_DATA");
  } else {
    if (typeof sellUrgency === "number" && sellUrgency >= 60) cautions.push("HEAVY_INSIDER_SELLING");
    if (typeof buySignal === "number" && buySignal >= 60) cautions.push("INSIDER_BUYING_SUPPORT");
  }
  const finish = (verdict) => ({ verdict, cautionFlags: cautions.length > 0 ? cautions.join(", ") : "NONE" });

  if (disqualified) return finish("AVOID_DISQUALIFIED");

  if (tier === TIER.MOONSHOT) {
    if (typeof moonshotScore !== "number") return finish("INSUFFICIENT_DATA");
    if (moonshotScore >= 65) return finish("MOONSHOT_WATCH");
    if (moonshotScore >= 45) return finish("SPECULATIVE_WATCH");
    return finish("AVOID_WEAK_MOONSHOT");
  }

  if (typeof fundamentalsScore !== "number") return finish("INSUFFICIENT_DATA");

  let verdict;
  if (fundamentalsScore >= 70) verdict = "STRONG_CANDIDATE";
  else if (fundamentalsScore >= 55) verdict = "WATCH";
  else if (fundamentalsScore >= 40) verdict = "WEAK";
  else verdict = "AVOID_WEAK_FUNDAMENTALS";

  if (typeof reliabilityScore === "number" && reliabilityScore < 35 && (verdict === "STRONG_CANDIDATE" || verdict === "WATCH")) {
    cautions.push("INCONSISTENT_EARNINGS");
  }
  if (archetypeCode === "HIGH_RISK") {
    cautions.push("VOLATILE_AND_WEAK_FUNDAMENTALS");
  }

  return finish(verdict);
}
/*************************************************************
 * Sheet helper
 *************************************************************/
function _getOrCreateScoresSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.SCORES);
  if (!sheet) sheet = ss.insertSheet(SHEETS.SCORES);
  return sheet;
}