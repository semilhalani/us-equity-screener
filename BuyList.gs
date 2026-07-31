/*************************************************************
 * BuyList.gs — 📋 BUY LIST presentation sheet, the flagship
 * -----------------------------------------------------------
 * Combines the three long-term "how good, and how's it priced right
 * now" scenarios into one view, differentiated by Entry_Timing:
 *   BUY_NOW                — quality stock, price in a reasonable
 *                            mid-range zone of its 52-week range
 *   STRONG_BUY_UNDERVALUED — quality stock, price near its 52-week low
 *   WAIT_FOR_PULLBACK      — quality stock, price near its 52-week high
 * Moonshot tier is excluded — it has its own watchlist with its own
 * entirely different evaluation logic.
 *
 * LIVE PRICING: Current_Price, Pct_Of_52W_Range, and Entry_Timing are
 * written as GOOGLEFINANCE-based FORMULAS, not values Apps Script
 * fetches — GOOGLEFINANCE doesn't touch the UrlFetchApp quota and
 * resolves/recalculates on its own inside Sheets, so these three
 * columns stay live between refreshes without needing to re-run this
 * script. Apps Script writes the formulas once per run; Sheets keeps
 * them current. Give Sheets a few seconds after a run to resolve them
 * across every row.
 *
 * DATA-QUALITY GUARD (per council review of the scoring methodology —
 * see Scoring_System_Explained.md): a benchmark-lookup miss makes
 * Quality/Value/Growth fall back to a neutral 50 each. A
 * Fundamentals_Score built from two fake-neutral inputs and one real
 * one could still cross the STRONG_CANDIDATE/WATCH threshold and look
 * like a real recommendation. So this list requires Quality_Score,
 * Value_Score, AND Growth_Score to all be genuinely non-blank — not
 * just Fundamentals_Score crossing a threshold — before a ticker is
 * included.
 *
 * ENTRY POINT: runBuyListRefresh()
 *************************************************************/
function runBuyListRefresh() {
  const scoresSheet = _getOrCreateScoresSheet();
  const sLast = scoresSheet.getLastRow();
  if (sLast < 2) {
    SpreadsheetApp.getActive().toast("BuyList: Scores is empty — run runScoringEngineRefresh() first.");
    return;
  }
  const fundSheet = _getOrCreateFundamentalsSheet();
  const fLast = fundSheet.getLastRow();
  const week52ByTicker = new Map();
  if (fLast >= 2) {
    fundSheet.getRange(2, 1, fLast - 1, FUNDAMENTALS_HEADERS.length).getValues().forEach(r => {
      const t = String(r[0] || "").trim().toUpperCase();
      if (t) week52ByTicker.set(t, { low: r[3], high: r[4] }); // Week52_Low, Week52_High
    });
  }
  const scoresData = scoresSheet.getRange(2, 1, sLast - 1, SCORES_HEADERS.length).getValues();
  const candidates = scoresData.filter(r => {
    const tier = r[2], qualified = r[13], verdict = r[15];
    const qualityScore = r[5], valueScore = r[6], growthScore = r[7];
    if (tier === TIER.MOONSHOT) return false; // has its own watchlist
    if (qualified !== true) return false;
    if (verdict !== "STRONG_CANDIDATE" && verdict !== "WATCH") return false;
    // Data-quality guard — see header comment above.
    if (typeof qualityScore !== "number" || typeof valueScore !== "number" || typeof growthScore !== "number") return false;
    return true;
  }).sort((a, b) => {
    // Best Fundamentals_Score first — without this, rows land in
    // whatever arbitrary order Raw_Fundamentals happened to build them
    // in, which has nothing to do with how attractive each one is.
    const av = (typeof a[8] === "number") ? a[8] : -Infinity;
    const bv = (typeof b[8] === "number") ? b[8] : -Infinity;
    return bv - av;
  });
  const valueRows = candidates.map(r => {
    const ticker = r[0];
    const w52 = week52ByTicker.get(String(ticker).trim().toUpperCase()) || {};
    const week52Low = (typeof w52.low === "number") ? w52.low : "";
    const week52High = (typeof w52.high === "number") ? w52.high : "";
    return [
      ticker, r[1], r[2], r[3],   // Ticker, Name, Tier, Sector
      r[8], r[6],                 // Fundamentals_Score, Value_Score (already rounded in Scores)
      "",                         // Current_Price — formula, filled below
      week52Low, week52High,
      "", "",                     // Pct_Of_52W_Range, Entry_Timing — formulas, filled below
      r[15], r[16],                // Verdict, Caution_Flags
      new Date()
    ];
  });
  const sheet = _getOrCreateBuyListSheet();
  sheet.clearContents();
  sheet.getRange(1, 1, 1, BUY_LIST_HEADERS.length).setValues([BUY_LIST_HEADERS]);
  if (valueRows.length > 0) {
    sheet.getRange(2, 1, valueRows.length, BUY_LIST_HEADERS.length).setValues(valueRows);
    // Column letters per BUY_LIST_HEADERS: A=Ticker G=Current_Price(7)
    // H=Week52_Low(8) I=Week52_High(9) J=Pct_Of_52W_Range(10) K=Entry_Timing(11)
    const priceFormulas = [];
    const rangeFormulas = [];
    const timingFormulas = [];
    for (let i = 0; i < valueRows.length; i++) {
      const row = i + 2;
      priceFormulas.push([`=IFERROR(GOOGLEFINANCE(A${row},"price"),"")`]);
      rangeFormulas.push([`=IFERROR((G${row}-H${row})/(I${row}-H${row}),"")`]);
      timingFormulas.push([`=IFERROR(IF(J${row}="","",IF(J${row}<=0.25,"STRONG_BUY_UNDERVALUED",IF(J${row}>=0.75,"WAIT_FOR_PULLBACK","BUY_NOW"))),"")`]);
    }
    sheet.getRange(2, 7, priceFormulas.length, 1).setFormulas(priceFormulas);
    sheet.getRange(2, 10, rangeFormulas.length, 1).setFormulas(rangeFormulas);
    sheet.getRange(2, 11, timingFormulas.length, 1).setFormulas(timingFormulas);
  }
  formatSheet(sheet, BUY_LIST_HEADERS.length);
  SpreadsheetApp.getActive().toast(
    `BuyList: ${valueRows.length} qualified candidates listed. Current price / entry timing are live formulas — ` +
    `give Sheets a few seconds to resolve GOOGLEFINANCE after this run.`
  );
}
function _getOrCreateBuyListSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.BUY_LIST);
  if (!sheet) sheet = ss.insertSheet(SHEETS.BUY_LIST);
  return sheet;
}