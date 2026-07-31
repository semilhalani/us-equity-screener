/*************************************************************
 * MoonshotWatchlist.gs — 🎯 MOONSHOT WATCHLIST presentation sheet
 * -----------------------------------------------------------
 * Pure computation from Scores — no API calls, no chunking needed,
 * full rebuild every run (same synchronous pattern as
 * RefData.gs/TierBenchmarks.gs). Filters Scores down to Tier4_Moonshot
 * tickers that passed the disqualifier gates, sorted by Moonshot_Score
 * descending.
 *
 * ENTRY POINT: runMoonshotWatchlistRefresh()
 *************************************************************/
function runMoonshotWatchlistRefresh() {
  const scoresSheet = _getOrCreateScoresSheet();
  const last = scoresSheet.getLastRow();
  if (last < 2) {
    SpreadsheetApp.getActive().toast("MoonshotWatchlist: Scores is empty — run runScoringEngineRefresh() first.");
    return;
  }
  const data = scoresSheet.getRange(2, 1, last - 1, SCORES_HEADERS.length).getValues();
  const rows = data
    // Tier, Qualified, AND Moonshot_Score must be a real number — a
    // qualified moonshot ticker with no fundamentals data to score
    // (Verdict=INSUFFICIENT_DATA in Scores) has nothing to rank it by
    // and doesn't belong mixed in with real 0-100 scored tickers.
    .filter(r => r[2] === TIER.MOONSHOT && r[13] === true && typeof r[9] === "number")
    .map(r => [
      r[0], r[1], r[3], r[4], r[9], r[16], new Date() // Ticker,Name,Sector,Industry,Moonshot_Score,Caution_Flags
    ])
    .sort((a, b) => {
      const av = (typeof a[4] === "number") ? a[4] : -Infinity;
      const bv = (typeof b[4] === "number") ? b[4] : -Infinity;
      return bv - av;
    });
  const sheet = _getOrCreateMoonshotWatchlistSheet();
  sheet.clearContents();
  sheet.getRange(1, 1, 1, MOONSHOT_WATCHLIST_HEADERS.length).setValues([MOONSHOT_WATCHLIST_HEADERS]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, MOONSHOT_WATCHLIST_HEADERS.length).setValues(rows);
  }
  formatSheet(sheet, MOONSHOT_WATCHLIST_HEADERS.length);
  SpreadsheetApp.getActive().toast(`MoonshotWatchlist: ${rows.length} qualified moonshot tickers listed.`);
}
function _getOrCreateMoonshotWatchlistSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.MOONSHOT);
  if (!sheet) sheet = ss.insertSheet(SHEETS.MOONSHOT);
  return sheet;
}