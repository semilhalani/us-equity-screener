/*************************************************************
 * EarningsAlerts.gs — ⚠️ EARNINGS ALERTS presentation sheet
 * -----------------------------------------------------------
 * Pure computation from Raw_Fundamentals (forward-looking earnings
 * date, already scraped from Finviz by Fundamentals.gs) joined with
 * Scores for Volatility/Reliability/Archetype. No API calls, no
 * Raw_Calendar needed — that sheet's original purpose (real report
 * dates via /calendar/earnings) turned out unnecessary for alerts
 * specifically, since Finviz's forward-looking field already covers
 * this.
 *
 * DATE PARSING: Finviz's scraped Earnings_Date has no year (e.g.
 * "Jul 28"). _inferEarningsDate() assumes the current year unless
 * that lands more than 5 days in the past, in which case it assumes
 * next year — handles dates scraped near a year boundary. Returns
 * null (never a guess) if the text can't be parsed at all.
 *
 * ENTRY POINT: runEarningsAlertsRefresh()
 *************************************************************/
function runEarningsAlertsRefresh() {
  const fundSheet = _getOrCreateFundamentalsSheet();
  const fLast = fundSheet.getLastRow();
  if (fLast < 2) {
    SpreadsheetApp.getActive().toast("EarningsAlerts: Raw_Fundamentals is empty — run runFundamentalsRefresh() first.");
    return;
  }
  const uniSheet = _getOrCreateUniverseSheet();
  const uniLast = uniSheet.getLastRow();
  const nameByTicker = new Map();
  if (uniLast >= 2) {
    uniSheet.getRange(2, 1, uniLast - 1, 2).getValues().forEach(r => {
      const t = String(r[0] || "").trim().toUpperCase();
      if (t) nameByTicker.set(t, r[1]);
    });
  }
  const scoresSheet = _getOrCreateScoresSheet();
  const sLast = scoresSheet.getLastRow();
  const scoreByTicker = new Map();
  if (sLast >= 2) {
    scoresSheet.getRange(2, 1, sLast - 1, SCORES_HEADERS.length).getValues().forEach(r => {
      const t = String(r[0] || "").trim().toUpperCase();
      if (t) scoreByTicker.set(t, { volatility: r[10], reliability: r[11], archetype: r[12] });
    });
  }
  const fundData = fundSheet.getRange(2, 1, fLast - 1, FUNDAMENTALS_HEADERS.length).getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const windowEnd = new Date(today.getTime() + EARNINGS_ALERT_WINDOW_DAYS * 24 * 60 * 60 * 1000);
  const rows = [];
  fundData.forEach(r => {
    const ticker = String(r[0] || "").trim().toUpperCase();
    if (!ticker) return;
    const earningsDateRaw = r[27]; // Earnings_Date
    const earningsSession = r[28]; // Earnings_Session
    const parsedDate = _inferEarningsDate(earningsDateRaw, today);
    if (!parsedDate) return; // couldn't parse — skip, don't guess
    if (parsedDate < today || parsedDate > windowEnd) return; // outside the alert window
    const daysUntil = Math.round((parsedDate.getTime() - today.getTime()) / (24 * 60 * 60 * 1000));
    const sc = scoreByTicker.get(ticker) || {};
    rows.push([
      ticker, nameByTicker.get(ticker) || "", earningsDateRaw, earningsSession || "", daysUntil,
      (typeof sc.volatility === "number") ? sc.volatility : "",
      (typeof sc.reliability === "number") ? sc.reliability : "",
      sc.archetype || "",
      new Date()
    ]);
  });
  rows.sort((a, b) => a[4] - b[4]); // Days_Until ascending
  const sheet = _getOrCreateEarningsAlertsSheet();
  sheet.clearContents();
  sheet.getRange(1, 1, 1, EARNINGS_ALERTS_HEADERS.length).setValues([EARNINGS_ALERTS_HEADERS]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, EARNINGS_ALERTS_HEADERS.length).setValues(rows);
  }
  formatSheet(sheet, EARNINGS_ALERTS_HEADERS.length);
  SpreadsheetApp.getActive().toast(`EarningsAlerts: ${rows.length} tickers reporting within ${EARNINGS_ALERT_WINDOW_DAYS} days.`);
}
function _inferEarningsDate(raw, today) {
  if (!raw) return null;
  const str = String(raw).trim();
  if (!str) return null;
  const currentYear = today.getFullYear();
  let candidate = new Date(`${str} ${currentYear}`);
  if (isNaN(candidate.getTime())) return null;
  candidate.setHours(0, 0, 0, 0);
  const fiveDaysAgo = new Date(today.getTime() - 5 * 24 * 60 * 60 * 1000);
  if (candidate < fiveDaysAgo) {
    candidate = new Date(`${str} ${currentYear + 1}`);
    candidate.setHours(0, 0, 0, 0);
  }
  return candidate;
}
function _getOrCreateEarningsAlertsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.EARNINGS_ALERTS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.EARNINGS_ALERTS);
  return sheet;
}