/*************************************************************
 * InsiderFeed.gs — 🔍 INSIDER FEED presentation sheet
 * -----------------------------------------------------------
 * Pure computation from Raw_Insiders + Raw_Universe — no API calls,
 * no chunking needed, full rebuild every run. Shows open-market
 * Purchase(P)/Sale(S) transactions from the last
 * INSIDER_FEED_WINDOW_DAYS days, above the INSIDER_FEED_MIN_VALUE
 * materiality floor, newest first. Grants/exercises/gifts/other
 * codes are excluded — same scope Insiders.gs already restricts to
 * for Buy_Signal_Score/Sell_Urgency_Score, since those are the codes
 * that actually carry investment signal.
 *
 * ENTRY POINT: runInsiderFeedRefresh()
 *************************************************************/
function runInsiderFeedRefresh() {
  const insidersSheet = _getOrCreateInsidersSheet();
  const iLast = insidersSheet.getLastRow();
  if (iLast < 2) {
    SpreadsheetApp.getActive().toast("InsiderFeed: Raw_Insiders is empty — run runInsidersRefresh() first.");
    return;
  }
  const uniSheet = _getOrCreateUniverseSheet();
  const uniLast = uniSheet.getLastRow();
  const uniInfoByTicker = new Map();
  if (uniLast >= 2) {
    const uniData = uniSheet.getRange(2, 1, uniLast - 1, UNIVERSE_HEADERS.length).getValues();
    uniData.forEach(r => {
      const t = String(r[0] || "").trim().toUpperCase();
      if (t) uniInfoByTicker.set(t, { name: r[1], marketCap: r[5] });
    });
  }
  const windowCutoffMs = Date.now() - (INSIDER_FEED_WINDOW_DAYS * 24 * 60 * 60 * 1000);
  const insidersData = insidersSheet.getRange(2, 1, iLast - 1, INSIDERS_HEADERS.length).getValues();
  // First pass: collect qualifying transactions + build a cluster-detection
  // map (distinct insiders per Ticker+Code within the window — multiple
  // different people buying/selling together is a stronger signal than
  // one person transacting repeatedly).
  const qualifying = [];
  const clusterKeyCounts = new Map(); // "TICKER|CODE" -> Set of insider names
  insidersData.forEach(r => {
    const ticker = String(r[0] || "").trim().toUpperCase();
    const insiderName = r[1];
    const txDate = r[2] instanceof Date ? r[2] : new Date(r[2]);
    const code = r[7];
    const shares = Number(r[5]) || 0;
    const price = Number(r[6]) || 0;
    if (!ticker || isNaN(txDate.getTime())) return;
    if (txDate.getTime() < windowCutoffMs) return;
    if (code !== "P" && code !== "S") return; // open-market purchase/sale only
    const value = shares * price;
    if (value < INSIDER_FEED_MIN_VALUE) return; // below materiality floor
    const uni = uniInfoByTicker.get(ticker) || {};
    const clusterKey = `${ticker}|${code}`;
    if (!clusterKeyCounts.has(clusterKey)) clusterKeyCounts.set(clusterKey, new Set());
    clusterKeyCounts.get(clusterKey).add(insiderName);
    qualifying.push({ ticker, name: uni.name || "", insiderName, txDate, code, shares, value, marketCap: uni.marketCap, clusterKey });
  });
  const rows = qualifying
    .sort((a, b) => b.txDate - a.txDate)
    .map(q => {
      const pctOfMarketCap = (typeof q.marketCap === "number" && q.marketCap > 0) ? (q.value / q.marketCap) : "";
      const clusterFlag = clusterKeyCounts.get(q.clusterKey).size >= 2;
      return [
        q.ticker, q.name, q.insiderName, q.txDate, q.code,
        (q.code === "P" ? "Purchase" : "Sale"),
        q.shares, q.value, pctOfMarketCap, clusterFlag, new Date()
      ];
    });
  const sheet = _getOrCreateInsiderFeedSheet();
  sheet.clearContents();
  sheet.getRange(1, 1, 1, INSIDER_FEED_HEADERS.length).setValues([INSIDER_FEED_HEADERS]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, INSIDER_FEED_HEADERS.length).setValues(rows);
  }
  formatSheet(sheet, INSIDER_FEED_HEADERS.length);
  SpreadsheetApp.getActive().toast(`InsiderFeed: ${rows.length} transactions in the last ${INSIDER_FEED_WINDOW_DAYS} days above $${INSIDER_FEED_MIN_VALUE.toLocaleString()}.`);
}
function _getOrCreateInsiderFeedSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.INSIDER_FEED);
  if (!sheet) sheet = ss.insertSheet(SHEETS.INSIDER_FEED);
  return sheet;
}