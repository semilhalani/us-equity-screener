/*************************************************************
 * RefData.gs — Industry & Sector benchmark averages
 * -----------------------------------------------------------
 * Scrapes Finviz's groups tables (one page each, no pagination
 * needed — unlike the screener, groups.ashx returns the whole
 * table in a single load). Fast enough to run synchronously,
 * no chunking required.
 *
 * COLUMN LAYOUT BELOW IS NOT A GUESS — it's reverse-engineered
 * from your own "Sam" file's working VLOOKUP formulas:
 *   VLOOKUP(x, 'US Industries'!B:D, 3, FALSE)  -> col D = PE
 *   VLOOKUP(x, 'US Industries'!B:K, 10, FALSE) -> col K = EPS past 5Y
 *   VLOOKUP(x, 'US Industries'!B:M, 12, FALSE) -> col M = Sales past 5Y
 * That fixes columns B, D, K, M exactly — everything between
 * them follows Finviz's standard groups-table order.
 *
 * IMPORTANT: this is the Finviz-published, cap-BLENDED average.
 * Useful context, but a mega-cap and a nano-cap in the same
 * industry don't really belong on the same benchmark. Once
 * Fundamentals.gs has populated every ticker, Raw_TierBenchmarks
 * (a later script) computes our OWN tier-aware medians instead.
 * This sheet stays as the reference/cross-check copy of what
 * Finviz itself reports.
 *
 * ENTRY POINT: runRefDataRefresh()
 *************************************************************/

function runRefDataRefresh() {
  _refreshGroupsTable("https://finviz.com/groups.ashx?g=industry&v=120&o=name", SHEETS.REF_INDUSTRIES, "Industry");
  _refreshGroupsTable("https://finviz.com/groups.ashx?g=sector&v=120&o=name", SHEETS.REF_SECTORS, "Sector");
  SpreadsheetApp.getActive().toast("RefData: industries + sectors refreshed.");
}

function _refreshGroupsTable(url, sheetName, labelForLog) {
  let html = "";
  try {
    const res = safeFetch(url, { headers: { "User-Agent": "Mozilla/5.0" } }, 4);
    html = res.getContentText();
  } catch (e) {
    Logger.log(`RefData: fetch failed for ${labelForLog}: ${e}`);
    return;
  }

  const tableRegex = /<table[^>]*\bgroups_table\b[^>]*>[\s\S]*?<\/table>/gi;
  const tables = html.match(tableRegex) || [];
  if (tables.length === 0) {
    Logger.log(`RefData: no groups_table found for ${labelForLog}`);
    return;
  }
  const tableHtml = tables.reduce((a, b) => (b.length > a.length ? b : a));

  const rowRegex = /<tr[\s\S]*?<\/tr>/gi;
  const rowsHtml = tableHtml.match(rowRegex) || [];
  if (rowsHtml.length === 0) {
    Logger.log(`RefData: no rows found for ${labelForLog}`);
    return;
  }

  const dataRows = [];
  rowsHtml.forEach(trHtml => {
    const thMatch = trHtml.match(/<th[^>]*>([\s\S]*?)<\/th>/gi);
    const tdMatch = trHtml.match(/<td[^>]*>([\s\S]*?)<\/td>/gi);
    if (thMatch && !tdMatch) return; // pure header row, skip
    if (tdMatch && tdMatch.length > 0) {
      dataRows.push(tdMatch.map(_normalizeGroupsCell));
    }
  });

  if (dataRows.length === 0) {
    Logger.log(`RefData: no data rows parsed for ${labelForLog}`);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, REF_HEADERS.length).setValues([REF_HEADERS]);

  // Defensive pad/trim — if Finviz ever tweaks column count slightly,
  // this stops setValues() from throwing a dimension-mismatch error.
  const fixedRows = dataRows.map(r => {
    const row = r.slice(0, REF_HEADERS.length);
    while (row.length < REF_HEADERS.length) row.push("");
    return row;
  });

  sheet.getRange(2, 1, fixedRows.length, REF_HEADERS.length).setValues(fixedRows);
  formatSheet(sheet, REF_HEADERS.length);
  Logger.log(`RefData: wrote ${fixedRows.length} rows for ${labelForLog}`);
}

function _normalizeGroupsCell(cellHtml) {
  if (!cellHtml) return "";
  const text = cellHtml.replace(/<[^>]+>/g, "");
  return htmlDecode(text);
}