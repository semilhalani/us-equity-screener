/*************************************************************
 * DebugTools.gs — quick status checks
 * -----------------------------------------------------------
 * Run checkPipelineStatus() any time you want to know "is
 * something still running, and how far did it get" without
 * digging through the Executions log.
 *************************************************************/
function checkPipelineStatus() {
  const props = PropertiesService.getScriptProperties();
  const module = props.getProperty(PROP.CHUNK_MODULE);
  const index = props.getProperty(PROP.CHUNK_INDEX);
  const total = props.getProperty(PROP.CHUNK_TOTAL);
  const size = props.getProperty(PROP.CHUNK_SIZE);
  if (!module) {
    Logger.log("No chunk process currently running or queued.");
    return;
  }
  const pct = (total && Number(total) > 0) ? Math.round((Number(index) / Number(total)) * 100) : 0;
  Logger.log(`Active process: ${module}`);
  Logger.log(`Progress: ${index} / ${total} (${pct}%), chunk size ${size}`);
  const triggers = ScriptApp.getProjectTriggers();
  const workerTrigger = triggers.find(t => t.getHandlerFunction() === "runChunkWorker");
  Logger.log(workerTrigger
    ? "Worker trigger: ACTIVE (should tick roughly every 1 minute)"
    : "Worker trigger: MISSING — autoResumeChunk should recreate it within the hour, or just call the same runX function again."
  );
}
/**
 * showAllTriggers() — lists every trigger currently registered,
 * useful if something seems stuck and you want to see what's
 * actually scheduled.
 */
function showAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    Logger.log("No triggers currently registered.");
    return;
  }
  triggers.forEach(t => {
    Logger.log(`Function: ${t.getHandlerFunction()}, Type: ${t.getEventType()}`);
  });
}
/**
 * checkCellUsage() — reports every sheet's cell footprint, sorted by
 * ALLOCATED size (getMaxRows() x getMaxColumns()) descending, which is
 * what actually counts against the workbook's hard 10,000,000-cell
 * limit — NOT how many cells actually contain data. A sheet that's had
 * clearContents() called on it (rebuilt in place, like Raw_Universe or
 * Raw_Fundamentals) keeps its OLD row/column count even though the
 * visible data is smaller now; only actually deleting rows/columns (or
 * the whole sheet) shrinks the allocated grid. This is why "used" and
 * "allocated" are reported separately below — a big gap between them
 * on any sheet is wasted grid space counting against the limit for no
 * reason. Hidden sheets (the *Scratch sheets) count too.
 *
 * Run this any time the 10M-cell error comes up, or just to sanity-check
 * things periodically as the universe grows.
 */
function checkCellUsage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const rows = sheets.map(sheet => {
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    return {
      name: sheet.getName(),
      maxRows, maxCols,
      allocatedCells: maxRows * maxCols,
      lastRow, lastCol,
      usedCells: lastRow * lastCol
    };
  });
  rows.sort((a, b) => b.allocatedCells - a.allocatedCells);
  let totalAllocated = 0;
  Logger.log("=== Cell usage by sheet (sorted by ALLOCATED cells, descending) ===");
  rows.forEach(r => {
    totalAllocated += r.allocatedCells;
    const wasted = r.allocatedCells - r.usedCells;
    Logger.log(
      `${r.name}: allocated ${r.maxRows} x ${r.maxCols} = ${r.allocatedCells} cells | ` +
      `used ${r.lastRow} x ${r.lastCol} = ${r.usedCells} cells | wasted (allocated-but-empty): ${wasted}`
    );
  });
  Logger.log(`=== TOTAL allocated cells across workbook: ${totalAllocated} (hard limit: 10,000,000) ===`);
}
/**
 * oneTimeCleanup_SharesHistoryAndGatesSchema() — RUN ONCE, then feel
 * free to delete this function. Does two things as part of retiring
 * the free Shares_Outstanding-history dilution method:
 *
 * 1. Deletes Raw_SharesHistory entirely (it's no longer written to or
 *    read by anything once the Constants.gs / Universe.gs /
 *    DisqualifierGates.gs updates are in place).
 * 2. Clears Raw_DisqualifierGates completely, INCLUDING its header row.
 *    This is necessary, not optional: the column layout changed (Notes
 *    columns replaced with clean Dilution/Zero_Revenue/Cash_Runway
 *    columns), but _ensureDisqualifierGatesHeaders() only writes new
 *    headers when the header row is currently EMPTY — so without this
 *    step, the OLD header text would silently stick around right above
 *    NEWLY-shaped data rows on your next run, which is exactly the kind
 *    of silent mismatch this project has been careful to avoid
 *    elsewhere. Since this sheet currently only has ~20 test rows in
 *    it, there's nothing meaningful to lose by clearing it.
 */
function oneTimeCleanup_SharesHistoryAndGatesSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sharesHistorySheet = ss.getSheetByName("Raw_SharesHistory");
  if (sharesHistorySheet) {
    const rowCount = sharesHistorySheet.getLastRow();
    ss.deleteSheet(sharesHistorySheet);
    Logger.log(`Deleted Raw_SharesHistory (had ${rowCount} rows).`);
  } else {
    Logger.log("Raw_SharesHistory not found — already deleted, or never created.");
  }
  const gatesSheet = ss.getSheetByName(SHEETS.DISQUALIFIER_GATES);
  if (gatesSheet) {
    const rowCount = gatesSheet.getLastRow();
    gatesSheet.clearContents();
    Logger.log(`Cleared Raw_DisqualifierGates (had ${rowCount} rows including header) — will be rebuilt with the new column layout on the next run.`);
  } else {
    Logger.log("Raw_DisqualifierGates not found — nothing to clear.");
  }
  SpreadsheetApp.getActive().toast("Cleanup complete — see the Executions log for details.");
}