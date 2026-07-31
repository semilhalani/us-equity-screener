/*************************************************************
 * ChunkEngine.gs — Shared chunked-execution infrastructure
 * -----------------------------------------------------------
 * KEY DESIGN PRINCIPLE (fixes the old system's fragility):
 * We never store ticker arrays or large payloads in Script
 * Properties — Properties have a ~9KB per-value limit, and
 * the old system's PROP.FUND_INPUT_SHEET bug came from
 * exactly this mistake (storing a JSON ticker array in a
 * property, then confusing it with a sheet-name property
 * elsewhere).
 *
 * Instead: every worker function receives (startIndex, chunkSize)
 * and is responsible for deriving its own slice of work — either
 * by reading a row range from a sheet (the normal case), or by
 * indexing into a small global constant array (for tiny, fixed
 * lists like our 9 Finviz filters).
 *
 * ALSO FIXED: overlapping trigger executions. If a chunk takes
 * longer than the 1-minute trigger interval, the next tick could
 * fire while the previous one is still running, racing on the
 * same CHUNK_INDEX. LockService.getScriptLock() prevents this —
 * a busy tick just skips cleanly instead of corrupting state.
 *
 * USAGE:
 *   startChunkProcess(moduleName, workerFnName, finalizeFnName, totalItems, chunkSize?)
 *   worker signature:   function myWorker(startIndex, chunkSize) { ...; return nextIndex; }
 *   finalize signature: function myFinalize() { ... }
 *************************************************************/

function startChunkProcess(moduleName, workerFnName, finalizeFnName, totalItems, chunkSize) {
  stopChunkProcess(`Replacing previous chunk process for ${moduleName}…`);

  const props = PropertiesService.getScriptProperties();
  const size = chunkSize || DEFAULT_CHUNK_SIZE;

  props.setProperty(PROP.CHUNK_MODULE, moduleName);
  props.setProperty(PROP.CHUNK_WORKER, workerFnName);
  props.setProperty(PROP.CHUNK_FINALIZE, finalizeFnName || "");
  props.setProperty(PROP.CHUNK_INDEX, "0");
  props.setProperty(PROP.CHUNK_TOTAL, String(totalItems || 0));
  props.setProperty(PROP.CHUNK_SIZE, String(size));

  Utilities.sleep(1000); // let previous triggers fully clear

  try {
    ScriptApp.newTrigger("runChunkWorker").timeBased().everyMinutes(1).create();
    ScriptApp.newTrigger("autoResumeChunk").timeBased().everyHours(1).create();
  } catch (e) {
    Logger.log(`startChunkProcess: trigger creation failed: ${e}`);
  }

  Logger.log(`Chunk process started for ${moduleName}, totalItems=${totalItems}, chunkSize=${size}`);
  SpreadsheetApp.getActive().toast(`${moduleName}: started (${totalItems} items, chunk=${size})`);
}

function stopChunkProcess(reason = "Stopped chunking") {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    const fn = t.getHandlerFunction();
    if (fn === "runChunkWorker" || fn === "autoResumeChunk") {
      try { ScriptApp.deleteTrigger(t); } catch (e) { Logger.log(`Error deleting trigger: ${e}`); }
    }
  }
  const props = PropertiesService.getScriptProperties();
  [PROP.CHUNK_MODULE, PROP.CHUNK_WORKER, PROP.CHUNK_FINALIZE,
   PROP.CHUNK_INDEX, PROP.CHUNK_TOTAL, PROP.CHUNK_SIZE]
   .forEach(k => props.deleteProperty(k));
  Logger.log(reason);
}

/*************************************************************
 * runChunkWorker — the 1-minute trigger target.
 * Wrapped in LockService so overlapping fires can't race.
 *************************************************************/
function runChunkWorker() {
  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(5000);
  if (!gotLock) {
    Logger.log("runChunkWorker: another invocation is still running — skipping this tick.");
    return;
  }
  try {
    _runChunkWorkerInner();
  } finally {
    lock.releaseLock();
  }
}

function _runChunkWorkerInner() {
  const props = PropertiesService.getScriptProperties();
  const workerName = props.getProperty(PROP.CHUNK_WORKER);
  const finalizeName = props.getProperty(PROP.CHUNK_FINALIZE);
  const index = parseInt(props.getProperty(PROP.CHUNK_INDEX) || "0", 10);
  const total = parseInt(props.getProperty(PROP.CHUNK_TOTAL) || "0", 10);
  const size = parseInt(props.getProperty(PROP.CHUNK_SIZE) || String(DEFAULT_CHUNK_SIZE), 10);

  if (!workerName || isNaN(total) || total <= 0) {
    return stopChunkProcess("❌ No active chunk process.");
  }

  const workerFn = this[workerName];
  if (typeof workerFn !== "function") {
    return stopChunkProcess(`❌ Worker function not found: ${workerName}`);
  }

  Logger.log(`Chunk worker running: ${workerName} @ index=${index}/${total}`);

  let nextIndex;
  try {
    nextIndex = workerFn(index, size);
  } catch (e) {
    Logger.log(`Worker threw error: ${e}`);
    return; // leave index unchanged — retried next minute
  }

  if (!isFinite(nextIndex) || nextIndex < index) {
    Logger.log(`Invalid nextIndex returned: ${nextIndex}`);
    return;
  }

  if (nextIndex >= total) {
    const moduleName = props.getProperty(PROP.CHUNK_MODULE);
    if (finalizeName && typeof this[finalizeName] === "function") {
      try { this[finalizeName](); } catch (e) { Logger.log(`Finalizer threw error: ${e}`); }
    }
    stopChunkProcess(`Completed all chunks for ${moduleName}.`);
    return;
  }

  props.setProperty(PROP.CHUNK_INDEX, String(nextIndex));
}

/*************************************************************
 * autoResumeChunk — hourly safety net.
 * Recreates the minute worker trigger if it was ever deleted
 * (e.g. Apps Script quota hiccup) while a job is still pending.
 *************************************************************/
function autoResumeChunk() {
  const props = PropertiesService.getScriptProperties();
  const index = parseInt(props.getProperty(PROP.CHUNK_INDEX) || "0", 10);
  const total = parseInt(props.getProperty(PROP.CHUNK_TOTAL) || "0", 10);
  if (isNaN(index) || isNaN(total) || total <= 0) return;
  if (index >= total) return;

  const triggers = ScriptApp.getProjectTriggers();
  const workerExists = triggers.some(t => t.getHandlerFunction() === "runChunkWorker");
  if (!workerExists) {
    ScriptApp.newTrigger("runChunkWorker").timeBased().everyMinutes(1).create();
    Logger.log("autoResumeChunk: recreated missing runChunkWorker trigger.");
  }
}