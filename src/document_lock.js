// Execution-local ownership lets population call analysis without reacquiring.
let analyzerDocumentLockAlerts_ = null;

function withAnalyzerDocumentLock_(action, timer) {
  if (analyzerDocumentLockAlerts_ !== null) return action();

  const lock = LockService.getDocumentLock();
  if (!lock) throw new Error("Analyzer updates require a document lock.");
  const acquisitionStartedAt = Date.now();
  const alerts = [];
  let lockAcquired = false;
  let result;
  try {
    try {
      lock.waitLock(30000);
      lockAcquired = true;
    } finally {
      if (typeof endAnalyzerPhase_ === "function") {
        endAnalyzerPhase_(timer, "documentLockAcquisition", acquisitionStartedAt);
      }
    }
    analyzerDocumentLockAlerts_ = alerts;
    result = action();
  } finally {
    if (lockAcquired) {
      try {
        // Commit buffered spreadsheet writes while other editors are excluded.
        SpreadsheetApp.flush();
      } finally {
        analyzerDocumentLockAlerts_ = null;
        const releaseStartedAt = Date.now();
        try {
          lock.releaseLock();
        } finally {
          if (typeof endAnalyzerPhase_ === "function") {
            endAnalyzerPhase_(timer, "documentLockRelease", releaseStartedAt);
          }
        }
      }
    }
  }

  // Alerts suspend execution and discard locks, so display them after release.
  alerts.forEach(alert => alert.ui.alert.apply(alert.ui, alert.args));
  return result;
}

function showAnalyzerAlert_(ui, ...args) {
  if (analyzerDocumentLockAlerts_ !== null) {
    analyzerDocumentLockAlerts_.push({ ui, args });
    return;
  }
  return ui.alert.apply(ui, args);
}
