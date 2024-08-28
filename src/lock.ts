export function scriptLock<T extends Function>(fn: T) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(1000)) {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        "Another script is running in the background. Please wait for it to finish then try again"
      );
      return;
    }
    try {
      fn();
    } finally {
      lock.releaseLock();
    }
  }

function documentLock<T extends Function>(fn: T) {
    const lock = LockService.getDocumentLock();
    if (!lock.tryLock(1000)) {
      const ui = SpreadsheetApp.getUi();
      ui.alert("Document is locked. Please wait and try again later");
      return;
    }
    try {
      fn();
    } finally {
      lock.releaseLock();
    }
  }