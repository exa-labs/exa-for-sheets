/**
 * WebsetsSync.gs — Orchestrates the Websets → Sheet sync workflow.
 *
 * Coordinates fetching items from the Exa Websets API (or demo data),
 * writing them to the "Webset Data" sheet, and managing time-driven
 * auto-refresh triggers.
 *
 * Depends on:
 *   - WebsetsClient.gs  → getWebset(), listAllWebsetItems()
 *   - SheetWriter.gs     → syncItemsToSheet()
 *   - Config.gs          → getWebsetsWebsetId(), isWebsetsDemoMode(),
 *                          setWebsetsLastSync()
 *   - Demo.gs            → getWebsetsDemoItems(), getWebsetsDemoInfo()
 *   - Code.gs            → getApiKey()
 */

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

/** @const {string} Name of the function invoked by the time-driven trigger. */
var WEBSETS_TRIGGER_FUNCTION = 'websetsAutoRefreshSync';

/** @const {number} Auto-refresh interval in minutes. */
var WEBSETS_TRIGGER_INTERVAL = 15;

// ---------------------------------------------------------------------------
// 1. Manual sync
// ---------------------------------------------------------------------------

/**
 * Performs a manual sync triggered from the sidebar or menu.
 *
 * Steps:
 *   1. Determine data source (demo vs. live API).
 *   2. Validate configuration (API key + Webset ID for live mode).
 *   3. Fetch items.
 *   4. Write items to the sheet.
 *   5. Record the sync timestamp.
 *   6. Show a result toast.
 *
 * @return {{ added: number, updated: number }|{ added: number, updated: number, error: string }}
 */
function websetsSyncNow() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var demoMode = isWebsetsDemoMode();
    var items;

    if (demoMode) {
      // ----- Demo mode -----
      ss.toast('Syncing with demo data…', 'Websets Sync', 3);
      items = getWebsetsDemoItems();
    } else {
      // ----- Live mode — validate config first -----
      var apiKey = getApiKey();
      if (!apiKey) {
        SpreadsheetApp.getUi().alert(
          'Websets Sync',
          'No API key found. Please add your Exa API key in the sidebar settings.',
          SpreadsheetApp.getUi().ButtonSet.OK
        );
        return { added: 0, updated: 0, error: 'Missing API key.' };
      }

      var websetId = getWebsetsWebsetId();
      if (!websetId) {
        SpreadsheetApp.getUi().alert(
          'Websets Sync',
          'No Webset ID configured. Please enter your Webset ID in the sidebar settings.',
          SpreadsheetApp.getUi().ButtonSet.OK
        );
        return { added: 0, updated: 0, error: 'Missing Webset ID.' };
      }

      ss.toast('Syncing with Exa Websets API…', 'Websets Sync', 3);
      items = listAllWebsetItems(apiKey, websetId);
    }

    // Write items to the sheet.
    var result = syncItemsToSheet(items);

    // Record the sync timestamp.
    setWebsetsLastSync(new Date().toISOString());

    // Show result toast.
    ss.toast(
      'Sync complete: ' + result.added + ' added, ' + result.updated + ' updated.',
      'Websets Sync',
      5
    );

    return { added: result.added, updated: result.updated };
  } catch (e) {
    Logger.log('[WebsetsSync] websetsSyncNow error: ' + e.message);
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Sync failed: ' + e.message,
        'Websets Sync Error',
        8
      );
    } catch (_) {
      // Toast may fail if called outside an interactive context.
    }
    return { added: 0, updated: 0, error: String(e.message || e) };
  }
}

// ---------------------------------------------------------------------------
// 2. Verify connection
// ---------------------------------------------------------------------------

/**
 * Tests the connection to a Webset by calling the API.
 *
 * @param {string} apiKey   The Exa API key.
 * @param {string} websetId The Webset ID to test.
 * @return {{ valid: boolean, title: string }|{ valid: boolean, error: string }}
 */
function websetsVerifyConnection(apiKey, websetId) {
  try {
    if (!apiKey) {
      return { valid: false, error: 'API key is required.' };
    }
    if (!websetId) {
      return { valid: false, error: 'Webset ID is required.' };
    }

    // Demo mode — return mock info without hitting the API.
    if (apiKey === 'demo-key' || isWebsetsDemoMode()) {
      var demoInfo = getWebsetsDemoInfo();
      return { valid: true, title: demoInfo.title || 'Demo Webset' };
    }

    var webset = getWebset(apiKey, websetId);
    return { valid: true, title: webset.title || webset.id || websetId };
  } catch (e) {
    Logger.log('[WebsetsSync] websetsVerifyConnection error: ' + e.message);
    return { valid: false, error: String(e.message || e) };
  }
}


// ---------------------------------------------------------------------------
// 3. Enable auto-refresh
// ---------------------------------------------------------------------------

/**
 * Creates a time-driven trigger that runs websetsAutoRefreshSync every
 * WEBSETS_TRIGGER_INTERVAL minutes.  Any existing triggers for the same
 * function are removed first to avoid duplicates.
 *
 * @return {{ success: boolean }|{ success: boolean, error: string }}
 */
function websetsEnableAutoRefresh() {
  try {
    // Remove any existing triggers for the auto-refresh function.
    websetsRemoveTriggers_();

    // Create a new time-driven trigger.
    ScriptApp.newTrigger(WEBSETS_TRIGGER_FUNCTION)
      .timeBased()
      .everyMinutes(WEBSETS_TRIGGER_INTERVAL)
      .create();

    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Auto-refresh enabled — syncing every ' + WEBSETS_TRIGGER_INTERVAL + ' minutes.',
        'Websets Sync',
        5
      );
    } catch (_) {
      // Toast may fail in non-interactive context.
    }

    return { success: true };
  } catch (e) {
    Logger.log('[WebsetsSync] websetsEnableAutoRefresh error: ' + e.message);
    return { success: false, error: String(e.message || e) };
  }
}

// ---------------------------------------------------------------------------
// 4. Disable auto-refresh
// ---------------------------------------------------------------------------

/**
 * Removes all time-driven triggers for websetsAutoRefreshSync.
 *
 * @return {{ success: boolean }}
 */
function websetsDisableAutoRefresh() {
  try {
    websetsRemoveTriggers_();

    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Auto-refresh disabled.',
        'Websets Sync',
        5
      );
    } catch (_) {
      // Toast may fail in non-interactive context.
    }

    return { success: true };
  } catch (e) {
    Logger.log('[WebsetsSync] websetsDisableAutoRefresh error: ' + e.message);
    return { success: true }; // Treat cleanup errors as non-fatal.
  }
}

// ---------------------------------------------------------------------------
// 5. Trigger-driven sync (runs silently — NO UI calls)
// ---------------------------------------------------------------------------

/**
 * Called by the time-driven trigger.  Fetches items from the API (or demo
 * source) and syncs them to the sheet.  Runs without any UI interactions
 * (no toast, no alert) because trigger executions don't have UI access.
 *
 * If the required configuration is missing, the function returns silently
 * so that the trigger doesn't throw an error every interval.
 */
function websetsAutoRefreshSync() {
  try {
    var demoMode = isWebsetsDemoMode();
    var items;

    if (demoMode) {
      items = getWebsetsDemoItems();
    } else {
      var apiKey = getApiKey();
      var websetId = getWebsetsWebsetId();

      // No config → exit silently.
      if (!apiKey || !websetId) {
        Logger.log('[WebsetsSync] Auto-refresh skipped: missing API key or Webset ID.');
        return;
      }

      items = listAllWebsetItems(apiKey, websetId);
    }

    var result = syncItemsToSheet(items);

    // Record the sync timestamp.
    setWebsetsLastSync(new Date().toISOString());

    Logger.log(
      '[WebsetsSync] Auto-refresh complete: ' +
        result.added + ' added, ' +
        result.updated + ' updated.'
    );
  } catch (e) {
    Logger.log('[WebsetsSync] Auto-refresh error: ' + e.message);
  }
}

// ---------------------------------------------------------------------------
// 6. Auto-refresh status
// ---------------------------------------------------------------------------

/**
 * Checks whether a time-driven trigger for websetsAutoRefreshSync currently
 * exists.
 *
 * @return {{ enabled: boolean }}
 */
function websetsGetAutoRefreshStatus() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === WEBSETS_TRIGGER_FUNCTION) {
        return { enabled: true };
      }
    }
    return { enabled: false };
  } catch (e) {
    Logger.log('[WebsetsSync] websetsGetAutoRefreshStatus error: ' + e.message);
    return { enabled: false };
  }
}

// ---------------------------------------------------------------------------
// Private helpers
// ---------------------------------------------------------------------------

/**
 * Removes all project triggers that target WEBSETS_TRIGGER_FUNCTION.
 *
 * @private
 */
function websetsRemoveTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === WEBSETS_TRIGGER_FUNCTION) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
