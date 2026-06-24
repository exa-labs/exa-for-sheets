/**
 * Config.gs — Websets-specific configuration storage.
 *
 * Uses ScriptProperties (shared across all users of the same script/sheet)
 * for Webset ID, demo-mode flag, and last-sync timestamp.
 *
 * The API key is intentionally NOT managed here — it lives in
 * UserProperties via the existing Code.gs getApiKey() / saveApiKey().
 *
 * Every property key is prefixed with WEBSETS_ to avoid collisions
 * with the original Exa search code.
 */

// ---------------------------------------------------------------------------
// Property-key constants
// ---------------------------------------------------------------------------

var WEBSETS_ID_KEY        = 'WEBSETS_WEBSET_ID';
var WEBSETS_DEMO_KEY      = 'WEBSETS_DEMO_MODE';
var WEBSETS_LAST_SYNC_KEY = 'WEBSETS_LAST_SYNC';

// ---------------------------------------------------------------------------
// Webset ID
// ---------------------------------------------------------------------------

/**
 * Save a Webset ID to ScriptProperties.
 *
 * @param {string} id  The Webset ID to persist.
 */
function saveWebsetsWebsetId(id) {
  try {
    PropertiesService.getDocumentProperties().setProperty(WEBSETS_ID_KEY, id);
    CacheService.getDocumentCache().put(WEBSETS_ID_KEY, id, 21600);
  } catch (e) {
    try {
      CacheService.getDocumentCache().put(WEBSETS_ID_KEY, id, 21600);
    } catch (e2) {
      Logger.log('saveWebsetsWebsetId error: ' + e);
      throw new Error('Storage permission denied (multiple Google accounts bug). Try an Incognito window.');
    }
  }
}

/**
 * Retrieve the stored Webset ID.
 *
 * @return {string|null}  The Webset ID, or null if not yet configured.
 */
function getWebsetsWebsetId() {
  try {
    var val = PropertiesService.getDocumentProperties().getProperty(WEBSETS_ID_KEY);
    if (!val) val = CacheService.getDocumentCache().get(WEBSETS_ID_KEY);
    return val;
  } catch (e) {
    try {
      return CacheService.getDocumentCache().get(WEBSETS_ID_KEY);
    } catch (e2) {
      Logger.log('getWebsetsWebsetId error: ' + e);
      return null;
    }
  }
}

// ---------------------------------------------------------------------------
// Demo mode
// ---------------------------------------------------------------------------

/**
 * Enable or disable demo mode.
 *
 * The boolean is stored as the string "true" or "false" because
 * PropertiesService only supports string values.
 *
 * @param {boolean} enabled  Whether demo mode should be on.
 */
function setWebsetsDemoMode(enabled) {
  var strVal = String(!!enabled);
  try {
    PropertiesService.getDocumentProperties().setProperty(WEBSETS_DEMO_KEY, strVal);
    CacheService.getDocumentCache().put(WEBSETS_DEMO_KEY, strVal, 21600);
  } catch (e) {
    try {
      CacheService.getDocumentCache().put(WEBSETS_DEMO_KEY, strVal, 21600);
    } catch (e2) {
      Logger.log('setWebsetsDemoMode error: ' + e);
      throw new Error('Storage permission denied (multiple Google accounts bug). Try an Incognito window.');
    }
  }
}

/**
 * Check whether demo mode is currently enabled.
 *
 * @return {boolean}  true if demo mode is on; false otherwise (including
 *                    when the property has never been set).
 */
function isWebsetsDemoMode() {
  try {
    var value = PropertiesService.getDocumentProperties().getProperty(WEBSETS_DEMO_KEY);
    if (!value) value = CacheService.getDocumentCache().get(WEBSETS_DEMO_KEY);
    return value === 'true';
  } catch (e) {
    try {
      var value = CacheService.getDocumentCache().get(WEBSETS_DEMO_KEY);
      return value === 'true';
    } catch (e2) {
      Logger.log('isWebsetsDemoMode error: ' + e);
      return false;
    }
  }
}

// ---------------------------------------------------------------------------
// Last sync timestamp
// ---------------------------------------------------------------------------

/**
 * Record the last successful sync time.
 *
 * @param {string} timestamp  An ISO-8601 timestamp string
 *                            (e.g. new Date().toISOString()).
 */
function setWebsetsLastSync(timestamp) {
  try {
    PropertiesService.getDocumentProperties().setProperty(WEBSETS_LAST_SYNC_KEY, timestamp);
    CacheService.getDocumentCache().put(WEBSETS_LAST_SYNC_KEY, timestamp, 21600);
  } catch (e) {
    try {
      CacheService.getDocumentCache().put(WEBSETS_LAST_SYNC_KEY, timestamp, 21600);
    } catch (e2) {
      Logger.log('setWebsetsLastSync error: ' + e);
    }
  }
}

/**
 * Retrieve the last recorded sync time.
 *
 * @return {string|null}  ISO-8601 timestamp, or null if no sync has occurred.
 */
function getWebsetsLastSync() {
  try {
    var val = PropertiesService.getDocumentProperties().getProperty(WEBSETS_LAST_SYNC_KEY);
    if (!val) val = CacheService.getDocumentCache().get(WEBSETS_LAST_SYNC_KEY);
    return val;
  } catch (e) {
    try {
      return CacheService.getDocumentCache().get(WEBSETS_LAST_SYNC_KEY);
    } catch (e2) {
      Logger.log('getWebsetsLastSync error: ' + e);
      return null;
    }
  }
}

// ---------------------------------------------------------------------------
// Bulk helpers
// ---------------------------------------------------------------------------

/**
 * Delete every WEBSETS_ property from ScriptProperties.
 *
 * This iterates all script-level properties and removes any whose key
 * starts with "WEBSETS_", leaving the rest (e.g. original Exa keys) intact.
 */
function clearWebsetsConfig() {
  try {
    var store = PropertiesService.getDocumentProperties();
    var all = store.getProperties();
    Object.keys(all).forEach(function (key) {
      if (key.indexOf('WEBSETS_') === 0) store.deleteProperty(key);
    });
  } catch (e) {}

  try {
    var cache = CacheService.getDocumentCache();
    cache.remove(WEBSETS_ID_KEY);
    cache.remove(WEBSETS_DEMO_KEY);
    cache.remove(WEBSETS_LAST_SYNC_KEY);
  } catch (e2) {}
}

/**
 * Return a snapshot of all Websets configuration — handy for populating
 * the sidebar UI in a single server call.
 *
 * @return {{websetId: string|null,
 *           demoMode: boolean,
 *           lastSync: string|null,
 *           hasApiKey: boolean}}
 */
function getWebsetsConfig() {
  var hasKey = false;
  try {
    hasKey = !!getApiKey(); // getApiKey() lives in Code.gs
  } catch (e) {
    Logger.log('getWebsetsConfig API key error: ' + e);
  }

  return {
    websetId:  getWebsetsWebsetId(),
    demoMode:  isWebsetsDemoMode(),
    lastSync:  getWebsetsLastSync(),
    hasApiKey: hasKey
  };
}

/**
 * Save the Webset ID from the sidebar settings form.
 *
 * Only the Webset ID is persisted here; the API key is saved separately
 * by the existing Code.gs saveApiKey() function.
 *
 * @param {string} websetId  The Webset ID to store.
 */
function saveWebsetsConfig(websetId) {
  saveWebsetsWebsetId(websetId);
}
