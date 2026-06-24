/**
 * WebsetsClient.gs
 *
 * Thin wrapper around UrlFetchApp for the Exa Websets REST API.
 * All functions are global (plain Apps Script style, no modules).
 *
 * Depends on:
 *   - getApiKey()  (defined in Code.gs – reads from UserProperties)
 *
 * API docs base URL: https://api.exa.ai/websets/v0
 */

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

/** @const {string} Base URL for the Exa Websets API (v0). */
var WEBSETS_BASE_URL = 'https://api.exa.ai/websets/v0';

/** @const {number} Maximum number of pagination loops to prevent runaways. */
var WEBSETS_MAX_PAGES = 100;

// ---------------------------------------------------------------------------
// Private Helpers (underscore suffix = file-private by convention)
// ---------------------------------------------------------------------------

/**
 * Builds the common fetch options object used by every Websets API request.
 *
 * @param {string} apiKey - The Exa API key.
 * @return {Object} A UrlFetchApp options object with auth and content headers.
 * @private
 */
function websets_fetchOptions_(apiKey) {
  return {
    headers: {
      'x-api-key': apiKey,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
}

/**
 * Maps an HTTP error response to a concise, human-friendly error message
 * and throws an Error so callers can handle it uniformly.
 *
 * @param {HTTPResponse} response - The UrlFetchApp response object.
 * @param {string} context - Short description of what we were doing
 *     (e.g. "fetching webset", "listing items page 3").
 * @throws {Error} Always throws with a descriptive message.
 * @private
 */
function websets_friendlyError_(response, context) {
  var code = response.getResponseCode();
  var body = response.getContentText();
  var detail = '';

  // Try to extract the API's own error message for extra context.
  try {
    var parsed = JSON.parse(body);
    detail = parsed.error || parsed.message || '';
  } catch (_) {
    // Body wasn't JSON – that's fine, we have the status code.
  }

  var message;
  switch (code) {
    case 401:
      message = 'Invalid or unauthorized API key. Websets requires a Pro plan.';
      break;
    case 403:
      message = 'Access denied. Your plan may not include Websets.';
      break;
    case 404:
      message = 'Webset not found. Check the ID.';
      break;
    case 429:
      message = 'Rate limit hit. Wait and try again.';
      break;
    default:
      message = 'Websets API error (HTTP ' + code + ')' +
                (detail ? ': ' + detail : '.');
  }

  throw new Error('[Websets] ' + context + ' — ' + message);
}

// ---------------------------------------------------------------------------
// Public Functions
// ---------------------------------------------------------------------------

/**
 * Fetches a single webset by its ID.
 *
 * GET /websets/{websetId}
 *
 * @param {string} apiKey   - The Exa API key.
 * @param {string} websetId - The ID of the webset to retrieve.
 * @return {Object} Parsed JSON object representing the webset
 *     (contains id, title, status, and other metadata).
 * @throws {Error} If the HTTP response is not 200.
 */
function getWebset(apiKey, websetId) {
  var url = WEBSETS_BASE_URL + '/websets/' + encodeURIComponent(websetId);
  var options = websets_fetchOptions_(apiKey);
  options.method = 'get';

  var response = UrlFetchApp.fetch(url, options);

  if (response.getResponseCode() !== 200) {
    websets_friendlyError_(response, 'fetching webset "' + websetId + '"');
  }

  return JSON.parse(response.getContentText());
}

/**
 * Fetches ALL items for a given webset, automatically paginating through
 * every page of results.
 *
 * GET /websets/{websetId}/items?cursor=…
 *
 * The Exa API may return either of two response shapes:
 *   Primary:  { data: [...], pagination: { next: "cursor" } }
 *   Fallback: { items: [...], cursor: "cursor" }
 * This function handles both transparently.
 *
 * @param {string} apiKey   - The Exa API key.
 * @param {string} websetId - The webset whose items we want.
 * @return {Object[]} Flat array of all item objects. Each item typically has:
 *     id, url, name/title, verifiedAt, status, enrichments, entity.
 * @throws {Error} If any page request returns a non-200 status.
 */
function listAllWebsetItems(apiKey, websetId) {
  var allItems = [];
  var cursor = null;
  var page = 0;

  do {
    page++;

    // Safety valve – never loop more than WEBSETS_MAX_PAGES times.
    if (page > WEBSETS_MAX_PAGES) {
      Logger.log('[Websets] Hit safety limit of ' + WEBSETS_MAX_PAGES +
                 ' pages while fetching items for webset "' + websetId + '".');
      break;
    }

    // Build the URL, appending the cursor when we have one.
    var url = WEBSETS_BASE_URL + '/websets/' + encodeURIComponent(websetId) + '/items';
    if (cursor) {
      url += '?cursor=' + encodeURIComponent(cursor);
    }

    var options = websets_fetchOptions_(apiKey);
    options.method = 'get';

    var response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() !== 200) {
      websets_friendlyError_(response, 'listing items (page ' + page + ') for webset "' + websetId + '"');
    }

    var body = JSON.parse(response.getContentText());

    // ------------------------------------------------------------------
    // Handle both possible API response shapes.
    // ------------------------------------------------------------------

    // Items array: prefer body.data, fall back to body.items.
    var pageItems = body.data || body.items || [];
    allItems = allItems.concat(pageItems);

    // Next-page cursor: prefer body.pagination.next, fall back to body.cursor.
    var nextCursor = null;
    if (body.pagination && body.pagination.next) {
      nextCursor = body.pagination.next;
    } else if (body.cursor) {
      nextCursor = body.cursor;
    }

    cursor = nextCursor;

    Logger.log('[Websets] Page ' + page + ': received ' + pageItems.length +
               ' items (total so far: ' + allItems.length + ')');

  } while (cursor);

  Logger.log('[Websets] Finished fetching items for webset "' + websetId +
             '". Total items: ' + allItems.length + ' across ' + page + ' page(s).');

  return allItems;
}
