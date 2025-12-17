function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Exa AI')
    .addItem('Open Sidebar', 'showSidebar')
    .addSeparator()
    .addItem('About/Help', 'showAbout')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Exa AI');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showAbout() {
  var ui = SpreadsheetApp.getUi();
  var message = 'Exa AI for Google Sheets\n\n' +
                'Version: 1.0.0\n\n' +
                'This add-on provides powerful AI-driven search and analysis capabilities using the Exa API.\n\n' +
                'Key Features:\n' +
                '- EXA_ANSWER: Generate AI answers from web searches\n' +
                '- EXA_SEARCH: Perform web searches\n' +
                '- EXA_CONTENTS: Extract content from URLs\n' +
                '- EXA_FINDSIMILAR: Find similar web pages\n\n' +
                'For detailed documentation and support, open the sidebar and navigate to the Docs tab.\n\n' +
                'Visit https://exa.ai for more information about the Exa API.\n\n' +
                'Visit https://github.com/exa-labs/exa-sheets the source code.';
  
  ui.alert('About Exa AI', message, ui.ButtonSet.OK);
}

/**
 * Retrieves all stored API keys
 * @return {Object} Object containing all saved API keys and metadata
 */
function getAllApiKeys() {
  const keysJson = PropertiesService.getUserProperties().getProperty('EXA_API_KEYS');
  if (!keysJson) {
    return { keys: {}, activeKeyName: null };
  }
  
  try {
    return JSON.parse(keysJson);
  } catch (e) {
    // If there's an error parsing the JSON, return empty structure
    return { keys: {}, activeKeyName: null };
  }
}

/**
 * Saves API keys data to UserProperties
 * @param {Object} keysData Object containing all keys and the active key name
 */
function saveAllApiKeys(keysData) {
  PropertiesService.getUserProperties().setProperty('EXA_API_KEYS', JSON.stringify(keysData));
}

/**
 * Helper function to create structured error responses
 * @param {string} code Error code for categorization
 * @param {string} message User-friendly error message
 * @param {string} correlationId Unique ID for tracking this request in logs
 * @param {Object} details Optional additional error details
 * @return {Object} Structured error response
 */
function fail(code, message, correlationId, details) {
  return { 
    success: false, 
    code: code,
    message: message,
    correlationId: correlationId,
    details: details || null
  };
}

/**
 * Saves a new API key (simplified version for the new UI)
 * @param {string} key The Exa API key to save
 * @param {string} reqId Optional request ID from client for correlation
 * @return {Object} Status object with success flag, message, and correlation ID
 */
function saveApiKey(key, reqId) {
  const correlationId = reqId || Utilities.getUuid();
  
  try {
    // Validate input
    if (!key || typeof key !== 'string' || !key.trim()) {
      console.error(JSON.stringify({
        correlationId: correlationId,
        where: 'saveApiKey',
        code: 'VALIDATION_EMPTY_KEY',
        error: 'Empty or invalid API key provided'
      }));
      return fail('VALIDATION_EMPTY_KEY', 'API key is required.', correlationId);
    }
    
    // Use a default name since we're managing a single key
    const name = "default";
    
    // Get all existing keys
    const keysData = getAllApiKeys();
    
    // Add or update the key with metadata
    const now = new Date().toISOString();
    keysData.keys[name] = {
      key: key,  // Store the actual API key
      created: keysData.keys[name] ? keysData.keys[name].created : now, // Keep original created date if updating
      lastUsed: now,
      // First few and last few characters for display, rest is masked with more dots
      displayKey: `${key.substring(0, 4)}${'.'.repeat(15)}${key.substring(key.length - 4)}`
    };
    
    // Set as active key
    keysData.activeKeyName = name;
    
    // Save back to properties with error handling
    try {
      saveAllApiKeys(keysData);
    } catch (e) {
      const errorMsg = String(e);
      const errorStack = e.stack || '';
      
      // Detect specific storage errors
      let code = 'STORAGE_WRITE_FAILED';
      let userMessage = 'Failed to save API key. Please try again.';
      
      if (errorMsg.includes('Service invoked too many times')) {
        code = 'STORAGE_RATE_LIMIT';
        userMessage = 'Too many requests. Please wait a moment and try again.';
      } else if (errorMsg.includes('exceeded maximum size')) {
        code = 'STORAGE_SIZE_EXCEEDED';
        userMessage = 'Storage limit exceeded. Please contact support.';
      }
      
      console.error(JSON.stringify({
        correlationId: correlationId,
        where: 'saveApiKey',
        code: code,
        error: errorMsg,
        stack: errorStack
      }));
      
      return fail(code, userMessage, correlationId);
    }
    
    return { 
      success: true, 
      message: 'API key saved successfully.',
      correlationId: correlationId
    };
    
  } catch (e) {
    // Catch any unexpected errors
    const errorMsg = String(e);
    const errorStack = e.stack || '';
    
    console.error(JSON.stringify({
      correlationId: correlationId,
      where: 'saveApiKey',
      code: 'INTERNAL',
      error: errorMsg,
      stack: errorStack
    }));
    
    return fail('INTERNAL', 'Unexpected error while saving API key.', correlationId);
  }
}

/**
 * Deletes an API key by name
 * @param {string} name The name of the key to delete
 * @return {Object} Status object with success flag and message
 */
function deleteApiKey(name) {
  const keysData = getAllApiKeys();
  
  if (!keysData.keys[name]) {
    return { 
      success: false, 
      message: `Key "${name}" not found.`
    };
  }
  
  // Delete the key
  delete keysData.keys[name];
  
  // If we deleted the active key, set a new active key or clear it
  if (keysData.activeKeyName === name) {
    const remainingKeys = Object.keys(keysData.keys);
    keysData.activeKeyName = remainingKeys.length > 0 ? remainingKeys[0] : null;
  }
  
  // Save the changes
  saveAllApiKeys(keysData);
  
  return { 
    success: true, 
    message: `Key "${name}" deleted successfully.`
  };
}

/**
 * Sets the active API key by name
 * @param {string} name The name of the key to set as active
 * @return {Object} Status object with success flag and message
 */
function setActiveApiKey(name) {
  const keysData = getAllApiKeys();
  
  if (!keysData.keys[name]) {
    return { 
      success: false, 
      message: `Key "${name}" not found.`
    };
  }
  
  // Set the active key
  keysData.activeKeyName = name;
  
  // Update last used timestamp
  keysData.keys[name].lastUsed = new Date().toISOString();
  
  // Save the changes
  saveAllApiKeys(keysData);
  
  return { 
    success: true, 
    message: `Key "${name}" is now active.`
  };
}

/**
 * Gets the currently active API key value for use in API calls
 * @return {string|null} The active API key or null if no key is set
 */
function getApiKey() {
  const keysData = getAllApiKeys();
  
  if (!keysData.activeKeyName || !keysData.keys[keysData.activeKeyName]) {
    return null;
  }
  
  // Update last used timestamp
  keysData.keys[keysData.activeKeyName].lastUsed = new Date().toISOString();
  saveAllApiKeys(keysData);
  
  return keysData.keys[keysData.activeKeyName].key;
}

/**
 * Gets information about the active key for display in UI
 * @return {Object} Object with active key info or null if no active key
 */
function getActiveKeyInfo() {
  const keysData = getAllApiKeys();
  
  if (!keysData.activeKeyName || !keysData.keys[keysData.activeKeyName]) {
    return null;
  }
  
  const activeKey = keysData.keys[keysData.activeKeyName];
  return {
    name: keysData.activeKeyName,
    displayKey: activeKey.displayKey,
    created: activeKey.created,
    lastUsed: activeKey.lastUsed
  };
}

/**
 * Helper function that formats API keys data for display in the UI
 * @return {Object} Object with activeKey and keys properties
 */
function getAllApiKeysForUI() {
  const keysData = getAllApiKeys();
  const result = {
    keys: {},
    activeKey: null
  };
  
  // Process each key to create the UI representation
  if (keysData && keysData.keys) {
    Object.entries(keysData.keys).forEach(([name, keyData]) => {
      result.keys[name] = {
        displayKey: keyData.displayKey,
        created: keyData.created,
        lastUsed: keyData.lastUsed
      };
    });
  }
  
  // Set the active key info
  if (keysData.activeKeyName && keysData.keys[keysData.activeKeyName]) {
    const activeKey = keysData.keys[keysData.activeKeyName];
    result.activeKey = {
      name: keysData.activeKeyName,
      displayKey: activeKey.displayKey,
      created: activeKey.created,
      lastUsed: activeKey.lastUsed
    };
  }
  
  return result;
}


/**
 * Simplified remove API key function for the new UI
 * @return {Object} Status object with success flag and message
 */
function removeApiKey() {
  // Clear all keys
  PropertiesService.getUserProperties().deleteProperty('EXA_API_KEYS');
  
  return { 
    success: true, 
    message: 'API key removed successfully.'
  };
}

/**
 * Get API key info for the simplified UI
 * @return {Object|null} Object with displayKey and created date, or null if no key
 */
function getApiKeyForUI() {
  const keysData = getAllApiKeys();
  
  if (!keysData.activeKeyName || !keysData.keys[keysData.activeKeyName]) {
    return null;
  }
  
  const activeKey = keysData.keys[keysData.activeKeyName];
  return {
    displayKey: activeKey.displayKey,
    created: activeKey.created
  };
}

/**
 * Ensures the user has authorized the add-on by touching PropertiesService.
 * This should be called on a user gesture (button click) to ensure the OAuth
 * consent flow can display properly without being blocked by pop-up blockers.
 * @return {boolean} Always returns true if authorization succeeds
 */
function ensureAuthorized() {
  PropertiesService.getUserProperties().getProperty('EXA_API_KEYS');
  return true;
}

/**
 * Queries the Exa /answer endpoint to provide an AI-generated answer based on search results.
 * Allows adding prefix/suffix text and optionally includes source citations.
 * By default, extracts and returns only the core answer text before any inline citations like " ([Source](URL)...)".
 *
 * @param {string} prompt The main question or prompt to send to Exa. Can be a cell reference.
 * @param {string} [prefix=""] Optional. Text to add before the main prompt.
 * @param {string} [suffix=""] Optional. Text to add after the main prompt.
 * @param {boolean} [includeCitations=FALSE] Optional. If TRUE, returns the full answer string as received from the API (including inline citations) AND appends any additional citations from the API's 'citations' array. Defaults to FALSE (extracts core answer only).
 * @return {string} The core answer, the full answer with citations, or an error message.
 * @customfunction
 */
function EXA_ANSWER(prompt, prefix, suffix, includeCitations) {
  const apiKey = getApiKey();
  if (!apiKey) return "No API key set. Please set your API key in the Exa AI sidebar.";

  // --- Parameter Validation and Processing ---
  if (!prompt || typeof prompt !== 'string' || prompt.trim() === "") {
    return "Please provide a valid prompt/question.";
  }

  const finalPrompt = `${prefix || ''} ${prompt} ${suffix || ''}`.trim();
  // Explicitly check if includeCitations is exactly TRUE.
  const shouldShowFullAnswerWithCitations = includeCitations === true;

  // --- API Call ---
  try {
    const response = UrlFetchApp.fetch("https://api.exa.ai/answer", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ query: finalPrompt }),
      headers: { "x-api-key": apiKey, "x-exa-integration": "exa-for-sheets", "User-Agent": "exa-for-sheets 1.0" },
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    // --- Response Handling ---
    if (responseCode === 200) {
      const result = JSON.parse(responseBody);

      if (result && typeof result.answer === 'string') {
        const fullAnswerFromApi = result.answer;
        let finalOutput = fullAnswerFromApi; // Default to the full answer

        if (!shouldShowFullAnswerWithCitations) {
          // --- Default Behavior: Strip Inline Citations ---
          // Remove inline citation patterns like " ([Source](URL))" or " ([Source](URL), [Source2](URL2))"
          // This regex matches citations in the format: space followed by opening paren, bracket, content, closing bracket, paren
          // It handles multiple consecutive citations separated by commas
          finalOutput = fullAnswerFromApi.replace(/\s+\(\[([^\]]+)\]\([^\)]+\)(?:,\s*\[([^\]]+)\]\([^\)]+\))*\)/g, '').trim();

          // If no inline citations were found, use the full answer
          if (finalOutput === '') {
            finalOutput = fullAnswerFromApi.trim();
          }

        } else {
          // --- Include Citations Behavior ---
          // Use the full answer from API, and append from the citations array if present
           finalOutput = fullAnswerFromApi.trim(); // Start with the full API answer

           if (result.citations && Array.isArray(result.citations) && result.citations.length > 0) {
              const formattedCitations = result.citations.map(citation => {
                  const title = citation.title || 'Source';
                  const url = citation.url;
                  return url ? `([${title}](${url}))` : null;
              })
              .filter(c => c !== null)
              .join(', ');

              if (formattedCitations.length > 0) {
                  const separator = (finalOutput.slice(-1) !== ' ' && finalOutput.slice(-1) !== '\n') ? ' ' : '';
                  finalOutput += separator + formattedCitations; // Append the extra citations
              }
           }
           // If includeCitations is true but no citations array, finalOutput remains the fullAnswerFromApi
        }

        return finalOutput.trim(); // Return the processed output

      } else {
        return "API returned a valid response, but no 'answer' field (string) was found.";
      }
    } else if (responseCode === 401) {
      return "API Error: Invalid API Key.";
    } else { // Handle other errors
      let errorMessage = `API Error: Status ${responseCode}.`;
      try {
        const errorResult = JSON.parse(responseBody);
        errorMessage += ` Message: ${errorResult.error || responseBody}`;
      } catch (e) {
        errorMessage += ` Response: ${responseBody}`;
      }
      return errorMessage;
    }
  } catch (e) {
    Logger.log(`EXA_ANSWER Error: ${e} for prompt: ${finalPrompt}`);
    return `Script Error: ${e.message}`;
  }
}

/**
 * Retrieves the text content of a given URL using the Exa /contents endpoint.
 *
 * @param {string} url The full URL (including http/https) to fetch content from.
 * @return {string} The main text content of the URL or an error message.
 * @customfunction
 */
function EXA_CONTENTS(url) {
  const apiKey = getApiKey();
  if (!apiKey) return "No API key set. Please set your API key in the Exa AI sidebar.";

  // Basic URL validation
  if (!url || typeof url !== 'string' || !url.startsWith('http')) {
      return "Please provide a valid URL starting with http or https.";
  }

  try {
    const response = UrlFetchApp.fetch("https://api.exa.ai/contents", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ urls: [url] }), // Exa's /contents endpoint expects an array of URLs
      headers: { "x-api-key": apiKey, "x-exa-integration": "exa-for-sheets", "User-Agent": "exa-for-sheets 1.0" },
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
        const result = JSON.parse(responseBody);
        // The response structure for /contents typically wraps results in a 'results' array
        const contentData = result.results && result.results[0];
        if (contentData) {
            // Prioritize text content based on Exa's common response structure
            return (contentData.text || contentData.highlights || "No relevant content found in response.").trim();
        } else {
            return "API returned successfully, but no content data found for this URL.";
        }
    } else if (responseCode === 401) {
        return "API Error: Invalid API Key. Please check your key in the menu.";
    } else {
        // Try to parse error message from API if possible, otherwise return generic error
        let errorMessage = `API Error: Received status code ${responseCode}.`;
        try {
            const errorResult = JSON.parse(responseBody);
            errorMessage += ` Message: ${errorResult.error || responseBody}`;
        } catch (e) {
            errorMessage += ` Response: ${responseBody}`;
        }
        return errorMessage;
    }
  } catch (e) {
    return `Script Error: ${e.message}`;
  }
}

/**
 * Finds URLs similar to the input URL using the Exa /findSimilar endpoint, with optional filters.
 * Returns a vertical list of similar URLs.
 *
 * @param {string} url The URL to find similar links for (must include http/https).
 * @param {number} [numResults=1] Optional. The maximum number of similar URLs to return (1-10). Defaults to 1.
 * @param {string} [includeDomainsStr=""] Optional. Comma-separated list of domains to restrict results to (e.g., "example.com,anotherexample.org").
 * @param {string} [excludeDomainsStr=""] Optional. Comma-separated list of domains to exclude from results (e.g., "exclude.net,badsite.co").
 * @param {string} [includeTextStr=""] Optional. A phrase that MUST be present in the content of result pages.
 * @param {string} [excludeTextStr=""] Optional. A phrase that MUST NOT be present in the content of result pages.
 * @return {string[][]} A vertical array of similar URLs, or a single cell error message.
 * @customfunction
 */
function EXA_FINDSIMILAR(url, numResults, includeDomainsStr, excludeDomainsStr, includeTextStr, excludeTextStr) {
  const apiKey = getApiKey();
  if (!apiKey) return [["No API key set. Please set your API key in the Exa AI sidebar."]];

  // --- Parameter Validation and Processing ---
  if (!url || typeof url !== 'string' || !url.startsWith('http')) {
    return [["Please provide a valid URL starting with http or https."]];
  }

  // Validate and set numResults (sensible default and limits)
  const count = (typeof numResults === 'number' && numResults >= 1 && numResults <= 10)
                ? Math.floor(numResults)
                : 1; // Default to 1 if invalid, NaN, or outside 1-10 range

  // Process domain lists (comma-separated string to array)
  const processDomains = (domainStr) => {
    if (typeof domainStr === 'string' && domainStr.trim() !== '') {
      return domainStr.split(',').map(d => d.trim()).filter(d => d.length > 0);
    }
    return null; // Return null if empty or not a string
  };

  const includeDomains = processDomains(includeDomainsStr);
  const excludeDomains = processDomains(excludeDomainsStr);

  // Process text filters (use the string directly if provided)
  const includeText = (typeof includeTextStr === 'string' && includeTextStr.trim() !== '') ? [includeTextStr.trim()] : null;
  const excludeText = (typeof excludeTextStr === 'string' && excludeTextStr.trim() !== '') ? [excludeTextStr.trim()] : null;
  // Note: Exa API docs mention limit of 1 string, 5 words for text filters. We send as array[1].

  // --- Build API Payload ---
  const payload = {
    url: url,
    numResults: count,
    excludeSourceDomain: true // Good default to avoid getting the input URL back
  };

  if (includeDomains && includeDomains.length > 0) {
    payload.includeDomains = includeDomains;
  }
  if (excludeDomains && excludeDomains.length > 0) {
    payload.excludeDomains = excludeDomains;
  }
  if (includeText && includeText.length > 0) {
      // Ensure only one item is sent if API has that restriction
     payload.includeText = includeText.slice(0, 1);
  }
  if (excludeText && excludeText.length > 0) {
      // Ensure only one item is sent if API has that restriction
     payload.excludeText = excludeText.slice(0, 1);
  }

  // --- API Call and Response Handling ---
  try {
    const response = UrlFetchApp.fetch("https://api.exa.ai/findSimilar", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      headers: { "x-api-key": apiKey, "x-exa-integration": "exa-for-sheets", "User-Agent": "exa-for-sheets 1.0" },
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const result = JSON.parse(responseBody);
      if (result && result.results && result.results.length > 0) {
        return result.results.map(item => [item.url || "N/A"]); // Map URLs, provide fallback
      } else {
        return [["No similar URLs found matching the criteria."]];
      }
    } else if (responseCode === 401) {
      return [["API Error: Invalid API Key."]];
    } else if (responseCode === 400) {
        // Handle potential bad request errors (e.g., invalid filters)
        let errorMessage = `API Error (Bad Request): Status ${responseCode}.`;
        try {
            const errorResult = JSON.parse(responseBody);
            errorMessage += ` Message: ${errorResult.error || responseBody}`;
        } catch (e) {
            errorMessage += ` Response: ${responseBody}`;
        }
        return [[errorMessage]];
    }
    else { // Handle other errors
      let errorMessage = `API Error: Status ${responseCode}.`;
      try {
        const errorResult = JSON.parse(responseBody);
        errorMessage += ` Message: ${errorResult.error || responseBody}`;
      } catch (e) {
        errorMessage += ` Response: ${responseBody}`;
      }
      return [[errorMessage]];
    }
  } catch (e) {
    // Catch script execution errors (e.g., network issues)
    Logger.log(`EXA_FINDSIMILAR Error: ${e} for payload: ${JSON.stringify(payload)}`); // Log for debugging
    return [[`Script Error: ${e.message}`]];
  }
}

/**
 * Searches the web using the Exa /search endpoint based on a query.
 * Returns a vertical list of result URLs.
 *
 * @param {string} query The search query.
 * @param {number} [numResults=1] Optional. The maximum number of result URLs to return. Defaults to 1.
 * @param {string} [searchType="auto"] Optional. The type of search ('auto', 'neural', 'keyword'). Defaults to 'auto'.
 * @param {string} [prefix=""] Optional. Text to add before the main query.
 * @param {string} [suffix=""] Optional. Text to add after the main query.
 * @return {string[]} An array of result URLs or a single cell error message.
 * @customfunction
 */
function EXA_SEARCH(query, numResults, searchType, prefix, suffix) {
  const apiKey = getApiKey();
  if (!apiKey) return [["No API key set. Please set your API key in the Exa AI sidebar."]];

  if (!query || typeof query !== 'string' || query.trim() === "") {
    return [["Please provide a valid search query."]];
  }

  // Process the query with optional prefix and suffix
  const finalQuery = `${prefix || ''} ${query} ${suffix || ''}`.trim();

  const count = (typeof numResults === 'number' && numResults > 0 && numResults <= 10) ? Math.floor(numResults) : 1; // Default to 1, max 10 for free tier
  const type = (searchType && ['auto', 'neural', 'keyword'].includes(searchType)) ? searchType : 'auto'; // Default to 'auto'

  try {
    const response = UrlFetchApp.fetch("https://api.exa.ai/search", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        query: finalQuery,
        numResults: count,
        type: type,
        useAutoprompt: (type !== 'keyword') // Enable autoprompt for neural/auto by default
      }),
      headers: { "x-api-key": apiKey, "x-exa-integration": "exa-for-sheets", "User-Agent": "exa-for-sheets 1.0" },
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

     if (responseCode === 200) {
      const result = JSON.parse(responseBody);
      if (result && result.results && result.results.length > 0) {
        // Map results to a 2D array for vertical spill
        return result.results.map(item => [item.url]);
      } else {
        return [["API returned successfully, but no search results found."]];
      }
    } else if (responseCode === 401) {
      return [["❌ API Error: Invalid API Key. Please check your key in the menu."]];
    } else {
      let errorMessage = `API Error: Status ${responseCode}.`;
      try {
        const errorResult = JSON.parse(responseBody);
        errorMessage += ` Message: ${errorResult.error || responseBody}`;
      } catch (e) {
        errorMessage += ` Response: ${responseBody}`;
      }
      return [[errorMessage]];
    }
  } catch (e) {
    return [[`Script Error: ${e.message}`]];
  }
}

/**
 * Refreshes all selected cells containing Exa functions by forcing recalculation.
 * Processes all cells in parallel for optimal performance.
 * Properly handles array-returning functions by clearing spilled values.
 * 
 * @param {string} operation The operation to perform (always 'refresh')
 * @return {Object} Result object with success flag and message
 */
function processBatchOperation(operation) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const selection = sheet.getActiveRange();
    
    if (!selection) {
      return { 
        success: false, 
        message: 'No cells selected. Please select cells containing Exa functions.'
      };
    }
    
    // Get all formulas and filter for Exa functions
    const formulas = selection.getFormulas();
    const exaCells = [];
    
    formulas.forEach((row, rowIndex) => {
      row.forEach((formula, colIndex) => {
        if (formula && formula.toUpperCase().match(/^=EXA_/)) {
          exaCells.push({
            cell: selection.getCell(rowIndex + 1, colIndex + 1),
            formula: formula,
            row: selection.getRow() + rowIndex,
            col: selection.getColumn() + colIndex
          });
        }
      });
    });
    
    if (exaCells.length === 0) {
      const totalCells = formulas.flat().length;
      const cellText = totalCells === 1 ? 'cell' : 'cells';
      return { 
        success: false, 
        message: `No Exa functions found in the ${totalCells} selected ${cellText}.`
      };
    }
    
    // For each Exa function cell, clear potential spilled values
    exaCells.forEach(item => {
      // Clear the formula cell
      item.cell.setFormula('');
      
      // Clear potential spilled values below and to the right
      // Array formulas can spill vertically (for EXA_SEARCH, EXA_FINDSIMILAR)
      // We'll clear up to 100 rows below and 10 columns to the right to be safe
      const maxRow = Math.min(item.row + 100, sheet.getMaxRows());
      const maxCol = Math.min(item.col + 10, sheet.getMaxColumns());
      
      if (maxRow > item.row || maxCol > item.col) {
        const numRows = maxRow - item.row + 1;
        const numCols = maxCol - item.col + 1;
        const spillRange = sheet.getRange(item.row, item.col, numRows, numCols);
        
        // Only clear cells that don't have formulas (these are spilled values)
        const spillFormulas = spillRange.getFormulas();
        const spillValues = spillRange.getValues();
        
        spillFormulas.forEach((formulaRow, rowIdx) => {
          formulaRow.forEach((formula, colIdx) => {
            // Skip the first cell (it's the formula cell we already cleared)
            if (rowIdx === 0 && colIdx === 0) return;
            
            // If cell has no formula but has a value, it's likely a spilled value
            if (!formula && spillValues[rowIdx][colIdx] !== '') {
              sheet.getRange(item.row + rowIdx, item.col + colIdx).clear();
            }
          });
        });
      }
    });
    
    SpreadsheetApp.flush();
    
    // Restore all formulas at once
    exaCells.forEach(item => item.cell.setFormula(item.formula));
    SpreadsheetApp.flush();
    
    const cellText = exaCells.length === 1 ? 'cell' : 'cells';
    return { 
      success: true, 
      message: `Successfully refreshed ${exaCells.length} ${cellText}.`
    };
    
  } catch (e) {
    Logger.log(`Error in processBatchOperation: ${e}`);
    return { 
      success: false, 
      message: `Operation failed: ${e.message}`
    };
  }
}
