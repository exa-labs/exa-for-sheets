// Rate limiting configuration
var RATE_LIMIT_CONFIG = {
  maxRetries: 3,
  initialDelayMs: 1000,
  maxDelayMs: 10000
};

var EXA_AGENT_CONFIG = {
  endpoint: "https://api.exa.ai/agent/runs",
  betaHeader: "agent-2026-05-07",
  defaultRows: 50
};

var EXA_AGENT_VALID_EFFORTS = ['low', 'medium', 'high', 'xhigh', 'auto'];

/**
 * Makes an HTTP request with exponential backoff retry on rate limit (429) errors.
 * @param {string} url The URL to fetch
 * @param {Object} options UrlFetchApp options
 * @return {HTTPResponse} The response object
 */
function fetchWithRetry(url, options) {
  var lastError = null;
  var delay = RATE_LIMIT_CONFIG.initialDelayMs;

  for (var attempt = 0; attempt <= RATE_LIMIT_CONFIG.maxRetries; attempt++) {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();

    if (code !== 429) {
      return response;
    }

    // Rate limited - check if we should retry
    if (attempt >= RATE_LIMIT_CONFIG.maxRetries) {
      return response; // Return the 429 response after max retries
    }

    // Respect Retry-After when present, including HTTP-date values.
    var headers = response.getHeaders ? response.getHeaders() : {};
    var retryAfter = headers['Retry-After'] || headers['retry-after'];
    if (retryAfter) {
      delay = getRetryAfterMs(response);
    }

    // Cap the delay.
    delay = Math.min(delay, RATE_LIMIT_CONFIG.maxDelayMs);

    Logger.log('Rate limited (429). Retrying in ' + delay + 'ms (attempt ' + (attempt + 1) + '/' + RATE_LIMIT_CONFIG.maxRetries + ')');
    Utilities.sleep(delay);

    // Exponential backoff for next attempt
    delay = delay * 2;
  }

  return response;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Exa AI')
    .addItem('Open Sidebar', 'showSidebar')
    .addSeparator()
    .addItem('About/Help', 'showAbout')
    .addToUi();

  // Websets Sync — added by websets-sheets-sync fork
  SpreadsheetApp.getUi()
    .createMenu('Websets Sync')
    .addItem('Open Sidebar', 'showWebsetsSidebar')
    .addSeparator()
    .addItem('Sync Now', 'websetsSyncNow')
    .addSeparator()
    .addItem('Enable Auto-Refresh (15 min)', 'websetsEnableAutoRefresh')
    .addItem('Disable Auto-Refresh', 'websetsDisableAutoRefresh')
    .addSeparator()
    .addItem('Reset Websets Config', 'websetsClearConfig')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Exa AI');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Opens the Websets Sync configuration sidebar.
 * Added by websets-sheets-sync fork.
 */
function showWebsetsSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('WebsetsSidebar')
    .setTitle('Websets Sync');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Clears all Websets configuration (wrapper for menu item).
 * Added by websets-sheets-sync fork.
 */
function websetsClearConfig() {
  clearWebsetsConfig();
  SpreadsheetApp.getUi().alert('Websets Sync configuration has been reset.');
}

function showAbout() {
  var ui = SpreadsheetApp.getUi();
  var message = 'Exa AI for Google Sheets\n\n' +
                'Version: 2.0.5\n\n' +
                'Use Exa Agent to generate tables and fill blank cells.\n\n' +
                'Use the EXA formula inside a cell when you want one answer.\n\n' +
                'Open the sidebar to add your API key and start using Exa.\n\n' +
                'Learn more at https://exa.ai';

  ui.alert('About Exa AI', message, ui.ButtonSet.OK);
}

/**
 * Retrieves all stored API keys
 * @return {Object} Object containing all saved API keys and metadata
 */
function getAllApiKeys() {
  let keysJson = null;
  try {
    keysJson = PropertiesService.getUserProperties().getProperty('EXA_API_KEYS');
    if (!keysJson) keysJson = CacheService.getUserCache().get('EXA_API_KEYS');
  } catch (e) {
    try {
      keysJson = CacheService.getUserCache().get('EXA_API_KEYS');
    } catch (e2) {
      return { keys: {}, activeKeyName: null };
    }
  }
  
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
  var str = JSON.stringify(keysData);
  try {
    PropertiesService.getUserProperties().setProperty('EXA_API_KEYS', str);
    CacheService.getUserCache().put('EXA_API_KEYS', str, 21600); // Also cache it
  } catch (e) {
    try {
      CacheService.getUserCache().put('EXA_API_KEYS', str, 21600);
    } catch (e2) {
      Logger.log('saveAllApiKeys storage error: ' + e);
      throw new Error('Storage permission denied. If you have multiple Google accounts logged in, please use an Incognito window.');
    }
  }
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
  try {
    PropertiesService.getUserProperties().deleteProperty('EXA_API_KEYS');
  } catch (e) {
    Logger.log('removeApiKey storage error: ' + e);
  }

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
  try {
    PropertiesService.getUserProperties().getProperty('EXA_API_KEYS');
    return true;
  } catch (e) {
    try {
      CacheService.getUserCache().get('EXA_API_KEYS');
      return true;
    } catch (e2) {
      return false;
    }
  }
}

/**
 * Simple AI-powered data enrichment using Exa. This is the recommended function for most use cases.
 * Just describe what information you want about the data in the referenced cell.
 * Uses /search with outputSchema for structured text output.
 *
 * Examples:
 *   =EXA("Return only the company website URL", A1)
 *   =EXA("Return only the company headcount", A1)
 *   =EXA("Return only the CEO name", A1)
 *   =EXA("Return the Amazon rating of this product", A1)
 *
 * For advanced options (system prompt, output schema, citations), use EXA_ANSWER instead.
 *
 * @param {string} prompt What information you want (e.g., "Return only the company website URL").
 * @param {string} [context=""] Optional. Cell reference or text to enrich (e.g., company name in A1).
 * @return {string} The requested information or an error message.
 * @customfunction
 */
function EXA(prompt, context) {
  const apiKey = getApiKey();
  if (!apiKey) return "No API key set. Please set your API key in the Exa AI sidebar.";

  if (!prompt || typeof prompt !== 'string' || prompt.trim() === "") {
    return "Please provide a valid prompt/question.";
  }

  const query = context ? `${prompt}: ${context}` : prompt;

  const systemPrompt = 'Follow the user\'s formatting instructions exactly. ' +
    'Return only the requested information with no extra commentary. ' +
    'No citations, no markdown formatting, no brackets like [1][2].';

  const payload = {
    query: query,
    numResults: 10,
    type: 'auto',
    stream: false,
    systemPrompt: systemPrompt,
    outputSchema: {
      type: 'text',
      description: prompt
    },
    contents: {
      highlights: {
        maxCharacters: 4000
      }
    }
  };

  try {
    const response = fetchWithRetry("https://api.exa.ai/search", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      headers: { "x-api-key": apiKey, "x-exa-integration": "exa-for-sheets", "User-Agent": "exa-for-sheets 2.0" },
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const result = JSON.parse(responseBody);
      if (result.output && result.output.content) {
        // Strip any remaining inline citation markers like [1], [2][3], etc.
        return result.output.content.replace(/\s*\[\d+\](\[\d+\])*/g, '').trim();
      } else if (result.results && result.results.length > 0) {
        return result.results[0].url;
      }
      return "No results found.";
    } else if (responseCode === 401) {
      return "API Error: Invalid API Key. Please check your key in the Exa AI sidebar.";
    } else if (responseCode === 429) {
      return "API Error: Rate limit exceeded. Please wait a moment and try again.";
    } else {
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
    return `Script Error: ${e.message}`;
  }
}

/**
 * Queries the Exa /answer endpoint to provide an AI-generated answer based on search results.
 * Allows adding prefix/suffix text and optionally includes source citations.
 * By default, extracts and returns only the core answer text before any inline citations like " ([Source](URL)...)".
 *
 * For structured output, generate schemas at: https://dashboard.exa.ai/playground/answer
 *
 * @param {string} prompt The main question or prompt to send to Exa. Can be a cell reference.
 * @param {string} [prefix=""] Optional. Text to add before the main prompt.
 * @param {string} [suffix=""] Optional. Text to add after the main prompt.
 * @param {boolean} [includeCitations=FALSE] Optional. If TRUE, appends source citations. Defaults to FALSE.
 * @param {string} [systemPrompt=""] Optional. System instructions to control output format (e.g., "only return a number"). Uses chat completions endpoint.
 * @param {string} [outputSchema=""] Optional. JSON schema for structured output (e.g., '{"type":"object","properties":{"value":{"type":"number"}},"required":["value"]}').
 * @param {boolean} [returnRawJson=FALSE] Optional. If TRUE and outputSchema is provided, returns raw JSON instead of extracted value.
 * @param {string} [type=""] Optional. Search type: 'auto' (default), 'neural', 'fast', or 'deep'. Deep search provides more thorough results.
 * @return {string} The answer, or structured JSON if outputSchema is provided.
 * @customfunction
 */
function EXA_ANSWER(prompt, prefix, suffix, includeCitations, systemPrompt, outputSchema, returnRawJson, type) {
  const apiKey = getApiKey();
  if (!apiKey) return "No API key set. Please set your API key in the Exa AI sidebar.";

  // --- Parameter Validation and Processing ---
  if (!prompt || typeof prompt !== 'string' || prompt.trim() === "") {
    return "Please provide a valid prompt/question.";
  }

  const finalPrompt = `${prefix || ''} ${prompt} ${suffix || ''}`.trim();
  const shouldShowFullAnswerWithCitations = includeCitations === true;
  const hasSystemPrompt = typeof systemPrompt === 'string' && systemPrompt.trim() !== '';

  // Parse outputSchema if provided
  let parsedSchema = null;
  if (typeof outputSchema === 'string' && outputSchema.trim() !== '') {
    try {
      parsedSchema = JSON.parse(outputSchema);
    } catch (e) {
      return "Invalid outputSchema: must be valid JSON.";
    }
  }

  // Validate search type — defaults to 'deep' for richer results
  const validTypes = ['auto', 'neural', 'fast', 'deep'];
  const searchType = (typeof type === 'string' && validTypes.includes(type.toLowerCase()))
    ? type.toLowerCase()
    : 'deep';

  // --- API Call ---
  try {
    let response;
    const useChatCompletions = hasSystemPrompt || parsedSchema;

    if (useChatCompletions) {
      // Use chat completions endpoint for systemPrompt OR outputSchema (OpenAI-compatible format)
      const messages = [];
      if (hasSystemPrompt) {
        messages.push({ role: "system", content: systemPrompt.trim() });
      }
      messages.push({ role: "user", content: finalPrompt });

      const chatPayload = { model: "exa", messages: messages };
      const extraBody = {};
      if (parsedSchema) {
        extraBody.outputSchema = parsedSchema;
      }
      if (searchType) {
        extraBody.type = searchType;
      }
      if (Object.keys(extraBody).length > 0) {
        chatPayload.extraBody = extraBody;
      }

      response = fetchWithRetry("https://api.exa.ai/chat/completions", {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(chatPayload),
        headers: { "Authorization": `Bearer ${apiKey}`, "x-exa-integration": "exa-for-sheets", "User-Agent": "exa-for-sheets 2.0" },
        muteHttpExceptions: true
      });
    } else {
      // Use /answer endpoint (no systemPrompt, no outputSchema)
      const answerPayload = { query: finalPrompt };
      if (searchType) {
        answerPayload.type = searchType;
      }
      response = fetchWithRetry("https://api.exa.ai/answer", {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(answerPayload),
        headers: { "x-api-key": apiKey, "x-exa-integration": "exa-for-sheets", "User-Agent": "exa-for-sheets 2.0" },
        muteHttpExceptions: true
      });
    }

    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    // --- Response Handling ---
    if (responseCode === 200) {
      const result = JSON.parse(responseBody);

      let fullAnswerFromApi;
      let citations = [];

      if (useChatCompletions) {
        // Chat completions response format (for systemPrompt and/or outputSchema)
        // Response is always in choices[0].message.content format

        if (result.choices && result.choices[0] && result.choices[0].message) {
          const messageContent = result.choices[0].message.content;
          citations = result.choices[0].message.citations || [];

          if (parsedSchema) {
            // With outputSchema: content is a JSON string that needs to be parsed
            try {
              const answerObj = JSON.parse(messageContent);

              if (typeof answerObj === 'object' && answerObj !== null) {
                if (returnRawJson === true) {
                  fullAnswerFromApi = JSON.stringify(answerObj, null, 2);
                } else {
                  // Extract value: if single key, return just the value; otherwise return formatted JSON
                  const keys = Object.keys(answerObj);
                  if (keys.length === 1) {
                    fullAnswerFromApi = String(answerObj[keys[0]]);
                  } else {
                    fullAnswerFromApi = JSON.stringify(answerObj, null, 2);
                  }
                }
              } else {
                fullAnswerFromApi = messageContent;
              }
            } catch (parseError) {
              // If JSON parsing fails, return the content as-is
              fullAnswerFromApi = messageContent;
            }
          } else {
            // Without outputSchema (systemPrompt only): use content directly
            fullAnswerFromApi = messageContent;
          }
        } else {
          return "API returned a valid response, but no message content was found.";
        }
      } else {
        // /answer endpoint response format (no systemPrompt, no outputSchema)
        citations = result.citations || [];

        if (result && typeof result.answer === 'string') {
          fullAnswerFromApi = result.answer;
        } else {
          return "API returned a valid response, but no 'answer' field was found.";
        }
      }

      let finalOutput = fullAnswerFromApi;

      // Regex to match inline citations like " ([Source](URL))" or " ([Source](URL), [Source2](URL2))"
      const inlineCitationRegex = /\s*\(\[([^\]]+)\]\(([^\)]+)\)(?:,\s*\[([^\]]+)\]\(([^\)]+)\))*\)/g;

      // Always strip inline citations from the answer text for cleaner output
      const cleanAnswer = fullAnswerFromApi.replace(inlineCitationRegex, '').trim();

      if (!shouldShowFullAnswerWithCitations) {
        finalOutput = cleanAnswer || fullAnswerFromApi.trim();
      } else {
        finalOutput = cleanAnswer || fullAnswerFromApi.trim();

        const allCitations = [];

        if (Array.isArray(citations) && citations.length > 0) {
          citations.forEach(citation => {
            const title = citation.title || 'Source';
            const url = citation.url;
            if (url) {
              allCitations.push(`[${title}](${url})`);
            }
          });
        }

        if (allCitations.length > 0) {
          finalOutput += '\n\nSources:\n' + allCitations.map((c, i) => `${i + 1}. ${c}`).join('\n');
        }
      }

      return finalOutput.trim();

    } else if (responseCode === 401) {
      return "API Error: Invalid API Key.";
    } else if (responseCode === 429) {
      return "API Error: Rate limit exceeded. Please wait a moment and try again.";
    } else {
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
    const response = fetchWithRetry("https://api.exa.ai/contents", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ urls: [url] }),
      headers: { "x-api-key": apiKey, "x-exa-integration": "exa-for-sheets", "User-Agent": "exa-for-sheets 1.1" },
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
        const result = JSON.parse(responseBody);
        const contentData = result.results && result.results[0];
        if (contentData) {
            return (contentData.text || contentData.highlights || "No relevant content found in response.").trim();
        } else {
            return "API returned successfully, but no content data found for this URL.";
        }
    } else if (responseCode === 401) {
        return "API Error: Invalid API Key. Please check your key in the menu.";
    } else if (responseCode === 429) {
        return "API Error: Rate limit exceeded. Please wait a moment and try again.";
    } else {
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
    const response = fetchWithRetry("https://api.exa.ai/findSimilar", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      headers: { "x-api-key": apiKey, "x-exa-integration": "exa-for-sheets", "User-Agent": "exa-for-sheets 1.1" },
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const result = JSON.parse(responseBody);
      if (result && result.results && result.results.length > 0) {
        return result.results.map(item => [item.url || "N/A"]);
      } else {
        return [["No similar URLs found matching the criteria."]];
      }
    } else if (responseCode === 401) {
      return [["API Error: Invalid API Key."]];
    } else if (responseCode === 429) {
      return [["API Error: Rate limit exceeded. Please wait a moment and try again."]];
    } else if (responseCode === 400) {
        let errorMessage = `API Error (Bad Request): Status ${responseCode}.`;
        try {
            const errorResult = JSON.parse(responseBody);
            errorMessage += ` Message: ${errorResult.error || responseBody}`;
        } catch (e) {
            errorMessage += ` Response: ${responseBody}`;
        }
        return [[errorMessage]];
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
    Logger.log(`EXA_FINDSIMILAR Error: ${e} for payload: ${JSON.stringify(payload)}`);
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
 * @param {string} [includeDomainsStr=""] Optional. Comma-separated list of domains to restrict results to (e.g., "linkedin.com,crunchbase.com").
 * @param {string} [excludeDomainsStr=""] Optional. Comma-separated list of domains to exclude from results (e.g., "wikipedia.org,reddit.com").
 * @param {string} [category=""] Optional. Filter by content category: "company", "research paper", "news", "github", "personal site", "pdf", "financial report", "people".
 * @param {number} [highlightsMaxChars=0] Optional. If > 0, requests content highlights with this max character limit per result.
 * @param {string} [outputSchemaJson=""] Optional. JSON string for outputSchema (e.g., '{"type":"text","description":"summarize"}') to get synthesized output.
 * @return {string[]|string} An array of result URLs, synthesized output text, or a single cell error message.
 * @customfunction
 */
function EXA_SEARCH(query, numResults, searchType, prefix, suffix, includeDomainsStr, excludeDomainsStr, category, highlightsMaxChars, outputSchemaJson) {
  const apiKey = getApiKey();
  if (!apiKey) return [["No API key set. Please set your API key in the Exa AI sidebar."]];

  if (!query || typeof query !== 'string' || query.trim() === "") {
    return [["Please provide a valid search query."]];
  }

  // Process the query with optional prefix and suffix
  const finalQuery = `${prefix || ''} ${query} ${suffix || ''}`.trim();

  const count = (typeof numResults === 'number' && numResults > 0 && numResults <= 10) ? Math.floor(numResults) : 1;
  const type = (searchType && ['auto', 'neural', 'keyword'].includes(searchType)) ? searchType : 'auto';

  // Process domain lists (comma-separated string to array)
  const processDomains = (domainStr) => {
    if (typeof domainStr === 'string' && domainStr.trim() !== '') {
      return domainStr.split(',').map(d => d.trim()).filter(d => d.length > 0);
    }
    return null;
  };

  const includeDomains = processDomains(includeDomainsStr);
  const excludeDomains = processDomains(excludeDomainsStr);

  // Validate category if provided
  const validCategories = ['company', 'research paper', 'news', 'github', 'personal site', 'pdf', 'financial report', 'people'];
  const categoryValue = (typeof category === 'string' && category.trim() !== '' && validCategories.includes(category.toLowerCase()))
    ? category.toLowerCase()
    : null;

  // Build payload
  const payload = {
    query: finalQuery,
    numResults: count,
    type: type,
    stream: false,
    useAutoprompt: (type !== 'keyword')
  };

  if (includeDomains && includeDomains.length > 0) {
    payload.includeDomains = includeDomains;
  }
  if (excludeDomains && excludeDomains.length > 0) {
    payload.excludeDomains = excludeDomains;
  }
  if (categoryValue) {
    payload.category = categoryValue;
  }

  // Add contents.highlights if maxCharacters specified
  if (typeof highlightsMaxChars === 'number' && highlightsMaxChars > 0) {
    payload.contents = { highlights: { maxCharacters: highlightsMaxChars } };
  }

  // Parse outputSchema if provided as JSON string
  if (typeof outputSchemaJson === 'string' && outputSchemaJson.trim() !== '') {
    try {
      payload.outputSchema = JSON.parse(outputSchemaJson);
    } catch (e) {
      return [["Invalid outputSchema: must be valid JSON."]];
    }
  }

  try {
    const response = fetchWithRetry("https://api.exa.ai/search", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      headers: { "x-api-key": apiKey, "x-exa-integration": "exa-for-sheets", "User-Agent": "exa-for-sheets 2.0" },
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const result = JSON.parse(responseBody);

      // If outputSchema was used, return the synthesized output text
      if (result.output && result.output.content) {
        return [[result.output.content]];
      }

      if (result && result.results && result.results.length > 0) {
        return result.results.map(item => [item.url]);
      } else {
        return [["API returned successfully, but no search results found."]];
      }
    } else if (responseCode === 401) {
      return [["API Error: Invalid API Key. Please check your key in the menu."]];
    } else if (responseCode === 429) {
      return [["API Error: Rate limit exceeded. Please wait a moment and try again."]];
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
 * Normalizes a user-facing column label into a stable JSON property key.
 * @param {string} label The column label typed in the sidebar
 * @param {Object} usedKeys Set-like object of previously used keys
 * @return {string} A unique, schema-safe property key
 */
function normalizeAgentColumnKey(label, usedKeys) {
  var key = String(label || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');

  if (!key) {
    key = 'field';
  }

  var baseKey = key;
  var suffix = 2;
  while (usedKeys[key]) {
    key = baseKey + '_' + suffix;
    suffix++;
  }

  usedKeys[key] = true;
  return key;
}

/**
 * Converts a schema key back into a readable label when no label was supplied.
 * @param {string} key The schema key
 * @return {string} Human-readable label
 */
function humanizeAgentColumnLabel(key) {
  return String(key || '')
    .replace(/_/g, ' ')
    .replace(/\b\w/g, function(ch) { return ch.toUpperCase(); });
}

/**
 * Parses a sidebar column list into labels and schema keys.
 * Commas, semicolons, pipes, tabs, and new lines are treated as exact separators.
 * Plain text without separators is passed to the Agent as column guidance.
 * @param {string} columnsText Column text from the sidebar
 * @return {Object} Parse result with success flag and columns
 */
function parseAgentColumns(columnsText) {
  var rawText = (typeof columnsText === 'string' && columnsText.trim())
    ? columnsText.trim()
    : '';

  if (!rawText) {
    return { success: true, columns: [], autoColumns: true, columnInstructions: '' };
  }

  var hasExplicitSeparators = /[,;\n\t|]/.test(rawText);
  if (!hasExplicitSeparators) {
    return {
      success: true,
      columns: [],
      autoColumns: true,
      columnInstructions: rawText
    };
  }

  var labels = rawText
    .split(/[,;\n\t|]+/)
    .map(function(label) { return label.trim(); })
    .filter(function(label) { return label.length > 0; });

  if (labels.length === 0) {
    return { success: true, columns: [], autoColumns: true, columnInstructions: '' };
  }

  var usedKeys = {};
  var columns = labels.map(function(label) {
    var key = normalizeAgentColumnKey(label, usedKeys);
    var isUrlList = key === 'source_urls' || key === 'sources' || key === 'urls' || key.match(/_urls$/);
    return {
      label: label,
      key: key,
      type: isUrlList ? 'array' : 'string'
    };
  });

  return { success: true, columns: columns, autoColumns: false, columnInstructions: rawText };
}

/**
 * Builds the JSON schema sent to /agent/runs for a spreadsheet table.
 * @param {Object[]} columns Parsed column definitions
 * @param {number} rowLimit Maximum number of rows to request
 * @return {Object} JSON schema
 */
function buildAgentTableOutputSchema(columns, rowLimit) {
  if (!columns || columns.length === 0) {
    return {
      type: 'object',
      required: ['columns', 'rows'],
      additionalProperties: false,
      properties: {
        columns: {
          type: 'array',
          description: 'Spreadsheet column headers inferred from the user request, in display order.',
          minItems: 1,
          items: { type: 'string' }
        },
        rows: {
          type: 'array',
          description: 'Spreadsheet rows. Each row has cells in the same order as the columns array.',
          minItems: 1,
          maxItems: rowLimit,
          items: {
            type: 'object',
            required: ['cells'],
            additionalProperties: false,
            properties: {
              cells: {
                type: 'array',
                description: 'Cell values for this row, in the exact same order as the columns array.',
                minItems: 1,
                items: { type: 'string' }
              }
            }
          }
        }
      }
    };
  }

  var properties = {};
  var required = [];

  columns.forEach(function(column) {
    required.push(column.key);
    if (column.type === 'array') {
      properties[column.key] = {
        type: 'array',
        description: 'Values for the "' + column.label + '" column.',
        items: { type: 'string' }
      };
    } else {
      properties[column.key] = {
        type: 'string',
        description: 'Value for the "' + column.label + '" column.'
      };
    }
  });

  return {
    type: 'object',
    required: ['rows'],
    additionalProperties: false,
    properties: {
      rows: {
        type: 'array',
        description: 'Spreadsheet rows matching the requested table.',
        minItems: 1,
        maxItems: rowLimit,
        items: {
          type: 'object',
          required: required,
          additionalProperties: false,
          properties: properties
        }
      }
    }
  };
}

/**
 * Builds a clear table-oriented Agent prompt while preserving the user's ask.
 * @param {string} prompt User prompt
 * @param {Object[]} columns Parsed columns
 * @param {number} rowLimit Maximum number of rows
 * @param {string} columnInstructions Optional free-text column guidance
 * @return {string} Prompt for Exa Agent
 */
function buildAgentTablePrompt(prompt, columns, rowLimit, columnInstructions) {
  var basePrompt = String(prompt || '').trim() + '\n\n' +
    'Create a spreadsheet-ready table. Return no more than ' + rowLimit + ' rows. ' +
    'Use the row count requested by the user when it is within that limit. ' +
    'If the user did not request a count, return the most useful set of rows up to that limit. ' +
    'Use web research. Do not rely on memory. ' +
    'Every row must be a real, source-verifiable entity that matches the request. ' +
    'Do not pad the table with weak matches or duplicates just to reach the row count. ' +
    'Try to fill the table as completely as public sources allow, but use empty strings for unknown values rather than inventing facts. ' +
    'Prefer official, primary, or otherwise high-quality sources when possible. ' +
    'Do not return notes, summaries, metadata, or separate partial-results arrays unless the user explicitly asks for them.';

  if (!columns || columns.length === 0) {
    var columnGuidance = columnInstructions
      ? ' Use these exact columns if possible: ' + columnInstructions + '.'
      : ' Infer the best concise column headers from the user request.';
    return basePrompt + columnGuidance + ' ' +
      'Return a columns array and rows where each row is an object with a cells array. ' +
      'Each cells array must contain values in the exact same order as the columns array. ' +
      'Do not return empty row objects.';
  }

  var columnList = columns.map(function(column) {
    return column.label + ' (schema key: ' + column.key + ')';
  }).join(', ');

  return basePrompt + ' Populate exactly these columns: ' + columnList + '.';
}

/**
 * Validates and normalizes the Agent effort setting.
 * @param {string} effort User-selected effort
 * @return {string} Valid effort
 */
function normalizeAgentEffort(effort) {
  var value = (typeof effort === 'string' && effort.trim()) ? effort.trim().toLowerCase() : 'auto';
  return EXA_AGENT_VALID_EFFORTS.includes(value) ? value : 'auto';
}

/**
 * Infers a useful row count from prompts like "top 50", "10 companies", or "return 25 rows".
 * @param {string} prompt User prompt
 * @param {*} explicitRowLimit Optional hidden/configured row count
 * @return {number} Requested row count or the default
 */
function inferAgentRowLimit(prompt, explicitRowLimit) {
  var configured = parseInt(explicitRowLimit, 10);
  if (configured && configured > 0) {
    return configured;
  }

  var text = String(prompt || '').toLowerCase();
  var patterns = [
    /\btop\s+(\d+)\b/,
    /\bfirst\s+(\d+)\b/,
    /\b(\d+)\s+(?:rows?|companies|items|people|startups|leads|results|entries)\b/,
    /\b(?:return|find|list|give me)\s+(\d+)\b/
  ];

  for (var i = 0; i < patterns.length; i++) {
    var match = text.match(patterns[i]);
    if (match) {
      var inferred = parseInt(match[1], 10);
      if (inferred && inferred > 0) {
        return inferred;
      }
    }
  }

  return EXA_AGENT_CONFIG.defaultRows;
}

/**
 * Starts an Exa Agent run designed to produce a spreadsheet table.
 * Called from the sidebar.
 *
 * @param {Object} config Sidebar options
 * @return {Object} Result with run summary and table config
 */
function startAgentTableRun(config) {
  const apiKey = getApiKey();
  if (!apiKey) {
    return {
      success: false,
      message: 'No API key set. Please set your API key in the Exa AI sidebar.'
    };
  }

  config = config || {};
  var prompt = (typeof config.prompt === 'string') ? config.prompt.trim() : '';
  if (!prompt) {
    return { success: false, message: 'Please enter an Agent prompt.' };
  }

  var rowLimit = inferAgentRowLimit(prompt, config.rowLimit);

  var parsedColumns = parseAgentColumns(config.columns);
  if (!parsedColumns.success) {
    return parsedColumns;
  }

  var columns = parsedColumns.columns;
  var autoColumns = parsedColumns.autoColumns === true;
  var effort = normalizeAgentEffort(config.effort);
  var requestedStartCell = normalizeAgentStartCell(config.startCell);
  if (requestedStartCell && !isValidAgentStartCell(requestedStartCell)) {
    return {
      success: false,
      message: 'Start cell must be a single cell like A1 or D5.'
    };
  }
  var sheet = SpreadsheetApp.getActiveSheet();
  var sheetContext = getAgentSheetContext(sheet);
  var startCell = requestedStartCell || getCurrentAgentStartCell(sheet);
  var outputSchema = buildAgentTableOutputSchema(columns, rowLimit);

  var payload = {
    query: buildAgentTablePrompt(prompt, columns, rowLimit, parsedColumns.columnInstructions || ''),
    systemPrompt: 'You are filling a Google Sheets table. Return precise, concise cell values and satisfy the provided JSON schema.',
    outputSchema: outputSchema,
    effort: effort,
    metadata: {
      integration: 'exa-for-sheets',
      mode: 'agent-table',
      rowLimit: String(rowLimit),
      columns: columns.map(function(column) { return column.key; }).join(','),
      autoColumns: String(autoColumns),
      columnInstructions: parsedColumns.columnInstructions || '',
      startCell: startCell
    }
  };

  try {
    var response = fetchWithRetry(EXA_AGENT_CONFIG.endpoint, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      headers: {
        'x-api-key': apiKey,
        'Exa-Beta': EXA_AGENT_CONFIG.betaHeader,
        'x-exa-integration': 'exa-for-sheets',
        'User-Agent': 'exa-for-sheets 2.0'
      },
      muteHttpExceptions: true
    });

    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();

    if (responseCode === 200) {
      var run = JSON.parse(responseBody);
      return {
        success: true,
        message: 'Agent run started.',
        run: summarizeAgentRun(run),
        tableConfig: {
          rootKey: 'rows',
          columns: columns,
          rowLimit: rowLimit,
          autoColumns: autoColumns,
          startCell: startCell,
          sheetId: sheetContext.sheetId,
          sheetName: sheetContext.sheetName
        }
      };
    }

    return {
      success: false,
      message: parseAgentApiError(responseCode, responseBody, 'Failed to start Agent run.')
    };
  } catch (e) {
    Logger.log('startAgentTableRun Error: ' + e);
    return { success: false, message: 'Script Error: ' + e.message };
  }
}

/**
 * Builds fill jobs from the active selection without starting Agent runs.
 * Called from the sidebar before queueing one or more Agent fill runs.
 *
 * @param {Object} config Sidebar options
 * @return {Object} Fill preparation result
 */
function prepareAgentFillJobs(config) {
  const apiKey = getApiKey();
  if (!apiKey) {
    return {
      success: false,
      message: 'No API key set. Please set your API key in the Exa AI sidebar.'
    };
  }

  config = config || {};
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var selection = sheet && sheet.getActiveRange ? sheet.getActiveRange() : null;
    if (!selection) {
      return { success: false, message: 'Select cells to fill first.' };
    }

    var result = buildAgentFillJobsFromSelection(sheet, selection, {
      overwrite: config.overwrite === true,
      instructions: readAgentFillInstructions(config.instructions)
    });

    if (!result.success) {
      return result;
    }

    return {
      success: true,
      message: 'Ready to fill selection.',
      jobs: result.jobs,
      totalCells: result.totalCells,
      totalJobs: result.jobs.length,
      selectedRange: result.selectedRange || '',
      skippedCells: result.skippedCells || 0
    };
  } catch (e) {
    Logger.log('prepareAgentFillJobs Error: ' + e);
    return { success: false, message: 'Could not read the selected cells. Try selecting the cells again.' };
  }
}

/**
 * Starts one Agent run for a prepared fill job.
 *
 * @param {Object} job Prepared fill job
 * @param {Object} config Sidebar options
 * @return {Object} Result with run summary and fill config
 */
function startAgentFillRun(job, config) {
  const apiKey = getApiKey();
  if (!apiKey) {
    return { success: false, message: 'No API key set.' };
  }

  job = job || {};
  config = config || {};
  if (!Array.isArray(job.targets) || job.targets.length === 0) {
    return { success: false, message: 'No blank cells selected.' };
  }

  var effort = normalizeAgentEffort(config.effort);
  var instructions = readAgentFillInstructions(config.instructions || job.instructions);
  var isContinueTable = job.mode === 'continue-table';
  var payload = {
    query: isContinueTable ? buildAgentContinueTablePrompt(job, instructions) : buildAgentFillPrompt(job, instructions),
    systemPrompt: isContinueTable
      ? 'You are continuing a Google Sheets table. Return concise, verified cell values and satisfy the provided JSON schema.'
      : 'You are filling selected cells in a Google Sheet. Return concise, verified values and satisfy the provided JSON schema.',
    input: {
      data: job.inputRows || []
    },
    outputSchema: buildAgentFillOutputSchema(job),
    effort: effort,
    metadata: {
      integration: 'exa-for-sheets',
      mode: isContinueTable ? 'agent-continue-table' : 'agent-fill',
      jobId: String(job.jobId || ''),
      targetCells: String(job.targets.length),
      targetRows: String(countUniqueAgentFillRows(job.targets)),
      fields: (job.fields || []).map(function(field) { return field.fieldId; }).join(','),
      overwrite: String(job.overwrite === true)
    }
  };

  try {
    var response = fetchWithRetry(EXA_AGENT_CONFIG.endpoint, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      headers: {
        'x-api-key': apiKey,
        'Exa-Beta': EXA_AGENT_CONFIG.betaHeader,
        'x-exa-integration': 'exa-for-sheets',
        'User-Agent': 'exa-for-sheets 2.0'
      },
      muteHttpExceptions: true
    });

    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();

    if (responseCode === 200) {
      var run = JSON.parse(responseBody);
      return {
        success: true,
        message: 'Fill started.',
        run: summarizeAgentRun(run),
        fillConfig: sanitizeAgentFillConfig(job)
      };
    }

    return {
      success: false,
      message: parseAgentApiError(responseCode, responseBody, 'Failed to start fill run.')
    };
  } catch (e) {
    Logger.log('startAgentFillRun Error: ' + e);
    return { success: false, message: 'Script Error: ' + e.message };
  }
}

/**
 * Writes a completed Agent fill run back into selected target cells.
 *
 * @param {string} runId Exa Agent run ID
 * @param {Object} fillConfig Fill config returned by startAgentFillRun
 * @return {Object} Write result
 */
function writeAgentFillRunToSheet(runId, fillConfig) {
  const apiKey = getApiKey();
  if (!apiKey) {
    return { success: false, message: 'No API key set.' };
  }

  var fetched = fetchAgentRun(apiKey, runId);
  if (!fetched.success) {
    return fetched;
  }

  var run = fetched.run;
  if (run.status !== 'completed') {
    return {
      success: false,
      message: 'Agent run is not complete yet. Current status: ' + run.status
    };
  }

  var extracted = extractAgentFillValues(run, fillConfig || {});
  if (!extracted.success) {
    return extracted;
  }

  try {
    var sheetResult = getAgentWriteSheet(fillConfig || {});
    if (!sheetResult.success) {
      return sheetResult;
    }

    var sheet = sheetResult.sheet;
    var filled = 0;
    var blank = 0;
    var skipped = 0;
    var overwrite = fillConfig && fillConfig.overwrite === true;

    extracted.targets.forEach(function(target) {
      var cellRange = sheet.getRange(target.row, target.column);
      var currentValue = cellRange.getValue ? cellRange.getValue() : '';
      var currentFormula = cellRange.getFormula ? cellRange.getFormula() : '';
      if (!overwrite && !isAgentFillCellBlank(currentValue, currentFormula)) {
        skipped++;
        return;
      }

      if (target.value === '' || target.value === null || target.value === undefined) {
        blank++;
        return;
      }

      cellRange.setValue(formatAgentCellValue(target.value));
      filled++;
    });

    SpreadsheetApp.flush();

    return {
      success: true,
      message: blank > 0 || skipped > 0 ? 'Done. Some cells were left blank.' : 'Done.',
      runId: run.id,
      filled: filled,
      blank: blank,
      skipped: skipped,
      costDollars: run.costDollars || null,
      usage: run.usage || null
    };
  } catch (e) {
    Logger.log('writeAgentFillRunToSheet Error: ' + e);
    return { success: false, message: 'Could not write the filled cells. Try again.' };
  }
}

/**
 * Builds prepared fill jobs from an active range.
 *
 * @param {Sheet} sheet Active sheet
 * @param {Range} selection Active selection
 * @param {Object} options Fill options
 * @return {Object} Jobs result
 */
function buildAgentFillJobsFromSelection(sheet, selection, options) {
  options = options || {};
  var overwrite = options.overwrite === true;
  var startRow = selection.getRow();
  var startCol = selection.getColumn();
  var numRows = selection.getNumRows ? selection.getNumRows() : 1;
  var numCols = selection.getNumColumns ? selection.getNumColumns() : 1;
  var endRow = startRow + numRows - 1;
  var endCol = startCol + numCols - 1;
  var selectionValues = selection.getValues ? selection.getValues() : [[selection.getValue ? selection.getValue() : '']];
  var selectionFormulas = selection.getFormulas ? selection.getFormulas() : createBlankMatrix(numRows, numCols);
  var lastRow = Math.max(endRow, sheet.getLastRow ? sheet.getLastRow() : endRow);
  var lastCol = Math.max(endCol, sheet.getLastColumn ? sheet.getLastColumn() : endCol);
  var scanEndCol = Math.min(lastCol, Math.max(endCol, 50));
  var scanNumCols = scanEndCol;
  var headerRow = detectAgentFillHeaderRow(sheet, startRow, startCol, endCol, scanNumCols);
  var table = buildAgentFillTableInfo(sheet, headerRow, startCol, endCol, scanNumCols, startRow, endRow);
  var tableValues = readSheetValues(sheet, table.dataStartRow, table.startCol, table.dataEndRow - table.dataStartRow + 1, table.endCol - table.startCol + 1);
  var rowContexts = buildAgentFillRowContexts(tableValues, table.dataStartRow, table);
  var rowContextByNumber = {};
  rowContexts.forEach(function(rowContext) {
    rowContextByNumber[rowContext.rowNumber] = rowContext;
  });

  var targets = [];
  var continuationCandidates = [];
  var skippedCells = 0;
  var blankCandidateCells = 0;
  var cellsWithoutRowContext = 0;
  for (var rowOffset = 0; rowOffset < numRows; rowOffset++) {
    for (var colOffset = 0; colOffset < numCols; colOffset++) {
      var row = startRow + rowOffset;
      var col = startCol + colOffset;
      if (headerRow && row === headerRow) {
        skippedCells++;
        continue;
      }

      var value = selectionValues[rowOffset] ? selectionValues[rowOffset][colOffset] : '';
      var formula = selectionFormulas[rowOffset] ? selectionFormulas[rowOffset][colOffset] : '';
      if (!overwrite && !isAgentFillCellBlank(value, formula)) {
        skippedCells++;
        continue;
      }

      blankCandidateCells++;
      var field = table.fieldByColumn[col] || createAgentFillFallbackField(col);
      var target = {
        cell: formatCellA1(row, col),
        row: row,
        column: col,
        rowId: createAgentFillRowId(row),
        fieldId: field.fieldId,
        title: field.title,
        columnLetter: formatColumnLabel(col)
      };
      var rowContext = rowContextByNumber[row];
      if (!rowContext || !rowHasAgentFillContext(rowContext)) {
        cellsWithoutRowContext++;
        continuationCandidates.push(target);
        continue;
      }

      targets.push(target);
    }
  }

  var selectedRange = selection.getA1Notation ? selection.getA1Notation() : formatRangeA1(startRow, startCol, numRows, numCols);
  var jobs = buildAgentFillRowJobs({
    targets: targets,
    table: table,
    rowContexts: rowContexts,
    rowContextByNumber: rowContextByNumber,
    sheet: sheet,
    selectedRange: selectedRange,
    overwrite: overwrite,
    instructions: readAgentFillInstructions(options.instructions)
  });
  var continuation = buildAgentContinueTableJobs({
    candidates: continuationCandidates,
    table: table,
    rowContexts: rowContexts,
    rowContextByNumber: rowContextByNumber,
    sheet: sheet,
    selectedRange: selectedRange,
    overwrite: overwrite,
    instructions: readAgentFillInstructions(options.instructions)
  });

  if (continuation.success) {
    jobs = jobs.concat(continuation.jobs);
    skippedCells += continuation.skippedCells || 0;
  } else {
    skippedCells += continuationCandidates.length;
  }

  if (jobs.length === 0) {
    if (blankCandidateCells > 0 && cellsWithoutRowContext === blankCandidateCells) {
      return {
        success: false,
        message: 'Select blanks in rows that already have some data.'
      };
    }
    return { success: false, message: overwrite ? 'No fillable cells selected.' : 'No blank cells selected.' };
  }

  return {
    success: true,
    jobs: jobs,
    totalCells: targets.length + (continuation.success ? continuation.totalCells : 0),
    selectedRange: selectedRange,
    skippedCells: skippedCells
  };
}

function buildAgentFillRowJobs(args) {
  var targets = args.targets || [];
  if (targets.length === 0) {
    return [];
  }

  var selectedRows = uniqueNumbers(targets.map(function(target) { return target.row; }));
  var selectedFields = uniqueAgentFillFields(targets, args.table);
  var exampleRows = pickAgentFillExampleRows(args.rowContexts, selectedRows, 5);
  var chunkSize = selectedFields.length <= 1 ? 10 : (selectedRows.length === 1 ? 1 : 3);
  var rowChunks = chunkArray(selectedRows, chunkSize);

  return rowChunks.map(function(rowChunk, index) {
    var rowSet = {};
    rowChunk.forEach(function(row) { rowSet[row] = true; });
    var jobTargets = targets.filter(function(target) { return rowSet[target.row] === true; });
    var jobFields = uniqueAgentFillFields(jobTargets, args.table);
    var jobRows = rowChunk.map(function(row) { return args.rowContextByNumber[row]; }).filter(Boolean);
    return {
      jobId: 'fill_' + (index + 1),
      mode: 'fill',
      sheetId: args.sheet.getSheetId ? args.sheet.getSheetId() : null,
      sheetName: args.sheet.getName ? args.sheet.getName() : '',
      selectedRange: args.selectedRange,
      overwrite: args.overwrite,
      instructions: args.instructions,
      headerRow: args.table.headerRow || null,
      fields: jobFields,
      targets: jobTargets,
      inputRows: jobRows.concat(exampleRows),
      targetRowIds: rowChunk.map(createAgentFillRowId)
    };
  });
}

function buildAgentContinueTableJobs(args) {
  var candidates = args.candidates || [];
  if (candidates.length === 0) {
    return { success: false, jobs: [], totalCells: 0 };
  }

  var table = args.table || {};
  if (!table.headerRow || !Array.isArray(table.fields) || table.fields.length < 2) {
    return { success: false, jobs: [], totalCells: 0 };
  }

  var targets = candidates.filter(function(target) {
    var field = table.fieldByColumn && table.fieldByColumn[target.column];
    return field && field.hasHeader === true;
  });
  var skippedCells = candidates.length - targets.length;
  if (targets.length === 0) {
    return { success: false, jobs: [], totalCells: 0 };
  }

  var selectedRows = uniqueNumbers(targets.map(function(target) { return target.row; }));
  if (selectedRows.length === 0) {
    return { success: false, jobs: [], totalCells: 0 };
  }

  var allSelectedRowsAreBlank = selectedRows.every(function(row) {
    return !rowHasAgentFillContext(args.rowContextByNumber[row]);
  });
  if (!allSelectedRowsAreBlank) {
    return { success: false, jobs: [], totalCells: 0 };
  }

  var exampleRows = pickAgentContinueTableExampleRows(args.rowContexts, selectedRows, 10);
  if (exampleRows.length < 2) {
    return { success: false, jobs: [], totalCells: 0 };
  }

  var firstSelectedRow = selectedRows[0];
  var lastExampleRow = Math.max.apply(null, exampleRows.map(function(row) { return row.rowNumber || 0; }));
  if (!lastExampleRow || firstSelectedRow > lastExampleRow + 5) {
    return { success: false, jobs: [], totalCells: 0 };
  }

  var fields = uniqueAgentFillFields(targets, table);
  return {
    success: true,
    totalCells: targets.length,
    skippedCells: skippedCells,
    jobs: [{
      jobId: 'continue_1',
      mode: 'continue-table',
      sheetId: args.sheet.getSheetId ? args.sheet.getSheetId() : null,
      sheetName: args.sheet.getName ? args.sheet.getName() : '',
      selectedRange: args.selectedRange,
      overwrite: args.overwrite,
      instructions: args.instructions,
      headerRow: table.headerRow || null,
      fields: fields,
      targets: targets,
      inputRows: exampleRows,
      targetRowIds: selectedRows.map(createAgentFillRowId)
    }]
  };
}

/**
 * Starts from the selection and finds the nearest likely header row above it.
 */
function detectAgentFillHeaderRow(sheet, selectionStartRow, selectionStartCol, selectionEndCol, scanNumCols) {
  if (selectionStartRow <= 1) {
    return 0;
  }

  var scanStartRow = Math.max(1, selectionStartRow - 500);
  var scanRows = selectionStartRow - scanStartRow;
  if (scanRows <= 0) {
    return 0;
  }

  var values = readSheetValues(sheet, scanStartRow, 1, scanRows, scanNumCols);
  var topOfBlockIndex = values.length - 1;
  while (topOfBlockIndex >= 0 && countNonEmptyCells([values[topOfBlockIndex]]) > 0) {
    topOfBlockIndex--;
  }

  var candidateIndex = topOfBlockIndex + 1;
  if (candidateIndex < values.length) {
    var candidateRow = values[candidateIndex];
    var candidateCount = countNonEmptyCells([candidateRow]);
    if (candidateCount >= 2 && (scanStartRow === 1 || candidateIndex > 0)) {
      return scanStartRow + candidateIndex;
    }
  }

  var frozenHeaderRow = getLikelyFrozenHeaderRow(sheet, selectionStartRow, scanNumCols);
  if (frozenHeaderRow) {
    return frozenHeaderRow;
  }

  var topHeaderRow = getLikelyTopHeaderRow(sheet, selectionStartRow, scanNumCols);
  if (topHeaderRow) {
    return topHeaderRow;
  }

  if (selectionStartRow > 1) {
    var firstRow = readSheetValues(sheet, 1, 1, 1, scanNumCols)[0] || [];
    if (countNonEmptyCells([firstRow]) >= 2) {
      return 1;
    }
  }

  return 0;
}

/**
 * Uses frozen rows as a strong hint for sheets with long tables.
 */
function getLikelyFrozenHeaderRow(sheet, selectionStartRow, scanNumCols) {
  try {
    var frozenRows = sheet.getFrozenRows ? sheet.getFrozenRows() : 0;
    if (!frozenRows || frozenRows >= selectionStartRow) {
      return 0;
    }

    var frozenRow = readSheetValues(sheet, frozenRows, 1, 1, scanNumCols)[0] || [];
    return countNonEmptyCells([frozenRow]) >= 2 ? frozenRows : 0;
  } catch (e) {
    Logger.log('getLikelyFrozenHeaderRow skipped: ' + e);
    return 0;
  }
}

/**
 * Finds a likely table header near the top of the sheet for long selections.
 */
function getLikelyTopHeaderRow(sheet, selectionStartRow, scanNumCols) {
  var rowsToCheck = Math.min(20, selectionStartRow - 1);
  if (rowsToCheck <= 0) {
    return 0;
  }

  var topRows = readSheetValues(sheet, 1, 1, rowsToCheck, scanNumCols);
  for (var i = 0; i < topRows.length; i++) {
    if (countNonEmptyCells([topRows[i]]) >= 2) {
      return i + 1;
    }
  }

  return 0;
}

/**
 * Builds field metadata for the table around the selected range.
 */
function buildAgentFillTableInfo(sheet, headerRow, selectionStartCol, selectionEndCol, scanNumCols, selectionStartRow, selectionEndRow) {
  var startCol = selectionStartCol;
  var endCol = selectionEndCol;
  var headers = [];

  if (headerRow) {
    var headerValues = readSheetValues(sheet, headerRow, 1, 1, scanNumCols)[0] || [];
    var bounds = findAgentFillHeaderBounds(headerValues, selectionStartCol, selectionEndCol);
    startCol = bounds.startCol;
    endCol = bounds.endCol;
    headers = headerValues.slice(startCol - 1, endCol);
  } else {
    startCol = 1;
    endCol = Math.max(selectionEndCol, Math.min(scanNumCols, selectionEndCol + 3));
    for (var col = startCol; col <= endCol; col++) {
      headers.push(formatColumnLabel(col));
    }
  }

  var fields = [];
  var fieldByColumn = {};
  var usedKeys = {};
  for (var currentCol = startCol; currentCol <= endCol; currentCol++) {
    var rawHeader = headers[currentCol - startCol];
    var hasHeader = !isBlankValue(rawHeader) && String(rawHeader).trim() !== '';
    var title = String(rawHeader || '').trim() || formatColumnLabel(currentCol);
    var fieldId = normalizeAgentColumnKey(title, usedKeys);
    var field = {
      fieldId: fieldId,
      title: title,
      column: currentCol,
      columnLetter: formatColumnLabel(currentCol),
      hasHeader: hasHeader
    };
    fields.push(field);
    fieldByColumn[currentCol] = field;
  }

  var lastRow = Math.max(sheet.getLastRow ? sheet.getLastRow() : selectionEndRow, selectionEndRow);
  var dataStartRow = headerRow
    ? Math.max(headerRow + 1, selectionStartRow - 5)
    : Math.max(1, selectionStartRow - 5);
  var dataEndRow = Math.min(lastRow, selectionEndRow + 5);
  return {
    startCol: startCol,
    endCol: endCol,
    headerRow: headerRow || null,
    dataStartRow: dataStartRow,
    dataEndRow: Math.max(dataEndRow, dataStartRow),
    fields: fields,
    fieldByColumn: fieldByColumn
  };
}

/**
 * Finds the broad header span that contains the selected columns.
 */
function findAgentFillHeaderBounds(headerValues, selectionStartCol, selectionEndCol) {
  var first = 0;
  var last = 0;
  for (var i = 0; i < headerValues.length; i++) {
    if (!isBlankValue(headerValues[i])) {
      if (!first) first = i + 1;
      last = i + 1;
    }
  }

  if (!first || !last) {
    return { startCol: selectionStartCol, endCol: selectionEndCol };
  }

  return {
    startCol: Math.min(first, selectionStartCol),
    endCol: Math.max(last, selectionEndCol)
  };
}

/**
 * Converts table rows into structured row context for Agent input.
 */
function buildAgentFillRowContexts(tableValues, dataStartRow, table) {
  return tableValues.map(function(rowValues, index) {
    var rowNumber = dataStartRow + index;
    var cells = table.fields.map(function(field, fieldIndex) {
      return {
        fieldId: field.fieldId,
        title: field.title,
        column: field.columnLetter,
        value: formatAgentInputCellValue(rowValues[fieldIndex])
      };
    });

    return {
      rowId: createAgentFillRowId(rowNumber),
      rowNumber: rowNumber,
      cells: cells
    };
  });
}

/**
 * Builds a prompt for one fill job.
 */
function buildAgentFillPrompt(job, instructions) {
  var fieldLines = (job.fields || []).map(function(field) {
    return '- ' + field.fieldId + ' (' + field.title + ', column ' + field.columnLetter + ')';
  }).join('\n');

  var targetLines = (job.targets || []).map(function(target) {
    return '- ' + target.cell + ': rowId ' + target.rowId + ', field ' + target.fieldId + ' (' + target.title + ')';
  }).join('\n');

  var guidance = instructions
    ? '\nUser guidance: ' + instructions
    : '';

  return [
    'Fill selected blank cells in a Google Sheets table.',
    'Use input.data as the spreadsheet context. Each row has cells with field ids, titles, columns, and current values.',
    'Use web research. Do not rely on memory.',
    'Fill only the requested target cells listed below.',
    'Keep values short, clean, and spreadsheet-ready.',
    'Return null when a value cannot be verified from reliable public sources.',
    'Do not invent values. Do not return notes, citations, sources, markdown, or extra fields.',
    '',
    'Target fields:',
    fieldLines,
    '',
    'Target cells:',
    targetLines,
    guidance,
    '',
    'Return exactly one rows item for each target rowId. Preserve rowId exactly and put cell values under values.<fieldId>.'
  ].filter(function(part) { return part !== ''; }).join('\n');
}

/**
 * Builds a prompt for continuing blank rows directly below an existing table.
 */
function buildAgentContinueTablePrompt(job, instructions) {
  var fieldLines = (job.fields || []).map(function(field) {
    return '- ' + field.fieldId + ' (' + field.title + ', column ' + field.columnLetter + ')';
  }).join('\n');

  var targetRows = formatAgentTargetRows(job.targets || []);
  var guidance = instructions
    ? '\nUser guidance: ' + instructions
    : '';

  return [
    'Continue the existing Google Sheets table into the selected blank rows.',
    'Use input.data as the table context. It contains real rows already in the sheet, including nearby rows before the blank selection.',
    'Use web research. Do not rely on memory.',
    'Fill only the requested target cells listed below.',
    'Keep the same topic, ordering, formatting, and value style as the existing rows.',
    'If a selected column is a rank, index, date, year, or other sequence, continue the sequence from the existing rows.',
    'Do not repeat entities already present in input.data.',
    'Every new row must be a real, source-verifiable entity that belongs next in the table.',
    'Keep values short, clean, and spreadsheet-ready.',
    'Return null when a value cannot be verified from reliable public sources.',
    'Do not invent values. Do not return notes, citations, sources, markdown, or extra fields.',
    '',
    'Target fields:',
    fieldLines,
    '',
    'Target rows and cells:',
    targetRows,
    guidance,
    '',
    'Return exactly one rows item for each target rowId. Preserve rowId exactly and put cell values under values.<fieldId>.'
  ].filter(function(part) { return part !== ''; }).join('\n');
}

function formatAgentTargetRows(targets) {
  var rows = {};
  var order = [];
  targets.forEach(function(target) {
    if (!rows[target.rowId]) {
      rows[target.rowId] = {
        row: target.row,
        cells: []
      };
      order.push(target.rowId);
    }
    rows[target.rowId].cells.push(target.cell + ' = ' + target.fieldId + ' (' + target.title + ')');
  });

  order.sort(function(a, b) {
    return (rows[a].row || 0) - (rows[b].row || 0);
  });

  return order.map(function(rowId) {
    return '- ' + rowId + ' (sheet row ' + rows[rowId].row + '): ' + rows[rowId].cells.join(', ');
  }).join('\n');
}

/**
 * Builds the strict output schema for one fill job.
 */
function buildAgentFillOutputSchema(job) {
  var rowIds = (job.targetRowIds && job.targetRowIds.length)
    ? job.targetRowIds
    : uniqueStrings((job.targets || []).map(function(target) { return target.rowId; }));
  var properties = {};
  var requiredFields = [];
  (job.fields || []).forEach(function(field) {
    requiredFields.push(field.fieldId);
    properties[field.fieldId] = {
      anyOf: [
        { type: 'string' },
        { type: 'number' },
        { type: 'boolean' },
        { type: 'array', items: { type: 'string' } },
        { type: 'null' }
      ],
      description: 'Value for ' + field.title + '. Return null if unavailable.'
    };
  });
  var valuesSchema = {
    type: 'object',
    additionalProperties: false,
    properties: properties
  };
  if (job && job.mode === 'continue-table' && requiredFields.length > 0) {
    valuesSchema.required = requiredFields;
  }

  return {
    type: 'object',
    required: ['rows'],
    additionalProperties: false,
    properties: {
      rows: {
        type: 'array',
        minItems: rowIds.length,
        maxItems: rowIds.length,
        items: {
          type: 'object',
          required: ['rowId', 'values'],
          additionalProperties: false,
          properties: {
            rowId: {
              type: 'string',
              enum: rowIds
            },
            values: valuesSchema
          }
        }
      }
    }
  };
}

/**
 * Extracts and validates fill values from a completed run.
 */
function extractAgentFillValues(run, fillConfig) {
  var structured = run && run.output ? run.output.structured : null;
  if (!structured) {
    return { success: false, message: 'Agent run completed, but no structured output was returned.' };
  }

  var records = Array.isArray(structured)
    ? structured
    : (Array.isArray(structured.rows) ? structured.rows : []);
  if (records.length === 0) {
    return { success: false, message: 'Agent did not return any filled values.' };
  }

  var targets = Array.isArray(fillConfig.targets) ? fillConfig.targets : [];
  var allowed = {};
  targets.forEach(function(target) {
    allowed[target.rowId + '::' + target.fieldId] = target;
  });

  var valuesByTarget = {};
  records.forEach(function(record) {
    if (!record || typeof record !== 'object') return;
    var rowId = String(record.rowId || record.row_id || '');
    var values = record.values && typeof record.values === 'object' ? record.values : record;
    Object.keys(values).forEach(function(fieldId) {
      var key = rowId + '::' + fieldId;
      if (allowed[key]) {
        valuesByTarget[key] = values[fieldId];
      }
    });
  });

  return {
    success: true,
    targets: targets.map(function(target) {
      var key = target.rowId + '::' + target.fieldId;
      return {
        row: target.row,
        column: target.column,
        cell: target.cell,
        rowId: target.rowId,
        fieldId: target.fieldId,
        value: Object.prototype.hasOwnProperty.call(valuesByTarget, key) ? valuesByTarget[key] : ''
      };
    })
  };
}

/**
 * Keeps only the fill config fields needed for writing a result.
 */
function sanitizeAgentFillConfig(job) {
  return {
    jobId: String(job.jobId || ''),
    sheetId: job.sheetId === null || job.sheetId === undefined ? null : Number(job.sheetId),
    sheetName: String(job.sheetName || ''),
    overwrite: job.overwrite === true,
    targets: (job.targets || []).map(function(target) {
      return {
        cell: target.cell,
        row: target.row,
        column: target.column,
        rowId: target.rowId,
        fieldId: target.fieldId,
        title: target.title
      };
    }),
    fields: (job.fields || []).map(function(field) {
      return {
        fieldId: field.fieldId,
        title: field.title,
        column: field.column,
        columnLetter: field.columnLetter
      };
    })
  };
}

function readAgentFillInstructions(instructions) {
  return (typeof instructions === 'string') ? instructions.trim() : '';
}

function isAgentFillCellBlank(value, formula) {
  return isBlankValue(value) && isBlankValue(formula);
}

function isBlankValue(value) {
  return value === '' || value === null || value === undefined;
}

function formatAgentInputCellValue(value) {
  if (value === null || value === undefined) return '';
  if (isDateLikeValue(value)) {
    try {
      return value.toISOString();
    } catch (e) {
      return String(value);
    }
  }
  if (Array.isArray(value)) {
    return value.map(formatAgentInputCellValue).join(', ');
  }
  if (typeof value === 'object') {
    try {
      return JSON.stringify(value);
    } catch (e) {
      return String(value);
    }
  }
  return value;
}

function isDateLikeValue(value) {
  return Object.prototype.toString.call(value) === '[object Date]' ||
    (value && typeof value === 'object' && typeof value.toISOString === 'function' && typeof value.getTime === 'function');
}

function rowHasAgentFillContext(rowContext) {
  return !!(rowContext && Array.isArray(rowContext.cells) && rowContext.cells.some(function(cell) {
    return !isBlankValue(cell.value);
  }));
}

function pickAgentFillExampleRows(rowContexts, selectedRows, limit) {
  var selected = {};
  selectedRows.forEach(function(row) { selected[row] = true; });
  return rowContexts
    .filter(function(rowContext) {
      return selected[rowContext.rowNumber] !== true && rowHasAgentFillContext(rowContext);
    })
    .slice(0, limit)
    .map(function(rowContext) {
      return {
        rowId: rowContext.rowId,
        rowNumber: rowContext.rowNumber,
        example: true,
        cells: rowContext.cells
      };
    });
}

function pickAgentContinueTableExampleRows(rowContexts, selectedRows, limit) {
  var selected = {};
  selectedRows.forEach(function(row) { selected[row] = true; });
  var firstSelectedRow = selectedRows[0] || 0;
  var examples = rowContexts
    .filter(function(rowContext) {
      return rowContext.rowNumber < firstSelectedRow &&
        selected[rowContext.rowNumber] !== true &&
        rowHasAgentFillContext(rowContext);
    })
    .slice(-limit)
    .map(function(rowContext) {
      return {
        rowId: rowContext.rowId,
        rowNumber: rowContext.rowNumber,
        example: true,
        cells: rowContext.cells
      };
    });

  return examples;
}

function uniqueAgentFillFields(targets, table) {
  var seen = {};
  var fields = [];
  targets.forEach(function(target) {
    if (seen[target.fieldId]) return;
    seen[target.fieldId] = true;
    var field = table.fieldByColumn[target.column] || createAgentFillFallbackField(target.column);
    fields.push(field);
  });
  return fields;
}

function createAgentFillFallbackField(col) {
  return {
    fieldId: normalizeAgentColumnKey(formatColumnLabel(col), {}),
    title: formatColumnLabel(col),
    column: col,
    columnLetter: formatColumnLabel(col),
    hasHeader: false
  };
}

function countUniqueAgentFillRows(targets) {
  return uniqueStrings((targets || []).map(function(target) { return target.rowId; })).length;
}

function createAgentFillRowId(row) {
  return 'row_' + row;
}

function readSheetValues(sheet, row, col, numRows, numCols) {
  if (numRows <= 0 || numCols <= 0) {
    return [];
  }
  return sheet.getRange(row, col, numRows, numCols).getValues();
}

function createBlankMatrix(numRows, numCols) {
  var rows = [];
  for (var r = 0; r < numRows; r++) {
    var row = [];
    for (var c = 0; c < numCols; c++) {
      row.push('');
    }
    rows.push(row);
  }
  return rows;
}

function uniqueNumbers(values) {
  var seen = {};
  var result = [];
  values.forEach(function(value) {
    var numberValue = Number(value);
    if (!numberValue || seen[numberValue]) return;
    seen[numberValue] = true;
    result.push(numberValue);
  });
  return result.sort(function(a, b) { return a - b; });
}

function uniqueStrings(values) {
  var seen = {};
  var result = [];
  values.forEach(function(value) {
    var stringValue = String(value || '');
    if (!stringValue || seen[stringValue]) return;
    seen[stringValue] = true;
    result.push(stringValue);
  });
  return result;
}

function chunkArray(values, chunkSize) {
  var chunks = [];
  for (var i = 0; i < values.length; i += chunkSize) {
    chunks.push(values.slice(i, i + chunkSize));
  }
  return chunks;
}

/**
 * Gets the current status of an Agent run.
 * @param {string} runId Exa Agent run ID
 * @return {Object} Run status result
 */
function getAgentRunStatus(runId) {
  const apiKey = getApiKey();
  if (!apiKey) {
    return { success: false, message: 'No API key set.' };
  }

  var fetched = fetchAgentRun(apiKey, runId);
  if (!fetched.success) {
    return fetched;
  }

  return {
    success: true,
    run: summarizeAgentRun(fetched.run)
  };
}

/**
 * Cancels a queued or running Agent run.
 * @param {string} runId Exa Agent run ID
 * @return {Object} Cancel result
 */
function cancelAgentRun(runId) {
  const apiKey = getApiKey();
  if (!apiKey) {
    return { success: false, message: 'No API key set.' };
  }

  if (!isValidAgentRunId(runId)) {
    return { success: false, message: 'Invalid Agent run ID.' };
  }

  try {
    var response = fetchWithRetry(EXA_AGENT_CONFIG.endpoint + '/' + encodeURIComponent(runId) + '/cancel', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({}),
      headers: {
        'x-api-key': apiKey,
        'Exa-Beta': EXA_AGENT_CONFIG.betaHeader,
        'x-exa-integration': 'exa-for-sheets',
        'User-Agent': 'exa-for-sheets 2.0'
      },
      muteHttpExceptions: true
    });

    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();

    if (responseCode === 200) {
      return {
        success: true,
        message: 'Agent run cancelled.',
        run: summarizeAgentRun(JSON.parse(responseBody))
      };
    }

    return {
      success: false,
      message: parseAgentApiError(responseCode, responseBody, 'Failed to cancel Agent run.')
    };
  } catch (e) {
    Logger.log('cancelAgentRun Error: ' + e);
    return { success: false, message: 'Script Error: ' + e.message };
  }
}

/**
 * Writes a completed Agent run's structured output into the active sheet.
 * @param {string} runId Exa Agent run ID
 * @param {Object} tableConfig Returned from startAgentTableRun
 * @param {Object} writeOptions includeHeaders, includeSources, overwrite
 * @return {Object} Write result
 */
function writeAgentRunToSheet(runId, tableConfig, writeOptions) {
  const apiKey = getApiKey();
  if (!apiKey) {
    return { success: false, message: 'No API key set.' };
  }

  var fetched = fetchAgentRun(apiKey, runId);
  if (!fetched.success) {
    return fetched;
  }

  var run = fetched.run;
  if (run.status !== 'completed') {
    return {
      success: false,
      message: 'Agent run is not complete yet. Current status: ' + run.status
    };
  }

  writeOptions = writeOptions || {};
  var tableResult = buildAgentSheetValues(
    run,
    tableConfig || {},
    writeOptions.includeHeaders !== false,
    writeOptions.includeSources === true
  );

  if (!tableResult.success) {
    return tableResult;
  }

  try {
    var sheetResult = getAgentWriteSheet(tableConfig || {});
    if (!sheetResult.success) {
      return sheetResult;
    }

    var sheet = sheetResult.sheet;
    var startRange = null;
    var requestedStartCell = normalizeAgentStartCell(
      tableConfig && tableConfig.startCell ? tableConfig.startCell : writeOptions.startCell
    );
    if (requestedStartCell) {
      if (!isValidAgentStartCell(requestedStartCell)) {
        return {
          success: false,
          message: 'Start cell must be a single cell like A1 or D5.'
        };
      }
      startRange = sheet.getRange(requestedStartCell);
    } else {
      startRange = sheet.getActiveRange();
    }
    if (!startRange) {
      startRange = sheet.getRange(1, 1);
    }

    var startRow = startRange.getRow();
    var startCol = startRange.getColumn();
    var values = tableResult.values;
    var numRows = values.length;
    var numCols = values[0].length;

    ensureSheetSize(sheet, startRow + numRows - 1, startCol + numCols - 1);

    var targetRange = sheet.getRange(startRow, startCol, numRows, numCols);
    if (writeOptions.overwrite !== true) {
      var existing = targetRange.getValues();
      var occupiedCount = countNonEmptyCells(existing);
      if (occupiedCount > 0) {
        return {
          success: false,
          needsConfirmation: true,
          code: 'TARGET_NOT_EMPTY',
          message: 'The target range ' + formatRangeA1(startRow, startCol, numRows, numCols) + ' contains ' + occupiedCount + ' non-empty cell(s).',
          rows: numRows,
          columns: numCols
        };
      }
    }

    targetRange.setValues(values);
    styleAgentTable(sheet, startRow, startCol, numRows, numCols, writeOptions.includeHeaders !== false);
    SpreadsheetApp.flush();

    return {
      success: true,
      message: 'Wrote table starting at ' + formatCellA1(startRow, startCol) + '.',
      runId: run.id,
      rows: tableResult.dataRows,
      columns: numCols,
      costDollars: run.costDollars || null,
      usage: run.usage || null
    };
  } catch (e) {
    Logger.log('writeAgentRunToSheet Error: ' + e);
    return { success: false, message: 'Operation failed: ' + e.message };
  }
}

/**
 * Fetches a run from the Agent API.
 * @param {string} apiKey Exa API key
 * @param {string} runId Agent run ID
 * @return {Object} Fetch result
 */
function fetchAgentRun(apiKey, runId) {
  if (!isValidAgentRunId(runId)) {
    return { success: false, message: 'Invalid Agent run ID.' };
  }

  try {
    var response = UrlFetchApp.fetch(EXA_AGENT_CONFIG.endpoint + '/' + encodeURIComponent(runId), {
      method: 'get',
      headers: {
        'x-api-key': apiKey,
        'Exa-Beta': EXA_AGENT_CONFIG.betaHeader,
        'x-exa-integration': 'exa-for-sheets',
        'User-Agent': 'exa-for-sheets 2.0'
      },
      muteHttpExceptions: true
    });

    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();

    if (responseCode === 200) {
      return { success: true, run: JSON.parse(responseBody) };
    }

    if (responseCode === 429) {
      return {
        success: false,
        retryable: true,
        retryAfterMs: getRetryAfterMs(response),
        message: 'Exa Agent is rate limited. Waiting before checking again.'
      };
    }

    return {
      success: false,
      message: parseAgentApiError(responseCode, responseBody, 'Failed to fetch Agent run.')
    };
  } catch (e) {
    Logger.log('fetchAgentRun Error: ' + e);
    return { success: false, message: 'Script Error: ' + e.message };
  }
}

/**
 * Creates a compact run object for sidebar polling.
 * @param {Object} run Full Agent run
 * @return {Object} Summary
 */
function summarizeAgentRun(run) {
  run = run || {};
  return {
    id: run.id || '',
    status: run.status || '',
    stopReason: run.stopReason || null,
    createdAt: run.createdAt || null,
    completedAt: run.completedAt || null,
    outputText: run.output && run.output.text ? run.output.text : '',
    hasStructured: !!(run.output && run.output.structured),
    rowCount: countAgentStructuredRows(run, 'rows'),
    usage: run.usage || null,
    costDollars: run.costDollars || null
  };
}

/**
 * Builds a 2D values matrix from Agent structured output.
 * @param {Object} run Agent run
 * @param {Object} tableConfig Table config
 * @param {boolean} includeHeaders Whether to include headers
 * @param {boolean} includeSources Whether to append a Sources column from grounding
 * @return {Object} Values result
 */
function buildAgentSheetValues(run, tableConfig, includeHeaders, includeSources) {
  var structured = run && run.output ? run.output.structured : null;
  if (!structured) {
    return { success: false, message: 'Agent run completed, but no structured output was returned.' };
  }

  var rootKey = tableConfig && tableConfig.rootKey ? tableConfig.rootKey : 'rows';
  var records = null;
  var actualRootKey = rootKey;

  if (Array.isArray(structured)) {
    records = structured;
    actualRootKey = '';
  } else if (Array.isArray(structured[rootKey])) {
    records = structured[rootKey];
  } else {
    Object.keys(structured).some(function(key) {
      if (Array.isArray(structured[key])) {
        records = structured[key];
        actualRootKey = key;
        return true;
      }
      return false;
    });
  }

  if (!records || records.length === 0) {
    return { success: false, message: 'Agent structured output did not contain any table rows.' };
  }

  var columns = normalizeAgentTableConfigColumns(tableConfig, records[0], structured.columns);
  if (columns.length === 0) {
    return { success: false, message: 'Could not infer table columns from Agent output.' };
  }

  var values = [];
  if (includeHeaders) {
    var headerRow = columns.map(function(column) { return column.label || humanizeAgentColumnLabel(column.key); });
    if (includeSources) {
      headerRow.push('Sources');
    }
    values.push(headerRow);
  }

  var dataValues = [];
  records.forEach(function(record, rowIndex) {
    var row = columns.map(function(column, columnIndex) {
      return formatAgentCellValue(getAgentRecordValue(record, column, columnIndex));
    });

    if (includeSources) {
      row.push(getAgentGroundingSourcesForRow(run, actualRootKey, rowIndex));
    }

    if (rowHasAgentValue(row)) {
      dataValues.push(row);
    }
  });

  if (dataValues.length === 0) {
    return {
      success: false,
      message: 'Agent returned column headers but no usable data rows. Try running again, or open More options and add the columns you want.'
    };
  }

  values = values.concat(dataValues);

  return {
    success: true,
    values: values,
    dataRows: dataValues.length,
    dataColumns: columns.length + (includeSources ? 1 : 0)
  };
}

/**
 * Normalizes stored table columns or infers them from a record.
 * @param {Object} tableConfig Table config
 * @param {Object} firstRecord First output record
 * @return {Object[]} Columns
 */
function normalizeAgentTableConfigColumns(tableConfig, firstRecord, generatedColumnLabels) {
  if (tableConfig && Array.isArray(tableConfig.columns) && tableConfig.columns.length > 0) {
    return tableConfig.columns
      .filter(function(column) { return column && column.key; })
      .map(function(column) {
        return {
          key: column.key,
          label: column.label || humanizeAgentColumnLabel(column.key),
          type: column.type || 'string'
        };
      });
  }

  if (Array.isArray(generatedColumnLabels) && generatedColumnLabels.length > 0) {
    return generatedColumnLabels
      .map(function(label) { return String(label || '').trim(); })
      .filter(function(label) { return label.length > 0; })
      .map(function(label) {
        return {
          key: label,
          label: label,
          type: 'string'
        };
      });
  }

  if (!firstRecord || typeof firstRecord !== 'object' || Array.isArray(firstRecord)) {
    return [];
  }

  return Object.keys(firstRecord).map(function(key) {
    return {
      key: key,
      label: humanizeAgentColumnLabel(key),
      type: Array.isArray(firstRecord[key]) ? 'array' : 'string'
    };
  });
}

/**
 * Reads a cell value from an Agent row, allowing generated labels to differ in casing or separators.
 * @param {Object|Array} record Agent row object
 * @param {Object} column Column definition
 * @param {number} columnIndex Column index
 * @return {*} Raw cell value
 */
function getAgentRecordValue(record, column, columnIndex) {
  if (Array.isArray(record)) {
    return record[columnIndex];
  }

  if (!record || typeof record !== 'object' || !column) {
    return '';
  }

  if (Array.isArray(record.cells)) {
    return record.cells[columnIndex];
  }

  if (Object.prototype.hasOwnProperty.call(record, column.key)) {
    return record[column.key];
  }

  var candidates = [normalizeAgentColumnKey(column.key, {})];
  if (column.label) {
    candidates.push(normalizeAgentColumnKey(column.label, {}));
  }

  var keys = Object.keys(record);
  for (var i = 0; i < keys.length; i++) {
    var normalizedKey = normalizeAgentColumnKey(keys[i], {});
    if (candidates.indexOf(normalizedKey) !== -1) {
      return record[keys[i]];
    }
  }

  return '';
}

/**
 * Checks whether a rendered row has at least one real value.
 * @param {Array} row Rendered row values
 * @return {boolean} True when the row has content
 */
function rowHasAgentValue(row) {
  return row.some(function(value) {
    if (value === null || value === undefined) return false;
    return String(value).trim() !== '';
  });
}

/**
 * Counts rows in structured Agent output.
 * @param {Object} run Agent run
 * @param {string} rootKey Expected root array key
 * @return {number} Row count
 */
function countAgentStructuredRows(run, rootKey) {
  var structured = run && run.output ? run.output.structured : null;
  if (!structured) return 0;
  if (Array.isArray(structured)) return structured.length;
  if (rootKey && Array.isArray(structured[rootKey])) return structured[rootKey].length;

  var count = 0;
  Object.keys(structured).some(function(key) {
    if (Array.isArray(structured[key])) {
      count = structured[key].length;
      return true;
    }
    return false;
  });
  return count;
}

/**
 * Formats a structured value for a Google Sheets cell.
 * @param {*} value Cell value
 * @return {string|number|boolean} Sheets-compatible value
 */
function formatAgentCellValue(value) {
  if (value === null || value === undefined) {
    return '';
  }

  if (Array.isArray(value)) {
    return value.map(function(item) {
      return (item && typeof item === 'object') ? JSON.stringify(item) : String(item);
    }).join(', ');
  }

  if (typeof value === 'object') {
    return JSON.stringify(value);
  }

  return value;
}

/**
 * Gets unique citation URLs for a row from Agent grounding metadata.
 * @param {Object} run Agent run
 * @param {string} rootKey Structured output root key
 * @param {number} rowIndex Row index
 * @return {string} Comma-separated URLs
 */
function getAgentGroundingSourcesForRow(run, rootKey, rowIndex) {
  var grounding = run && run.output && Array.isArray(run.output.grounding) ? run.output.grounding : [];
  if (!rootKey) {
    return '';
  }

  var prefix = 'structured.' + rootKey + '[' + rowIndex + ']';
  var seen = {};
  var urls = [];

  grounding.forEach(function(item) {
    if (!item || typeof item.field !== 'string' || item.field.indexOf(prefix) !== 0) {
      return;
    }

    (item.citations || []).forEach(function(citation) {
      if (citation && citation.url && !seen[citation.url]) {
        seen[citation.url] = true;
        urls.push(citation.url);
      }
    });
  });

  return urls.join(', ');
}

/**
 * Parses an API error response into a user-facing message.
 * @param {number} responseCode HTTP status
 * @param {string} responseBody Raw body
 * @param {string} fallback Fallback message
 * @return {string} User-facing error
 */
function parseAgentApiError(responseCode, responseBody, fallback) {
  var message = fallback + ' Status ' + responseCode + '.';
  try {
    var parsed = JSON.parse(responseBody);
    var detail = parsed.error || parsed.message || responseBody;
    message += ' Message: ' + (typeof detail === 'string' ? detail : JSON.stringify(detail));
  } catch (e) {
    if (responseBody) {
      message += ' Response: ' + responseBody;
    }
  }
  return message;
}

/**
 * Reads Retry-After from an HTTP response and converts it to milliseconds.
 * @param {HTTPResponse} response UrlFetchApp response
 * @return {number} Delay in milliseconds
 */
function getRetryAfterMs(response) {
  try {
    var headers = response && response.getHeaders ? response.getHeaders() : {};
    var retryAfter = headers['Retry-After'] || headers['retry-after'];
    if (!retryAfter) return 10000;

    var seconds = parseInt(retryAfter, 10);
    if (seconds && seconds > 0) {
      return Math.min(seconds * 1000, 60000);
    }

    var retryDate = new Date(retryAfter).getTime();
    if (!isNaN(retryDate)) {
      return Math.min(Math.max(retryDate - Date.now(), 1000), 60000);
    }
  } catch (e) {
    Logger.log('getRetryAfterMs skipped: ' + e);
  }

  return 10000;
}

/**
 * Validates an Agent run ID before placing it in a URL.
 * @param {string} runId Agent run ID
 * @return {boolean} True if valid
 */
function isValidAgentRunId(runId) {
  return typeof runId === 'string' && /^[A-Za-z0-9_.:-]+$/.test(runId);
}

/**
 * Normalizes a user-entered start cell.
 * @param {*} startCell Start cell from the sidebar
 * @return {string} Uppercase cell address or empty string
 */
function normalizeAgentStartCell(startCell) {
  return (typeof startCell === 'string') ? startCell.trim().toUpperCase() : '';
}

/**
 * Validates a simple A1-style single-cell address.
 * @param {string} startCell Cell address
 * @return {boolean} True if valid
 */
function isValidAgentStartCell(startCell) {
  return typeof startCell === 'string' && /^[A-Z]{1,3}[1-9][0-9]*$/.test(startCell);
}

/**
 * Reads the currently selected cell when the Agent run starts.
 * @return {string} Active cell address, or A1 if unavailable
 */
function getCurrentAgentStartCell(sheet) {
  try {
    sheet = sheet || SpreadsheetApp.getActiveSheet();
    var activeRange = sheet && sheet.getActiveRange ? sheet.getActiveRange() : null;
    if (activeRange && activeRange.getRow && activeRange.getColumn) {
      return formatCellA1(activeRange.getRow(), activeRange.getColumn());
    }
  } catch (e) {
    Logger.log('getCurrentAgentStartCell skipped: ' + e);
  }

  return 'A1';
}

/**
 * Captures the sheet identity for long-running Agent writes.
 * @param {Sheet} sheet Target sheet
 * @return {Object} Sheet identity
 */
function getAgentSheetContext(sheet) {
  return {
    sheetId: sheet && sheet.getSheetId ? sheet.getSheetId() : null,
    sheetName: sheet && sheet.getName ? sheet.getName() : ''
  };
}

/**
 * Finds the sheet that was active when an Agent run started.
 * Falls back to the active sheet only for older configs that did not store sheet identity.
 * @param {Object} config Table/fill write config
 * @return {Object} Result with sheet or user-facing error
 */
function getAgentWriteSheet(config) {
  config = config || {};
  var sheetId = config.sheetId === null || config.sheetId === undefined ? null : Number(config.sheetId);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet ? SpreadsheetApp.getActiveSpreadsheet() : null;

  if (sheetId !== null && !isNaN(sheetId) && spreadsheet && spreadsheet.getSheets) {
    var sheets = spreadsheet.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId && Number(sheets[i].getSheetId()) === sheetId) {
        return { success: true, sheet: sheets[i] };
      }
    }

    return {
      success: false,
      message: 'The sheet used for this Agent run could not be found. Run it again from the sheet you want to update.'
    };
  }

  if (config.sheetName && spreadsheet && spreadsheet.getSheetByName) {
    var namedSheet = spreadsheet.getSheetByName(config.sheetName);
    if (namedSheet) {
      return { success: true, sheet: namedSheet };
    }
  }

  var activeSheet = SpreadsheetApp.getActiveSheet ? SpreadsheetApp.getActiveSheet() : null;
  if (activeSheet) {
    return { success: true, sheet: activeSheet };
  }

  return {
    success: false,
    message: 'Could not find the sheet to write to.'
  };
}

/**
 * Counts non-empty values in a 2D range array.
 * @param {Array[]} values Range values
 * @return {number} Count
 */
function countNonEmptyCells(values) {
  var count = 0;
  values.forEach(function(row) {
    row.forEach(function(value) {
      if (value !== '' && value !== null && value !== undefined) {
        count++;
      }
    });
  });
  return count;
}

/**
 * Expands the active sheet when the table would exceed current bounds.
 * @param {Sheet} sheet Active sheet
 * @param {number} requiredRows Last required row
 * @param {number} requiredCols Last required column
 */
function ensureSheetSize(sheet, requiredRows, requiredCols) {
  if (sheet.getMaxRows && sheet.insertRowsAfter && requiredRows > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }

  if (sheet.getMaxColumns && sheet.insertColumnsAfter && requiredCols > sheet.getMaxColumns()) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredCols - sheet.getMaxColumns());
  }
}

/**
 * Applies light header styling when supported by the Sheets runtime.
 * @param {Sheet} sheet Active sheet
 * @param {number} startRow Start row
 * @param {number} startCol Start column
 * @param {number} numRows Number of rows
 * @param {number} numCols Number of columns
 * @param {boolean} includeHeaders Whether headers were written
 */
function styleAgentTable(sheet, startRow, startCol, numRows, numCols, includeHeaders) {
  try {
    if (includeHeaders && sheet.getRange) {
      var headerRange = sheet.getRange(startRow, startCol, 1, numCols);
      if (headerRange.setFontWeight) headerRange.setFontWeight('bold');
      if (headerRange.setBackground) headerRange.setBackground('#EFF6FF');
    }

    if (sheet.autoResizeColumns) {
      sheet.autoResizeColumns(startCol, numCols);
    }
  } catch (e) {
    Logger.log('styleAgentTable skipped: ' + e);
  }
}

/**
 * Formats a cell coordinate as A1 notation.
 * @param {number} row Row number
 * @param {number} col Column number
 * @return {string} A1 coordinate
 */
function formatCellA1(row, col) {
  return formatColumnLabel(col) + row;
}

/**
 * Formats a one-based column number as A1-style letters.
 * @param {number} col Column number
 * @return {string} Column letters
 */
function formatColumnLabel(col) {
  var letters = '';
  var current = col;
  while (current > 0) {
    var remainder = (current - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    current = Math.floor((current - 1) / 26);
  }
  return letters;
}

/**
 * Formats a range coordinate as A1 notation.
 * @param {number} row Start row
 * @param {number} col Start column
 * @param {number} numRows Number of rows
 * @param {number} numCols Number of columns
 * @return {string} A1 range
 */
function formatRangeA1(row, col, numRows, numCols) {
  return formatCellA1(row, col) + ':' + formatCellA1(row + numRows - 1, col + numCols - 1);
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
        // Match =EXA( or =EXA_ to include both simplified EXA() and EXA_ANSWER, EXA_SEARCH, etc.
        if (formula && formula.toUpperCase().match(/^=EXA[_(]/)) {
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

/**
 * Converts selected cells containing Exa functions to their static values.
 * This prevents automatic recalculation and unexpected API charges.
 * The formulas are replaced with their current values, so they won't refresh.
 *
 * @return {Object} Result object with success flag and message
 */
function convertToValues() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const selection = sheet.getActiveRange();

    if (!selection) {
      return {
        success: false,
        message: 'No cells selected. Please select cells containing Exa functions.'
      };
    }

    const formulas = selection.getFormulas();
    const values = selection.getValues();
    const exaCells = [];

    formulas.forEach((row, rowIndex) => {
      row.forEach((formula, colIndex) => {
        // Match =EXA( or =EXA_ to include both simplified EXA() and EXA_ANSWER, EXA_SEARCH, etc.
        if (formula && formula.toUpperCase().match(/^=EXA[_(]/)) {
          exaCells.push({
            row: rowIndex,
            col: colIndex,
            formula: formula,
            value: values[rowIndex][colIndex]
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

    // Replace formulas with their values
    exaCells.forEach(item => {
      const cell = selection.getCell(item.row + 1, item.col + 1);
      cell.setValue(item.value);
    });

    SpreadsheetApp.flush();

    const cellText = exaCells.length === 1 ? 'cell' : 'cells';
    return {
      success: true,
      message: `Converted ${exaCells.length} ${cellText} to static values. These cells will no longer auto-refresh.`
    };

  } catch (e) {
    Logger.log(`Error in convertToValues: ${e}`);
    return {
      success: false,
      message: `Operation failed: ${e.message}`
    };
  }
}
