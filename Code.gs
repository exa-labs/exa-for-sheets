// Rate limiting configuration
var RATE_LIMIT_CONFIG = {
  maxRetries: 3,
  initialDelayMs: 1000,
  maxDelayMs: 10000
};

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
    
    // Check for Retry-After header
    var retryAfter = response.getHeaders()['Retry-After'];
    if (retryAfter) {
      delay = parseInt(retryAfter, 10) * 1000;
    }
    
    // Cap the delay
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
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Exa AI');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showAbout() {
  var ui = SpreadsheetApp.getUi();
  var message = 'Exa AI for Google Sheets\n\n' +
                'Version: 2.0.3\n\n' +
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
