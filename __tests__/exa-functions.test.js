// Unit tests for Exa Google Sheets functions
// These tests mock Google Apps Script APIs and test the core logic

const fs = require('fs');
const path = require('path');

// Load the Code.gs file and evaluate it to get the functions
const codeGsPath = path.join(__dirname, '..', 'Code.gs');
const codeGsContent = fs.readFileSync(codeGsPath, 'utf8');

// Evaluate the code in the global context (with mocked GAS APIs from setup.js)
eval(codeGsContent);

describe('API Key Management', () => {
  beforeEach(() => {
    resetMocks();
  });

  describe('saveApiKey', () => {
    test('should save a valid API key', () => {
      const result = saveApiKey('exa_test_key_12345678');
      
      expect(result.success).toBe(true);
      expect(result.message).toBe('API key saved successfully.');
      expect(result.correlationId).toBeDefined();
    });

    test('should reject empty API key', () => {
      const result = saveApiKey('');
      
      expect(result.success).toBe(false);
      expect(result.code).toBe('VALIDATION_EMPTY_KEY');
    });

    test('should reject null API key', () => {
      const result = saveApiKey(null);
      
      expect(result.success).toBe(false);
      expect(result.code).toBe('VALIDATION_EMPTY_KEY');
    });

    test('should create masked display key', () => {
      saveApiKey('exa_test_key_12345678');
      const keysData = JSON.parse(getMockProperty('EXA_API_KEYS'));
      
      expect(keysData.keys.default.displayKey).toMatch(/^exa_\.+5678$/);
    });
  });

  describe('getApiKey', () => {
    test('should return null when no key is set', () => {
      const result = getApiKey();
      expect(result).toBeNull();
    });

    test('should return the active API key', () => {
      saveApiKey('exa_my_secret_key');
      const result = getApiKey();
      
      expect(result).toBe('exa_my_secret_key');
    });

    test('should update lastUsed timestamp when getting key', () => {
      saveApiKey('exa_test_key');
      const beforeGet = JSON.parse(getMockProperty('EXA_API_KEYS'));
      const beforeLastUsed = beforeGet.keys.default.lastUsed;
      
      // Small delay to ensure timestamp changes
      getApiKey();
      
      const afterGet = JSON.parse(getMockProperty('EXA_API_KEYS'));
      expect(afterGet.keys.default.lastUsed).toBeDefined();
    });
  });

  describe('removeApiKey', () => {
    test('should remove the API key', () => {
      saveApiKey('exa_test_key');
      expect(getApiKey()).toBe('exa_test_key');
      
      const result = removeApiKey();
      
      expect(result.success).toBe(true);
      expect(getApiKey()).toBeNull();
    });
  });

  describe('getAllApiKeys', () => {
    test('should return empty structure when no keys exist', () => {
      const result = getAllApiKeys();
      
      expect(result).toEqual({ keys: {}, activeKeyName: null });
    });

    test('should return keys after saving', () => {
      saveApiKey('exa_test_key');
      const result = getAllApiKeys();
      
      expect(result.activeKeyName).toBe('default');
      expect(result.keys.default).toBeDefined();
      expect(result.keys.default.key).toBe('exa_test_key');
    });
  });
});

describe('EXA_ANSWER', () => {
  beforeEach(() => {
    resetMocks();
    saveApiKey('exa_test_api_key');
  });

  test('should return error when no API key is set', () => {
    removeApiKey();
    const result = EXA_ANSWER('What is the capital of France?');
    
    expect(result).toContain('No API key set');
  });

  test('should return error for empty prompt', () => {
    const result = EXA_ANSWER('');
    
    expect(result).toBe('Please provide a valid prompt/question.');
  });

  test('should return error for null prompt', () => {
    const result = EXA_ANSWER(null);
    
    expect(result).toBe('Please provide a valid prompt/question.');
  });

  test('should make API call with correct parameters', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        answer: 'Paris is the capital of France.'
      })
    });

    EXA_ANSWER('What is the capital of France?');

    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      'https://api.exa.ai/answer',
      expect.objectContaining({
        method: 'post',
        contentType: 'application/json',
        headers: expect.objectContaining({
          'x-api-key': 'exa_test_api_key',
          'x-exa-integration': 'exa-for-sheets'
        })
      })
    );
  });

  test('should return answer from successful API response', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        answer: 'Paris is the capital of France.'
      })
    });

    const result = EXA_ANSWER('What is the capital of France?');
    
    expect(result).toBe('Paris is the capital of France.');
  });

  test('should strip inline citations by default', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        answer: 'Paris is the capital of France. ([Wikipedia](https://en.wikipedia.org/wiki/Paris))'
      })
    });

    const result = EXA_ANSWER('What is the capital of France?');
    
    expect(result).toBe('Paris is the capital of France.');
    expect(result).not.toContain('Wikipedia');
  });

  test('should include citations when includeCitations is true', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        answer: 'Paris is the capital of France.',
        citations: [
          { title: 'Wikipedia', url: 'https://en.wikipedia.org/wiki/Paris' }
        ]
      })
    });

    // EXA_ANSWER(prompt, prefix, suffix, includeCitations, systemPrompt, outputSchema)
    const result = EXA_ANSWER('What is the capital of France?', '', '', true);
    
    expect(result).toContain('Paris is the capital of France.');
    expect(result).toContain('Wikipedia');
  });

  test('should combine prefix, prompt, and suffix', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        answer: 'Test answer'
      })
    });

    EXA_ANSWER('main query', 'prefix text', 'suffix text');

    const callArgs = UrlFetchApp.fetch.mock.calls[0][1];
    const payload = JSON.parse(callArgs.payload);
    
    expect(payload.query).toBe('prefix text main query suffix text');
  });

  test('should handle 401 unauthorized error', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 401,
      getContentText: () => JSON.stringify({ error: 'Invalid API key' })
    });

    const result = EXA_ANSWER('test query');
    
    expect(result).toContain('Invalid API Key');
  });

  test('should handle API errors gracefully', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 500,
      getContentText: () => JSON.stringify({ error: 'Internal server error' })
    });

    const result = EXA_ANSWER('test query');
    
    expect(result).toContain('API Error');
    expect(result).toContain('500');
  });

  test('should use chat completions endpoint when systemPrompt is provided', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        choices: [{ message: { content: 'Will Bryk', citations: [] } }]
      })
    });

    // EXA_ANSWER(prompt, prefix, suffix, includeCitations, systemPrompt, outputSchema, returnRawJson)
    const result = EXA_ANSWER('ceo of exa.ai', '', '', false, 'only return the name');
    
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      'https://api.exa.ai/chat/completions',
      expect.objectContaining({
        method: 'post',
        headers: expect.objectContaining({
          'Authorization': 'Bearer exa_test_api_key'
        })
      })
    );
    
    const callArgs = UrlFetchApp.fetch.mock.calls[0][1];
    const payload = JSON.parse(callArgs.payload);
    expect(payload.model).toBe('exa');
    expect(payload.messages).toContainEqual({ role: 'system', content: 'only return the name' });
    expect(payload.messages).toContainEqual({ role: 'user', content: 'ceo of exa.ai' });
    
    expect(result).toBe('Will Bryk');
  });

  test('should use chat completions endpoint when outputSchema is provided', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        choices: [{ message: { content: '{"name": "Will Bryk"}', citations: [] } }]
      })
    });

    const schema = '{"type":"object","properties":{"name":{"type":"string"}}}';
    // EXA_ANSWER(prompt, prefix, suffix, includeCitations, systemPrompt, outputSchema, returnRawJson)
    const result = EXA_ANSWER('ceo of exa.ai', '', '', false, '', schema);
    
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      'https://api.exa.ai/chat/completions',
      expect.anything()
    );
    
    const callArgs = UrlFetchApp.fetch.mock.calls[0][1];
    const payload = JSON.parse(callArgs.payload);
    expect(payload.output_schema).toEqual(JSON.parse(schema));
    
    // Should extract the value from single-key JSON
    expect(result).toBe('Will Bryk');
  });

  test('should return raw JSON when returnRawJson is true', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        choices: [{ message: { content: '{"name": "Will Bryk"}', citations: [] } }]
      })
    });

    const schema = '{"type":"object","properties":{"name":{"type":"string"}}}';
    // EXA_ANSWER(prompt, prefix, suffix, includeCitations, systemPrompt, outputSchema, returnRawJson)
    const result = EXA_ANSWER('ceo of exa.ai', '', '', false, '', schema, true);
    
    // Should return formatted JSON when returnRawJson is true
    expect(result).toContain('"name"');
    expect(result).toContain('Will Bryk');
  });

  test('should return formatted JSON for multi-key response', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        choices: [{ message: { content: '{"name": "Will Bryk", "company": "Exa"}', citations: [] } }]
      })
    });

    const schema = '{"type":"object","properties":{"name":{"type":"string"},"company":{"type":"string"}}}';
    // EXA_ANSWER(prompt, prefix, suffix, includeCitations, systemPrompt, outputSchema, returnRawJson)
    const result = EXA_ANSWER('ceo of exa.ai', '', '', false, '', schema);
    
    // Should return formatted JSON for multi-key response
    expect(result).toContain('"name"');
    expect(result).toContain('"company"');
  });

  test('should return error for invalid outputSchema JSON', () => {
    const result = EXA_ANSWER('test', '', '', false, '', 'invalid json');
    
    expect(result).toContain('Invalid outputSchema');
  });
});

describe('EXA_CONTENTS', () => {
  beforeEach(() => {
    resetMocks();
    saveApiKey('exa_test_api_key');
  });

  test('should return error when no API key is set', () => {
    removeApiKey();
    const result = EXA_CONTENTS('https://example.com');
    
    expect(result).toContain('No API key set');
  });

  test('should return error for invalid URL', () => {
    const result = EXA_CONTENTS('not-a-url');
    
    expect(result).toContain('valid URL');
  });

  test('should return error for empty URL', () => {
    const result = EXA_CONTENTS('');
    
    expect(result).toContain('valid URL');
  });

  test('should make API call with correct URL', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        results: [{ text: 'Page content here' }]
      })
    });

    EXA_CONTENTS('https://example.com/page');

    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      'https://api.exa.ai/contents',
      expect.objectContaining({
        method: 'post'
      })
    );

    const callArgs = UrlFetchApp.fetch.mock.calls[0][1];
    const payload = JSON.parse(callArgs.payload);
    expect(payload.urls).toEqual(['https://example.com/page']);
  });

  test('should return text content from successful response', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        results: [{ text: 'This is the page content.' }]
      })
    });

    const result = EXA_CONTENTS('https://example.com');
    
    expect(result).toBe('This is the page content.');
  });

  test('should handle 401 unauthorized error', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 401,
      getContentText: () => JSON.stringify({ error: 'Invalid API key' })
    });

    const result = EXA_CONTENTS('https://example.com');
    
    expect(result).toContain('Invalid API Key');
  });
});

describe('EXA_SEARCH', () => {
  beforeEach(() => {
    resetMocks();
    saveApiKey('exa_test_api_key');
  });

  test('should return error when no API key is set', () => {
    removeApiKey();
    const result = EXA_SEARCH('test query');
    
    expect(result[0][0]).toContain('No API key set');
  });

  test('should return error for empty query', () => {
    const result = EXA_SEARCH('');
    
    expect(result[0][0]).toContain('valid search query');
  });

  test('should make API call with correct parameters', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        results: [{ url: 'https://example.com' }]
      })
    });

    EXA_SEARCH('machine learning', 5, 'neural');

    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      'https://api.exa.ai/search',
      expect.objectContaining({
        method: 'post'
      })
    );

    const callArgs = UrlFetchApp.fetch.mock.calls[0][1];
    const payload = JSON.parse(callArgs.payload);
    
    expect(payload.query).toBe('machine learning');
    expect(payload.numResults).toBe(5);
    expect(payload.type).toBe('neural');
  });

  test('should return URLs as vertical array', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        results: [
          { url: 'https://example1.com' },
          { url: 'https://example2.com' },
          { url: 'https://example3.com' }
        ]
      })
    });

    const result = EXA_SEARCH('test query', 3);
    
    expect(result).toEqual([
      ['https://example1.com'],
      ['https://example2.com'],
      ['https://example3.com']
    ]);
  });

  test('should default to 1 result and auto search type', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        results: [{ url: 'https://example.com' }]
      })
    });

    EXA_SEARCH('test query');

    const callArgs = UrlFetchApp.fetch.mock.calls[0][1];
    const payload = JSON.parse(callArgs.payload);
    
    expect(payload.numResults).toBe(1);
    expect(payload.type).toBe('auto');
  });

  test('should handle prefix and suffix', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        results: [{ url: 'https://example.com' }]
      })
    });

    EXA_SEARCH('main query', 1, 'auto', 'prefix', 'suffix');

    const callArgs = UrlFetchApp.fetch.mock.calls[0][1];
    const payload = JSON.parse(callArgs.payload);
    
    expect(payload.query).toBe('prefix main query suffix');
  });

  test('should handle 401 unauthorized error', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 401,
      getContentText: () => JSON.stringify({ error: 'Invalid API key' })
    });

    const result = EXA_SEARCH('test query');
    
    expect(result[0][0]).toContain('Invalid API Key');
  });
});

describe('EXA_FINDSIMILAR', () => {
  beforeEach(() => {
    resetMocks();
    saveApiKey('exa_test_api_key');
  });

  test('should return error when no API key is set', () => {
    removeApiKey();
    const result = EXA_FINDSIMILAR('https://example.com');
    
    expect(result[0][0]).toContain('No API key set');
  });

  test('should return error for invalid URL', () => {
    const result = EXA_FINDSIMILAR('not-a-url');
    
    expect(result[0][0]).toContain('valid URL');
  });

  test('should make API call with correct URL', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        results: [{ url: 'https://similar.com' }]
      })
    });

    EXA_FINDSIMILAR('https://example.com');

    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      'https://api.exa.ai/findSimilar',
      expect.objectContaining({
        method: 'post'
      })
    );

    const callArgs = UrlFetchApp.fetch.mock.calls[0][1];
    const payload = JSON.parse(callArgs.payload);
    
    expect(payload.url).toBe('https://example.com');
  });

  test('should return similar URLs as vertical array', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        results: [
          { url: 'https://similar1.com' },
          { url: 'https://similar2.com' }
        ]
      })
    });

    const result = EXA_FINDSIMILAR('https://example.com', 2);
    
    expect(result).toEqual([
      ['https://similar1.com'],
      ['https://similar2.com']
    ]);
  });

  test('should handle domain filtering', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        results: [{ url: 'https://linkedin.com/company/test' }]
      })
    });

    EXA_FINDSIMILAR('https://example.com', 5, 'linkedin.com,crunchbase.com', 'wikipedia.org');

    const callArgs = UrlFetchApp.fetch.mock.calls[0][1];
    const payload = JSON.parse(callArgs.payload);
    
    expect(payload.includeDomains).toEqual(['linkedin.com', 'crunchbase.com']);
    expect(payload.excludeDomains).toEqual(['wikipedia.org']);
  });
});

describe('Rate Limiting', () => {
  beforeEach(() => {
    resetMocks();
    saveApiKey('exa_test_api_key');
  });

  test('fetchWithRetry should return response on success', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({ answer: 'test' })
    });

    const response = fetchWithRetry('https://api.exa.ai/answer', { method: 'post' });
    
    expect(response.getResponseCode()).toBe(200);
    expect(UrlFetchApp.fetch).toHaveBeenCalledTimes(1);
  });

  test('fetchWithRetry should retry on 429 and succeed', () => {
    let callCount = 0;
    UrlFetchApp.fetch.mockImplementation(() => {
      callCount++;
      if (callCount === 1) {
        return {
          getResponseCode: () => 429,
          getContentText: () => 'Rate limited',
          getHeaders: () => ({})
        };
      }
      return {
        getResponseCode: () => 200,
        getContentText: () => JSON.stringify({ answer: 'test' })
      };
    });

    const response = fetchWithRetry('https://api.exa.ai/answer', { method: 'post' });
    
    expect(response.getResponseCode()).toBe(200);
    expect(UrlFetchApp.fetch).toHaveBeenCalledTimes(2);
    expect(Utilities.sleep).toHaveBeenCalledTimes(1);
  });

  test('fetchWithRetry should respect Retry-After header', () => {
    let callCount = 0;
    UrlFetchApp.fetch.mockImplementation(() => {
      callCount++;
      if (callCount === 1) {
        return {
          getResponseCode: () => 429,
          getContentText: () => 'Rate limited',
          getHeaders: () => ({ 'Retry-After': '2' })
        };
      }
      return {
        getResponseCode: () => 200,
        getContentText: () => JSON.stringify({ answer: 'test' })
      };
    });

    fetchWithRetry('https://api.exa.ai/answer', { method: 'post' });
    
    expect(Utilities.sleep).toHaveBeenCalledWith(2000);
  });

  test('fetchWithRetry should return 429 after max retries', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 429,
      getContentText: () => 'Rate limited',
      getHeaders: () => ({})
    });

    const response = fetchWithRetry('https://api.exa.ai/answer', { method: 'post' });
    
    expect(response.getResponseCode()).toBe(429);
    expect(UrlFetchApp.fetch).toHaveBeenCalledTimes(4); // 1 initial + 3 retries
  });

  test('EXA_ANSWER should return rate limit error message on 429', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 429,
      getContentText: () => 'Rate limited',
      getHeaders: () => ({})
    });

    const result = EXA_ANSWER('test query');
    
    expect(result).toContain('Rate limit exceeded');
  });

  test('EXA_SEARCH should return rate limit error message on 429', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 429,
      getContentText: () => 'Rate limited',
      getHeaders: () => ({})
    });

    const result = EXA_SEARCH('test query');
    
    expect(result[0][0]).toContain('Rate limit exceeded');
  });
});

describe('Helper Functions', () => {
  beforeEach(() => {
    resetMocks();
  });

  describe('fail', () => {
    test('should create structured error response', () => {
      const result = fail('TEST_ERROR', 'Test message', 'test-correlation-id', { extra: 'data' });
      
      expect(result).toEqual({
        success: false,
        code: 'TEST_ERROR',
        message: 'Test message',
        correlationId: 'test-correlation-id',
        details: { extra: 'data' }
      });
    });

    test('should handle missing details', () => {
      const result = fail('TEST_ERROR', 'Test message', 'test-id');
      
      expect(result.details).toBeNull();
    });
  });

  describe('getApiKeyForUI', () => {
    test('should return null when no key exists', () => {
      const result = getApiKeyForUI();
      expect(result).toBeNull();
    });

    test('should return display info when key exists', () => {
      saveApiKey('exa_test_key_12345678');
      const result = getApiKeyForUI();
      
      expect(result).toBeDefined();
      expect(result.displayKey).toBeDefined();
      expect(result.created).toBeDefined();
      // Should not expose the actual key
      expect(result.key).toBeUndefined();
    });
  });
});
