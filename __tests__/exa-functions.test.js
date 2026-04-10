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

  test('should use chat completions endpoint with output_schema when outputSchema is provided', () => {
    // Chat completions response format with JSON string in content
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        choices: [{ message: { content: JSON.stringify({ name: 'Will Bryk' }), citations: [] } }]
      })
    });

    const schema = '{"type":"object","properties":{"name":{"type":"string"}}}';
    // EXA_ANSWER(prompt, prefix, suffix, includeCitations, systemPrompt, outputSchema, returnRawJson, type)
    const result = EXA_ANSWER('ceo of exa.ai', '', '', false, '', schema);
    
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      'https://api.exa.ai/chat/completions',
      expect.anything()
    );
    
    const callArgs = UrlFetchApp.fetch.mock.calls[0][1];
    const payload = JSON.parse(callArgs.payload);
    expect(payload.extraBody.outputSchema).toEqual(JSON.parse(schema));
    expect(payload.model).toBe('exa');
    
    // Should extract the value from single-key object
    expect(result).toBe('Will Bryk');
  });

  test('should return raw JSON when returnRawJson is true', () => {
    // Chat completions response format with JSON string in content
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        choices: [{ message: { content: JSON.stringify({ name: 'Will Bryk' }), citations: [] } }]
      })
    });

    const schema = '{"type":"object","properties":{"name":{"type":"string"}}}';
    // EXA_ANSWER(prompt, prefix, suffix, includeCitations, systemPrompt, outputSchema, returnRawJson, type)
    const result = EXA_ANSWER('ceo of exa.ai', '', '', false, '', schema, true);
    
    // Should return formatted JSON when returnRawJson is true
    expect(result).toContain('"name"');
    expect(result).toContain('Will Bryk');
  });

  test('should return formatted JSON for multi-key response', () => {
    // Chat completions response format with JSON string in content
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        choices: [{ message: { content: JSON.stringify({ name: 'Will Bryk', company: 'Exa' }), citations: [] } }]
      })
    });

    const schema = '{"type":"object","properties":{"name":{"type":"string"},"company":{"type":"string"}}}';
    // EXA_ANSWER(prompt, prefix, suffix, includeCitations, systemPrompt, outputSchema, returnRawJson, type)
    const result = EXA_ANSWER('ceo of exa.ai', '', '', false, '', schema);
    
    // Should return formatted JSON for multi-key response
    expect(result).toContain('"name"');
    expect(result).toContain('"company"');
  });

  test('should return JSON with all schema-defined properties for complex multi-property schema', () => {
    const complexAnswer = {
      id: 'exa-001',
      name: 'Exa AI',
      description: 'Neural search engine company',
      status: 'active',
      createdDate: '2022-01-15',
      lastModified: '2024-03-20',
      category: 'AI/ML',
      priority: 'high',
      tags: ['search', 'ai', 'neural'],
      metadata: {
        version: '2.0',
        author: 'Will Bryk'
      }
    };

    // Chat completions response format with JSON string in content
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        choices: [{ message: { content: JSON.stringify(complexAnswer), citations: [] } }]
      })
    });

    const schema = JSON.stringify({
      type: 'object',
      required: ['id', 'name', 'description', 'status', 'createdDate', 'lastModified', 'category', 'priority', 'tags', 'metadata'],
      additionalProperties: false,
      properties: {
        id: { type: 'string', description: 'Unique identifier' },
        name: { type: 'string', description: 'Display name or title' },
        description: { type: 'string', description: 'Detailed description' },
        status: { type: 'string', description: 'Current status' },
        createdDate: { type: 'string', description: 'Date when created' },
        lastModified: { type: 'string', description: 'Date when last modified' },
        category: { type: 'string', description: 'Category classification' },
        priority: { type: 'string', description: 'Priority level' },
        tags: { type: 'array', description: 'List of tags', items: { type: 'string' } },
        metadata: {
          type: 'object',
          description: 'Additional metadata',
          properties: {
            version: { type: 'string', description: 'Version number' },
            author: { type: 'string', description: 'Author or creator' }
          },
          additionalProperties: false
        }
      }
    });

    const result = EXA_ANSWER('describe exa.ai company', '', '', false, '', schema, true);
    const parsed = JSON.parse(result);

    expect(parsed.id).toBe('exa-001');
    expect(parsed.name).toBe('Exa AI');
    expect(parsed.description).toBe('Neural search engine company');
    expect(parsed.status).toBe('active');
    expect(parsed.createdDate).toBe('2022-01-15');
    expect(parsed.lastModified).toBe('2024-03-20');
    expect(parsed.category).toBe('AI/ML');
    expect(parsed.priority).toBe('high');
    expect(parsed.tags).toEqual(['search', 'ai', 'neural']);
    expect(parsed.metadata.version).toBe('2.0');
    expect(parsed.metadata.author).toBe('Will Bryk');
  });

  test('should return error for invalid outputSchema JSON', () => {
    const result = EXA_ANSWER('test', '', '', false, '', 'invalid json');
    
    expect(result).toContain('Invalid outputSchema');
  });

  test('should pass type parameter to /answer endpoint', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        answer: 'Deep answer result'
      })
    });

    // EXA_ANSWER(prompt, prefix, suffix, includeCitations, systemPrompt, outputSchema, returnRawJson, type)
    EXA_ANSWER('What is Exa?', '', '', false, '', '', false, 'deep');

    const callArgs = UrlFetchApp.fetch.mock.calls[0][1];
    const payload = JSON.parse(callArgs.payload);
    
    expect(payload.type).toBe('deep');
    expect(payload.query).toBe('What is Exa?');
  });

  test('should default to deep type when type is not provided', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        answer: 'Default answer'
      })
    });

    EXA_ANSWER('test query');

    const callArgs = UrlFetchApp.fetch.mock.calls[0][1];
    const payload = JSON.parse(callArgs.payload);
    
    expect(payload.type).toBe('deep');
  });

  test('should fall back to deep for invalid type parameter', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        answer: 'Answer'
      })
    });

    EXA_ANSWER('test', '', '', false, '', '', false, 'invalid_type');

    const callArgs = UrlFetchApp.fetch.mock.calls[0][1];
    const payload = JSON.parse(callArgs.payload);
    
    expect(payload.type).toBe('deep');
  });

  test('should pass type via extraBody when using chat completions', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        choices: [{ message: { content: 'Deep result', citations: [] } }]
      })
    });

    EXA_ANSWER('test query', '', '', false, 'be concise', '', false, 'deep');

    const callArgs = UrlFetchApp.fetch.mock.calls[0][1];
    const payload = JSON.parse(callArgs.payload);
    
    expect(payload.extraBody.type).toBe('deep');
  });

  test('should accept all valid type values', () => {
    const validTypes = ['auto', 'neural', 'fast', 'deep'];
    
    for (const type of validTypes) {
      resetMocks();
      saveApiKey('exa_test_api_key');
      
      UrlFetchApp.fetch.mockReturnValue({
        getResponseCode: () => 200,
        getContentText: () => JSON.stringify({ answer: `Answer for ${type}` })
      });

      EXA_ANSWER('test', '', '', false, '', '', false, type);

      const callArgs = UrlFetchApp.fetch.mock.calls[0][1];
      const payload = JSON.parse(callArgs.payload);
      expect(payload.type).toBe(type);
    }
  });
});

describe('EXA (simplified wrapper)', () => {
  beforeEach(() => {
    resetMocks();
    saveApiKey('exa_test_api_key');
  });

  test('should return error when no API key is set', () => {
    removeApiKey();
    const result = EXA('What is the CEO name?', 'Exa AI');
    
    expect(result).toContain('No API key set');
  });

  test('should use /search with systemPrompt, outputSchema, and highlights', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        output: { content: 'Will Bryk' },
        results: []
      })
    });

    EXA('Return only the CEO name', 'Exa AI');

    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      'https://api.exa.ai/search',
      expect.objectContaining({ method: 'post' })
    );

    const callArgs = UrlFetchApp.fetch.mock.calls[0][1];
    const payload = JSON.parse(callArgs.payload);
    
    expect(payload.query).toBe('Return only the CEO name: Exa AI');
    expect(payload.type).toBe('auto');
    expect(payload.numResults).toBe(10);
    expect(payload.systemPrompt).toContain('No citations');
    expect(payload.outputSchema).toEqual({ type: 'text', description: 'Return only the CEO name' });
    expect(payload.contents).toEqual({ highlights: { maxCharacters: 4000 } });
  });

  test('should combine prompt and context in query', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        output: { content: 'https://exa.ai' },
        results: []
      })
    });

    EXA('Return only the company website URL', 'Exa AI');

    const callArgs = UrlFetchApp.fetch.mock.calls[0][1];
    const payload = JSON.parse(callArgs.payload);
    
    expect(payload.query).toBe('Return only the company website URL: Exa AI');
    expect(payload.outputSchema.description).toBe('Return only the company website URL');
  });

  test('should work without context', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        output: { content: 'Paris' },
        results: []
      })
    });

    EXA('What is the capital of France?');

    const callArgs = UrlFetchApp.fetch.mock.calls[0][1];
    const payload = JSON.parse(callArgs.payload);
    
    expect(payload.query).toBe('What is the capital of France?');
    expect(payload.type).toBe('auto');
  });

  test('should return output.content from response', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        output: { content: '150 employees' },
        results: []
      })
    });

    const result = EXA('Return only the company headcount', 'Exa AI');
    
    expect(result).toBe('150 employees');
  });

  test('should strip inline citation markers from output', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        output: { content: 'Sam Altman is the CEO of OpenAI [1][2][3]' },
        results: []
      })
    });

    const result = EXA('who is the ceo of OpenAI');
    
    expect(result).toBe('Sam Altman is the CEO of OpenAI');
  });

  test('should use systemPrompt and outputSchema description for formatting', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        output: { content: 'Sam Altman' },
        results: []
      })
    });

    EXA('who is the ceo of OpenAI, please ensure you only use the first and last name');

    const callArgs = UrlFetchApp.fetch.mock.calls[0][1];
    const payload = JSON.parse(callArgs.payload);
    
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      'https://api.exa.ai/search',
      expect.anything()
    );
    expect(payload.systemPrompt).toContain('formatting instructions');
    expect(payload.systemPrompt).toContain('No citations');
    expect(payload.outputSchema.type).toBe('text');
    expect(payload.outputSchema.description).toBe('who is the ceo of OpenAI, please ensure you only use the first and last name');
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
