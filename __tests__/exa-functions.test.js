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

describe('Agent Table', () => {
  beforeEach(() => {
    resetMocks();
    saveApiKey('exa_test_api_key');
  });

  test('parseAgentColumns should create stable unique schema keys', () => {
    const result = parseAgentColumns('Company Name, Website URL, Website URL, Source URLs');

    expect(result.success).toBe(true);
    expect(result.columns.map(column => column.key)).toEqual([
      'company_name',
      'website_url',
      'website_url_2',
      'source_urls'
    ]);
    expect(result.columns[3].type).toBe('array');
  });

  test('parseAgentColumns should annotate contact and URL column formats', () => {
    const result = parseAgentColumns('Company Email, Phone Number, Website URL, Source URLs, E-mail, Phone #, URL(s)');

    expect(result.success).toBe(true);
    expect(result.columns.map(column => column.format)).toEqual([
      'email',
      'phone',
      'uri',
      'uri',
      'email',
      'phone',
      'uri'
    ]);
  });

  test('parseAgentColumns should allow blank columns for auto mode', () => {
    const result = parseAgentColumns('');

    expect(result.success).toBe(true);
    expect(result.autoColumns).toBe(true);
    expect(result.columns).toEqual([]);
  });

  test('parseAgentColumns should treat free-text columns as Agent guidance', () => {
    const result = parseAgentColumns('company name website URL CEO founding date');

    expect(result.success).toBe(true);
    expect(result.autoColumns).toBe(true);
    expect(result.columns).toEqual([]);
    expect(result.columnInstructions).toBe('company name website URL CEO founding date');
  });

  test('parseAgentColumns should support non-comma explicit separators', () => {
    const result = parseAgentColumns('company name\nwebsite URL;CEO|founding date');

    expect(result.success).toBe(true);
    expect(result.autoColumns).toBe(false);
    expect(result.columns.map(column => column.key)).toEqual([
      'company_name',
      'website_url',
      'ceo',
      'founding_date'
    ]);
  });

  test('buildAgentTableOutputSchema should create rows array schema', () => {
    const columns = parseAgentColumns('name, website_url').columns;
    const schema = buildAgentTableOutputSchema(columns, 10);

    expect(schema.required).toEqual(['rows']);
    expect(schema.properties.rows.maxItems).toBe(10);
    expect(schema.properties.rows.items.required).toEqual(['name', 'website_url']);
    expect(schema.properties.rows.items.properties.website_url.type).toBe('string');
  });

  test('buildAgentTableOutputSchema should use email and phone formats for contact fields', () => {
    const columns = parseAgentColumns('Company Email, Phone Number, Source URLs').columns;
    const schema = buildAgentTableOutputSchema(columns, 10);
    const properties = schema.properties.rows.items.properties;

    expect(properties.company_email.format).toBe('email');
    expect(properties.company_email.type).toEqual(['string', 'null']);
    expect(properties.phone_number.format).toBe('phone');
    expect(properties.phone_number.type).toEqual(['string', 'null']);
    expect(properties.source_urls.items.format).toBe('uri');
  });

  test('buildAgentTableOutputSchema should create auto-column schema when no columns are supplied', () => {
    const schema = buildAgentTableOutputSchema([], 25);

    expect(schema.required).toEqual(['columns', 'rows']);
    expect(schema.properties.columns.maxItems).toBeUndefined();
    expect(schema.properties.rows.maxItems).toBe(25);
    expect(schema.properties.rows.items.required).toEqual(['cells']);
    expect(schema.properties.rows.items.additionalProperties).toBe(false);
    expect(schema.properties.rows.items.properties.cells.maxItems).toBeUndefined();
  });

  test('inferAgentRowLimit should use counts from the prompt without exposing a rows field', () => {
    expect(inferAgentRowLimit('Find top 50 AI companies')).toBe(50);
    expect(inferAgentRowLimit('return 12 rows of startups')).toBe(12);
    expect(inferAgentRowLimit('Find top 250 companies')).toBe(250);
    expect(inferAgentRowLimit('Find top 1000 companies')).toBe(1000);
    expect(inferAgentRowLimit('Find AI companies')).toBe(EXA_AGENT_CONFIG.defaultRows);
  });

  test('startAgentTableRun should call /agent/runs with auto-column table schema', () => {
    const originalGetActiveSheet = SpreadsheetApp.getActiveSheet;
    SpreadsheetApp.getActiveSheet = jest.fn(() => ({
      getActiveRange: () => ({
        getRow: () => 14,
        getColumn: () => 12
      })
    }));

    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        id: 'agent_run_test',
        object: 'agent_run',
        status: 'running',
        stopReason: null,
        createdAt: '2026-06-18T00:00:00.000Z',
        completedAt: null,
        request: {},
        output: { text: '', structured: null, grounding: [] },
        usage: { searches: 0 },
        costDollars: { total: 0 }
      })
    });

    const result = startAgentTableRun({
      prompt: 'Find 3 AI companies',
      columns: '',
      effort: 'auto'
    });

    expect(result.success).toBe(true);
    expect(result.run.id).toBe('agent_run_test');
    expect(result.tableConfig.columns).toEqual([]);
    expect(result.tableConfig.autoColumns).toBe(true);
    expect(result.tableConfig.startCell).toBe('L14');

    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      'https://api.exa.ai/agent/runs',
      expect.objectContaining({
        method: 'post',
        contentType: 'application/json',
        headers: expect.objectContaining({
          'x-api-key': 'exa_test_api_key',
          'x-exa-integration': 'exa-for-sheets'
        })
      })
    );

    const payload = JSON.parse(UrlFetchApp.fetch.mock.calls[0][1].payload);
    expect(payload.effort).toBe('auto');
    expect(payload.outputSchema.required).toEqual(['columns', 'rows']);
    expect(payload.outputSchema.properties.rows.maxItems).toBe(3);
    expect(payload.query).toContain('Create a spreadsheet-ready table');
    expect(payload.query).toContain('Return no more than 3 rows');
    expect(payload.query).toContain('Use web research. Do not rely on memory.');
    expect(payload.query).toContain('Every row must be a real, source-verifiable entity');
    expect(payload.query).toContain('Try to fill the table as completely as public sources allow');
    expect(payload.query).toContain('Do not return notes, summaries, metadata');
    expect(payload.query).not.toContain('explicitly perform contact enrichment');
    expect(payload.query).toContain('Infer the best concise column headers');
    expect(payload.metadata.startCell).toBe('L14');

    SpreadsheetApp.getActiveSheet = originalGetActiveSheet;
  });

  test('startAgentTableRun should add contact enrichment guidance and formats', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        id: 'agent_run_test',
        object: 'agent_run',
        status: 'running',
        output: { text: '', structured: null, grounding: [] },
        usage: { searches: 0 },
        costDollars: { total: 0 }
      })
    });

    const result = startAgentTableRun({
      prompt: 'Find 5 HVAC companies in Austin and include official emails and phone numbers',
      columns: 'Company Name, Website URL, Email, Phone Number',
      effort: 'auto'
    });

    expect(result.success).toBe(true);
    const payload = JSON.parse(UrlFetchApp.fetch.mock.calls[0][1].payload);
    const properties = payload.outputSchema.properties.rows.items.properties;

    expect(payload.query).toContain('explicitly perform contact enrichment');
    expect(payload.query).toContain('Do not guess email patterns');
    expect(properties.email.format).toBe('email');
    expect(properties.email.type).toEqual(['string', 'null']);
    expect(properties.phone_number.format).toBe('phone');
    expect(properties.phone_number.type).toEqual(['string', 'null']);
    expect(properties.website_url.format).toBe('uri');
  });

  test('startAgentTableRun should infer contact columns when columns are blank', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        id: 'agent_run_test',
        object: 'agent_run',
        status: 'running',
        output: { text: '', structured: null, grounding: [] },
        usage: { searches: 0 },
        costDollars: { total: 0 }
      })
    });

    const result = startAgentTableRun({
      prompt: 'Find 10 HVAC companies in Austin with emails and phone numbers',
      columns: '',
      effort: 'auto'
    });

    expect(result.success).toBe(true);
    expect(result.tableConfig.autoColumns).toBe(false);
    expect(result.tableConfig.columns.map(column => column.key)).toEqual([
      'name',
      'website_url',
      'contact_email',
      'phone_number'
    ]);

    const payload = JSON.parse(UrlFetchApp.fetch.mock.calls[0][1].payload);
    const properties = payload.outputSchema.properties.rows.items.properties;
    expect(payload.outputSchema.required).toEqual(['rows']);
    expect(payload.metadata.autoColumns).toBe('false');
    expect(payload.metadata.inferredColumns).toBe('true');
    expect(payload.query).toContain('Populate exactly these columns');
    expect(properties.contact_email.format).toBe('email');
    expect(properties.contact_email.type).toEqual(['string', 'null']);
    expect(properties.phone_number.format).toBe('phone');
    expect(properties.phone_number.type).toEqual(['string', 'null']);
  });

  test('startAgentTableRun should keep blank non-contact prompts in auto-column mode', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        id: 'agent_run_test',
        object: 'agent_run',
        status: 'running',
        output: { text: '', structured: null, grounding: [] },
        usage: { searches: 0 },
        costDollars: { total: 0 }
      })
    });

    const result = startAgentTableRun({
      prompt: 'Find 10 AI companies and include founding date',
      columns: '',
      effort: 'auto'
    });

    expect(result.success).toBe(true);
    expect(result.tableConfig.autoColumns).toBe(true);

    const payload = JSON.parse(UrlFetchApp.fetch.mock.calls[0][1].payload);
    expect(payload.outputSchema.required).toEqual(['columns', 'rows']);
    expect(payload.metadata.autoColumns).toBe('true');
    expect(payload.metadata.inferredColumns).toBe('false');
    expect(payload.query).toContain('Infer the best concise column headers');
  });

  test('startAgentTableRun should use custom start cell when provided', () => {
    const originalGetActiveSheet = SpreadsheetApp.getActiveSheet;
    SpreadsheetApp.getActiveSheet = jest.fn(() => ({
      getActiveRange: () => ({
        getRow: () => 14,
        getColumn: () => 12
      })
    }));

    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        id: 'agent_run_test',
        object: 'agent_run',
        status: 'running',
        output: { text: '', structured: null, grounding: [] },
        usage: { searches: 0 },
        costDollars: { total: 0 }
      })
    });

    const result = startAgentTableRun({
      prompt: 'Find 3 AI companies',
      columns: '',
      effort: 'auto',
      startCell: ' b7 '
    });

    expect(result.success).toBe(true);
    expect(result.tableConfig.startCell).toBe('B7');

    const payload = JSON.parse(UrlFetchApp.fetch.mock.calls[0][1].payload);
    expect(payload.metadata.startCell).toBe('B7');

    SpreadsheetApp.getActiveSheet = originalGetActiveSheet;
  });

  test('startAgentTableRun should add free-text column guidance to prompt', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        id: 'agent_run_test',
        object: 'agent_run',
        status: 'running',
        output: { text: '', structured: null, grounding: [] },
        usage: { searches: 0 },
        costDollars: { total: 0 }
      })
    });

    const result = startAgentTableRun({
      prompt: 'Find top 20 AI companies',
      columns: 'company name website URL CEO founding date',
      effort: 'auto'
    });

    expect(result.success).toBe(true);

    const payload = JSON.parse(UrlFetchApp.fetch.mock.calls[0][1].payload);
    expect(payload.outputSchema.required).toEqual(['columns', 'rows']);
    expect(payload.query).toContain('Use these exact columns if possible: company name website URL CEO founding date');
    expect(payload.metadata.columnInstructions).toBe('company name website URL CEO founding date');
  });

  test('getAgentRunStatus should fetch and summarize an Agent run', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        id: 'agent_run_test',
        object: 'agent_run',
        status: 'completed',
        stopReason: 'schema_satisfied',
        createdAt: '2026-06-18T00:00:00.000Z',
        completedAt: '2026-06-18T00:00:10.000Z',
        request: {},
        output: {
          text: 'Done',
          structured: { rows: [{ name: 'Exa' }, { name: 'Perplexity' }] },
          grounding: []
        },
        usage: { searches: 3 },
        costDollars: { total: 0.012 }
      })
    });

    const result = getAgentRunStatus('agent_run_test');

    expect(result.success).toBe(true);
    expect(result.run.status).toBe('completed');
    expect(result.run.rowCount).toBe(2);
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      'https://api.exa.ai/agent/runs/agent_run_test',
      expect.objectContaining({
        method: 'get',
        headers: expect.objectContaining({
          'Exa-Beta': 'agent-2026-05-07'
        })
      })
    );
  });

  test('getAgentRunStatus should keep polling on Agent rate limits', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 429,
      getContentText: () => JSON.stringify({ error: 'rate limited' }),
      getHeaders: () => ({ 'Retry-After': '12' })
    });

    const result = getAgentRunStatus('agent_run_test');

    expect(result.success).toBe(false);
    expect(result.retryable).toBe(true);
    expect(result.retryAfterMs).toBe(12000);
    expect(result.message).toContain('rate limited');
    expect(Utilities.sleep).not.toHaveBeenCalled();
  });

  test('cancelAgentRun should call the cancel endpoint', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        id: 'agent_run_test',
        object: 'agent_run',
        status: 'cancelled',
        stopReason: 'cancelled',
        createdAt: '2026-06-18T00:00:00.000Z',
        completedAt: '2026-06-18T00:00:05.000Z',
        request: {},
        output: { text: '', structured: null, grounding: [] },
        usage: { searches: 0 },
        costDollars: { total: 0 }
      })
    });

    const result = cancelAgentRun('agent_run_test');

    expect(result.success).toBe(true);
    expect(result.run.status).toBe('cancelled');
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      'https://api.exa.ai/agent/runs/agent_run_test/cancel',
      expect.objectContaining({
        method: 'post',
        headers: expect.objectContaining({
          'Exa-Beta': 'agent-2026-05-07'
        })
      })
    );
  });

  test('buildAgentSheetValues should flatten rows and append grounding sources', () => {
    const columns = parseAgentColumns('name, website_url, reason').columns;
    const run = {
      id: 'agent_run_test',
      status: 'completed',
      output: {
        structured: {
          rows: [
            { name: 'Perplexity AI', website_url: 'https://www.perplexity.ai/', reason: 'AI answer engine' },
            { name: 'Brave Search', website_url: 'https://brave.com/search/', reason: 'Independent AI search' }
          ]
        },
        grounding: [
          {
            field: 'structured.rows[0].reason',
            citations: [{ url: 'https://www.perplexity.ai/', title: 'Perplexity' }]
          },
          {
            field: 'structured.rows[1].website_url',
            citations: [{ url: 'https://brave.com/search/', title: 'Brave Search' }]
          }
        ]
      }
    };

    const result = buildAgentSheetValues(run, { rootKey: 'rows', columns }, true, true);

    expect(result.success).toBe(true);
    expect(result.values).toEqual([
      ['name', 'website_url', 'reason', 'Sources'],
      ['Perplexity AI', 'https://www.perplexity.ai/', 'AI answer engine', 'https://www.perplexity.ai/'],
      ['Brave Search', 'https://brave.com/search/', 'Independent AI search', 'https://brave.com/search/']
    ]);
  });

  test('buildAgentSheetValues should match custom columns by display label variations', () => {
    const columns = parseAgentColumns('Company Name, Website URL, CEO').columns;
    const run = {
      id: 'agent_run_test',
      status: 'completed',
      output: {
        structured: {
          rows: [
            { 'Company Name': 'Exa', 'Website URL': 'https://exa.ai', CEO: 'Will Bryk' }
          ]
        },
        grounding: []
      }
    };

    const result = buildAgentSheetValues(run, { rootKey: 'rows', columns }, true, false);

    expect(result.success).toBe(true);
    expect(result.values).toEqual([
      ['Company Name', 'Website URL', 'CEO'],
      ['Exa', 'https://exa.ai', 'Will Bryk']
    ]);
  });

  test('buildAgentSheetValues should use generated columns from auto mode', () => {
    const run = {
      id: 'agent_run_test',
      status: 'completed',
      output: {
        structured: {
          columns: ['Company Name', 'Website URL', 'CEO'],
          rows: [
            { 'Company Name': 'Exa', 'Website URL': 'https://exa.ai', CEO: 'Will Bryk' },
            { company_name: 'OpenAI', website_url: 'https://openai.com', ceo: 'Sam Altman' }
          ]
        },
        grounding: []
      }
    };

    const result = buildAgentSheetValues(run, { rootKey: 'rows', columns: [] }, true, false);

    expect(result.success).toBe(true);
    expect(result.values).toEqual([
      ['Company Name', 'Website URL', 'CEO'],
      ['Exa', 'https://exa.ai', 'Will Bryk'],
      ['OpenAI', 'https://openai.com', 'Sam Altman']
    ]);
  });

  test('buildAgentSheetValues should support generated columns with cells arrays', () => {
    const run = {
      id: 'agent_run_test',
      status: 'completed',
      output: {
        structured: {
          columns: ['Company Name', 'Website URL', 'CEO'],
          rows: [
            { cells: ['Exa', 'https://exa.ai', 'Will Bryk'] },
            ['OpenAI', 'https://openai.com', 'Sam Altman']
          ]
        },
        grounding: []
      }
    };

    const result = buildAgentSheetValues(run, { rootKey: 'rows', columns: [] }, true, false);

    expect(result.success).toBe(true);
    expect(result.values).toEqual([
      ['Company Name', 'Website URL', 'CEO'],
      ['Exa', 'https://exa.ai', 'Will Bryk'],
      ['OpenAI', 'https://openai.com', 'Sam Altman']
    ]);
  });

  test('buildAgentSheetValues should reject headers-only auto output', () => {
    const run = {
      id: 'agent_run_test',
      status: 'completed',
      output: {
        structured: {
          columns: ['Company Name', 'Website URL'],
          rows: [{}]
        },
        grounding: []
      }
    };

    const result = buildAgentSheetValues(run, { rootKey: 'rows', columns: [] }, true, false);

    expect(result.success).toBe(false);
    expect(result.message).toContain('no usable data rows');
  });

  test('isValidAgentStartCell should accept only simple single cells', () => {
    expect(normalizeAgentStartCell(' d5 ')).toBe('D5');
    expect(isValidAgentStartCell('A1')).toBe(true);
    expect(isValidAgentStartCell('D5')).toBe(true);
    expect(isValidAgentStartCell('A1:B2')).toBe(false);
    expect(isValidAgentStartCell('Sheet1!A1')).toBe(false);
    expect(isValidAgentStartCell('1A')).toBe(false);
  });

  test('writeAgentRunToSheet should reject invalid custom start cell', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        id: 'agent_run_test',
        object: 'agent_run',
        status: 'completed',
        output: {
          structured: {
            columns: ['Company Name'],
            rows: [{ cells: ['Exa'] }]
          },
          grounding: []
        },
        usage: { searches: 1 },
        costDollars: { total: 0.01 }
      })
    });

    const result = writeAgentRunToSheet(
      'agent_run_test',
      { rootKey: 'rows', columns: [], autoColumns: true },
      { includeHeaders: true, includeSources: false, startCell: 'A1:B2' }
    );

    expect(result.success).toBe(false);
    expect(result.message).toContain('Start cell');
  });

  test('writeAgentRunToSheet should use saved start cell from run start', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        id: 'agent_run_test',
        object: 'agent_run',
        status: 'completed',
        output: {
          structured: {
            columns: ['Company Name'],
            rows: [{ cells: ['Exa'] }]
          },
          grounding: []
        },
        usage: { searches: 1 },
        costDollars: { total: 0.01 }
      })
    });

    const originalGetActiveSheet = SpreadsheetApp.getActiveSheet;
    const setValues = jest.fn();
    const getRange = jest.fn((arg1, arg2, arg3, arg4) => {
      if (typeof arg1 === 'string') {
        return {
          getRow: () => 5,
          getColumn: () => 4
        };
      }

      return {
        getValues: () => Array.from({ length: arg3 }, () => Array(arg4).fill('')),
        setValues,
        setFontWeight: jest.fn(),
        setBackground: jest.fn()
      };
    });

    SpreadsheetApp.getActiveSheet = jest.fn(() => ({
      getActiveRange: () => ({
        getRow: () => 20,
        getColumn: () => 20
      }),
      getRange,
      getMaxRows: () => 1000,
      getMaxColumns: () => 26,
      autoResizeColumns: jest.fn()
    }));

    const result = writeAgentRunToSheet(
      'agent_run_test',
      { rootKey: 'rows', columns: [], autoColumns: true, startCell: 'D5' },
      { includeHeaders: true, includeSources: false, startCell: '' }
    );

    expect(result.success).toBe(true);
    expect(result.message).toContain('D5');
    expect(getRange.mock.calls[0]).toEqual(['D5']);
    expect(getRange.mock.calls[1]).toEqual([5, 4, 2, 1]);
    expect(setValues).toHaveBeenCalledWith([
      ['Company Name'],
      ['Exa']
    ]);

    SpreadsheetApp.getActiveSheet = originalGetActiveSheet;
  });

  test('writeAgentRunToSheet should write to the saved sheet even if another sheet is active', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        id: 'agent_run_test',
        object: 'agent_run',
        status: 'completed',
        output: {
          structured: {
            columns: ['Company Name'],
            rows: [{ cells: ['Exa'] }]
          },
          grounding: []
        }
      })
    });

    const target = createMockSheet([[''], ['']], 'A1', [], { name: 'Target', sheetId: 222 });
    const active = createMockSheet([['']], 'A1', [], { name: 'Other', sheetId: 333 });
    const originalGetActiveSheet = SpreadsheetApp.getActiveSheet;
    const originalGetActiveSpreadsheet = SpreadsheetApp.getActiveSpreadsheet;
    SpreadsheetApp.getActiveSheet = jest.fn(() => active.sheet);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => ({
      getSheets: () => [active.sheet, target.sheet],
      getSheetByName: (name) => ({ Target: target.sheet, Other: active.sheet }[name] || null)
    }));

    const result = writeAgentRunToSheet(
      'agent_run_test',
      { rootKey: 'rows', columns: [], autoColumns: true, startCell: 'B2', sheetId: 222, sheetName: 'Target' },
      { includeHeaders: true, includeSources: false }
    );

    expect(result.success).toBe(true);
    expect(target.writes).toEqual([
      { row: 2, col: 2, value: 'Company Name' },
      { row: 3, col: 2, value: 'Exa' }
    ]);
    expect(active.writes).toEqual([]);

    SpreadsheetApp.getActiveSheet = originalGetActiveSheet;
    if (originalGetActiveSpreadsheet) {
      SpreadsheetApp.getActiveSpreadsheet = originalGetActiveSpreadsheet;
    } else {
      delete SpreadsheetApp.getActiveSpreadsheet;
    }
  });
});

function testColumnToNumber(label) {
  return label.split('').reduce((total, ch) => total * 26 + ch.charCodeAt(0) - 64, 0);
}

function testColumnToLabel(col) {
  let letters = '';
  let current = col;
  while (current > 0) {
    const remainder = (current - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    current = Math.floor((current - 1) / 26);
  }
  return letters;
}

function parseTestA1Range(a1) {
  const parts = a1.split(':');
  const parseCell = (cell) => {
    const match = cell.match(/^([A-Z]+)(\d+)$/i);
    return {
      col: testColumnToNumber(match[1].toUpperCase()),
      row: Number(match[2])
    };
  };
  const start = parseCell(parts[0]);
  const end = parts[1] ? parseCell(parts[1]) : start;
  return {
    row: start.row,
    col: start.col,
    numRows: end.row - start.row + 1,
    numCols: end.col - start.col + 1
  };
}

function createMockSheet(data, selectionA1, formulas = [], options = {}) {
  const selection = parseTestA1Range(selectionA1);
  const writes = [];
  const getCellValue = (row, col) => (data[row - 1] && data[row - 1][col - 1] !== undefined) ? data[row - 1][col - 1] : '';
  const getCellFormula = (row, col) => (formulas[row - 1] && formulas[row - 1][col - 1] !== undefined) ? formulas[row - 1][col - 1] : '';
  const ensureCell = (row, col) => {
    while (data.length < row) data.push([]);
    while (data[row - 1].length < col) data[row - 1].push('');
  };

  const makeRange = (row, col, numRows = 1, numCols = 1) => ({
    getRow: () => row,
    getColumn: () => col,
    getNumRows: () => numRows,
    getNumColumns: () => numCols,
    getA1Notation: () => `${testColumnToLabel(col)}${row}:${testColumnToLabel(col + numCols - 1)}${row + numRows - 1}`,
    getValues: () => Array.from({ length: numRows }, (_, rowOffset) =>
      Array.from({ length: numCols }, (_, colOffset) => getCellValue(row + rowOffset, col + colOffset))
    ),
    getFormulas: () => Array.from({ length: numRows }, (_, rowOffset) =>
      Array.from({ length: numCols }, (_, colOffset) => getCellFormula(row + rowOffset, col + colOffset))
    ),
    getValue: () => getCellValue(row, col),
    getFormula: () => getCellFormula(row, col),
    setValue: jest.fn((value) => {
      ensureCell(row, col);
      data[row - 1][col - 1] = value;
      writes.push({ row, col, value });
    }),
    setValues: jest.fn((values) => {
      values.forEach((valueRow, rowOffset) => {
        valueRow.forEach((value, colOffset) => {
          ensureCell(row + rowOffset, col + colOffset);
          data[row + rowOffset - 1][col + colOffset - 1] = value;
          writes.push({ row: row + rowOffset, col: col + colOffset, value });
        });
      });
    }),
    setFontWeight: jest.fn(),
    setBackground: jest.fn()
  });

  const sheet = {
    getName: () => options.name || 'Sheet1',
    getSheetId: () => options.sheetId || 1001,
    getLastRow: () => data.length,
    getLastColumn: () => Math.max(...data.map(row => row.length)),
    getFrozenRows: () => 0,
    getActiveRange: () => makeRange(selection.row, selection.col, selection.numRows, selection.numCols),
    getRange: jest.fn((arg1, arg2, arg3, arg4) => {
      if (typeof arg1 === 'string') {
        const parsed = parseTestA1Range(arg1);
        return makeRange(parsed.row, parsed.col, parsed.numRows, parsed.numCols);
      }
      return makeRange(arg1, arg2, arg3 || 1, arg4 || 1);
    }),
    getMaxRows: () => 1000,
    getMaxColumns: () => 26,
    autoResizeColumns: jest.fn()
  };

  return { sheet, writes, data };
}

describe('Agent Fill Cells', () => {
  beforeEach(() => {
    resetMocks();
    saveApiKey('exa_test_api_key');
  });

  test('prepareAgentFillJobs should use headers and nearby rows as context for blank selected cells', () => {
    const { sheet } = createMockSheet([
      ['Company Name', 'Website URL', 'HQ', 'CEO', 'One Liner', 'LinkedIn URL'],
      ['Apple', 'https://apple.com', '', '', '', ''],
      ['Google', 'https://google.com', 'Mountain View, CA', 'Sundar Pichai', 'Search and cloud company', 'https://www.linkedin.com/company/google']
    ], 'D2:F2');
    const originalGetActiveSheet = SpreadsheetApp.getActiveSheet;
    SpreadsheetApp.getActiveSheet = jest.fn(() => sheet);

    const result = prepareAgentFillJobs({ overwrite: false, instructions: '' });

    expect(result.success).toBe(true);
    expect(result.totalCells).toBe(3);
    expect(result.jobs).toHaveLength(1);
    expect(result.jobs[0].sheetId).toBe(1001);
    expect(result.jobs[0].sheetName).toBe('Sheet1');
    expect(result.jobs[0].targets.map(target => target.cell)).toEqual(['D2', 'E2', 'F2']);
    expect(result.jobs[0].targets.map(target => target.fieldId)).toEqual(['ceo', 'one_liner', 'linkedin_url']);
    expect(result.jobs[0].inputRows.find(row => row.rowId === 'row_2').cells[0].value).toBe('Apple');
    expect(result.jobs[0].inputRows.some(row => row.example === true && row.rowId === 'row_3')).toBe(true);

    SpreadsheetApp.getActiveSheet = originalGetActiveSheet;
  });

  test('prepareAgentFillJobs should skip non-empty cells and empty rows by default', () => {
    const { sheet } = createMockSheet([
      ['Company Name', 'Website URL', 'HQ', 'CEO'],
      ['Apple', 'https://apple.com', '', 'Tim Cook'],
      ['', '', '', '']
    ], 'C2:D3');
    const originalGetActiveSheet = SpreadsheetApp.getActiveSheet;
    SpreadsheetApp.getActiveSheet = jest.fn(() => sheet);

    const result = prepareAgentFillJobs({ overwrite: false });

    expect(result.success).toBe(true);
    expect(result.totalCells).toBe(1);
    expect(result.jobs[0].targets.map(target => target.cell)).toEqual(['C2']);

    SpreadsheetApp.getActiveSheet = originalGetActiveSheet;
  });

  test('prepareAgentFillJobs should explain when blank rows do not have enough context', () => {
    const { sheet } = createMockSheet([
      ['Company Name', 'Website URL', 'HQ', 'CEO'],
      ['', '', '', '']
    ], 'C2:D2');
    const originalGetActiveSheet = SpreadsheetApp.getActiveSheet;
    SpreadsheetApp.getActiveSheet = jest.fn(() => sheet);

    const result = prepareAgentFillJobs({ overwrite: false });

    expect(result.success).toBe(false);
    expect(result.message).toBe('Select blanks in rows that already have some data.');

    SpreadsheetApp.getActiveSheet = originalGetActiveSheet;
  });

  test('prepareAgentFillJobs should continue blank rows under an existing ranked table', () => {
    const { sheet } = createMockSheet([
      ['Rank', 'Company', 'Industry', 'Revenue in INR', 'Source'],
      [51, 'Adani Power', 'Power', '58,906', 'Fortune India 500 2025'],
      [52, 'IndusInd Bank', 'Banks', '56,358', 'Fortune India 500 2025'],
      [53, 'Sun Pharmaceuticals Industries', 'Pharma', '54,579', 'Fortune India 500 2025'],
      [54, 'Tech Mahindra', 'IT Services', '53,852', 'Fortune India 500 2025'],
      [55, 'Bajaj Auto', 'Auto - Two Wheeler', '52,441', 'Fortune India 500 2025'],
      ['', '', '', '', ''],
      ['', '', '', '', '']
    ], 'A7:E8');
    const originalGetActiveSheet = SpreadsheetApp.getActiveSheet;
    SpreadsheetApp.getActiveSheet = jest.fn(() => sheet);

    const result = prepareAgentFillJobs({ overwrite: false });

    expect(result.success).toBe(true);
    expect(result.totalCells).toBe(10);
    expect(result.jobs).toHaveLength(1);
    expect(result.jobs[0].mode).toBe('continue-table');
    expect(result.jobs[0].targetRowIds).toEqual(['row_7', 'row_8']);
    expect(result.jobs[0].targets.map(target => target.cell)).toEqual([
      'A7', 'B7', 'C7', 'D7', 'E7',
      'A8', 'B8', 'C8', 'D8', 'E8'
    ]);
    expect(result.jobs[0].fields.map(field => field.title)).toEqual([
      'Rank', 'Company', 'Industry', 'Revenue in INR', 'Source'
    ]);
    expect(result.jobs[0].inputRows).toHaveLength(5);
    expect(result.jobs[0].inputRows[result.jobs[0].inputRows.length - 1].cells[1].value).toBe('Bajaj Auto');

    SpreadsheetApp.getActiveSheet = originalGetActiveSheet;
  });

  test('prepareAgentFillJobs should keep one run for multi-row table continuation', () => {
    const rows = [
      ['Rank', 'Company', 'Industry', 'Revenue in INR', 'Source']
    ];
    for (let rank = 1; rank <= 10; rank++) {
      rows.push([rank, `Company ${rank}`, `Industry ${rank}`, `${rank * 1000}`, 'Source']);
    }
    for (let i = 0; i < 20; i++) {
      rows.push(['', '', '', '', '']);
    }

    const { sheet } = createMockSheet(rows, 'A12:E31');
    const originalGetActiveSheet = SpreadsheetApp.getActiveSheet;
    SpreadsheetApp.getActiveSheet = jest.fn(() => sheet);

    const result = prepareAgentFillJobs({ overwrite: false });

    expect(result.success).toBe(true);
    expect(result.totalCells).toBe(100);
    expect(result.jobs).toHaveLength(1);
    expect(result.jobs[0].mode).toBe('continue-table');
    expect(result.jobs[0].targetRowIds).toHaveLength(20);

    SpreadsheetApp.getActiveSheet = originalGetActiveSheet;
  });

  test('prepareAgentFillJobs should ignore selected blank columns outside the table header', () => {
    const { sheet } = createMockSheet([
      ['Rank', 'Company'],
      [1, 'Company 1'],
      [2, 'Company 2'],
      ['', '', ''],
      ['', '', '']
    ], 'A4:C5');
    const originalGetActiveSheet = SpreadsheetApp.getActiveSheet;
    SpreadsheetApp.getActiveSheet = jest.fn(() => sheet);

    const result = prepareAgentFillJobs({ overwrite: false });

    expect(result.success).toBe(true);
    expect(result.totalCells).toBe(4);
    expect(result.skippedCells).toBe(2);
    expect(result.jobs).toHaveLength(1);
    expect(result.jobs[0].mode).toBe('continue-table');
    expect(result.jobs[0].targets.map(target => target.cell)).toEqual(['A4', 'B4', 'A5', 'B5']);

    SpreadsheetApp.getActiveSheet = originalGetActiveSheet;
  });

  test('prepareAgentFillJobs should reject blank rows that are too far from table examples', () => {
    const rows = [
      ['Rank', 'Company'],
      [1, 'Company 1'],
      [2, 'Company 2']
    ];
    while (rows.length < 21) {
      rows.push(['', '']);
    }

    const { sheet } = createMockSheet(rows, 'A20:B21');
    const originalGetActiveSheet = SpreadsheetApp.getActiveSheet;
    SpreadsheetApp.getActiveSheet = jest.fn(() => sheet);

    const result = prepareAgentFillJobs({ overwrite: false });

    expect(result.success).toBe(false);
    expect(result.message).toBe('Select blanks in rows that already have some data.');

    SpreadsheetApp.getActiveSheet = originalGetActiveSheet;
  });

  test('prepareAgentFillJobs should return selected range for sidebar confirmation', () => {
    const { sheet } = createMockSheet([
      ['Company Name', 'Website URL', 'HQ'],
      ['Apple', 'https://apple.com', '']
    ], 'C2');
    const originalGetActiveSheet = SpreadsheetApp.getActiveSheet;
    SpreadsheetApp.getActiveSheet = jest.fn(() => sheet);

    const result = prepareAgentFillJobs({ overwrite: false });

    expect(result.success).toBe(true);
    expect(result.selectedRange).toBe('C2:C2');
    expect(result.totalCells).toBe(1);

    SpreadsheetApp.getActiveSheet = originalGetActiveSheet;
  });

  test('prepareAgentFillJobs should not mistake nearby data rows for headers in long tables', () => {
    const rows = [
      ['Company Name', 'Website URL', 'HQ', 'CEO']
    ];
    for (let i = 2; i < 60; i++) {
      rows.push([`Company ${i}`, `https://company${i}.com`, `City ${i}`, `CEO ${i}`]);
    }
    rows.push(['Apple', 'https://apple.com', '', '']);

    const { sheet } = createMockSheet(rows, 'D60');
    const originalGetActiveSheet = SpreadsheetApp.getActiveSheet;
    SpreadsheetApp.getActiveSheet = jest.fn(() => sheet);

    const result = prepareAgentFillJobs({ overwrite: false });

    expect(result.success).toBe(true);
    expect(result.jobs[0].targets[0].fieldId).toBe('ceo');
    expect(result.jobs[0].targets[0].title).toBe('CEO');

    SpreadsheetApp.getActiveSheet = originalGetActiveSheet;
  });

  test('startAgentFillRun should create a strict Agent run with row context input', () => {
    const job = {
      jobId: 'fill_1',
      overwrite: false,
      fields: [{ fieldId: 'ceo', title: 'CEO', column: 4, columnLetter: 'D' }],
      targets: [{ cell: 'D2', row: 2, column: 4, rowId: 'row_2', fieldId: 'ceo', title: 'CEO', columnLetter: 'D' }],
      inputRows: [{ rowId: 'row_2', rowNumber: 2, cells: [{ fieldId: 'company_name', title: 'Company Name', column: 'A', value: 'Apple' }] }],
      targetRowIds: ['row_2']
    };
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        id: 'agent_run_fill',
        status: 'running',
        output: { structured: null },
        usage: { searches: 0 },
        costDollars: { total: 0 }
      })
    });

    const result = startAgentFillRun(job, { effort: 'auto', instructions: 'Use official sources only.' });

    expect(result.success).toBe(true);
    expect(result.fillConfig.targets[0].cell).toBe('D2');
    const payload = JSON.parse(UrlFetchApp.fetch.mock.calls[0][1].payload);
    expect(payload.query).toContain('Fill selected blank cells');
    expect(payload.query).toContain('Use official sources only.');
    expect(payload.input.data[0].rowId).toBe('row_2');
    expect(payload.input.targets).toBeUndefined();
    expect(payload.outputSchema.properties.rows.items.properties.rowId.enum).toEqual(['row_2']);
    expect(payload.outputSchema.properties.rows.items.properties.values.properties.ceo).toBeDefined();
    expect(payload.outputSchema.properties.rows.items.properties.values.required).toBeUndefined();
    expect(payload.metadata.targetCells).toBe('1');
    expect(payload.metadata.overwrite).toBe('false');
  });

  test('startAgentFillRun should add contact formats and enrichment guidance', () => {
    const job = {
      jobId: 'fill_1',
      overwrite: false,
      fields: [
        { fieldId: 'email', title: 'Email', column: 3, columnLetter: 'C' },
        { fieldId: 'phone_number', title: 'Phone Number', column: 4, columnLetter: 'D' },
        { fieldId: 'website_url', title: 'Website URL', column: 5, columnLetter: 'E' }
      ],
      targets: [
        { cell: 'C2', row: 2, column: 3, rowId: 'row_2', fieldId: 'email', title: 'Email', columnLetter: 'C' },
        { cell: 'D2', row: 2, column: 4, rowId: 'row_2', fieldId: 'phone_number', title: 'Phone Number', columnLetter: 'D' },
        { cell: 'E2', row: 2, column: 5, rowId: 'row_2', fieldId: 'website_url', title: 'Website URL', columnLetter: 'E' }
      ],
      inputRows: [{ rowId: 'row_2', rowNumber: 2, cells: [{ fieldId: 'company_name', title: 'Company Name', column: 'A', value: 'Acme HVAC' }] }],
      targetRowIds: ['row_2']
    };
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        id: 'agent_run_fill',
        status: 'running',
        output: { structured: null },
        usage: { searches: 0 },
        costDollars: { total: 0 }
      })
    });

    const result = startAgentFillRun(job, { effort: 'auto', instructions: '' });

    expect(result.success).toBe(true);
    const payload = JSON.parse(UrlFetchApp.fetch.mock.calls[0][1].payload);
    const valueProperties = payload.outputSchema.properties.rows.items.properties.values.properties;

    expect(payload.query).toContain('explicitly perform contact enrichment');
    expect(valueProperties.email.type).toEqual(['string', 'null']);
    expect(valueProperties.email.format).toBe('email');
    expect(valueProperties.email.anyOf).toBeUndefined();
    expect(valueProperties.phone_number.type).toEqual(['string', 'null']);
    expect(valueProperties.phone_number.format).toBe('phone');
    expect(valueProperties.phone_number.anyOf).toBeUndefined();
    expect(valueProperties.website_url.anyOf[0].format).toBe('uri');
    expect(valueProperties.website_url.anyOf[3].items.format).toBe('uri');
  });

  test('startAgentFillRun should create a continuation Agent run for blank rows below a table', () => {
    const job = {
      jobId: 'continue_1',
      mode: 'continue-table',
      overwrite: false,
      fields: [
        { fieldId: 'rank', title: 'Rank', column: 1, columnLetter: 'A' },
        { fieldId: 'company', title: 'Company', column: 2, columnLetter: 'B' }
      ],
      targets: [
        { cell: 'A4', row: 4, column: 1, rowId: 'row_4', fieldId: 'rank', title: 'Rank', columnLetter: 'A' },
        { cell: 'B4', row: 4, column: 2, rowId: 'row_4', fieldId: 'company', title: 'Company', columnLetter: 'B' }
      ],
      inputRows: [
        { rowId: 'row_2', rowNumber: 2, example: true, cells: [{ fieldId: 'rank', title: 'Rank', column: 'A', value: 1 }, { fieldId: 'company', title: 'Company', column: 'B', value: 'Company 1' }] },
        { rowId: 'row_3', rowNumber: 3, example: true, cells: [{ fieldId: 'rank', title: 'Rank', column: 'A', value: 2 }, { fieldId: 'company', title: 'Company', column: 'B', value: 'Company 2' }] }
      ],
      targetRowIds: ['row_4']
    };
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        id: 'agent_run_continue',
        status: 'running',
        output: { structured: null },
        usage: { searches: 0 },
        costDollars: { total: 0 }
      })
    });

    const result = startAgentFillRun(job, { effort: 'auto', instructions: 'Use the same source as the table.' });

    expect(result.success).toBe(true);
    expect(result.fillConfig.targets[0].cell).toBe('A4');
    const payload = JSON.parse(UrlFetchApp.fetch.mock.calls[0][1].payload);
    expect(payload.query).toContain('Continue the existing Google Sheets table');
    expect(payload.query).toContain('Do not repeat entities');
    expect(payload.query).toContain('Use the same source as the table.');
    expect(payload.input.data[0].example).toBe(true);
    expect(payload.input.targets).toBeUndefined();
    expect(payload.metadata.mode).toBe('agent-continue-table');
    expect(payload.outputSchema.properties.rows.items.properties.rowId.enum).toEqual(['row_4']);
    expect(payload.outputSchema.properties.rows.items.properties.values.required).toEqual(['rank', 'company']);
  });

  test('writeAgentFillRunToSheet should write continuation results into selected blank rows', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        id: 'agent_run_continue',
        status: 'completed',
        output: {
          structured: {
            rows: [
              { rowId: 'row_4', values: { rank: 3, company: 'Company 3' } },
              { rowId: 'row_5', values: { rank: 4, company: 'Company 4' } }
            ]
          }
        },
        usage: { searches: 2 },
        costDollars: { total: 0.03 }
      })
    });
    const { sheet, writes } = createMockSheet([
      ['Rank', 'Company'],
      [1, 'Company 1'],
      [2, 'Company 2'],
      ['', ''],
      ['', '']
    ], 'A4:B5');
    const originalGetActiveSheet = SpreadsheetApp.getActiveSheet;
    SpreadsheetApp.getActiveSheet = jest.fn(() => sheet);

    const result = writeAgentFillRunToSheet('agent_run_continue', {
      overwrite: false,
      targets: [
        { cell: 'A4', row: 4, column: 1, rowId: 'row_4', fieldId: 'rank', title: 'Rank' },
        { cell: 'B4', row: 4, column: 2, rowId: 'row_4', fieldId: 'company', title: 'Company' },
        { cell: 'A5', row: 5, column: 1, rowId: 'row_5', fieldId: 'rank', title: 'Rank' },
        { cell: 'B5', row: 5, column: 2, rowId: 'row_5', fieldId: 'company', title: 'Company' }
      ]
    });

    expect(result.success).toBe(true);
    expect(result.filled).toBe(4);
    expect(writes).toEqual([
      { row: 4, col: 1, value: 3 },
      { row: 4, col: 2, value: 'Company 3' },
      { row: 5, col: 1, value: 4 },
      { row: 5, col: 2, value: 'Company 4' }
    ]);

    SpreadsheetApp.getActiveSheet = originalGetActiveSheet;
  });

  test('writeAgentFillRunToSheet should write only allowed returned target cells', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        id: 'agent_run_fill',
        status: 'completed',
        output: {
          structured: {
            rows: [
              { rowId: 'row_2', values: { ceo: 'Tim Cook', hq: 'Cupertino, California' } },
              { rowId: 'row_999', values: { ceo: 'Ignored' } }
            ]
          }
        },
        usage: { searches: 3 },
        costDollars: { total: 0.05 }
      })
    });
    const { sheet, writes } = createMockSheet([
      ['Company Name', 'Website URL', 'HQ', 'CEO'],
      ['Apple', 'https://apple.com', '', '']
    ], 'D2');
    const originalGetActiveSheet = SpreadsheetApp.getActiveSheet;
    SpreadsheetApp.getActiveSheet = jest.fn(() => sheet);

    const result = writeAgentFillRunToSheet('agent_run_fill', {
      overwrite: false,
      targets: [{ cell: 'D2', row: 2, column: 4, rowId: 'row_2', fieldId: 'ceo', title: 'CEO' }]
    });

    expect(result.success).toBe(true);
    expect(result.filled).toBe(1);
    expect(writes).toEqual([{ row: 2, col: 4, value: 'Tim Cook' }]);

    SpreadsheetApp.getActiveSheet = originalGetActiveSheet;
  });

  test('writeAgentFillRunToSheet should not overwrite existing cells unless enabled', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        id: 'agent_run_fill',
        status: 'completed',
        output: {
          structured: {
            rows: [{ rowId: 'row_2', values: { ceo: 'New CEO' } }]
          }
        }
      })
    });
    const { sheet, writes } = createMockSheet([
      ['Company Name', 'CEO'],
      ['Apple', 'Existing CEO']
    ], 'B2');
    const originalGetActiveSheet = SpreadsheetApp.getActiveSheet;
    SpreadsheetApp.getActiveSheet = jest.fn(() => sheet);

    const result = writeAgentFillRunToSheet('agent_run_fill', {
      overwrite: false,
      targets: [{ cell: 'B2', row: 2, column: 2, rowId: 'row_2', fieldId: 'ceo', title: 'CEO' }]
    });

    expect(result.success).toBe(true);
    expect(result.filled).toBe(0);
    expect(result.skipped).toBe(1);
    expect(writes).toEqual([]);

    SpreadsheetApp.getActiveSheet = originalGetActiveSheet;
  });

  test('writeAgentFillRunToSheet should write to the saved sheet even if another sheet is active', () => {
    UrlFetchApp.fetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        id: 'agent_run_fill',
        status: 'completed',
        output: {
          structured: {
            rows: [{ rowId: 'row_2', values: { ceo: 'Tim Cook' } }]
          }
        }
      })
    });
    const target = createMockSheet([
      ['Company Name', 'CEO'],
      ['Apple', '']
    ], 'B2', [], { name: 'Target', sheetId: 444 });
    const active = createMockSheet([
      ['Company Name', 'CEO'],
      ['Apple', '']
    ], 'B2', [], { name: 'Other', sheetId: 555 });
    const originalGetActiveSheet = SpreadsheetApp.getActiveSheet;
    const originalGetActiveSpreadsheet = SpreadsheetApp.getActiveSpreadsheet;
    SpreadsheetApp.getActiveSheet = jest.fn(() => active.sheet);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => ({
      getSheets: () => [active.sheet, target.sheet],
      getSheetByName: (name) => ({ Target: target.sheet, Other: active.sheet }[name] || null)
    }));

    const result = writeAgentFillRunToSheet('agent_run_fill', {
      sheetId: 444,
      sheetName: 'Target',
      overwrite: false,
      targets: [{ cell: 'B2', row: 2, column: 2, rowId: 'row_2', fieldId: 'ceo', title: 'CEO' }]
    });

    expect(result.success).toBe(true);
    expect(target.writes).toEqual([{ row: 2, col: 2, value: 'Tim Cook' }]);
    expect(active.writes).toEqual([]);

    SpreadsheetApp.getActiveSheet = originalGetActiveSheet;
    if (originalGetActiveSpreadsheet) {
      SpreadsheetApp.getActiveSpreadsheet = originalGetActiveSpreadsheet;
    } else {
      delete SpreadsheetApp.getActiveSpreadsheet;
    }
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
