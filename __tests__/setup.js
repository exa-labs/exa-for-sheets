// Mock Google Apps Script APIs for testing

// Storage for mocked properties
const mockUserProperties = {};

// Mock PropertiesService
global.PropertiesService = {
  getUserProperties: () => ({
    getProperty: (key) => mockUserProperties[key] || null,
    setProperty: (key, value) => {
      mockUserProperties[key] = value;
    },
    deleteProperty: (key) => {
      delete mockUserProperties[key];
    }
  })
};

// Mock UrlFetchApp
global.UrlFetchApp = {
  fetch: jest.fn()
};

// Mock SpreadsheetApp
global.SpreadsheetApp = {
  getUi: () => ({
    createMenu: () => ({
      addItem: () => ({ addSeparator: () => ({ addItem: () => ({ addToUi: () => {} }) }) }),
      addToUi: () => {}
    }),
    showSidebar: () => {},
    alert: () => {},
    ButtonSet: { OK: 'OK' }
  }),
  getActiveSheet: () => ({
    getActiveRange: () => null,
    getRange: () => ({
      getFormulas: () => [],
      getValues: () => [],
      getValue: () => '',
      setValue: () => {},
      setFormula: () => {},
      clear: () => {},
      getCell: () => ({
        setValue: () => {},
        setFormula: () => {}
      }),
      getRow: () => 1,
      getColumn: () => 1
    }),
    getMaxRows: () => 1000,
    getMaxColumns: () => 26
  }),
  flush: () => {}
};

// Mock HtmlService
global.HtmlService = {
  createHtmlOutputFromFile: () => ({
    setTitle: () => ({})
  })
};

// Mock Utilities
global.Utilities = {
  getUuid: () => 'test-uuid-' + Math.random().toString(36).substr(2, 9),
  sleep: jest.fn()
};

// Mock Logger
global.Logger = {
  log: jest.fn()
};

// Helper to reset mocks between tests
global.resetMocks = () => {
  Object.keys(mockUserProperties).forEach(key => delete mockUserProperties[key]);
  jest.clearAllMocks();
};

// Helper to set mock properties
global.setMockProperty = (key, value) => {
  mockUserProperties[key] = value;
};

// Helper to get mock properties
global.getMockProperty = (key) => {
  return mockUserProperties[key];
};
