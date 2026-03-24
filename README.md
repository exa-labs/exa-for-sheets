# Exa Google Sheets Extension

## Description

This Google Apps Script integration brings the power of the Exa API directly into Google Sheets, allowing you to search the web, query information, find similar content, and extract website content without leaving your spreadsheet.

## Features

### Custom Sheet Functions
* **EXA:** Simplified function for quick data enrichment - just describe what you want
* **EXA_ANSWER:** Query the Exa AI with advanced options (system prompts, structured output)
* **EXA_SEARCH:** Search the web with domain and category filtering
* **EXA_CONTENTS:** Extract the text content from specific URLs
* **EXA_FINDSIMILAR:** Find URLs similar to a provided reference URL

### Sidebar Interface
* **API Key Management:** Securely save your Exa API key
* **Documentation:** Built-in reference for all Exa functions and their parameters
* **Batch Refresh:** Refresh multiple Exa function cells at once
* **Convert to Values:** Freeze results to prevent auto-refresh charges

## Setup & Installation

1.  **Open your Google Sheet:** Go to the Google Sheet where you want to use this script.
2.  **Open the Script Editor:** Click on "Extensions" > "Apps Script".
3.  **Copy the Code:**
    *   Copy the contents of `Code.gs` from this project and paste it into the `Code.gs` file in the Apps Script editor. Overwrite any existing template code.
    *   Create a new HTML file in the editor (File > New > HTML file). Name it `Sidebar.html` (ensure the name matches exactly, including capitalization).
    *   Copy the contents of the `Sidebar.html` file from this project and paste it into the newly created `Sidebar.html` file in the editor.
4.  **Save the Project:** Click the floppy disk icon (Save project) and give your script project a name (e.g., "Exa for Sheets").
5.  **Refresh Your Sheet:** Close the Apps Script editor tab and refresh your Google Sheet browser tab.
6.  **Authorize the Script:**
    *   After refreshing, a new custom menu item "Exa AI" should appear. Click on it, then select "Open Sidebar".
    *   A dialog box will pop up asking for authorization. Review the permissions requested (it will need access to external services, script properties, and the current spreadsheet) and click "Allow". You might need to go through an "Advanced" > "Go to (project name)" flow if Google warns it's an unverified app.
7.  **Set Your API Key:**
    *   Once authorized, the sidebar will open with the API Key tab active.
    *   Go to [https://exa.ai/](https://exa.ai/) to get your API key if you don't have one.
    *   Paste your Exa API key into the input field in the sidebar and click "Save API Key".
    *   You should see a success message "API key saved successfully.".

## Using the Sidebar

The sidebar offers three main tabs:

### 1. API Key
* Set or update your Exa API key
* View status of key operations

### 2. Batch
* Refresh selected cells containing Exa functions
* Status updates for refresh operations

### 3. Documentation
* Comprehensive documentation for all Exa functions
* Parameter descriptions and return value information
* Quick reference for function syntax

## Using Exa Functions

### EXA (Simplified)
```
=EXA(prompt, [context])
```
The easiest way to enrich data. Just describe what information you want.

**Examples:**
```
=EXA("Return the CEO name", A1)
=EXA("Return the company website URL", A1)
=EXA("Return the Amazon rating of this product", A1)
```

### EXA_ANSWER
```
=EXA_ANSWER(prompt, [prefix], [suffix], [includeCitations], [systemPrompt], [outputSchema], [returnRawJson])
```
Advanced AI answers with full control over output format.

**Parameters:**
* `prompt` (required): The main question or prompt
* `prefix` (optional): Text to add before the main prompt
* `suffix` (optional): Text to add after the main prompt
* `includeCitations` (optional): If TRUE, includes source citations (Default: FALSE)
* `systemPrompt` (optional): Control output format (e.g., "only return a number")
* `outputSchema` (optional): JSON schema for structured output. Generate at https://dashboard.exa.ai/playground/answer
* `returnRawJson` (optional): If TRUE with outputSchema, returns raw JSON instead of extracted value

**Examples:**
```
=EXA_ANSWER("OpenAI CEO", "", "", FALSE, "only return a name")
=EXA_ANSWER("ceo of exa.ai", "", "", FALSE, "", "{""type"":""object"",""properties"":{""name"":{""type"":""string""}}}")
```

### EXA_SEARCH
```
=EXA_SEARCH(query, [numResults], [searchType], [prefix], [suffix], [includeDomainsStr], [excludeDomainsStr], [category])
```
Searches the web and returns a vertical list of URLs.

**Parameters:**
* `query` (required): The search query
* `numResults` (optional): Number of results to return (1-10, Default: 1)
* `searchType` (optional): "auto", "neural", or "keyword" (Default: "auto")
* `prefix` (optional): Text to add before the main query
* `suffix` (optional): Text to add after the main query
* `includeDomainsStr` (optional): Comma-separated domains to include (e.g., "linkedin.com,crunchbase.com")
* `excludeDomainsStr` (optional): Comma-separated domains to exclude
* `category` (optional): Filter by type - "company", "research paper", "news", "github", "pdf", etc.

**Examples:**
```
=EXA_SEARCH("AI startups", 5, "auto", "", "", "linkedin.com,crunchbase.com")
=EXA_SEARCH("transformer architecture", 5, "auto", "", "", "", "", "research paper")
```

### EXA_CONTENTS
```
=EXA_CONTENTS(url)
```
Retrieves the text content from a specified URL.

**Parameters:**
* `url` (required): The full URL to extract content from (must start with http/https)

### EXA_FINDSIMILAR
```
=EXA_FINDSIMILAR(url, [numResults], [includeDomainsStr], [excludeDomainsStr], [includeTextStr], [excludeTextStr])
```
Finds URLs similar to the input URL.

**Parameters:**
* `url` (required): The reference URL to find similar content
* `numResults` (optional): Number of results to return (1-10, Default: 1)
* `includeDomainsStr` (optional): Comma-separated list of domains to include
* `excludeDomainsStr` (optional): Comma-separated list of domains to exclude
* `includeTextStr` (optional): Phrase that must appear in results
* `excludeTextStr` (optional): Phrase that must not appear in results

## Batch Refresh

To use batch refresh:

1. Select cells containing Exa functions in your sheet
2. Open the sidebar and navigate to the "Batch" tab
3. Click **Refresh Selected Cells** to re-execute the Exa functions in selected cells

## Notes

* Exa API requests count against your Exa usage quota
* For best performance, avoid excessive function calls in large sheets
* Functions will automatically refresh when their inputs change or when the sheet is reopened
* Use "Convert to Values" in the sidebar to freeze results and prevent auto-refresh charges
* Rate limiting: The add-on automatically retries up to 3 times with exponential backoff on rate limit errors

---

## Privacy & Security

* Your Exa API key is stored securely using Google Apps Script's User Properties service
* The key is only accessible to your Google account
* No data is stored outside of your Google account and the Exa API
* View our [Privacy Policy](https://exa.ai/exa-for-sheets/privacy-policy) for details on data handling and Google API compliance

## Development

1. Install dependencies:
   ```bash
   npm install
   ```

2. Login to Google:
   ```bash
   npm run login
   ```

3. Create a new Google Apps Script project:
   ```bash
   npm run create
   ```

4. Push the code:
   ```bash
   npm run push
   ```

5. Run tests:
   ```bash
   npm test
   ```

6. Bump version (updates package.json, Code.gs, and CHANGELOG.md):
   ```bash
   npm run bump patch  # or minor, major
   ```
