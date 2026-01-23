# Changelog

All notable changes to Exa for Google Sheets will be documented in this file.

## [1.1.0] - 2026-01-22

### New Features

**Simplified =EXA() Function**
A new streamlined function for data enrichment that's easier to use than the full EXA_ANSWER. Just describe what information you want about the data in a cell.

```
=EXA("Return only the company website URL", A1)
=EXA("Return only the company headcount", A1)
=EXA("Return only the CEO name", A1)
=EXA("Return the Amazon rating of this product", A1)
```

**Convert to Values (Freeze Results)**
New batch operation in the sidebar that converts Exa function results to static values. This prevents automatic recalculation and unexpected API charges when the sheet refreshes.

**Domain Filtering for EXA_SEARCH**
Added `includeDomainsStr` and `excludeDomainsStr` parameters to EXA_SEARCH, allowing you to restrict or exclude specific domains from search results.

```
=EXA_SEARCH("AI startups", 5, "auto", "", "", "linkedin.com,crunchbase.com")
=EXA_SEARCH("machine learning", 5, "auto", "", "", "", "wikipedia.org")
```

**Category Filtering for EXA_SEARCH**
Added `category` parameter to filter results by content type. Available categories: "company", "research paper", "news", "github", "tweet", "personal site", "pdf", "financial report", "people".

```
=EXA_SEARCH("OpenAI", 5, "auto", "", "", "", "", "company")
=EXA_SEARCH("transformer architecture", 5, "auto", "", "", "", "", "research paper")
=EXA_SEARCH("GPT-5 release", 5, "auto", "", "", "", "", "news")
```

**Rate Limiting with Automatic Retry**
API calls now automatically retry with exponential backoff when rate limited (HTTP 429). The add-on will retry up to 3 times with increasing delays, respecting the Retry-After header when provided.

### Improvements

**Better Citation Formatting**
Citations in EXA_ANSWER responses are now displayed in a clean numbered list at the end of the answer, rather than scattered inline throughout the text.

Before:
```
The company was founded in 2020 ([Source](url1)) and has raised $50M ([Source](url2)).
```

After:
```
The company was founded in 2020 and has raised $50M.

Sources:
1. [Company Profile](url1)
2. [Funding News](url2)
```

**API Usage Documentation**
Added prominent warnings and tips in the sidebar documentation about controlling API usage and preventing unexpected charges from auto-refresh behavior.

**Clearer Rate Limit Error Messages**
When rate limits are exceeded after all retries, the error message now clearly indicates the issue: "API Error: Rate limit exceeded. Please wait a moment and try again."

### Bug Fixes

**Fixed =EXA() Function Recognition in Batch Operations**
The "Convert to Values" and "Refresh Selected Cells" operations now correctly recognize the new simplified =EXA() function, not just the =EXA_ANSWER, =EXA_SEARCH, etc. functions.

### Technical

**Unit Test Suite**
Added comprehensive Jest test suite with 48 tests covering API key management, all Exa functions, rate limiting, and helper functions. Run with `npm test`.

---

## [1.0.0] - Initial Release

Initial release of Exa for Google Sheets with the following functions:
- EXA_ANSWER: AI-powered answers from web searches
- EXA_SEARCH: Web search returning URLs
- EXA_CONTENTS: Extract content from URLs
- EXA_FINDSIMILAR: Find similar web pages
