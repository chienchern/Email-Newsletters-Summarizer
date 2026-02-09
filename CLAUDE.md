# Email Newsletters Summarizer - Project Guidelines

This file contains project-specific guidelines for Claude Code when working on the Email Newsletters Summarizer project. These rules supplement the global `~/.claude/CLAUDE.md` guidelines.

## 1. Documentation Standards

### README Synchronization Rule

**CRITICAL:** The README.md must be updated whenever code changes affect documented functionality. Before committing any code changes, verify these sections remain accurate:

- **Features list** (README.md lines 5-12): Update when adding/removing capabilities
- **Setup instructions** (README.md lines 14-54): Update when installation steps change
- **Configuration Options table** (README.md lines 57-65): Update when adding/removing Script Properties or changing their purpose
- **Code examples** (README.md lines 32-39): Update when API calls or configuration methods change
- **Model references** (README.md line 66): Update when changing Claude model versions

**When to update:**
- Adding or modifying configuration constants
- Changing Script Properties requirements
- Modifying feature behavior
- Updating API endpoints or parameters
- Changing the Claude model being used

**Best practice:** Include README updates in the same commit as the related code changes.

## 2. Security & Configuration Management

**Security Rules:**
- Never hardcode sensitive values (API keys, Google Doc IDs, email addresses) in code
- Always use Script Properties (`PropertiesService.getScriptProperties()`) for configuration
- When adding new configuration values, update README Section 2 with setup instructions
- Verify `.gitignore` excludes any local credential files

**Configuration Pattern:**
```javascript
const ANTHROPIC_API_KEY = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
const GOOGLE_DOC_ID = PropertiesService.getScriptProperties().getProperty('GOOGLE_DOC_ID');
```

## 3. Code Style & Structure

**Architecture:**
- Single-file architecture (`email-newsletters-summarizer.js`) is intentional due to Google Apps Script constraints
- Do not split into multiple files unless Google Apps Script capabilities change

**Organization:**
- Keep configuration constants at the top (lines 1-86 in current version)
- Maintain existing function organization pattern:
  1. Configuration constants
  2. Main entry point (`runDailyDigest`)
  3. Email processing functions
  4. API integration functions
  5. Document writing functions
  6. Helper/utility functions

**Style Conventions:**
- Use descriptive function names that indicate purpose
- Follow existing indentation and formatting style (2-space indentation)
- Maintain consistent error handling patterns
- Use `Logger.log()` for debugging output

## 4. Testing Approach

**No Automated Tests:**
This project has no automated test suite. All testing is manual using the Google Apps Script editor.

**Pre-Commit Testing Checklist:**
1. Open the Apps Script editor
2. Run `runDailyDigest()` function
3. Verify console logs show expected behavior
4. Test with sample emails (both plain text and HTML)
5. Check the Google Doc output for correct formatting
6. Verify all Script Properties are being read correctly

**Test Scenarios:**
- Empty inbox (no newsletters)
- Single newsletter
- Multiple newsletters
- HTML-formatted emails
- Plain text emails
- Edge cases (very long emails, special characters)

## 5. Commit Conventions

**Follow Global CLAUDE.md:**
- Use conventional commit style (concise subject line, optional body)
- Never include Claude session links or URLs in commit messages

**Project-Specific Conventions:**
- When updating both code and README, use a single commit with a clear message
- Example: `feat: add email filtering by sender` (includes both code changes and README updates)
- When fixing bugs, reference the specific function or feature area
- Example: `fix: handle HTML entities in email subject lines`

**Commit Message Structure:**
```
<type>: <brief description>

Optional body explaining why the change was made and any
important context about the implementation.
```

**Common types:** `feat`, `fix`, `docs`, `refactor`, `chore`

---

## Quick Reference

- **Main code file:** `email-newsletters-summarizer.js`
- **Documentation:** `README.md`
- **Security:** All secrets in Script Properties (never in code)
- **Testing:** Manual testing via Apps Script editor
- **Architecture:** Single-file Google Apps Script
