# Email Newsletters Summarizer

A Google Apps Script that automatically summarizes your email newsletters using Gemini AI and outputs them to a Google Doc.

## Features

- Fetches newsletters from Gmail by label
- Summarizes content using Gemini 2.5 Flash API
- Outputs formatted summaries to Google Docs
- Skips admin/transactional emails automatically
- Tracks processed emails to avoid duplicates
- Progress logging for debugging

## Setup

### 1. Create a Google Apps Script Project

1. Go to [script.google.com](https://script.google.com)
2. Create a new project
3. Copy the contents of `email-newsletters-summarizer.js` into the editor

### 2. Configure Script Properties

1. Go to **Project Settings** (gear icon)
2. Under **Script Properties**, add:
   - `GEMINI_API_KEY`: Your Gemini API key from [Google AI Studio](https://aistudio.google.com/apikey)
   - `GOOGLE_DOC_ID`: The ID of your target Google Doc (from the URL: `docs.google.com/document/d/YOUR_DOC_ID/edit`)

### 3. Update Configuration

Edit the constants at the top of the script:

```javascript
const GMAIL = {
  LABEL: 'Newsletters',  // Your Gmail label
  NEWER_THAN: '7d',      // Time range
  // ...
};
```

**Note:** The Google Doc ID is now configured in Script Properties (step 2), not in the code.

### 4. Set Up Gmail Filter

Create a Gmail filter to label incoming newsletters:
1. Gmail Settings → Filters → Create new filter
2. Set criteria (e.g., `(list:(substack.com) OR from:(tldrnewsletter.com)) -subject:("payment" OR "receipt" OR "subscription" OR "subscribe" OR "unsubscribed")` in the 'Has the words' field)
3. Apply label (e.g., `Newsletters`)

### 5. Run Manually or Schedule

- **Manual**: Click Run → `runDailyDigest`
- **Scheduled**: Triggers → Add trigger → `runDailyDigest` → Time-driven

## Configuration Options

| Section | Option | Description |
|---------|--------|-------------|
| `GMAIL` | `LABEL` | Gmail label to search |
| `GMAIL` | `NEWER_THAN` | Time range (e.g., `7d`, `30d`) |
| `GMAIL` | `MAX_THREADS` | Max emails per run |
| `GEMINI` | `MAX_OUTPUT_TOKENS` | Summary length limit |
| `GEMINI` | `DELAY_MS` | Delay between API calls |
| `PROCESSING` | `MIN_CONTENT_LENGTH` | Skip short emails |

**Note on Gemini Model**: The script uses `gemini-2.5-flash` (line 29 in the code). If this model becomes unavailable, you can update the `API_URL` in the `GEMINI` config to use a different model like `gemini-1.5-flash` or check [Google AI Studio](https://aistudio.google.com) for the latest available models.

## License

MIT
