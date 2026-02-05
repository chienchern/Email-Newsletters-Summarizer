// ============================================================================
// CONFIGURATION
// ============================================================================

// --- Gmail ---
const GMAIL = {
  LABEL: 'Newsletters',
  NEWER_THAN: '8d',
  MAX_THREADS: 50,
  EXCLUDE_SUBJECTS: [
    // Add phrases to exclude from Gmail search
    // Example: 'payment receipt', 'is going live'
  ],
  SKIP_PATTERNS: [
    // Regex patterns to skip admin/transactional emails by subject
    /unsubscribe/i,
    /subscription (confirmed|updated|cancelled)/i,
    /preferences (updated|saved|changed)/i,
    /successfully (removed|unsubscribed)/i,
    /confirm your (email|subscription)/i,
    /welcome to .* newsletter/i,
    /you('ve| have) been (added|removed)/i,
    /manage your subscription/i,
  ],
};

// --- Gemini API ---
const GEMINI = {
  API_URL: 'https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent',
  MAX_OUTPUT_TOKENS: 2048,
  DELAY_MS: 2000,
  PROMPT: `
You are an executive assistant. Distill this newsletter into key takeaways only.

INPUT:
{{CONTENT}}

RULES:
- Maximum 10 bullet points total
- Each bullet: 1-2 sentences max
- Bold the topic (e.g., "**Topic:** key insight")
- Only include genuinely important or actionable information
- Skip fluff, intros, outros, and promotional content
- If nothing valuable, return "STATUS: SKIP"

Be ruthlessly concise. Output Markdown.
`,
};

// --- Processing ---
const PROCESSING = {
  MAX_CONTENT_LENGTH: 25000,   // Max chars to send to API
  MIN_CONTENT_LENGTH: 500,     // Skip emails shorter than this
  MIN_SUMMARY_LENGTH: 50,      // Skip summaries shorter than this
};

// --- Storage ---
const STORAGE = {
  PROCESSED_IDS_KEY: 'PROCESSED_MESSAGE_IDS',
  MAX_STORED_IDS: 500,
};

// --- Google Doc ---
const DOC = {
  ID: '1pkvgDMwSNC-Ccw8WLh8ODlLhNY81M96SBvzz--Y4Tjs',
  STYLES: {
    header: {
      [DocumentApp.Attribute.FONT_FAMILY]: 'Roboto',
      [DocumentApp.Attribute.FOREGROUND_COLOR]: '#444444',
    },
    title: {
      [DocumentApp.Attribute.FONT_FAMILY]: 'Roboto',
      [DocumentApp.Attribute.FOREGROUND_COLOR]: '#1155cc',
    },
    footer: {
      [DocumentApp.Attribute.FONT_FAMILY]: 'Roboto',
      [DocumentApp.Attribute.FONT_SIZE]: 8,
      [DocumentApp.Attribute.FOREGROUND_COLOR]: '#666666',
    },
    body: {
      [DocumentApp.Attribute.FONT_FAMILY]: 'Merriweather',
      [DocumentApp.Attribute.FONT_SIZE]: 10,
      [DocumentApp.Attribute.LINE_SPACING]: 1.15,
    },
  },
};

// --- Helpers ---
function buildGmailQuery() {
  const base = `label:${GMAIL.LABEL} newer_than:${GMAIL.NEWER_THAN}`;
  const exclusions = GMAIL.EXCLUDE_SUBJECTS
    .map(s => `-subject:"${s}"`)
    .join(' ');
  return exclusions ? `${base} ${exclusions}` : base;
}

// ====================
// MAIN ENTRY POINT
// ====================

function runDailyDigest() {
  console.log('=== Starting Daily Digest ===');

  const apiKey = getApiKey();
  if (!apiKey) return;

  const doc = DocumentApp.openById(DOC.ID);
  const body = doc.getBody();
  const processedIds = getProcessedIds();
  console.log(`Loaded ${processedIds.size} previously processed message IDs`);

  insertDateHeader(body);

  console.log(`Searching emails: ${buildGmailQuery()}`);
  const threads = GmailApp.search(buildGmailQuery(), 0, GMAIL.MAX_THREADS);
  console.log(`Found ${threads.length} threads`);

  if (threads.length === 0) {
    console.log('No threads to process');
    insertEmptyState(body);
    doc.saveAndClose();
    return;
  }

  const newlyProcessedIds = processThreads(threads, body, processedIds, apiKey);

  if (newlyProcessedIds.length > 0) {
    saveProcessedIds(processedIds, newlyProcessedIds);
  }

  console.log(`=== Completed: ${newlyProcessedIds.length} emails summarized ===`);
  doc.saveAndClose();
}

// ====================
// THREAD PROCESSING
// ====================

function processThreads(threads, body, processedIds, apiKey) {
  const newlyProcessedIds = [];
  const total = threads.length;

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const progress = `[${i + 1}/${total}]`;
    const result = processSingleThread(thread, body, processedIds, apiKey, progress);

    if (result.processed) {
      newlyProcessedIds.push(result.messageId);
    }
  }

  return newlyProcessedIds;
}

function processSingleThread(thread, body, processedIds, apiKey, progress) {
  const msg = getLatestMessage(thread);
  const subject = msg.getSubject();
  const messageId = msg.getId();

  console.log(`${progress} Processing: "${subject}"`);

  if (processedIds.has(messageId)) {
    console.log(`  ↳ Skipped: already processed`);
    return { processed: false };
  }

  if (isAdminEmail(subject)) {
    console.log(`  ↳ Skipped: admin/transactional email`);
    return { processed: false };
  }

  const content = extractEmailContent(msg);
  if (content.length < PROCESSING.MIN_CONTENT_LENGTH) {
    console.log(`  ↳ Skipped: content too short (${content.length} chars)`);
    return { processed: false };
  }

  try {
    console.log(`  ↳ Calling Gemini API...`);
    const result = callGeminiAPI(apiKey, content);

    if (result.error) {
      console.error(`  ↳ API error: ${result.error}`);
      return { processed: false };
    }

    if (result.summary.includes("STATUS: SKIP")) {
      console.log(`  ↳ Skipped: marked as skip by API`);
      return { processed: false };
    }

    if (result.summary.length < PROCESSING.MIN_SUMMARY_LENGTH) {
      console.log(`  ↳ Skipped: summary too short (${result.summary.length} chars)`);
      return { processed: false };
    }

    insertArticleBlock(body, {
      subject: subject,
      sender: msg.getFrom(),
      permalink: thread.getPermalink(),
      summary: result.summary
    });

    thread.markRead();
    console.log(`  ↳ Success: summarized (${result.summary.length} chars)`);
    Utilities.sleep(GEMINI.DELAY_MS);

    return { processed: true, messageId };

  } catch (e) {
    console.error(`  ↳ Error: ${e.message}`);
    return { processed: false };
  }
}

function getLatestMessage(thread) {
  const messages = thread.getMessages();
  return messages[messages.length - 1];
}

/**
 * Extract email content, preferring plain text but falling back to HTML
 */
function extractEmailContent(msg) {
  let content = msg.getPlainBody();

  // If plain text is too short, try HTML and strip tags
  if (content.length < PROCESSING.MIN_CONTENT_LENGTH) {
    const htmlBody = msg.getBody();
    if (htmlBody) {
      content = stripHtmlTags(htmlBody);
    }
  }

  return content.substring(0, PROCESSING.MAX_CONTENT_LENGTH);
}

/**
 * Strip HTML tags and clean up whitespace
 */
function stripHtmlTags(html) {
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Check if email is administrative/transactional (not a newsletter)
 */
function isAdminEmail(subject) {
  return GMAIL.SKIP_PATTERNS.some(pattern => pattern.test(subject));
}

// ====================
// DOCUMENT INSERTION
// ====================

function insertDateHeader(body) {
  const dateStr = new Date().toLocaleDateString('en-US', {
    weekday: 'long',
    month: 'short',
    day: 'numeric'
  });

  const header = body.insertParagraph(0, `\n📅 INTELLIGENCE BRIEF: ${dateStr}`);
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  header.setAttributes(DOC.STYLES.header);
}

function insertEmptyState(body) {
  body.insertParagraph(1, "No new updates.")
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setItalic(true);
}

function insertArticleBlock(body, article) {
  // Insert in reverse order since we're prepending at index 1
  body.insertHorizontalRule(1);
  insertFooter(body, article.sender, article.permalink);
  insertMarkdownContent(body, article.summary, 1);
  insertTitle(body, article.subject);
}

function insertTitle(body, subject) {
  const title = body.insertParagraph(1, subject);
  title.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  title.setAttributes(DOC.STYLES.title);
}

function insertFooter(body, sender, permalink) {
  const footerText = `Source: ${sender} | `;
  const linkText = "Open Email";

  const footer = body.insertParagraph(1, footerText);
  footer.appendText(linkText);
  footer.editAsText().setLinkUrl(footerText.length, footerText.length + linkText.length - 1, permalink);
  footer.setAttributes(DOC.STYLES.footer);
}

// ====================
// MARKDOWN PARSING
// ====================

function insertMarkdownContent(body, text, index) {
  const lines = text.split('\n').reverse();

  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line) continue;

    const parsed = parseLine(line);
    const element = insertLineElement(body, index, parsed);

    element.setAttributes(DOC.STYLES.body);
    applyBoldRanges(element, parsed.boldRanges, parsed.text);
  }
}

function parseLine(line) {
  const isBullet = line.startsWith('* ') || line.startsWith('- ');
  const content = isBullet ? line.substring(2) : line;
  const { text, boldRanges } = extractBoldRanges(content);

  return { text, boldRanges, isBullet };
}

function extractBoldRanges(line) {
  const boldRanges = [];
  const boldPattern = /\*\*(.*?)\*\*/g;
  let match;
  let cleanLine = '';
  let lastIndex = 0;

  while ((match = boldPattern.exec(line)) !== null) {
    cleanLine += line.substring(lastIndex, match.index);

    const boldStart = cleanLine.length;
    cleanLine += match[1];
    const boldEnd = cleanLine.length - 1;

    boldRanges.push({ start: boldStart, end: boldEnd });
    lastIndex = match.index + match[0].length;
  }

  cleanLine += line.substring(lastIndex);

  return { text: cleanLine, boldRanges };
}

function insertLineElement(body, index, parsed) {
  if (parsed.isBullet) {
    const item = body.insertListItem(index, parsed.text);
    item.setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET);
    return item;
  }
  return body.insertParagraph(index, parsed.text);
}

function applyBoldRanges(element, boldRanges, text) {
  const textElement = element.editAsText();

  for (const range of boldRanges) {
    if (range.start <= range.end && range.end < text.length) {
      textElement.setBold(range.start, range.end, true);
    }
  }
}

// ====================
// GEMINI API
// ====================

function callGeminiAPI(key, text) {
  const prompt = buildPrompt(text);
  const response = makeApiRequest(key, prompt);
  return parseApiResponse(response);
}

function buildPrompt(text) {
  return GEMINI.PROMPT.replace('{{CONTENT}}', text);
}

function makeApiRequest(key, prompt) {
  const url = `${GEMINI.API_URL}?key=${key}`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.3,
      maxOutputTokens: GEMINI.MAX_OUTPUT_TOKENS
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  return UrlFetchApp.fetch(url, options);
}

function parseApiResponse(response) {
  try {
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode !== 200) {
      const errorJson = JSON.parse(responseText);
      return { summary: '', error: errorJson.error?.message || `HTTP ${responseCode}` };
    }

    const json = JSON.parse(responseText);

    if (json.error) {
      return { summary: '', error: json.error.message };
    }

    if (json.candidates?.[0]?.finishReason === 'SAFETY') {
      return { summary: '', error: 'Content blocked by safety filters' };
    }

    const summary = json.candidates?.[0]?.content?.parts?.[0]?.text;
    if (!summary) {
      return { summary: '', error: 'No content in API response' };
    }

    return { summary, error: null };

  } catch (e) {
    return { summary: '', error: `Request failed: ${e.message}` };
  }
}

// ====================
// PERSISTENCE
// ====================

function getApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  if (!apiKey) {
    console.error('GEMINI_API_KEY not set in Script Properties');
    return null;
  }

  return apiKey;
}

function getProcessedIds() {
  const stored = PropertiesService.getScriptProperties().getProperty(STORAGE.PROCESSED_IDS_KEY);

  if (!stored) return new Set();

  try {
    return new Set(JSON.parse(stored));
  } catch (e) {
    return new Set();
  }
}

function saveProcessedIds(existingSet, newIds) {
  const allIds = [...existingSet, ...newIds];
  const trimmedIds = allIds.slice(-STORAGE.MAX_STORED_IDS);

  PropertiesService.getScriptProperties().setProperty(
    STORAGE.PROCESSED_IDS_KEY,
    JSON.stringify(trimmedIds)
  );
}
