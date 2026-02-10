// ============================================================================
// CONFIGURATION
// ============================================================================

// --- Gmail ---
const GMAIL = {
  LABEL: 'Newsletters',
  NEWER_THAN: '1d',
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
  MAX_OUTPUT_TOKENS: 4096,
  DELAY_MS: 2000,
  PROMPT: `
You are an executive assistant. Distill this newsletter into key takeaways and classify it by theme.

INPUT:
{{CONTENT}}

RULES:
- Maximum 5 bullet points total
- Each bullet: 1-2 sentences max
- Bold the topic (e.g., "**Topic:** key insight")
- Only include genuinely important or actionable information
- Skip fluff, intros, outros, and promotional content
- Classify the newsletter into ONE theme category

OUTPUT FORMAT (respond with valid JSON - no markdown code blocks):
{
  "theme": "EXACTLY one of these strings: Tech News | Product Updates | Industry Analysis | AI & ML | Business Strategy | Developer Tools | Marketing | Finance | Design | Other",
  "summary": "Markdown formatted bullets here"
}

JSON FORMATTING - CRITICAL:
- Output ONLY the JSON object, no markdown code blocks
- Properly escape quotes in summary text (use \" for quotes inside strings)
- Keep newlines as \n within the JSON string
- Test that your output is valid JSON before returning

THEME CLASSIFICATION - CRITICAL:
Use the EXACT theme string from the list above. Do NOT paraphrase, rephrase, or create variations.
Examples:
- ✅ "Tech News" (correct)
- ❌ "Tech Headlines" (incorrect - will cause misclassification)
- ❌ "Technology News" (incorrect)

SPECIAL CASES:
- If nothing valuable, return: {"theme": "SKIP", "summary": "STATUS: SKIP"}

Be ruthlessly concise.
`,
  SYNTHESIS_PROMPT: `
You are an executive assistant. Synthesize key insights from multiple newsletter summaries on a common theme.

INPUT SUMMARIES:
{{SUMMARIES}}

RULES:
- Maximum 7 bullet points total
- Each bullet: 1-2 sentences max
- Bold the topic (e.g., "**Topic:** key insight")
- Find common threads across newsletters - don't just concatenate
- If multiple newsletters mention the same topic, synthesize into ONE bullet
- Prioritize the most important/actionable information
- Skip redundant or minor details

OUTPUT: Markdown formatted bullets only (no theme classification needed)

Be ruthlessly concise. This is an executive summary of summaries.
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

// --- Article Collection ---
const ARTICLE_COLLECTION = {
  THEME_ORDER: [
    'Tech News',
    'AI & ML',
    'Product Updates',
    'Developer Tools',
    'Business Strategy',
    'Industry Analysis',
    'Marketing',
    'Finance',
    'Design',
    'Other'
  ],
  DEFAULT_THEME: 'Other',
  THEME_KEYWORDS: {
    // For fuzzy matching when AI deviates from exact theme names
    'tech': 'Tech News',
    'technology': 'Tech News',
    'headlines': 'Tech News',
    'ai': 'AI & ML',
    'artificial intelligence': 'AI & ML',
    'machine learning': 'AI & ML',
    'ml': 'AI & ML',
    'product': 'Product Updates',
    'release': 'Product Updates',
    'launch': 'Product Updates',
    'tool': 'Developer Tools',
    'developer': 'Developer Tools',
    'business': 'Business Strategy',
    'strategy': 'Business Strategy',
    'industry': 'Industry Analysis',
    'analysis': 'Industry Analysis',
    'market': 'Industry Analysis',
    'marketing': 'Marketing',
    'finance': 'Finance',
    'financial': 'Finance',
    'design': 'Design'
  }
};

// --- Google Doc ---
const DOC = {
  ID: PropertiesService.getScriptProperties().getProperty('GOOGLE_DOC_ID'),
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

  // PHASE 1: Collect phase
  console.log('\n--- Phase 1: Collecting articles ---');
  const collectedArticles = processThreads(threads, processedIds, apiKey);

  if (collectedArticles.length === 0) {
    console.log('No new summaries to insert');
    insertEmptyState(body);
    doc.saveAndClose();
    return;
  }

  // PHASE 2: Group by theme
  console.log('\n--- Phase 2: Grouping by theme ---');
  const groupedByTheme = groupArticlesByTheme(collectedArticles);
  console.log(`Grouped into ${Object.keys(groupedByTheme).length} themes`);

  // PHASE 3: Synthesize cross-newsletter summaries (NEW)
  console.log('\n--- Phase 3: Synthesizing theme summaries ---');
  const synthesizedThemes = {};

  for (const [theme, articles] of Object.entries(groupedByTheme)) {
    const synthesizedSummary = synthesizeThemeSummary(apiKey, articles, theme);
    synthesizedThemes[theme] = {
      articles: articles,
      synthesizedSummary: synthesizedSummary
    };

    // Add delay between synthesis calls
    if (synthesizedSummary) {
      Utilities.sleep(GEMINI.DELAY_MS);
    }
  }

  // PHASE 4: Insert individual articles first
  console.log('\n--- Phase 4: Inserting individual articles ---');
  insertIndividualArticles(body, collectedArticles);

  // PHASE 5: Insert master summary (appears at top)
  console.log('\n--- Phase 5: Inserting master summary ---');
  insertMasterSummary(body, synthesizedThemes);

  // Extract message IDs for persistence
  const newlyProcessedIds = collectedArticles.map(a => a.messageId);
  if (newlyProcessedIds.length > 0) {
    saveProcessedIds(processedIds, newlyProcessedIds);
  }

  console.log(`\n=== Completed: ${collectedArticles.length} emails summarized across ${Object.keys(synthesizedThemes).length} themes ===`);
  doc.saveAndClose();
}

// ====================
// THREAD PROCESSING
// ====================

function processThreads(threads, processedIds, apiKey) {
  const collectedArticles = [];
  const total = threads.length;

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const progress = `[${i + 1}/${total}]`;
    const result = processSingleThread(thread, processedIds, apiKey, progress);

    if (result.article) {
      collectedArticles.push(result.article);
    }
  }

  return collectedArticles;
}

function processSingleThread(thread, processedIds, apiKey, progress) {
  const msg = getLatestMessage(thread);
  const subject = msg.getSubject();
  const messageId = msg.getId();

  console.log(`${progress} Processing: "${subject}"`);

  if (processedIds.has(messageId)) {
    console.log(`  ↳ Skipped: already processed`);
    return {};
  }

  if (isAdminEmail(subject)) {
    console.log(`  ↳ Skipped: admin/transactional email`);
    return {};
  }

  const content = extractEmailContent(msg);
  if (content.length < PROCESSING.MIN_CONTENT_LENGTH) {
    console.log(`  ↳ Skipped: content too short (${content.length} chars)`);
    return {};
  }

  try {
    console.log(`  ↳ Calling Gemini API...`);
    const prompt = buildNewsletterSummaryPrompt(content);
    const result = callGeminiAPIWithPrompt(apiKey, prompt);

    if (result.error) {
      console.error(`  ↳ API error: ${result.error}`);
      return {};
    }

    // NEW: Parse JSON response
    const parsed = parseApiJsonResponse(result.summary);

    if (!parsed || parsed.theme === 'SKIP') {
      console.log(`  ↳ Skipped: marked as skip by API`);
      return {};
    }

    if (parsed.summary.length < PROCESSING.MIN_SUMMARY_LENGTH) {
      console.log(`  ↳ Skipped: summary too short`);
      return {};
    }

    thread.markRead();
    console.log(`  ↳ Success: theme="${parsed.theme}", summary length=${parsed.summary.length} chars`);
    Utilities.sleep(GEMINI.DELAY_MS);

    return {
      article: {
        messageId: messageId,
        subject: subject,
        sender: msg.getFrom(),
        permalink: thread.getPermalink(),
        theme: parsed.theme,
        summary: parsed.summary,
        timestamp: msg.getDate()
      }
    };

  } catch (e) {
    console.error(`  ↳ Error: ${e.message}`);
    return {};
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
// ARTICLE COLLECTION & GROUPING
// ====================

function groupArticlesByTheme(articles) {
  const grouped = {};

  // Initialize all themes
  ARTICLE_COLLECTION.THEME_ORDER.forEach(theme => {
    grouped[theme] = [];
  });

  // Group articles
  articles.forEach(article => {
    const theme = article.theme || ARTICLE_COLLECTION.DEFAULT_THEME;
    if (!grouped[theme]) {
      grouped[theme] = [];
    }
    grouped[theme].push(article);
  });

  // Sort within each theme by timestamp (newest first)
  Object.keys(grouped).forEach(theme => {
    grouped[theme].sort((a, b) => b.timestamp - a.timestamp);
  });

  // Remove empty themes
  const result = {};
  ARTICLE_COLLECTION.THEME_ORDER.forEach(theme => {
    if (grouped[theme] && grouped[theme].length > 0) {
      result[theme] = grouped[theme];
    }
  });

  return result;
}

function insertMasterSummary(body, synthesizedThemes) {
  const themes = Object.keys(synthesizedThemes);
  if (themes.length === 0) return;

  // Insert divider before individual articles
  body.insertHorizontalRule(1);

  // Insert theme sections (in reverse since we prepend)
  themes.reverse().forEach(theme => {
    const themeData = synthesizedThemes[theme];

    // Insert synthesized bullets
    if (themeData.synthesizedSummary) {
      insertMarkdownContent(body, themeData.synthesizedSummary, 1);
    } else {
      // Fallback: list article subjects if synthesis failed
      themeData.articles.slice().reverse().forEach(article => {
        const bullet = body.insertListItem(1, article.subject);
        bullet.setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET);
        bullet.setAttributes(DOC.STYLES.body);
      });
    }

    // Insert theme header
    const count = themeData.articles.length;
    const countLabel = count === 1 ? '1 newsletter' : `${count} newsletters`;
    const themeHeader = body.insertParagraph(1, `${theme} (${countLabel})`);
    themeHeader.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    themeHeader.setAttributes(DOC.STYLES.header);
  });

  // Insert master summary title
  const title = body.insertParagraph(1, '🧭 Master Summary');
  title.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  title.setAttributes(DOC.STYLES.title);
}

function insertIndividualArticles(body, articles) {
  const sortedArticles = sortArticlesForDisplay(articles);

  // Insert in reverse (prepending at index 1, after master summary)
  sortedArticles.reverse().forEach(article => {
    insertSingleArticle(body, article);
  });
}

function sortArticlesForDisplay(articles) {
  const themeOrder = ARTICLE_COLLECTION.THEME_ORDER;

  return articles.sort((a, b) => {
    // Sort by theme order first
    const themeIndexA = themeOrder.indexOf(a.theme);
    const themeIndexB = themeOrder.indexOf(b.theme);

    if (themeIndexA !== themeIndexB) {
      return themeIndexA - themeIndexB;
    }

    // Then by timestamp (newest first within theme)
    return b.timestamp - a.timestamp;
  });
}

function insertSingleArticle(body, article) {
  body.insertHorizontalRule(1);
  insertFooter(body, article.sender, article.permalink);
  insertMarkdownContent(body, article.summary, 1);
  insertTitle(body, article.subject);
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

  const header = body.insertParagraph(0, `📅 INTELLIGENCE BRIEF: ${dateStr}`);
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  header.setAttributes(DOC.STYLES.header);
}

function insertEmptyState(body) {
  const timeStr = new Date().toLocaleTimeString('en-US', {
    hour: 'numeric',
    minute: '2-digit',
    timeZoneName: 'short'
  });
  body.insertParagraph(1, `No new newsletters today. (Checked at ${timeStr})`)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setItalic(true);
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

function callGeminiAPIWithPrompt(key, prompt) {
  const response = makeApiRequest(key, prompt);
  return parseApiResponse(response);
}

function buildNewsletterSummaryPrompt(content) {
  return GEMINI.PROMPT.replace('{{CONTENT}}', content);
}

function buildMasterSummaryPrompt(summaries) {
  return GEMINI.SYNTHESIS_PROMPT.replace('{{SUMMARIES}}', summaries);
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

    const finishReason = json.candidates?.[0]?.finishReason;

    if (finishReason === 'SAFETY') {
      return { summary: '', error: 'Content blocked by safety filters' };
    }

    if (finishReason === 'MAX_TOKENS') {
      console.warn('⚠️ Response truncated due to MAX_TOKENS limit');
      // Continue processing but log warning
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

function findBestThemeMatch(aiTheme) {
  const normalized = aiTheme.toLowerCase().trim();

  // Try exact match first (case-insensitive)
  const exactMatch = ARTICLE_COLLECTION.THEME_ORDER.find(
    t => t.toLowerCase() === normalized
  );
  if (exactMatch) return exactMatch;

  // Try keyword matching
  for (const [keyword, theme] of Object.entries(ARTICLE_COLLECTION.THEME_KEYWORDS)) {
    if (normalized.includes(keyword)) {
      console.warn(`Fuzzy matched "${aiTheme}" → "${theme}" via keyword "${keyword}"`);
      return theme;
    }
  }

  // No match found
  console.warn(`No match for theme "${aiTheme}", using default`);
  return ARTICLE_COLLECTION.DEFAULT_THEME;
}

function fixUnescapedQuotesInJson(jsonText) {
  // Fix malformed JSON by extracting fields and reconstructing
  // This handles cases where the AI returns unescaped quotes in the summary

  try {
    // Extract theme field
    const themeMatch = jsonText.match(/"theme"\s*:\s*"([^"]+)"/);
    if (!themeMatch) {
      return jsonText; // Can't fix, return as-is
    }
    const theme = themeMatch[1];

    // Extract summary field - more lenient pattern
    // Look for "summary": " and then grab everything until the final "}
    const summaryMatch = jsonText.match(/"summary"\s*:\s*"([\s\S]*)"[\s\n]*}/);
    if (!summaryMatch) {
      return jsonText; // Can't fix, return as-is
    }

    let summary = summaryMatch[1];

    // Remove any trailing quote and whitespace that might be part of the closing
    summary = summary.replace(/"\s*$/, '');

    // Now reconstruct valid JSON with properly escaped summary
    const fixedJson = {
      theme: theme,
      summary: summary
    };

    return JSON.stringify(fixedJson);

  } catch (e) {
    console.warn('Could not fix malformed JSON: ' + e.message);
    return jsonText;
  }
}

function parseApiJsonResponse(responseText) {
  try {
    // Strip markdown code blocks if present
    let cleanedText = responseText.trim();

    // Remove ```json and ``` markers
    if (cleanedText.startsWith('```json')) {
      cleanedText = cleanedText.replace(/^```json\s*\n?/, '').replace(/\n?```\s*$/, '');
    } else if (cleanedText.startsWith('```')) {
      cleanedText = cleanedText.replace(/^```\s*\n?/, '').replace(/\n?```\s*$/, '');
    }

    // First attempt: try parsing as-is
    let json;
    try {
      json = JSON.parse(cleanedText.trim());
    } catch (parseError) {
      // Second attempt: fix unescaped quotes in summary field
      console.warn('Initial JSON parse failed, attempting to fix unescaped quotes...');
      cleanedText = fixUnescapedQuotesInJson(cleanedText);
      json = JSON.parse(cleanedText.trim());
    }

    // Validate JSON structure
    if (!json || typeof json !== 'object') {
      console.error('API response is not a JSON object');
      console.error('Response preview: ' + responseText.substring(0, 200));
      return null;
    }

    if (!json.theme || !json.summary) {
      console.error('Invalid JSON structure - missing theme or summary');
      console.error('Received: ' + JSON.stringify(json).substring(0, 200));
      return null;
    }

    // Ensure summary is a string
    const summary = typeof json.summary === 'string'
      ? json.summary.trim()
      : String(json.summary).trim();

    const theme = String(json.theme).trim();

    // Handle SKIP special case
    if (theme === 'SKIP') {
      return { theme: 'SKIP', summary: summary };
    }

    // Use fuzzy matching to find best theme
    const matchedTheme = findBestThemeMatch(theme);

    return { theme: matchedTheme, summary: summary };

  } catch (e) {
    console.error(`Failed to parse JSON: ${e.message}`);
    console.error(`Response length: ${responseText.length} characters`);
    console.error('Response preview (first 500 chars): ' + responseText.substring(0, 500));
    console.error('Response end (last 200 chars): ' + responseText.substring(Math.max(0, responseText.length - 200)));

    // Fallback for non-JSON responses
    if (responseText.includes('STATUS: SKIP')) {
      return { theme: 'SKIP', summary: responseText };
    }

    // Last resort: use the raw response as summary
    return {
      theme: ARTICLE_COLLECTION.DEFAULT_THEME,
      summary: responseText.replace(/```json|```/g, '').trim()
    };
  }
}

function synthesizeThemeSummary(apiKey, articles, themeName) {
  if (articles.length === 0) return null;

  // If only one article, return its summary directly (no need to synthesize)
  if (articles.length === 1) {
    return articles[0].summary;
  }

  // Combine all summaries with newsletter subject as context
  const combinedSummaries = articles.map(a =>
    `Newsletter: "${a.subject}"\n${a.summary}`
  ).join('\n\n---\n\n');

  const prompt = buildMasterSummaryPrompt(combinedSummaries);

  console.log(`  Synthesizing ${articles.length} articles for theme "${themeName}"...`);

  const result = callGeminiAPIWithPrompt(apiKey, prompt);

  if (result.error) {
    console.error(`  Synthesis error: ${result.error}`);
    return null;
  }

  console.log(`  Synthesis complete (${result.summary.length} chars)`);
  return result.summary;
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
