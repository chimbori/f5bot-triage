/**
 * F5Bot Gmail -> Google Sheet pipeline.
 *
 * Three-stage process, run in order by runAll():
 *   1. importF5BotEmails()      - searches Gmail using GMAIL_QUERY and pulls matching
 *                                 messages into the sheet, skipping any Message ID
 *                                 already present so existing rows (and their
 *                                 extracted/flagged data) are never touched
 *   2. extractSpreadsheetFields() - parses each unprocessed row's HTML Body into structured columns,
 *                                 and colors the keyword's occurrences within the Comment text
 *   3. deleteFlaggedMessages()  - deletes from Gmail any message whose row is marked "d" in Action
 *                                 and removes the row from the spreadsheet (auto-creates the Action
 *                                 column on first run if missing)
 *
 * SETUP
 * 1. Apps Script editor > paste this in.
 * 2. Add the Advanced Gmail service (`Gmail`) in Apps Script project services.
 * 3. Run `runAll`, or run the three functions individually.
 *
 * Note: notify_() is used instead of calling SpreadsheetApp.getUi().alert()
 * directly, since getUi() throws when a function is run from a context with
 * no bound UI (e.g. a time-driven trigger, or run via the script editor
 * without the spreadsheet open). It falls back to Logger.log() in that case.
 */

// --- Constants ---

const SHEET_ID = '10QwlMIfzd2yKHixDfx1iaUF0Sk4toFkgvY8JYrQ3hQ4';
const SHEET_NAME = 'F5Bot Mentions';
const GMAIL_QUERY = 'label:companies/f5bot -label:trash'; // customize this to match your Gmail search criteria

const SUBJECT_COLUMN_NAME = 'Subject';
const DATE_COLUMN_NAME = 'Date';
const BODY_COLUMN_NAME = 'HTML Body';
const MESSAGE_ID_COLUMN_NAME = 'Message ID';
const ACTION_COLUMN_NAME = 'Action';
const BOILERPLATE = 'Do you have comments or suggestions about F5Bot?';
const NEW_COLUMNS = ['Keyword', 'Subreddit', 'Link', 'Comment'];
const KEYWORD_HIGHLIGHT_COLOR = '#8B0000'; // deep red; change to any hex color to adjust the keyword highlight
const BLOCKED_SUBREDDITS = [
  '/r/airbnb_hosts/',
  '/r/ShortTermRentals/',
  '/r/dailygames/',
  '/r/TraktRejects/',
];

/**
 * Entry point: runs the full pipeline end to end.
 * Import new messages -> extract fields -> delete any flagged messages.
 */
function runAll() {
  importF5BotEmails();
  extractSpreadsheetFields();
  deleteFlaggedMessages();
}

/**
 * Stage 1: Finds Gmail messages matching GMAIL_QUERY and appends any that
 * aren't already in the sheet (matched by Message ID). Also removes any
 * rows from the sheet whose Message ID is no longer in Gmail (e.g. the
 * message was deleted or no longer matches the query). Existing rows that
 * are still in Gmail are left completely untouched - their cells (including
 * any extracted fields or Action flags) are never overwritten.
 */
function importF5BotEmails() {
  emptyGmailTrash_();
  const sheet = getSheet_();

  // Ensure the header row exists.
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([SUBJECT_COLUMN_NAME, BODY_COLUMN_NAME, DATE_COLUMN_NAME, MESSAGE_ID_COLUMN_NAME]);
    sheet.setFrozenRows(1);
  }

  const colIndex = getColumnIndex_(sheet);
  const messageIdCol = colIndex[MESSAGE_ID_COLUMN_NAME];
  if (!messageIdCol) {
    throw new Error(`Could not find a column named "${MESSAGE_ID_COLUMN_NAME}".`);
  }

  const existingIds = getExistingMessageIds_(sheet, messageIdCol);

  // Search Gmail using the configured query
  const threads = GmailApp.search(GMAIL_QUERY);

  // Build set of Message IDs currently in Gmail matching the query
  const gmailIds = new Set();
  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      if (!message.isInTrash()) {
        gmailIds.add(message.getId());
      }
    });
  });

  // Remove rows whose Message ID is not in Gmail anymore
  const lastRow = sheet.getLastRow();
  const rowsToDelete = [];
  for (let r = 2; r <= lastRow; r++) {
    const messageId = sheet.getRange(r, messageIdCol).getValue();
    if (messageId && messageId.toString().trim() !== '' && !gmailIds.has(messageId.toString().trim())) {
      rowsToDelete.push(r);
    }
  }

  // Delete in reverse order to keep row numbers valid
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }

  if (rowsToDelete.length > 0) {
    Logger.log(`Removed ${rowsToDelete.length} row(s) with deleted Gmail messages.`);
  }

  // Now add any new messages from Gmail that aren't already in the sheet
  const lastCol = sheet.getLastColumn();
  const subjectIdx = colIndex[SUBJECT_COLUMN_NAME] ? colIndex[SUBJECT_COLUMN_NAME] - 1 : null;
  const bodyIdx = colIndex[BODY_COLUMN_NAME] ? colIndex[BODY_COLUMN_NAME] - 1 : null;
  const dateIdx = colIndex[DATE_COLUMN_NAME] ? colIndex[DATE_COLUMN_NAME] - 1 : null;
  const messageIdIdx = messageIdCol - 1;

  const newRows = [];
  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const id = message.getId();
      if (existingIds.has(id)) return; // already in the sheet; leave that row alone

      const row = new Array(lastCol).fill('');
      if (subjectIdx !== null) row[subjectIdx] = message.getSubject();
      if (bodyIdx !== null) row[bodyIdx] = message.getBody();
      if (dateIdx !== null) row[dateIdx] = message.getDate();
      row[messageIdIdx] = id;
      newRows.push(row);
    });
  });

  if (newRows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    const range = sheet.getRange(startRow, 1, newRows.length, lastCol);
    range.setValues(newRows);
    range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // preserve line breaks visually
  }

  Logger.log(`Added ${newRows.length} new row(s) from Gmail`);
}

/**
 * Stage 2: For every unprocessed row (Keyword cell still empty), parses
 * the HTML Body to pull out Keyword, Subreddit, Link, and Comment into
 * their own columns, then trims the HTML Body down to just the
 * message content (stripping the F5Bot boilerplate/footer). Every
 * occurrence of the keyword within the Comment text is colored using
 * KEYWORD_HIGHLIGHT_COLOR.
 */
function extractSpreadsheetFields() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // no data rows

  const colIndex = getColumnIndex_(sheet);

  const bodyCol = colIndex[BODY_COLUMN_NAME];
  if (!bodyCol) {
    throw new Error(`Could not find a column named "${BODY_COLUMN_NAME}". Update BODY_COLUMN_NAME.`);
  }

  // Add any missing output columns
  NEW_COLUMNS.forEach(name => {
    if (!colIndex[name]) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(name);
      colIndex[name] = newCol;
    }
  });
  if (!colIndex[ACTION_COLUMN_NAME]) {
    const actionCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, actionCol).setValue(ACTION_COLUMN_NAME);
    colIndex[ACTION_COLUMN_NAME] = actionCol;
  }

  const finalLastCol = sheet.getLastColumn();
  const dataRange = sheet.getRange(2, 1, lastRow - 1, finalLastCol);
  const data = dataRange.getValues();

  const bodyIdx = colIndex[BODY_COLUMN_NAME] - 1;
  const keywordIdx = colIndex['Keyword'] - 1;
  const subredditIdx = colIndex['Subreddit'] - 1;
  const linkIdx = colIndex['Link'] - 1;
  const commentIdx = colIndex['Comment'] - 1;
  const actionIdx = colIndex[ACTION_COLUMN_NAME] - 1;
  const messageIdIdx = colIndex[MESSAGE_ID_COLUMN_NAME] - 1;
  const commentCol = colIndex['Comment']; // 1-based, needed for rich text writes below

  let processedCount = 0;
  let markedBlockedCount = 0;
  const rowsToHighlight = []; // { rowNumber, comment, keyword }

  for (let r = 0; r < data.length; r++) {
    const row = data[r];

    const existingSubreddit = row[subredditIdx] ? row[subredditIdx].toString().trim() : '';
    if (existingSubreddit && isBlockedSubreddit_(existingSubreddit)) {
      row[actionIdx] = 'd';
      markedBlockedCount++;
      Logger.log(`Blocked subreddit match: row ${r + 2}, Message ID ${row[messageIdIdx]}, Subreddit ${existingSubreddit}; marked Action=d.`);
    }

    // Skip rows already processed
    if (row[keywordIdx] && row[keywordIdx].toString().trim() !== '') continue;

    const rawBody = row[bodyIdx];
    if (!rawBody || rawBody.toString().trim() === '') continue;

    const body = rawBody.toString();

    // Extract fields from the original HTML before it gets trimmed below.
    const keyword = extractKeyword_(body);
    const comment = extractComment_(body);
    const subreddit = extractSubreddit_(body);

    row[keywordIdx] = keyword;
    row[subredditIdx] = subreddit;
    row[linkIdx] = extractLink_(body);
    row[commentIdx] = comment;
    row[bodyIdx] = stripBoilerplate_(body);

    if (isBlockedSubreddit_(subreddit)) {
      row[actionIdx] = 'd';
      markedBlockedCount++;
      Logger.log(`Blocked subreddit match: row ${r + 2}, Message ID ${row[messageIdIdx]}, Subreddit ${subreddit}; marked Action=d.`);
    }

    if (keyword && comment) {
      rowsToHighlight.push({ rowNumber: r + 2, comment, keyword }); // +2: data starts at row 2, r is 0-based
    }

    processedCount++;
  }

  dataRange.setValues(data);
  rowsToHighlight.forEach(({ rowNumber, comment, keyword }) => {
    highlightKeywordInComment_(sheet, rowNumber, commentCol, comment, keyword);
  });

  notify_(`Extracted Keyword, Subreddit, Link, and Comment for ${processedCount} row(s). Marked ${markedBlockedCount} blocked row(s) for deletion.`);
}

/**
 * Stage 3: For every row with "d" in the Action column, deletes the
 * corresponding Gmail message (moves it to Trash) using its Message ID,
 * then removes the entire row from the spreadsheet (compacting rows so
 * no blanks remain). Auto-creates the Action column on first run if it
 * doesn't exist yet.
 */
function deleteFlaggedMessages() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const colIndex = getColumnIndex_(sheet);

  // Create the Action column if it doesn't exist yet - nothing to delete either way.
  if (!colIndex[ACTION_COLUMN_NAME]) {
    const newCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, newCol).setValue(ACTION_COLUMN_NAME);
    colIndex[ACTION_COLUMN_NAME] = newCol;
    Logger.log(`Created "${ACTION_COLUMN_NAME}" column. Nothing to delete yet.`);
    return;
  }

  const messageIdCol = colIndex[MESSAGE_ID_COLUMN_NAME];
  if (!messageIdCol) throw new Error(`Could not find a column named "${MESSAGE_ID_COLUMN_NAME}".`);

  const lastCol = sheet.getLastColumn();
  const actionIdx = colIndex[ACTION_COLUMN_NAME] - 1;
  const messageIdIdx = messageIdCol - 1;

  let deletedCount = 0;
  let errorCount = 0;
  const rowsToDelete = []; // track row numbers (1-based) to delete, in reverse order

  // First pass: identify which rows to delete
  for (let r = 2; r <= lastRow; r++) {
    const row = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
    const flag = row[actionIdx] ? row[actionIdx].toString().trim().toLowerCase() : '';
    if (flag !== 'd') continue;

    const messageId = row[messageIdIdx] ? row[messageIdIdx].toString().trim() : '';
    if (!messageId) {
      errorCount++;
      rowsToDelete.push(r); // queue for deletion even without valid message ID
      continue;
    }

    try {
      const message = GmailApp.getMessageById(messageId);
      message.moveToTrash();
      deletedCount++;
    } catch (e) {
      Logger.log(`Error deleting message ${messageId}: ${e.message}`);
      errorCount++;
    }

    rowsToDelete.push(r); // queue for deletion
  }

  // Second pass: delete rows in reverse order so row numbers stay valid
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }

  notify_(`Deleted ${deletedCount} message(s). ${errorCount} error(s).`);
}

// --- Private helpers ---

/** Permanently deletes every thread currently in Gmail Trash. */
function emptyGmailTrash_() {
  if (typeof Gmail === 'undefined') {
    throw new Error('Enable the Advanced Gmail service named "Gmail" in the Apps Script project services.');
  }

  const pageSize = 100;
  let deletedCount = 0;
  let trashThreads;

  do {
    trashThreads = GmailApp.getTrashThreads(0, pageSize);
    trashThreads.forEach(thread => {
      Gmail.Users.Threads.remove('me', thread.getId());
      deletedCount++;
    });
  } while (trashThreads.length === pageSize);

  if (deletedCount > 0) {
    Logger.log(`Permanently deleted ${deletedCount} thread(s) from Gmail Trash.`);
  }
}

/** Gets the F5Bot Mentions sheet, creating it if it doesn't exist. */
function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

/** Builds a map of header name -> 1-based column index from the sheet's first row. */
function getColumnIndex_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const colIndex = {};
  headers.forEach((h, i) => { if (h) colIndex[h.toString().trim()] = i + 1; });
  return colIndex;
}

/** Reads the given Message ID column and returns the set of IDs already present in the sheet. */
function getExistingMessageIds_(sheet, messageIdCol) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();
  const values = sheet.getRange(2, messageIdCol, lastRow - 1, 1).getValues();
  const ids = values.map(row => row[0] ? row[0].toString().trim() : '').filter(id => id !== '');
  return new Set(ids);
}

/**
 * Shows a UI alert if a bound spreadsheet UI is available (e.g. running from
 * the editor with the sheet open); otherwise logs the message instead, so
 * this works from triggers and other non-UI execution contexts.
 */
function notify_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log(message);
  }
}

/** Truncates the HTML body at the F5Bot boilerplate line, discarding the footer/ads that follow. */
function stripBoilerplate_(html) {
  const idx = html.indexOf(BOILERPLATE);
  if (idx === -1) return html;
  const closeTagIdx = html.indexOf('</p>', idx);
  return closeTagIdx === -1
    ? html.substring(0, idx).trim()
    : html.substring(0, closeTagIdx + '</p>'.length).trim();
}

/** Extracts the search keyword from the <h2>Keyword: "..."</h2> line. */
function extractKeyword_(html) {
  const heading = html.match(/<h2\b[^>]*>\s*Keyword:\s*(?:&quot;|&#34;|&#x22;|")\s*([\s\S]*?)\s*(?:&quot;|&#34;|&#x22;|")\s*<\/h2>/i);
  if (!heading) return '';

  return heading[1]
    .replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&')
    .replace(/&quot;|&#34;|&#x22;/gi, '"')
    .trim();
}

/** Extracts the subreddit (e.g. "/r/CompetitiveApex/") from the post description. */
function extractSubreddit_(html) {
  const m = html.match(/(\/r\/[^\/)]+\/)/i);
  return m ? m[1] : '';
}

/** Returns true when an extracted subreddit is on the blocklist. */
function isBlockedSubreddit_(subreddit) {
  const value = subreddit.toString().trim().toLowerCase();
  return BLOCKED_SUBREDDITS.some(blocked => value === blocked.toLowerCase());
}

/** Extracts and URL-decodes the target link from the first f5bot.com/url?u= redirect. */
function extractLink_(html) {
  const m = html.match(/href=['"]https:\/\/f5bot\.com\/url\?u=([^&'"]+)/i);
  if (!m) return '';
  try {
    return decodeURIComponent(m[1]);
  } catch (e) {
    return m[1];
  }
}

/** Extracts the matched comment/post text from the monospace <span>, stripping nested tags. */
function extractComment_(html) {
  const m = html.match(/<span[^>]*font-family:[^>]*monospace[^>]*>([\s\S]*?)<\/span>/i);
  if (!m) return '';
  // Strip any nested tags (e.g. <strong>) and collapse whitespace
  return m[1]
    .replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&')
    .replace(/&nbsp;/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Applies a colored foreground to the keyword within a single Comment cell,
 * using KEYWORD_HIGHLIGHT_COLOR. All case-insensitive occurrences of the
 * keyword in the comment text are styled.
 */
function highlightKeywordInComment_(sheet, rowNumber, commentCol, comment, keyword) {
  const cell = sheet.getRange(rowNumber, commentCol);

  const builder = SpreadsheetApp.newRichTextValue().setText(comment);
  const lowerComment = comment.toLowerCase();
  const lowerKeyword = keyword.toLowerCase();

  let searchFrom = 0;
  let idx;
  while ((idx = lowerComment.indexOf(lowerKeyword, searchFrom)) !== -1) {
    const end = idx + keyword.length;
    builder.setTextStyle(idx, end, SpreadsheetApp.newTextStyle()
      .setForegroundColor(KEYWORD_HIGHLIGHT_COLOR)
      .build());
    searchFrom = end;
  }

  cell.setRichTextValue(builder.build());
}
