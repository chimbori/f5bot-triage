# F5Bot Spreadsheet Bridge

Google Apps Script pipeline that imports F5Bot Gmail alerts into a Google Sheet, extracts useful fields, highlights keyword matches, and removes unwanted subreddit messages.

## Setup

1. Open the target Google Sheet and open **Extensions > Apps Script**.
2. Paste the contents of `f5bot-spreadsheet-bridge.js` into the Apps Script project.
3. Update `SHEET_ID`, `SHEET_NAME`, and `GMAIL_QUERY` if needed.
4. Add the Advanced Gmail service:
   - In Apps Script, open **Services**.
   - Add the service named `Gmail`.
   - Use the default identifier `Gmail`.
5. Grant the requested Gmail and Sheets permissions when prompted.
6. Run `runAll()` manually once to authorize the script and verify the setup.

The Advanced Gmail service is required because the script permanently deletes messages that are already in Gmail Trash before importing new messages.

## Pipeline

`runAll()` runs these stages in order:

1. **Import**
   - Permanently deletes all threads currently in Gmail Trash.
   - Searches Gmail using `GMAIL_QUERY`.
   - Removes spreadsheet rows whose Message IDs are no longer in Gmail.
   - Imports new messages without overwriting existing rows.

2. **Extract**
   - Creates missing output columns: `Keyword`, `Subreddit`, `Link`, `Comment`, and `Action`.
   - Extracts fields from each message's HTML body.
   - Trims the stored HTML body at the F5Bot reply prompt.
   - Highlights keyword occurrences in the `Comment` cell.
   - Marks blocked subreddit rows with `Action = d`.

3. **Delete**
   - Finds rows marked `d` in the `Action` column.
   - Moves the corresponding Gmail messages to Trash.
   - Removes those rows from the spreadsheet.

## Blocked Subreddits

The initial blocklist is defined near the top of the script:

```javascript
const BLOCKED_SUBREDDITS = [
  '/r/airbnb_hosts/',
  '/r/ShortTermRentals/',
  '/r/dailygames/',
  '/r/TraktRejects/',
];
```

Add subreddit paths in the same format, including the leading and trailing slash. Matching is case-insensitive and uses the extracted `Subreddit` column value.

## Local Tests

The repository includes Node.js tests that use the supplied sample email files:

```sh
node --test f5bot-spreadsheet-bridge.test.js
```

The tests do not connect to Gmail or Google Sheets; Apps Script services are mocked locally.

## Files

- `f5bot-spreadsheet-bridge.js` - Apps Script implementation.
- `f5bot-spreadsheet-bridge.test.js` - Node.js regression tests.
- `sample1.txt`, `sample2.txt`, `sample3.txt` - sample F5Bot HTML emails used by the tests.
