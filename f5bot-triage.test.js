const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const projectDir = __dirname;
const sampleFiles = ['sample1.txt', 'sample2.txt', 'sample3.txt'];

function loadSample(fileName) {
  let html = fs.readFileSync(path.join(projectDir, fileName), 'utf8').trim();
  if (html.startsWith('"') && html.endsWith('"')) {
    html = html.slice(1, -1).replace(/""/g, '"');
  }
  return html;
}

function createScriptContext() {
  const alerts = [];
  const logs = [];
  const context = {
    Logger: { log: message => logs.push(message) },
    SpreadsheetApp: {
      getUi: () => ({ alert: message => alerts.push(message) }),
      newRichTextValue: () => {
        const richText = {
          text: '',
          styles: [],
          setText(text) {
            this.text = text;
            return this;
          },
          setTextStyle(start, end, style) {
            this.styles.push({ start, end, style });
            return this;
          },
          build() {
            return { text: this.text, styles: this.styles };
          },
        };
        return richText;
      },
      newTextStyle: () => {
        let foregroundColor;
        return {
          setForegroundColor(color) {
            foregroundColor = color;
            return this;
          },
          build() {
            return { foregroundColor };
          },
        };
      },
    },
  };

  vm.createContext(context);
  vm.runInContext(fs.readFileSync(path.join(projectDir, 'f5bot-triage.js'), 'utf8'), context);
  return { context, alerts, logs };
}

const { context, alerts, logs } = createScriptContext();

test('extracts fields from every sample email', () => {
  const expected = [
    { subreddit: '/r/peloton/', link: 'https://www.reddit.com/r/peloton/comments/1wfsugg/c/pa09cnp?context=3' },
    { subreddit: '/r/ProtonMail/', link: 'https://www.reddit.com/r/ProtonMail/comments/1whhv0g/c/pa3qqgu?context=3' },
    { subreddit: '/r/airbnb_hosts/', link: 'https://www.reddit.com/r/airbnb_hosts/comments/1wiwaiu/c/pafnndj?context=3' },
  ];

  sampleFiles.forEach((fileName, index) => {
    const html = loadSample(fileName);
    const commentSnippet = html.match(/<span[^>]*font-family:[^>]*monospace[^>]*>[\s\S]*?<\/span>/i)[0];
    const headingSnippet = html.match(/<h2\b[^>]*>[\s\S]*?<\/h2>/i)[0];
    const subredditSnippet = html.match(/<span[^>]*f5-meta[^>]*>Reddit Comments \([\s\S]*?<\/span>/i)[0];
    const linkSnippet = html.match(/<a href="https:\/\/f5bot\.com\/url\?u=[^>]+>/i)[0];

    assert.equal(context.extractKeyword_(headingSnippet), 'ical');
    assert.equal(context.extractSubreddit_(subredditSnippet), expected[index].subreddit);
    assert.equal(context.extractLink_(linkSnippet), expected[index].link);
    assert.match(context.extractComment_(commentSnippet), /ical/i);
    assert.equal(context.stripBoilerplate_(html).includes('Do you have comments or suggestions'), true);
  });
});

test('identifies blocked subreddits in sample emails', () => {
  assert.equal(context.isBlockedSubreddit_('/r/peloton/'), false);
  assert.equal(context.isBlockedSubreddit_('/r/ProtonMail/'), false);
  assert.equal(context.isBlockedSubreddit_(context.extractSubreddit_(loadSample('sample3.txt'))), true);
  assert.equal(context.isBlockedSubreddit_('/r/AIRBNB_HOSTS/'), true);
});

test('imports blocked messages before extraction marks them for deletion', () => {
  const testSetup = createScriptContext();
  const testContext = testSetup.context;
  const values = [['Subject', 'HTML Body', 'Date', 'Message ID']];
  const sheet = {
    getLastRow: () => values.length,
    getLastColumn: () => values[0].length,
    appendRow: row => values.push([...row]),
    setFrozenRows: () => { },
    getRange: (row, column, rowCount = 1, columnCount = 1) => ({
      getValues: () => values.slice(row - 1, row - 1 + rowCount).map(value => Array.from(
        { length: columnCount },
        (_, index) => value[column - 1 + index] ?? '',
      )),
      setValue: value => { values[row - 1][column - 1] = value; },
      setValues: rows => rows.forEach((value, index) => { values[row - 1 + index] = [...value]; }),
      setWrapStrategy: () => { },
      setRichTextValue: () => { },
      getValue: () => values[row - 1][column - 1],
    }),
  };
  const messages = [
    {
      getId: () => 'blocked-id',
      getBody: () => loadSample('sample3.txt'),
      getSubject: () => 'Blocked',
      getDate: () => '2026-09-21',
      isInTrash: () => false,
    },
    {
      getId: () => 'allowed-id',
      getBody: () => loadSample('sample1.txt'),
      getSubject: () => 'Allowed',
      getDate: () => '2026-09-21',
      isInTrash: () => false,
    },
  ];

  testContext.SpreadsheetApp.openById = () => ({ getSheetByName: () => sheet });
  testContext.SpreadsheetApp.WrapStrategy = { WRAP: 'WRAP' };
  const callOrder = [];
  testContext.Gmail = { Users: { Threads: { remove: () => { } } } };
  testContext.GmailApp = {
    getTrashThreads: () => {
      callOrder.push('trash');
      return [];
    },
    search: () => {
      callOrder.push('search');
      return [{ getMessages: () => messages }];
    },
  };
  testContext.importF5BotEmails();

  assert.deepEqual(callOrder, ['trash', 'search']);
  assert.deepEqual(values[0], ['Subject', 'HTML Body', 'Date', 'Message ID']);
  assert.equal(values[1][3], 'blocked-id');
  assert.equal(values[2][3], 'allowed-id');

  testContext.extractSpreadsheetFields();

  assert.equal(values[0][5], 'Subreddit');
  assert.equal(values[0][8], 'Action');
  assert.equal(values[1][5], '/r/airbnb_hosts/');
  assert.equal(values[1][8], 'd');
  assert.equal(values[2][5], '/r/peloton/');
  assert.equal(values[2][8], '');
  assert.match(testSetup.logs.find(log => log.includes('Blocked subreddit match:')), /Subreddit \/r\/airbnb_hosts\/; marked Action=d\./);
});

test('returns empty values when parser snippets are absent', () => {
  assert.equal(context.extractKeyword_('<h2>Subject only</h2>'), '');
  assert.equal(context.extractSubreddit_('<p>No subreddit</p>'), '');
  assert.equal(context.extractLink_('<p>No link</p>'), '');
  assert.equal(context.extractComment_('<p>No comment</p>'), '');
  const html = '<p>Message only</p>';
  assert.equal(context.stripBoilerplate_(html), html);
});

test('getColumnIndex_ maps trimmed headers to one-based columns', () => {
  const sheet = { getLastColumn: () => 3, getRange: () => ({ getValues: () => [[' Date ', '', 'Comment']] }) };
  assert.deepEqual({ ...context.getColumnIndex_(sheet) }, { Date: 1, Comment: 3 });
});

test('getExistingMessageIds_ returns trimmed non-empty IDs', () => {
  const sheet = { getLastRow: () => 4, getRange: () => ({ getValues: () => [[' id-1 '], [''], ['id-2']] }) };
  assert.deepEqual([...context.getExistingMessageIds_(sheet, 2)], ['id-1', 'id-2']);
});

test('getSheet_ returns an existing sheet or creates one', () => {
  const existing = { name: 'existing' };
  const existingSpreadsheet = { getSheetByName: () => existing };
  context.SpreadsheetApp.openById = () => existingSpreadsheet;
  assert.equal(context.getSheet_(), existing);

  const created = { name: 'created' };
  const newSpreadsheet = { getSheetByName: () => null, insertSheet: () => created };
  context.SpreadsheetApp.openById = () => newSpreadsheet;
  assert.equal(context.getSheet_(), created);
});

test('notify_ uses the UI and falls back to Logger', () => {
  context.SpreadsheetApp.getUi = () => ({ alert: message => alerts.push(message) });
  context.notify_('processed');
  assert.equal(alerts.at(-1), 'processed');

  context.SpreadsheetApp.getUi = () => { throw new Error('no UI'); };
  context.notify_('logged');
  assert.equal(logs.at(-1), 'logged');
});

test('highlightKeywordInComment_ styles every case-insensitive occurrence', () => {
  let richText;
  const sheet = {
    getRange: () => ({ setRichTextValue: value => { richText = value; } }),
  };

  context.highlightKeywordInComment_(sheet, 2, 4, 'ical iCal', 'ical');

  assert.equal(richText.text, 'ical iCal');
  assert.deepEqual(richText.styles.map(style => [style.start, style.end]), [[0, 4], [5, 9]]);
});
