const debugMode = true;
const VERSION = "1.11.6";

/**
 * Runs automatically when the sheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Scholar Tool')
    .addItem('Open Search Panel', 'showSidebar')
    .addSeparator()
    .addItem('Snowball Search', 'showSnowballbar')
    .addItem('Citation Crawl', 'showCrawlV2bar')
    .addSeparator()
    .addItem('Load Selected Log Row', 'resumeFromLog')
    .addItem('Cancel Crawl', 'cancelCrawlV2')
    .addItem('Show Crawl Progress', 'debugCrawlV2Progress')
    .addSeparator()
    .addItem('Set SerpAPI Key', 'promptForKey')
    .addItem('Set OpenAlex Email', 'promptForOpenAlexEmail')
    .addItem('Set Semantic Scholar API Key', 'promptForS2ApiKey')
    .addSeparator()
    .addItem('Debug: Property Storage Usage', 'debugPropertyStorage')
    .addItem('Debug: Clean Up Legacy v1 Crawl Properties', 'cleanupLegacyCrawlProperties')
    .addSeparator()
    .addItem('Version ' + VERSION, 'showVersion')
    .addToUi();
    
  // Note: Auto-opening sidebars only works after the user has 
  // authorized the script in this specific document.
  showSidebar();
}

/**
 * Opens the sidebar. Ensure the HTML file is named "Sidebar" (case sensitive).
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Integrated Scholar Search')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}
function showSnowballbar() {
  const html = HtmlService.createHtmlOutputFromFile('snowball_panel')
    .setTitle('OpenAlex Snowballer')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}
// A modeless dialog rather than a sidebar — sidebars are hard-capped at
// 300px by Sheets itself (setWidth() is silently ignored for showSidebar()),
// so a genuinely wider panel means a floating dialog instead. Modeless, not
// modal, so it doesn't block interacting with the sheet while a crawl runs.
//
// As of v22 this is the only citation crawl pipeline (the original v1
// forward/backward-only crawler has been archived — see crawl.js's header
// comment and archive/). Internal names (crawl_v2_panel.html, CRAWL2_*
// script properties, the *V2 function suffixes throughout crawl_v2.js) are
// kept as-is deliberately, not renamed to drop the "2"/"V2" — there may be
// stopped-but-resumable old v1 crawls whose script properties use the
// original unprefixed names, and reusing those exact names here risked
// colliding with and corrupting that state. Only the user-facing labels
// (this dialog's title, the menu entry) have been promoted to plain
// "Citation Crawl".
function showCrawlV2bar() {
  const html = HtmlService.createHtmlOutputFromFile('crawl_v2_panel')
    .setWidth(900)
    .setHeight(700);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Citation Crawl');
}

function promptForKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Set SerpAPI Key", "Paste your SerpApi Key here:", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty('SERPAPI_KEY', response.getResponseText());
    ui.alert("API Key saved successfully!");
  }
}

function showVersion() {
  SpreadsheetApp.getUi().alert('Scholar Tool — Version ' + VERSION);
}

function promptForOpenAlexEmail() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Set OpenAlex Email", "Enter your email address for the OpenAlex polite pool (free access):", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty('OPENALEX_EMAIL', response.getResponseText().trim());
    ui.alert("OpenAlex email saved successfully!");
  }
}

function promptForS2ApiKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Set Semantic Scholar API Key",
    "Paste your Semantic Scholar API key here.\n\n" +
    "Apply at: https://www.semanticscholar.org/product/api#api-key\n\n" +
    "Leave blank and click OK to clear a previously saved key.",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() == ui.Button.OK) {
    const key = response.getResponseText().trim();
    if (key) {
      PropertiesService.getScriptProperties().setProperty('S2_API_KEY', key);
      ui.alert("Semantic Scholar API key saved successfully!");
    } else {
      PropertiesService.getScriptProperties().deleteProperty('S2_API_KEY');
      ui.alert("Semantic Scholar API key cleared.");
    }
  }
}

// Diagnostic for PropertiesService's total-store quota (500KB, script-wide
// across every feature — crawl v2, snowball, saved keys, everything) being
// exceeded. Reports every currently-stored key with its approximate byte
// size, largest first, so a "property storage quota" error can be traced
// to an actual cause instead of guessed at.
function debugPropertyStorage() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var keys = Object.keys(props);
  var sized = keys.map(function(k) {
    var v = props[k] || '';
    return { key: k, bytes: k.length + v.length };
  }).sort(function(a, b) { return b.bytes - a.bytes; });
  var total = sized.reduce(function(sum, s) { return sum + s.bytes; }, 0);
  var lines = sized.slice(0, 25).map(function(s) {
    return s.key + ': ' + s.bytes + ' bytes';
  });
  var msg = 'Total properties: ' + keys.length +
    '\nTotal approx size: ' + total + ' / ~500,000 bytes\n\n' +
    'Largest 25:\n' + lines.join('\n');
  SpreadsheetApp.getUi().alert('Property Storage Usage', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// Deletes script properties left over from the retired v1 crawler (archived
// in archive/crawl_v1_full.js — the live pipeline is v2/CRAWL2_* only, see
// showCrawlV2bar's comment). No current code path reads these; their only
// effect is counting against the 500KB total-storage quota. Confirmed with
// the user before deleting anything, since this isn't reversible.
function cleanupLegacyCrawlProperties() {
  var ui = SpreadsheetApp.getUi();
  var legacyKeys = [
    'CRAWL_ACTIVE_SHEET', 'CRAWL_BACKWARD_DEPTH', 'CRAWL_BACKWARD_DONE',
    'CRAWL_BACKWARD_IDX', 'CRAWL_BATCH_NUM', 'CRAWL_CONSEC_FAILURES',
    'CRAWL_DIRECTION', 'CRAWL_EXPANDING_BACKWARD', 'CRAWL_EXPAND_BACKWARD',
    'CRAWL_LOG_ROW', 'CRAWL_MATCHES_ONLY', 'CRAWL_MAX_DEPTH',
    'CRAWL_MAX_PAPERS', 'CRAWL_PHASE', 'CRAWL_RUN_BACKWARD',
    'CRAWL_TRIGGER_ID', 'CRAWL_YEAR_BOUND', 'CRAWL_YEAR_CEILING',
    'CRAWL_YEAR_FLOOR'
  ];
  var props = PropertiesService.getScriptProperties();
  var present = legacyKeys.filter(function(k) { return props.getProperty(k) !== null; });
  if (present.length === 0) {
    ui.alert('No legacy v1 crawl properties found — nothing to clean up.');
    return;
  }
  var response = ui.alert(
    'Delete legacy v1 crawl properties?',
    'Found ' + present.length + ' leftover key(s) from the retired v1 crawler:\n\n' +
    present.join('\n') + '\n\nThese are not read by any current code path. Delete them?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;
  present.forEach(function(k) { props.deleteProperty(k); });
  ui.alert('Deleted ' + present.length + ' legacy propert' + (present.length === 1 ? 'y' : 'ies') + '.');
}