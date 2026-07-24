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
    .addItem('Citation Crawl', 'showCrawlbar')
    .addSeparator()
    .addItem('Load Selected Log Row', 'resumeFromLog')
    .addItem('Cancel Crawl', 'cancelCrawl')
    .addSeparator()
    .addItem('Set SerpAPI Key', 'promptForKey')
    .addItem('Set OpenAlex Email', 'promptForOpenAlexEmail')
    .addItem('Set Semantic Scholar API Key', 'promptForS2ApiKey')
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
function showCrawlbar() {
  const html = HtmlService.createHtmlOutputFromFile('crawl_panel')
    .setWidth(450)
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