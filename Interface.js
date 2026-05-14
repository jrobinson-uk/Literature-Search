const debugMode = true;

/**
 * Runs automatically when the sheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Scholar Tool')
    .addItem('Open Search Panel', 'showSidebar')
    .addSeparator()
      .addItem('Start Snowball', 'showSnowballbar')
    .addSeparator()
    .addItem('Set API Key', 'promptForKey')
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
function promptForKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Set API Key", "Paste your SerpApi Key here:", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty('SERPAPI_KEY', response.getResponseText());
    ui.alert("API Key saved successfully!");
  }
}