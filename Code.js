function processSearch(formData) {
  const ui = SpreadsheetApp.getUi();
  const apiKey = PropertiesService.getScriptProperties().getProperty('SERPAPI_KEY');
  
  if (!apiKey) { ui.alert("No API Key found. Use the menu to set it."); return; }

  // Build the query string
  let queryParts = [];
  if (formData.allWords) queryParts.push(formData.allWords);
  if (formData.exact)    queryParts.push(`"${formData.exact}"`);
  if (formData.atLeast)  queryParts.push(`(${formData.atLeast.split(" ").join(" OR ")})`);
  if (formData.without)  queryParts.push(`-${formData.without.split(" ").join(" -")}`);
  if (formData.author)   queryParts.push(`author:"${formData.author}"`);
  if (formData.source)   queryParts.push(`source:"${formData.source}"`);

  let finalQuery = queryParts.join(" ");
  if (formData.scope === "title") finalQuery = `allintitle: ${finalQuery}`;

  try {
    let url = `https://serpapi.com/search.json?engine=google_scholar&q=${encodeURIComponent(finalQuery)}&api_key=${apiKey}`;
    if (formData.ylo) url += `&as_ylo=${formData.ylo}`;
    if (formData.yhi) url += `&as_yhi=${formData.yhi}`;

    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    
    if (debugMode) console.log("DEBUG DATA:", JSON.stringify(data, null, 2));

    const totalHits = data.search_information?.total_results || 0;
    logSearch(finalQuery, totalHits, formData, 1);

    const proceed = ui.alert("Stats", `Found ~${totalHits.toLocaleString()} results. List titles in 'Results'?`, ui.ButtonSet.YES_NO);
    if (proceed !== ui.Button.YES) return;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let resultsSheet = ss.getSheetByName("Results") || ss.insertSheet("Results");
    
    resultsSheet.clear().clearFormats();

    const headers = [["Title", "Link", "DOI (Extracted)", "PDF Link", "Citations", "Source Summary"]];
    resultsSheet.getRange(1, 1, 1, 6).setValues(headers).setFontWeight("bold").setBackground("#4285f4").setFontColor("white");
    resultsSheet.setFrozenRows(1);

    const results = data.organic_results || [];
    const titleRows = results.map(r => [r.title]);
    resultsSheet.getRange(2, 1, titleRows.length, 1).setValues(titleRows).setWrap(true).setVerticalAlignment("middle");

    SpreadsheetApp.flush();

    const getMeta = ui.alert("Titles Loaded", "Import full metadata?", ui.ButtonSet.YES_NO);
    if (getMeta === ui.Button.YES) {
      const doiRegex = /10\.\d{4,9}\/[-._;()/:A-Z0-9]+/i;
      const metaRows = results.map(r => {
        const foundDoi = r.link ? r.link.match(doiRegex) : null;
        const pdfLink = (r.resources && r.resources[0].link && r.resources[0].link.endsWith('.pdf')) ? r.resources[0].link : "None";
        return [
          r.link || "N/A",
          foundDoi ? foundDoi[0] : "Not found",
          pdfLink,
          r.inline_links?.cited_by?.total || 0,
          r.publication_info?.summary || ""
        ];
      });
      resultsSheet.getRange(2, 2, metaRows.length, 5).setValues(metaRows).setWrap(true).setVerticalAlignment("middle");
      resultsSheet.autoResizeColumns(2, 5);
      resultsSheet.setColumnWidth(1, 350);
    }
  } catch (e) { ui.alert("Error: " + e.message); }
}

/**
 * Appends a new row to the Log sheet in the specific requested order.
 */
/**
 * Appends a new row to the Log sheet with a direct link to the live Scholar search.
 */
/**
 * Appends a new row to the Log sheet with a spacer column (D) for grouping.
 */
function logSearch(query, hits, formData, callCount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("Log") || ss.insertSheet("Log");
  
  if (logSheet.getLastRow() === 0) {
    // Column D is left empty as a spacer for grouping
    const row1 = ["TimeStamp", "Scholar Link", "Search Query", "", "All Words", "Exact Phrase", "At least one of", "Without words", "Where", "Author", "Published in", "Years", "", "API Calls", "Results"];
    const row2 = ["", "", "", "", "", "", "", "", "", "", "", "From", "To", "", ""];
    
    logSheet.getRange(1, 1, 1, 15).setValues([row1]).setFontWeight("bold").setBackground("#eeeeee");
    logSheet.getRange(2, 1, 1, 15).setValues([row2]).setFontWeight("bold").setBackground("#eeeeee");
    
    logSheet.getRange("L1:M1").merge().setHorizontalAlignment("center");
    logSheet.setFrozenRows(2);
  }

  let scholarUrl = `https://scholar.google.com/scholar?q=${encodeURIComponent(query)}`;
  if (formData.ylo) scholarUrl += `&as_ylo=${formData.ylo}`;
  if (formData.yhi) scholarUrl += `&as_yhi=${formData.yhi}`;

  const rowData = [
    new Date(),             // A: TimeStamp
    scholarUrl,             // B: Scholar Link
    query,                  // C: Combined Search Query
    "",                     // D: EMPTY SPACER COLUMN
    formData.allWords || "",// E: All Words
    formData.exact || "",    // F: Exact Phrase
    formData.atLeast || "",  // G: At least one of
    formData.without || "",  // H: Without words
    formData.scope || "",    // I: Where
    formData.author || "",   // J: Author
    formData.source || "",   // K: Published in
    formData.ylo || "",      // L: Years From
    formData.yhi || "",      // M: Years To
    callCount,              // N: API Calls
    hits.toLocaleString()    // O: Results
  ];

  logSheet.appendRow(rowData);
  
  const lastRow = logSheet.getLastRow();
  logSheet.getRange(lastRow, 2).setFormula(`=HYPERLINK("${scholarUrl}", "🌐 View on Scholar")`);
}

/**
 * Updated to account for the spacer in Column D.
 */
function getSelectionFromLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  if (sheet.getName() !== "Log") {
    return { error: "Please click on a row in the 'Log' sheet first." };
  }
  
  const activeRow = sheet.getActiveRange().getRow();
  if (activeRow < 3) return { error: "Please select a data row." };
  
  const data = sheet.getRange(activeRow, 1, 1, 15).getValues()[0];
  
  return {
    allWords: data[4], // Col E (Index 4)
    exact:    data[5], // Col F
    atLeast:  data[6], // Col G
    without:  data[7], // Col H
    scope:    data[8], // Col I
    author:   data[9], // Col J
    source:   data[10],// Col K
    ylo:      data[11],// Col L
    yhi:      data[12] // Col M
  };
}