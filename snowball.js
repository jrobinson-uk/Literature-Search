function runDiagnostic(input) {
  const EMAIL = "your-email@example.com"; // OpenAlex Polite Pool
  const cleanInput = input.trim();
  
  try {
    // Determine if it's a DOI or a Title/ID
    let url = `https://api.openalex.org/works?search=${encodeURIComponent(cleanInput)}&mailto=${EMAIL}`;
    if (cleanInput.includes('10.') || cleanInput.startsWith('W')) {
      // Use the direct ID/DOI resolver
      let id = cleanInput;
      if (cleanInput.includes('10.')) id = `https://doi.org/${cleanInput}`;
      url = `https://api.openalex.org/works/${id}?mailto=${EMAIL}`;
    }

    const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    const result = JSON.parse(response.getContentText());
    
    // OpenAlex returns a single object for direct ID lookups, 
    // but a results array for searches.
    const paper = result.results ? result.results[0] : result;

    if (!paper || !paper.id) {
      return "DIAGNOSTIC FAILED: No paper found for that input.";
    }

    // Capture the metadata we care about
    const report = {
      "Title": paper.title,
      "OpenAlex ID": paper.id,
      "DOI": paper.doi,
      "Citations (Total)": paper.cited_by_count,
      "Citations (API Link)": paper.cited_by_api_url,
      "References (Count)": paper.referenced_works ? paper.referenced_works.length : 0,
      "Year": paper.publication_year,
      "Versions Count": paper.versions_count || "Not listed"
    };

    // Format for display
    let displayMsg = "--- METADATA RECOVERED ---\n\n";
    for (let key in report) {
      displayMsg += `${key}: ${report[key]}\n`;
    }
    
    return displayMsg;

  } catch (e) {
    return `DIAGNOSTIC ERROR: ${e.message}`;
  }
}