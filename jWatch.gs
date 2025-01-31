function filterAndCategorizeEmails() {
  const CONFIG = {
    SEARCH_QUERY: '("Lamaranmu untuk posisi" OR "Terima kasih atas minat Anda" OR "sekarang telah kedaluwarsa")',
    SHEET_NAME: "watchlist",
    HEADERS: ["Date", "Application Name", "Company Name", "Status"]
  };

  const threads = fetchAllThreads(CONFIG.SEARCH_QUERY);
  const categorizedEmails = processThreads(threads);
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAME, CONFIG.HEADERS);
  writeToSheet(sheet, CONFIG.HEADERS, categorizedEmails);
  Logger.log("Processed %s emails. Data updated successfully!", categorizedEmails.length);
}

function fetchAllThreads(query) {
  let threads = [];
  let start = 0;
  const batchSize = 500;

  while (true) {
    const batch = GmailApp.search(query, start, batchSize);
    if (batch.length === 0) break;
    threads = threads.concat(batch);
    start += batchSize;
  }

  return threads;
}

function processThreads(threads) {
  return threads.flatMap(thread => 
    thread.getMessages().map(message => processMessage(message)).filter(Boolean)
  );
}

function processMessage(message) {
  const body = message.getBody().replace(/<[^>]+>/g,'').replace(/\s+/g,' ').trim();
  const date = message.getDate();

  if (body.includes("Lamaranmu untuk posisi") && body.includes("berhasil dikirimkan ke")) {
    const match = parseMatch(body, 
      /Lamaranmu untuk posisi\s+(.*?)\s+berhasil/,
      /berhasil dikirimkan ke\s+(.*?)\s*&#8202;/
    );
    return match && [date, match[0], match[1].replace(/&amp;/g,'&'), "Sent"];
  }

  if (body.includes("Terima kasih atas minat Anda pada") && body.includes("Sayangnya")) {
    const match = parseMatch(body,
      /pada\s*(.*?)\s*lowongan/,
      /di\s+([^.\n<]+?)\s*\.\s*Sayangnya/
    );
    return match && [date, match[0], match[1].replace(/&amp;/g,'&'), "Rejected"];
  }

  if (body.includes("sekarang telah kedaluwarsa")) {
    const match = parseMatch(body,
      /Pekerjaan\s+(.*?)\s+yang Anda lamar di/,
      /di\s+([^.\n<]+?)\s+sekarang telah kedaluwarsa/
    );
    if (match) {
      const companyName = match[1].replace(/telah kedaluwarsa.*/, '').trim();
      return [date, match[0], companyName, "Rejected"];
    }
  }

  return null;
}

function parseMatch(body, ...regexes) {
  const matches = regexes.map(r => {
    const match = body.match(r);
    return match ? match[1].trim() : null;
  });
  return matches.every(Boolean) ? matches : null;
}

function getOrCreateSheet(sheetName, headers) {
  const spreadsheet = SpreadsheetApp.getActive() || SpreadsheetApp.create("Jobstreet Watch");
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.getActiveSheet().setName(sheetName);
    sheet.appendRow(headers);
  } 
  else if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
  return sheet;
}

function writeToSheet(sheet, headers, data) {
  if (!data.length) return;
  const processedData = data.map(row => row.map(field => 
    typeof field === 'string' ? HtmlService.createHtmlOutput(field).getContent() : field
  ));
  sheet.getRange(sheet.getLastRow()+1, 1, data.length, headers.length).setValues(processedData);
}
