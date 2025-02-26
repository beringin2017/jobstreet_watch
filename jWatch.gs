function filterAndCategorizeEmails() {
  const CONFIG = {
    SEARCH_QUERY: '("Lamaranmu untuk posisi" OR "Terima kasih atas minat Anda" OR "sekarang telah kedaluwarsa" OR "Maaf, lamaranmu untuk" OR subject:"Maaf, lamaranmu untuk")',
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
  const messagesArray = GmailApp.getMessagesForThreads(threads);
  let result = [];
  for (let i = 0; i < messagesArray.length; i++) {
    for (let j = 0; j < messagesArray[i].length; j++) {
      const processed = processMessage(messagesArray[i][j]);
      if (processed) result.push(processed);
    }
  }
  return result;
}

const regexSentPosition = /Lamaranmu untuk posisi\s+(.*?)\s+berhasil/;
const regexSentCompany = /berhasil dikirimkan ke\s+(.*?)\s*&#8202;/;
const regexRejectedPosition = /pada\s*(.*?)\s*lowongan/;
const regexRejectedCompany = /di\s+([^.\n<]+?)\s*\.\s*Sayangnya/;
const regexSubjectPosition = /lamaranmu untuk\s+(.*?)\s+di/;
const regexSubjectCompany = /di\s+(.*?)\s+belum sesuai/;

function processMessage(message) {
  const rawBody = message.getBody();
  const body = rawBody.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
  const date = message.getDate();
  const subject = message.getSubject();
  if (body.includes("Lamaranmu untuk posisi") && body.includes("berhasil dikirimkan ke")) {
    const posMatch = body.match(regexSentPosition);
    const compMatch = body.match(regexSentCompany);
    if (posMatch && compMatch) {
      return [date, posMatch[1].trim(), compMatch[1].trim().replace(/&amp;/g, '&'), "Sent"];
    }
  }
  if (body.includes("Terima kasih atas minat Anda pada") && body.includes("Sayangnya")) {
    const posMatch = body.match(regexRejectedPosition);
    const compMatch = body.match(regexRejectedCompany);
    if (posMatch && compMatch) {
      return [date, posMatch[1].trim(), compMatch[1].trim().replace(/&amp;/g, '&'), "Rejected"];
    }
  }
  if (body.includes("sekarang telah kedaluwarsa")) {
    const expiredData = parseExpiredJob(body);
    if (expiredData) {
      const [positionName, companyName] = expiredData;
      return [date, positionName, companyName, "Rejected"];
    }
  }
  if (subject.includes("Maaf, lamaranmu untuk") && subject.includes("belum sesuai")) {
    const posMatch = subject.match(regexSubjectPosition);
    const compMatch = subject.match(regexSubjectCompany);
    if (posMatch && compMatch) {
      return [date, posMatch[1].trim(), compMatch[1].trim(), "Rejected"];
    }
  }
  return null;
}

function parseExpiredJob(body) {
  const startToken = "Pekerjaan ";
  const endToken = " sekarang telah kedaluwarsa";
  const startIndex = body.lastIndexOf(startToken);
  if (startIndex === -1) return null;
  const endIndex = body.indexOf(endToken, startIndex);
  if (endIndex === -1) return null;
  let snippet = body.substring(startIndex + startToken.length, endIndex).trim();
  snippet = snippet
    .replace(/yang Anda lamar/gi, '')
    .replace(/telah kedaluwarsa/gi, '')
    .replace(/Hai\s+[^,]+,/gi, '')
    .replace(/Pekerjaan/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
  const diIndex = snippet.lastIndexOf(" di ");
  if (diIndex === -1) return null;
  const position = snippet.substring(0, diIndex).trim();
  const company = snippet.substring(diIndex + 4).trim();
  if (!position || !company) return null;
  return [position, company];
}

function getOrCreateSheet(sheetName, headers) {
  const spreadsheet = SpreadsheetApp.getActive() || SpreadsheetApp.create("Job Applications Watch");
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.getActiveSheet();
    if (sheet.getName() !== sheetName) {
      sheet.setName(sheetName);
    }
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
    }
  }
  return sheet;
}

function decodeHtmlEntities(text) {
  return text
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#039;/g, "'")
    .replace(/&#8202;/g, ' ');
}

function writeToSheet(sheet, headers, data) {
  if (!data.length) return;
  const processedData = data.map(row => row.map(field => {
    if (typeof field === 'string') {
      return decodeHtmlEntities(field);
    }
    return field;
  }));
  sheet.getRange(sheet.getLastRow() + 1, 1, processedData.length, headers.length).setValues(processedData);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Job Applications')
    .addItem('Update Application Data', 'filterAndCategorizeEmails')
    .addToUi();
}
