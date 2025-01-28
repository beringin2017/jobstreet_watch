function filterAndCategorizeEmails() {
  var searchQuery = '("Lamaranmu untuk posisi" OR "Terima kasih atas minat Anda")';
  var threads = GmailApp.search(searchQuery);
  Logger.log("Found " + threads.length + " threads matching the search query.");

  var categorizedEmails = [];

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var body = stripHtml(message.getBody());
      
      // Check for "Sent Application" message
      if (body.includes("Lamaranmu untuk posisi") && body.includes("berhasil dikirimkan ke")) {
        var applicationNameMatch = body.match(/Lamaranmu untuk posisi\s+(.*?)\s+berhasil/);
        var companyNameMatch = body.match(/berhasil dikirimkan ke\s+(.*?)\s*&#8202;/);

        if (applicationNameMatch && companyNameMatch) {
          var applicationName = decodeHtmlEntities(applicationNameMatch[1].trim());
          var companyName = cleanCompanyName(companyNameMatch[1].trim());
          
          categorizedEmails.push([
            message.getDate(),
            applicationName,
            companyName,
            "Sent"
          ]);
        } else {
          Logger.log("No match for Sent Application in this email.");
        }
      } 
      
      // Check for "Application Status" message
      else if (body.includes("Terima kasih atas minat Anda pada") && body.includes("Sayangnya")) {
        var applicationNameMatch = body.match(/pada\s*(.*?)\s*lowongan/);
        var companyNameMatch = body.match(/di\s+([^\.\n<]+)\s*\.\s*Sayangnya/);  // Updated regex
        
        if (applicationNameMatch && companyNameMatch) {
          var applicationName = decodeHtmlEntities(applicationNameMatch[1] || "N/A");  // Decode HTML entities and handle null
          var companyName = cleanCompanyName(companyNameMatch[1] || "N/A");  // Clean the company name and handle null
          var status = "Rejected"; // Status is "Rejected" for Application Status
          Logger.log("Matched Application Status: " + applicationName + " - " + companyName + " - " + status);

          categorizedEmails.push([
            message.getDate(),
            applicationName,
            companyName,
            status,  // Status is "Rejected" for Application Status
          ]);
        } else {
          Logger.log("No match for Application Status in this email.");
        }
      } else {
        Logger.log("No match for this email.");
      }
    }
  }

  // Write to "Sheet 1"
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) spreadsheet = SpreadsheetApp.create("Jobseek Watch");

  // Access the "Sheet 1"
  var sheet = spreadsheet.getSheetByName("Sheet 1");
  if (!sheet) {
    sheet = spreadsheet.insertSheet("Sheet 1"); // Create "Sheet 1" if it doesn't exist
  }

  // Updated header to match 4 columns: Date, Application Name, Company Name, Status
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", "Application Name", "Company Name", "Status"]);
  }

  if (categorizedEmails.length > 0) {
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, categorizedEmails.length, categorizedEmails[0].length).setValues(categorizedEmails);
  }

  // After data is inserted, clean up HTML entities from the sheet
  cleanHtmlEntitiesFromSheet(sheet);

  Logger.log("Emails categorized and exported successfully!");
}

// Helper function to remove HTML tags from the email body
function stripHtml(html) {
  // Remove <style> blocks and all HTML tags
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '') // Remove CSS
    .replace(/<[^>]+>/g, '') // Remove HTML tags
    .replace(/\s+/g, ' ') // Collapse whitespace
    .trim();
}
// Helper function to clean the company name, removing unwanted text after &#8202;
function cleanCompanyName(companyName) {
  return companyName
    .replace(/&amp;/g, '&') // Fix HTML entities
    .replace(/\s+/g, ' ') // Normalize spaces
    .trim();
}

// Helper function to decode HTML entities (like &amp;) into their ASCII characters
function decodeHtmlEntities(input) {
  var doc = HtmlService.createHtmlOutput(input);
  return doc.getContent(); // Decodes HTML entities into their respective characters
}

// Helper function to clean HTML entities from the entire sheet
function cleanHtmlEntitiesFromSheet(sheet) {
  var range = sheet.getDataRange(); // Get the entire data range
  var values = range.getValues(); // Get the values in the range

  // Loop through all rows and columns to clean HTML entities
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (typeof values[i][j] === 'string') {
        values[i][j] = decodeHtmlEntities(values[i][j]); // Decode HTML entities
      }
    }
  }

  // Set the cleaned values back into the sheet
  range.setValues(values);
}
