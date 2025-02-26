# jobstreet_watch
This Google Apps Script automates the process of fetching and categorizing job application emails from Gmail and logs them into a Google Spreadsheet. It extracts key details such as application date, position, company name, and status (Sent, Rejected, or Expired).

## Features
- Automatically fetches job application emails based on specific keywords.
- Extracts relevant details using regex pattern matching.
- Categorizes applications into Sent, Rejected, and Expired.
- Updates an existing Google Spreadsheet without creating new sheets.
- Renames the default active sheet to "watchlist" for consistency.

## Setup Instructions
1. Open [Google Apps Script](https://script.google.com/) and create a new project.
2. Copy and paste the provided script into the editor.
3. Save and authorize the script to access Gmail and Google Sheets.
4. Ensure that your Google Spreadsheet is open and active before running the script.
5. Run `filterAndCategorizeEmails` to process and log application data.

## Usage
- The script runs on-demand via the "Job Applications" menu in Google Sheets.
- It scans Gmail for job-related emails and logs structured data into the spreadsheet.
- The script intelligently processes messages, removing unnecessary text and extracting useful details.


![image](https://github.com/user-attachments/assets/0832f5a1-f1f6-4d80-9857-4aa883f03598)


