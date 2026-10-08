function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getBooks";
  
  // Return all books as JSON for the web app
  if (action === "getBooks") {
    const data = sheet.getDataRange().getValues();
    const books = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      books.push({
        id: i,
        title: row[0].toString(),
        subtitle: row[1] ? row[1].toString() : "",
        category: row[2] ? row[2].toString() : "Uncategorized",
        copies: parseInt(row[3], 10) || 1,
        status: row[4] ? row[4].toString() : "Available",
        publisher: row[5] ? row[5].toString() : "",
        checkedOutBy: row[6] ? row[6].toString() : null,
        checkedOutDate: row[7] ? row[7].toString() : null
      });
    }
    return ContentService.createTextOutput(JSON.stringify(books))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Handle checkout and return actions
  const bookTitle = e.parameter.title || "";
  const borrower = e.parameter.borrower || "";
  const phone = e.parameter.phone || "";
  const address = e.parameter.address || "";
  const today = Utilities.formatDate(new Date(), "GMT-7", "yyyy-MM-dd");

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === bookTitle.trim().toLowerCase()) {
      const row = i + 1;
      if (action === "checkout") {
        sheet.getRange(row, 5).setValue("Checked Out"); // Col E
        sheet.getRange(row, 7).setValue(borrower + " (" + phone + ")"); // Col G
        sheet.getRange(row, 8).setValue(today); // Col H
      } else if (action === "return") {
        sheet.getRange(row, 5).setValue("Available");
        sheet.getRange(row, 7).setValue("");
        sheet.getRange(row, 8).setValue("");
      }
      break;
    }
  }

  // Member logging
  if (action === "checkout" && phone) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let memberSheet = ss.getSheetByName("Members");
    if (!memberSheet) {
      memberSheet = ss.insertSheet("Members");
      memberSheet.appendRow(["Phone (Member ID)", "Full Name", "Address", "Last Borrowed Date"]);
    }
    const members = memberSheet.getDataRange().getValues();
    let found = false;
    for (let j = 1; j < members.length; j++) {
      if (members[j][0] && members[j][0].toString() === phone.toString()) {
        memberSheet.getRange(j + 1, 4).setValue(today);
        found = true;
        break;
      }
    }
    if (!found) {
      memberSheet.appendRow([phone, borrower, address, today]);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}
