// Sheet indices:
// 0: Deal Processing Sheet
// 1: Supplier Deals Sheet
// 2: Client Deals Sheet
// 3: Partner List
// 4: Funnel List

function initialSetupTrigger() {
  console.log("initialSetupTrigger started");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];  // Deal Processing Sheet
  var cell = sheet.getRange("J1");
  
  // Create dropdown list
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(['Ready', 'Processing'], true).build();
  cell.setDataValidation(rule);
  
  // Set default value and color
  cell.setValue("Ready");
  cell.setBackground("green");
  console.log("initialSetupTrigger completed");
}

function processAndTransfer() {
  console.log("processAndTransfer started");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheets()[0];  // Deal Processing Sheet
  var sheet2 = ss.getSheets()[1];  // Supplier Deals Sheet
  var funnelSheet = ss.getSheets()[4];  // Funnel List Sheet

  if (!sheet1 || !sheet2 || !funnelSheet) {
    console.error("One or more sheets not found");
    return;
  }

  var lastRow = sheet1.getLastRow();
  var data = sheet1.getRange(3, 1, lastRow - 2, 9).getValues();
  console.log(`Processing ${data.length} rows`);

  var today = new Date();
  var dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var newRows = [];
  var funnelsToAdd = new Set();
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][1] && data[i][2]) {  // Ensure essential data is present
      console.log(`Processing row ${i + 3}`);
      var dealId = generateDealId(data[i], today);
      var formattedFunnels = data[i][7].replace(/[\[\]]/g, '').split(',').map(f => f.trim()).join(', ');
      var newRow = [
        dateString,  // Date
        dealId,      // Deal ID
        data[i][1],  // Partner
        "",          // Partner Priority (new column in Supplier Deals Sheet)
        data[i][2],  // Geo
        data[i][3],  // Language
        data[i][4],  // Source
        data[i][5],  // CPA
        data[i][6],  // CRG
        data[i][8],  // CR
        "",          // CPL (empty for now)
        formattedFunnels,  // Funnels (formatted)
        "",          // EPL (empty, to be calculated manually)
        "",          // Quality (empty)
        "",          // Affiliate/Brand Interested (empty)
        ""           // FULL DEAL (removed, now calculated in the sheet)
      ];
      newRows.push(newRow);
      
      // Add to funnels to be processed later
      formattedFunnels.split(', ').forEach(funnel => {
        if (funnel !== "") funnelsToAdd.add(funnel);
      });
    }
  }

  // Batch append rows to Supplier Deals Sheet at the correct position
  if (newRows.length > 0) {
    var lastNonEmptyRow = getLastNonEmptyRow(sheet2);  // Find the correct last non-empty row
    sheet2.getRange(lastNonEmptyRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    console.log("Rows appended to Supplier Deals Sheet in bulk");
  }

  // Process funnels in bulk
  processFunnels(Array.from(funnelsToAdd), funnelSheet);

  // Clear processed rows from sheet1 in bulk
  sheet1.getRange(3, 1, lastRow - 2, 9).clearContent();
  console.log("Cleared processed rows from Deal Processing Sheet in bulk");

  // Reset J1 to "Ready" and green
  var triggerCell = sheet1.getRange("J1");
  triggerCell.setValue("Ready");
  triggerCell.setBackground("green");
  console.log("processAndTransfer completed");
}

// Helper function to find the last non-empty row in the Supplier Deals Sheet
function getLastNonEmptyRow(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  
  // Traverse rows from bottom to top to find the actual last row with data
  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i].some(cell => cell !== "")) {
      return i + 1;  // Return the row number (i is 0-based)
    }
  }
  return 0;  // No data found, return row 0 (first available row)
}

function processFunnels(funnels, funnelSheet) {
  console.log("processFunnels started");
  if (!funnelSheet) {
    console.error("Funnel sheet not found");
    return;
  }

  var lastRow = funnelSheet.getLastRow();
  var existingFunnels = lastRow > 0 ? funnelSheet.getRange("A1:A" + lastRow).getValues().flat() : [];
  var funnelsToAdd = funnels.filter(funnel => !existingFunnels.includes(funnel));

  if (funnelsToAdd.length > 0) {
    var newFunnelRows = funnelsToAdd.map(funnel => [funnel]);
    funnelSheet.getRange(funnelSheet.getLastRow() + 1, 1, newFunnelRows.length, 1).setValues(newFunnelRows);
    console.log(`${funnelsToAdd.length} new funnels added in bulk`);
  }
  console.log("processFunnels completed");
}

function generateDealId(rowData, date) {
  var geo = rowData[2].toUpperCase().replace(/\s+/g, '');  // Remove spaces from geo code as well
  var partner = rowData[1].replace(/\s+/g, '');  // Remove spaces from partner name
  var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "ddMMyyyy");
  return `${geo}-${partner}-${formattedDate}`;
}

function onEdit(e) {
  console.log("onEdit triggered");
  var range = e.range;
  var sheet = range.getSheet();
  var value = range.getValue();

  console.log(`Sheet index: ${sheet.getIndex()}, Range: ${range.getA1Notation()}, Value: ${value}`);

  if (sheet.getIndex() === 1) {  // Deal Processing Sheet
    if (range.getA1Notation() === "J1") {
      if (value === "Processing") {
        range.setBackground("yellow");
        try {
          processAndTransfer();
        } catch (error) {
          console.error(`Error in processAndTransfer: ${error}`);
          range.setValue("Ready");
          range.setBackground("red");
        }
      } else if (value === "Ready") {
        range.setBackground("green");
      }
    } else if (range.getColumn() === 1 && range.getRow() >= 3) {
      processInSheet1();
    }
  }
  console.log("onEdit completed");
}

function processInSheet1() {
  console.log("processInSheet1 started");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheets()[0];  // Deal Processing Sheet
  
  var lastRow = sheet1.getLastRow();
  var data = sheet1.getRange(3, 1, lastRow - 2, 1).getValues();
  
  for (var i = 0; i < data.length; i++) {
    var dealData = data[i][0];
    if (dealData) {
      console.log(`Processing row ${i + 3}`);
      var parts = dealData.split("-");
      
      if (parts.length >= 6) {
        sheet1.getRange(i + 3, 2).setValue(parts[0].trim());  // Partner
        sheet1.getRange(i + 3, 3).setValue(parts[1].trim());  // Geo Code
        sheet1.getRange(i + 3, 4).setValue(parts[2].trim());  // Language
        sheet1.getRange(i + 3, 5).setValue(parts[3].trim());  // Source
        sheet1.getRange(i + 3, 6).setValue(parts[4].trim());  // CPA
        sheet1.getRange(i + 3, 7).setValue(Number(parts[5].trim()) / 100);  // CRG
        sheet1.getRange(i + 3, 8).setValue(parts.slice(6, -1).join("-").trim());  // Funnels
        sheet1.getRange(i + 3, 9).setValue(parts[parts.length - 1].trim());  // CR
        console.log(`Row ${i + 3} processed successfully`);
      } else {
        sheet1.getRange(i + 3, 2).setValue("Invalid Format");
        console.warn(`Invalid format in row ${i + 3}`);
      }
    }
  }
  console.log("processInSheet1 completed");
}