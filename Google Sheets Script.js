function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Custom Menu');
  menu.addItem('Set Up Trigger', 'setupTrigger');
  menu.addItem('Update Partner Types', 'updatePartnerTypes');
  menu.addToUi();
}

function setupTrigger() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var cell = sheet.getRange("J1");
  
  // Create dropdown list
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(['Ready', 'Processing'], true).build();
  cell.setDataValidation(rule);
  
  // Set default value and color
  cell.setValue("Ready");
  cell.setBackground("green");
}

function processAndTransfer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheets()[0];
  var sheet2 = ss.getSheets()[1];
  var funnelSheet = ss.getSheets()[4]; // 5th sheet for funnel list
  
  var lastRow = sheet1.getLastRow();
  var data = sheet1.getRange(2, 1, lastRow - 1, 9).getValues();
  
  var today = new Date();
  var dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) {
      var dealId = generateDealId(data[i], today);
      var newRow = [
        dateString,  // Date
        dealId,      // Deal ID
        data[i][1],  // Partner
        data[i][2],  // Geo
        data[i][3],  // Language
        data[i][4],  // Source
        data[i][5],  // CPA
        data[i][6],  // CRG
        "",          // CPL (empty for now)
        data[i][7],  // Funnels
        "",          // EPL (empty, to be calculated manually)
        "",          // Quality (empty)
        "",          // Affiliate/Brand Interested (empty)
        formatFullDeal(data[i])  // FULL DEAL
      ];
      sheet2.appendRow(newRow);
      
      // Process funnels
      processFunnels(data[i][7], funnelSheet);
      
      // Clear the processed row from sheet1
      sheet1.getRange(i + 2, 1, 1, 9).clearContent();
      
      // Force the spreadsheet to update its display
      SpreadsheetApp.flush();
    }
  }
  
  // Reset J1 to "Ready" and green
  var triggerCell = sheet1.getRange("J1");
  triggerCell.setValue("Ready");
  triggerCell.setBackground("green");
  
  // Update partner types
  updatePartnerTypes();
}

function processFunnels(funnelsString, funnelSheet) {
  // Remove brackets and split by comma
  var funnels = funnelsString.replace(/[\[\]]/g, '').split(',').map(funnel => funnel.trim());
  var existingFunnels = funnelSheet.getRange("A1:A" + funnelSheet.getLastRow()).getValues().flat();
  
  funnels.forEach(funnel => {
    if (!existingFunnels.includes(funnel) && funnel !== "") {
      funnelSheet.appendRow([funnel]);
    }
  });
}

function generateDealId(rowData, date) {
  var geo = rowData[2].toUpperCase();
  var partner = rowData[1].replace(/\s+/g, '');  // Remove spaces from partner name
  var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "ddMMyyyy");
  return `${geo}-${partner}-${formattedDate}`;
}

function formatFullDeal(rowData) {
  var source = rowData[4];
  var geo = rowData[2];
  var cpa = rowData[5];
  var crg = rowData[6];
  var funnels = rowData[7];
  var language = rowData[3];
  var cr = parseInt(crg*100) + 2;
  
  return `[${source}] ${geo} ${cpa}+${crg*100}% (Expected CR ${cr}%)\nFunnel(s): ${funnels} || Language: ${language}`;
}

function updatePartnerTypes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dealSheet = ss.getSheets()[1];  // 2nd sheet (deal tracker)
  var partnerSheet = ss.getSheets()[3];  // 4th sheet (partner list)
  
  // Get partner data
  var partnerData = partnerSheet.getRange("A2:B" + partnerSheet.getLastRow()).getValues();
  var partnerTypes = {};
  partnerData.forEach(row => {
    partnerTypes[row[0]] = row[1];  // Assuming column A is Partner Name and B is Type
  });
  
  // Update deal sheet
  var lastRow = dealSheet.getLastRow();
  var partnerColumn = 3;  // Column C
  var typeColumn = 13;    // Column M (Affiliate/Brand Interested)
  
  for (var i = 2; i <= lastRow; i++) {
    var partner = dealSheet.getRange(i, partnerColumn).getValue();
    var typeCell = dealSheet.getRange(i, typeColumn);
    
    if (partner in partnerTypes) {
      var type = partnerTypes[partner];
      typeCell.setValue(type);
      
      // Set color based on type
      if (type.toLowerCase() === "brand") {
        typeCell.setBackground("#b7e1cd");  // Light green
      } else if (type.toLowerCase() === "network") {
        typeCell.setBackground("#fce8b2");  // Light yellow
      } else {
        typeCell.setBackground(null);  // Clear background
      }
    } else {
      typeCell.setValue("");
      typeCell.setBackground(null);  // Clear background
    }
  }
  
  // Set data validation for the type column
  var types = Object.values(partnerTypes).filter((v, i, a) => a.indexOf(v) === i);  // Get unique types
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(types, true).build();
  dealSheet.getRange(2, typeColumn, lastRow - 1, 1).setDataValidation(rule);
}

function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var value = range.getValue();
  
  if (sheet.getIndex() === 1) {
    if (range.getA1Notation() === "J1") {
      if (value === "Processing") {
        range.setBackground("yellow");
        processAndTransfer();
      } else if (value === "Ready") {
        range.setBackground("green");
      }
    } else if (range.getColumn() === 1 && range.getRow() > 1) {
      processInSheet1();
    }
  }
}

function processInSheet1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheets()[0];
  
  var lastRow = sheet1.getLastRow();
  var data = sheet1.getRange(2, 1, lastRow - 1, 1).getValues();
  
  for (var i = 0; i < data.length; i++) {
    var dealData = data[i][0];
    if (dealData) {
      var parts = dealData.split("-");
      
      if (parts.length >= 6) {
        var partner = parts[0].trim();
        var geoCode = parts[1].trim();
        var language = parts[2].trim();
        var source = parts[3].trim();
        var cpa = parts[4].trim();
        var crg = parts[5].trim();
        var funnels = parts.slice(6, -1).join("-").trim();
        var cr = parts[parts.length - 1].trim();
        
        if (language.toLowerCase() === "fb" || language.toLowerCase() === "google") {
          source = language;
          language = "Native";
        }
        
        if (source.toLowerCase() === "fb") {
          source = "Facebook";
        }
        
        sheet1.getRange(i + 2, 2).setValue(partner);
        sheet1.getRange(i + 2, 3).setValue(geoCode);
        sheet1.getRange(i + 2, 4).setValue(language);
        sheet1.getRange(i + 2, 5).setValue(source);
        sheet1.getRange(i + 2, 6).setValue(cpa);
        sheet1.getRange(i + 2, 7).setValue(crg/100);
        sheet1.getRange(i + 2, 8).setValue(funnels);
        sheet1.getRange(i + 2, 9).setValue(cr + "%"); // CR column
      } else {
        sheet1.getRange(i + 2, 2).setValue("Invalid Format");
      }
    }
  }
}
