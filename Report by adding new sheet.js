const sheetImport = 'PTO-Import';
const sheetReport = 'PTO-Report';

function createReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetImport);
  var rowLast = sheet.getLastRow();
  var rangeN = sheet.getRange('N2:N' + rowLast);
  var rangeD = sheet.getRange('D2:D'+ rowLast);
  var rangeB = sheet.getRange('A2:A'+ rowLast);
  var rangeG = sheet.getRange('G2:G'+ rowLast);
  var rangeH = sheet.getRange('H2:H'+ rowLast);
  var rangeI = sheet.getRange('I2:I'+ rowLast);
  var rangeJ = sheet.getRange('J2:J'+ rowLast);
  var rangeK = sheet.getRange('K2:K'+ rowLast);
  var valuesN = rangeN.getValues();
  var valuesD = rangeD.getValues();
  var valuesB = rangeB.getValues();
  var valuesG = rangeG.getValues();
  var valuesH = rangeH.getValues();
  var valuesI = rangeI.getValues();
  var valuesJ = rangeJ.getValues();
  var valuesK = rangeK.getValues();
  var reportData = [['Manager', 'Reports','Max PTO' ,'Available', 'Taken', 'Planned PTO', 'Remaining']];

  
  for (var i = 0; i < valuesN.length; i++) {
    var currentValueN = valuesN[i][0];
    var matchingValuesB = [];
    var matchingValuesG = [];
    var matchingValuesH = [];
    var matchingValuesI = [];
    var matchingValuesJ = [];
    var matchingValuesK = [];
    
    for (var j = 0; j < valuesD.length; j++) {
      var currentValueD = valuesD[j][0];
      
      if (currentValueN === currentValueD) {
        matchingValuesB.push(valuesB[j][0]);
      }
    }
    
    for (var k = 0; k < matchingValuesB.length; k++) {
      var currentMatchingValueB = matchingValuesB[k];
      
      for (var l = 0; l < valuesB.length; l++) {
        if (currentMatchingValueB === valuesB[l][0]) {
          matchingValuesG.push(valuesG[l][0]);
          matchingValuesH.push(valuesH[l][0]);
          matchingValuesI.push(valuesI[l][0]);
          matchingValuesJ.push(valuesJ[l][0]);
          matchingValuesK.push(valuesK[l][0]);
        }
      }
    }
    
    for (var m = 0; m < matchingValuesB.length; m++) {
      var row = [];
      row.push(currentValueN);
      row.push(matchingValuesB[m]);
      row.push(matchingValuesG[m]);
      row.push(matchingValuesH[m]);
      row.push(matchingValuesI[m]);
      row.push(matchingValuesJ[m]);
      row.push(matchingValuesK[m]);
      reportData.push(row);
    }
  }
  
  var reportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetReport);
  reportSheet.getRange('A1:K').clearContent();
  var reportRange = reportSheet.getRange(1, 1, reportData.length, 7);
  reportRange.setValues(reportData);
  
  // Apply formatting to the report sheet
  reportSheet.getRange("A1:G1").setFontWeight("bold");
  reportSheet.autoResizeColumns(1, 7);
  

}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Manual trigger")
    .addItem("Update Report", "createReport")
    .addItem("Send Reports to Managers","sendPTOReports")
    .addToUi();
}
