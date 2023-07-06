const Importsheet = 'PTO-Import';
const Reportsheet = 'PTO-Report';

function sPTOReports() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Step 1: Get managers' names and email addresses
  var managersSheet = spreadsheet.getSheetByName(Reportsheet);
  var managersRange = managersSheet.getRange("N1:O");
  var managersData = managersRange.getValues();
  
  var managers = {};
  
  for (var i = 1; i < managersData.length; i++) {
    var managerName = managersData[i][0];
    var managerEmail = managersData[i][1];
    
    if (managerName && managerEmail) {
      managers[managerName] = managerEmail;
    }
  }
  
  // Step 2: Collect reportees' data for each manager
  var reporteesData = {};
  var reporteesSheet = spreadsheet.getSheetByName(Reportsheet);
  var lastRow = reporteesSheet.getLastRow();
  var dataRange = reporteesSheet.getRange("A1:G" + lastRow);
  var dataValues = dataRange.getValues();
  
  for (var j = 0; j < dataValues.length; j++) {
    var reporteeName = dataValues[j][1];
    var managerName = dataValues[j][0];
    
    if (managerName && managers[managerName]) {
      if (!reporteesData[managerName]) {
        reporteesData[managerName] = [];
      }
      
      reporteesData[managerName].push(dataValues[j]);
    }
  }
  
  // Step 3: Send PTO reports to managers' email addresses
  for (var manager in reporteesData) {
    var reportees = reporteesData[manager];
    var managerEmail = managers[manager];
    
    var htmlTable = '<table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;background-color: transparent; margin: 0 auto;">';
    htmlTable += '<thead style="background-color: #f2f2f2;">';
    htmlTable += '<tr>';
    htmlTable += '<th style="border: 1px solid #dddddd; padding: 10px; text-align: center;" colspan="6">' +  ' PTO Report - ' + reporteesSheet.getRange("Z1").getValue() + '</th>';
    htmlTable += '</tr>';
    htmlTable += '<tr>';
    htmlTable += '<th style="border: 1px solid #dddddd; padding: 10px; text-align: left;">Report Name</th>';
    htmlTable += '<th style="border: 1px solid #dddddd; padding: 10px; text-align: center;">Max PTO</th>';
    htmlTable += '<th style="border: 1px solid #dddddd; padding: 10px; text-align: center;">Available</th>';
    htmlTable += '<th style="border: 1px solid #dddddd; padding: 10px; text-align: center;">Taken</th>';
    htmlTable += '<th style="border: 1px solid #dddddd; padding: 10px; text-align: center;">Planned PTO</th>';
    htmlTable += '<th style="border: 1px solid #dddddd; padding: 10px; text-align: center;">Remaining</th>';
    htmlTable += '</tr>';
    htmlTable += '</thead>';
    htmlTable += '<tbody>';
    
    for (var k = 0; k < reportees.length; k++) {
      var reportee = reportees[k];
      
      htmlTable += '<tr>';
      
      for (var l = 1; l < reportee.length; l++) {
        var cellBackgroundColor = reporteesSheet.getRange(k + 2, l + 1).getBackground();
        var cellAlignment = (l === 1) ? 'left' : 'center';
        
        htmlTable += '<td style="border: 1px solid #dddddd; padding: 10px; background-color: ' + cellBackgroundColor + '; text-align: ' + cellAlignment + ';">' + reportee[l] + '</td>';
      }
      
      htmlTable += '</tr>';
    }
    
    htmlTable += '</tbody>';
    htmlTable += '</table>';
    
    var recipient = managerEmail;
    var subject = 'PTO Report - ' + reporteesSheet.getRange("Z1").getValue();
    var body = '<div style="background-color: #ffffff; padding: 20px; text-align: center;">';
    body += '<h2 style="font-size: 15px; margin-bottom: 10px; text-align: left;">Hi ' + manager + ',</h2>';
    body += '<p style="text-align: left;">Below is the updated PTO report for all your reportees currently till ' + reporteesSheet.getRange("Z1").getValue() + '</p>';
    body += htmlTable;
    body += '<p style="text-align: left;">Feel free to reach out if you have any questions/concerns.</p>';
    body += '<p style="text-align: left;">Kindly,</p>';
    body += '<p style="text-align: left;">Support-Operations</p>';
    body += '</div>';
    
    GmailApp.sendEmail(recipient, subject, '', { htmlBody: body });
  }
}
