function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateEmployeeSheet(data.employeeId);
    
    if (data.type === 'login') {
      sheet.appendRow([data.date, data.timestamp, '', '']);
    } else {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const loginTimeCell = sheet.getRange(lastRow, 2);
        const loginTime = loginTimeCell.getDisplayValue() || loginTimeCell.getValue();
        sheet.getRange(lastRow, 3).setValue(data.timestamp);
        sheet.getRange(lastRow, 4).setValue('Calculated');
      }
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({success: true}))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeaders({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type'
      });
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({success: false, error: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeaders({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type'
      });
  }
}

function doOptions(e) {
  return ContentService
    .createTextOutput('')
    .setHeaders({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
}

function getOrCreateEmployeeSheet(employeeId) {
  const sheetName = `Employee_${employeeId}`;
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
    sheet.getRange(1, 1, 1, 4).setValues([['Date', 'Login Time', 'Logout Time', 'Working Hours']]);
  }
  
  return sheet;
}

