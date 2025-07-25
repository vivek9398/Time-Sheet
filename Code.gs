function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = 'Employee_' + data.employeeId;
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(['Date', 'Login Time', 'Logout Time', 'Working Hours']);
    }
    
    if (data.type === 'login') {
      sheet.appendRow([data.date, data.timestamp, '', '']);
    } else {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const loginTime = sheet.getRange(lastRow, 2).getValue();
        const workingHours = calculateHours(loginTime, data.timestamp);
        sheet.getRange(lastRow, 3).setValue(data.timestamp);
        sheet.getRange(lastRow, 4).setValue(workingHours);
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

function calculateHours(loginTime, logoutTime) {
  try {
    const parseTime = (timeStr) => {
      const parts = timeStr.split(', ');
      const datePart = parts[0].split('/');
      const timePart = parts[1].split(' ');
      const time = timePart[0].split(':');
      const period = timePart[1];
      
      let hours = parseInt(time[0]);
      if (period === 'pm' && hours !== 12) hours += 12;
      if (period === 'am' && hours === 12) hours = 0;
      
      return new Date(datePart[2], datePart[1] - 1, datePart[0], hours, parseInt(time[1]), parseInt(time[2]));
    };
    
    const login = parseTime(loginTime);
    const logout = parseTime(logoutTime);
    const diffMs = logout - login;
    const hours = Math.floor(diffMs / (1000 * 60 * 60));
    const minutes = Math.floor((diffMs % (1000 * 60 * 60)) / (1000 * 60));
    return hours + 'h ' + minutes + 'm';
  } catch (error) {
    return 'Error';
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

