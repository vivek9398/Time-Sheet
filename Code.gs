function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateEmployeeSheet(data.employeeId);
    
    if (data.type === 'login') {
      sheet.appendRow([data.date, data.timestamp, '', '']);
    } else {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const loginTime = sheet.getRange(lastRow, 2).getValue();
        const workingHours = calculateWorkingHours(loginTime, data.timestamp);
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

function calculateWorkingHours(loginTime, logoutTime) {
  try {
    // Parse Indian time format: "25/07/2025, 04:38:09 pm"
    const parseTime = (timeStr) => {
      const [datePart, timePart] = timeStr.split(', ');
      const [day, month, year] = datePart.split('/');
      const [time, period] = timePart.split(' ');
      const [hours, minutes, seconds] = time.split(':');
      
      let hour24 = parseInt(hours);
      if (period.toLowerCase() === 'pm' && hour24 !== 12) hour24 += 12;
      if (period.toLowerCase() === 'am' && hour24 === 12) hour24 = 0;
      
      return new Date(year, month - 1, day, hour24, parseInt(minutes), parseInt(seconds));
    };
    
    const login = parseTime(loginTime);
    const logout = parseTime(logoutTime);
    const diffMs = logout - login;
    const hours = Math.floor(diffMs / (1000 * 60 * 60));
    const minutes = Math.floor((diffMs % (1000 * 60 * 60)) / (1000 * 60));
    return `${hours}h ${minutes}m`;
  } catch (error) {
    return 'Error';
  }
}