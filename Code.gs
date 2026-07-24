function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Employee Attendance System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. MASTER LOGIN (Bypass for Admin)
function verifyLogin(empId, password) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Employees");
  var data = sheet.getDataRange().getValues();
  for(var i = 1; i < data.length; i++) {
    // .toString().trim() use kiya hai taaki Number aur Text ka issue na aaye
    if(data[i][0].toString().trim() === empId.toString().trim() && data[i][3].toString().trim() === password.toString().trim()) {
      
      // Bypass: Ab jo bhi login karega (jaise 101), usko direct Admin panel dikhega
      return { success: true, name: data[i][1], role: "Admin" }; 
    }
  }
  return { success: false, message: "Invalid ID or Password" };
}

// 2. PUNCH IN / OUT (With Fixes)
function punchIn(empId, location) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");
  var date = new Date().toLocaleDateString();
  var time = new Date().toLocaleTimeString();
  sheet.appendRow([empId, date, time, "", location, ""]);
  return "Punched In Successfully at " + time;
}

function punchOut(empId, location) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");
  var data = sheet.getDataRange().getValues();
  var time = new Date().toLocaleTimeString();
  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][0].toString().trim() === empId.toString().trim() && data[i][3] === "") { 
      sheet.getRange(i + 1, 4).setValue(time); 
      sheet.getRange(i + 1, 6).setValue(location); 
      return "Punched Out Successfully at " + time;
    }
  }
  return "Error: Please Punch In first!";
}
// 3. HISTORY DASHBOARD (Clean Date & Time Fix)
function getMyHistory(empId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");
    
    // SABSE BADA JAADU YAHI HAI: getValues() ki jagah getDisplayValues() likha hai
    // Isse exactly wahi uthega jo sheet mein dikhta hai (jaise 24/07/2026 aur 10:15 AM)
    var data = sheet.getDataRange().getDisplayValues(); 
    var history = [];
    
    for(var i = 1; i < data.length; i++) { 
      if(data[i][0].toString().trim() === empId.toString().trim()) {
        
        var dDate = data[i][1] ? data[i][1] : "-";
        var dIn = data[i][2] ? data[i][2] : "-";
        var dOut = data[i][3] ? data[i][3] : "-";
        
        history.push([dDate, dIn, dOut]); 
      }
    }
    return history;
  } catch(e) {
    return [["Error", String(e), ""]];
  }
}
// 4. EMPLOYEE MANAGEMENT (Isme bhi text format fix kar diya)
function getAllEmployees() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Employees");
  var data = sheet.getDataRange().getValues();
  var emps = [];
  for(var i = 1; i < data.length; i++) {
    emps.push([String(data[i][0]), String(data[i][1]), String(data[i][4])]); 
  }
  return emps;
}
