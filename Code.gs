/**
 * برج التوفيق | Tower Management System - Google Apps Script Backend
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (action === 'getAllData') {
      var result = {
        units: getTableData(ss, 'Units'),
        tenants: getTableData(ss, 'Tenants'),
        owners: getTableData(ss, 'Owners'),
        contracts: getTableData(ss, 'Contracts'),
        rentDue: getTableData(ss, 'RentDue'),
        rentPayments: getTableData(ss, 'RentPayments'),
        withdrawals: getTableData(ss, 'Withdrawals'),
        expenses: getTableData(ss, 'Expenses'),
        openingBalances: getTableData(ss, 'OpeningBalances'),
        ownerAllocations: getTableData(ss, 'OwnerAllocations'),
        cashLedger: getTableData(ss, 'CashLedger'),
        auditLog: getTableData(ss, 'AuditLog'),
        settings: getSettings(ss)
      };
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: result })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'saveData') {
      var table = data.table;
      var rows = data.data;
      saveTableData(ss, table, rows);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Unknown action' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("Tawfiq Tower Management System API is running successfully.");
}

function getTableData(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // Initialize default headers if needed
    return [];
  }
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var headers = rows[0];
  var data = [];
  for (var i = 1; i < rows.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = rows[i][j];
    }
    data.push(obj);
  }
  return data;
}

function saveTableData(ss, sheetName, dataArray) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  sheet.clear();
  if (!dataArray || dataArray.length === 0) return;
  
  var keys = Object.keys(dataArray[0]);
  var output = [keys];
  for (var i = 0; i < dataArray.length; i++) {
    var row = [];
    for (var j = 0; j < keys.length; j++) {
      row.push(dataArray[i][keys[j]] !== undefined ? dataArray[i][keys[j]] : '');
    }
    output.push(row);
  }
  sheet.getRange(1, 1, output.length, keys.setValues ? keys.length : keys.length).setValues(output);
}

function getSettings(ss) {
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) return { towerName: "برج التوفيق", address: "القاهرة، مصر", currency: "جنيه مصري", financialMonth: "2026-10" };
  var rows = sheet.getDataRange().getValues();
  var settings = { towerName: "برج التوفيق", address: "القاهرة، مصر", currency: "جنيه مصري", financialMonth: "2026-10" };
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] === 'towerName') settings.towerName = rows[i][1];
    if (rows[i][0] === 'address') settings.address = rows[i][1];
    if (rows[i][0] === 'currency') settings.currency = rows[i][1];
    if (rows[i][0] === 'financialMonth') settings.financialMonth = rows[i][1];
  }
  return settings;
}

function setupTawfiqTowerSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNames = [
    'Units', 'Tenants', 'Owners', 'Contracts', 'RentDue', 
    'RentPayments', 'Withdrawals', 'Expenses', 'OpeningBalances', 
    'OwnerAllocations', 'CashLedger', 'AuditLog', 'Settings'
  ];
  
  for (var i = 0; i < sheetNames.length; i++) {
    var name = sheetNames[i];
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      ss.insertSheet(name);
    }
  }
  SpreadsheetApp.getUi().alert('تم إعداد جداول برج التوفيق بنجاح!');
}
