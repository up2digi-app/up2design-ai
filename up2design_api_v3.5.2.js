/**
 * UP2Design CMS API v3.5.2
 * READ + WRITE ENGINE (HOTFIX LEADS PARSER & WHATSAPP TEXT FORMAT)
 * Author: Senior Google Apps Script Developer
 */

const SHEET = SpreadsheetApp.getActiveSpreadsheet();

/**
 * ==========================================
 * DO GET (READ API)
 * ==========================================
 */
function doGet(e) {
  e = e || {};
  e.parameter = e.parameter || {};
  
  var action = e.parameter.action || "all";
  var success = true;
  var responseData = null;

  try {
    // PATCH: TEST ENDPOINT DENGAN RESPONSE CUSTOM
    if (action === "test") {
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        status: "online",
        timestamp: new Date().toISOString()
      })).setMimeType(ContentService.MimeType.JSON);
    }

    switch (action) {
      case "settings": responseData = getSettings(); break;
      case "portfolio": responseData = getPortfolio(); break;
      case "pricing": responseData = getPricing(); break;
      case "pricing_features": responseData = getPricingFeatures(); break;
      case "faq": responseData = getFaq(); break;
      case "testimonial": responseData = getTestimonial(); break;
      case "social": responseData = getSocial(); break;
      case "leads": responseData = getLeads(); break;
      case "all":
        responseData = {
          settings: getSettings(),
          portfolio: getPortfolio(),
          pricing: getPricing(),
          pricing_features: getPricingFeatures(),
          faq: getFaq(),
          testimonial: getTestimonial(),
          social: getSocial(),
          leads: getLeads()
        };
        break;
      default:
        success = false;
        responseData = { message: "Invalid GET action endpoint." };
    }
  } catch (error) {
    success = false;
    responseData = { message: error.toString(), stack: error.stack };
  }

  return formatResponse(success, action, responseData);
}

/**
 * ==========================================
 * DO POST (WRITE & LOGIN API)
 * ==========================================
 */
function doPost(e) {
  e = e || {};
  e.parameter = e.parameter || {};
  
  var action = e.parameter.action;
  var success = false;
  var responseData = null;

  try {
    switch (action) {
      case "login":
        responseData = loginUser(e.parameter.username, e.parameter.password);
        success = responseData.success;
        if(success) responseData = responseData.data;
        break;
        
      case "update_settings":
        var payload = JSON.parse(e.parameter.payload);
        responseData = updateSettings(payload);
        success = responseData.success;
        break;
        
      case "create":
        var sheetName = e.parameter.sheet_name;
        var payload = JSON.parse(e.parameter.payload);
        responseData = createRecord(sheetName, payload);
        success = responseData.success;
        break;
        
      case "update":
        var sheetName = e.parameter.sheet_name;
        var recordId = e.parameter.record_id;
        var payload = JSON.parse(e.parameter.payload);
        responseData = updateRecord(sheetName, recordId, payload);
        success = responseData.success;
        break;
        
      case "delete":
        var sheetName = e.parameter.sheet_name;
        var recordId = e.parameter.record_id;
        var role = e.parameter.role || 'editor'; // Default to editor for safety if not passed
        responseData = deleteRecord(sheetName, recordId, role);
        success = responseData.success;
        break;
        
      case "create_lead":
        var payload = {};
        if (e.postData && e.postData.contents) {
          payload = JSON.parse(e.postData.contents);
        } else if (e.parameter && e.parameter.payload) {
          payload = JSON.parse(e.parameter.payload);
        }
        responseData = createLead(payload);
        success = responseData.success;
        break;
        
      default:
        responseData = { message: "Invalid POST action endpoint." };
    }
  } catch (error) {
    success = false;
    responseData = { message: error.toString(), stack: error.stack };
  }

  return formatResponse(success, action, responseData);
}

/**
 * ==========================================
 * READ HANDLERS
 * ==========================================
 */
function getSettings() {
  var data = getSheetDataAsObjects('SETTINGS');
  var settingsObj = {};
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    // Dynamic Mapping: memastikan field_name terdeteksi sempurna
    if (row.field_name && row.field_name.toString().trim() !== "") {
      settingsObj[row.field_name.toString().trim()] = row.value;
    }
  }
  return settingsObj;
}

function getPortfolio() {
  return sortData(getSheetDataAsObjects('PORTFOLIO').filter(function(r) { return isTrue(r.active); }));
}

function getPricing() {
  var pricing = sortData(getSheetDataAsObjects('PRICING').filter(function(r) { return isTrue(r.active); }));
  var features = sortData(getSheetDataAsObjects('PRICING_FEATURES'));
  
  for (var i = 0; i < pricing.length; i++) {
    var pkg = pricing[i];
    pkg.features = features.filter(function(f) { return f.package_id == pkg.package_id; }).map(function(f) { return f.feature; });
  }
  return pricing;
}

function getPricingFeatures() {
  // Mengambil fitur harga secara terpisah dengan filter & sort
  var data = getSheetDataAsObjects('PRICING_FEATURES');
  var activeData = data.filter(function(r) {
    // Safe fallback jika kolom active belum ada di Spreadsheet PRICING_FEATURES
    if (r.active === undefined) return true;
    return isTrue(r.active);
  });
  return sortData(activeData);
}

function getFaq() {
  return sortData(getSheetDataAsObjects('FAQ').filter(function(r) { return isTrue(r.active); }));
}

function getTestimonial() {
  return sortData(getSheetDataAsObjects('TESTIMONIAL').filter(function(r) { return isTrue(r.active); }));
}

function getSocial() {
  return sortData(getSheetDataAsObjects('SOCIAL').filter(function(r) { return isTrue(r.active); }));
}

function getLeads() {
  // Menampilkan data leads (Tanpa filter, Diurutkan berdasarkan created_at DESC)
  var data = getSheetDataAsObjects('LEADS');
  return data.sort(function(a, b) {
    var dateA = new Date(a.created_at || 0).getTime();
    var dateB = new Date(b.created_at || 0).getTime();
    return dateB - dateA; // Sort DESC
  });
}

/**
 * ==========================================
 * WRITE & LOGIN HANDLERS
 * ==========================================
 */
function loginUser(username, password) {
  if (!username || !password) return { success: false, message: "Username dan Password wajib diisi." };
  
  var sheet = SHEET.getSheetByName('USERS');
  if (!sheet) return { success: false, message: "Sheet USERS tidak ditemukan." };
  
  var data = getSheetDataAsObjects('USERS');
  var passwordHash = hashPassword(password);
  
  for (var i = 0; i < data.length; i++) {
    var user = data[i];
    
    // Fallback Backward Compatibility: Cek hash SHA256 atau plain text
    var isHashMatch = (user.password_hash === passwordHash);
    var isPlainMatch = (user.password_hash === password);
    
    if (user.username === username && (isHashMatch || isPlainMatch)) {
      if (!isTrue(user.status)) {
        return { success: false, message: "Akun Anda tidak aktif. Hubungi Administrator." };
      }
      
      // Update Last Login dan migrasi Hash Password
      var now = new Date().toISOString();
      var rowIndex = findRowIndexById('USERS', 'id', user.id);
      if (rowIndex > -1) {
        var headers = sheet.getDataRange().getValues()[0];
        
        // Update Kolom last_login
        var colIndex = headers.indexOf('last_login') + 1;
        if (colIndex > 0) sheet.getRange(rowIndex, colIndex).setValue(now);
        
        // Auto-Migrasi: Jika password masih plain text, enkripsi dan simpan sebagai SHA256
        if (isPlainMatch && !isHashMatch) {
          var pwdIndex = headers.indexOf('password_hash') + 1;
          if (pwdIndex > 0) sheet.getRange(rowIndex, pwdIndex).setValue(passwordHash);
        }
      }
      
      return {
        success: true,
        data: {
          id: user.id,
          username: user.username,
          full_name: user.full_name,
          role: user.role,
          last_login: now
        }
      };
    }
  }
  
  return { success: false, message: "Username atau Password salah." };
}

function updateSettings(payload) {
  var sheet = SHEET.getSheetByName('SETTINGS');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var fieldCol = headers.indexOf('field_name');
  var valCol = headers.indexOf('value');
  
  if (fieldCol === -1 || valCol === -1) {
    fieldCol = 0; valCol = 1; // Fallback index jika header tidak sesuai
  }

  for (var key in payload) {
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][fieldCol] === key) {
        sheet.getRange(i + 1, valCol + 1).setValue(payload[key]);
        found = true;
        break;
      }
    }
    // Auto-Append: Jika field (misal gas_api_url) belum ada di Spreadsheet, tambahkan baris baru otomatis
    if (!found) {
      var newRow = new Array(headers.length).fill("");
      newRow[fieldCol] = key;
      newRow[valCol] = payload[key];
      sheet.appendRow(newRow);
    }
  }
  
  return { success: true, message: "Settings saved successfully." };
}

function createRecord(sheetName, payload) {
  var sheet = SHEET.getSheetByName(sheetName);
  if (!sheet) return { success: false, message: "Sheet " + sheetName + " tidak ditemukan." };
  
  var headers = sheet.getDataRange().getValues()[0];
  var newRow = new Array(headers.length).fill("");
  
  // Auto Generate ID
  var idField = sheetName === 'PRICING' ? 'package_id' : 'id';
  if (!payload[idField]) payload[idField] = generateNextId();
  
  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    if (payload[header] !== undefined) {
      newRow[i] = payload[header];
    }
  }
  
  sheet.appendRow(newRow);
  return { success: true, message: "Data berhasil ditambahkan." };
}

function updateRecord(sheetName, recordId, payload) {
  var sheet = SHEET.getSheetByName(sheetName);
  var idField = sheetName === 'PRICING' ? 'package_id' : 'id';
  var rowIndex = findRowIndexById(sheetName, idField, recordId);
  
  if (rowIndex === -1) return { success: false, message: "Data tidak ditemukan." };
  
  var headers = sheet.getDataRange().getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    if (payload[header] !== undefined) {
      sheet.getRange(rowIndex, i + 1).setValue(payload[header]);
    }
  }
  
  return { success: true, message: "Data berhasil diupdate." };
}

function deleteRecord(sheetName, recordId, role) {
  // ROLE ENGINE: Editor cannot delete
  if (role === 'editor') {
    return { success: false, message: "UNAUTHORIZED: Editor tidak memiliki akses untuk menghapus data." };
  }
  
  var sheet = SHEET.getSheetByName(sheetName);
  var idField = sheetName === 'PRICING' ? 'package_id' : 'id';
  var rowIndex = findRowIndexById(sheetName, idField, recordId);
  
  if (rowIndex === -1) return { success: false, message: "Data tidak ditemukan." };
  
  sheet.deleteRow(rowIndex);
  return { success: true, message: "Data berhasil dihapus." };
}

function createLead(payload) {
  if (!payload.name || !payload.whatsapp || !payload.service) {
    return { success: false, message: "Field wajib belum lengkap." };
  }

  var sheet = SHEET.getSheetByName('LEADS');
  if (!sheet) return { success: false, message: "Sheet LEADS tidak ditemukan." };
  
  var headers = sheet.getDataRange().getValues()[0];
  var newRow = new Array(headers.length).fill("");
  
  var leadId = generateNextId();
  
  // PATCH: Force WhatsApp number to be string explicitly
  const whatsapp = "'" + String(payload.whatsapp || "").trim();
  
  var leadData = {
    id: leadId,
    created_at: new Date().toISOString(),
    name: payload.name,
    whatsapp: whatsapp, // Using the forced string variable
    service: payload.service,
    notes: payload.notes || "",
    source: "Landing Page",
    status: "NEW"
  };
  
  // Dynamic mapping allows columns to be rearranged without breaking the code
  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    if (leadData[header] !== undefined) {
      newRow[i] = leadData[header];
    }
  }
  
  sheet.appendRow(newRow);
  return { success: true, message: "Lead berhasil dikirim", lead_id: leadData.id };
}

/**
 * ==========================================
 * HELPER FUNCTIONS
 * ==========================================
 */
function getSheetDataAsObjects(sheetName) {
  var sheet = SHEET.getSheetByName(sheetName);
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var headers = data[0];
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      if (headers[j]) obj[headers[j]] = row[j];
    }
    result.push(obj);
  }
  
  return result;
}

function findRowIndexById(sheetName, idColumnName, searchId) {
  var sheet = SHEET.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idIndex = headers.indexOf(idColumnName);
  
  if (idIndex === -1) return -1;
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][idIndex].toString() === searchId.toString()) {
      return i + 1; // +1 because sheet rows are 1-indexed
    }
  }
  return -1;
}

function generateNextId() {
  return 'ID-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
}

function hashPassword(password) {
  var signature = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return signature.map(function(byte) {
      var v = (byte < 0) ? 256 + byte : byte;
      return ("0" + v.toString(16)).slice(-2);
  }).join("").toLowerCase();
}

function isTrue(value) {
  if (value === true || value === 1) return true;
  if (typeof value === 'string') {
    var lowerVal = value.toLowerCase().trim();
    return lowerVal === 'true' || lowerVal === 'yes' || lowerVal === '1';
  }
  return false;
}

function sortData(dataArray) {
  return dataArray.sort(function(a, b) {
    var orderA = parseFloat(a.sort_order);
    var orderB = parseFloat(b.sort_order);
    if (isNaN(orderA)) orderA = 99999;
    if (isNaN(orderB)) orderB = 99999;
    return orderA - orderB;
  });
}

function formatResponse(success, action, data) {
  var timestamp = new Date().toISOString();
  var responseObj = {
    success: success,
    action: action,
    generated_at: timestamp
  };
  
  if (success) {
    responseObj.data = data;
  } else {
    responseObj.message = data && data.message ? data.message : "An error occurred.";
    if (data && data.stack) responseObj.stack = data.stack;
  }
  
  return ContentService.createTextOutput(JSON.stringify(responseObj))
    .setMimeType(ContentService.MimeType.JSON);
}
