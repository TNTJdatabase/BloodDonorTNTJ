/**
 * ============================================================
 * BloodHeroTNTJ — Blood Donor Registration System
 * Backend: Google Apps Script
 * ============================================================
 */

const CONFIG = {
  SHEET_NAME: "BloodDB",
  DONOR_PREFIX: "DNR",
  TIMEZONE: "Asia/Kolkata",
  DATE_FORMAT: "dd/MM/yyyy hh:mm:ss a",
  DATE_ONLY_FORMAT: "dd/MM/yyyy",
  MIN_AGE: 18,
  MAX_AGE: 65,
  DUPLICATE_WINDOW_SECONDS: 30,
  LOCK_TIMEOUT_MS: 30000,
  PHONE_REGEX: /^\+?[0-9]{10,15}$/
};

const BLOOD_GROUPS = [
  "A+", "A-", "B+", "B-", "AB+", "AB-", "O+", "O-",
  "A1+", "A1-", "A2+", "A2-", "A1B+", "A1B-", "A2B+", "A2B-",
  "dont-know"
];

const GENDERS = ["male", "female", "other"];

const DONATION_TYPES = {
  THREE_MONTHS: "3months",
  TWO_MONTHS: "2months",
  ONE_MONTH: "1month",
  LESS_THAN_ONE_MONTH: "less1month",
  NEVER: "never"
};

const CELL_COLOURS = {
  GREEN: "#9ef01a",
  ORANGE: "#ffba08",
  NONE: null
};

const ELIGIBILITY_RULES = {
  male: { greenDays: 90, orangeDays: 60 },
  female: { greenDays: 120, orangeDays: 90 },
  other: { greenDays: 90, orangeDays: 60 }
};

const COLUMNS = {
  TIMESTAMP: 1, DONOR_ID: 2, NAME: 3, AGE: 4, GENDER: 5,
  CONTACT: 6, BLOOD_GROUP: 7, ADDRESS: 8, AREA: 9, LAST_DONATED: 10
};

const HEADERS = [
  "Timestamp", "Donor ID", "Name", "Age", "Gender",
  "Contact", "Blood Group", "Address", "Area", "Last Donated"
];

const TOTAL_COLUMNS = Object.keys(COLUMNS).length;

// ============================================================
// ENTRY POINTS
// ============================================================

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse(errorResponse("No data received."));
    }
    const data = JSON.parse(e.postData.contents);
    return jsonResponse(registerDonor(data));
  } catch (err) {
    Logger.log("doPost error: " + err.stack);
    return jsonResponse(errorResponse("Server Error"));
  }
}

function doGet(e) {
  return jsonResponse({ success: true, message: "BloodHeroTNTJ backend is live." });
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("BloodHeroTNTJ")
    .addItem("Run duplicate cleanup & colors NOW", "processPendingUpdates")
    .addToUi();
}

// ============================================================
// CORE REGISTRATION LOGIC
// ============================================================

function registerDonor(rawData) {
  const data = sanitizeInput(rawData);
  const validation = validateInput(data);
  if (!validation.valid) return errorResponse(validation.message || "Validation Failed");

  const sheet = getOrCreateSheet();
  if (isRecentDuplicate(data.contact, data.blood_group)) {
    return errorResponse("Duplicate submission detected. Please wait before retrying.");
  }

  const lock = LockService.getScriptLock();
  let donorId;

  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    const timestamp = getCurrentTimestamp();
    const lastDonatedValue = calculateDonationDate(data.last_donation);
    donorId = generateNextDonorId(sheet);

    const rowNumber = writeRow(sheet, {
      timestamp: timestamp, donorId: donorId, name: data.name, age: data.age,
      gender: data.gender, contact: data.contact, bloodGroup: data.blood_group,
      address: data.address, area: data.area, lastDonated: lastDonatedValue
    });

    formatRowBasic(sheet, rowNumber, { contact: data.contact, lastDonated: lastDonatedValue });
    markRecentSubmission(data.contact, data.blood_group);

  } catch (err) {
    Logger.log("registerDonor error: " + err.stack);
    return errorResponse("Server Error");
  } finally {
    lock.releaseLock();
  }

  // Notice: No trigger scheduling here anymore! We let the dashboard handle it.
  return successResponse(donorId);
}

// ============================================================
// SANITIZATION & VALIDATION
// ============================================================

function sanitizeInput(raw) {
  raw = raw || {};
  return {
    name: toTitleCase(trimString(raw.name)),
    age: Number(trimString(raw.age)),
    gender: trimString(raw.gender).toLowerCase(),
    contact: trimString(raw.contact),
    blood_group: normalizeBloodGroup(trimString(raw.blood_group)),
    address: trimString(raw.address),
    area: trimString(raw.area),
    last_donation: trimString(raw.last_donation).toLowerCase()
  };
}

function trimString(value) { return (value == null) ? "" : String(value).trim(); }

function toTitleCase(str) {
  if (!str) return "";
  return str.toLowerCase().split(/\s+/).map(w => w.length ? w.charAt(0).toUpperCase() + w.slice(1) : "").join(" ");
}

function normalizeBloodGroup(value) {
  if (!value) return "";
  const upper = value.toUpperCase();
  for (let i = 0; i < BLOOD_GROUPS.length; i++) {
    if (BLOOD_GROUPS[i].toUpperCase() === upper) return BLOOD_GROUPS[i];
  }
  return value;
}

function validateInput(data) {
  if (!data.name) return { valid: false, message: "Name is required." };
  if (!Number.isFinite(data.age) || data.age < CONFIG.MIN_AGE || data.age > CONFIG.MAX_AGE) return { valid: false, message: `Age must be between ${CONFIG.MIN_AGE} and ${CONFIG.MAX_AGE}.` };
  if (!data.gender || GENDERS.indexOf(data.gender) === -1) return { valid: false, message: "Invalid gender." };
  if (!data.contact || !CONFIG.PHONE_REGEX.test(data.contact)) return { valid: false, message: "Invalid phone number." };
  if (!data.blood_group || BLOOD_GROUPS.indexOf(data.blood_group) === -1) return { valid: false, message: "Invalid blood group." };
  if (!data.address) return { valid: false, message: "Address is required." };
  if (!data.area) return { valid: false, message: "Area is required." };

  const validDonationTypes = Object.values(DONATION_TYPES);
  if (!data.last_donation || validDonationTypes.indexOf(data.last_donation) === -1) return { valid: false, message: "Invalid last donation date." };
  
  return { valid: true, message: "" };
}

// ============================================================
// DUPLICATE HANDLING
// ============================================================

function normalizePhoneForDuplicateKey(phone) {
  const digitsOnly = String(phone || "").replace(/\D/g, "");
  return digitsOnly || String(phone || "").trim();
}

function buildDuplicateKey(phone, bloodGroup) {
  const normalizedPhone = normalizePhoneForDuplicateKey(phone);
  const normalizedGroup = String(bloodGroup || "").trim().toUpperCase();
  return normalizedPhone + "|" + normalizedGroup;
}

function isRecentDuplicate(phone, bloodGroup) {
  return CacheService.getScriptCache().get("submit_" + buildDuplicateKey(phone, bloodGroup)) !== null;
}

function markRecentSubmission(phone, bloodGroup) {
  CacheService.getScriptCache().put("submit_" + buildDuplicateKey(phone, bloodGroup), "1", CONFIG.DUPLICATE_WINDOW_SECONDS);
}

function removeDuplicateDonors(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const numRows = lastRow - 1;
  const contactValues = sheet.getRange(2, COLUMNS.CONTACT, numRows, 1).getDisplayValues();
  const bloodGroupValues = sheet.getRange(2, COLUMNS.BLOOD_GROUP, numRows, 1).getDisplayValues();
  const rowsForKey = {};

  for (let i = 0; i < numRows; i++) {
    const rowNumber = i + 2;
    const key = buildDuplicateKey(contactValues[i][0], bloodGroupValues[i][0]);
    if (!rowsForKey[key]) rowsForKey[key] = [];
    rowsForKey[key].push(rowNumber);
  }

  const rowsToDelete = [];
  Object.keys(rowsForKey).forEach(key => {
    const rows = rowsForKey[key];
    if (rows.length > 1) {
      for (let j = 0; j < rows.length - 1; j++) rowsToDelete.push(rows[j]);
    }
  });

  rowsToDelete.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));
  return rowsToDelete.length;
}

// ============================================================
// ID GENERATION
// ============================================================

function getMaxDonorNumber(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const idValues = sheet.getRange(2, COLUMNS.DONOR_ID, lastRow - 1, 1).getDisplayValues();
  const prefixLen = CONFIG.DONOR_PREFIX.length;
  let max = 0;

  for (let i = 0; i < idValues.length; i++) {
    const idStr = idValues[i][0];
    if (idStr && idStr.indexOf(CONFIG.DONOR_PREFIX) === 0) {
      const num = parseInt(idStr.slice(prefixLen), 10);
      if (!isNaN(num) && num > max) max = num;
    }
  }
  return max;
}

function generateNextDonorId(sheet) {
  return CONFIG.DONOR_PREFIX + (getMaxDonorNumber(sheet) + 1);
}

function assignPendingDonorIds(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const numRows = lastRow - 1;
  const idRange = sheet.getRange(2, COLUMNS.DONOR_ID, numRows, 1);
  const idValues = idRange.getDisplayValues();
  let nextNumber = getMaxDonorNumber(sheet) + 1;
  let changed = false;
  const updatedValues = idValues.map(row => [row[0]]);

  for (let i = 0; i < numRows; i++) {
    if (!idValues[i][0]) {
      updatedValues[i][0] = CONFIG.DONOR_PREFIX + nextNumber;
      nextNumber++;
      changed = true;
    }
  }
  if (changed) idRange.setValues(updatedValues);
}

// ============================================================
// DATE / ELIGIBILITY LOGIC
// ============================================================

function calculateDonationDate(donationType) {
  const today = new Date();
  switch (donationType) {
    case DONATION_TYPES.THREE_MONTHS: return shiftMonths(today, -3);
    case DONATION_TYPES.TWO_MONTHS: return shiftMonths(today, -2);
    case DONATION_TYPES.ONE_MONTH: return shiftMonths(today, -1);
    case DONATION_TYPES.LESS_THAN_ONE_MONTH: return shiftMonths(today, 0);
    default: return 0;
  }
}

function shiftMonths(date, months) {
  const result = new Date(date.getTime());
  result.setMonth(result.getMonth() + months);
  return result;
}

function getEligibilityColour(gender, lastDonated) {
  if (!lastDonated) return CELL_COLOURS.GREEN;
  const rules = ELIGIBILITY_RULES[gender] || ELIGIBILITY_RULES.other;
  const msPerDay = 24 * 60 * 60 * 1000;
  const daysSince = Math.floor((new Date().getTime() - new Date(lastDonated).getTime()) / msPerDay);

  if (daysSince >= rules.greenDays) return CELL_COLOURS.GREEN;
  if (daysSince >= rules.orangeDays) return CELL_COLOURS.ORANGE;
  return CELL_COLOURS.NONE;
}

function recalculateAllEligibility(sheetArg) {
  const sheet = sheetArg || getOrCreateSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const numRows = lastRow - 1;
  const genderValues = sheet.getRange(2, COLUMNS.GENDER, numRows, 1).getDisplayValues();
  const lastDonatedValues = sheet.getRange(2, COLUMNS.LAST_DONATED, numRows, 1).getValues();
  const bloodGroupRange = sheet.getRange(2, COLUMNS.BLOOD_GROUP, numRows, 1);
  const backgrounds = [];

  for (let i = 0; i < numRows; i++) {
    backgrounds.push([getEligibilityColour((genderValues[i][0] || "").toLowerCase(), lastDonatedValues[i][0])]);
  }
  bloodGroupRange.setBackgrounds(backgrounds);
}

// ============================================================
// SHEET / ROW HELPERS
// ============================================================

function getCurrentTimestamp() { return new Date(); }

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.getRange(2, COLUMNS.TIMESTAMP, sheet.getMaxRows() - 1, 1).setNumberFormat(CONFIG.DATE_FORMAT);
    sheet.getRange(2, COLUMNS.LAST_DONATED, sheet.getMaxRows() - 1, 1).setNumberFormat(CONFIG.DATE_ONLY_FORMAT);
  }
  return sheet;
}

function writeRow(sheet, rowData) {
  const rowValues = new Array(TOTAL_COLUMNS);
  rowValues[COLUMNS.TIMESTAMP - 1] = rowData.timestamp;
  rowValues[COLUMNS.DONOR_ID - 1] = rowData.donorId;
  rowValues[COLUMNS.NAME - 1] = rowData.name;
  rowValues[COLUMNS.AGE - 1] = rowData.age;
  rowValues[COLUMNS.GENDER - 1] = rowData.gender;
  rowValues[COLUMNS.CONTACT - 1] = rowData.contact;
  rowValues[COLUMNS.BLOOD_GROUP - 1] = rowData.bloodGroup;
  rowValues[COLUMNS.ADDRESS - 1] = rowData.address;
  rowValues[COLUMNS.AREA - 1] = rowData.area;
  rowValues[COLUMNS.LAST_DONATED - 1] = rowData.lastDonated;
  sheet.appendRow(rowValues);
  return sheet.getLastRow();
}

function formatRowBasic(sheet, rowNumber, rowData) {
  sheet.getRange(rowNumber, COLUMNS.CONTACT).setFormula(`=HYPERLINK("tel:${rowData.contact}","${rowData.contact}")`);
  const lastDonatedCell = sheet.getRange(rowNumber, COLUMNS.LAST_DONATED);
  rowData.lastDonated === 0 ? lastDonatedCell.setValue("Never") : lastDonatedCell.setValue(rowData.lastDonated).setNumberFormat(CONFIG.DATE_ONLY_FORMAT);
}

// ============================================================
// DASHBOARD TARGET: The function you will select in the Triggers UI
// ============================================================

function processPendingUpdates() {
  try {
    const sheet = getOrCreateSheet();
    assignPendingDonorIds(sheet);
    removeDuplicateDonors(sheet);
    recalculateAllEligibility(sheet);
    SpreadsheetApp.flush();
  } catch (err) {
    Logger.log("Dashboard Trigger failed: " + err.stack);
  }
}

// ============================================================
// RESPONSE HELPERS
// ============================================================

function jsonResponse(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function errorResponse(msg) { return { success: false, message: msg }; }
function successResponse(id) { return { success: true, donorId: id, message: "Registration successful." }; }