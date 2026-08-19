// ═══════════════════════════════════════════════════════════
// KAZAM RWA ONBOARDING — Google Apps Script Backend v3
// ═══════════════════════════════════════════════════════════
//
// WHAT'S NEW IN V3:
// - FIX: handleSubmission built its row as a fixed 34-value array whose last
//   two entries (society_fee_11kw, agreement_tenure) landed in the sheet's
//   latitude / longitude columns. Every new society since KZ-RWA-185 wrote
//   junk into those two columns. The row is now built BY HEADER NAME, so it
//   can never drift again no matter how the sheet is reordered or extended.
// - NEW: flat_count, ev_2w_count, ev_4w_count captured on submit and update
// - society_fee_11kw and agreement_tenure now actually land in real columns
//
// DEPLOY ORDER (important):
//   1. Paste this file over the old Code.gs, then re-deploy the web app.
//      This step alone stops the latitude/longitude corruption. No config to
//      change — SPREADSHEET_ID / PARENT_FOLDER_ID / NOTIFY_EMAIL are yours.
//   2. Run addMissingColumns() from the editor  — adds the 5 sheet columns.
//   3. Run previewCoordCleanup(), read the log, then runCoordCleanup().
//   4. Later, when the CSV is ready: set BULK_CSV_FILE_ID, then
//      previewBulkImport() and runBulkImport().
//
// Steps 2-4 live in the ONE-TIME MAINTENANCE section at the bottom of this
// file. Nothing in the web app calls them, so once they are done you can
// select from that banner to the end of the file and delete it.
//
// SETUP (unchanged from v2):
// 1. A "CPO_List" tab with column A header "cpo_name" and rows Vida, Kazam
// 2. Paste this code in Extensions → Apps Script
// 3. Deploy → Web app → Execute as Me → Anyone → Deploy
// ═══════════════════════════════════════════════════════════

// ---- CONFIG ----
const SPREADSHEET_ID = "1fmNAhg0Dkd3fvTj-PK9S221kEFcl9IsGbgAJIoZY74k";
const PARENT_FOLDER_ID = "14Nylm59BeS9GZkhBrh6nUixffcbbMWAJ";
const NOTIFY_EMAIL = "rachit@kazam.in";

// ---- CORS HANDLER ----
function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}

// ---- MAIN POST HANDLER ----
function logRequest(source, content) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let logSheet = ss.getSheetByName("Debug_Log");
  if (!logSheet) {
    logSheet = ss.insertSheet("Debug_Log");
    logSheet.appendRow(["timestamp", "source", "content"]);
  }
  logSheet.appendRow([new Date().toISOString(), source, String(content).substring(0, 500)]);
}


function doPost(e) {
 try {
    logRequest("doPost", e.postData.contents);
    let raw = e.postData.contents;
    if (raw.startsWith("data=")) {
      raw = decodeURIComponent(raw.substring(5).replace(/\+/g, ' '));
    } else if (raw.startsWith("payload=")) {
      raw = decodeURIComponent(raw.substring(8).replace(/\+/g, ' '));
    }
    const data = JSON.parse(raw);
    const action = data.action || "submit";
    let result;

    switch (action) {
      case "submit":
        result = handleSubmission(data);
        break;
      case "update":
        result = handleUpdate(data);
        break;
      case "addCPO":
        result = handleAddCPO(data);
        break;
      default:
        result = { success: false, error: "Unknown action: " + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- GET HANDLER ----
function doGet(e) {
    logRequest("doGet", JSON.stringify(e.parameter));

  const action = (e && e.parameter && e.parameter.action) || "status";
  let result;

  switch (action) {
    case "getNextId":
      result = { success: true, nextId: getNextSocietyId() };
      break;
    case "getSocieties":
      result = { success: true, societies: getSocietyList() };
      break;
    case "getSociety":
      result = { success: true, society: getSocietyById(e.parameter.id) };
      break;
    case "getCPOs":
      result = { success: true, cpos: getCPOList() };
      break;
    default:
      result = {
        success: true,
        status: "Kazam Onboarding API v3 is running",
        nextId: getNextSocietyId(),
        cpos: getCPOList()
      };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════
// CPO LIST MANAGEMENT
// ═══════════════════════════════════════════════════════════

function getCPOList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("CPO_List");

  // Auto-create CPO_List sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet("CPO_List");
    sheet.getRange("A1").setValue("cpo_name");
    sheet.getRange("A2").setValue("Vida");
    sheet.getRange("A3").setValue("Kazam");
  }

  const data = sheet.getDataRange().getValues();
  const cpos = [];
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0]).trim();
    if (name && name !== "") cpos.push(name);
  }
  return cpos;
}

function handleAddCPO(data) {
  const cpoName = (data.cpo_name || "").trim();
  if (!cpoName) return { success: false, error: "CPO name is required" };

  const existing = getCPOList();
  if (existing.map(c => c.toLowerCase()).includes(cpoName.toLowerCase())) {
    return { success: false, error: "CPO already exists" };
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("CPO_List");
  if (!sheet) {
    sheet = ss.insertSheet("CPO_List");
    sheet.getRange("A1").setValue("cpo_name");
  }
  sheet.appendRow([cpoName]);

  return { success: true, cpos: getCPOList() };
}

// ═══════════════════════════════════════════════════════════
// SOCIETY LOOKUP
// ═══════════════════════════════════════════════════════════

function getNextSocietyId() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Society_Master");
  const data = sheet.getDataRange().getValues();
  let maxNum = 0;
  for (let i = 1; i < data.length; i++) {
    const match = String(data[i][0]).trim().match(/KZ-RWA-(\d+)/);
    if (match) maxNum = Math.max(maxNum, parseInt(match[1]));
  }
  return "KZ-RWA-" + String(maxNum + 1).padStart(3, "0");
}

function getSocietyList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Society_Master");
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const societies = [];
  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach((h, j) => { row[h] = data[i][j]; });
    societies.push({
      id: String(row.society_id || "").trim(),
      name: String(row.society_name || "").trim(),
      city: String(row.city || "").trim(),
      status: String(row.status || "Active").trim(),
    });
  }
  return societies;
}

function getSocietyById(societyId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Society_Master");
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === societyId) {
      const row = {};
      headers.forEach((h, j) => {
        let val = data[i][j];
        if (val === null || val === undefined) val = "";
        row[h] = val;
      });
      row._rowIndex = i + 1; // 1-based row number for updates
      return row;
    }
  }
  return null;
}

// ═══════════════════════════════════════════════════════════
// SHARED HELPERS FOR COUNT FIELDS
// ═══════════════════════════════════════════════════════════

// Whole number >= 0, or "" when not supplied. Blank must stay blank rather
// than becoming 0 — for the count fields "not collected yet" and "genuinely
// zero" are different answers, and the bulk CSV import relies on the difference.
// Also swallows the placeholder text humans put in spreadsheets.
function countOrBlank(v) {
  const raw = String(v === null || v === undefined ? "" : v).trim().replace(/,/g, "");
  if (raw === "") return "";
  if (/^(na|n\/a|nil|none|-|nan|unknown|tbd)$/i.test(raw)) return "";
  const n = Number(raw);
  if (!isFinite(n) || n < 0) return "";
  return Math.round(n);
}

// Every spreadsheet call in this file goes through openById — the script is
// standalone, not bound to the sheet, so getActiveSpreadsheet() returns null.
function masterSheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Society_Master");
  if (!sheet) throw new Error('No "Society_Master" tab in the spreadsheet.');
  return sheet;
}

function masterHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn())
              .getValues()[0]
              .map(h => String(h).trim());
}

// ═══════════════════════════════════════════════════════════
// NEW SOCIETY SUBMISSION
// ═══════════════════════════════════════════════════════════

function handleSubmission(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const societyId = getNextSocietyId();

  // Create Drive folder
  const folderName = societyId + " - " + data.society_name;
  const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
  const societyFolder = parentFolder.createFolder(folderName);
  societyFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Upload files
  let agreementLink = "";
  let electricityBillLink = "";

  if (data.agreement_file) {
    const f = uploadFile(societyFolder, "Agreement - " + data.society_name + ".pdf", data.agreement_file);
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    agreementLink = f.getUrl();
  }
  if (data.electricity_bill_file) {
    const f = uploadFile(societyFolder, "Electricity Bill - " + data.society_name + ".pdf", data.electricity_bill_file);
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    electricityBillLink = f.getUrl();
  }

  // Build master row
  const totalChargers = (parseInt(data.chargers_3_3kw) || 0) + (parseInt(data.chargers_7_4kw) || 0) + (parseInt(data.chargers_11kw) || 0);
  const primaryCpo = data.cpo_3_3kw || data.cpo_7_4kw || data.cpo_11kw || "Vida";
  const primaryFee = parseFloat(data.society_fee_3_3kw) || parseFloat(data.society_fee_7_4kw) || parseFloat(data.society_fee_11kw) || 0;

  const masterSheet = ss.getSheetByName("Society_Master");

  // Values keyed BY COLUMN NAME. Previously this was a positional array, which
  // silently wrote society_fee_11kw into `latitude` and agreement_tenure into
  // `longitude` once those two columns were added to the sheet. Keying by name
  // means adding, removing or reordering sheet columns can never misalign it.
  const values = {
    society_id:             societyId,
    society_name:           data.society_name || "",
    cpo_name:               primaryCpo,                                  // backward compat
    city:                   data.city || "",
    state:                  data.state || "",
    status:                 "Active",
    electricity_rate:       parseFloat(data.electricity_rate) || 0,
    electricity_duty:       data.electricity_duty || "Inclusive",
    finalized_rate:         parseFloat(data.finalized_rate) || 0,
    society_fee:            primaryFee,                                  // backward compat
    no_of_chargers:         totalChargers,                               // backward compat
    ac_holder:              data.ac_holder || "",
    bank_name:              data.bank_name || "",
    ifsc:                   data.ifsc || "",
    ac_number:              data.ac_number || "",
    poc_name:               data.poc_name || "",
    poc_phone:              data.poc_phone || "",
    rwa_email:              data.rwa_email || "",
    address:                data.address || "",
    pincode:                data.pincode || "",
    agreement_date:         data.agreement_date || "",
    agreement_link:         agreementLink,
    google_maps_link:       data.google_maps_link || "",
    electricity_bill_link:  electricityBillLink,
    chargers_3_3kw:         parseInt(data.chargers_3_3kw) || 0,
    cpo_3_3kw:              data.cpo_3_3kw || "",
    society_fee_3_3kw:      parseFloat(data.society_fee_3_3kw) || 0,
    chargers_7_4kw:         parseInt(data.chargers_7_4kw) || 0,
    cpo_7_4kw:              data.cpo_7_4kw || "",
    society_fee_7_4kw:      parseFloat(data.society_fee_7_4kw) || 0,
    chargers_11kw:          parseInt(data.chargers_11kw) || 0,
    cpo_11kw:               data.cpo_11kw || "",
    society_fee_11kw:       parseFloat(data.society_fee_11kw) || 0,
    agreement_tenure:       data.agreement_tenure || "",
    flat_count:             countOrBlank(data.flat_count),
    ev_2w_count:            countOrBlank(data.ev_2w_count),
    ev_4w_count:            countOrBlank(data.ev_4w_count),
    // latitude / longitude deliberately absent — geocoding owns those columns,
    // and anything not listed here is written as blank rather than clobbered.
  };

  const headers = masterHeaders_(masterSheet);

  const masterRow = headers.map(h => (values.hasOwnProperty(h) ? values[h] : ""));

  // Surface anything the form collects that the sheet has no column for,
  // instead of dropping it silently the way the old code did.
  const orphaned = Object.keys(values).filter(k => headers.indexOf(k) === -1);
  if (orphaned.length) {
    logRequest("handleSubmission", "No sheet column for: " + orphaned.join(", ") +
                                   " — run addMissingColumns()");
  }

  masterSheet.appendRow(masterRow);

  // Add to Rate_History
  const rateSheet = ss.getSheetByName("Rate_History");
  if (rateSheet) {
    rateSheet.appendRow([
      societyId, data.society_name, data.agreement_date || "", "",
      parseFloat(data.electricity_rate) || 0, data.electricity_duty || "Inclusive",
      parseFloat(data.finalized_rate) || 0, "Initial agreement"
    ]);
  }

  // User_Access
  const userSheet = ss.getSheetByName("User_Access");
  const emails = (data.rwa_emails || data.rwa_email || "")
    .split(",").map(e => e.trim().toLowerCase()).filter(e => e && e.includes("@"));
  for (const email of emails) {
    userSheet.appendRow([email, societyId, data.society_name]);
  }

  // Notification
  sendNotification(societyId, data, agreementLink, electricityBillLink, societyFolder.getUrl(), emails, "NEW");

  return {
    success: true,
    societyId: societyId,
    societyName: data.society_name,
    folderUrl: societyFolder.getUrl(),
    emailsAdded: emails.length,
    orphanedFields: orphaned
  };
}

// ═══════════════════════════════════════════════════════════
// UPDATE EXISTING SOCIETY
// ═══════════════════════════════════════════════════════════

function handleUpdate(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const masterSheet = ss.getSheetByName("Society_Master");
  const allData = masterSheet.getDataRange().getValues();
  const headers = allData[0].map(h => String(h).trim());
  const societyId = (data.society_id || "").trim();

  // Find the row
  let rowIndex = -1;
  for (let i = 1; i < allData.length; i++) {
    if (String(allData[i][0]).trim() === societyId) {
      rowIndex = i + 1; // 1-based for Sheets API
      break;
    }
  }
  if (rowIndex === -1) return { success: false, error: "Society not found: " + societyId };

  // Map of field name → column index
  const colMap = {};
  headers.forEach((h, j) => { colMap[h] = j + 1; }); // 1-based

  // Fields that can be updated (key = form field, value = sheet column name)
  const updatableFields = {
    society_name: "society_name",
    city: "city",
    state: "state",
    status: "status",
    electricity_rate: "electricity_rate",
    electricity_duty: "electricity_duty",
    finalized_rate: "finalized_rate",
    ac_holder: "ac_holder",
    bank_name: "bank_name",
    ifsc: "ifsc",
    ac_number: "ac_number",
    poc_name: "poc_name",
    poc_phone: "poc_phone",
    rwa_email: "rwa_email",
    address: "address",
    pincode: "pincode",
    agreement_date: "agreement_date",
    agreement_tenure: "agreement_tenure",
    google_maps_link: "google_maps_link",
    chargers_3_3kw: "chargers_3_3kw",
    cpo_3_3kw: "cpo_3_3kw",
    society_fee_3_3kw: "society_fee_3_3kw",
    chargers_7_4kw: "chargers_7_4kw",
    cpo_7_4kw: "cpo_7_4kw",
    society_fee_7_4kw: "society_fee_7_4kw",
    chargers_11kw: "chargers_11kw",
    cpo_11kw: "cpo_11kw",
    society_fee_11kw: "society_fee_11kw",
    flat_count: "flat_count",
    ev_2w_count: "ev_2w_count",
    ev_4w_count: "ev_4w_count",
  };

  // Count fields are left alone when the form sends a blank, so editing an
  // older society without the counts to hand cannot wipe a value that is
  // already there (or one the bulk CSV import has since filled in).
  const countFields = { flat_count: 1, ev_2w_count: 1, ev_4w_count: 1 };

  // Track what changed
  const changes = [];

  for (const [formField, colName] of Object.entries(updatableFields)) {
    if (data[formField] !== undefined && data[formField] !== null) {
      const col = colMap[colName];
      if (col) {
        let incoming = data[formField];
        if (countFields[formField]) {
          incoming = countOrBlank(incoming);
          if (incoming === "") continue;   // blank means "no answer", not "erase"
        }
        const oldVal = String(allData[rowIndex - 1][col - 1] || "");
        const newVal = String(incoming);
        if (oldVal.trim() !== newVal.trim()) {
          masterSheet.getRange(rowIndex, col).setValue(incoming);
          changes.push(colName + ": " + oldVal + " → " + newVal);
        }
      }
    }
  }

  // Handle file uploads for updates
  let agreementLink = "";
  let electricityBillLink = "";

  if (data.agreement_file) {
    // Find or create society folder
    const folder = getOrCreateSocietyFolder(societyId, data.society_name || "");
    const f = uploadFile(folder, "Agreement - " + (data.society_name || societyId) + ".pdf", data.agreement_file);
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    agreementLink = f.getUrl();
    const col = colMap["agreement_link"];
    if (col) {
      masterSheet.getRange(rowIndex, col).setValue(agreementLink);
      changes.push("agreement_link: uploaded new file");
    }
  }

  if (data.electricity_bill_file) {
    const folder = getOrCreateSocietyFolder(societyId, data.society_name || "");
    const f = uploadFile(folder, "Electricity Bill - " + (data.society_name || societyId) + ".pdf", data.electricity_bill_file);
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    electricityBillLink = f.getUrl();
    const col = colMap["electricity_bill_link"];
    if (col) {
      masterSheet.getRange(rowIndex, col).setValue(electricityBillLink);
      changes.push("electricity_bill_link: uploaded new file");
    }
  }

  // Update backward-compat columns
  const totalChargers = (parseInt(data.chargers_3_3kw) || 0) + (parseInt(data.chargers_7_4kw) || 0) + (parseInt(data.chargers_11kw) || 0);
  if (totalChargers > 0 && colMap["no_of_chargers"]) {
    masterSheet.getRange(rowIndex, colMap["no_of_chargers"]).setValue(totalChargers);
  }
  const primaryCpo = data.cpo_3_3kw || data.cpo_7_4kw || data.cpo_11kw;
  if (primaryCpo && colMap["cpo_name"]) {
    masterSheet.getRange(rowIndex, colMap["cpo_name"]).setValue(primaryCpo);
  }
  const primaryFee = parseFloat(data.society_fee_3_3kw) || parseFloat(data.society_fee_7_4kw) || parseFloat(data.society_fee_11kw) || 0;
  if (colMap["society_fee"]) {
    masterSheet.getRange(rowIndex, colMap["society_fee"]).setValue(primaryFee);
  }

  // Update User_Access if new emails provided
  if (data.rwa_emails) {
    const userSheet = ss.getSheetByName("User_Access");
    const newEmails = data.rwa_emails.split(",").map(e => e.trim().toLowerCase()).filter(e => e && e.includes("@"));
    // Get existing emails for this society
    const userData = userSheet.getDataRange().getValues();
    const existingEmails = new Set();
    for (let i = 1; i < userData.length; i++) {
      if (String(userData[i][1]).trim() === societyId) {
        existingEmails.add(String(userData[i][0]).trim().toLowerCase());
      }
    }
    // Add only new emails
    for (const email of newEmails) {
      if (!existingEmails.has(email)) {
        userSheet.appendRow([email, societyId, data.society_name || ""]);
        changes.push("User_Access: added " + email);
      }
    }
  }

  // Notification
  if (changes.length > 0 && NOTIFY_EMAIL) {
    try {
      const body = [
        "Society updated: " + societyId + " (" + (data.society_name || "") + ")",
        "",
        "Changes:",
        ...changes.map(c => "  • " + c),
        "",
        "Updated at: " + new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" }),
      ].join("\n");
      MailApp.sendEmail(NOTIFY_EMAIL, "Society Updated: " + (data.society_name || societyId), body);
    } catch (e) { /* non-critical */ }
  }

  return {
    success: true,
    societyId: societyId,
    changes: changes,
    message: changes.length > 0 ? changes.length + " field(s) updated" : "No changes detected"
  };
}

// ═══════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════

function getOrCreateSocietyFolder(societyId, societyName) {
  const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
  const folderName = societyId + " - " + societyName;
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  // Try partial match
  const allFolders = parentFolder.getFolders();
  while (allFolders.hasNext()) {
    const f = allFolders.next();
    if (f.getName().startsWith(societyId)) return f;
  }
  // Create new
  const newFolder = parentFolder.createFolder(folderName);
  newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return newFolder;
}

function uploadFile(folder, fileName, base64Data) {
  const parts = base64Data.split(",");
  const mimeMatch = parts[0].match(/data:(.*?);/);
  const mimeType = mimeMatch ? mimeMatch[1] : "application/pdf";
  const decoded = Utilities.base64Decode(parts[1]);
  const blob = Utilities.newBlob(decoded, mimeType, fileName);
  return folder.createFile(blob);
}

function sendNotification(societyId, data, agLink, ebLink, folderUrl, emails, type) {
  if (!NOTIFY_EMAIL) return;
  try {
    const prefix = type === "NEW" ? "New Society Onboarded" : "Society Updated";
    const subject = prefix + ": " + data.society_name + " (" + societyId + ")";
    const body = [
      prefix + " on the Kazam RWA Dashboard.",
      "",
      "Society ID: " + societyId,
      "Society Name: " + data.society_name,
      "City: " + data.city + ", " + data.state,
      "",
      "Society size:",
      "  Flats: " + (data.flat_count || "—"),
      "  Existing EVs: " + (data.ev_2w_count || 0) + " x 2W, " + (data.ev_4w_count || 0) + " x 4W",
      "",
      "Chargers:",
      "  3.3 kW: " + (data.chargers_3_3kw || 0) + " (CPO: " + (data.cpo_3_3kw || "—") + ")",
      "  7.4 kW: " + (data.chargers_7_4kw || 0) + " (CPO: " + (data.cpo_7_4kw || "—") + ")",
      "  11 kW: " + (data.chargers_11kw || 0) + " (CPO: " + (data.cpo_11kw || "—") + ")",
      "",
      "Rate: ₹" + data.finalized_rate + "/kWh",
      "POC: " + (data.poc_name || "") + " (" + (data.poc_phone || "") + ")",
      "Emails: " + emails.join(", "),
      "",
      "Agreement: " + (agLink || "Not uploaded"),
      "Drive Folder: " + (folderUrl || ""),
      "",
      "Submitted at: " + new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" }),
    ].join("\n");
    MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
  } catch (e) { /* non-critical */ }
}

function testSubmit() {
  const testData = {
    action: "submit",
    society_name: "TEST DELETE ME",
    cpo_3_3kw: "Vida",
    city: "Bengaluru",
    state: "Karnataka",
    electricity_rate: "6.92",
    electricity_duty: "Inclusive",
    finalized_rate: "6.92",
    chargers_3_3kw: "2",
    chargers_7_4kw: "0",
    chargers_11kw: "0",
    society_fee_3_3kw: "0",
    society_fee_7_4kw: "0",
    society_fee_11kw: "0",
    ac_holder: "Test Society",
    bank_name: "HDFC",
    ifsc: "HDFC0001234",
    ac_number: "12345678901234",
    poc_name: "Test User",
    poc_phone: "9999999999",
    rwa_emails: "test@kazam.in",
    address: "Test Address",
    pincode: "560001",
    agreement_date: "2026-04-27",
    agreement_tenure: "3 years",
    google_maps_link: "",
    flat_count: "240",
    ev_2w_count: "18",
    ev_4w_count: "6",
  };

  const result = handleSubmission(testData);
  Logger.log(JSON.stringify(result));
}

function authorizeDrive() {
  const folders = DriveApp.getRootFolder();
  Logger.log("Drive access OK: " + folders.getName());
}


// ═══════════════════════════════════════════════════════════
// ONE-TIME MAINTENANCE — delete this whole section once it is done
// ═══════════════════════════════════════════════════════════
// Nothing below is reachable from the web app. doPost and doGet never call
// any of it. Run these by hand from the Apps Script editor: pick the function
// from the dropdown, press Run, then read Executions / the log.
//
//   1. addMissingColumns()                          — adds the 5 sheet columns
//   2. previewCoordCleanup()  then runCoordCleanup()  — clears junk lat/long
//   3. previewBulkImport()    then runBulkImport()    — loads the counts CSV
//
// Every step has a preview that writes nothing. Read it before the real run.
// Once the sheet is fixed and the CSV is loaded, select from this banner to
// the end of the file and delete it — the web app is unaffected.

// The only value you need to set, and only for step 3.
// From https://drive.google.com/file/d/<THIS_PART>/view
const BULK_CSV_FILE_ID = "";

const COUNT_COLUMNS   = ["flat_count", "ev_2w_count", "ev_4w_count"];
const MISSING_COLUMNS = COUNT_COLUMNS.concat(["society_fee_11kw", "agreement_tenure"]);


// ── 1. ADD THE MISSING COLUMNS (idempotent) ─────────────────
// Appends to the right of the last column rather than inserting each name
// beside its logical neighbour. Both read and write paths are header-driven
// now, so position genuinely does not matter.
function addMissingColumns() {
  const sheet   = masterSheet_();
  const headers = masterHeaders_(sheet);
  const added   = [];

  MISSING_COLUMNS.forEach(name => {
    if (headers.indexOf(name) !== -1) return;       // already there
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue(name);
    headers.push(name);
    added.push(name);
  });

  if (added.length) {
    // Copy the existing header formatting across (format only, so the names
    // written above survive).
    const firstNew = sheet.getLastColumn() - added.length + 1;
    sheet.getRange(1, 1).copyTo(
      sheet.getRange(1, firstNew, 1, added.length),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );
  }

  const msg = added.length
    ? "Added " + added.length + " column(s): " + added.join(", ")
    : "Nothing to do — every column already exists.";
  Logger.log(msg);
  return msg;
}


// ── 2. CLEAR JUNK LEFT IN latitude / longitude ──────────────
// Until v3, handleSubmission appended a fixed 34-value array whose last two
// entries landed in latitude and longitude. Societies from KZ-RWA-185 onward
// therefore hold society_fee_11kw (a 0) in latitude and agreement_tenure
// (usually blank) in longitude. This blanks anything that is not a plausible
// coordinate so a geocoding pass can refill it. Real coordinates are kept.
const INDIA_LAT = [6.0, 38.0];
const INDIA_LON = [68.0, 98.0];

function previewCoordCleanup() { return coordCleanup_(true);  }
function runCoordCleanup()     { return coordCleanup_(false); }

function isPlausibleCoord_(lat, lon) {
  const la = Number(String(lat).trim());
  const lo = Number(String(lon).trim());
  if (!isFinite(la) || !isFinite(lo)) return false;
  return la >= INDIA_LAT[0] && la <= INDIA_LAT[1] &&
         lo >= INDIA_LON[0] && lo <= INDIA_LON[1];
}

function coordCleanup_(dryRun) {
  const sheet   = masterSheet_();
  const headers = masterHeaders_(sheet);
  const latCol  = headers.indexOf("latitude");
  const lonCol  = headers.indexOf("longitude");
  const idCol   = headers.indexOf("society_id");
  if (latCol === -1 || lonCol === -1) throw new Error("No latitude / longitude columns.");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("No data rows.");
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  const toClear = [];
  let kept = 0, alreadyBlank = 0;

  data.forEach((row, i) => {
    const lat = String(row[latCol] == null ? "" : row[latCol]).trim();
    const lon = String(row[lonCol] == null ? "" : row[lonCol]).trim();
    const id  = idCol === -1 ? "row " + (i + 2) : String(row[idCol]).trim();
    if (lat === "" && lon === "")        { alreadyBlank++; return; }
    if (isPlausibleCoord_(lat, lon))     { kept++;         return; }
    toClear.push({ rowNum: i + 2, id: id, lat: lat, lon: lon });
  });

  if (!dryRun && toClear.length) {
    toClear.forEach(t => {
      sheet.getRange(t.rowNum, latCol + 1).setValue("");
      sheet.getRange(t.rowNum, lonCol + 1).setValue("");
    });
    SpreadsheetApp.flush();
  }

  const lines = [
    dryRun ? "=== DRY RUN (nothing written) ===" : "=== CLEANUP COMPLETE ===",
    "Real coordinates kept : " + kept,
    "Already blank         : " + alreadyBlank,
    "Junk " + (dryRun ? "to clear" : "cleared") + "        : " + toClear.length,
    ""
  ];
  toClear.slice(0, 60).forEach(t => {
    lines.push("  " + t.id + "  lat=" + JSON.stringify(t.lat) + "  lon=" + JSON.stringify(t.lon));
  });
  if (toClear.length > 60) lines.push("  …and " + (toClear.length - 60) + " more");

  const out = lines.join("\n");
  Logger.log(out);
  return out;
}


// ── 3. BULK-LOAD COUNTS FOR OLDER RWAs ──────────────────────
// Manual CSVs never use the exact field names, so accept the common spellings.
const CSV_ALIASES = {
  society_id:   ["society_id", "id", "rwa_id", "kazam_id", "societyid"],
  society_name: ["society_name", "name", "society", "rwa_name", "societyname"],
  flat_count:   ["flat_count", "flats", "no_of_flats", "number_of_flats",
                 "total_flats", "flat", "units", "no_of_units"],
  ev_2w_count:  ["ev_2w_count", "2w", "ev_2w", "2w_ev", "two_wheeler",
                 "two_wheelers", "no_of_2w", "2w_count", "2_wheeler"],
  ev_4w_count:  ["ev_4w_count", "4w", "ev_4w", "4w_ev", "four_wheeler",
                 "four_wheelers", "no_of_4w", "4w_count", "4_wheeler"]
};

function normKey_(s)  { return String(s).toLowerCase().replace(/^﻿/, "").replace(/[^a-z0-9]/g, ""); }
function normName_(s) { return String(s).toLowerCase().replace(/[^a-z0-9]/g, ""); }

function mapCsvHeaders_(headerRow) {
  const map = {};
  headerRow.forEach((raw, idx) => {
    const k = normKey_(raw);
    Object.keys(CSV_ALIASES).forEach(field => {
      if (map[field] !== undefined) return;
      if (CSV_ALIASES[field].some(a => normKey_(a) === k)) map[field] = idx;
    });
  });
  return map;
}

function previewBulkImport() { return bulkImport_(true);  }
function runBulkImport()     { return bulkImport_(false); }

function bulkImport_(dryRun) {
  if (!BULK_CSV_FILE_ID) throw new Error("Set BULK_CSV_FILE_ID first.");

  const csv  = DriveApp.getFileById(BULK_CSV_FILE_ID).getBlob().getDataAsString("UTF-8");
  const rows = Utilities.parseCsv(csv);
  if (rows.length < 2) throw new Error("CSV has no data rows.");

  const cmap = mapCsvHeaders_(rows[0]);
  if (cmap.society_id === undefined && cmap.society_name === undefined) {
    throw new Error("CSV needs a society_id or society_name column. Found: " + rows[0].join(", "));
  }
  const present = COUNT_COLUMNS.filter(f => cmap[f] !== undefined);
  if (!present.length) {
    throw new Error("CSV has none of flat_count / ev_2w_count / ev_4w_count. Found: " + rows[0].join(", "));
  }

  const sheet   = masterSheet_();
  const headers = masterHeaders_(sheet);
  COUNT_COLUMNS.forEach(c => {
    if (headers.indexOf(c) === -1) throw new Error('Sheet is missing "' + c + '". Run addMissingColumns() first.');
  });

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("Master sheet has no data rows.");
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  const idCol   = headers.indexOf("society_id");
  const nameCol = headers.indexOf("society_name");

  // Names are only usable as a key when they are unambiguous.
  const byId = {}, byName = {}, nameDupes = {};
  data.forEach((row, i) => {
    if (idCol !== -1) {
      const id = String(row[idCol]).trim();
      if (id) byId[id] = i;
    }
    if (nameCol !== -1) {
      const nm = normName_(row[nameCol]);
      if (!nm) return;
      if (byName[nm] !== undefined) nameDupes[nm] = true; else byName[nm] = i;
    }
  });

  const colIdx = {};
  COUNT_COLUMNS.forEach(c => { colIdx[c] = headers.indexOf(c); });

  const updated = [], unmatched = [], ambiguous = [], noValues = [];
  let touchedCells = 0;

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row.join("").trim()) continue;                 // blank line

    const csvId   = cmap.society_id   !== undefined ? String(row[cmap.society_id]).trim()   : "";
    const csvName = cmap.society_name !== undefined ? String(row[cmap.society_name]).trim() : "";
    const label   = csvId || csvName || ("row " + (r + 1));

    let target = null;
    if (csvId && byId[csvId] !== undefined) {
      target = byId[csvId];
    } else if (csvName) {
      const nm = normName_(csvName);
      if (nameDupes[nm])            { ambiguous.push(label); continue; }
      if (byName[nm] !== undefined) { target = byName[nm]; }
    }
    if (target === null) { unmatched.push(label); continue; }

    const wrote = [];
    present.forEach(field => {
      const val = countOrBlank(row[cmap[field]]);
      if (val === "") return;                                                  // blank in CSV -> leave alone
      if (String(data[target][colIdx[field]]).trim() === String(val)) return;  // unchanged
      data[target][colIdx[field]] = val;
      wrote.push(field + "=" + val);
      touchedCells++;
    });

    if (wrote.length) updated.push(label + " → " + wrote.join(", "));
    else noValues.push(label);
  }

  if (!dryRun && touchedCells > 0) {
    COUNT_COLUMNS.forEach(c => {
      const ci = colIdx[c];
      sheet.getRange(2, ci + 1, data.length, 1).setValues(data.map(row => [row[ci]]));
    });
    SpreadsheetApp.flush();
  }

  const report = [
    dryRun ? "=== DRY RUN (nothing written) ===" : "=== IMPORT COMPLETE ===",
    "CSV columns used : " + present.join(", "),
    "Matched & changed: " + updated.length + " societies (" + touchedCells + " cells)",
    "Already correct  : " + noValues.length,
    "No match in sheet: " + unmatched.length,
    "Ambiguous name   : " + ambiguous.length,
    ""
  ];
  if (updated.length) {
    report.push("CHANGES:", updated.slice(0, 60).join("\n"));
    if (updated.length > 60) report.push("…and " + (updated.length - 60) + " more");
    report.push("");
  }
  if (unmatched.length) report.push("NO MATCH — fix these in the CSV:", unmatched.join(", "), "");
  if (ambiguous.length) report.push("AMBIGUOUS — duplicate society_name, add society_id:", ambiguous.join(", "), "");

  const out = report.join("\n");
  Logger.log(out);
  return out;
}
