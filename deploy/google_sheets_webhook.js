/**
 * UHBFT Pharmacy Near Miss Reporting - Google Sheets Webhook
 * 
 * SETUP INSTRUCTIONS (Takes ~60 seconds):
 * 1. Open Google Sheets (https://sheets.new) and name it: "Pharmacy Near Miss Master Log"
 * 2. Click: Extensions > Apps Script
 * 3. Delete any code in the editor and paste this entire script.
 * 4. Click the blue "Deploy" button (top right) > "New deployment"
 * 5. Click the gear icon next to "Select type" > "Web app"
 * 6. Set Description: "Near Miss Webhook"
 * 7. Set "Execute as": "Me"
 * 8. Set "Who has access": "Anyone"  <-- CRITICAL for phones to post without Google login prompt
 * 9. Click "Deploy", authorize permissions when prompted, and copy the "Web app URL"
 * 10. Paste this URL into Near Miss Logger (via the Cloud Sync settings or nearmiss_config.js)
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // Wait up to 15s to prevent concurrent write collisions
  } catch (f) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Lock timeout" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Auto-initialize headers on first run if sheet is blank
    if (sheet.getLastRow() === 0) {
      var headers = [
        "Event ID", "Submitted At", "Site", "Site Other", "Type", "Satellite?", "Satellite Name",
        "Completed By", "Completed On", "Time Of Day", "Error Categories",
        "Error Drug Name", "Error CMM Code", "Error VTM ID", "Error Molecule",
        "Error VMP ID", "Error VMP Name", "Error AMP ID", "Error AMPP ID",
        "Error GTIN-14", "Error Batch", "Error Expiry", "Expired?", "NTI Drug?",
        "Intended Drug Name", "Intended CMM Code", "Intended VTM ID", "Intended VMP ID",
        "Intended AMP ID", "Intended NTI?", "Qty Dispensed", "Qty Prescribed",
        "Contributing Factors", "Learning Note (Zero PID)"
      ];
      sheet.appendRow(headers);
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground("#005eb8"); // NHS Blue
      headerRange.setFontColor("#ffffff");
      headerRange.setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    var payload = {};
    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else {
      payload = e.parameter || {};
    }

    // Extract fields matching Near Miss canonical schema
    var err = payload.errorDrug || {};
    var intd = payload.intendedDrug || {};
    var cats = Array.isArray(payload.errorCategories) ? payload.errorCategories.join("; ") : (payload.errorCategories || "");
    var facts = Array.isArray(payload.contributingFactors) ? payload.contributingFactors.join("; ") : (payload.contributingFactors || "");

    var row = [
      payload.eventId || ("NM-" + new Date().getTime()),
      payload.submittedAt || new Date().toISOString(),
      payload.site || "",
      payload.siteOther || "",
      payload.nearMissType || "Dispensing",
      payload.isSatelliteUnit ? "Yes" : "No",
      payload.satelliteUnitName || "",
      payload.completedBy || "",
      payload.completedOn || "",
      payload.timeOfDay || "",
      cats,
      err.name || "",
      err.cmmCode || "",
      err.vtmId || "",
      err.vtmName || "",
      err.vmpId || "",
      err.vmpName || "",
      err.ampId || "",
      err.amppId || "",
      err.scannedGtin14 || "",
      err.batchNumber || "",
      err.expiryDate || "",
      err.isExpired ? "Yes" : "No",
      err.isNti ? "Yes" : "No",
      intd.name || "",
      intd.cmmCode || "",
      intd.vtmId || "",
      intd.vmpId || "",
      intd.ampId || "",
      intd.isNti ? "Yes" : "No",
      payload.qtyDispensed || "",
      payload.qtyPrescribed || "",
      facts,
      payload.learningNote || ""
    ];

    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      eventId: payload.eventId,
      rowNumber: sheet.getLastRow()
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "online",
    service: "UHBFT Near Miss Google Sheets Webhook",
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}
