/**
 * श्री फार्मसी — Google Sheets Sync Backend (v2)
 * ---------------------------------------------------
 * Deploy > New deployment > Web app असे करून URL मिळवा
 * आणि तो URL index.html मधील SHEET_API_URL मध्ये टाका.
 *
 * शीट्स (आपोआप तयार होतील): Medicines, Vendors, Sales, Purchases,
 * Khata, Returns, Settings
 */

const SHEET_NAMES = ["Medicines", "Vendors", "Sales", "Purchases", "Khata", "Returns", "Settings", "Users"];
const BACKUP_FOLDER_NAME = "Shree Pharmacy Backups";

/**
 * हे फंक्शन Apps Script एडिटरमधून मॅन्युअली एकदा चालवा (वरती फंक्शन ड्रॉपडाऊनमध्ये
 * "setupSheets" निवडून ▶ Run दाबा) — यामुळे सर्व 7 शीट्स लगेच हेडरसह तयार होतील,
 * ॲपमधून काहीही सिंक होण्याची वाट न बघता.
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = {
    Medicines: ["id","name","salt","mfr","batch","expiry","mrp","cost","stock","reorder","barcode","stripSize"],
    Vendors: ["id","name","phone","gst","outstanding","payments"],
    Sales: ["id","date","channel","customer","items","total"],
    Purchases: ["id","date","billNo","medicine","medId","vendor","vendorId","qty","cost"],
    Khata: ["id","name","phone","balance","payments"],
    Returns: ["id","date","medicine","medId","batch","vendor","vendorId","qty","amount","deducted"],
    Settings: ["shopName","address","owner","gstin","logo","developer"],
  };
  SHEET_NAMES.forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
    }
  });
  SpreadsheetApp.flush();
  return "सर्व शीट्स तयार झाल्या ✅";
}

function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || "load";
    if (action === "backupDrive") return backupToDrive_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = {};
    SHEET_NAMES.forEach(function (name) {
      data[name] = sheetToJSON_(ss.getSheetByName(name));
    });
    return jsonOut_({ status: "ok", data: data });
  } catch (err) {
    return jsonOut_({ status: "error", message: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    SHEET_NAMES.forEach(function (name) {
      if (!body[name]) return;
      let sh = ss.getSheetByName(name);
      if (!sh) sh = ss.insertSheet(name);
      sh.clearContents();

      const rows = body[name];
      if (rows && rows.length) {
        const headers = Object.keys(rows[0]);
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
        // "phone" कॉलम नेहमी टेक्स्ट (Plain text) फॉरमॅटमध्ये ठेवा,
        // जेणेकरून Google Sheets मोबाईल नंबरला आपोआप Number मध्ये बदलणार नाही.
        const phoneColIdx = headers.indexOf("phone");
        if (phoneColIdx !== -1) {
          sh.getRange(2, phoneColIdx + 1, Math.max(rows.length, 1), 1).setNumberFormat("@");
        }
        const values = rows.map(function (r) {
          return headers.map(function (h) {
            const v = r[h];
            if (h === "phone") return String(v);
            return typeof v === "object" && v !== null ? JSON.stringify(v) : v;
          });
        });
        sh.getRange(2, 1, values.length, headers.length).setValues(values);
      }
    });

    return jsonOut_({ status: "ok" });
  } catch (err) {
    return jsonOut_({ status: "error", message: String(err) });
  }
}

/** सध्याची स्प्रेडशीट "Shree Pharmacy Backups" फोल्डरमध्ये तारीख-वेळेसह कॉपी करते */
function backupToDrive_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const file = DriveApp.getFileById(ss.getId());
    let folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(BACKUP_FOLDER_NAME);
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
    const copy = file.makeCopy(ss.getName() + " — Backup " + stamp, folder);
    return jsonOut_({ status: "ok", fileUrl: copy.getUrl() });
  } catch (err) {
    return jsonOut_({ status: "error", message: String(err) });
  }
}

function sheetToJSON_(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1)
    .filter(function (row) { return row.join("") !== ""; })
    .map(function (row) {
      const obj = {};
      headers.forEach(function (h, i) {
        let v = row[i];
        if (h === "phone") { obj[h] = String(v).trim(); return; }
        if (typeof v === "string" && (v.startsWith("[") || v.startsWith("{"))) {
          try { v = JSON.parse(v); } catch (e2) { /* plain string ठेवा */ }
        }
        obj[h] = v;
      });
      return obj;
    });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
