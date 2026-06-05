// ============================================================
// Creative Tracker — Bulk Entry Automation
// Google Apps Script (bound to the tracker sheet)
// ============================================================
// INSTALL:
//   1. Open your Creative Tracker Google Sheet
//   2. Extensions → Apps Script
//   3. Replace any existing code with this file
//   4. Add a second file named "Sidebar" and paste Sidebar.html
//   5. Save both files, then refresh the sheet
//   6. A "Creative Tracker" menu will appear in the sheet
// ============================================================

// ── Config ───────────────────────────────────────────────────
// The script auto-detects the active sheet, but set this if
// you always want a specific tab.
const FORCE_SHEET_NAME = ""; // e.g. "Cetosomal" — leave blank to use active sheet

// Row number where headers live (S.No., Date Added, Product Name…)
// The script will search rows 1–20 for the header row automatically.
const HEADER_SEARCH_LIMIT = 20;

// ── Column identifiers (matched by header text) ───────────────
// These must match exactly what's in your header row.
const HEADER_MAP = {
  sno:             "S.No.",
  date:            "Date Added",
  product:         "Product Name",
  adName:          "Ad Name",
  drive45:         "Google Drive Link\n(4:5)",
  drive916:        "Google Drive Link\n(9:16)",
  live:            "Live?",
  canLive:         "Can take live?",
  raisedBy:        "Raised By",
  funnel:          "Funnel Type",
  intInf:          "INT / INF",
  adType:          "Ad Type",
  language:        "Language",
  person:          "Person Full Name",
  narrative:       "Broad Narrative - P0",
  adFormat:        "Ad Format",
  onboardingMonth: "Inf Onboarding Month",
  additionalInfo:  "Additional information",
  instagram:       "Instagram Profile",
  creatorType:     "Creator Type",
  ytLinks:         "YT Links",
  dateTakenLive:   "Date Taken Live",
  ytAdsStatus:     "YT Ads Status",
  ytAdName:        "YT Ad Name"
};

// ── Menu ─────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Creative Tracker")
    .addItem("Add New Entries…", "showSidebar")
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("Add New Entries")
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── Called by sidebar on load ─────────────────────────────────
function getSheetContext() {
  const { sheet, colIndex, headerRow } = resolveSheet();
  const nextSno = getNextSerialNumber(sheet, colIndex.sno, headerRow);
  const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  return {
    sheetName:  sheet.getName(),
    nextSno:    String(nextSno).padStart(5, "0"),
    today:      today,
    totalRows:  sheet.getLastRow() - headerRow
  };
}

// ── Called by sidebar: reads files from a Drive folder ────────
function getFolderFiles(folderUrl) {
  const folderId = extractDriveFolderId(folderUrl);
  if (!folderId) return { ok: false, msg: "Could not extract folder ID from that URL. Make sure it's a Google Drive folder link." };

  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    return { ok: false, msg: "Could not access folder. Make sure it's shared with your Google account." };
  }

  const files = [];
  const iter  = folder.getFiles();
  while (iter.hasNext()) {
    const f = iter.next();
    files.push({
      id:   f.getId(),
      name: f.getName(),
      url:  "https://drive.google.com/file/d/" + f.getId() + "/view?usp=drive_link",
      mimeType: f.getMimeType()
    });
  }

  // Sort by name so cuts appear in a predictable order
  files.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));

  if (files.length === 0) return { ok: false, msg: "No files found in that folder." };

  return { ok: true, files };
}

// ── Called by sidebar: writes rows to the sheet ───────────────
function addEntries(payload) {
  // payload = { shared: {...}, cuts: [{funnel, narrative, adFormat, ratio}, ...] }
  try {
    const { sheet, colIndex, headerRow } = resolveSheet();
    const totalCols = Object.keys(colIndex).length;
    // Use the actual last column from the sheet
    const lastCol   = sheet.getLastColumn();

    const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    const nextSno = getNextSerialNumber(sheet, colIndex.sno, headerRow);
    const shared  = payload.shared;
    const cuts    = payload.cuts;

    if (!cuts || cuts.length === 0) return { ok: false, msg: "No cuts to add." };

    // Find the row that has the formula for Ad Name so we can copy it
    const formulaSourceRow = findFormulaRow(sheet, colIndex.adName, headerRow);

    const newRowIndex = sheet.getLastRow() + 1;
    const rows = [];

    cuts.forEach((cut, i) => {
      const sno = String(nextSno + i).padStart(5, "0");

      // Build a sparse row using column indices
      const row = new Array(lastCol).fill("");

      setValue(row, colIndex.sno,             sno);
      setValue(row, colIndex.date,            shared.date || today);
      setValue(row, colIndex.product,         shared.product);
      // adName is left blank — formula fills it
      setValue(row, colIndex.drive45,         cut.ratio === "4:5"  ? cut.url : "");
      setValue(row, colIndex.drive916,        cut.ratio === "9:16" ? cut.url : "");
      setValue(row, colIndex.live,            shared.live    || "No");
      setValue(row, colIndex.canLive,         shared.canLive || "Yes");
      setValue(row, colIndex.raisedBy,        shared.raisedBy || "Pranav");
      setValue(row, colIndex.funnel,          cut.funnel);
      setValue(row, colIndex.intInf,          shared.intInf);
      setValue(row, colIndex.adType,          shared.adType);
      setValue(row, colIndex.language,        shared.language);
      setValue(row, colIndex.person,          shared.person || "None");
      setValue(row, colIndex.narrative,       cut.narrative);
      setValue(row, colIndex.adFormat,        cut.adFormat);
      setValue(row, colIndex.onboardingMonth, shared.onboardingMonth || "None");
      setValue(row, colIndex.additionalInfo,  shared.additionalInfo  || "");
      setValue(row, colIndex.instagram,       shared.instagram       || "");
      setValue(row, colIndex.creatorType,     shared.creatorType     || "");
      setValue(row, colIndex.ytLinks,         "");
      setValue(row, colIndex.dateTakenLive,   "");
      setValue(row, colIndex.ytAdsStatus,     "No");
      // ytAdName left blank — formula fills it

      rows.push(row);
    });

    // Write values
    sheet.getRange(newRowIndex, 1, rows.length, lastCol).setValues(rows);

    // Copy Ad Name formula (and YT Ad Name formula) from the source row
    if (formulaSourceRow) {
      const adNameCol   = colIndex.adName;
      const ytNameCol   = colIndex.ytAdName;

      [adNameCol, ytNameCol].forEach(col => {
        if (col == null) return;
        const srcCell = sheet.getRange(formulaSourceRow, col + 1);
        const formula = srcCell.getFormula();
        if (!formula) return;

        // Copy formula to each new row
        for (let i = 0; i < rows.length; i++) {
          const destCell = sheet.getRange(newRowIndex + i, col + 1);
          // Re-anchor row references in the formula
          const adjusted = adjustFormulaRow(formula, formulaSourceRow, newRowIndex + i);
          destCell.setFormula(adjusted);
        }
      });
    }

    return {
      ok:  true,
      msg: `✅ Added ${rows.length} entr${rows.length === 1 ? "y" : "ies"} starting at S.No. ${String(nextSno).padStart(5, "0")}.`
    };

  } catch (e) {
    return { ok: false, msg: "Error: " + e.message + "\n" + e.stack };
  }
}

// ── Sheet resolution ──────────────────────────────────────────
function resolveSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = FORCE_SHEET_NAME
    ? ss.getSheetByName(FORCE_SHEET_NAME)
    : ss.getActiveSheet();

  if (!sheet) throw new Error(`Sheet "${FORCE_SHEET_NAME}" not found.`);

  const { headerRow, colIndex } = detectHeaders(sheet);
  return { sheet, headerRow, colIndex };
}

function detectHeaders(sheet) {
  const limit  = Math.min(HEADER_SEARCH_LIMIT, sheet.getLastRow());
  const data   = sheet.getRange(1, 1, limit, sheet.getLastColumn()).getValues();

  // Build reverse map: header text → key
  const reverseMap = {};
  Object.entries(HEADER_MAP).forEach(([key, text]) => reverseMap[text] = key);

  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    const colIndex = {};
    let matches = 0;
    row.forEach((cell, c) => {
      const cellText = String(cell).trim();
      if (reverseMap[cellText]) {
        colIndex[reverseMap[cellText]] = c; // 0-indexed
        matches++;
      }
    });
    if (matches >= 6) { // enough columns found to be confident
      return { headerRow: r + 1, colIndex }; // headerRow is 1-indexed
    }
  }
  throw new Error("Could not find the header row. Check that your sheet tab is active and headers match expected values.");
}

// ── Helpers ───────────────────────────────────────────────────
function getNextSerialNumber(sheet, snoColIndex, headerRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRow) return 1;

  const values = sheet
    .getRange(headerRow + 1, snoColIndex + 1, lastRow - headerRow, 1)
    .getValues();

  let max = 0;
  values.forEach(([v]) => {
    const n = parseInt(v, 10);
    if (!isNaN(n) && n > max) max = n;
  });
  return max + 1;
}

function findFormulaRow(sheet, adNameColIndex, headerRow) {
  if (adNameColIndex == null) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRow) return null;

  // Walk from the last row upward to find a row with a formula in Ad Name col
  for (let r = lastRow; r > headerRow; r--) {
    const formula = sheet.getRange(r, adNameColIndex + 1).getFormula();
    if (formula) return r;
  }
  return null;
}

function adjustFormulaRow(formula, sourceRow, destRow) {
  // Replace absolute row references specific to sourceRow with destRow
  // This handles simple CONCATENATE/& style formulas that reference same-row cells
  return formula.replace(new RegExp(sourceRow, "g"), destRow);
}

function setValue(row, colIndex, value) {
  if (colIndex != null) row[colIndex] = value;
}

function extractDriveFolderId(url) {
  const m = url.match(/folders\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}

// ── Expose dropdown options to sidebar ───────────────────────
function getDropdownOptions() {
  return {
    funnel:    ["BOF", "MOF", "TOF"],
    intInf:    ["Inf", "Int", "Hair Warriors"],
    adType:    ["Reel", "Static", "Carousel"],
    language:  ["Hinglish", "English", "Hindi"],
    adFormat:  ["Educational", "Non Vo", "Statics", "Testimonial",
                "Carousel", "UGC", "Organic Reviews", "Hair Warriors"],
    narrative: ["Problem-Solution", "Results/Assurance", "Myth Busting",
                "Trust", "Features", "Why MM is different", "Safety",
                "Science", "BeforeAfter", "Transplant or PRP",
                "Product First", "Results/Assurance/ProductFirst"],
    creatorType: ["", "Meta Creator", "Hair Warrior"],
    live:      ["No", "Yes"],
    canLive:   ["Yes", "No"],
    ratio:     ["9:16", "4:5"]
  };
}
