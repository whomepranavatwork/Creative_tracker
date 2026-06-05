// ============================================================
// Creative Tracker — Bulk Entry Automation
// ============================================================

const FORCE_SHEET_NAME    = "";  // leave blank — user picks tab in sidebar
const HEADER_SEARCH_LIMIT = 20;

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

// ── Single init call: replaces getSheetNames + getDropdownOptions ─
// Returns everything the sidebar needs on load in one round trip.
function init() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ss.getSheets().map(s => s.getName());

  return {
    sheetNames,
    dropdownOptions: {
      funnel:          ["BOF", "MOF", "TOF"],
      intInf:          ["Inf", "Int", "Hair Warriors"],
      adType:          ["Reel", "Static", "Carousel"],
      language:        ["Hinglish", "English", "Hindi"],
      adFormat:        ["Educational", "Non Vo", "Statics", "Testimonial",
                        "Carousel", "UGC", "Organic Reviews", "Hair Warriors"],
      narrative:       ["Problem-Solution", "Results/Assurance", "Myth Busting",
                        "Trust", "Features", "Why MM is different", "Safety",
                        "Science", "BeforeAfter", "Transplant or PRP",
                        "Product First", "Results/Assurance/ProductFirst"],
      creatorType:     ["", "Meta Creator", "Hair Warrior"],
      onboardingMonth: ["", "Oct'25", "Nov'25", "Dec'25", "Jan'26", "Feb'26",
                        "March'26", "April'26", "May'26", "June'26"],
      live:            ["No", "Yes"],
      canLive:         ["Yes", "No"],
      ratio:           ["9:16", "4:5"]
    }
  };
}

// ── Called when user picks a tab ──────────────────────────────
// Returns context AND colIndex/headerRow so sidebar can cache them.
function getSheetContext(tabName) {
  const { sheet, colIndex, headerRow } = resolveSheet(tabName);
  const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  // Single batched read: get S.No. column and Date Added column together
  const lastRow   = sheet.getLastRow();
  const dataRows  = lastRow - headerRow;
  let nextSno     = 1;
  let lastDataRow = headerRow;

  if (dataRows > 0) {
    // Read S.No. and Date Added columns in one getValues call
    const snoCol  = colIndex.sno;
    const dateCol = colIndex.date ?? colIndex.drive916 ?? colIndex.drive45;
    const minCol  = Math.min(snoCol, dateCol) + 1;           // 1-indexed
    const maxCol  = Math.max(snoCol, dateCol) + 1;
    const width   = maxCol - minCol + 1;

    const data = sheet
      .getRange(headerRow + 1, minCol, dataRows, width)
      .getValues();

    const snoOffset  = snoCol  + 1 - minCol;
    const dateOffset = dateCol + 1 - minCol;

    let maxSno = 0;
    for (let i = data.length - 1; i >= 0; i--) {
      const n = parseInt(data[i][snoOffset], 10);
      if (!isNaN(n) && n > maxSno) maxSno = n;
      if (lastDataRow === headerRow && data[i][dateOffset] !== "") {
        lastDataRow = headerRow + 1 + i;
      }
    }
    nextSno = maxSno + 1;
  }

  return {
    sheetName:   sheet.getName(),
    nextSno:     String(nextSno).padStart(5, "0"),
    today,
    totalRows:   dataRows,
    lastDataRow,           // cached — passed back with addEntries
    colIndex,              // cached — passed back with addEntries
    headerRow              // cached — passed back with addEntries
  };
}

// ── Called by sidebar: reads files from a Drive folder ────────
function getFolderFiles(folderUrl) {
  const folderId = extractDriveFolderId(folderUrl);
  if (!folderId) return { ok: false, msg: "Could not extract folder ID. Make sure it's a Google Drive folder link." };

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
      url:  "https://drive.google.com/file/d/" + f.getId() + "/view?usp=drive_link"
    });
  }
  files.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));

  if (files.length === 0) return { ok: false, msg: "No files found in that folder." };
  return { ok: true, files };
}

// ── Called by sidebar: writes rows to the sheet ───────────────
// payload includes cached colIndex, headerRow, lastDataRow, nextSno
// so no sheet re-scanning is needed.
function addEntries(payload) {
  try {
    const tabName    = payload.tabName;
    const shared     = payload.shared;
    const cuts       = payload.cuts;
    const colIndex   = payload.colIndex;
    const headerRow  = payload.headerRow;
    const nextSno    = parseInt(payload.nextSno, 10);
    const firstNewRow = payload.lastDataRow + 1;

    if (!cuts || cuts.length === 0) return { ok: false, msg: "No cuts to add." };

    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const name   = tabName || FORCE_SHEET_NAME;
    const sheet  = name ? ss.getSheetByName(name) : ss.getActiveSheet();
    if (!sheet) throw new Error("Sheet \"" + name + "\" not found.");

    const lastCol = sheet.getLastColumn();
    const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    const rows    = [];

    cuts.forEach((cut, i) => {
      const sno = String(nextSno + i).padStart(5, "0");
      const row = new Array(lastCol).fill("");

      set(row, colIndex.sno,             sno);
      set(row, colIndex.date,            shared.date || today);
      set(row, colIndex.product,         shared.product);
      // adName: leave blank — formula in sheet handles it
      set(row, colIndex.drive45,         cut.ratio === "4:5"  ? cut.url : "");
      set(row, colIndex.drive916,        cut.ratio === "9:16" ? cut.url : "");
      set(row, colIndex.live,            shared.live    || "No");
      set(row, colIndex.canLive,         shared.canLive || "Yes");
      set(row, colIndex.raisedBy,        shared.raisedBy || "Pranav");
      set(row, colIndex.funnel,          cut.funnel);
      set(row, colIndex.intInf,          shared.intInf);
      set(row, colIndex.adType,          shared.adType);
      set(row, colIndex.language,        shared.language);
      set(row, colIndex.person,          shared.person || "None");
      set(row, colIndex.narrative,       cut.narrative);
      set(row, colIndex.adFormat,        cut.adFormat);
      set(row, colIndex.onboardingMonth, shared.onboardingMonth || "");
      set(row, colIndex.additionalInfo,  shared.additionalInfo  || "");
      set(row, colIndex.instagram,       shared.instagram       || "");
      set(row, colIndex.creatorType,     shared.creatorType     || "");
      set(row, colIndex.ytLinks,         "");
      set(row, colIndex.dateTakenLive,   "");
      set(row, colIndex.ytAdsStatus,     "No");
      // ytAdName: leave blank — formula handles it

      rows.push(row);
    });

    // Write values
    sheet.getRange(firstNewRow, 1, rows.length, lastCol).setValues(rows);

    // Hyperlink Drive link columns
    setHyperlinks(sheet, firstNewRow, rows, colIndex);

    // Copy Ad Name / YT Ad Name formulas from the row above
    const formulaSourceRow = firstNewRow - 1;
    if (formulaSourceRow > headerRow) {
      [colIndex.adName, colIndex.ytAdName].forEach(col => {
        if (col == null) return;
        const srcCell = sheet.getRange(formulaSourceRow, col + 1);
        if (!srcCell.getFormula()) return;
        sheet.getRange(firstNewRow, col + 1, rows.length, 1)
          .copyFrom(srcCell, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
      });
    }

    return {
      ok:  true,
      msg: `Added ${rows.length} entr${rows.length === 1 ? "y" : "ies"} (S.No. ${String(nextSno).padStart(5,"0")} – ${String(nextSno + rows.length - 1).padStart(5,"0")}) on sheet "${sheet.getName()}", rows ${firstNewRow}–${firstNewRow + rows.length - 1}.`,
      // Return updated values so sidebar can refresh without another server call
      nextSno:     String(nextSno + rows.length).padStart(5, "0"),
      lastDataRow: firstNewRow + rows.length - 1,
      totalRows:   payload.totalRows + rows.length
    };

  } catch (e) {
    return { ok: false, msg: e.message + "\n" + e.stack };
  }
}

// ── Sheet resolution (only used by getSheetContext) ───────────
function resolveSheet(tabName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const name  = tabName || FORCE_SHEET_NAME;
  const sheet = name ? ss.getSheetByName(name) : ss.getActiveSheet();
  if (!sheet) throw new Error("Sheet \"" + name + "\" not found.");

  const { headerRow, colIndex } = detectHeaders(sheet);
  return { sheet, headerRow, colIndex };
}

function detectHeaders(sheet) {
  const limit = Math.min(HEADER_SEARCH_LIMIT, sheet.getLastRow());
  // Only read the first 30 columns — headers are always in the left portion
  const readCols = Math.min(30, sheet.getLastColumn());
  const data  = sheet.getRange(1, 1, limit, readCols).getValues();

  const reverseMap = {};
  Object.entries(HEADER_MAP).forEach(([key, text]) => reverseMap[text] = key);

  for (let r = 0; r < data.length; r++) {
    const colIndex = {};
    let matches = 0;
    data[r].forEach((cell, c) => {
      const key = reverseMap[String(cell).trim()];
      if (key) { colIndex[key] = c; matches++; }
    });
    if (matches >= 6) return { headerRow: r + 1, colIndex };
  }

  throw new Error(
    "Could not find the header row (looked at rows 1–" + limit + "). " +
    "Make sure the correct sheet tab is selected."
  );
}

// ── Helpers ───────────────────────────────────────────────────
function setHyperlinks(sheet, firstNewRow, rows, colIndex) {
  [colIndex.drive45, colIndex.drive916].forEach(col => {
    if (col == null) return;
    const richValues = rows.map(row => {
      const url = row[col];
      if (!url) return SpreadsheetApp.newRichTextValue().setText("").build();
      return SpreadsheetApp.newRichTextValue()
        .setText(url).setLinkUrl(url).build();
    });
    sheet.getRange(firstNewRow, col + 1, rows.length, 1)
      .setRichTextValues(richValues.map(v => [v]));
  });
}

function set(row, col, value) {
  if (col != null) row[col] = value;
}

function extractDriveFolderId(url) {
  const m = String(url).match(/folders\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}
