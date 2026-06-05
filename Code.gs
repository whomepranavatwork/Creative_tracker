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
    .createMenu("Bias for Action")
    .addItem("Open", "showSidebar")
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("Bias for Action")
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── Single init call: replaces getSheetNames + getDropdownOptions ─
// Returns everything the sidebar needs on load in one round trip.
function init() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ss.getSheets().slice(0, 7).map(s => s.getName());

  // Read Buckets tab for dynamic dropdown values
  const buckets = readBuckets(ss);

  return {
    sheetNames,
    personMap: buckets.personMap,
    dropdownOptions: {
      product:         buckets.product,
      funnel:          buckets.funnel.length   ? buckets.funnel   : ["BOF", "MOF", "TOF"],
      intInf:          buckets.intInf.length   ? buckets.intInf   : ["Inf", "Int", "Hair Warriors"],
      adType:          buckets.adType.length   ? buckets.adType   : ["Reel", "Static", "Carousel"],
      language:        buckets.language.length ? buckets.language : ["Hinglish", "English", "Hindi"],
      adFormat:        buckets.adFormat.length ? buckets.adFormat : ["Educational", "Non Vo", "Statics",
                         "Testimonial", "Carousel", "UGC", "Organic Reviews", "Hair Warriors"],
      person:          ["None", ...buckets.person],
      narrative:       ["Problem-Solution", "Results/Assurance", "Myth Busting",
                        "Trust", "Features", "Why MM is different", "Safety",
                        "Science", "BeforeAfter", "Transplant or PRP",
                        "Product First", "Results/Assurance/ProductFirst"],
      creatorType:     ["", "Meta Creator", "Hair Warrior"],
      onboardingMonth: ["", "Oct'25", "Nov'25", "Dec'25", "Jan'26", "Feb'26",
                        "March'26", "April'26", "May'26", "June'26"],
      canLive:         ["Yes", "No"],
      ratio:           ["9:16", "4:5", "Both"]
    }
  };
}

function readBuckets(ss) {
  const result = { product: [], funnel: [], intInf: [], adType: [], language: [], adFormat: [], person: [], personMap: {} };
  const sheet = ss.getSheetByName("Buckets");
  if (!sheet) return result;

  // Columns (0-indexed): A=product, B=funnel, C=intInf, D=adType, E=language, F=adFormat, H=person, I=instagram
  const data = sheet.getDataRange().getValues();
  const seen = { product: new Set(), funnel: new Set(), intInf: new Set(), adType: new Set(), language: new Set(), adFormat: new Set(), person: new Set() };

  data.slice(1).forEach(row => {
    const add = (key, col) => {
      const v = String(row[col] || "").trim();
      if (v && !seen[key].has(v)) { seen[key].add(v); result[key].push(v); }
    };
    add("product",  0);
    add("funnel",   1);
    add("intInf",   2);
    add("adType",   3);
    add("language", 4);
    add("adFormat", 5);
    add("person",   7);
    const person = String(row[7] || "").trim();
    const ig     = String(row[8] || "").trim();
    if (person && ig) result.personMap[person] = ig;
  });

  return result;
}

// ── Called when user picks a tab ──────────────────────────────
// Returns context AND colIndex/headerRow so sidebar can cache them.
function getSheetContext(tabName) {
  const { sheet, colIndex, headerRow } = resolveSheet(tabName);
  const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  const lastRow  = sheet.getLastRow();
  const dataRows = lastRow - headerRow;
  let nextSno          = 1;
  let lastDataRow      = headerRow;
  let adNameFormula    = "";
  let adNameFormulaRow = null;

  if (dataRows > 0) {
    // Use Funnel Type to find lastDataRow — pre-filled rows never have it.
    // Date Added is unreliable because a previous buggy run may have written dates there.
    // Read S.No. and Funnel Type columns together in one API call.
    const snoCol    = colIndex.sno;
    const funnelCol = colIndex.funnel;
    const minCol    = Math.min(snoCol, funnelCol) + 1;
    const maxCol    = Math.max(snoCol, funnelCol) + 1;
    const width     = maxCol - minCol + 1;

    const data = sheet
      .getRange(headerRow + 1, minCol, dataRows, width)
      .getValues();

    const snoOffset    = snoCol    + 1 - minCol;
    const funnelOffset = funnelCol + 1 - minCol;

    // Walk from bottom — first row with a Funnel value is the last real entry
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i][funnelOffset] !== "") {
        lastDataRow = headerRow + 1 + i;
        // nextSno comes from THIS row's S.No., not the max across all rows
        const n = parseInt(data[i][snoOffset], 10);
        nextSno = isNaN(n) ? 1 : n + 1;
        break;
      }
    }

    // Scan the entire Ad Name column in one call to find a row with the formula.
    // We can't rely on lastDataRow alone because the script may have overwritten
    // that row's formula with setValues("") in a previous run.
    if (colIndex.adName != null) {
      const allFormulas = sheet
        .getRange(headerRow + 1, colIndex.adName + 1, lastRow - headerRow, 1)
        .getFormulas();
      for (let i = allFormulas.length - 1; i >= 0; i--) {
        if (allFormulas[i][0]) {
          adNameFormula    = allFormulas[i][0];
          adNameFormulaRow = headerRow + 1 + i; // absolute 1-indexed row
          break;
        }
      }
    }
  }

  return {
    sheetName:       sheet.getName(),
    nextSno:         String(nextSno).padStart(5, "0"),
    today,
    totalRows:       dataRows,
    lastDataRow,
    adNameFormula,
    adNameFormulaRow,
    colIndex,
    headerRow
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
    const tabName      = payload.tabName;
    const shared       = payload.shared;
    const cuts         = payload.cuts;
    const colIndex     = payload.colIndex;
    const headerRow    = payload.headerRow;
    const nextSno      = parseInt(payload.nextSno, 10);
    const firstNewRow  = payload.lastDataRow + 1;

    if (!cuts || cuts.length === 0) return { ok: false, msg: "No cuts to add." };

    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const name   = tabName || FORCE_SHEET_NAME;
    const sheet  = name ? ss.getSheetByName(name) : ss.getActiveSheet();
    if (!sheet) throw new Error("Sheet \"" + name + "\" not found.");

    const lastCol = sheet.getLastColumn();
    const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    const rows    = [];

    cuts.forEach((cut, i) => {
      const row = new Array(lastCol).fill("");

      set(row, colIndex.sno,             nextSno + i);
      set(row, colIndex.date,            shared.date || today);
      set(row, colIndex.product,         shared.product);
      // adName: leave blank — formula in sheet handles it
      set(row, colIndex.drive45,         (cut.ratio === "4:5"  || cut.ratio === "Both") ? cut.url : "");
      set(row, colIndex.drive916,        (cut.ratio === "9:16" || cut.ratio === "Both") ? cut.url : "");
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

    // Format S.No. column as zero-padded 5-digit (setValues converts strings to numbers)
    if (colIndex.sno != null) {
      sheet.getRange(firstNewRow, colIndex.sno + 1, rows.length, 1).setNumberFormat("00000");
    }

    // Hyperlink Drive link columns
    setHyperlinks(sheet, firstNewRow, rows, colIndex);

    // Copy Ad Name formula from the row where we actually found it
    // (adNameFormulaRow). copyTo adjusts relative references automatically,
    // so it doesn't matter if the source is a template row far below lastDataRow.
    if (colIndex.adName != null && payload.adNameFormulaRow) {
      const srcCell = sheet.getRange(payload.adNameFormulaRow, colIndex.adName + 1);
      srcCell.copyTo(
        sheet.getRange(firstNewRow, colIndex.adName + 1, rows.length, 1),
        SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false
      );
    }

    if (colIndex.ytAdName != null && payload.adNameFormulaRow) {
      const srcCell = sheet.getRange(payload.adNameFormulaRow, colIndex.ytAdName + 1);
      if (srcCell.getFormula()) {
        srcCell.copyTo(
          sheet.getRange(firstNewRow, colIndex.ytAdName + 1, rows.length, 1),
          SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false
        );
      }
    }

    return {
      ok:  true,
      msg: `${rows.length} row${rows.length === 1 ? "" : "s"} added to ${sheet.getName()}.`,
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
