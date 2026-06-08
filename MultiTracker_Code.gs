// ============================================================
// Creative Tracker — MULTI-TRACKER VERSION
// Deploy this file as "Code.gs" in a separate Apps Script project.
// Original single-tracker files (Code.gs, webapp.html, Sidebar.html)
// are NOT used by this version.
// ============================================================

// ── Tracker registry ─────────────────────────────────────────
// Add / remove trackers here. The key is the display name shown
// in the tracker selector; spreadsheetId is the Sheet ID from its URL.
const TRACKERS = {
  "Hair Stage": {
    spreadsheetId: "1EknvJXuzSZTNKnT3Xr3jsAbPct9yEgCzLbbehwZCJ2E",
    tabProductMap: {
      "BGK":          ["Beard Growth Kit", "Beard Activator Kit", "Beard Development Kit", "Beard Gummies"],
      "Biotin":       ["Biotin30"],
      "S1":           ["Stage 1"],
      "S2":           ["Stage2"],
      "S3":           ["Stage3"],
      "Form Testing": ["Selfasst"],
      "Cetosomal":    ["Advance Regime"]
    }
  },
  "Hair Stage 2": {
    spreadsheetId: "1DbcV58XCrHQqjqg1Kl6JaEsBHrpBA7jGiqGOB71wj1Y",
    tabProductMap: {
      "BGK":          ["Beard Growth Kit", "Beard Activator Kit", "Beard Development Kit", "Beard Gummies"],
      "Biotin":       ["Biotin30"],
      "S1":           ["Stage 1"],
      "S2":           ["Stage2"],
      "S3":           ["Stage3"],
      "Form Testing": ["Selfasst"],
      "Cetosomal":    ["Advance Regime"]
    }
  }
};

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

// ── Helpers ───────────────────────────────────────────────────
function getSpreadsheetFor(trackerName) {
  const t = TRACKERS[trackerName];
  if (!t) throw new Error("Unknown tracker: " + trackerName);
  return SpreadsheetApp.openById(t.spreadsheetId);
}

// ── Web app entry point ───────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile("webapp")
    .setTitle("Creative Tracker")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Called on page load: returns tracker list + first tracker data ──
function getTrackers() {
  const trackerNames = Object.keys(TRACKERS);
  const first = trackerNames[0];
  const data  = selectTracker(first);
  data.trackerNames = trackerNames;
  data.activeTracker = first;
  return data;
}

// ── Called when user switches tracker ────────────────────────
function selectTracker(trackerName) {
  const ss      = getSpreadsheetFor(trackerName);
  const t       = TRACKERS[trackerName];
  const sheetNames = ss.getSheets().slice(0, 7).map(s => s.getName());
  const buckets = readBuckets(ss);

  return {
    trackerName,
    sheetNames,
    tabProductMap: t.tabProductMap,
    personMap:     buckets.personMap,
    dropdownOptions: {
      product:         buckets.product,
      funnel:          buckets.funnel.length   ? buckets.funnel   : ["BOF", "MOF", "TOF"],
      intInf:          buckets.intInf.length   ? buckets.intInf   : ["Inf", "Int", "Hair Warriors"],
      adType:          buckets.adType.length   ? buckets.adType   : ["Reel", "Static", "Carousel"],
      language:        buckets.language.length ? buckets.language : ["Hinglish", "English", "Hindi"],
      adFormat:        buckets.adFormat.length ? buckets.adFormat : ["Educational", "Non Vo", "Statics",
                         "Testimonial", "Carousel", "UGC", "Organic Reviews", "Hair Warriors"],
      raisedBy:        ["Abhijeet","Varsha","Bhavya","Devansh","Lakshya","Purva","Jay","Cath","Disha","Aniket","Nivedita","Aakash","Karan","Pranav","Dipen","Sutanto"],
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
  const result = {
    product: [], funnel: [], intInf: [], adType: [], language: [],
    adFormat: [], person: [], personMap: {}
  };
  const sheet = ss.getSheetByName("Buckets");
  if (!sheet) return result;

  const data = sheet.getDataRange().getValues();
  const seen = {
    product: new Set(), funnel: new Set(), intInf: new Set(), adType: new Set(),
    language: new Set(), adFormat: new Set(), person: new Set()
  };

  data.slice(1).forEach(row => {
    const add = (key, col) => {
      if (col >= row.length) return;
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
    const ig     = String(row.length > 8 ? row[8] : "" || "").trim();
    if (person && ig) result.personMap[person] = ig;
  });

  return result;
}

// ── Called when user picks a tab ──────────────────────────────
function getSheetContext(tabName, trackerName) {
  const { sheet, colIndex, headerRow } = resolveSheet(tabName, trackerName);
  const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  const lastRow  = sheet.getLastRow();
  const dataRows = lastRow - headerRow;
  let nextSno          = 1;
  let lastDataRow      = headerRow;
  let adNameFormula    = "";
  let adNameFormulaRow = null;

  if (dataRows > 0) {
    const signalCols = [colIndex.drive916, colIndex.drive45, colIndex.date].filter(c => c != null);

    if (signalCols.length > 0) {
      const minCol = Math.min(...signalCols) + 1;
      const maxCol = Math.max(...signalCols) + 1;
      const data   = sheet.getRange(headerRow + 1, minCol, dataRows, maxCol - minCol + 1).getValues();

      for (let i = data.length - 1; i >= 0; i--) {
        const isReal = signalCols.some(c => data[i][c + 1 - minCol] !== "");
        if (isReal) { lastDataRow = headerRow + 1 + i; break; }
      }
    }

    if (colIndex.sno != null && lastDataRow > headerRow) {
      const snoData = sheet
        .getRange(headerRow + 1, colIndex.sno + 1, lastDataRow - headerRow, 1)
        .getValues();
      let maxSno = 0;
      snoData.forEach(r => { const n = parseInt(r[0], 10); if (!isNaN(n) && n > maxSno) maxSno = n; });
      nextSno = maxSno + 1;
    }

    if (colIndex.adName != null) {
      const allFormulas = sheet
        .getRange(headerRow + 1, colIndex.adName + 1, lastRow - headerRow, 1)
        .getFormulas();
      for (let i = allFormulas.length - 1; i >= 0; i--) {
        if (allFormulas[i][0]) {
          adNameFormula    = allFormulas[i][0];
          adNameFormulaRow = headerRow + 1 + i;
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

// ── Called by client: reads files from a Drive folder or single file ──
function getFolderFiles(driveUrl) {
  const fileMatch = String(driveUrl).match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (fileMatch) {
    try {
      const f = DriveApp.getFileById(fileMatch[1]);
      return { ok: true, files: [{
        id:   f.getId(),
        name: f.getName(),
        url:  "https://drive.google.com/file/d/" + f.getId() + "/view?usp=drive_link"
      }]};
    } catch (e) {
      return { ok: false, msg: "Could not access file. Make sure it's shared with your Google account." };
    }
  }

  const folderId = extractDriveFolderId(driveUrl);
  if (!folderId) return { ok: false, msg: "Could not extract folder or file ID. Paste a Google Drive folder or file link." };

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

// ── Called by client: writes rows to the sheet ───────────────
function addEntries(payload) {
  try {
    const tabName     = payload.tabName;
    const trackerName = payload.trackerName;
    const shared      = payload.shared;
    const cuts        = payload.cuts;
    const colIndex    = payload.colIndex;
    const headerRow   = payload.headerRow;
    const nextSno     = parseInt(payload.nextSno, 10);
    const firstNewRow = payload.lastDataRow + 1;

    if (!cuts || cuts.length === 0) return { ok: false, msg: "No cuts to add." };

    const ss    = getSpreadsheetFor(trackerName);
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) throw new Error("Sheet \"" + tabName + "\" not found.");

    const checkCol = colIndex.drive916 != null ? colIndex.drive916 : colIndex.drive45;
    if (checkCol != null) {
      const existing = sheet
        .getRange(firstNewRow, checkCol + 1, cuts.length, 1)
        .getValues();
      const conflict = existing.findIndex(r => r[0] !== "");
      if (conflict !== -1) {
        return {
          ok: false,
          msg: `Row ${firstNewRow + conflict} already has data — aborting to avoid overwrite. Refresh the page and try again.`
        };
      }
    }

    const lastCol = sheet.getLastColumn();
    const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    const rows    = [];

    cuts.forEach((cut, i) => {
      const row = new Array(lastCol).fill("");

      set(row, colIndex.sno,             nextSno + i);
      set(row, colIndex.date,            shared.date || today);
      set(row, colIndex.product,         shared.product);
      set(row, colIndex.drive45,         (cut.ratio === "4:5"  || cut.ratio === "Both") ? cut.url : "");
      set(row, colIndex.drive916,        (cut.ratio === "9:16" || cut.ratio === "Both") ? cut.url : "");
      set(row, colIndex.live,            "No");
      set(row, colIndex.canLive,         shared.canLive  || "Yes");
      set(row, colIndex.raisedBy,        shared.raisedBy || "");
      set(row, colIndex.funnel,          cut.funnel);
      set(row, colIndex.intInf,          shared.intInf);
      set(row, colIndex.adType,          shared.adType);
      set(row, colIndex.language,        shared.language);
      set(row, colIndex.person,          shared.person   || "None");
      set(row, colIndex.narrative,       cut.narrative);
      set(row, colIndex.adFormat,        cut.adFormat);
      set(row, colIndex.onboardingMonth, shared.onboardingMonth || "");
      set(row, colIndex.additionalInfo,  shared.additionalInfo  || "");
      set(row, colIndex.instagram,       shared.instagram       || "");
      set(row, colIndex.creatorType,     shared.creatorType     || "");
      set(row, colIndex.ytLinks,         "");
      set(row, colIndex.dateTakenLive,   "");
      set(row, colIndex.ytAdsStatus,     "No");

      rows.push(row);
    });

    sheet.getRange(firstNewRow, 1, rows.length, lastCol).setValues(rows);

    if (colIndex.sno != null) {
      sheet.getRange(firstNewRow, colIndex.sno + 1, rows.length, 1).setNumberFormat("00000");
    }

    setHyperlinks(sheet, firstNewRow, rows, colIndex);

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
      nextSno:     String(nextSno + rows.length).padStart(5, "0"),
      lastDataRow: firstNewRow + rows.length - 1,
      totalRows:   payload.totalRows + rows.length
    };

  } catch (e) {
    return { ok: false, msg: e.message + "\n" + e.stack };
  }
}

// ── Sheet resolution ──────────────────────────────────────────
function resolveSheet(tabName, trackerName) {
  const ss    = getSpreadsheetFor(trackerName);
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) throw new Error("Sheet \"" + tabName + "\" not found.");

  const { headerRow, colIndex } = detectHeaders(sheet);
  return { sheet, headerRow, colIndex };
}

function detectHeaders(sheet) {
  const limit    = Math.min(HEADER_SEARCH_LIMIT, sheet.getLastRow());
  const readCols = Math.min(sheet.getLastColumn(), 50);
  const data     = sheet.getRange(1, 1, limit, readCols).getValues();

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
      return SpreadsheetApp.newRichTextValue().setText(url).setLinkUrl(url).build();
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
