// ============================================================
// Creative Tracker — Bulk Entry Automation
// ============================================================

const SPREADSHEET_ID      = "1EknvJXuzSZTNKnT3Xr3jsAbPct9yEgCzLbbehwZCJ2E";
const FORCE_SHEET_NAME    = "";  // leave blank — user picks tab in sidebar
const HEADER_SEARCH_LIMIT = 20;

// Tab → allowed product names. Returned to the client via init() so changes
// here take effect on next page load — no webapp.html redeploy needed.
// Single-product tabs auto-select; multi-product tabs show a dropdown.
const TAB_PRODUCT_MAP = {
  "BGK":          ["Beard Growth Kit", "Beard Activator Kit", "Beard Development Kit", "Beard Gummies"],
  "Biotin":       ["Biotin30"],
  "S1":           ["Stage 1"],
  "S2":           ["Stage2"],
  "S3":           ["Stage3"],
  "Form Testing": ["Selfasst"],
  "Cetosomal":    ["Advance Regime"]
};

function getSpreadsheet() {
  return SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile("webapp")
    .setTitle("Creative Tracker")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

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
    .createMenu("Bias for Action!")
    .addItem("Open", "showSidebar")
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("Bias for Action!")
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── Single init call — returns everything the client needs on load ─
function init() {
  const ss = getSpreadsheet();
  const sheetNames = ss.getSheets().slice(0, 7).map(s => s.getName());

  const buckets = readBuckets(ss);

  return {
    sheetNames,
    tabProductMap: TAB_PRODUCT_MAP,
    personMap: buckets.personMap,
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

  // Columns (0-indexed): A=product, B=funnel, C=intInf, D=adType, E=language,
  // F=adFormat, H=person, I=instagram
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
    const snoCol    = colIndex.sno;
    const funnelCol = colIndex.funnel;

    // Guard: both columns must exist before reading
    if (snoCol != null && funnelCol != null) {
      const minCol = Math.min(snoCol, funnelCol) + 1;
      const maxCol = Math.max(snoCol, funnelCol) + 1;
      const width  = maxCol - minCol + 1;

      const data = sheet
        .getRange(headerRow + 1, minCol, dataRows, width)
        .getValues();

      const snoOffset    = snoCol    + 1 - minCol;
      const funnelOffset = funnelCol + 1 - minCol;

      for (let i = data.length - 1; i >= 0; i--) {
        if (data[i][funnelOffset] !== "") {
          lastDataRow = headerRow + 1 + i;
          const n = parseInt(data[i][snoOffset], 10);
          nextSno = isNaN(n) ? 1 : n + 1;
          break;
        }
      }
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
  // Single file link: drive.google.com/file/d/FILE_ID/...
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

  // Folder link
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
    const tabName   = payload.tabName;
    const shared    = payload.shared;
    const cuts      = payload.cuts;
    const colIndex  = payload.colIndex;
    const headerRow = payload.headerRow;

    if (!cuts || cuts.length === 0) return { ok: false, msg: "No cuts to add." };

    const ss    = getSpreadsheet();
    const name  = tabName || FORCE_SHEET_NAME;
    const sheet = name ? ss.getSheetByName(name) : ss.getActiveSheet();
    if (!sheet) throw new Error("Sheet \"" + name + "\" not found.");

    // Re-derive lastDataRow server-side to avoid race conditions from stale client state.
    // Scan funnel column from the bottom — same logic as getSheetContext().
    let lastDataRow = headerRow;
    let nextSno     = 1;
    const lastRow   = sheet.getLastRow();
    const dataRows  = lastRow - headerRow;

    if (dataRows > 0 && colIndex.sno != null && colIndex.funnel != null) {
      const minCol = Math.min(colIndex.sno, colIndex.funnel) + 1;
      const maxCol = Math.max(colIndex.sno, colIndex.funnel) + 1;
      const data   = sheet
        .getRange(headerRow + 1, minCol, dataRows, maxCol - minCol + 1)
        .getValues();
      const snoOff    = colIndex.sno    + 1 - minCol;
      const funnelOff = colIndex.funnel + 1 - minCol;
      for (let i = data.length - 1; i >= 0; i--) {
        if (data[i][funnelOff] !== "") {
          lastDataRow = headerRow + 1 + i;
          const n = parseInt(data[i][snoOff], 10);
          nextSno = isNaN(n) ? 1 : n + 1;
          break;
        }
      }
    } else {
      // Fallback to client value if funnel/sno columns missing
      lastDataRow = payload.lastDataRow;
      nextSno     = parseInt(payload.nextSno, 10);
    }

    const firstNewRow = lastDataRow + 1;

    // Overwrite guard: drive link columns are never pre-filled, so any value = real entry
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

    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    const n     = cuts.length;

    // Build per-column value arrays — only columns we explicitly own.
    // Writing column-by-column avoids triggering data validation rules on
    // adjacent columns we don't touch (the previous full-row setValues approach
    // wrote empty strings across the entire row, tripping sheet validations).
    const colVals = {}; // colIndex (0-based) → [val for row 0, val for row 1, ...]
    const col = (key, valFn) => {
      const c = colIndex[key];
      if (c == null) return;
      colVals[c] = cuts.map((cut, i) => valFn(cut, i));
    };

    col("sno",             (_, i) => nextSno + i);
    col("date",            ()     => shared.date || today);
    col("product",         ()     => shared.product);
    col("drive45",         cut    => (cut.ratio === "4:5"  || cut.ratio === "Both") ? cut.url : "");
    col("drive916",        cut    => (cut.ratio === "9:16" || cut.ratio === "Both") ? cut.url : "");
    col("live",            ()     => "No");
    col("canLive",         ()     => shared.canLive  || "Yes");
    col("raisedBy",        ()     => shared.raisedBy || "");
    col("funnel",          cut    => cut.funnel);
    col("intInf",          ()     => shared.intInf);
    col("adType",          ()     => shared.adType);
    col("language",        ()     => shared.language);
    col("person",          ()     => shared.person   || "");
    col("narrative",       cut    => cut.narrative);
    col("adFormat",        cut    => cut.adFormat);
    col("onboardingMonth", ()     => shared.onboardingMonth || "");
    col("additionalInfo",  ()     => shared.additionalInfo  || "");
    col("instagram",       ()     => shared.instagram       || "");
    col("creatorType",     ()     => shared.creatorType     || "");
    col("ytAdsStatus",     ()     => "No");

    // Write contiguous blocks of owned columns in one setValues call each.
    // Columns with gaps > 1 start a new block. This keeps API calls to 2–4
    // while never writing to unowned columns (which may have strict validation).
    const sorted = Object.entries(colVals)
      .map(([c, vals]) => ({ c: parseInt(c), vals }))
      .sort((a, b) => a.c - b.c);

    let i = 0;
    while (i < sorted.length) {
      let j = i;
      while (j + 1 < sorted.length && sorted[j + 1].c === sorted[j].c + 1) j++;

      const startCol = sorted[i].c;
      const width    = sorted[j].c - startCol + 1;
      const colMap   = {};
      for (let k = i; k <= j; k++) colMap[sorted[k].c] = sorted[k].vals;

      const data = Array.from({ length: n }, (_, r) =>
        Array.from({ length: width }, (__, o) => {
          const vals = colMap[startCol + o];
          return vals ? vals[r] : "";
        })
      );

      sheet.getRange(firstNewRow, startCol + 1, n, width).setValues(data);
      i = j + 1;
    }

    if (colIndex.sno != null) {
      sheet.getRange(firstNewRow, colIndex.sno + 1, n, 1).setNumberFormat("00000");
    }

    setHyperlinks(sheet, firstNewRow, cuts, colIndex);

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

    const firstSno = String(nextSno).padStart(5, "0");
    const lastSno  = String(nextSno + rows.length - 1).padStart(5, "0");
    const snoRange = rows.length === 1 ? firstSno : `${firstSno}–${lastSno}`;

    return {
      ok:  true,
      msg: `${rows.length} row${rows.length === 1 ? "" : "s"} added to ${sheet.getName()} · S.No. ${snoRange}`,
      nextSno:     String(nextSno + rows.length).padStart(5, "0"),
      lastDataRow: firstNewRow + rows.length - 1,
      totalRows:   dataRows + rows.length
    };

  } catch (e) {
    return { ok: false, msg: e.message + "\n" + e.stack };
  }
}

// ── Sheet resolution ──────────────────────────────────────────
function resolveSheet(tabName) {
  const ss    = getSpreadsheet();
  const name  = tabName || FORCE_SHEET_NAME;
  const sheet = name ? ss.getSheetByName(name) : ss.getActiveSheet();
  if (!sheet) throw new Error("Sheet \"" + name + "\" not found.");

  const { headerRow, colIndex } = detectHeaders(sheet);
  return { sheet, headerRow, colIndex };
}

function detectHeaders(sheet) {
  const limit    = Math.min(HEADER_SEARCH_LIMIT, sheet.getLastRow());
  const readCols = Math.min(sheet.getLastColumn(), 50); // covers any realistic tracker layout
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
function setHyperlinks(sheet, firstNewRow, cuts, colIndex) {
  const make = url => url
    ? SpreadsheetApp.newRichTextValue().setText(url).setLinkUrl(url).build()
    : SpreadsheetApp.newRichTextValue().setText("").build();

  if (colIndex.drive45 != null) {
    sheet.getRange(firstNewRow, colIndex.drive45 + 1, cuts.length, 1)
      .setRichTextValues(cuts.map(cut =>
        [(cut.ratio === "4:5" || cut.ratio === "Both") ? make(cut.url) : make("")]));
  }
  if (colIndex.drive916 != null) {
    sheet.getRange(firstNewRow, colIndex.drive916 + 1, cuts.length, 1)
      .setRichTextValues(cuts.map(cut =>
        [(cut.ratio === "9:16" || cut.ratio === "Both") ? make(cut.url) : make("")]));
  }
}

function set(row, col, value) {
  if (col != null) row[col] = value;
}

function extractDriveFolderId(url) {
  const m = String(url).match(/folders\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}
