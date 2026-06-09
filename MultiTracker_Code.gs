// ============================================================
// Creative Tracker — MULTI-TRACKER VERSION
// Deploy this file as "Code.gs" in a separate Apps Script project.
// Original single-tracker files (Code.gs, webapp.html, Sidebar.html)
// are NOT used by this version.
// ============================================================

// ── Tracker registry ─────────────────────────────────────────
const TRACKERS = {
  "Hair": {
    spreadsheetId: "1DbcV58XCrHQqjqg1Kl6JaEsBHrpBA7jGiqGOB71wj1Y",
    tabProductMap: {
      "Biotin":       ["Biotin30"],
      "S1":           ["Stage 1"],
      "S2":           ["Stage2"],
      "S3":           ["Stage3"],
      "Form Testing": ["Selfasst"],
      "Cetosomal":    ["Advance Regime"]
    }
  },
  "Beard": {
    spreadsheetId: "1lguNY_9CQIOk6Rsm9hsixJ-GTxMgCpCRhnZIAaxuh2c",
    skipColumns: ["canLive", "sno"],
    hideFields:  ["canLive"],
    tabProductMap: {
      "Beard": ["Beard Growth Kit", "Beard Activator Kit", "Beard Development Kit", "Beard Gummies"]
    }
  }
  // "Nutrition": { spreadsheetId: "REPLACE_WITH_ID", tabProductMap: { ... } }
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

// ── Script Properties key helpers ────────────────────────────
// bc1_ = buckets cache (tracker-level: dropdowns, sheetNames, personMap)
// sc1_ = sheet context (tab-level: headerRow, colIndex, formula info)
// ls1_ = live state   (tab-level: lastDataRow, nextSno, totalRows — updated on every submit)
function _spKey(prefix, trackerName, tabName) {
  return prefix + trackerName + (tabName ? "|" + tabName : "");
}

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

// ── Called on page load ───────────────────────────────────────
function getTrackers() {
  const trackerNames = Object.keys(TRACKERS);
  const first = trackerNames[0];
  const data  = selectTracker(first);
  data.trackerNames  = trackerNames;
  data.activeTracker = first;
  return data;
}

// ── Called when user switches tracker ────────────────────────
// Reads from Script Properties; falls back to sheet on cache miss.
function selectTracker(trackerName) {
  const props = PropertiesService.getScriptProperties();
  const key   = _spKey("bc1_", trackerName);
  const raw   = props.getProperty(key);
  if (raw) {
    try {
      const cached = JSON.parse(raw);
      cached.trackerName = trackerName;
      return cached;
    } catch (e) { /* fall through */ }
  }
  return _loadBucketsFromSheet(trackerName);
}

// Reads Buckets sheet, writes to Script Properties, returns result.
function _loadBucketsFromSheet(trackerName) {
  const ss      = getSpreadsheetFor(trackerName);
  const t       = TRACKERS[trackerName];
  const sheetNames = ss.getSheets()
    .map(s => s.getName())
    .filter(name => t.tabProductMap[name] !== undefined);
  const buckets = readBuckets(ss);

  const result = {
    trackerName,
    sheetNames,
    hideFields:    t.hideFields || [],
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

  try {
    PropertiesService.getScriptProperties()
      .setProperty(_spKey("bc1_", trackerName), JSON.stringify(result));
  } catch (e) { /* ignore write errors — will re-read next time */ }

  return result;
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
// Returns cached sheet context + live state; hits the sheet only on cache miss.
function getSheetContext(tabName, trackerName) {
  const props  = PropertiesService.getScriptProperties();
  const scKey  = _spKey("sc1_", trackerName, tabName);
  const lsKey  = _spKey("ls1_", trackerName, tabName);
  const today  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  let sc = null; // static: headerRow, colIndex, adNameFormula, adNameFormulaRow
  let ls = null; // live:   lastDataRow, nextSno, totalRows

  try { sc = JSON.parse(props.getProperty(scKey)); } catch (e) { sc = null; }
  try { ls = JSON.parse(props.getProperty(lsKey)); } catch (e) { ls = null; }

  if (!sc || !ls) {
    // Cache miss — read from sheet
    const ss    = getSpreadsheetFor(trackerName);
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) throw new Error("Sheet \"" + tabName + "\" not found.");

    if (!sc) {
      const detected = detectHeaders(sheet);
      sc = {
        headerRow:        detected.headerRow,
        colIndex:         detected.colIndex,
        adNameFormula:    "",
        adNameFormulaRow: null
      };
      // Scan for ad name formula (rarely changes — cached until manual refresh)
      if (detected.colIndex.adName != null) {
        const lastRow = sheet.getLastRow();
        if (lastRow > detected.headerRow) {
          const formulas = sheet
            .getRange(detected.headerRow + 1, detected.colIndex.adName + 1, lastRow - detected.headerRow, 1)
            .getFormulas();
          for (let i = formulas.length - 1; i >= 0; i--) {
            if (formulas[i][0]) {
              sc.adNameFormula    = formulas[i][0];
              sc.adNameFormulaRow = detected.headerRow + 1 + i;
              break;
            }
          }
        }
      }
      try { props.setProperty(scKey, JSON.stringify(sc)); } catch (e) {}
    }

    if (!ls) {
      ls = _computeLiveState(sheet, sc.colIndex, sc.headerRow);
      try { props.setProperty(lsKey, JSON.stringify(ls)); } catch (e) {}
    }
  }

  return {
    sheetName:       tabName,
    nextSno:         String(ls.nextSno).padStart(5, "0"),
    today,
    totalRows:       ls.totalRows,
    lastDataRow:     ls.lastDataRow,
    adNameFormula:   sc.adNameFormula,
    adNameFormulaRow: sc.adNameFormulaRow,
    colIndex:        sc.colIndex,
    headerRow:       sc.headerRow
  };
}

// Scans the sheet to find lastDataRow, nextSno, totalRows.
function _computeLiveState(sheet, colIndex, headerRow) {
  const lastRow  = sheet.getLastRow();
  const dataRows = lastRow - headerRow;
  let nextSno    = 1;
  let lastDataRow = headerRow;

  if (dataRows > 0) {
    const signalCols = [colIndex.drive916, colIndex.drive45, colIndex.date].filter(c => c != null);
    const scanCols   = colIndex.sno != null ? [...signalCols, colIndex.sno] : [...signalCols];

    if (scanCols.length > 0) {
      const minC = Math.min(...scanCols);
      const maxC = Math.max(...scanCols);
      const data = sheet.getRange(headerRow + 1, minC + 1, dataRows, maxC - minC + 1).getValues();

      for (let i = data.length - 1; i >= 0; i--) {
        if (signalCols.some(c => data[i][c - minC] !== "")) {
          lastDataRow = headerRow + 1 + i;
          break;
        }
      }

      if (colIndex.sno != null && lastDataRow > headerRow) {
        let maxSno = 0;
        const snoOff = colIndex.sno - minC;
        for (let i = 0; i <= lastDataRow - headerRow - 1; i++) {
          const n = parseInt(data[i][snoOff], 10);
          if (!isNaN(n) && n > maxSno) maxSno = n;
        }
        nextSno = maxSno + 1;
      }
    }
  }

  return { nextSno, lastDataRow, totalRows: dataRows };
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

    const lastCol  = sheet.getLastColumn();
    const today    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    const configSkip = (TRACKERS[trackerName].skipColumns || [])
      .map(k => colIndex[k]).filter(c => c != null);
    const sheetSkip  = getProtectedCols(sheet, lastCol);
    const skipCols   = new Set([...configSkip, ...sheetSkip]);

    const put = (row, col, val) => { if (!skipCols.has(col)) set(row, col, val); };
    const rows = [];

    cuts.forEach((cut, i) => {
      const row = new Array(lastCol).fill("");

      put(row, colIndex.sno,             nextSno + i);
      put(row, colIndex.date,            shared.date || today);
      put(row, colIndex.product,         shared.product);
      put(row, colIndex.drive45,         (cut.ratio === "4:5"  || cut.ratio === "Both") ? cut.url : "");
      put(row, colIndex.drive916,        (cut.ratio === "9:16" || cut.ratio === "Both") ? cut.url : "");
      put(row, colIndex.live,            "No");
      put(row, colIndex.canLive,         shared.canLive  || "Yes");
      put(row, colIndex.raisedBy,        shared.raisedBy || "");
      put(row, colIndex.funnel,          cut.funnel);
      put(row, colIndex.intInf,          shared.intInf);
      put(row, colIndex.adType,          shared.adType);
      put(row, colIndex.language,        shared.language);
      put(row, colIndex.person,          shared.person   || "None");
      put(row, colIndex.narrative,       cut.narrative);
      put(row, colIndex.adFormat,        cut.adFormat);
      put(row, colIndex.onboardingMonth, shared.onboardingMonth || "");
      put(row, colIndex.additionalInfo,  shared.additionalInfo  || "");
      put(row, colIndex.instagram,       shared.instagram       || "");
      put(row, colIndex.creatorType,     shared.creatorType     || "");
      put(row, colIndex.ytLinks,         "");
      put(row, colIndex.dateTakenLive,   "");
      put(row, colIndex.ytAdsStatus,     "No");

      rows.push(row);
    });

    writeRows(sheet, firstNewRow, rows, lastCol, skipCols);

    if (colIndex.sno != null && !skipCols.has(colIndex.sno)) {
      try {
        sheet.getRange(firstNewRow, colIndex.sno + 1, rows.length, 1).setNumberFormat("00000");
      } catch (e) { /* protected — skip format */ }
    }

    setHyperlinks(sheet, firstNewRow, rows, colIndex);

    if (colIndex.adName != null && payload.adNameFormulaRow && !skipCols.has(colIndex.adName)) {
      try {
        const srcCell = sheet.getRange(payload.adNameFormulaRow, colIndex.adName + 1);
        srcCell.copyTo(
          sheet.getRange(firstNewRow, colIndex.adName + 1, rows.length, 1),
          SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false
        );
      } catch (e) { /* protected — skip formula copy */ }
    }

    if (colIndex.ytAdName != null && payload.adNameFormulaRow && !skipCols.has(colIndex.ytAdName)) {
      try {
        const srcCell = sheet.getRange(payload.adNameFormulaRow, colIndex.ytAdName + 1);
        if (srcCell.getFormula()) {
          srcCell.copyTo(
            sheet.getRange(firstNewRow, colIndex.ytAdName + 1, rows.length, 1),
            SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false
          );
        }
      } catch (e) { /* protected — skip formula copy */ }
    }

    const newLastDataRow = firstNewRow + rows.length - 1;
    const newNextSno     = nextSno + rows.length;
    const newTotalRows   = payload.totalRows + rows.length;

    // Update live state cache so next tab switch is instant
    try {
      PropertiesService.getScriptProperties().setProperty(
        _spKey("ls1_", trackerName, tabName),
        JSON.stringify({ nextSno: newNextSno, lastDataRow: newLastDataRow, totalRows: newTotalRows })
      );
    } catch (e) { /* ignore cache write failure */ }

    return {
      ok:  true,
      msg: `${rows.length} row${rows.length === 1 ? "" : "s"} added to ${sheet.getName()}.`,
      nextSno:     String(newNextSno).padStart(5, "0"),
      lastDataRow: newLastDataRow,
      totalRows:   newTotalRows
    };

  } catch (e) {
    return { ok: false, msg: e.message + "\n" + e.stack };
  }
}

// ── Cache refresh — called by the Refresh button in the webapp ──
// Clears Script Properties for the given tracker and re-reads everything from the sheet.
// Returns full tracker data (same shape as selectTracker) so the client can update immediately.
function refreshCache(trackerName) {
  const props = PropertiesService.getScriptProperties();
  const t     = TRACKERS[trackerName];
  if (!t) return { ok: false, msg: "Unknown tracker: " + trackerName };

  // Delete all cached keys for this tracker
  props.deleteProperty(_spKey("bc1_", trackerName));
  Object.keys(t.tabProductMap).forEach(function(tabName) {
    props.deleteProperty(_spKey("sc1_", trackerName, tabName));
    props.deleteProperty(_spKey("ls1_", trackerName, tabName));
  });

  // Re-read Buckets from sheet and cache
  const data = _loadBucketsFromSheet(trackerName);

  // Pre-warm sheet context for all tabs (header scan + live state)
  Object.keys(t.tabProductMap).forEach(function(tabName) {
    try { getSheetContext(tabName, trackerName); } catch (e) { /* skip tabs with issues */ }
  });

  return data; // same shape as selectTracker — client can call applyTrackerData directly
}

// Run once manually from Apps Script editor after deploy, or set as a daily trigger.
function refreshAllCaches() {
  Object.keys(TRACKERS).forEach(function(name) {
    try { refreshCache(name); } catch (e) { /* don't let one tracker block others */ }
  });
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
  Object.entries(HEADER_MAP).forEach(([key, text]) => {
    reverseMap[text.toLowerCase()] = key;
  });

  for (let r = 0; r < data.length; r++) {
    const colIndex = {};
    let matches = 0;
    data[r].forEach((cell, c) => {
      const key = reverseMap[String(cell).trim().toLowerCase()];
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

function getProtectedCols(sheet, lastCol) {
  const blocked = new Array(lastCol).fill(false);
  try {
    sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) {
      try {
        if (p.canEdit()) return;
        const r = p.getRange();
        const c1 = r.getColumn() - 1;
        const c2 = c1 + r.getNumColumns() - 1;
        for (let c = Math.max(c1, 0); c <= Math.min(c2, lastCol - 1); c++) blocked[c] = true;
      } catch (pe) {}
    });
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) {
      try {
        if (p.canEdit()) return;
        blocked.fill(true);
        p.getUnprotectedRanges().forEach(function(r) {
          const c1 = r.getColumn() - 1;
          const c2 = c1 + r.getNumColumns() - 1;
          for (let c = Math.max(c1, 0); c <= Math.min(c2, lastCol - 1); c++) blocked[c] = false;
        });
      } catch (pe) {}
    });
  } catch (e) {}
  const result = new Set();
  blocked.forEach(function(b, i) { if (b) result.add(i); });
  return result;
}

function writeRows(sheet, firstNewRow, rows, lastCol, skipCols) {
  let segStart = null;
  for (let c = 0; c <= lastCol; c++) {
    const isSkipped = skipCols.has(c);
    if (!isSkipped && segStart === null) {
      segStart = c;
    } else if ((isSkipped || c === lastCol) && segStart !== null) {
      const segEnd  = c - 1;
      const segData = rows.map(function(row) { return row.slice(segStart, segEnd + 1); });
      sheet.getRange(firstNewRow, segStart + 1, rows.length, segEnd - segStart + 1).setValues(segData);
      segStart = null;
    }
  }
}

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
