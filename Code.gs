// ============================================================
// Creative Tracker — Bulk Entry Automation
// Google Apps Script (bound to the tracker sheet)
// ============================================================
// HOW TO INSTALL:
//   1. Open your Creative Tracker Google Sheet
//   2. Extensions → Apps Script
//   3. Delete any existing code, paste this file's contents
//   4. Paste Sidebar.html as a second file named "Sidebar"
//   5. Save, then refresh the sheet
//   6. A new menu "Creative Tracker" will appear
// ============================================================

const SHEET_NAME = "Cetosomal"; // Change to match your exact tab name
const HEADER_ROW = 9;           // Row where column headers live (S.No., Date Added…)
const DATA_START_ROW = 10;      // First data row (adjust if your sheet differs)

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
    .setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── Column map (0-indexed from column A = 0) ─────────────────
// Matches the CSV: S.No.(A=0), Date(B=1), Product(C=2), AdName(D=3),
// Drive4_5(E=4), Drive9_16(F=5), Live(G=6), CanLive(H=7), RaisedBy(I=8),
// FunnelType(J=9), IntInf(K=10), AdType(L=11), Language(M=12),
// PersonName(N=13), Narrative(O=14), AdFormat(P=15), OnboardingMonth(Q=16),
// AdditionalInfo(R=17), InstagramProfile(S=18), CreatorType(T=19),
// YTLinks(U=20), DateTakenLive(V=21), YTAdsStatus(W=22), YTAdName(X=23)
const COL = {
  sno: 0, date: 1, product: 2, adName: 3,
  drive45: 4, drive916: 5, live: 6, canLive: 7, raisedBy: 8,
  funnel: 9, intInf: 10, adType: 11, language: 12,
  person: 13, narrative: 14, adFormat: 15, onboardingMonth: 16,
  additionalInfo: 17, instagram: 18, creatorType: 19,
  ytLinks: 20, dateTakenLive: 21, ytAdsStatus: 22, ytAdName: 23
};
const TOTAL_COLS = 24;

// ── Main entry point called from the sidebar ──────────────────
function addEntries(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME)
      || ss.getSheets()[0]; // fallback to first sheet

    const links45  = splitLinks(formData.links45  || "");
    const links916 = splitLinks(formData.links916 || "");
    const maxRows  = Math.max(links45.length, links916.length);

    if (maxRows === 0) return { ok: false, msg: "Please provide at least one Drive link." };

    const nextSno = getNextSerialNumber(sheet);

    const today = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"
    );

    const rows = [];
    for (let i = 0; i < maxRows; i++) {
      const sno      = String(nextSno + i).padStart(5, "0");
      const link45   = links45[i]  || "";
      const link916  = links916[i] || "";
      const adName   = buildAdName(formData, sno);
      const ytAdName = buildYtAdName(formData, sno);

      const row = new Array(TOTAL_COLS).fill("");
      row[COL.sno]             = sno;
      row[COL.date]            = formData.date || today;
      row[COL.product]         = formData.product;
      row[COL.adName]          = adName;
      row[COL.drive45]         = link45;
      row[COL.drive916]        = link916;
      row[COL.live]            = formData.live || "No";
      row[COL.canLive]         = formData.canLive || "Yes";
      row[COL.raisedBy]        = formData.raisedBy || "Pranav";
      row[COL.funnel]          = formData.funnel;
      row[COL.intInf]          = formData.intInf;
      row[COL.adType]          = formData.adType;
      row[COL.language]        = formData.language;
      row[COL.person]          = formData.person || "None";
      row[COL.narrative]       = formData.narrative;
      row[COL.adFormat]        = formData.adFormat;
      row[COL.onboardingMonth] = formData.onboardingMonth || "None";
      row[COL.additionalInfo]  = formData.additionalInfo || "";
      row[COL.instagram]       = formData.instagram || "";
      row[COL.creatorType]     = formData.creatorType || "";
      row[COL.ytLinks]         = formData.ytLinks || "";
      row[COL.dateTakenLive]   = "";
      row[COL.ytAdsStatus]     = "No";
      row[COL.ytAdName]        = ytAdName;

      rows.push(row);
    }

    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rows.length, TOTAL_COLS).setValues(rows);

    return {
      ok: true,
      msg: `✅ Added ${rows.length} entr${rows.length === 1 ? "y" : "ies"} (S.No. ${String(nextSno).padStart(5,"0")} – ${String(nextSno + rows.length - 1).padStart(5,"0")}).`
    };
  } catch (e) {
    return { ok: false, msg: "Error: " + e.message };
  }
}

// ── Helpers ───────────────────────────────────────────────────

function getNextSerialNumber(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return 1;

  // Walk up from the last row to find the highest numeric S.No.
  const snoValues = sheet
    .getRange(DATA_START_ROW, COL.sno + 1, lastRow - DATA_START_ROW + 1, 1)
    .getValues();

  let max = 0;
  for (const [val] of snoValues) {
    const n = parseInt(val, 10);
    if (!isNaN(n) && n > max) max = n;
  }
  return max + 1;
}

function slugify(str) {
  if (!str) return "none";
  return str
    .toLowerCase()
    .trim()
    .replace(/\s+/g, "_")       // spaces → underscores
    .replace(/['']/g, "")       // remove apostrophes
    .replace(/[^a-z0-9_/\-]/g, ""); // keep letters, digits, _, /, -
}

function buildAdName(f, sno) {
  const parts = [
    slugify(f.product),
    slugify(f.funnel),
    slugify(f.intInf),
    slugify(f.adType),
    slugify(f.language),
    slugify(f.person || "none"),
    slugify(f.narrative),
    slugify(f.adFormat),
    sno
  ];
  return parts.join("_");
}

function buildYtAdName(f, sno) {
  // Base same as Ad Name
  const base = buildAdName(f, sno);
  // Append creator type slug if present and not Inf
  const ct = (f.creatorType || "").trim();
  if (ct && ct.toLowerCase() !== "inf") {
    return base.replace(`_${sno}`, "") + "_" + slugify(ct);
  }
  return base;
}

function splitLinks(text) {
  return text
    .split(/[\n,]+/)
    .map(s => s.trim())
    .filter(Boolean);
}

// ── Utility: expose dropdown options to the sidebar ───────────
function getDropdownOptions() {
  return {
    funnel:    ["BOF", "MOF", "TOF"],
    intInf:    ["Inf", "Int", "Hair Warriors"],
    adType:    ["Reel", "Static", "Carousel"],
    language:  ["Hinglish", "English", "Hindi"],
    adFormat:  ["Educational", "Non Vo", "Statics", "Testimonial", "Carousel",
                "UGC", "Organic Reviews", "Hair Warriors"],
    narrative: ["Problem-Solution", "Results/Assurance", "Myth Busting", "Trust",
                "Features", "Why MM is different", "Safety", "Science",
                "BeforeAfter", "Transplant or PRP", "Product First",
                "Results/Assurance/ProductFirst"],
    creatorType: ["", "Meta Creator", "Hair Warrior"],
    live:      ["No", "Yes"],
    canLive:   ["Yes", "No"]
  };
}
