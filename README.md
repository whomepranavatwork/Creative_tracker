# Creative Tracker — Multi-Tracker Bulk Entry

Google Apps Script + standalone web app for bulk-writing creative tracker rows from a Drive folder link.  
Paste a Drive folder → fill shared fields once → set per-cut details → submit. All rows written instantly.

Supports multiple trackers (Hair, Beard, Nutrition) each backed by a separate Google Sheet, selectable from the same web app URL.

---

## Files

| File | Purpose |
|------|---------|
| `MultiTracker_Code.gs` | Server-side Apps Script: reads sheets, writes rows, serves web app, manages cache |
| `MultiTracker_webapp.html` | Standalone full-page web app (share this URL with the team) |
| `Code.gs` | Original single-tracker server code (not used by multi-tracker deployment) |
| `webapp.html` | Original single-tracker web app (not used by multi-tracker deployment) |
| `Sidebar.html` | Sidebar version for use inside Google Sheets (not used by multi-tracker deployment) |
| `.clasp.json` | clasp deployment config (script ID) |

---

## One-time Setup

1. Open any Google Sheet → **Extensions → Apps Script**
2. Create two files: `MultiTracker_Code.gs` and `MultiTracker_webapp.html` with contents from this repo  
   (The Apps Script project does not need to be bound to a specific sheet — it uses `openById` to access all sheets)
3. In `MultiTracker_Code.gs`, update the `TRACKERS` registry with the correct `spreadsheetId` for each tracker (see below)
4. **Deploy → New deployment**
   - Type: **Web app**
   - Execute as: **Me** (or the service account that has edit access to all sheets)
   - Who has access: Anyone with Google account (or restrict to org)
5. Copy the URL and share with the team

> The URL is permanent. Subsequent code changes are deployed via **Deploy → Manage deployments → Edit (pencil icon) → Deploy** — same URL, no resharing needed.

> After the first deploy (or after adding a new tracker), run `refreshAllCaches()` once from the Apps Script editor to pre-warm the Script Properties cache for all trackers and tabs.

---

## Tracker Registry

All tracker configuration lives in the `TRACKERS` constant at the top of `MultiTracker_Code.gs`:

```js
const TRACKERS = {
  "Hair": {
    spreadsheetId: "...",
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
    spreadsheetId: "...",
    skipColumns: ["canLive", "sno"],
    hideFields:  ["canLive"],
    tabProductMap: {
      "Beard": ["Beard Growth Kit", "Beard Activator Kit", "Beard Development Kit", "Beard Gummies"]
    }
  },
  "Nutrition": {
    spreadsheetId: "...",
    tabProductMap: {
      "Shilajit":  ["Shilajit Gummies", "Creatine Powder"],
      "Creatine":  ["Creatine Powder", "Creatine Electrolyte"],
      "Magnesium": ["Magnesium Gummies", "Creatine Powder"]
    }
  }
};
```

**Per-tracker config options:**

| Key | Type | Description |
|-----|------|-------------|
| `spreadsheetId` | string | ID from the sheet URL: `.../spreadsheets/d/<ID>/edit` |
| `tabProductMap` | object | Maps sheet tab names → allowed product options for that tab |
| `skipColumns` | string[] | Column keys (from `HEADER_MAP`) whose cells are never written (e.g. auto-filled or locked) |
| `hideFields` | string[] | Column keys whose form fields are hidden in the UI |

To add a new tracker: add an entry to `TRACKERS`, then run `refreshAllCaches()`.

---

## Rules and Logic

### Tracker and Tab Selection

- **Tracker pills** appear at the top if more than one tracker is registered; clicking one switches context
- **Tab pills** show the tabs defined in `tabProductMap` for the active tracker
- Selecting a tab loads sheet context (entry count, next S.No.) via a single server call
- Switching tabs **clears the Drive folder and cut details** but **preserves all shared fields** — no rework needed if you pick the wrong tab
- Switching trackers resets everything including shared fields

### Product Name (Tab-Driven)

Product options are **set per tab** in `tabProductMap`. They do not read from the Buckets tab.

- **Single product tab**: product is auto-selected, no dropdown shown
- **Multi-product tab**: dropdown with a "— select product —" placeholder
- **Unmapped tab**: shows the full Buckets product list with a warning banner

### Drive Folder Loading

- Paste a Drive folder or individual file URL → loads automatically (no need to click Load)
- Typing or pasting a full `drive.google.com/drive/folders/` or `drive.google.com/file/d/` URL also triggers auto-load
- Files are sorted alphabetically (numeric-aware)
- Batches over 20 files show a warning banner
- The folder must be shared with the Google account executing the script

### Shared Fields (Step 2)

These apply to every file in the batch:

- Product Name, Date Added (defaults to today), INT/INF, Ad Type, Language, Can Take Live?
- Person Full Name → Instagram auto-fills from `personMap` (Buckets tab column I); clears when switched back to "None"
- Creator Type, Onboarding Month (only required when Person ≠ "None"), Raised By, Additional Info

**localStorage persistence**: INT/INF, Ad Type, Language, Person, Creator Type, Instagram, Raised By, Can Take Live? are saved after each successful submit and restored on next visit.

### Cut Details (Step 3)

Each file gets its own row. Per-cut fields:

- **Funnel Type** — from Buckets tab
- **Ratio** — 9:16, 4:5, or Both. Controls which drive link column receives the URL
- **Broad Narrative** — dropdown + "Other…" for free-text entry
- **Ad Format** — dropdown + "Other…" for free-text entry

**Same for all toggle**: on by default — one shared panel with a file checklist. Toggle off for a per-row table with individual dropdowns.

### Last-Row Detection

Finding the last real entry (and therefore where to write new rows) is the most critical operation. The logic uses two independent signals and reads at most `DATA_SCAN_LIMIT = 2000` rows from the header downward, regardless of `sheet.getLastRow()` (which can be inflated by stray values or formula pre-fill).

**Signals used (both must be present when both columns exist):**

1. **Drive link column** — only contains an `https://` URL when a real submission was made; never pre-filled
2. **Ad name column** — the ad name formula (`=LOWER(SUBSTITUTE(TEXTJOIN(...)))`) produces an empty string for blank rows, a short partial string (e.g. `advance_regime_00176`) when only product/S.No. is pre-filled, and a full name with 6+ underscore-separated segments only for complete rows

A row is classified as real only when:
- Both signals are present → requires **both** a drive URL and a full ad name
- Only drive columns present (no ad name column) → requires drive URL only
- Only ad name column present (no drive columns) → requires full ad name only

**S.No. column is NOT used as a signal** — it is formula-pre-filled for all rows in the sheet (including template rows far below the last real entry), so `sheet.getLastRow()` and the S.No. column are both unreliable for boundary detection.

### Overwrite Guard

Before writing, the script checks the target rows in the drive link column (`Google Drive Link (9:16)` preferred, falls back to `Google Drive Link (4:5)`). If any cell already contains an `https://` URL, the submit is **aborted** with an error. Placeholder values like "None" or "No" do not trigger the guard — only real URLs do.

### S.No. and Row Positioning

- **Header row auto-detection**: scans the first 20 rows for a row matching ≥ 6 known column headers from `HEADER_MAP`. No hardcoded row numbers.
- **S.No. value**: read from the sno cell at `lastDataRow`; new rows get `lastSno + 1, lastSno + 2, ...`
- **Format**: zero-padded 5 digits (`00001`), applied via `setNumberFormat("00000")`
- New rows are always written immediately after `lastDataRow`

### Ad Name Formula

- The script scans up to `DATA_SCAN_LIMIT` rows of the Ad Name column looking for a formula
- That formula is copied (with relative reference adjustment) into all newly written rows
- `YT Ad Name` formula is copied the same way if present

### Protected Columns

Before writing, the script checks sheet and range protections. Columns the running account cannot edit are skipped entirely (not written to). Tracker-level `skipColumns` adds additional columns to skip regardless of protection status (e.g. Beard's S.No. and Can Take Live? are auto-managed by sheet rules).

### Columns Written Per Row

| Column | Value |
|--------|-------|
| S.No. | Auto-incremented, zero-padded 5-digit |
| Date Added | From date field (defaults to today, dd/MM/yyyy) |
| Product Name | From tab-driven product field |
| Ad Name | Formula copied from existing row |
| Google Drive Link (4:5) | File URL if ratio is 4:5 or Both |
| Google Drive Link (9:16) | File URL if ratio is 9:16 or Both |
| Live? | Hardcoded "No" |
| Can Take Live? | From field (default: Yes) — hidden/skipped for Beard |
| Raised By | From field |
| Funnel Type | Per-cut field |
| INT / INF | From shared field |
| Ad Type | From shared field |
| Language | From shared field |
| Person Full Name | From shared field |
| Broad Narrative - P0 | Per-cut field |
| Ad Format | Per-cut field |
| Inf Onboarding Month | From shared field (blank for Int content) |
| Additional Information | From shared field |
| Instagram Profile | Auto-filled from personMap or manually entered |
| Creator Type | From shared field |
| YT Links | Written blank |
| Date Taken Live | Written blank |
| YT Ads Status | Hardcoded "No" |
| YT Ad Name | Formula copied from existing row |

Drive link cells are written as **rich text hyperlinks** (URL is both display text and link target).

---

## Caching (Script Properties)

All sheet reads are cached in **Script Properties** to keep the web app fast after the first load. Three cache tiers:

| Prefix | Scope | Contents |
|--------|-------|----------|
| `bc1_<tracker>` | Tracker-level | Dropdown options, sheetNames, tabProductMap, personMap |
| `sc1_<tracker>\|<tab>` | Tab-level | headerRow, colIndex, adName formula and formula row |
| `ls1_<tracker>\|<tab>` | Tab-level | lastDataRow, nextSno, totalRows — updated after every submit |

**Cache invalidation:**
- The **Refresh** button in the web app calls `refreshCache(trackerName)` — clears and re-reads all keys for the current tracker
- Run `refreshAllCaches()` from the Apps Script editor to clear and re-warm all trackers and tabs at once (use after deploy or when sheet structure changes)

---

## Dropdowns from the Buckets Tab

Each tracker's sheet must have a **Buckets** tab. The following options are read from it:

| Field | Buckets Column |
|-------|---------------|
| Product (fallback only) | A |
| Funnel Type | B |
| INT / INF | C |
| Ad Type | D |
| Language | E |
| Ad Format | F |
| Person Full Name | H |
| Instagram Profile | I (linked to Person via personMap) |

Built-in fallback defaults apply if a column is empty:
- Funnel: BOF, MOF, TOF
- INT/INF: Inf, Int, Hair Warriors
- Ad Type: Reel, Static, Carousel
- Language: Hinglish, English, Hindi
- Ad Format: Educational, Non Vo, Statics, Testimonial, Carousel, UGC, Organic Reviews, Hair Warriors

Narrative, ratio, creatorType, onboardingMonth, canLive options are hardcoded in `MultiTracker_Code.gs`.

---

## Key Constants in MultiTracker_Code.gs

```js
const TRACKERS = { ... };         // tracker registry — see above
const HEADER_SEARCH_LIMIT = 20;   // rows to scan when locating the header row
const DATA_SCAN_LIMIT     = 2000; // max rows to scan for last real entry (~1 year headroom)
```

`HEADER_MAP` maps internal field keys to exact column header strings as they appear in each sheet. Update it if a sheet uses non-standard header text.

---

## Diagnostics

`debugLiveState(trackerName, tabName)` can be run from the Apps Script editor to inspect what the row-detection logic sees on any tab:

```js
debugLiveState("Nutrition", "Creatine")
// Logs: headerRow, colIndex, sheet.getLastRow(), scan range,
//       first/last 5 rows with drive/adName/sno values,
//       and the final result (lastDataRow, nextSno, totalRows)
```

Run this whenever entries or S.No. counts look wrong. If the result is correct but the web app shows stale numbers, run `refreshAllCaches()` to clear the Script Properties cache.
