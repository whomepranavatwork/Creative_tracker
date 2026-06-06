# Creative Tracker — Bulk Entry Automation

Google Apps Script + standalone web app that bulk-writes creative tracker rows from a Drive folder link.  
Paste a Drive folder → fill shared fields once → set per-cut details → submit. All rows written instantly.

---

## Files

| File | Purpose |
|------|---------|
| `Code.gs` | Server-side Apps Script: reads the sheet, writes rows, serves the web app |
| `webapp.html` | Standalone full-page web app (share this URL with the team) |
| `Sidebar.html` | Sidebar version (opens inside Google Sheets) |
| `.clasp.json` | clasp deployment config (script ID) |

---

## One-time Setup

### Option A — Web App (recommended, works on any device)

1. Open the target Google Sheet → **Extensions → Apps Script**
2. Create / replace three files: `Code.gs`, `webapp.html`, `Sidebar.html` with contents from this repo
3. In `Code.gs`, set `SPREADSHEET_ID` to the Sheet's ID (from its URL: `.../spreadsheets/d/<ID>/edit`)
4. Update `TAB_PRODUCT_MAP` in `webapp.html` to match the Sheet's tab names and products (see below)
5. **Deploy → New deployment**
   - Type: **Web app**
   - Execute as: **User accessing the web app** ← important for edit history attribution
   - Who has access: Anyone with Google account (or restrict to org)
6. Copy the URL and share with the team

> The URL is permanent. Subsequent code changes are deployed via **Deploy → Manage deployments → Edit (pencil icon) → Deploy** — same URL, no resharing needed.

### Option B — Sidebar (inside Google Sheets only)

Steps 1–4 same as above. No deployment needed — the sidebar opens via the **Bias for Action!** menu that appears when the Sheet is opened.

---

## Migrating to a Different Sheet

1. Copy all three files into the new Sheet's Apps Script project
2. Change `SPREADSHEET_ID` in `Code.gs` to the new Sheet's ID
3. Update `TAB_PRODUCT_MAP` in `webapp.html` to match the new Sheet's tabs and products
4. If the new Sheet has different column headers, update `HEADER_MAP` in `Code.gs` to match
5. Deploy as a new Web app deployment (new URL) — follow Option A steps 5–6

---

## Rules and Logic Currently Active

### Tab Selection

- Tabs are shown as **pill buttons** at the top of the form (not a dropdown)
- Only the first 7 sheet tabs are shown (configurable via the `slice(0,7)` call in `init()`)
- Selecting a tab loads sheet context (next S.No., entry count) via a single server call
- Switching tabs **clears the Drive folder and cut details** but **preserves all shared fields** (Step 2) — so if a user picks the wrong tab after filling the form, they can switch without losing their work

### Product Name (Tab-Driven)

Product options are **hardcoded per tab** in `TAB_PRODUCT_MAP` — they do not read from the Buckets tab. This is intentional: the Buckets tab may list all products across all trackers; we only want the relevant ones per tab.

Current mapping:

| Tab | Product Options |
|-----|----------------|
| BGK | Beard Growth Kit, Beard Activator Kit, Beard Development Kit, Beard Gummies |
| Biotin | Biotin30 (auto-selected) |
| S1 | Stage 1 (auto-selected) |
| S2 | Stage2 (auto-selected) |
| S3 | Stage3 (auto-selected) |
| Form Testing | Selfasst (auto-selected) |
| Cetosomal | Advance Regime (auto-selected) |

- **Single product tabs**: product is auto-selected, no dropdown shown
- **Multi-product tabs (BGK)**: dropdown with a "— select product —" placeholder
- **Unknown tabs**: falls back to the full Buckets product list

To update: edit `TAB_PRODUCT_MAP` in `webapp.html`.

### Drive Folder Loading

- Paste a Drive folder URL → folder loads automatically (no need to click Load)
- Typing/dragging a full `drive.google.com/drive/folders/` URL also triggers auto-load
- Files are sorted alphabetically (numeric-aware sort)
- The folder must be shared with the script runner's Google account

### Shared Fields (Step 2)

These fields apply to every file in the batch:

- Product Name, Date Added (defaults to today), INT/INF, Ad Type, Language, Can Take Live?
- Person Full Name → Instagram auto-fills from the Buckets tab `personMap`
- Creator Type, Onboarding Month, Raised By (defaults to "Pranav"), Additional Info

**localStorage persistence**: INT/INF, Ad Type, Language, Person, Creator Type, Instagram, Raised By are saved to `localStorage` after each successful submit and restored on next visit. Product is excluded (tab-driven).

All dropdown options except Product and the narrative/format fields are read live from the **Buckets tab** on page load. Fallback defaults are built-in if a column is empty.

### Cut Details (Step 3)

Each file in the folder gets its own row. Per-cut fields:

- **Funnel Type** (from Buckets tab: BOF, MOF, TOF or similar)
- **Ratio** — 9:16, 4:5, or Both. Drives which Drive link column gets the URL:
  - `9:16` → writes to `Google Drive Link (9:16)` column only
  - `4:5` → writes to `Google Drive Link (4:5)` column only
  - `Both` → writes the same URL to both columns
- **Broad Narrative** — dropdown + "Other…" option for custom values
- **Ad Format** — dropdown + "Other…" option for custom values

**Same for all toggle**: on by default. Shows a shared panel with a file checklist. When toggled off, shows a per-row table where each file has individual dropdowns.

### Overwrite Guard

Before writing, the script checks the target rows in the **Drive link column** (`Google Drive Link (9:16)` preferred, falls back to `Google Drive Link (4:5)`). If any of those cells already has a value, the submit is **aborted** with an error message. This prevents accidental overwrites of existing entries.

Drive link columns are used as the guard because they are never pre-filled — product name, S.No., and funnel type may be pre-filled by other workflows.

### S.No. and Row Positioning

- The script **auto-detects the header row** by scanning the first 20 rows for a row matching at least 6 known column headers. No hardcoded row numbers.
- `lastDataRow` is determined by scanning the **Funnel Type column** from the bottom up — the last row with a non-empty Funnel value is treated as the last real entry. (Date Added is not used for this because it may have been written incorrectly by earlier runs.)
- S.No. is read from that last row and incremented. It is formatted as a zero-padded 5-digit number (`00001`).
- New rows are always written immediately after `lastDataRow`.

### Ad Name Formula

- The script scans the entire Ad Name column for the last row containing a formula
- That formula is copied (with relative reference adjustment) into all newly written rows
- `YT Ad Name` formula is copied the same way if present
- Ad Name and YT Ad Name cells are **left blank for value writing** — only the formula copy fills them

### Edit History Attribution

The web app is deployed as **"Execute as: User accessing the web app"**. This means Google Sheet edit history correctly attributes writes to the individual user who submitted the form, not the script owner.

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
| Can Take Live? | From field (default: Yes) |
| Raised By | From field (default: Pranav) |
| Funnel Type | Per-cut field |
| INT / INF | From shared field |
| Ad Type | From shared field |
| Language | From shared field |
| Person Full Name | From shared field |
| Broad Narrative - P0 | Per-cut field |
| Ad Format | Per-cut field |
| Inf Onboarding Month | From shared field |
| Additional Information | From shared field |
| Instagram Profile | Auto-filled from personMap or manually entered |
| Creator Type | From shared field |
| YT Links | Written blank |
| Date Taken Live | Written blank |
| YT Ads Status | Hardcoded "No" |
| YT Ad Name | Formula copied from existing row |

Drive link cells are written as **rich text hyperlinks** (the URL is both the display text and the link target).

---

## Dropdowns That Come from the Buckets Tab

The following options are read live from the **Buckets** sheet on page load:

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

If any column is empty, built-in fallback defaults are used:
- Funnel: BOF, MOF, TOF
- INT/INF: Inf, Int, Hair Warriors
- Ad Type: Reel, Static, Carousel
- Language: Hinglish, English, Hindi
- Ad Format: Educational, Non Vo, Statics, Testimonial, Carousel, UGC, Organic Reviews, Hair Warriors

Narrative and Can Take Live / ratio / creatorType / onboardingMonth options are hardcoded in `Code.gs` (they don't vary by Buckets data).

---

## Code.gs Key Constants

```js
const SPREADSHEET_ID   = "1EknvJXuzSZTNKnT3Xr3jsAbPct9yEgCzLbbehwZCJ2E"; // change per sheet
const FORCE_SHEET_NAME = "";       // leave blank — user picks tab in the form
const HEADER_SEARCH_LIMIT = 20;    // rows to scan when looking for the header row
```

`HEADER_MAP` maps internal field keys to exact column header strings as they appear in the Sheet. If column headers differ in a new sheet, update this map.
