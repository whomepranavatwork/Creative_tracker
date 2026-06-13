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
| `.clasp.json` | clasp deployment config (script ID) |
| `appsscript.json` | Apps Script manifest (runtime, timezone, scopes) |
| `.github/workflows/deploy-appsscript.yml` | GitHub Actions workflow — auto-deploys to Apps Script on every push to `main` |

---

## One-time Setup

1. Open any Google Sheet → **Extensions → Apps Script**
2. The project does not need to be bound to a specific sheet — it uses `openById` to access all sheets
3. In `MultiTracker_Code.gs`, update the `TRACKERS` registry with the correct `spreadsheetId` for each tracker (see below)
4. **Deploy → New deployment**
   - Type: **Web app**
   - Execute as: **Me** (or the service account that has edit access to all sheets)
   - Who has access: Anyone with Google account (or restrict to org)
5. Copy the URL and share with the team

> The URL is permanent. Pushes to `main` update the project code automatically via GitHub Actions (see below). The **live** web app is also redeployed automatically **only if** the `CLASP_DEPLOYMENT_ID` secret is configured — otherwise the live URL keeps serving the old version until you redeploy manually via **Deploy → Manage deployments → Edit (pencil icon) → New version → Deploy**.

> After the first deploy (or after adding a new tracker), run `refreshAllCaches()` once from the Apps Script editor to pre-warm the Script Properties cache for all trackers and tabs.

> **Apps Script API must be enabled** on the deploying Google account: visit `https://script.google.com/home/usersettings` and toggle **Google Apps Script API → On**. Required for clasp to push.

---

## GitHub Actions Auto-Deploy

Every push to `main` triggers `.github/workflows/deploy-appsscript.yml`, which installs clasp, runs `clasp push --force` (updates the project's editor code), and then — if configured — runs `clasp deploy -i <deployment-id>` to update the **live web app** to the new code while keeping the URL stable.

> **Important:** `clasp push` alone does NOT change what the live URL serves. Without the deployment step, the web app keeps running the previously deployed version.

**One-time secret setup:**

1. On your machine: `npx @google/clasp login` → completes OAuth in browser
2. Copy the full contents of `~/.clasprc.json` (Windows: `%USERPROFILE%\.clasprc.json`)
3. GitHub repo → **Settings → Secrets and variables → Actions → New repository secret**
   - Name: `CLASPRC_JSON`
   - Value: paste the JSON
4. Add a second secret for live redeploys:
   - In the Apps Script editor: **Deploy → Manage deployments** → copy the **Deployment ID** of the web app deployment
   - Name: `CLASP_DEPLOYMENT_ID`
   - Value: the deployment ID
5. Done — all future pushes to `main` update both the code and the live web app

The `refresh_token` in the JSON is long-lived; the secret does not need to be rotated unless you revoke access. If `CLASP_DEPLOYMENT_ID` is missing, the workflow still succeeds but prints a warning that a manual redeploy is needed.

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
    narratives: ["Pure & Natural/Ingredients", "Daily Energy", "Summer",
      "Clinically Studied", "Zero Added Sugar", /* ...full list in code... */ ],
    tabProductMap: {
      "Shilajit":  ["Shilajit Gummies", "Shilajit Gummies Advanced"],
      "Creatine":  ["Creatine Powder", "Creatine Electrolyte"],
      "Magnesium": ["Magnesium Gummies"]
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
| `narratives` | string[] | Optional override of the Broad Narrative dropdown options for this tracker (defaults to the Hair narrative list if omitted) |

To add a new tracker: add an entry to `TRACKERS`, then run `refreshAllCaches()`.

---

## Rules and Logic

### Tracker and Tab Selection

- The nav is a **single grouped view**: every registered tracker's tabs are shown at once, grouped under a tracker label (HAIR / NUTRITION / BEARD). There is no separate "pick a tracker first" step — you click the tab you want directly.
- Tabs come from each tracker's `tabProductMap`. Clicking a tab in a *different* tracker switches context (loads that tracker's dropdowns/people) and then activates the tab; clicking a tab in the *current* tracker just loads its sheet context.
- Cross-tracker loads use a monotonic request counter so rapid clicks always apply only the latest response (no stale tab "snapping back").
- Selecting a tab loads sheet context (entry count, next S.No., live sheet-view snapshot) via a single server call.
- Switching tabs **clears the Drive folder, cut details, and cached sheet context** immediately — no stale data from the previous tab can carry into a submit.
- Switching tabs within the same tracker **preserves all shared fields**; switching to a tab in another tracker resets tracker-scoped data (dropdowns, people) but keeps your shared field entries where still valid.
- The web app guards against browser **back/forward** navigating away from the single-page app (`history.pushState` + `popstate`).

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
- Person Full Name — **searchable combobox**, strict to Buckets names. Click the field (or the ▾ caret) to open the list, type to filter it live, and pick with the mouse or keyboard (↑/↓ to move, Enter to select, Esc to close). It is strict: only names already in the tracker's Buckets tab (plus "None") can be committed — if you type something that isn't a real name and click away, the field reverts. A hidden `#person` element holds the committed value (so submit, localStorage, and Instagram auto-fill read it unchanged). Selecting a name auto-fills Instagram from `personMap`; "None"/empty clears it
- **New person names**: the only way to add a name is the **"+ Add person"** button (top right). It writes the name (and Instagram) to the chosen tracker's **Buckets tab** (columns H/I); if the target is the active tracker, the new name is added to the combobox list and auto-selected. This keeps the sheet's own data-validation dropdowns consistent — free-form names can't be submitted
- Creator Type (incl. Meta Creator, YT Creator, Hair Warrior), Onboarding Month (only required when Person ≠ "None"), Raised By, Additional Info

**Required fields** are marked with a red `*`. Submitting with any required field empty highlights the missing fields and scrolls to the first one.

**localStorage persistence**: INT/INF, Ad Type, Language, Person, Creator Type, Instagram, Raised By, Can Take Live? are saved after each successful submit and restored on next visit.

### Cut Details (Step 3)

Each file gets its own row. Per-cut fields:

- **Funnel Type** — from Buckets tab
- **Ratio** — 9:16, 4:5, or Both. Controls which drive link column receives the URL
- **Broad Narrative** — dropdown + "Other…" for free-text entry
- **Ad Format** — dropdown + "Other…" for free-text entry

**Same for all toggle**: on by default — one shared panel with a file checklist. Toggle off for a per-row table with individual dropdowns.

**File list UI**: each file shows its name (clickable — opens the Drive link in a new tab) and a ▶ play button that opens an inline video preview modal without leaving the page.

### Last-Row Detection

Finding the last real entry (and therefore where to write new rows) is the most critical operation. The logic uses two independent signals and reads at most `DATA_SCAN_LIMIT = 2000` rows from the header downward, regardless of `sheet.getLastRow()` (which can be inflated by stray values or formula pre-fill).

**Signals used:**

1. **Drive link column** — only contains an `https://` URL when a real submission was made; never pre-filled
2. **Ad name column** — the ad name formula (`=LOWER(SUBSTITUTE(TEXTJOIN(...)))`) produces an empty string for blank rows, a short partial string (e.g. `advance_regime_00176`) when only product/S.No. is pre-filled, and a full name with **6 or more underscore-separated segments** only for complete rows

A row is classified as real when:
- Both columns present → **either** signal is sufficient (a drive URL is conclusive on its own even if the ad name formula hasn't fully resolved; the 6-segment threshold still filters out pre-filled partial ad names)
- Only drive columns present (no ad name column) → requires drive URL
- Only ad name column present (no drive columns) → requires full ad name

**S.No. column is NOT used as a signal** — it is formula-pre-filled for all rows in the sheet (including template rows far below the last real entry), so `sheet.getLastRow()` and the S.No. column are both unreliable for boundary detection.

### Concurrent Submit Safety

Every submit acquires a **script-level lock** (`LockService.getScriptLock().tryLock(10000)`) before touching the sheet. Inside the lock, `_computeLiveState` is re-run against the live sheet — the client-sent `lastDataRow` and `nextSno` values are never trusted. This eliminates the race condition where two simultaneous submits could target the same rows.

If the lock cannot be acquired within 10 seconds, the submit returns an error without writing anything.

### Overwrite Guard

Two guards run inside the submit lock before any write:

1. **Drive URL guard** — checks **both** drive link columns (`Google Drive Link (9:16)` and `Google Drive Link (4:5)`) across the target rows. If any cell already contains an `https://` URL, the submit is **aborted**. Placeholder values like "None" or "No" do not trigger it — only real URLs do.
2. **Manual-row guard** — checks the Product Name, Person Full Name, and Raised By cells across the target rows. Any non-empty value means a teammate started that row by hand (e.g. typed details before pasting the drive link) — the submit is **aborted** rather than overwriting their in-progress work.

### Schema Validation at Write Time

Every submit re-validates the cached column map against the live header row: each cached column index must still carry its expected header text. If columns were inserted, moved, or deleted since the cache was built, the schema (and formula row positions) is rebuilt from the sheet and the live-state cache is cleared — values can never silently land in the wrong columns.

### Column Write Safety

Only columns that were explicitly set by the submit are written. The server tracks a `writtenCols` Set and writes only those columns in contiguous batches — helper columns, formula columns, and anything not in the Set are never touched, even if they fall between two written columns.

### Write Verification (Read-Back)

After writing, the server calls `SpreadsheetApp.flush()` (forces pending writes to commit) and then **reads the rows back** to confirm they actually landed:

1. **S.No. anchor** — for each new row it checks the S.No. cell holds the expected number. A missing/wrong S.No. means the row was **not persisted at all** (e.g. a merged cell or row-level protection silently swallowed the write), which is reported distinctly: *"N row(s) were not written to the sheet (check for merged cells or row-level protection on rows X–Y)"*.
2. **Drive-link read-back** — confirms each URL landed in the correct ratio column (4:5 vs 9:16). A mismatch is reported as *"drive links did not read back correctly"*.

On the happy path the banner shows only the clean success message; diagnostics (cuts received, columns written, skipped/protected columns, validation rejections) are appended **only** when verification fails or the sheet rejected a value.

### S.No. and Row Positioning

- **Header row auto-detection**: scans the first 20 rows for a row matching ≥ 6 known column headers from `HEADER_MAP`. No hardcoded row numbers.
- **S.No. value**: read from the sno cell at `lastDataRow`; new rows get `lastSno + 1, lastSno + 2, ...`
- **Format**: zero-padded 5 digits (`00001`), applied via `setNumberFormat("00000")`
- New rows are always written immediately after `lastDataRow`

### Ad Name Formula

- The script independently scans up to `DATA_SCAN_LIMIT` rows of the **Ad Name** column for a formula, and separately scans the **YT Ad Name** column for its formula
- Both formulas are stored at their own row references (`adNameFormulaRow`, `ytAdNameFormulaRow`) — they do not need to be on the same row
- Each formula is copied (with relative reference adjustment) into all newly written rows

### Protected Columns

Before writing, the script checks sheet and range protections. Columns the running account cannot edit are skipped entirely (not written to). Tracker-level `skipColumns` adds additional columns to skip regardless of protection status (e.g. Beard's S.No. and Can Take Live? are auto-managed by sheet rules).

Hyperlink writing also respects `skipColumns` — protected drive columns are not touched by the hyperlink step.

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
| YT Ad Name | Formula copied from existing row (independently discovered) |

Drive link cells are written as **rich text hyperlinks** (URL is both display text and link target).

---

## Sheet Maintenance Standards

The detection and write logic relies on a few conventions. Keep these true and the app keeps working:

- **Header row stays at row 1** with the exact header text in `HEADER_MAP`. Don't insert rows above it. Renaming a header? The cache auto-rebuilds on the next load, or clear `sc1_<tracker>|<tab>` from Script Properties to force it.
- **Tab names must match the keys in `tabProductMap`.** If you rename a sheet tab, update the matching key in `TRACKERS` (in `MultiTracker_Code.gs`) — otherwise that tab won't appear in the nav / product mapping.
- **Leave every cell blank below the last real entry.** The boundary detector finds the last row using the **Drive link** column and the **Ad Name formula** (a full ad name has 6+ underscore segments). It does **not** trust S.No. or `sheet.getLastRow()`. So:
  - **Never pre-fill the Ad Name formula into empty template rows** — every such row then looks "real" and new entries write in the wrong place. The form fills the formula itself on submit.
  - Don't pre-fill S.No. in blank rows (harmless for detection, but redundant — the form writes it).
- **Don't replace the Ad Name formula with a static value.** A hardcoded string won't update; the form copies the live formula down. If a formula cell was manually cleared, the form re-fills it on the next write to that row.
- **Buckets tab is the source of dropdowns/people** — don't delete or rename it. After editing Buckets, hit **Refresh dropdowns** (or run `refreshAllCaches()`).
- **Protection / merged-cell changes** can silently block writes (see Write Verification). After changing protected ranges, refresh the cache so `protectedCols` is rebuilt.

---

## Caching (Script Properties)

All sheet reads are cached in **Script Properties** to keep the web app fast after the first load. Three cache tiers:

| Prefix | Scope | Contents |
|--------|-------|----------|
| `bc1_<tracker>` | Tracker-level | Dropdown options, sheetNames, tabProductMap (everything except person data) |
| `bc1p_<tracker>` | Tracker-level | personMap + person list — split from `bc1_` to stay under the 9 KB per-property limit. Treated as atomic with `bc1_`: if either key is missing, both are rebuilt from the sheet |
| `sc1_<tracker>\|<tab>` | Tab-level | headerRow, colIndex, adNameFormulaRow, ytAdNameFormulaRow, adName formula — re-validated against the live header row on every submit |

**Live state (entry count, last row, next S.No.) is never cached.** It is recomputed from the sheet on every tab load and again inside every submit lock — so the count shown in the app always reflects rows added manually or from other sessions, and writes always target the true next row.

**Cache invalidation:**
- The **Refresh dropdowns** button in the web app calls `refreshCache(trackerName, activeTab)` — clears all keys for the current tracker and pre-warms the active tab (other tabs rebuild lazily on first use)
- Run `refreshAllCaches()` from the Apps Script editor to clear and re-warm all trackers and all tabs at once (use after deploy or when sheet structure changes)
- A schema mismatch detected at submit time auto-rebuilds `sc1_` for that tab — no manual action needed

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
const TRACKERS = { ... };               // tracker registry — see above
const DEPLOY_STAMP = "11 Jun 2026, ..."; // shown in the topbar; bump MANUALLY when cutting a deploy
const HEADER_SEARCH_LIMIT = 20;         // rows to scan when locating the header row
const DATA_SCAN_LIMIT     = 2000;       // max rows to scan for last real entry (~1 year headroom)
```

> `DEPLOY_STAMP` is a **manual** string (an earlier attempt to auto-inject it in CI was reverted). Update it by hand when you want the topbar's "deployed …" label to reflect a notable release.

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
