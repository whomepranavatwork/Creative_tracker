# Creative Tracker — Bulk Entry Automation

Google Apps Script that adds a sidebar to your Creative Tracker sheet.  
Paste one Drive folder link → fill shared fields once → set per-cut Funnel / Narrative / Format → done.

---

## One-time Installation

1. Open your **Creative Tracker Google Sheet**
2. Click **Extensions → Apps Script**
3. Delete all existing code in the default `Code.gs` file
4. Paste the contents of `Code.gs` from this repo
5. Click **+ (New file) → HTML**, name it exactly `Sidebar`
6. Paste the contents of `Sidebar.html` from this repo
7. Click **Save** (or Ctrl+S)
8. Close the Apps Script tab and **refresh your sheet**
9. A **"Creative Tracker"** menu will appear in the top menu bar

---

## Workflow (per batch)

1. **Creative Tracker → Add New Entries…** opens the sidebar
2. **Step 1** — Paste the Drive folder link containing the cuts, click **Load**  
   The script lists all files in the folder automatically.
3. **Step 2** — Fill the shared fields once:  
   Product, INT/INF, Ad Type, Language, Person, Creator Type, Instagram, etc.
4. **Step 3** — For each cut, set:
   - **Funnel Type** (BOF / MOF / TOF)
   - **Broad Narrative**
   - **Ad Format**
   - **Ratio** (9:16 or 4:5) — defaults to the value chosen in Step 2
5. Click **Add to Sheet** — all rows are written instantly with correct S.No., Drive links, and all fields.

The **Ad Name formula** is copied automatically from the previous row, so the sheet generates it as usual.

---

## Notes

- The script auto-detects the header row and column positions — no hardcoded row numbers.
- S.No. auto-increments from the highest existing value.
- If the folder has files for both ratios (e.g. a 4:5 and a 9:16 version), load the folder, set each row's ratio individually.
- To lock the script to a specific sheet tab, set `FORCE_SHEET_NAME` at the top of `Code.gs`.
