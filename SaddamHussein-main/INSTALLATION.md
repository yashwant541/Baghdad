# Bahrain–Iraq Recon Studio: Dataiku installation

1. Create a **Standard Webapp** in the Dataiku project.
2. Copy `webapp.html`, `webapp.css`, `webapp.js`, and `backend.py` into the corresponding WebApp editors.
3. Copy the entire `bahrain_iraq_algorithm` folder into the project Python library. The library name must remain `bahrain_iraq_algorithm`.
4. Add `pandas`, `openpyxl`, and `numpy` to the Dataiku code environment used by the WebApp.
5. Restart the backend and open the WebApp.

## Workflow
- Upload the Trial Balance; the engine detects its header row and builds the summarized pivot plus a group/sub‑group hierarchy tree (split from the BS Mapping column).
- Upload up to three submission workbooks. Every sheet is inspected — header row, currency columns (USD/IQD/BHD/…, including a `'000`/`million` scale) and a Total column are detected automatically, but nothing is read until you tick the sheets you want (and optionally correct the header row) in the sheet picker.
- Extraction reads each selected sheet structurally: bold/indented rows become section and sub‑section headers, "Total …" rows are flagged, and every numeric cell is tagged with its currency, scale and exact cell reference.
- The engine proposes candidate reconciliation rules automatically: Assets↔Assets, sub‑group↔sub‑group, and currency‑for‑currency / Total‑for‑Total, with an inferred sign convention. Nothing is applied until you approve each suggestion in Mapping Studio; add manual rules for anything it missed.
- Run the reconciliation, then review the full Reconciliation Map — every TB group and sub‑group with its approved submission evidence (file, sheet, cell, amount) — before exporting the audit workbook.

## Notes
- Files are stored only in a temporary backend session directory; no managed folder is used.
- The generated TB Pivot is a normal Excel summary sheet, not an Excel PivotTable cache object. It contains the requested hierarchy, currency columns, and summed Adjusted Balance.
- The exported workbook adds an Auto Suggestions sheet and a Reconciliation Map sheet (flattened lineage) alongside the original summary/pivot/detail/submission/results sheets.
- For multi-user production deployment, replace the in-memory `SESSIONS` dictionary with an approved shared session store and apply the organisation's retention and access controls.
