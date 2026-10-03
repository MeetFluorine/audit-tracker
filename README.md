# PV Inventory Recon Tracker

Static site. Compares a Specific Substore Stock Report (Excel) with an Airtel i360 PV audit preview (PDF) in the browser. No backend; files never leave the user's machine.

## Files
- `index.html` page markup
- `css/styles.css` styles
- `js/core.js` parsing, normalisation, matching
- `js/ui.js` upload, dashboard, filters, table
- `js/export.js` Excel export
- SheetJS and PDF.js load from cdnjs.

## Deploy on GitHub Pages
1. Push these files to the repository root.
2. Settings > Pages > Deploy from branch > `main` / root.
3. Optional: add a custom domain under Pages settings.
