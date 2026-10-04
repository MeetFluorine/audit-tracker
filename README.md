<div align="center">

# 📊 PV Inventory Recon Tracker

**Find physical-verification variances in seconds, right in your browser.**

Compare a *Specific Substore Stock Report* (Excel) with an *Airtel i360 PV audit preview* (PDF), see exactly what is missing, extra or different, and download a clean, formatted Excel report.

[![Live Demo](https://img.shields.io/badge/🚀_Live_Demo-Open_App-7030A0?style=for-the-badge)](https://meetfluorine.github.io/audit-tracker/)

![Static site](https://img.shields.io/badge/Static_Site-No_Backend-1d8a4e?style=flat-square)
![Privacy](https://img.shields.io/badge/Privacy-Files_never_leave_your_machine-2563eb?style=flat-square)
![Hosting](https://img.shields.io/badge/Hosted_on-GitHub_Pages-111827?style=flat-square)

### 🔗 https://meetfluorine.github.io/audit-tracker/

</div>

---

## ✨ What it does

| Step | What happens |
|---|---|
| **1. Upload** | Drop the stock report (`.xlsx`) and the audit preview (`.pdf`). Columns and substore are detected automatically. |
| **2. Match** | Serialised items are matched by serial number, non-serialised items by item code and quantity. |
| **3. Review** | A dashboard shows missing, extra and quantity differences, with tabs, filters, search and sorting. |
| **4. Download** | One click exports a formatted Excel report, ready to share. |

## 🧩 Highlights

- 🔍 **Serial and quantity matching** in one run, with item-code mismatches, duplicate serials and unreadable rows flagged separately
- 🗂️ **Review tabs:** Needs attention, Missing in audit, Extra in audit, Quantity difference, Other issues, Matched, All records
- 🎛️ **Filters and search** by type, item code, serial number, quality, inventory status and free text
- 🔒 **Private by design:** everything runs client-side, nothing is uploaded anywhere
- 📱 **Responsive** layout that works on desktop and mobile

## 📥 The Excel report

The downloaded workbook is styled and ready to use:

| Sheet | Contents |
|---|---|
| **Summary** | Purple title banner, substore and generated time, then one row per item code: `Item_Code`, `Item_Description`, `Type`, `System_Stock`, `Scanned_in_Audit`, `Matched`, `Missing`, `Extra`, `Result` |
| **Overall Variance** | Every variance record in one place |
| **Missing Serialised** | Serial numbers in system stock but not found in audit |
| **Missing Non-Serialised** | Item codes where audit quantity is below system stock |
| **Extra in Audit** | Serials not in stock, and item codes where audit is above stock |
| **Other Issues** | Code mismatch, duplicate serial, unreadable row (only when present) |

Every table is a real Excel Table (banded rows, filter buttons, centred values). Filename format: `PV_Inventory_Variance_<SUBSTORE>_<YYYY-MM-DD>.xlsx`

## 📁 Project structure

```
pv-recon-tracker/
├── index.html        page markup
├── css/styles.css    styles
└── js/
    ├── core.js       parsing, normalisation, matching engine
    ├── ui.js         upload, dashboard, filters, table
    └── export.js     styled Excel export
```

## 🛠️ Built with

- **Vanilla JavaScript**, HTML and CSS (no build step)
- [SheetJS](https://sheetjs.com/) reads the stock Excel
- [PDF.js](https://mozilla.github.io/pdf.js/) reads the audit PDF
- [ExcelJS](https://github.com/exceljs/exceljs) writes the styled report

All libraries load from cdnjs, so an internet connection is needed on first load.

## 🚀 Deploy on GitHub Pages

1. Push these files to the repository root.
2. **Settings → Pages → Deploy from branch → `main` / root.**
3. Optional: add a custom domain under Pages settings.

Live site: **https://meetfluorine.github.io/audit-tracker/**

## 💻 Run locally

Open `index.html` in a browser, or serve the folder:

```bash
python3 -m http.server 8000
# then visit http://localhost:8000
```

---

<div align="center">

Made for faster, cleaner inventory audits. 📦

</div>
