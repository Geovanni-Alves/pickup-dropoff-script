# Pickup & Drop-Off Scripts (Google Apps Script + Google Sheets)

This repository contains two Google Apps Script projects designed to automate and manage student pickup and drop-off workflows using Google Sheets.

These scripts support route organization, attendance tracking, van assignment, and export tools for after-school transportation programs.


---

## 📁 Structure

pickup-dropoff-scripts/
├── pickup/       # Pickup script project
├── dropoff/      # Drop-off script project
└── README.md

Each folder is a standalone Apps Script project managed via `clasp`.

---

## 🚀 Getting Started

1. **Install clasp**

   npm install -g @google/clasp
   clasp login

2. **Clone or link projects**

   cd pickup
   clasp clone <PICKUP_SCRIPT_ID>

   cd ../dropoff
   clasp clone <DROPOFF_SCRIPT_ID>

3. **Push changes back to Google Sheets**

   clasp push

---

## 🔐 .gitignore

Make sure your repo ignores local bindings and credentials:

.clasp.json
.clasprc.json

---

## 📄 License

MIT License — use freely with attribution.
