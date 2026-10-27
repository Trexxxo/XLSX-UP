# XLSX.UP

XLSX.UP is a browser-based, open-source worksheet protection remover for Excel .xlsx workbooks. No backend, account, or file uploads to a server.

## How to use

1. Open https://trexxxo.github.io/XLSX-UP/
2. Drag and drop your .xlsx file (or click browse files).
3. Review which worksheets are protected.
4. Click Unlock & download .xlsx.
5. Your browser downloads original-name-unprotected.xlsx. Your original file remains untouched.

This is equivalent to renaming .xlsx to .zip, opening xl/worksheets/*.xml, removing the sheetProtection XML element, rebuilding the archive, and renaming it back to .xlsx.

## Features

- Modern responsive drag-and-drop interface.
- Detects and lists protected worksheets and their names where available.
- Removes worksheet-level editing protection from all protected worksheets.
- Keeps cell data, formulas, styling, and other workbook ZIP entries by modifying only protection-related worksheet XML.
- Downloads a new .xlsx copy without overwriting the original.
- Local processing in the browser using JSZip 3.10.1 loaded from the cdnjs CDN. The workbook itself is never transmitted to a server.
- Clear status and ZIP generation progress.

## Limitations

- Standard .xlsx (Office Open XML) only.
- Cannot open files encrypted with a password-to-open. The workbook must be readable as a ZIP.
- Does not remove workbook-structure protection, VBA project passwords, or other forms of encryption.
- Does not support .xls, .xlsm, .xlsb.
- Digital signatures may become invalid after editing.
- Very large workbooks need substantial browser memory.
- Only edit workbooks you own or have permission to modify.

## Run locally

Serve the project root with: python -m http.server 8000

Then visit http://localhost:8000. Internet is needed on first load to fetch JSZip from cdnjs.

## Publish with GitHub Pages

The workflow in .github/workflows/pages.yml verifies the code and deploys on pushes to main.

If GitHub Pages has not yet been enabled, visit Settings > Pages > Build and deployment > Source > GitHub Actions.

Then rerun the Deploy XLSX.UP workflow from Actions (or push a commit).

Expected site URL: https://trexxxo.github.io/XLSX-UP/

## Tests

Run: node --test tests/*.test.cjs

## Tech

HTML, CSS, browser JavaScript, JSZip 3.10.1, GitHub Pages.
