# Presswork — Excel to PDF (LibreOffice + auto table styling)

Converts `.xlsx`, `.xls`, `.xlsm`, `.csv`, and `.ods` files to PDF.

Unlike a plain LibreOffice conversion, this version **always produces a
proper-looking table** regardless of what formatting the original file had:

1. A small Python step (`prepare.py`, using `openpyxl`) opens the file and:
   - draws real borders around every used cell (so gridlines always show)
   - bolds and shades the header row
   - sets the page to fit all columns on one page width, letting rows flow
     across as many pages as needed
   - repeats the header row on every printed page
2. LibreOffice then exports that styled workbook to PDF.

This fixes the classic problem where a wide spreadsheet gets split into
separate "column groups" that print as unrelated pages — everything is now
forced to fit one page's width first.

## What's included

```
presswork-excel/
├── server.js        Node/Express backend — orchestrates styling + conversion
├── prepare.py        Python script that applies table styling (via openpyxl)
├── package.json
├── Dockerfile        Bundles Node + LibreOffice + Python/openpyxl
├── public/
│   └── index.html    The website
└── README.md
```

## Run it locally

You now need **three** things installed (LibreOffice and Python are usually
quick installs if you don't have them):

1. **Node.js** (v18+) — https://nodejs.org
2. **LibreOffice** — https://www.libreoffice.org/download
3. **Python 3** with the `openpyxl` package:
   - **Windows:** install Python from https://python.org (check "Add to PATH" during install), then:
     ```
     pip install openpyxl
     ```
   - **macOS:** Python 3 is usually preinstalled; if not, `brew install python`, then:
     ```
     pip3 install openpyxl
     ```
   - **Linux:**
     ```
     sudo apt-get install python3 python3-openpyxl
     ```

Then, from this folder:
```
npm install
npm start
```

If Windows can't find LibreOffice, set the path before starting (same as the
Word converter):
```
$env:SOFFICE_PATH="C:\Program Files\LibreOffice\program\soffice.exe"
npm start
```

If Windows can't find Python, or you have both `python` and `python3`
commands and it picks the wrong one, set this too:
```
$env:PYTHON_PATH="python"
npm start
```

Open `http://localhost:3000` and try converting a spreadsheet.

## Publish it

Same as before — the included `Dockerfile` now installs LibreOffice **and**
Python/openpyxl in one image, so deployment to Render/Railway/Fly.io needs
no extra steps:

1. Push this folder to a GitHub repo (e.g. `presswork-excel`).
2. On Render: **New +** → **Web Service** → select the repo → it
   auto-detects the `Dockerfile` → **Create Web Service**.
3. Wait for the build, then use the `https://your-app.onrender.com` URL.

## Notes

- Max upload size is 25 MB (`MAX_FILE_SIZE` in `server.js`).
- `.xlsx` / `.xlsm` files are styled directly. `.xls` / `.ods` files are
  first converted to `.xlsx` by LibreOffice, then styled the same way.
- If the Python styling step fails for any reason, the server automatically
  falls back to converting the original file without styling, rather than
  failing the whole request.
- Because styling is now always applied, any borders/colors already in your
  original spreadsheet will be overridden with the standard table look
  (thin gray borders, shaded bold header). If you'd rather it only add
  styling when none exists, let me know and I can make that conditional.
