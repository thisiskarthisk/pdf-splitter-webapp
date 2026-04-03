╔══════════════════════════════════════════════════════════╗
║           PDF CHAPTER SPLITTER — SETUP GUIDE            ║
╚══════════════════════════════════════════════════════════╝

📁 FOLDER STRUCTURE
────────────────────
pdf_splitter/
│
├── main.py                  ← Flask backend (main server)
├── requirements.txt        ← Python packages needed
├── run.bat                 ← Double-click to run on Windows
├── run.sh                  ← Run on Mac / Linux
├── README.txt              ← This file
│
├── templates/
│   └── index.html          ← Web UI (auto-loaded by Flask)
│
├── uploads/                ← Auto-created (PDF uploads stored here)
└── outputs/                ← Auto-created (split PDFs saved here)


══════════════════════════════════════════════════════════
 STEP 1 — INSTALL PYTHON
══════════════════════════════════════════════════════════

Download Python 3.8 or newer from: https://www.python.org/downloads/
✅ During install on Windows: CHECK "Add Python to PATH"


══════════════════════════════════════════════════════════
 STEP 2 — RUN THE APP
══════════════════════════════════════════════════════════

 WINDOWS:
   Double-click  →  run.bat
   (It auto-installs packages and starts the server)

 MAC / LINUX:
   Open Terminal in this folder and run:
     chmod +x run.sh
     ./run.sh

 MANUAL (any OS):
   Open a terminal / command prompt in this folder:
     pip install flask pymupdf python-docx openpyxl requests gunicorn Pillow
     python3.10 main.py


══════════════════════════════════════════════════════════
 STEP 3 — OPEN IN BROWSER
══════════════════════════════════════════════════════════

After the server starts, open your browser and go to:

   👉  http://127.0.0.1:5000

You will see the PDF Chapter Splitter web app.


══════════════════════════════════════════════════════════
 HOW TO USE
══════════════════════════════════════════════════════════

 TAB 1 — CHAPTER SPLIT:
   1. Click "Upload your PDF Textbook" (or drag & drop)
   2. Wait a few seconds — chapters are auto-detected
   3. Click "Download" on any chapter card  OR
      Click "Download All ZIP" to get everything at once

 TAB 2 — BOOKMARK EXTRACT:
   1. Upload any PDF
   2. Left side shows all bookmarks — click one to auto-fill pages
   3. OR type Start Page and End Page manually
   4. Click "Extract & Download PDF"


══════════════════════════════════════════════════════════
 WORKS WITH
══════════════════════════════════════════════════════════

  ✅ All Tamil Nadu 11th / 12th Std textbooks
  ✅ Any PDF that has bookmarks (Table of Contents)
  ✅ Statistics, Accountancy, Economics, Commerce,
     Biology, Physics, Chemistry, Political Science, etc.


══════════════════════════════════════════════════════════
 STOP THE SERVER
══════════════════════════════════════════════════════════

  Press  CTRL + C  in the terminal window to stop.


══════════════════════════════════════════════════════════
 TROUBLESHOOTING
══════════════════════════════════════════════════════════

  ❌ "pip is not recognized"
     → Python not installed or not added to PATH
     → Re-install Python and check "Add to PATH"

  ❌ "Port 5000 already in use"
     → Open app.py, change port=5000 to port=5001
     → Open browser at http://127.0.0.1:5001

  ❌ "No chapters found"
     → PDF has no bookmarks
     → Use Tab 2 (Bookmark Extract) and enter page numbers manually

══════════════════════════════════════════════════════════
