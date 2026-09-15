# iCanvas needfinding kit

Three files that follow the persona process from the Week 02–03 slides (collect → find patterns → cluster segments → write a persona) for the hypothetical **iCanvas** lecture management system.

| File | What it is | Who uses it |
|---|---|---|
| `iCanvas_needfinding.xlsx` | The data source: questionnaire + codebook, 124 **simulated** respondents on the `Responses` sheet, live-formula statistics (1-Patterns), k-means segments (2-Segments), a full persona with evidence table, POV and HMWs (3-Persona) | You, when preparing or explaining the analysis |
| `vis.html` | Opens with a **Simulate users** button; pressing it reads the `Responses` sheet of the workbook, replays the simulated submissions one by one and builds the same three analysis steps in the browser (patterns, k-means with k = 2…8 and a PCA map, a generated and editable persona per segment) | You, on the projector |
| `user.html` | The same 34-item questionnaire as a stand-alone, phone-friendly web page; every question can be skipped | Students, during the lecture |

**The data is simulated.** The 124 respondents were generated (seed 2026) from six latent behavioural profiles with realistic noise (one straight-liner, skipped optional items, near-duplicate open answers). The README sheet inside the workbook says the same. Nothing in it is a statement about real Tilburg University users. `user.html` does not feed the workbook; the workbook is the fixed teaching dataset.

## How vis.html works

Open `vis.html` and press **Simulate users**. The page then finds its data in one of two ways:

1. **Served from a web server, next to the workbook** (Canvas Files, GitHub Pages, SURFdrive, your own site, or `python3 -m http.server` in the folder): the page fetches `iCanvas_needfinding.xlsx`, reads the `Responses` sheet with SheetJS (loaded from cdnjs) and starts the replay. The header chip shows *N = 124 · Responses sheet of iCanvas_needfinding.xlsx*.
2. **Opened as a local file** (double-click): browsers block a page from reading files next to it, so the page falls back to an embedded copy of the same simulated users and says so. Drop the workbook on the **Load data** tab to read it directly.

The replay bar shows submissions arriving in the order of their `submitted` timestamps (Pause, speed 1–8×, **Load all now** to skip, Replay again). The statistics update as responses arrive; clustering and the persona are recomputed on the complete sample at the end. Segments and sizes match the workbook (34 / 29 / 20 / 17 / 15 / 9).

The **Load data** tab also accepts another copy of the workbook, a CSV, or JSON files (replace or append), which is how a class's own answers could be analysed if you ever want to; nothing is written back to the Excel file.

## Lecture workflow

1. Unzip the kit and keep `user.html`, `vis.html` and `iCanvas_needfinding.xlsx` in the same folder (host them together if you serve them).
2. Show the link or QR code for `user.html`. Students go through the five sections on their phones (about 7 minutes); every question can be skipped and the Next button always moves on. On the thank-you screen they can download or copy their answers if you want to look at them.
3. Meanwhile open `vis.html` on the projector and press **Simulate users**: the simulated submissions stream in, then walk through tab 1 → 2 → 3. Change k and press **Recompute**, pick a segment on tab 3, edit the persona card in place, and **Print / save PDF** for a handout.

If you also want to collect the class's answers automatically, `user.html` can POST each response to any endpoint that accepts JSON: set `CONFIG.SUBMIT_ENDPOINT` at the top of its script, or share the link as `user.html?endpoint=<URL>`. A minimal Google Apps Script that appends rows to a sheet is:

```javascript
const SHEET_NAME = "responses";
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.tryLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    const data = JSON.parse(e.postData.contents);
    if (sh.getLastRow() === 0) sh.appendRow(Object.keys(data));
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    sh.appendRow(headers.map(h => data[h] === undefined ? "" : data[h]));
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } finally { lock.releaseLock(); }
}
function doGet() { return ContentService.createTextOutput("iCanvas survey endpoint is live."); }
```

Deploy it as a web app (Execute as *Me*, access *Anyone*) and use the `/exec` URL. This is optional and independent of the workbook.

## Survey design choices (mapped to the slides)

- Demographics as in Muller et al. Table 55.1, with age in ranges and no identifiers (privacy). Accessibility needs and outside commitments are asked so the **extreme users** from the Participants slide can be found.
- Habits are asked as behaviour in a fixed recall window ("last 7 days", "most recent deadline"), not as opinions; frequency and difficulty scales use the numerical equivalences of Table 55.2.
- Section C uses 7-point Likert statements with an **inverted duplicate pair** (C1/C2); the workbook and vis.html compute Cronbach's α for the pair.
- One **ranking** question (D1) and three **open-ended grand-tour** questions (E1–E3) phrased non-leading: "the last time it frustrated you", not "what do you dislike".
- No "Is feature X important to you?", no "What would you like in a tool?", no hypotheticals.

## What is where in the workbook

- **README** — purpose, stakeholder map, method, sheet map, coding conventions, data provenance.
- **Questionnaire / Codes** — every item with type, options, numeric coding, column name and the slide-based rationale; the code→label→value table the formulas use.
- **Responses** — one row per simulated respondent; grey derived columns are INDEX/MATCH recodes, purple columns are the k-means segment and PCA coordinates (pasted, because Excel cannot run k-means). This is the sheet vis.html reads.
- **1-Patterns** — all statistics are live formulas (COUNTIF/AVERAGEIFS/CORREL) on Responses; charts on the right; "Key patterns" narrative at the bottom.
- **2-Segments** — k selection (elbow + silhouette), segment sizes, profile matrix with per-row colour scale, role × segment, descriptions with quotes, PCA scatter.
- **3-Persona** — the Sanne de Vries card in the template layout, an evidence table (segment vs all, live), the 4-sentence POV, five HMWs, and one-line persona candidates for the other five segments.

## Regenerating or changing the questions

The `source/` folder contains the pipeline: `questions.json` is the single source of truth for the Excel codebook, `user.html` and `vis.html`. Edit it, then run `python3 gen.py && python3 analyze.py && python3 finalize.py && python3 build_xlsx.py && python3 build_html.py` (needs pandas, scikit-learn, openpyxl; recalculate the workbook once in Excel or LibreOffice).
