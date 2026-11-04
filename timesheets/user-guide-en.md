# Timesheets – TDS s.r.o.

*English version of [`user-guide-sk.md`](user-guide-sk.md), the guide the office
staff use. The app's buttons are in Slovak; their English meaning is given in
brackets.*

An app that turns scanned paper **PRACOVNÉ VÝKAZY** (an employee's monthly
timesheet, by job) into an Excel workbook: it adds up the hours worked on each
job, multiplies them by the hourly rate and calculates the amount in €. The
whole table and the rate stay **fully editable** in Excel, in case the program
misreads something.

Handwriting is read with the **Gemini API** (Google, the free model
`gemini-flash-latest`), which needs internet and a free API key. If Gemini fails
(e.g. the daily quota is used up), the app tries **Azure AI Document
Intelligence** (optional, also with a free tier, see 4b) and only as a last
resort local Tesseract (works offline, but is the least accurate on handwriting).

---

## 1. Why everything needs checking

The timesheets are **handwritten**, not filled in by machine. Automatic
handwriting recognition (OCR/AI) will never be 100 % accurate, so the app:

1. reads everything at once with Gemini (name, month, year, all rows and hours);
   on a real timesheet it read all 9 job numbers and almost all the hours
   correctly, which local Tesseract (without Gemini) could not do at all,
2. **ALWAYS** shows the whole timesheet on screen as an editable table before
   anything is written to Excel,
3. lets you open the original scan side by side at any time with
   **"🔍 Zobraziť naskenovanú stranu"** (show scanned page) to check and correct
   the values.

Still, always check the job numbers and the hours on densely written rows.

---

## 2. How it works

1. **"📄 Načítať výkaz (PDF)"** (load timesheet) – choose one or more scanned
   timesheets. On each page the app detects the rotation, the name, month and
   year, and the hours per job (columns day 1-31).
2. The timesheet appears in the list on the left as **"Čaká na kontrolu"**
   (awaiting review). Click it to see the whole table on the right.
3. **Check and correct** the values (compare with the scan via 🔍). If the
   program missed a whole job row, add it with **"➕ Pridať riadok zákazky"**
   (add job row).
4. **"✅ Schváliť"** (approve) – writes the timesheet to `data/dochadzka.xlsx`
   (sheet **Záznamy**) and recalculates the summary for each job. The original
   PDF is moved to `vykazy_spracovane/`.
5. **"📊 Otvoriť súhrn (Excel)"** (open summary) – opens `data/dochadzka.xlsx`
   on the sheet **Súhrn podľa zákazok** (summary by job: total hours × hourly
   rate = €).

One PDF file = one employee for one month (even with several pages, when there
were more than 9 jobs). If you scan several employees at once, save them as
separate PDF files.

Processing one page with Gemini takes about 30-70 seconds (it goes over the
internet), so a multi-page timesheet can easily take a few minutes. The app stays
usable meanwhile (it runs in the background, with a progress bar).

---

## 3. The workbook `data/dochadzka.xlsx`

| Sheet | Contents |
|---|---|
| **Nastavenia** (settings) | The default hourly rate (€). Change it here; it applies wherever the rate was not overwritten by hand for a specific job. |
| **Záznamy** (records) | Raw data: one row = employee × month × year × job, with columns Day 1-31. **This is the source of truth**: a fix here (or in the app before approval) carries through to the summary. |
| **Súhrn podľa zákazok** (summary by job) | One row per job: total hours (automatic, from Záznamy), hourly rate (by default a link to Nastavenia; can be overwritten by hand for any job) and the total in €. |

The **"Spolu hodín"** (total hours) and **"Spolu v EUR"** (total in EUR) columns
on the summary sheet are recalculated every time a timesheet is approved; do
not edit them by hand (the change would be lost). The **"Hodinová sadzba"**
(hourly rate) column, on the other hand, is there for you to edit as needed.

If you approve a timesheet again (e.g. after a bigger correction) and its source
file is already in Záznamy, the app asks whether to replace the old records or
add them as new ones (watch out for duplicates with "add").

---

## 4. Gemini API key (once, about 2 minutes)

1. Open <https://aistudio.google.com/apikey> and sign in with a Google account
   (no payment card needed; this is the free tier).
2. Click **"Create API key"** and copy the generated key.
3. Put it as the only line in a file called **`gemini_api_key.txt`** in the app's
   folder (next to `app.py`).

The free tier limits the number of requests per day and minute, but at the usual
volume (around 15-20 timesheets a month) the limit is practically out of reach.
The key is a private credential: do not share `gemini_api_key.txt` with anyone or
put it anywhere public.

**Note:** on the free tier some specific (date- or version-named) models have zero
quota and return error 429 on the very first call, so the app uses the model
`gemini-flash-latest` by default (an alias for the currently available free
model), which works reliably. To use another model, change `gemini_model` in
`config.json`, but first check that your account has a non-zero quota for it
(Google's error message always says so).

Without this key (or without internet) the app automatically tries Azure Document
Intelligence (4b) if it is set up, otherwise local Tesseract straight away. That
works, but less accurately (more orange, unrecognised fields to check by hand).

---

## 4b. Azure Document Intelligence key (optional cloud fallback)

This path is used only when Gemini fails for a page (e.g. the quota is used up),
so it is optional; the app works without it (it goes straight to Tesseract).

1. Open <https://portal.azure.com>, sign in (no payment card needed for the free
   F0 tier) and create a **"Document Intelligence"** resource (Create a resource →
   search for "Document Intelligence" → pricing tier **Free F0**, about 500 pages
   a month).
2. Once the resource exists, go to **"Keys and Endpoint"** and copy the
   **Endpoint** (e.g. `https://your-resource.cognitiveservices.azure.com/`) and
   **KEY 1**.
3. In `config.json`, set `azure_di_endpoint` to the copied endpoint.
4. Put the key as the only line in a file called **`azure_di_key.txt`** in the
   app's folder (next to `app.py`); the exact file name is set by
   `azure_di_kluc_subor` in `config.json`.

The key is a private credential: do not share `azure_di_key.txt` with anyone or
put it anywhere public. If `azure_di_endpoint` is empty or the key file is
missing, the app simply skips Azure without an error (as with a missing Gemini
key).

---

## 5. Installation (once)

Python and all the packages it needs (openpyxl, ttkbootstrap, pdfplumber,
pytesseract, Pillow, google-genai, azure-ai-documentintelligence) are already
installed on the office computer. To run it on another computer:

1. Install Python 3.10+ (tick "Add Python to PATH").
2. Install [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) (the
   Slovak language data is in the `tessdata/` folder; it is only the last offline
   fallback, when both Gemini and Azure fail).
3. In the app's folder, run: `pip install -r requirements.txt`
4. Create `gemini_api_key.txt` as in section 4 above.
5. Optionally create `azure_di_key.txt` as in section 4b above.

## 6. Daily use

Double-click **`start.bat`** (or run `python app.py`).

---

## 7. Files

| File / folder | Purpose |
|---|---|
| `app.py` | the main app (window) |
| `extractor.py` | reading the PDF scan (rotation, Gemini/Azure/Tesseract OCR) |
| `template.py` | calibrated layout of the printed form (used by the Azure and Tesseract fallbacks) |
| `config.json` | settings (rate, folders, Gemini/Azure access) |
| `gemini_api_key.txt` | your private Gemini API key (created as in section 4) |
| `azure_di_key.txt` | optional private Azure Document Intelligence key (created as in 4b) |
| `vykazy/` | where you can save scanned PDFs before loading them |
| `vykazy_spracovane/` | approved PDFs are moved here (archive) |
| `tessdata/` | language data for the local OCR fallback (Slovak + English) |

---

## 8. Troubleshooting

**Name/month/year not read** – rare when Gemini is running; if it happens (or the
local Tesseract fallback is in use), fill them in by hand at the top of the table
before approving.

**Many orange (unrecognised) fields** – the local fallback was used (check
`gemini_api_key.txt` and the internet connection). Type in the values by hand
from the 🔍 scan preview.

**Error 429 / "quota exceeded"** – either the free daily limit is used up (try
again the next day), or `gemini_model` in `config.json` points to a model with no
free quota on your account (switch back to `gemini-flash-latest`). The app then
switches to Azure automatically (if set up, see 4b) and only then to local
Tesseract; the exact reason is shown in the yellow or blue notice above the table
for that timesheet. With an exhausted quota the app skips that service for 2
minutes, so the rest of the batch does not wait on pointless retries.

**"⚠️ Over rozloženie"** (check layout) in the list – the app was not sure the
page is the standard form (e.g. a badly tilted scan). Check every single value
carefully.

**The Excel file cannot be saved** – close `dochadzka.xlsx` in Excel; the app
cannot write to it while it is open (it offers to try again).
