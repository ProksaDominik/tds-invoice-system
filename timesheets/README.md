# Timesheets — handwritten timesheet recognition

Companion app to the invoice pipeline. Workers at TDS s.r.o. fill in a paper
monthly timesheet by hand: hours per day for each job order. This app reads
the scanned sheets and writes the hours into the shared payroll workbook,
after a person has checked every value.

## Pipeline

1. **Render and orient.** Each PDF page is rendered at 300 DPI. The form is
   scanned in landscape onto portrait pages, so all four rotations are tried
   and the one where OCR finds the form's header keywords wins.
2. **Recognise**, with a fallback chain:
   - **Gemini** (`gemini-flash-latest`, structured JSON output) reads the whole
     page in one call: name, month, year and every job-order row with its
     daily hours.
   - **Azure AI Document Intelligence** if Gemini fails (e.g. quota exhausted).
     Its word boxes are mapped onto a calibrated template of the form.
   - **Tesseract**, fully offline: each day cell is classified as empty,
     crossed out (weekend) or filled by its share of dark pixels, and only
     filled cells are OCR'd.
3. **Group pages.** Consecutive pages of one worker are merged. A page whose
   name could not be read is kept separate, so that two people's hours are
   never merged by mistake.
4. **Review.** The full grid (days 1–31 × job orders) is shown as editable
   fields next to the scanned page. Job orders unknown to the payroll workbook
   are highlighted before approval.
5. **Write.** On approval, the hours and rate go into the employee's column
   block on the month's sheet, on the row of each job order listed in the
   `SUMAR` sheet. A missing month sheet is created from the January template.

## Setup

On Windows, `setup.bat` installs Python, Tesseract and the dependencies and
creates a desktop shortcut. Then copy `config.example.json` to `config.json`
and set the payroll workbook path (`subor_mzdy_siet`; `{rok}`
is replaced by the year), the employee → column mapping
(`zamestnanci_stlpce`) and the API-key files for Gemini and, optionally,
Azure. Employee names and keys stay in `config.json` and the key files, which
are not versioned.

The end-user manual is in Slovak: [`user-guide-sk.md`](user-guide-sk.md).
Tests for the payroll writer and the autofill cache live in the repository's
[`tests/`](../tests/) folder.
