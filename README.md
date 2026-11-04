# Invoice processing with my own OCR and extraction models

[![CI](https://github.com/ProksaDominik/tds-invoice-system/actions/workflows/ci.yml/badge.svg)](https://github.com/ProksaDominik/tds-invoice-system/actions/workflows/ci.yml)
![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue)
![License: MIT](https://img.shields.io/badge/license-MIT-green)

A desktop app that turns PDF invoices into accounting records. It is used
every day at **TDS s.r.o.**, a Slovak HVAC company: it reads each invoice
(digital or scanned), pulls out the data, and after a person has checked it,
writes it to the company's Excel ledgers. Everything runs on the office PC;
invoices never leave the company.

On top of the app I trained two models of my own: an **OCR model** that reads
text from scanned pages (a CRNN, trained from scratch) and a **layout model**
that finds the invoice data on the page (LiLT, fine-tuned).

## Results in one minute

**Data:** 437 real company invoices (238 issued, 184 received; 160 digital, 277
scanned) and 500 generated received invoices (300 for training, 200 for
testing). Tested on 82 real invoices from suppliers the models never saw, two fields each (amount
to pay and variable symbol): 164 values, checked against the payment QR code
printed on each invoice:

| Pipeline | Right | Wrong | Missed | Every field right |
|---|---:|---:|---:|---:|
| Today's app: rules, Gemma 4 for gaps | 151 | 1 | 12 | 69 of 82 |
| Fair version: Gemma 4 also asked for the variable symbol | 158 | 1 | 5 | 76 of 82 |
| Rules, my LiLT model for gaps | 159 | 1 | 4 | 77 of 82 |
| **My own stack only:** my OCR + rules + my LiLT (no PaddleOCR, no Gemma) | **159** | **1** | **4** | **77 of 82** |

- **My OCR model** gets 2.2% of characters wrong, against 7.5% for PaddleOCR (the
  OCR the app uses), on the same 1,386 real text lines. On whole pages it finds
  97% of the words and reads 95% of them exactly.
- **My own stack** matches the PaddleOCR version value for value, and reads a
  scanned invoice in 4.6 s instead of 11.6 s.
- **Gemma vs my model:** most of the gap closed once Gemma was asked for the
  variable symbol too; the app had simply never asked it. Both are now close.
- **Every field:** on 100 fresh generated invoices where every answer is
  known (supplier, IČO, VAT number, dates, amounts: 12 fields), the app gets
  88% of the values right, and 97% of the values it fills in are correct. The
  test also found a bug: the rules only knew the Slovak label "IČO", not Czech
  "IČ" or English "Company ID". After the fix that field went from 61% to 100%.
- **What did not work:** adding 300 generated invoices to LiLT's training did
  not make it better on real invoices.

**See it yourself:** [`notebooks/pipeline.ipynb`](notebooks/pipeline.ipynb)
trains both models with live charts and reproduces every number above (about
2½ hours on a laptop GPU, or a 15-minute quick mode).

## Version 2.0 (in progress)

I am working on a second version that is not tied to one company:

- **Universal:** any company's invoices, configured instead of hard-coded
  (today the rules know TDS's own name, ID and ledger layout).
- **A database** instead of Excel workbooks, so every invoice, correction and
  approval is stored and searchable, and every correction becomes new
  training data.
- **Docker:** the extractor as a small web service with a review page, started
  with one `docker compose up`.
- **Faster and leaner:** my own OCR and LiLT model in the app itself, so it
  needs neither PaddleOCR nor an LLM server.

The repository also contains a companion app, [`timesheets/`](timesheets/),
which reads handwritten monthly timesheets and writes the hours to the
payroll workbook.

## The problem

Received invoices come from dozens of suppliers, each with its own layout.
Some are digital PDFs, most arrive as scans, and some have a broken text layer
(the page looks fine, but the embedded text is missing digits). Fields move
around, labels wrap across lines, two-column layouts get merged into one line
by the PDF reader, and OCR turns `IČO:` into `ICO;`. One template, or one
regex, cannot cover that.

## How it works

```mermaid
flowchart LR
    pdf[PDF invoice] --> layer{Text layer?}
    layer -- yes --> txt[pdfplumber]
    layer -- no --> ocr[PaddleOCR PP-OCRv6<br/>216 DPI]
    txt --> rules
    ocr --> rules[Rule-based extraction<br/>+ arithmetic validation]
    rules --> gap{Received invoice and a<br/>critical field still empty?}
    gap -- yes --> llm[Local LLM fill-in<br/>Ollama, JSON schema]
    llm --> check[Re-validate;<br/>drop LLM values that<br/>break the arithmetic]
    gap -- no --> stage
    check --> stage[Staging workbook<br/>'pending review']
    stage --> review[Human review in GUI<br/>edit / approve / reject]
    review --> out[Ledgers + PDF archive<br/>+ calendar reminders]
```

1. **Text.** `pdfplumber` reads digital PDFs; pages with no usable text are
   rendered at 216 DPI and read with PaddleOCR (PP-OCRv6 medium detection and
   multilingual recognition, which covers Slovak diacritics). OCR returns loose
   text boxes, so they are regrouped into lines by vertical position and read
   left to right, with a wide gap kept as a column break: labels and values end
   up on one line, the same shape the rules get from a PDF text layer. It runs
   on an NVIDIA GPU when one is present.
2. **Rules.** Targeted extractors find the invoice number, supplier and
   customer, company IDs (IČO), issue/supply/due dates, net amount, VAT, total
   and payment reference. They handle labels wrapped onto the next line, values
   that the PDF reader moved far from their label, several amounts merged into
   one line, `§69` reverse-charge references that look like thousands, and
   common OCR misreads of separators.
3. **Validation.** Net + VAT must equal the total (±0.02 €); a missing amount
   is derived only when the other two are known; the implied VAT rate must be
   realistic (0–45 %); dates more than three years from today are rejected as
   OCR misreads. Values that fail are cleared rather than guessed, so the
   reviewer sees an empty field instead of a plausible wrong one.
4. **Classification.** Direction comes from the company's own ERP template
   (issued invoices carry fixed phrases) and from where the company's IČO and
   name appear. Domestic vs foreign comes from the VAT-ID country prefix, not
   the text language, because international platforms localise their invoices
   into Slovak.
5. **LLM fill-in.** Only for *received* invoices, and only for the critical
   fields the rules left empty, a local model (Gemma 4, `gemma4:e4b-it-qat`
   via Ollama, with thinking switched off) is asked for the missing values
   with a JSON-schema-constrained response. Its answers pass the same validation; if they break the
   arithmetic, the rule-based values are restored.
6. **Human in the loop.** Every record lands in a staging workbook first and
   reaches the ledgers only after approval. The staging copy is kept as an
   audit trail; approved PDFs are filed in the archive, received invoices
   under the company's internal archive number (DF/zDF).

### Design decisions worth knowing

- **Rules first, LLM for the gaps.** On the office CPU one call to the earlier
  3B model took 30–60 s, while the rules finish in milliseconds and are deterministic. The
  LLM is skipped entirely when the rules already filled every critical field.
- **The LLM never overrides the rules.** An earlier version asked the LLM to
  re-read every invoice and cross-check the regex values. In practice the
  earlier 3B model occasionally swapped the net amount and the total, and the
  "cross-check" destroyed correct values, so it was removed.
- **Thinking is switched off.** Gemma 4 reasons before answering by default;
  the hidden reasoning used up the whole response-token cap, the JSON answer
  came back empty and the fill-in silently found nothing. Every call now sends
  `think=false` (models that do not support the flag are called without it).
- **A vision-model fallback was tried and removed.** Reading amounts from the
  rendered page with a VLM rescued a few invoices with corrupted text layers,
  but was too slow on office hardware to justify.
- **Empty beats wrong.** In a review workflow a missing value costs the
  reviewer a few seconds; a wrong value that looks right can slip into the
  ledger. Validation therefore prefers clearing a field to keeping a doubtful
  one.

## Evaluation

[`evaluation/`](evaluation/) measures the extractor on real invoices by
re-running it on archived PDFs and comparing the output with ground truth from
one of two sources:

- **Values a person approved in the app** (staging workbook or an Invoice Desk
  database): every reviewable field.
- **Payment QR codes and archive file names**, with no manual labelling.
  Slovak invoices carry a PAY by square code and Czech ones a QR Platba code,
  generated by the issuer's accounting system from the invoice data; the
  extractor never reads them. They give the amount to pay and the variable
  symbol. Archive file names add the direction (received invoices are filed
  under the internal number assigned on approval) and, for issued invoices,
  the invoice number. Documents with two different payment codes, or an
  issued invoice whose QR variable symbol differs from its file name, get no
  labels from that source. The QR amount is the amount *to pay*: on an
  invoice with a retention or a deducted advance it differs from the total, so
  such invoices count as misses for the total.

  The QR code's date is **not** used, although it is meant to be the due date.
  The first scored run put due-date accuracy at 7 %, so the labels were checked
  before the extractor: on 220 of 242 issued invoices the QR date is exactly
  14, 30 or 60 days (the payment terms) before the printed due date: the
  invoicing system writes the issue date there, and suppliers differ.

```bash
python -m evaluation --pdfs <archived PDFs> --zaloha <staging workbook>      # approved values
python -m evaluation.payment_qr --pdfs <PDFs> --out labels_qr.csv             # QR + file-name labels
python -m evaluation --pdfs <PDFs> --labels labels_qr.csv                     # ... scored with them
python -m evaluation.rescore --run <run> --labels <fixed.csv> --out <dir>     # re-score saved predictions
python -m evaluation --pdfs <PDFs> --llm gemma4:e4b-it-qat                    # + LLM setup
python -m evaluation --pdfs <PDFs>                                            # fill rate + speed only
```

It reports per-field accuracy (= recall), precision and F1, the same pooled
over all fields, the share of invoices with every field correct, and time per
stage (text layer, OCR, parsing, LLM). Only fields a reviewer can actually edit
in the app are scored; fields that are stored but never shown for review would
just compare the extractor with itself. The report contains aggregates only;
per-invoice output stays in the gitignored `eval_results/` folder.

### Results

437 real office documents (2026) with ground truth from payment QR codes and
archive file names; 160 digital, 277 scanned. Laptop with an RTX 5050 (8 GB).

| Field | Documents | Rules: accuracy / precision | Rules + Gemma 4: accuracy / precision |
|---|---:|---:|---:|
| Direction (received / issued) | 422 | 99.5 % / 99.5 % | 99.5 % / 99.5 % |
| Invoice number (issued) | 238 | 100 % / 100 % | 100 % / 100 % |
| Amount to pay | 352 | 90.3 % / 98.5 % | **96.3 %** / 98.5 % |
| Variable symbol | 352 | 85.2 % / 100 % | 85.2 % / 100 % |
| **All fields pooled** | 1,364 | 93.5 % / 99.5 % (F1 96.4 %) | **95.1 % / 99.5 % (F1 97.2 %)** |

- **Misses are mostly empty fields, not wrong values**: precision stays at
  99.5 %. The variable symbol is never wrong, but it is missing on 52 invoices,
  44 of them received ones, 31 of those digital: the rules do not recognise
  some suppliers' layouts. This is the gap the trained model closes (below).
- Of the 5 totals counted as wrong, the QR amount is 50 %, 9.5 % or 60 % of
  the extracted total, consistent with advance and retention invoices, where
  the amount to pay legitimately differs from the total.
- Gemma 4 was called on 132 of 437 documents (median 2.4 s per call).
- Speed per document (median): 0.15 s for a digital PDF, 11.6 s for a scanned
  one (median 3 pages, ~3.5 s per page with PaddleOCR on the GPU). On a CPU
  without CUDA PaddleOCR took 56–100 s per page, so the office PC needs an
  NVIDIA GPU for scans.

## Trained models

Both models are trained from the same office documents, split into train /
validation / test by **supplier** (303 / 46 / 96 documents from 153 / 21 / 62
supplier groups: QR IBAN, else IČO or name), so every test number is on
suppliers the model never saw. Code in [`training/`](training/); every number
below comes from one run of [`notebooks/pipeline.ipynb`](notebooks/pipeline.ipynb)
on a laptop with an RTX 5050 (8 GB).

### OCR: a CRNN text recognizer trained from scratch

CNN → two-layer BiLSTM → CTC, 8.5 M parameters, no pretrained weights. Trained
on synthetic invoice lines (generated field patterns, 30 fonts, scan-like
degradation that keeps diacritics readable) mixed 70 / 30 with 5,314 real lines
cut from digital invoices of training suppliers, whose text layer gives an exact
transcription. 10,000 steps, 40 minutes.

Test: 1,386 real lines from test-split suppliers; PaddleOCR's recognizer (the
app's model) scored on exactly the same crops:

| | Characters wrong | ... spaces ignored | Lines exact, spaces ignored |
|---|---:|---:|---:|
| **Own CRNN** | **2.2 %** | **2.1 %** | **90.2 %** |
| PaddleOCR PP-OCRv6 (pretrained) | 7.5 % | 2.9 % | 75.2 % |

Spaces are reported separately because a PDF text layer and an OCR engine space
words differently.

**Whole pages.** The CRNN reads one line at a time; [`training/ocr/detect.py`](training/ocr/detect.py)
finds the lines without any learned model (local threshold, straightening,
erasing table rules, joining letters into phrases). On 39 whole digital test
pages, compared with their text layer: 97.2 % of the words found, 95.3 % read
exactly, 0.6 % of characters wrong, 1.1 s per page.

### Extraction: LiLT fine-tuned on QR-labelled invoices

[LiLT](https://huggingface.co/nielsr/lilt-xlm-roberta-base) (XLM-RoBERTa base,
MIT licence) reads each word with its position on the page. It is fine-tuned to
tag the amount to pay and the variable symbol; training labels are automatic:
words that read as the QR values are tagged. Trained on the real training
invoices plus 300 generated received invoices
([`synthetic_invoices.py`](training/extraction/synthetic_invoices.py)), 8 epochs,
58 minutes.

Test split, 82 invoices, 164 values:

| | Right | Wrong | Missed | Accuracy | Precision | Every field right |
|---|---:|---:|---:|---:|---:|---:|
| Rules | 150 | 1 | 13 | 91.5 % | 99.3 % | 68 |
| Rules + Gemma 4 (today's app) | 151 | 1 | 12 | 92.1 % | 99.3 % | 69 |
| Rules + Gemma 4, also asked for the variable symbol | 158 | 1 | 5 | 96.3 % | 99.4 % | 76 |
| LiLT alone | 157 | 3 | 4 | 95.7 % | 98.1 % | 76 |
| **Rules, then LiLT for empty fields** | **159** | **1** | **4** | **97.0 %** | **99.4 %** | **77** |
| Own stack: own OCR → rules → LiLT | 159 | 1 | 4 | 97.0 % | 99.4 % | 77 |

- **Rules first, the model only for gaps:** LiLT alone picks a wrong total more
  often than the rules (3 vs 1), so it fills what the rules left empty and never
  overrides them.
- **The fair comparison:** the app asks Gemma only for amounts and dates. Asked
  for the variable symbol as well, Gemma closes most of the gap (96.3 % vs
  97.0 %); with 82 invoices the two are within each other's range.
- **Generated invoices did not help:** an earlier LiLT trained on real invoices
  only (20 epochs) reached 98.8 % in the same setting. The two runs also differ
  in length, but there is no sign that the generated data improved results on
  real invoices.
- **Own stack:** scanned pages read by the own OCR instead of PaddleOCR give
  the same 159 / 164; a scanned invoice takes 4.6 s (median) instead of 11.6 s.
- *Limits:* 82 test invoices (13 received), two scored fields, and the
  combinations are scored on saved predictions; LiLT is not wired into the app
  yet.

### Every field, on generated invoices

The payment QR codes only answer for two fields. For the others, 100 generated
received invoices (a fresh set, not the ones LiLT trained on and not the one
used to find the IČO bug; 56 of them scanned) have every answer known. The rules
and Gemma never saw a generated invoice, so this tests the extractor on
unfamiliar suppliers. The layouts all come from one generator, though, so it is
**not** accuracy on real invoices. 1,200 values; accuracy, with precision
(filled-in values that are right) in brackets:

| Field | PaddleOCR → rules | **PaddleOCR → rules + Gemma 4 (the app)** | Own OCR → rules |
|---|---:|---:|---:|
| Direction | 100 % (100 %) | 100 % (100 %) | 100 % (100 %) |
| Invoice number | 100 % (100 %) | 100 % (100 %) | 93 % (93.0 %) |
| Supplier name | 81 % (97.6 %) | 81 % (97.6 %) | 75 % (92.6 %) |
| Supplier IČO | 100 % (100 %) | 100 % (100 %) | 99 % (100 %) |
| Supplier VAT number | 100 % (100 %) | 100 % (100 %) | 98 % (100 %) |
| Issue date | 69 % (100 %) | 69 % (100 %) | 65 % (98.5 %) |
| Delivery date | 60 % (100 %) | 83 % (97.6 %) | 60 % (100 %) |
| Due date | 78 % (100 %) | 99 % (100 %) | 73 % (100 %) |
| Net amount | 69 % (86.2 %) | 91 % (91.9 %) | 62 % (86.1 %) |
| VAT amount | 60 % (75.9 %) | 87 % (89.7 %) | 54 % (75.0 %) |
| Amount to pay | 73 % (85.9 %) | 84 % (85.7 %) | 65 % (83.3 %) |
| Variable symbol | 67 % (98.5 %) | 67 % (98.5 %) | 63 % (98.4 %) |
| **All fields** | **79.8 % (95.5 %)** | **88.4 % (96.6 %)** | **75.6 % (94.2 %)** |
| Every field right | 11 of 100 | 22 of 100 | 9 of 100 |

- **Gemma earns its place on amounts and dates** (+8.6 points pooled). It is not
  asked for names, IDs or the variable symbol.
- **The IČO bug.** The first run of this test (on another generated set) put the
  supplier's IČO at 59 %: the rules only knew the label "IČO", while Czech
  invoices write "IČ" and English ones "Company ID". After the fix it is 100 %.
  On the real archive the fix finds the IČO on 15 more pages, and every
  QR-scored field on those documents stayed exactly the same.
- **Amount to pay:** 15 % of the generated invoices have a deposit, so the
  amount to pay differs from the total. The extractor often takes the total,
  which causes most of the wrong values in that row and in the VAT row.
- **VAT number:** the new field (the counterparty's IČ DPH, never the company's
  own) is right on every invoice that has one.

### What the measurements caught

- **The QR date is not the due date.** The first scored run put due-date
  accuracy at 7 %. The labels were checked before the extractor: on 220 of 242
  issued invoices the QR date is the issue date. The field was dropped.
- **A barcode drawn with a font.** One invoice's Aztec code has a text layer of
  random letters; its 324 "lines" were 73 % of the first OCR test errors.
- **Large title text cut in two.** The own OCR first read the issued invoice
  number right on only 74 % of scans: the number is printed in a title font
  twice the normal size, and the line detector split it ("202600" + "12").
  Rejoining pieces by their own height fixed all 45 cases.
- **Gemma was never asked.** "LiLT beats Gemma by 6 points" turned out to be
  mostly the app not asking Gemma for the variable symbol.
- **Gemma 4 thinks before answering.** The hidden reasoning used up the token
  budget and the JSON answer came back empty; calls now switch thinking off.
- **Laptop GPU limits.** Batches that overflowed the 8 GB were spilled to
  system RAM by Windows (0.17 → 2.8 training steps per second after the fix),
  and cuDNN's LSTM dropout made every finished run abort on exit.

## Tests

`pytest` runs 178 tests in about two seconds, with no PDFs, OCR or LLM needed (7 of them need the training environment and are skipped without it). The
PDF reader is stubbed and the pipeline runs on synthetic invoice texts
(received, issued, foreign reverse-charge). The tests cover the parsing rules,
the arithmetic validation, the LLM gate (never called without config or for
issued invoices; only asked for missing critical fields; a fill-in that breaks
net + VAT = total is discarded), regrouping OCR text boxes into lines, the
Ollama call (thinking off), the ledger helpers, the timesheet payroll writer,
the evaluation metrics, re-scoring, the payment-QR decoder that produces
ground truth, and the training data pipeline (charset and CTC decoding,
synthetic lines, word labels, the supplier split, the own OCR's line detection, the invoice generator) and the counterparty VAT number. CI runs ruff and the tests
on Windows and Linux, Python 3.10 and 3.12.

**Known limitations** (each one is a test or an evaluation case):

- An inline `Customer s.r.o., IČO: 12345678` layout keeps the `IČO` suffix in
  the party name (documented by a strict `xfail` test).
- Company names without a Slovak/Czech legal form (e.g. `… Ireland Limited`)
  are often not found by the rules and are filled in by the reviewer.
- Company IDs are validated by format only; the check-digit validation of
  earlier versions is not part of the current pipeline.

## Running it

Windows (the production platform):

```powershell
powershell -ExecutionPolicy Bypass -File scripts\install.ps1   # deps, PaddlePaddle (GPU/CPU), config.json, shortcut
```

Manually, on any OS:

```bash
python -m venv .venv && .venv/Scripts/activate    # or: source .venv/bin/activate
pip install -r requirements.txt
cp config.example.json config.json                # then set ico_firmy and file paths
python app.py
```

Optional components (the app skips a step if its component is missing):

| Component | Setup |
|---|---|
| OCR for scans | [PaddleOCR](https://github.com/PaddlePaddle/PaddleOCR) is in `requirements.txt`; its engine is installed separately: `pip install paddlepaddle` (CPU), or `paddlepaddle-gpu` from the [PaddlePaddle index](https://www.paddlepaddle.org.cn/en/install/quick) for an NVIDIA GPU. Models (~135 MB) download on the first scan. |
| LLM fill-in | [Ollama](https://ollama.com), then `ollama pull gemma4:e4b-it-qat` (6.1 GB) |
| Calendar reminders | Google OAuth `credentials.json` next to `app.py` (see the user guide, §2.4) |

Development:

```bash
pip install -r requirements-dev.txt
pytest
ruff check .
```

## Configuration

`config.json` is machine-specific and not versioned; start from
[`config.example.json`](config.example.json).

| Key | Meaning |
|---|---|
| `ico_firmy`, `varianty_nazvu` | Company ID and name variants, used to tell received from issued invoices |
| `ic_dph_firmy` | Own VAT number, never reported as the counterparty's (OCR misreads of it included) |
| `subor_prijate`, `subor_vystavene`, `subor_zaloha` | Ledger workbooks and the staging workbook |
| `priecinok_invoices` | Inbox folder with scanned invoices |
| `priecinok_faktury`, `priecinok_faktury_pre_skladnikov` | Archive for approved PDFs and the warehouse copy |
| `archivacia_tuzemske_od`, `archivacia_zahranicne_od` | Starting points of the internal archive numbering |
| `ollama_zapnute`, `ollama_model`, `ollama_pokusy` | LLM fill-in switch, model and retries |
| `kalendar_zapnuty`, `hromadny_termin`, `hromadne_pripomienky_dni` | Calendar reminders |

## Repository layout

| Path | Contents |
|---|---|
| `app.py` | GUI: upload, review table, edit / approve / reject |
| `extractor.py` | Extraction pipeline: text and OCR, rules, validation, LLM fill-in |
| `excel_manager.py` | Staging workbook, ledgers, archive numbering and PDF filing |
| `calendar_manager.py` | Google Calendar due-date and summary reminders |
| `evaluation/` | Accuracy and speed measurement against approved invoices or payment-QR ground truth |
| `training/` | Own models: CRNN text recognizer and line detection (`ocr/`), LiLT extraction and the invoice generator (`extraction/`), supplier split |
| `notebooks/pipeline.ipynb` | Trains both models and measures everything, step by step, with live charts |
| `tests/` | Unit and pipeline tests on synthetic data |
| `timesheets/` | Companion app for handwritten timesheets |
| `docs/user-guide-en.md`, `docs/user-guide-sk.md` | User guide (English, and the Slovak original the office uses) |
| `scripts/install.ps1` | One-click setup on a new Windows PC |

Comments, docstrings, documentation and tests are in English. Function and
variable names, config keys and the app's on-screen text are in Slovak, the
language of the people who use the app.

## Data handling

The repository contains no customer data. Invoices, ledgers, machine
configuration and credentials are gitignored, and the tests use made-up
invoices. Apart from creating calendar events, nothing is sent to external
services; the LLM runs locally through Ollama.

## License

[MIT](LICENSE)
