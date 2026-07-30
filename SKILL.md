---
name: tu-bbs-report-format
description: Fix formatting in Tribhuvan University BBS and MBS project reports and theses (.docx). Use this skill whenever someone mentions a BBS project report, MGT 401, a TU or Tribhuvan University thesis, Faculty of Management formatting, or describes any of these symptoms in a Word document — every page showing the same number, roman numerals turning into 1 2 3, page numbers disappearing after editing in Google Docs, paragraphs that look different sizes despite all saying 12pt, words squished together in a table of contents, huge gaps between words in a bibliography, headings stranded at the bottom of a page, captions separated from their table, tables sticking out past the margin, blank pages appearing from nowhere, or table of contents page numbers that do not match the real pages. Also use it when asked to check or verify a report's page numbers, or to apply TU formatting rules to any academic document.
---

# TU BBS Report Formatting

Repairs the formatting bugs that circulate inside inherited Word templates used
for Tribhuvan University BBS 4th-year project reports (MGT 401: Final Project).

These bugs are structural — they live in the document's XML, survive being
copied from student to student, and cannot be fixed by retyping text.

## Before you start

**Always work on a copy.** These scripts rewrite the file's internals.

**Ask for the .docx, never a PDF.** A PDF has no editable structure, so
"fixing" one means rebuilding it from extracted text — which risks silently
altering numbers in tables. If the user only has a PDF, say so plainly and warn
them they must proofread every figure afterwards.

**Install the dependencies:** `pip install lxml pdfplumber`

## Diagnose first

Run the verifier before changing anything, so you know what you're dealing with
and can show the user the difference afterwards:

```bash
python scripts/verify_pages.py report.docx
python scripts/fix_page_numbers.py report.docx --check
```

Read `references/troubleshooting.md` when a symptom isn't obvious. It's indexed
by what the user sees rather than by cause.

## Repair order matters

Run these in sequence. Each one shifts pagination, so verification must come
last.

```bash
python scripts/fix_page_numbers.py report.docx -o step1.docx
python scripts/fix_fonts.py       step1.docx  -o step2.docx
python scripts/fix_layout.py      step2.docx  -o step3.docx
python scripts/verify_pages.py    step3.docx
```

| Script | Fixes |
|---|---|
| `fix_page_numbers.py` | literal text in footers, missing `lowerRoman`, malformed PAGE fields |
| `fix_fonts.py` | mixed typefaces, undersized list markers, embedded fonts |
| `fix_layout.py` | TOC word spacing, orphaned headings, empty spacing paragraphs |
| `verify_pages.py` | TOC claims vs actual pagination, stranded headings, margins |

## After repairing

`verify_pages.py` will list page-reference mismatches. **These must be corrected
by hand or by script — the tools do not update TOC numbers automatically**,
because the correct value depends on the final layout.

For each mismatch, edit the number in the Table of Contents, List of Tables or
List of Figures entry, then re-run the verifier. Repeat until it reports none.

Tell the user to open the result in Word and press `Ctrl+A` then `F9` to refresh
the page-number fields.

## The TU rules

| Setting | Requirement |
|---|---|
| Paper | White, 8½ × 11 in (`pgSz w=12240 h=15840`) |
| Font | Times New Roman 12pt (`sz=24`) |
| Line spacing | 1.5 (`w:line="360" w:lineRule="auto"`), justified |
| Margins | ≥1 in all sides; 1.5 in left is common for binding |
| Preliminary pages | lower-case roman from **ii** on the Declaration, centred footer, title page unnumbered |
| Body pages | arabic restarting at **1**, upper right |
| Tables | caption **above** |
| Figures | caption **below** |
| Bibliography | APA, hanging indent, left-aligned not justified |
| Length | ~8,000–10,000 words, roughly 30–35 body pages |

Front matter order: Declaration, Supervisor's Recommendation, Endorsement,
Acknowledgements, Table of Contents, List of Tables, List of Figures,
Abbreviations.

Three chapters: Introduction; Results and Analysis; Summary and Conclusion.

**Chapter 1 contains** background, profile of the organisation, objectives,
rationale, method, review of literature, limitations. It does **not** include a
Statement of the Problem — that isn't in the TU specification, and supervisors
often ask for its removal.

## Things the guideline doesn't decide

Don't guess at these. Tell the user to ask their supervisor:

- **Cover paper colour** — often assigned by concentration group, varies by campus
- **Letterhead** — TU requires campus letterhead for the Supervisor's
  Recommendation and Endorsement pages; those two are printed separately onto
  sheets the supervisor supplies
- **Caption placement** — some supervisors want table captions below the table,
  contradicting the guideline. Follow the supervisor, but say once that it
  departs from the written rule.

## Checks worth running unprompted

If the report contains financial ratios, **recompute each one from the figures
printed beside it**. In one real submission, two of forty ratios didn't
reconcile with their own inputs — recomputable by anyone with a calculator.

Also confirm every author cited in the text appears in the Bibliography, and
that every listed abbreviation is actually used. Both map directly onto marking
criteria.

## What not to do

Don't use `python-docx` — the bugs live in parts of the XML it can't reach.

Don't apply `keepNext` to every row of a table to stop it splitting. It makes
the table unbreakable, pushes it wholesale onto the next page and leaves large
white gaps. On a 43-page report this cost a page and left six pages under 62%
full. Apply it to the last row only.

Don't reorder `pPr` children carelessly — OOXML mandates a specific order and
Word rejects files that violate it. `fix_layout.py` includes a reordering pass;
reuse it rather than writing your own.

Don't let the user edit in Google Docs after this. It strips `lowerRoman`,
collapses PAGE fields and reintroduces mixed fonts.
