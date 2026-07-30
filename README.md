# TU BBS Report Format

Formatting toolkit for **Tribhuvan University BBS 4th-year project reports**
(MGT 401: Final Project), Faculty of Management.

Fixes the formatting bugs that get passed from student to student inside
inherited Word templates.

---

## Does your report have any of these?

- [ ] Every page shows the same number — `II` on all of them
- [ ] Roman numerals turned into 1, 2, 3 after you edited the file
- [ ] Some paragraphs look bigger than others, even though everything says 12pt
- [ ] Words squished together in the Table of Contents — `TableofContents`
- [ ] Enormous gaps between words in the Bibliography
- [ ] A heading stranded alone at the bottom of a page
- [ ] A table caption on one page and its table on the next
- [ ] Tables sticking out past the right margin
- [ ] Blank pages appearing from nowhere
- [ ] Table of Contents page numbers that don't match the real pages

**None of these are your fault.** They live in the XML underneath your document,
and they survive being copied from one student to the next. Retyping the text
won't fix any of them.

→ Start with **[references/troubleshooting.md](references/troubleshooting.md)**,
which is indexed by symptom.

---

## What's here

```
tu-bbs-report-format/
├── scripts/
│   └── fix_fonts.py              normalise every font to one typeface
├── references/
│   └── troubleshooting.md        symptom → cause → fix, plus a final checklist
└── LICENSE                       MIT
```

More scripts are being added — see [Roadmap](#roadmap).

---

## Quick start

Requires Python 3 and one library:

```bash
pip install lxml
```

Then, on a **copy** of your report:

```bash
python scripts/fix_fonts.py my-report.docx -o my-report-fixed.docx
```

Always work on a copy. These scripts rewrite the file's internals.

---

## Why the scripts edit XML directly

A `.docx` is a ZIP archive of XML files. Most Python libraries for Word documents
can't reach the parts where these bugs live — theme fonts, section properties,
page-number fields, list-marker formatting. So the scripts unzip the document,
edit the XML with `lxml`, and zip it back up.

That also means they only touch formatting. **Your text is never retyped, so it
cannot drift.**

---

## The TU rules these enforce

From the Faculty of Management guideline for BBS project report writing:

| Setting | Requirement |
|---|---|
| Paper | White, 8½ × 11 inch |
| Font | Times New Roman, 12pt |
| Line spacing | 1.5, justified |
| Margins | At least 1 inch on all sides |
| Preliminary pages | Lower-case roman, starting at **ii** on the Declaration, centred in footer |
| Body pages | Arabic, restarting at **1**, upper right corner |
| Tables | Number and caption **above** the table |
| Figures | Number and caption **below** the figure |
| Bibliography | APA style, hanging indent |
| Length | ~8,000–10,000 words (roughly 30–35 pages) |

Chapter structure is fixed at three: Introduction, Results and Analysis, and
Summary and Conclusion. Note that **Statement of the Problem is not part of the
TU Chapter 1 specification** — the listed contents are background, profile of
the organisation, objectives, rationale, method, review of literature, and
limitations.

---

## Things the guideline doesn't cover

Some requirements are set by your campus, not by the Faculty. Ask your
supervisor rather than guessing:

- **Cover paper colour.** Often assigned by concentration group. One widely
  circulated scheme is light orange for Accounting, light green for Finance,
  light blue for Management and yellow for Marketing — but campuses vary, and
  it is not in the official guideline.
- **Letterhead.** TU specifies campus letterhead for the Supervisor's
  Recommendation and Endorsement pages. Your supervisor supplies the sheets;
  you print those two pages onto them separately.
- **Caption placement.** Some supervisors ask for table captions *below* the
  table, which contradicts the guideline. Follow your supervisor.

---

## Before you submit

The full checklist is at the end of the troubleshooting guide. The three most
often missed:

1. **Verify TOC page numbers against the actual document, last.** They're typed
   by hand and every edit shifts them.
2. **Recompute each ratio from the figures printed beside it.** Evaluators do.
3. **Check every author cited in the text appears in the Bibliography.**
   References carries marks.

---

## Roadmap

- [x] `fix_fonts.py` — normalise all fonts
- [x] `troubleshooting.md` — symptom-indexed reference
- [ ] `fix_page_numbers.py` — PAGE fields, roman/arabic sections
- [ ] `fix_toc.py` — left-align entries, strip stray tabs, set tab stops
- [ ] `fix_spacing.py` — remove empty spacing paragraphs, apply keepNext
- [ ] `verify_pages.py` — check TOC against real pagination
- [ ] `assets/tu-template.docx` — clean starting template
- [ ] `SKILL.md` — packaged as a Claude Skill

---

## Contributing

If you hit a formatting bug that isn't documented here, please
[open an issue](../../issues) describing **what you saw** — the symptom, not the
cause. That's what helps the next student find it.

Corrections to the TU rules table are especially welcome if your campus
interprets something differently.

---

## A note on academic honesty

This repository contains formatting tools and a blank template. It deliberately
contains **no sample report content**.

Copying another student's report — in whole or in part — makes your submission
unacceptable, and under TU rules a degree can be quashed retroactively if
plagiarism is found after it has been awarded. Use these tools on **your own
work**.

---

## Licence

MIT — free to use, copy, modify and share.

Built by **Dipen Ad**, BBS 2026, after spending far too long fixing these problems by hand.
