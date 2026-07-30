# Troubleshooting

Formatting bugs found in real Tribhuvan University BBS project reports, listed
by **what you see** rather than what causes it — because that's how you'll be
searching at 2am the night before submission.

Every problem here came from a template passed between students. None of them
are your fault, and none are fixed by retyping the text.

---

## Quick index

| What you're seeing | Jump to |
|---|---|
| Every page shows the same number | [Page numbering](#page-numbering) |
| Roman numerals showing as 1, 2, 3 | [Page numbering](#page-numbering) |
| Page numbers vanished after editing in Google Docs | [Page numbering](#page-numbering) |
| Some paragraphs look bigger than others | [Fonts](#fonts) |
| Minus signs look wrong in tables | [Fonts](#fonts) |
| Bullets smaller than the text | [Fonts](#fonts) |
| Words squished together in the Table of Contents | [Table of contents](#table-of-contents) |
| Huge gaps between words in the Bibliography | [Justification](#justification) |
| Page numbers in the TOC are wrong | [Table of contents](#table-of-contents) |
| Heading alone at the bottom of a page | [Orphans and widows](#orphans-and-widows) |
| Caption separated from its table | [Orphans and widows](#orphans-and-widows) |
| Tables sticking out past the right margin | [Tables](#tables) |
| Tables splitting across two pages | [Tables](#tables) |
| Big empty gaps in the middle of pages | [Spacing](#spacing) |
| Blank pages appearing from nowhere | [Spacing](#spacing) |

---

## Page numbering

### Every page shows the same number

**What you see:** The Declaration, Supervisor's Recommendation, Endorsement and
every other preliminary page all display `II`. Or all display `2`.

**Cause:** The footer contains the *literal text* "II" rather than a page-number
field. Somebody typed it once and it repeats on every page using that footer.

**How to confirm it:** Open `word/footer1.xml` and look for `<w:t>II</w:t>`.
A real page number looks like `<w:instrText> PAGE </w:instrText>` instead.

**Fix:** Replace the literal text with a PAGE field built from separate runs:

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>ii</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

In Word directly: delete the text, then Insert → Page Number.

---

### Roman numerals showing as 1, 2, 3

**Cause:** The preliminary section's `pgNumType` has no format attribute, so it
falls back to decimal.

**Fix:** In that section's `sectPr`:

```xml
<w:pgNumType w:fmt="lowerRoman" w:start="1"/>
```

TU requires lower-case roman numerals on preliminary pages, starting at **ii**
on the Declaration. The title page counts as i but is not numbered — that needs
`<w:titlePg w:val="1"/>` in the same section.

---

### Page numbers vanished after editing in Google Docs

**Cause:** Round-tripping a .docx through Google Docs or some online converters
strips `w:fmt="lowerRoman"` and collapses PAGE fields into a single malformed
run with `begin`, `instrText`, `separate` and `end` all inside one `<w:r>`.

**Fix:** Re-apply both fixes above. Then **stop editing in Google Docs** — it
will undo them again. Use Word or LibreOffice for the final stages.

---

## Fonts

### Some paragraphs look bigger than others

**What you see:** One paragraph appears heavier or taller than the one below it,
even though selecting both shows 12pt.

**Cause:** Multiple fonts in one document. Real example from a submitted report:

| Font | Where it was hiding |
|---|---|
| Georgia | the document default |
| Gungsuh (a Korean serif) | 95 runs of pasted body text |
| Cardo | embedded in the file, on 5 runs |
| Arial | every bullet and numbered list marker |
| Calibri / Cambria | the theme's major and minor fonts |

At the same point size these have different x-heights, so they *look* like
different sizes.

**Fix:** Rewrite every `rFonts` element across `theme1.xml`, `styles.xml`
(including `docDefaults`), `document.xml`, `numbering.xml`, and all headers and
footers. Missing any one of them leaves the inconsistency visible.

Use `scripts/fix_fonts.py`.

---

### Minus signs look wrong in tables

**What you see:** In trend calculation tables, `−0.43` renders in a sans-serif
face while the numbers beside it are serif.

**Cause:** Glyph fallback. The run's font doesn't contain U+2212 (the true minus
sign), so the renderer substitutes another face for that single character.

**Fix:** Normalise fonts as above. Times New Roman contains U+2212, so the
fallback stops once the run font is correct.

---

### Bullets smaller than the text they label

**Cause:** `numbering.xml` carries its own `sz` value, independent of your body
text. A common inherited value is 22 (11pt) against 24 (12pt) body text.

**Fix:** Set every `sz` and `szCs` in `numbering.xml` to match your body size.

---

## Table of contents

### Words squished together

**What you see:** `TableofContents`, `1.1BackgroundoftheStudy`,
`CHAPTERI:INTRODUCTION` — spaces gone entirely.

**Cause:** The paragraph is **justified** *and* contains a tab stop. Justification
stretches a line to both margins; when a tab already pushes the page number to
the right edge, the renderer compensates by crushing inter-word spacing to zero.

**Fix:** Two changes together —

1. Change `<w:jc w:val="both"/>` to `<w:jc w:val="left"/>`
2. Set one right-aligned tab stop at the text width

```xml
<w:tabs><w:tab w:val="right" w:leader="none" w:pos="8640"/></w:tabs>
```

**Getting the tab position right:** it equals page width minus both margins.
For US Letter with a 1.5in binding margin: 12240 − 2160 − 1440 = **8640** twips.
A common wrong value is 9072, which overshoots the margin.

---

### Stray leading tabs

**What you see:** Long entries like "Supervisor's Recommendation" wrap onto two
lines for no obvious reason.

**Cause:** A tab character *before* the label, so the text starts at a tab stop
instead of the margin and has less room.

**Fix:** Rebuild each entry as a single run containing exactly:
`<w:t>label</w:t><w:tab/><w:t>number</w:t>` — no leading or trailing tabs.

---

### Page numbers in the TOC are wrong

**Cause:** Nobody updated them after the last edit. This is the single most
common error in submitted reports, because the numbers are typed by hand and
every edit shifts them.

**Fix:** Verify against reality, don't trust the file. Render to PDF, extract
which page each heading actually falls on, and compare against what the TOC
claims. Use `scripts/verify_pages.py`.

Do this **last**, after every other change. Any edit invalidates it.

---

## Justification

### Huge gaps between words in the Bibliography

**What you see:**
`Practitioner's     Guide     (5th     ed.).     John     Wiley     &     Sons.`

**Cause:** Justified alignment on reference entries. APA references end in long
unbreakable URLs, so the renderer has nowhere to absorb slack and stretches the
word spacing instead.

**Fix:** Left-align the entries, keep the hanging indent:

```xml
<w:jc w:val="left"/>
<w:ind w:left="720" w:hanging="720"/>
```

This is also what APA actually specifies — reference lists are flush left.

---

## Orphans and widows

### Heading stranded at the bottom of a page

**What you see:** `2.1.1 Current Ratio` sits alone at the foot of a page and its
paragraph starts overleaf.

**Cause:** That heading style is missing `keepNext`. Check *every* level —
in one real report Headings 1, 2 and 3 had it and Heading 4 did not, so only
third-level sub-headings ever stranded.

**Fix:** In `styles.xml`, for each heading style:

```xml
<w:keepNext w:val="1"/><w:keepLines w:val="1"/>
```

---

### Caption separated from its table

**Cause:** No keep relationship between the caption paragraph and the table.

**Fix depends on where the caption sits:**

- **Caption above the table** (TU standard) — put `keepNext` on the caption
- **Caption below the table** — put `keepNext` on the paragraphs in the table's
  **last row only**

**Do not** put `keepNext` on every row. That makes the whole table unbreakable,
which pushes it wholesale onto the next page and leaves large white gaps. Tested
on a 43-page report: locking all rows added a page and left six pages under 62%
full.

---

## Tables

### Sticking out past the right margin

**Cause:** Table width was set for 1in side margins, but the document uses a
1.5in binding margin. A table declared 9360 twips (6.5in) inside an 8640-twip
(6.0in) text area overflows by half an inch.

**Fix:** Scale `gridCol` widths proportionally so the total plus `tblInd` is at
most your text width, and update each cell's `tcW` to match.

**Check your text width:** page width − left margin − right margin. Letter with
1.5in binding: 12240 − 2160 − 1440 = **8640** twips.

---

### Splitting across two pages

TU says: *"Keep tables from breaking across pages unless the table is too large
for a single page."*

**Fix:** `<w:cantSplit w:val="1"/>` in each row's `trPr` stops rows splitting
mid-height. Keeping a whole table together needs `keepNext` on all rows — but
read the warning above first, because it costs pages.

**Judgement call:** if a table genuinely can't fit alongside the text above it,
splitting is the lesser evil. Compare the page count and the amount of white
space both ways before deciding.

---

## Spacing

### Big empty gaps in the middle of pages

**Cause:** Empty paragraphs used as manual spacing, stacked on top of the
spacing the heading and table styles already apply. One report carried **40** of
them, each adding roughly 18pt.

**How to spot them:** paragraphs with no text, no image and no break, sitting
immediately before a heading or a table.

**Fix:** Delete them and let the styles handle spacing. Removing 40 from a
45-page report recovered a full page and eliminated 36 of 37 oversized gaps.

---

### Blank pages appearing from nowhere

**Cause:** A run of empty paragraphs pushes content past the page boundary, and
the next paragraph then begins with a page break — producing a page containing
nothing but a footer.

**Fix:** Delete the empty paragraphs. Don't delete the page break; it's doing
its job.

---

## Before you submit

Run these checks in order. Each one takes a minute and catches something an
evaluator would.

- [ ] Page numbers run i–x then 1–n with no gaps or repeats
- [ ] Every TOC entry matches where the section actually falls
- [ ] Every List of Tables / Figures entry matches its caption's real page
- [ ] Each ratio recomputes correctly from the inputs printed beside it
- [ ] Every author cited in the text appears in the Bibliography
- [ ] Every abbreviation listed is actually used in the body
- [ ] No heading alone at the bottom of a page
- [ ] No caption separated from its table or figure
- [ ] One typeface throughout
- [ ] Nothing crossing the margins
- [ ] No blank pages

The arithmetic check matters more than it sounds. In one submitted report two
of forty ratios didn't reconcile with the figures printed in the same table —
recomputable by anyone with a calculator, in a document that had been read by
several people.
