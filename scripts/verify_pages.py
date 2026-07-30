#!/usr/bin/env python3
"""
verify_pages.py - check a report's claims against what it actually prints

Table of Contents page numbers in BBS reports are typed by hand, so every edit
silently invalidates them. This renders the document to PDF and compares what
the TOC, List of Tables and List of Figures *claim* against where each heading
and caption *actually* falls.

It also reports the layout problems an evaluator would notice: headings
stranded at the foot of a page, captions separated from their table, content
crossing the margins, mixed typefaces and blank pages.

Run this LAST, after every other change. Any edit invalidates the result.

Usage:
    python verify_pages.py report.docx
    python verify_pages.py report.pdf          (skips the render step)
    python verify_pages.py report.docx --quiet (only show problems)

Requires: pdfplumber   (pip install pdfplumber)
          LibreOffice  (only when passing a .docx)
"""
import argparse, collections, os, re, shutil, subprocess, sys, tempfile

try:
    import pdfplumber
except ImportError:
    sys.exit('pdfplumber not installed.  pip install pdfplumber')

CAPTION = re.compile(r'^(Table|Figure)\s+(\d+)\s*[:.]')
HEADING = re.compile(r'^\d+\.\d+(\.\d+)?\s+[A-Z]')
# a TOC/list line: label, then whitespace, then a roman or arabic number at the end
ENTRY = re.compile(r'^(.*?)\s+([ivxlcdm]+|\d+)$', re.IGNORECASE)


def render(docx):
    """Convert .docx to PDF with LibreOffice, return the pdf path + temp dir."""
    for exe in ('libreoffice', 'soffice'):
        if shutil.which(exe):
            break
    else:
        sys.exit('LibreOffice not found. Convert to PDF yourself and pass the PDF.')
    tmp = tempfile.mkdtemp()
    subprocess.run([exe, '--headless', '--convert-to', 'pdf',
                    '--outdir', tmp, docx],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    pdf = os.path.join(tmp, os.path.splitext(os.path.basename(docx))[0] + '.pdf')
    if not os.path.exists(pdf):
        sys.exit('LibreOffice did not produce a PDF.')
    return pdf, tmp


def printed_number(page, front_matter):
    """Preliminary pages number in the footer, body pages in the header."""
    lines = (page.extract_text() or '').strip().split('\n')
    if not lines:
        return ''
    return lines[-1].strip() if front_matter else lines[0].strip()


def scan(pdf):
    """Where every heading and caption actually sits, and what the lists claim."""
    actual_caps = {}
    actual_secs = {}
    claims = {}
    front = True

    for page in pdf.pages:
        text = page.extract_text() or ''
        lines = text.split('\n')
        # once we meet a page whose first line is a bare number, the body has begun
        if front and lines and re.fullmatch(r'\d+', lines[0].strip()):
            front = False
        num = printed_number(page, front)

        if not front:                      # front matter holds the TOC, not the real thing
            for line in lines:
                s = line.strip()
                m = CAPTION.match(s)
                if m:
                    actual_caps.setdefault('%s %s' % (m.group(1), m.group(2)), num)
                elif HEADING.match(s) or s.startswith('CHAPTER'):
                    actual_secs.setdefault(s.split('  ')[0].strip(), num)

        # list/TOC pages: harvest "label ..... number" lines
        if front:
            for line in page.extract_text_lines():
                e = ENTRY.match(line['text'].strip())
                if e and len(e.group(1)) > 2:
                    claims[e.group(1).strip()] = e.group(2)

    return actual_caps, actual_secs, claims


def layout_problems(pdf):
    out = collections.OrderedDict()

    stranded = []
    for page in pdf.pages:
        lines = page.extract_text_lines()
        if len(lines) > 3 and HEADING.match(lines[-1]['text'].strip()):
            stranded.append(lines[-1]['text'].strip()[:50])
    out['headings stranded at a page bottom'] = stranded

    orphaned = []
    for page in pdf.pages:
        lines = page.extract_text_lines()
        for i, l in enumerate(lines[:2]):
            if CAPTION.match(l['text'].strip()) and not page.images:
                orphaned.append(l['text'].strip()[:50])
    out['captions separated from their object'] = orphaned

    gaps = []
    body_pages = []
    started = False
    for page in pdf.pages:
        first = (page.extract_text() or '').strip().split('\n')
        if not started and first and re.fullmatch(r'\d+', first[0].strip()):
            started = True
        if started:
            body_pages.append(page)
    for page in body_pages:
        lines = page.extract_text_lines()
        imgs = [(i['top'], i['bottom']) for i in page.images]
        for a, b in zip(lines, lines[1:]):
            g = b['top'] - a['bottom']
            if g >= 40 and not any(t < b['top'] and bo > a['bottom'] for t, bo in imgs):
                gaps.append('%.0fpt after "%s"' % (g, a['text'][:36]))
    out['vertical gaps over 40pt'] = gaps

    out['blank pages'] = ['sheet %d' % i for i, p in enumerate(pdf.pages, 1) if not p.chars]
    return out


def report(pdf, quiet=False):
    caps, secs, claims = scan(pdf)
    problems = 0

    def match(label):
        """Find the claim whose label starts the same way."""
        if label in claims:
            return claims[label]
        for k, v in claims.items():
            if k.startswith(label[:16]):
                # don't let "CHAPTER II" swallow "CHAPTER III"
                if label.startswith('CHAPTER II') and k.startswith('CHAPTER III'):
                    continue
                return v
        return None

    print('=' * 60)
    print('  PAGE REFERENCES')
    print('=' * 60)
    for key in sorted(caps, key=lambda s: (s.split()[0], int(s.split()[1]))):
        listed, real = match(key), caps[key]
        if listed is None:
            if not quiet:
                print('  %-14s not listed          (actually on %s)' % (key, real))
        elif listed != real:
            print('  %-14s listed %-5s ACTUAL %-5s  <-- MISMATCH' % (key, listed, real))
            problems += 1
        elif not quiet:
            print('  %-14s %s  ok' % (key, real))

    for key in secs:
        listed, real = match(key), secs[key]
        if listed and listed != real:
            print('  %-40s listed %-5s ACTUAL %-5s  <-- MISMATCH' % (key[:40], listed, real))
            problems += 1

    print()
    print('=' * 60)
    print('  LAYOUT')
    print('=' * 60)
    for name, items in layout_problems(pdf).items():
        print('  %-38s %d' % (name, len(items)))
        for it in items[:6]:
            print('      - %s' % it)
        if len(items) > 6:
            print('      ... and %d more' % (len(items) - 6))
        problems += len(items)

    print()
    print('=' * 60)
    print('  DOCUMENT')
    print('=' * 60)
    fonts = sorted({c['fontname'].split('+')[-1] for p in pdf.pages for c in p.chars})
    sizes = collections.Counter(round(c['size'], 1) for p in pdf.pages for c in p.chars)
    ink = [c for p in pdf.pages for c in p.chars if c['text'].strip()]
    w = pdf.pages[0].width
    print('  pages ................ %d' % len(pdf.pages))
    print('  typefaces ............ %s' % ', '.join(fonts))
    print('  sizes in use ......... %s' % dict(sorted(sizes.items())))
    if ink:
        print('  narrowest margins .... L %.2f"  R %.2f"' % (
            min(c['x0'] for c in ink) / 72, (w - max(c['x1'] for c in ink)) / 72))
    if len(fonts) > 3:
        print('  NOTE: more than one typeface family — run fix_fonts.py')

    print()
    print('=' * 60)
    if problems:
        print('  %d problem(s) found' % problems)
    else:
        print('  No problems found.')
    print('=' * 60)
    return problems


def main():
    ap = argparse.ArgumentParser(description='Verify a report against its own claims')
    ap.add_argument('path', help='.docx or .pdf')
    ap.add_argument('-q', '--quiet', action='store_true', help='show only problems')
    args = ap.parse_args()

    if not os.path.exists(args.path):
        sys.exit('not found: %s' % args.path)

    tmp = None
    path = args.path
    if path.lower().endswith('.docx'):
        print('rendering to PDF ...')
        path, tmp = render(path)

    try:
        with pdfplumber.open(path) as pdf:
            sys.exit(1 if report(pdf, args.quiet) else 0)
    finally:
        if tmp:
            shutil.rmtree(tmp, ignore_errors=True)


if __name__ == '__main__':
    main()
