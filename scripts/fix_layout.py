#!/usr/bin/env python3
"""
fix_layout.py - repair TOC spacing, orphaned headings and stray padding

Three separate bugs, all of which come bundled in inherited BBS templates:

1. TABLE OF CONTENTS WORD SPACING
   Entries are justified AND contain a tab stop. Justification stretches a line
   to both margins; when a tab already pushes the page number to the right edge
   there is nothing left to stretch, so the renderer crushes inter-word spacing
   to zero. You get "TableofContents" and "1.1BackgroundoftheStudy".
   Fix: left-align the entries, one right tab stop at the text width.

2. ORPHANED HEADINGS
   A heading style missing keepNext lets the heading sit alone at the foot of a
   page while its paragraph starts overleaf. Commonly only one level is
   affected — check all four.

3. EMPTY SPACING PARAGRAPHS
   Blank paragraphs used as manual spacing, stacked on top of the spacing the
   styles already apply. One real report carried forty of them.

Usage:
    python fix_layout.py report.docx
    python fix_layout.py report.docx --text-width 8640
    python fix_layout.py report.docx --skip-toc --skip-empties

Requires: lxml
"""
import argparse, os, re, shutil, sys, tempfile, zipfile
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XML = '{http://www.w3.org/XML/1998/namespace}'

# schema-mandated order of pPr children; getting this wrong makes Word reject the file
PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
             'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd',
             'tabs', 'suppressAutoHyphens', 'bidi', 'spacing', 'ind',
             'contextualSpacing', 'jc', 'textDirection', 'textAlignment',
             'outlineLvl', 'rPr', 'sectPr']


def unpack(docx, workdir):
    with zipfile.ZipFile(docx) as z:
        z.extractall(workdir)


def repack(workdir, out):
    if os.path.exists(out):
        os.remove(out)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(workdir):
            for name in files:
                full = os.path.join(root, name)
                z.write(full, os.path.relpath(full, workdir))


def text_of(el):
    return ''.join(t.text or '' for t in el.iter(W + 't')).strip()


def style_of(p):
    st = p.find(W + 'pPr/' + W + 'pStyle')
    return st.get(W + 'val') if st is not None else ''


def reorder(tree):
    """Sort every pPr's children into schema order."""
    fixed = 0
    for pPr in tree.getroot().iter(W + 'pPr'):
        kids = list(pPr)
        names = [etree.QName(k).localname for k in kids]
        want = sorted(kids, key=lambda k: PPR_ORDER.index(etree.QName(k).localname)
                      if etree.QName(k).localname in PPR_ORDER else len(PPR_ORDER))
        if [etree.QName(k).localname for k in want] != names:
            for k in kids:
                pPr.remove(k)
            for k in want:
                pPr.append(k)
            fixed += 1
    return fixed


def detect_text_width(workdir):
    """text width = page width - left margin - right margin, in twips"""
    d = open(os.path.join(workdir, 'word', 'document.xml'), encoding='utf-8').read()
    sz = re.search(r'<w:pgSz[^>]*w:w="(\d+)"', d)
    mar = re.search(r'<w:pgMar[^>]*w:left="([\d.]+)"[^>]*w:right="([\d.]+)"', d)
    if not mar:
        mar = re.search(r'<w:pgMar[^>]*w:right="([\d.]+)"[^>]*w:left="([\d.]+)"', d)
        if mar:
            right, left = mar.groups()
        else:
            return 8640
    else:
        left, right = mar.groups()
    width = int(sz.group(1)) if sz else 12240
    return int(width - float(left) - float(right))


def fix_toc(body, tab_pos):
    """Left-align list/TOC entries and give them one right tab stop."""
    n = 0
    for p in body.iter(W + 'p'):
        pPr = p.find(W + 'pPr')
        if pPr is None or pPr.find(W + 'tabs') is None:
            continue
        ts = [t for t in p.iter(W + 't')]
        if len(ts) < 2:
            continue

        jc = pPr.find(W + 'jc')
        if jc is None:
            jc = etree.SubElement(pPr, W + 'jc')
        jc.set(W + 'val', 'left')

        for old in pPr.findall(W + 'tabs'):
            pPr.remove(old)
        tabs = etree.SubElement(pPr, W + 'tabs')
        tab = etree.SubElement(tabs, W + 'tab')
        tab.set(W + 'val', 'right')
        tab.set(W + 'leader', 'none')
        tab.set(W + 'pos', str(tab_pos))

        # rebuild as exactly: label <tab> number, no stray leading/trailing tabs
        label = (ts[0].text or '').strip()
        number = (ts[-1].text or '').strip()
        runs = p.findall(W + 'r')
        keep = next((r for r in runs if r.find(W + 't') is not None), None)
        if keep is None:
            continue
        for r in runs:
            if r is not keep:
                p.remove(r)
        for t in keep.findall(W + 't'):
            keep.remove(t)
        for tb in keep.findall(W + 'tab'):
            keep.remove(tb)
        t1 = etree.SubElement(keep, W + 't')
        t1.set(XML + 'space', 'preserve')
        t1.text = label
        etree.SubElement(keep, W + 'tab')
        t2 = etree.SubElement(keep, W + 't')
        t2.set(XML + 'space', 'preserve')
        t2.text = number
        n += 1
    return n


def fix_headings(styles):
    """Every heading level needs keepNext, or it will strand at a page bottom."""
    n = 0
    for st in styles.getroot().iter(W + 'style'):
        sid = st.get(W + 'styleId') or ''
        if not sid.startswith('Heading'):
            continue
        pPr = st.find(W + 'pPr')
        if pPr is None:
            pPr = etree.SubElement(st, W + 'pPr')
        for tag in ('keepLines', 'keepNext'):
            if pPr.find(W + tag) is None:
                el = etree.Element(W + tag)
                el.set(W + 'val', '1')
                pPr.insert(0, el)
                n += 1
    return n


def drop_empties(body):
    """Delete blank paragraphs sitting immediately before a heading or table."""
    els = list(body)
    doomed = []
    for i, e in enumerate(els):
        if etree.QName(e).localname != 'p':
            continue
        if text_of(e) or e.find('.//' + W + 'drawing') is not None \
           or e.find('.//' + W + 'br') is not None \
           or e.find(W + 'pPr/' + W + 'sectPr') is not None:
            continue
        for j in range(i + 1, len(els)):
            nxt = els[j]
            if etree.QName(nxt).localname == 'tbl':
                doomed.append(e)
                break
            if etree.QName(nxt).localname != 'p':
                break
            if not text_of(nxt):
                continue
            if style_of(nxt).startswith('Heading'):
                doomed.append(e)
            break
    for e in doomed:
        body.remove(e)
    return len(doomed)


def main():
    ap = argparse.ArgumentParser(description='Fix TOC spacing, orphans and padding')
    ap.add_argument('docx')
    ap.add_argument('-o', '--output')
    ap.add_argument('--text-width', type=int,
                    help='twips; default is detected from the page setup')
    ap.add_argument('--skip-toc', action='store_true')
    ap.add_argument('--skip-headings', action='store_true')
    ap.add_argument('--skip-empties', action='store_true')
    args = ap.parse_args()

    if not os.path.exists(args.docx):
        sys.exit('not found: %s' % args.docx)

    workdir = tempfile.mkdtemp()
    try:
        unpack(args.docx, workdir)
        doc_path = os.path.join(workdir, 'word', 'document.xml')
        sty_path = os.path.join(workdir, 'word', 'styles.xml')

        width = args.text_width or detect_text_width(workdir)
        print('text width ............... %d twips (%.2f in)' % (width, width / 1440))

        tree = etree.parse(doc_path)
        body = tree.getroot().find(W + 'body')

        if not args.skip_toc:
            print('TOC/list entries fixed ... %d' % fix_toc(body, width))
        if not args.skip_empties:
            print('empty paragraphs removed . %d' % drop_empties(body))
        reorder(tree)
        tree.write(doc_path, xml_declaration=True, encoding='UTF-8', standalone=True)

        if not args.skip_headings:
            styles = etree.parse(sty_path)
            print('heading keep settings .... %d added' % fix_headings(styles))
            reorder(styles)
            styles.write(sty_path, xml_declaration=True, encoding='UTF-8', standalone=True)

        out = args.output or args.docx
        repack(workdir, out)
        print('\nwrote %s' % out)
        print('Now run verify_pages.py — page numbers will have shifted.')
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


if __name__ == '__main__':
    main()
