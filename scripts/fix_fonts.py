#!/usr/bin/env python3
"""
fix_fonts.py - force a single typeface throughout a .docx

Inherited BBS templates commonly carry four or five fonts at once (Georgia as
the default, Gungsuh on pasted runs, an embedded Cardo, Arial on list markers,
Calibri/Cambria in the theme). At the same point size these have different
x-heights, so paragraphs look inconsistently sized even though every run says
12pt. This normalises every font reference the file contains.

Usage:
    python fix_fonts.py report.docx
    python fix_fonts.py report.docx --font "Times New Roman" --size 12
    python fix_fonts.py report.docx -o fixed.docx

Requires: lxml   (pip install lxml)
"""
import argparse, os, shutil, sys, tempfile, zipfile
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
FONT_ATTRS = ('ascii', 'hAnsi', 'cs', 'eastAsia')


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


def load(path):
    return etree.parse(path)


def save(tree, path):
    tree.write(path, xml_declaration=True, encoding='UTF-8', standalone=True)


def set_theme(workdir, font):
    path = os.path.join(workdir, 'word', 'theme', 'theme1.xml')
    if not os.path.exists(path):
        return 0
    tree = load(path)
    n = 0
    for kind in ('majorFont', 'minorFont'):
        node = tree.getroot().find('.//' + A + kind)
        if node is None:
            continue
        latin = node.find(A + 'latin')
        if latin is not None and latin.get('typeface') != font:
            latin.set('typeface', font)
            n += 1
    save(tree, path)
    return n


def set_rfonts(path, font):
    """Rewrite every rFonts element in one part."""
    if not os.path.exists(path):
        return 0
    tree = load(path)
    n = 0
    for rf in tree.getroot().iter(W + 'rFonts'):
        for attr in FONT_ATTRS:
            if rf.get(W + attr) != font:
                rf.set(W + attr, font)
                n += 1
    save(tree, path)
    return n


def set_default(workdir, font, half_points):
    """docDefaults is what every unstyled run inherits from."""
    path = os.path.join(workdir, 'word', 'styles.xml')
    tree = load(path)
    rpr = tree.getroot().find(W + 'docDefaults/' + W + 'rPrDefault/' + W + 'rPr')
    if rpr is None:
        return False
    if rpr.find(W + 'rFonts') is None:
        rf = etree.Element(W + 'rFonts')
        for attr in FONT_ATTRS:
            rf.set(W + attr, font)
        rpr.insert(0, rf)
    for tag in ('sz', 'szCs'):
        el = rpr.find(W + tag)
        if el is None:
            el = etree.SubElement(rpr, W + tag)
        el.set(W + 'val', str(half_points))
    save(tree, path)
    return True


def match_marker_size(workdir, half_points):
    """List bullets/numbers often sit a point smaller than the text they label."""
    path = os.path.join(workdir, 'word', 'numbering.xml')
    if not os.path.exists(path):
        return 0
    tree = load(path)
    n = 0
    for tag in ('sz', 'szCs'):
        for el in tree.getroot().iter(W + tag):
            if el.get(W + 'val') != str(half_points):
                el.set(W + 'val', str(half_points))
                n += 1
    save(tree, path)
    return n


def drop_embedded(workdir, font):
    """Remove embedded font files; they bloat the docx and are now unused."""
    fonts_dir = os.path.join(workdir, 'word', 'fonts')
    removed = 0
    if os.path.isdir(fonts_dir):
        removed = len(os.listdir(fonts_dir))
        shutil.rmtree(fonts_dir)

    table = os.path.join(workdir, 'word', 'fontTable.xml')
    if os.path.exists(table):
        ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        with open(table, 'w', encoding='utf-8') as f:
            f.write(
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r'
                '<w:fonts xmlns:w="%s" xmlns:r="%s">'
                '<w:font w:name="%s"><w:charset w:val="00"/>'
                '<w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>'
                '</w:fonts>' % (ns, rel, font))

    rels = os.path.join(workdir, 'word', '_rels', 'fontTable.xml.rels')
    if os.path.exists(rels):
        with open(rels, 'w', encoding='utf-8') as f:
            f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/'
                    'package/2006/relationships"/>')
    return removed


def main():
    ap = argparse.ArgumentParser(description='Normalise all fonts in a .docx')
    ap.add_argument('docx')
    ap.add_argument('-o', '--output', help='default: overwrite input')
    ap.add_argument('--font', default='Times New Roman')
    ap.add_argument('--size', type=float, default=12.0, help='points')
    args = ap.parse_args()

    if not os.path.exists(args.docx):
        sys.exit('not found: %s' % args.docx)

    half = int(args.size * 2)          # OOXML stores half-points
    out = args.output or args.docx
    workdir = tempfile.mkdtemp()

    try:
        unpack(args.docx, workdir)

        print('theme fonts .............', set_theme(workdir, args.font))
        print('docDefaults .............', 'set' if set_default(workdir, args.font, half) else 'not found')

        total = 0
        for part in ('styles.xml', 'document.xml', 'numbering.xml',
                     'header1.xml', 'header2.xml', 'header3.xml',
                     'footer1.xml', 'footer2.xml', 'footer3.xml'):
            total += set_rfonts(os.path.join(workdir, 'word', part), args.font)
        print('font attributes rewritten', total)
        print('list marker size ........', match_marker_size(workdir, half))
        print('embedded fonts removed ..', drop_embedded(workdir, args.font))

        repack(workdir, out)
        print('\nwrote %s  (all text now %s %gpt)' % (out, args.font, args.size))
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


if __name__ == '__main__':
    main()
