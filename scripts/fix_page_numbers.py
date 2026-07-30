#!/usr/bin/env python3
"""
fix_page_numbers.py - repair page numbering in a .docx

Symptoms this fixes:
  * every preliminary page shows the same number (usually "II")
  * roman numerals rendering as 1, 2, 3
  * numbering that breaks again each time the file passes through Google Docs

Cause: the footer holds *literal text* instead of a PAGE field, and/or the
preliminary section is missing w:fmt="lowerRoman". Online converters strip both.

TU requires:
  preliminary pages  lower-case roman, starting at "ii" on the Declaration,
                     centred in the footer, title page unnumbered
  body pages         arabic, restarting at 1, upper right corner

Usage:
    python fix_page_numbers.py report.docx
    python fix_page_numbers.py report.docx -o fixed.docx
    python fix_page_numbers.py report.docx --check     (report only, no changes)

Requires: lxml
"""
import argparse, os, re, shutil, sys, tempfile, zipfile
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


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


def page_field(align, seed):
    """A PAGE field must live in separate runs. Collapsing them into one run is
    what Google Docs does, and Word then renders it unreliably."""
    return (
        '<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/>'
        '<w:jc w:val="%s"/></w:pPr>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:t>%s</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' % (align, seed))


def rebuild_part(path, tag, align, seed):
    """Replace whatever is in a header/footer with a clean PAGE field."""
    if not os.path.exists(path):
        return False
    s = open(path, encoding='utf-8').read()
    if '<w:p' not in s:
        return False
    start = s.index('<w:p')
    end = s.index('</w:%s>' % tag)
    open(path, 'w', encoding='utf-8').write(s[:start] + page_field(align, seed) + s[end:])
    return True


def inspect(workdir):
    """Report what's currently wrong, without changing anything."""
    issues = []
    doc = os.path.join(workdir, 'word', 'document.xml')
    d = open(doc, encoding='utf-8').read()

    sects = re.findall(r'<w:sectPr.*?</w:sectPr>', d, re.S)
    if len(sects) < 2:
        issues.append('only %d section(s) found — a report needs two, '
                      'one for preliminary pages and one for the body' % len(sects))
    for i, sec in enumerate(sects):
        num = re.search(r'<w:pgNumType[^/]*/>', sec)
        label = 'preliminary' if i == 0 else 'body'
        if not num:
            issues.append('%s section has no pgNumType' % label)
        elif i == 0 and 'lowerRoman' not in num.group():
            issues.append('preliminary section missing w:fmt="lowerRoman" '
                          '— roman numerals will render as 1, 2, 3')
        if i == 0 and 'titlePg' not in sec:
            issues.append('preliminary section missing titlePg '
                          '— the title page will be numbered')

    for name, tag in (('footer1.xml', 'ftr'), ('footer2.xml', 'ftr'),
                      ('footer3.xml', 'ftr'), ('header1.xml', 'hdr')):
        path = os.path.join(workdir, 'word', name)
        if not os.path.exists(path):
            continue
        s = open(path, encoding='utf-8').read()
        if 'PAGE' not in s:
            literal = re.search(r'<w:t[^>]*>([^<]{1,6})</w:t>', s)
            if literal and literal.group(1).strip():
                issues.append('%s contains literal text "%s" instead of a PAGE '
                              'field — every page will show that same value'
                              % (name, literal.group(1).strip()))
        elif s.count('<w:r>') < 3:
            issues.append('%s has a malformed PAGE field (all parts in one run)'
                          % name)
    return issues


def main():
    ap = argparse.ArgumentParser(description='Repair .docx page numbering')
    ap.add_argument('docx')
    ap.add_argument('-o', '--output', help='default: overwrite input')
    ap.add_argument('--check', action='store_true', help='report only')
    ap.add_argument('--body-position', choices=['header', 'footer'], default='header',
                    help='where body page numbers go (TU: header, upper right)')
    args = ap.parse_args()

    if not os.path.exists(args.docx):
        sys.exit('not found: %s' % args.docx)

    workdir = tempfile.mkdtemp()
    try:
        unpack(args.docx, workdir)

        issues = inspect(workdir)
        if issues:
            print('Problems found:')
            for i in issues:
                print('  - %s' % i)
        else:
            print('No page-numbering problems found.')
        if args.check:
            return
        print()

        doc = os.path.join(workdir, 'word', 'document.xml')
        d = open(doc, encoding='utf-8').read()

        # preliminary section -> lower-case roman
        before = d
        d = re.sub(r'<w:pgNumType(?![^/]*w:fmt)([^/]*)/>(\s*<w:titlePg)',
                   r'<w:pgNumType w:fmt="lowerRoman"\1/>\2', d, count=1)
        if d != before:
            print('preliminary section set to lowerRoman')

        # body section -> explicit decimal restart
        parts = d.rsplit('<w:pgNumType', 1)
        if len(parts) == 2 and 'w:fmt' not in parts[1][:60]:
            d = parts[0] + '<w:pgNumType w:fmt="decimal"' + parts[1]
            print('body section set to decimal, restarting at 1')

        open(doc, 'w', encoding='utf-8').write(d)

        wdir = os.path.join(workdir, 'word')
        if rebuild_part(os.path.join(wdir, 'footer1.xml'), 'ftr', 'center', 'ii'):
            print('footer1 rebuilt as a centred PAGE field')
        if args.body_position == 'header':
            if rebuild_part(os.path.join(wdir, 'header1.xml'), 'hdr', 'right', '1'):
                print('header1 rebuilt as a right-aligned PAGE field')
        else:
            if rebuild_part(os.path.join(wdir, 'footer3.xml'), 'ftr', 'center', '1'):
                print('footer3 rebuilt as a centred PAGE field')

        out = args.output or args.docx
        repack(workdir, out)
        print('\nwrote %s' % out)
        print('Open in Word and press Ctrl+A then F9 to refresh the fields.')
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


if __name__ == '__main__':
    main()
