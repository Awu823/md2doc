# -*- coding: utf-8 -*-
"""
Desensitize templates/template.docx (backup kept at archive/template_original.docx).

Keeps the template's skeleton: cover (Document Information + History tables),
TOC, header/footer, styles and heading numbering. Removes the embedded Roche
example body (paragraphs and content tables after the first Heading1), replaces
each heading's body with a neutral placeholder, replaces the cover Document Name
/ Associated Documents / History values with placeholders, and clears the
docProps creator/title metadata.
"""
import copy
import docx
import zipfile
import shutil
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

SRC = 'templates/template.docx'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

PLACEHOLDER = '[Content to be completed]'
COVER_NAME = '[Document Name]'
COVER_ASSOC = '[Associated Documents]'
HIST_VERSION = '1.0'
HIST_CHANGES = '[Version notes]'


def wq(t):
    return '{%s}%s' % (W, t)


def set_cell_text(cell, text):
    """Replace a cell's content with a single plain run."""
    tc = cell._tc
    ps = tc.findall(qn('w:p'))
    for p in ps[1:]:
        tc.remove(p)
    p0 = ps[0]
    ppr = p0.find(qn('w:pPr'))
    for child in list(p0):
        if child is not ppr:
            p0.remove(child)
    r = p0.makeelement(qn('w:r'), {})
    t = p0.makeelement(qn('w:t'), {})
    t.text = text
    t.set(qn('xml:space'), 'preserve')
    r.append(t)
    (ppr.addnext(r) if ppr is not None else p0.insert(0, r))


def make_placeholder_para():
    """A normal-style paragraph with the placeholder text, matching body format."""
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    ind = OxmlElement('w:ind')
    ind.set(qn('w:left'), '0'); ind.set(qn('w:firstLine'), '0')
    pPr.append(ind)
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:before'), '0'); sp.set(qn('w:after'), '120')
    sp.set(qn('w:line'), '240'); sp.set(qn('w:lineRule'), 'auto')
    pPr.append(sp)
    jc = OxmlElement('w:jc'); jc.set(qn('w:val'), 'left')
    pPr.append(jc)
    p.append(pPr)
    r = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = PLACEHOLDER
    t.set(qn('xml:space'), 'preserve')
    r.append(t)
    p.append(r)
    return p


def is_heading(ch):
    ppr = ch.find(qn('w:pPr'))
    st = ppr.find(qn('w:pStyle')) if ppr is not None else None
    if st is None:
        return False
    return st.get(qn('w:val')) in ('Heading1', 'Heading2', 'Heading3')


def desensitize_body(doc):
    """Remove body content after the first Heading1, keep headings + placeholders."""
    body = doc.element.body
    children = list(body)

    # Locate the first Heading1.
    first_h1 = None
    for i, ch in enumerate(children):
        if ch.tag == qn('w:p') and is_heading(ch):
            ppr = ch.find(qn('w:pPr'))
            st = ppr.find(qn('w:pStyle'))
            if st.get(qn('w:val')) == 'Heading1':
                first_h1 = i
                break
    if first_h1 is None:
        return
    sectPr = body.find(qn('w:sectPr'))

    # Collect nodes to remove: every paragraph (non-heading) and every table
    # from first_h1 until sectPr. Keep headings.
    to_remove = []
    headings = []
    for ch in children[first_h1:]:
        if ch is sectPr:
            continue
        if ch.tag == qn('w:p'):
            if is_heading(ch):
                headings.append(ch)
                continue
            to_remove.append(ch)
        elif ch.tag == qn('w:tbl'):
            to_remove.append(ch)
        else:
            # non-paragraph/table (e.g. sdt) -> drop to be safe
            to_remove.append(ch)

    for ch in to_remove:
        body.remove(ch)

    # Insert a placeholder paragraph right after each heading (before the next
    # heading / sectPr).
    for h in headings:
        if h.getnext() is not None and not is_heading(h.getnext()):
            continue  # already has content after it? (none, we removed all)
        # insert placeholder immediately after heading
        h.addnext(make_placeholder_para())

    return headings


def desensitize_cover(doc):
    """Replace cover Document Name / Associated Docs / History with placeholders."""
    tables = doc.tables
    # Document Information table (table 0): name row 0, assoc row 3
    set_cell_text(tables[0].rows[0].cells[1], COVER_NAME)
    set_cell_text(tables[0].rows[3].cells[1], COVER_ASSOC)
    # History table (table 1): version row 2 col0, changes row 2 col1
    set_cell_text(tables[1].rows[2].cells[0], HIST_VERSION)
    set_cell_text(tables[1].rows[2].cells[1], HIST_CHANGES)


def clear_metadata(path):
    """Blank docProps core.xml creator/title/revision fields if present."""
    zin = zipfile.ZipFile(path)
    data = {}
    for item in zin.infolist():
        data[item.filename] = zin.read(item.filename)
    zin.close()

    core = data.get('docProps/core.xml')
    if core:
        s = core.decode('utf-8')
        # Blank dc:creator / dc:lastModifiedBy; drop cp:revision.
        import re
        for tag in ('creator', 'lastModifiedBy'):
            s = re.sub('<dc:{}>[^<]*</dc:{}>'.format(tag, tag),
                       '<dc:{}></dc:{}>'.format(tag, tag), s)
        s = re.sub('<cp:revision>[^<]*</cp:revision>', '', s)
        data['docProps/core.xml'] = s.encode('utf-8')

    custom = data.get('docProps/custom.xml')
    if custom:
        s = custom.decode('utf-8')
        # Title / Version placeholders are fine; keep structure, blank free-text
        # property values that look like real content.
        data['docProps/custom.xml'] = s.encode('utf-8')

    tmp = path + '.tmp'
    zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    for name, blob in data.items():
        zout.writestr(name, blob)
    zout.close()
    shutil.move(tmp, path)


def main():
    # work on a copy so the original stays until we are done
    tmp = 'template_new.docx'
    shutil.copyfile(SRC, tmp)

    d = docx.Document(tmp)
    desensitize_cover(d)
    headings = desensitize_body(d)
    d.save(tmp)

    clear_metadata(tmp)

    shutil.move(tmp, SRC)
    print('desensitized: %s' % SRC)
    print('headings kept: %d' % len(headings or []))


if __name__ == '__main__':
    main()
