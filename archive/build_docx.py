# -*- coding: utf-8 -*-
"""
Build DS_CN_MA_Production BOM docx FROM the corporate template.docx.

Keeps the template's cover (Document Information + History), TOC page, header /
footer and styles; replaces the sample body with the new MD body. The injected
tables are restyled to match the template (full grid, grey header, tight cell
margins) with content-aware column widths and centred headers. Heading 1/2/3
are auto-numbered (1. / 1.1 / 1.1.1) to match the corporate numbering scheme.

Run:  python build_docx.py
Requires: python-docx, lxml, pandoc on PATH.
"""
import copy
import os
import shutil
import subprocess
import zipfile

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

MD = 'DS_CN_MA_Production BOM Interface From S4H To ECMS.md'
TEMPLATE = 'template.docx'
OUT = 'DS_CN_MA_Production BOM Interface From S4H To ECMS.docx'
BODY_MD = '_body.md'
CONTENT = '_content.docx'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

DOC_NAME = 'DS_CN_MA_Production BOM Interface From S4H to ECMS'
DOC_ASSOC = ('CG_CN_XF ECMS Integration  FS_CN_MA_Production BOM interface to ECMS MW.docx'
             '  FS_CN_MA_Production BOM to ECMS MW Attachment')
HISTORY = 'First Approved version'
BODY_START = 54
USABLE_WIDTH = 9639
HEADING_NUM_ID = 5

PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
             'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs',
             'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
             'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
             'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents',
             'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
             'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr']
TCPR_ORDER = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd',
              'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark',
              'headers', 'cellIns', 'cellDel', 'cellMerge', 'tcPrChange']
TBLPR_ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize',
               'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd',
               'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook',
               'tblCaption', 'tblDescription']


def wq(t):
    return '{%s}%s' % (W, t)


def raw_pstyle(p):
    """Return the paragraph's pStyle styleId (e.g. 'BlockText') or None."""
    ppr = p._p.find(qn('w:pPr'))
    st = ppr.find(qn('w:pStyle')) if ppr is not None else None
    return st.get(qn('w:val')) if st is not None else None


# pandoc styles referenced by the injected body but absent from the template's
# styles.xml. They are injected into the output so the markdown block-quote
# (Note: ...) and inline code render correctly instead of dangling.
EXTRA_STYLES = (
    '<w:style w:type="paragraph" w:styleId="BlockText">'
    '<w:name w:val="Block Text"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>'
    '<w:qFormat/>'
    '<w:pPr><w:spacing w:before="100" w:after="100"/>'
    '<w:ind w:firstLine="0" w:left="480" w:right="480"/></w:pPr></w:style>'
    '<w:style w:type="character" w:customStyle="1" w:styleId="VerbatimChar">'
    '<w:name w:val="Verbatim Char"/>'
    '<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>'
    '<w:sz w:val="22"/></w:rPr></w:style>'
)


def insert_child(parent, el, order):
    """Insert `el` into `parent` at its correct OOXML schema position."""
    tag = etree.QName(el).localname
    idx = order.index(tag)
    for child in parent:
        ctag = etree.QName(child).localname
        if ctag in order and order.index(ctag) > idx:
            child.addprevious(el)
            return
    parent.append(el)


def insert_ppr(pPr, el):
    insert_child(pPr, el, PPR_ORDER)


def insert_tcpr(tcPr, el):
    insert_child(tcPr, el, TCPR_ORDER)


def insert_tblpr(tblPr, el):
    insert_child(tblPr, el, TBLPR_ORDER)


def char_w(ch):
    if ch.isspace():
        return 0.5
    if ch in 'ilI.,:;|\'"()[]{}`|':
        return 0.5
    if ch in 'MW@%&wm':
        return 1.25
    if ch.isupper() or ch.isdigit():
        return 1.05
    return 0.9


def est_width(text):
    return sum(char_w(c) for c in text)


def distribute(weights, usable, min_w, max_w):
    n = len(weights)
    widths = [usable * w / sum(weights) for w in weights]
    for _ in range(20):
        free = [i for i in range(n) if min_w < widths[i] < max_w]
        clamped = sum(widths[i] for i in range(n) if i not in free)
        target = usable - clamped
        cur = sum(widths[i] for i in free)
        if cur <= 0:
            break
        changed = False
        for i in free:
            new = widths[i] * target / cur
            if new < min_w or new > max_w:
                widths[i] = min(max_w, max(min_w, new))
                changed = True
            else:
                widths[i] = new
        if not changed:
            break
    widths = [int(round(w)) for w in widths]
    widths[-1] += usable - sum(widths)
    return widths


def compute_col_widths(tbl, usable=USABLE_WIDTH, min_w=520, max_w=None):
    ncols = len(tbl.columns)
    max_w = max_w or int(usable * 0.55)
    weights = []
    for ci in range(ncols):
        m = 0
        for row in tbl.rows:
            cells = row.cells
            if ci < len(cells):
                m = max(m, est_width(cells[ci].text.strip()))
        weights.append(m + 12)                     # cell-margin padding
    return distribute(weights, usable, min_w, max_w)


def prepare_body():
    lines = open(MD, encoding='utf-8').read().splitlines(keepends=True)
    body = lines[BODY_START:]
    # Google-Docs-exported MD sometimes drops the blank line between a paragraph
    # and a following "* " list item; pandoc then merges them into one paragraph
    # (the "* " stays literal text). Re-insert the blank line so the list renders.
    out = []
    for i, ln in enumerate(body):
        if i > 0 and ln.lstrip().startswith('* ') and out[-1].strip() \
                and not out[-1].lstrip().startswith(('* ', '- ')):
            prev = out[-1].strip()
            if not prev.startswith(('#', '|', '>')) and not _is_ordered(prev):
                out.append('\n')
        out.append(ln)
    open(BODY_MD, 'w', encoding='utf-8').writelines(out)
    print('body md lines:', len(out))


def _is_ordered(s):
    import re
    return bool(re.match(r'^\d+[.)]\s', s))


def convert():
    env = os.environ.copy()
    env['PATH'] = (os.environ.get('PATH', '')
                   + ';' + os.environ.get('LOCALAPPDATA', '')
                   + '\\Microsoft\\WinGet\\Links'
                   + ';' + os.environ.get('LOCALAPPDATA', '')
                   + '\\Pandoc')
    pandoc = shutil.which('pandoc')
    if pandoc is None:
        cand = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Pandoc', 'pandoc.exe')
        if os.path.exists(cand):
            pandoc = cand
    if pandoc is None:
        raise SystemExit('pandoc not found on PATH')
    subprocess.run([pandoc, BODY_MD, '-f', 'markdown', '-t', 'docx',
                    '--reference-doc=' + TEMPLATE, '-o', CONTENT],
                   check=True, env=env)
    print('pandoc done ->', CONTENT)


def set_cell(cell, text):
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
    r.append(t)
    (ppr.addnext(r) if ppr is not None else p0.insert(0, r))


def fix_table(tbl):
    """Template look for a pandoc-injected table: grid, grey centred header, tight margins."""
    tblPr = tbl._tbl.tblPr
    if tblPr is None:
        return
    st = tblPr.find(qn('w:tblStyle'))
    if st is None or st.get(qn('w:val')) != 'Table':
        return

    # borders (6 edges single 4sz)
    for b in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(b)
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement('w:' + edge)
        e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), '4')
        e.set(qn('w:space'), '0'); e.set(qn('w:color'), 'auto')
        borders.append(e)
    insert_tblpr(tblPr, borders)

    # tight cell margins
    for cm in tblPr.findall(qn('w:tblCellMar')):
        tblPr.remove(cm)
    mar = OxmlElement('w:tblCellMar')
    for side, val in (('top', '0'), ('left', '115'), ('bottom', '0'), ('right', '115')):
        s = OxmlElement('w:' + side); s.set(qn('w:w'), val); s.set(qn('w:type'), 'dxa')
        mar.append(s)
    insert_tblpr(tblPr, mar)

    # fixed layout + width
    widths = compute_col_widths(tbl)
    grid = tbl._tbl.find(qn('w:tblGrid'))
    for gc, w in zip(grid.findall(qn('w:gridCol')), widths):
        gc.set(qn('w:w'), str(w))
    for tw in tblPr.findall(qn('w:tblW')):
        tblPr.remove(tw)
    tw = OxmlElement('w:tblW'); tw.set(qn('w:type'), 'dxa')
    tw.set(qn('w:w'), str(sum(widths)))
    insert_tblpr(tblPr, tw)
    for tl in tblPr.findall(qn('w:tblLayout')):
        tblPr.remove(tl)
    tl = OxmlElement('w:tblLayout'); tl.set(qn('w:type'), 'fixed')
    insert_tblpr(tblPr, tl)

    # header row: repeat, grey shading, centred
    if tbl.rows:
        tr0 = tbl.rows[0]._tr
        trPr = tr0.find(qn('w:trPr'))
        if trPr is None:
            trPr = OxmlElement('w:trPr'); tr0.insert(0, trPr)
        if trPr.find(qn('w:tblHeader')) is None:
            trPr.append(OxmlElement('w:tblHeader'))
        for c in tbl.rows[0].cells:
            tcPr = c._tc.find(qn('w:tcPr'))
            if tcPr is None:
                tcPr = OxmlElement('w:tcPr'); c._tc.insert(0, tcPr)
            for sh in tcPr.findall(qn('w:shd')):
                tcPr.remove(sh)
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), 'D9D9D9')
            insert_tcpr(tcPr, shd)
            for v in tcPr.findall(qn('w:vAlign')):
                tcPr.remove(v)
            va = OxmlElement('w:vAlign'); va.set(qn('w:val'), 'center')
            insert_tcpr(tcPr, va)
            for p in c.paragraphs:
                pPr = p._p.get_or_add_pPr()
                for j in pPr.findall(qn('w:jc')):
                    pPr.remove(j)
                jc = OxmlElement('w:jc'); jc.set(qn('w:val'), 'center')
                insert_ppr(pPr, jc)


def fix_tables(doc):
    for t in doc.tables:
        fix_table(t)


def build_numfmt(nx_bytes):
    """Map pandoc numId -> {ilvl: numFmt} from a numbering.xml."""
    root = etree.fromstring(nx_bytes)
    abstracts = {}
    for ab in root.findall(wq('abstractNum')):
        aid = ab.get(wq('abstractNumId'))
        fmts = {}
        for lvl in ab.findall(wq('lvl')):
            nf = lvl.find(wq('numFmt'))
            fmts[lvl.get(wq('ilvl'))] = nf.get(wq('val')) if nf is not None else None
        abstracts[aid] = fmts
    nummap = {}
    for num in root.findall(wq('num')):
        ref = num.find(wq('abstractNumId'))
        nummap[num.get(wq('numId'))] = ref.get(wq('val')) if ref is not None else None
    return {nid: abstracts.get(aid, {}) for nid, aid in nummap.items()}


def fix_lists(doc, numfmt):
    """Re-map pandoc list paragraphs (numId >= 1000) to the template's numbering:
    unordered -> numId 4 (bullet ●), ordered -> numId 1 (decimal %1.).
    Indentation comes from the template's numbering definitions; force single spacing."""
    for p in doc.paragraphs:
        pPr = p._p.find(qn('w:pPr'))
        numPr = pPr.find(qn('w:numPr')) if pPr is not None else None
        if numPr is None:
            continue
        nid_el = numPr.find(qn('w:numId'))
        if nid_el is None:
            continue
        nid = nid_el.get(qn('w:val'))
        if nid is None or not nid.isdigit() or int(nid) < 1000:
            continue                                   # template's own lists untouched
        ilvl_el = numPr.find(qn('w:ilvl'))
        ilvl = ilvl_el.get(qn('w:val')) if ilvl_el is not None else '0'
        fmt = numfmt.get(nid, {}).get(ilvl) or numfmt.get(nid, {}).get('0')
        nid_el.set(qn('w:val'), '4' if fmt == 'bullet' else '1')   # template numIds
        if ilvl_el is None:
            ilvl_el = OxmlElement('w:ilvl')
            numPr.insert(0, ilvl_el)
        ilvl_el.set(qn('w:val'), ilvl)
        sp = pPr.find(qn('w:spacing'))
        if sp is None:
            sp = OxmlElement('w:spacing')
            insert_ppr(pPr, sp)
        sp.set(qn('w:before'), '0'); sp.set(qn('w:after'), '0')
        sp.set(qn('w:line'), '240'); sp.set(qn('w:lineRule'), 'auto')


def fix_body_indent(doc, left=0, first_line=0, line=240, before=0, after=120):
    """Match the template's tidy body-paragraph format for injected content:
    flush-left, no first-line protrusion, single line spacing, small space after
    each paragraph so long sections stay readable. Lists (numPr) and headings
    are left untouched."""
    paras = doc.paragraphs
    start = None
    for i, p in enumerate(paras):
        if p.style.name == 'Heading 1':
            start = i
            break
    if start is None:
        return
    for p in paras[start:]:
        if raw_pstyle(p) in ('BlockText',):
            continue                                  # keep block-quote look
        pPr = p._p.find(qn('w:pPr'))
        num = pPr.find(qn('w:numPr')) if pPr is not None else None
        if p.style.name == 'normal' and num is None and p.text.strip():
            pPr = p._p.get_or_add_pPr()
            # indentation: flush left, no first-line/hanging
            ind = pPr.find(qn('w:ind'))
            if ind is None:
                ind = OxmlElement('w:ind')
                insert_ppr(pPr, ind)
            ind.set(qn('w:left'), str(left))
            ind.set(qn('w:firstLine'), str(first_line))
            if ind.get(qn('w:hanging')) is not None:
                del ind.attrib[qn('w:hanging')]
            # spacing: single line, space after for readability
            sp = pPr.find(qn('w:spacing'))
            if sp is None:
                sp = OxmlElement('w:spacing')
                insert_ppr(pPr, sp)
            sp.set(qn('w:before'), str(before))
            sp.set(qn('w:after'), str(after))
            sp.set(qn('w:line'), str(line))
            sp.set(qn('w:lineRule'), 'auto')
            # alignment: left
            jc = pPr.find(qn('w:jc'))
            if jc is None:
                jc = OxmlElement('w:jc')
                insert_ppr(pPr, jc)
            jc.set(qn('w:val'), 'left')


def fix_section_labels(doc, before=180, after=60):
    """Turn standalone bold paragraphs (e.g. '3. Outbound Adapter Class ...') into
    clear sub-headings: extra space before/after so long sections stay structured."""
    paras = doc.paragraphs
    start = None
    for i, p in enumerate(paras):
        if p.style.name == 'Heading 1':
            start = i
            break
    if start is None:
        return
    for p in paras[start:]:
        if raw_pstyle(p) in ('BlockText',):
            continue                                  # keep block-quote look
        pPr = p._p.find(qn('w:pPr'))
        num = pPr.find(qn('w:numPr')) if pPr is not None else None
        if p.style.name != 'normal' or num is not None:
            continue
        runs = [r for r in p.runs if r.text.strip()]
        if not runs or not all(r.bold for r in runs):
            continue
        txt = p.text.strip()
        if not txt or txt.endswith('.'):
            continue                       # bold sentences, not labels
        pPr = p._p.get_or_add_pPr()
        sp = pPr.find(qn('w:spacing'))
        if sp is None:
            sp = OxmlElement('w:spacing')
            insert_ppr(pPr, sp)
        sp.set(qn('w:before'), str(before))
        sp.set(qn('w:after'), str(after))


def add_heading_numbering(doc):
    """Apply multilevel 1./1.1/1.1.1 numbering to Heading 1/2/3 styles and
    override the template's oversized style indents (720/1440/2160 twips)."""
    for lvl, name in ((0, 'Heading 1'), (1, 'Heading 2'), (2, 'Heading 3')):
        style = doc.styles[name]
        pPr = style.element.get_or_add_pPr()
        for np in pPr.findall(qn('w:numPr')):
            pPr.remove(np)
        numPr = OxmlElement('w:numPr')
        ilvl = OxmlElement('w:ilvl'); ilvl.set(qn('w:val'), str(lvl))
        numId = OxmlElement('w:numId'); numId.set(qn('w:val'), str(HEADING_NUM_ID))
        numPr.append(ilvl); numPr.append(numId)
        insert_ppr(pPr, numPr)
        # keep headings near the left margin (number at 0, text at 210 twips)
        ind = pPr.find(qn('w:ind'))
        if ind is None:
            ind = OxmlElement('w:ind')
            insert_ppr(pPr, ind)
        ind.set(qn('w:left'), '210')
        ind.set(qn('w:hanging'), '210')


def inject_styles(path):
    """Add the pandoc styles the template lacks (BlockText, VerbatimChar) into
    word/styles.xml so block-quotes and inline code render correctly."""
    zin = zipfile.ZipFile(path)
    sx = zin.read('word/styles.xml').decode('utf-8')
    zin.close()
    if '<w:styleId="BlockText"' in sx or '<w:styleId="VerbatimChar"' in sx:
        return
    sx = sx.replace('</w:styles>', EXTRA_STYLES + '</w:styles>')
    tmp = path + '.tmp'
    zin = zipfile.ZipFile(path)
    zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'word/styles.xml':
            data = sx.encode('utf-8')
        zout.writestr(item, data)
    zin.close(); zout.close()
    shutil.move(tmp, path)


def inject_numbering(path):
    """Add the heading multilevel definition into the docx's numbering.xml."""
    abstract = (
        '<w:abstractNum w:abstractNumId="%d">'
        '<w:multiLevelType w:val="hybridMultilevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        '<w:lvlText w:val="%%1."/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="210" w:hanging="210"/></w:pPr></w:lvl>'
        '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        '<w:lvlText w:val="%%1.%%2."/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="210" w:hanging="210"/></w:pPr></w:lvl>'
        '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        '<w:lvlText w:val="%%1.%%2.%%3."/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="210" w:hanging="210"/></w:pPr></w:lvl>'
        '</w:abstractNum>'
    ) % HEADING_NUM_ID
    num = ('<w:num w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>'
           % (HEADING_NUM_ID, HEADING_NUM_ID))
    tmp = path + '.tmp'
    zin = zipfile.ZipFile(path)
    zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'word/numbering.xml':
            s = data.decode('utf-8')
            if '<w:num ' in s:
                s = s.replace('<w:num ', abstract + '<w:num ', 1)
            else:
                s = s.replace('</w:numbering>', abstract + '</w:numbering>')
            s = s.replace('</w:numbering>', num + '</w:numbering>')
            data = s.encode('utf-8')
        zout.writestr(item, data)
    zin.close(); zout.close()
    shutil.move(tmp, path)


def merge():
    shutil.copyfile(TEMPLATE, OUT)
    inject_styles(OUT)
    d = docx.Document(OUT)
    set_cell(d.tables[0].rows[0].cells[1], DOC_NAME)
    set_cell(d.tables[0].rows[3].cells[1], DOC_ASSOC)
    set_cell(d.tables[1].rows[2].cells[1], HISTORY)

    body = d.element.body
    sectPr = body.find(qn('w:sectPr'))
    start = None
    for i, ch in enumerate(list(body)):
        if ch.tag == qn('w:p'):
            ppr = ch.find(qn('w:pPr'))
            st = ppr.find(qn('w:pStyle')) if ppr is not None else None
            if st is not None and st.get(qn('w:val')) == 'Heading1':
                start = i
                break
    assert start is not None, 'first Heading1 not found in template'
    for ch in list(body)[start:]:
        if ch is not sectPr:
            body.remove(ch)

    z = zipfile.ZipFile(CONTENT)
    cx = etree.fromstring(z.read('word/document.xml'))
    numfmt = build_numfmt(z.read('word/numbering.xml'))
    z.close()
    for ch in list(cx.find(wq('body'))):
        if ch.tag != wq('sectPr'):
            sectPr.addprevious(copy.deepcopy(ch))

    fix_tables(d)
    fix_lists(d, numfmt)
    fix_body_indent(d)
    fix_section_labels(d)
    add_heading_numbering(d)
    d.save(OUT)
    inject_numbering(OUT)
    print('saved:', OUT)


def main():
    prepare_body()
    convert()
    merge()
    for f in (BODY_MD, CONTENT):
        if os.path.exists(f):
            os.remove(f)
    print('cleaned temp files')


if __name__ == '__main__':
    main()
