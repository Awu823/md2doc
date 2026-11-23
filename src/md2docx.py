# -*- coding: utf-8 -*-
"""
Generic Markdown -> corporate Word (.docx) converter.

Reusable, configurable CLI that converts a Markdown document into a Word file
styled after a corporate .docx template. Unlike a plain pandoc run, this tool:

  * Keeps the template's cover (Document Information + History), TOC page,
    header / footer and styles; only the sample body is replaced.
  * Auto-detects the body start (first ``# `` heading) and the cover metadata
    (Document Name / Associated Documents / History) from the Markdown, or
    accepts them via CLI / config.
  * Restyles injected tables (full grid, grey centred header, tight margins,
    content-aware column widths), auto-numbers Heading 1/2/3 (1. / 1.1 / 1.1.1),
    re-maps pandoc lists onto the template's bullet numbering, and injects the
    styles the template lacks (BlockText, VerbatimChar).

Usage
-----
    python md2docx.py [options] input.md

Options
-------
    -t, --template FILE    corporate .docx template   (default: template.docx)
    -o, --output FILE      output .docx               (default: input.docx)
    --body-start N         body start line index in the MD (default: auto)
    --doc-name TEXT        cover Document Name        (default: auto from MD)
    --doc-assoc TEXT       cover Associated Documents (default: auto from MD)
    --history TEXT         cover History value        (default: auto from MD)
    --cover-name T,R,C     cell of Document Name      (default: 0,0,1)
    --cover-assoc T,R,C    cell of Associated Docs    (default: 0,3,1)
    --cover-hist T,R,C     cell of History            (default: 1,2,1)
    --config FILE          JSON file with any of the above keys
    --keep-temp            keep intermediate _body.md / _content.docx

Requires: python-docx, lxml, pandoc on PATH.

Example
-------
    python md2docx.py --config md2docx.json
    python md2docx.py my_doc.md -t template.docx -o my_doc.docx
"""
import argparse
import copy
import json
import os
import re
import shutil
import subprocess
import zipfile

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
USABLE_WIDTH = 9639
HEADING_NUM_ID = 5

# Project root = parent of this file's directory (src/).
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# Intermediate files live in the project root, independent of the cwd.
BODY_MD = os.path.join(ROOT, '_body.md')
CONTENT = os.path.join(ROOT, '_content.docx')


def project_path(*parts):
    return os.path.join(ROOT, *parts)


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


# --------------------------------------------------------------------------- #
#  Config / auto-detection
# --------------------------------------------------------------------------- #
def default_config():
    """Built-in defaults so a bare `python md2docx.py file.md` just works."""
    return {
        'template': project_path('templates', 'template.docx'),
        'body_start': None,          # None -> auto-detect first '# ' heading
        'doc_name': None,            # None -> auto-detect from MD cover table
        'doc_assoc': None,
        'history': None,
        'cover_name': (0, 0, 1),     # (table_idx, row, col)
        'cover_assoc': (0, 3, 1),
        'cover_hist': (1, 2, 1),
    }


def _cell(triple):
    if triple is None:
        return None
    if isinstance(triple, (list, tuple)):
        return tuple(int(x) for x in triple)
    return tuple(int(x) for x in str(triple).split(','))


def parse_cell(text):
    """'0,0,1' -> (0, 0, 1)."""
    return tuple(int(x) for x in str(text).split(','))


def load_config(args):
    cfg = default_config()
    if args.config:
        with open(args.config, encoding='utf-8') as fh:
            data = json.load(fh)
        for k in ('body_start', 'doc_name', 'doc_assoc', 'history', 'template'):
            if k in data:
                cfg[k] = data[k]
        for k in ('cover_name', 'cover_assoc', 'cover_hist'):
            if k in data:
                cfg[k] = _cell(data[k])
    if args.template:
        cfg['template'] = args.template
    if args.body_start is not None:
        cfg['body_start'] = args.body_start
    if args.doc_name is not None:
        cfg['doc_name'] = args.doc_name
    if args.doc_assoc is not None:
        cfg['doc_assoc'] = args.doc_assoc
    if args.history is not None:
        cfg['history'] = args.history
    if args.cover_name:
        cfg['cover_name'] = parse_cell(args.cover_name)
    if args.cover_assoc:
        cfg['cover_assoc'] = parse_cell(args.cover_assoc)
    if args.cover_hist:
        cfg['cover_hist'] = parse_cell(args.cover_hist)
    cfg['md'] = args.md or data.get('md')
    if cfg['md'] is None:
        raise SystemExit('no input .md file given (positional or --config md)')
    if args.output or data.get('output'):
        cfg['output'] = args.output or data.get('output')
    elif cfg['md'].startswith('input'):
        cfg['output'] = project_path('output',
                                     os.path.splitext(os.path.basename(cfg['md']))[0] + '.docx')
    else:
        cfg['output'] = os.path.splitext(cfg['md'])[0] + '.docx'
    return cfg


def _unescape(s):
    return s.replace('\\_', '_').replace('\\-', '-').replace('\\.', '.') \
            .replace('\\ ', ' ').replace('\\*', '*').replace('\\`', '`')


def detect_body_start(lines):
    """First line starting with '# ' (H1). Returns index or None."""
    for i, ln in enumerate(lines):
        if ln.lstrip().startswith('# ') and not ln.lstrip().startswith('## '):
            return i
    return None


def detect_cover_meta(lines):
    """Extract Document Name / Associated Documents / History from the MD."""
    doc_name = doc_assoc = history = None
    for ln in lines:
        m = re.search(r'\|\s*\*\*Document Name\*\*\s*\|\s*(.*?)\s*\|', ln)
        if m:
            doc_name = _unescape(m.group(1))
        m = re.search(r'\|\s*\*\*Associated Documents\*\*\s*\|\s*(.*?)\s*\|', ln)
        if m:
            doc_assoc = _unescape(m.group(1))
        m = re.search(r'\|\s*\*\*Changes\*\*\s*\|\s*(.*?)\s*\|', ln)
        # History row: | version | changes |  (not the header)
        if not m:
            m = re.search(r'^\|\s*([0-9][^|]*)\s*\|\s*([^|]+)\s*\|', ln)
        if m and not m.group(1).strip().lower().startswith('version'):
            history = _unescape(m.group(2))
    return doc_name, doc_assoc, history


def wq(t):
    return '{%s}%s' % (W, t)


def raw_pstyle(p):
    """Return the paragraph's pStyle styleId (e.g. 'BlockText') or None."""
    ppr = p._p.find(qn('w:pPr'))
    st = ppr.find(qn('w:pStyle')) if ppr is not None else None
    return st.get(qn('w:val')) if st is not None else None


# --------------------------------------------------------------------------- #
#  OOXML helpers
# --------------------------------------------------------------------------- #
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


# --------------------------------------------------------------------------- #
#  Steps
# --------------------------------------------------------------------------- #
def prepare_body(cfg):
    lines = open(cfg['md'], encoding='utf-8').read().splitlines(keepends=True)
    body_start = cfg['body_start']
    if body_start is None:
        body_start = detect_body_start(lines)
        if body_start is None:
            raise SystemExit('could not auto-detect body start; use --body-start')
    body = lines[body_start:]
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
    print('body start line:', body_start, '| body md lines:', len(out))
    return lines, body_start


def _is_ordered(s):
    return bool(re.match(r'^\d+[.)]\s', s))


def convert(cfg):
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
                    '--reference-doc=' + cfg['template'], '-o', CONTENT],
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
            ind = pPr.find(qn('w:ind'))
            if ind is None:
                ind = OxmlElement('w:ind')
                insert_ppr(pPr, ind)
            ind.set(qn('w:left'), str(left))
            ind.set(qn('w:firstLine'), str(first_line))
            if ind.get(qn('w:hanging')) is not None:
                del ind.attrib[qn('w:hanging')]
            sp = pPr.find(qn('w:spacing'))
            if sp is None:
                sp = OxmlElement('w:spacing')
                insert_ppr(pPr, sp)
            sp.set(qn('w:before'), str(before))
            sp.set(qn('w:after'), str(after))
            sp.set(qn('w:line'), str(line))
            sp.set(qn('w:lineRule'), 'auto')
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


def _set_cover(doc, triple, text):
    if not triple:
        return
    ti, ri, ci = triple
    try:
        set_cell(doc.tables[ti].rows[ri].cells[ci], text)
    except IndexError:
        print('warning: cover cell %s out of range' % (triple,))


def merge(cfg):
    out = cfg['output']
    shutil.copyfile(cfg['template'], out)
    inject_styles(out)
    d = docx.Document(out)
    _set_cover(d, cfg['cover_name'], cfg['doc_name'])
    _set_cover(d, cfg['cover_assoc'], cfg['doc_assoc'])
    _set_cover(d, cfg['cover_hist'], cfg['history'])

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
    d.save(out)
    inject_numbering(out)
    print('saved:', out)


def _resolve_input(path):
    """Resolve an input .md path; fall back to project input/ dir."""
    if os.path.exists(path):
        return path
    cand = project_path('input', path)
    if os.path.exists(cand):
        return cand
    if not os.path.splitext(path)[1]:
        cand2 = project_path('input', path + '.md')
        if os.path.exists(cand2):
            return cand2
    raise SystemExit('input file not found: %s (also tried project input/)' % path)


def _resolve_output(path):
    """Resolve an output path; put it under project output/ if not specified."""
    if not os.path.dirname(path):
        return project_path('output', os.path.basename(path))
    return path


def _resolve_template(path):
    """Resolve the template path; fall back to project templates/ dir."""
    if os.path.exists(path):
        return path
    cand = project_path('templates', path)
    if os.path.exists(cand):
        return cand
    raise SystemExit('template not found: %s (also tried project templates/)' % path)


def main(argv=None):
    ap = argparse.ArgumentParser(
        description='Convert a Markdown doc into a corporate Word template.')
    ap.add_argument('md', nargs='?', help='input .md file')
    ap.add_argument('-t', '--template', help='corporate .docx template')
    ap.add_argument('-o', '--output', help='output .docx')
    ap.add_argument('--body-start', type=int, help='body start line index')
    ap.add_argument('--doc-name', help='cover Document Name')
    ap.add_argument('--doc-assoc', help='cover Associated Documents')
    ap.add_argument('--history', help='cover History value')
    ap.add_argument('--cover-name', help='Document Name cell as T,R,C')
    ap.add_argument('--cover-assoc', help='Associated Docs cell as T,R,C')
    ap.add_argument('--cover-hist', help='History cell as T,R,C')
    ap.add_argument('--config', help='JSON config file')
    ap.add_argument('--keep-temp', action='store_true',
                    help='keep intermediate _body.md / _content.docx')
    args = ap.parse_args(argv)

    if args.config:
        cfg = load_config(args)
    else:
        if not args.md:
            ap.error('an input .md file is required (or use --config)')
        cfg = load_config(args)

    # Resolve paths relative to the project root when not found as given.
    cfg['md'] = _resolve_input(cfg['md'])
    cfg['output'] = _resolve_output(cfg['output'])
    cfg['template'] = _resolve_template(cfg['template'])

    # Auto-detect cover metadata that was not supplied / configured.
    lines = open(cfg['md'], encoding='utf-8').read().splitlines()
    doc_name, doc_assoc, history = detect_cover_meta(lines)
    cfg['doc_name'] = cfg['doc_name'] or doc_name
    cfg['doc_assoc'] = cfg['doc_assoc'] or doc_assoc
    cfg['history'] = cfg['history'] or history

    prepare_body(cfg)
    convert(cfg)
    merge(cfg)
    if not args.keep_temp:
        for f in (BODY_MD, CONTENT):
            if os.path.exists(f):
                os.remove(f)
    print('cleaned temp files')


if __name__ == '__main__':
    main()
