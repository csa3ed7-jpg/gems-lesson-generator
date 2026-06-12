import copy
import os
import shutil
import zipfile
import re
from lxml import etree

NSMAP = {
    'a':  'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p':  'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r':  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

def _t(text):
    """Build a simple <a:r><a:rPr/><a:t>text</a:t></a:r>"""
    r = etree.Element('{http://schemas.openxmlformats.org/drawingml/2006/main}r')
    rpr = etree.SubElement(r, '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
    rpr.set('lang', 'en-US')
    rpr.set('dirty', '0')
    t = etree.SubElement(r, '{http://schemas.openxmlformats.org/drawingml/2006/main}t')
    t.text = str(text)
    return r

def _tb(text):
    """Bold run"""
    r = _t(text)
    r[0].set('b', '1')
    return r

def _para(*runs):
    p = etree.Element('{http://schemas.openxmlformats.org/drawingml/2006/main}p')
    for run in runs:
        p.append(run)
    return p

def _bullet_para(text, bold_prefix=None):
    p = etree.Element('{http://schemas.openxmlformats.org/drawingml/2006/main}p')
    pPr = etree.SubElement(p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
    pPr.set('marL', '342900')
    pPr.set('indent', '-342900')
    buChar = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}buChar')
    buChar.set('char', '•')
    if bold_prefix:
        p.append(_tb(bold_prefix + ': '))
        p.append(_t(text))
    else:
        p.append(_t(text))
    return p

def set_txbody(sp, paragraphs):
    """Replace txBody content with given list of <a:p> elements"""
    NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    txBody = sp.find('.//{%s}txBody' % NS)
    if txBody is None:
        return
    # Remove existing <a:p> elements
    for p in txBody.findall('{%s}p' % NS):
        txBody.remove(p)
    for para in paragraphs:
        txBody.append(copy.deepcopy(para))

def get_placeholders(slide_xml):
    """Return dict of ph_type/idx -> sp element"""
    NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    tree = etree.fromstring(slide_xml.encode() if isinstance(slide_xml, str) else slide_xml)
    result = {}
    for sp in tree.findall('.//{%s}sp' % NS_P):
        nvPr = sp.find('.//{%s}nvPr' % NS_P)
        if nvPr is not None:
            ph = nvPr.find('{%s}ph' % NS_P)
            if ph is not None:
                ph_type = ph.get('type', 'body')
                idx = ph.get('idx', '0')
                key = ph_type if ph_type != 'body' else f'body_{idx}'
                result[key] = sp
    return result, tree

def tree_to_str(tree):
    return etree.tostring(tree, xml_declaration=True, encoding='utf-8', pretty_print=True).decode()


def build_pptx(lesson, grade, week, lesson_type, template_path, output_path):
    """Fill template PPTX with lesson data and save to output_path"""
    import tempfile
    
    tmp_dir = tempfile.mkdtemp()
    try:
        # Unzip template
        with zipfile.ZipFile(template_path, 'r') as z:
            z.extractall(tmp_dir)
        
        slides_dir = os.path.join(tmp_dir, 'ppt', 'slides')
        
        def read_slide(n):
            path = os.path.join(slides_dir, f'slide{n}.xml')
            with open(path, 'rb') as f:
                return f.read()
        
        def write_slide(n, tree):
            path = os.path.join(slides_dir, f'slide{n}.xml')
            xml_str = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', pretty_print=True)
            with open(path, 'wb') as f:
                f.write(xml_str)
        
        def parse(n):
            return etree.fromstring(read_slide(n))
        
        NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
        
        def get_sp(tree, name_contains=None, ph_type=None, ph_idx=None):
            for sp in tree.findall('.//{%s}sp' % NS_P):
                cNvPr = sp.find('.//{%s}cNvPr' % NS_P)
                if name_contains and cNvPr is not None:
                    if name_contains.lower() in cNvPr.get('name','').lower():
                        return sp
                if ph_type is not None or ph_idx is not None:
                    ph = sp.find('.//{%s}ph' % NS_P)
                    if ph is not None:
                        if ph_type and ph.get('type') == ph_type:
                            return sp
                        if ph_idx is not None and ph.get('idx') == str(ph_idx):
                            return sp
            return None
        
        def set_sp_text(sp, paragraphs):
            if sp is None: return
            txBody = sp.find('.//{%s}txBody' % NS_A)
            if txBody is None: return
            for p in txBody.findall('{%s}p' % NS_A):
                txBody.remove(p)
            for para in paragraphs:
                txBody.append(copy.deepcopy(para))

        def p(*runs):
            el = etree.Element('{%s}p' % NS_A)
            for r in runs: el.append(r)
            return el
        
        def r(text, bold=False, size=None):
            el = etree.Element('{%s}r' % NS_A)
            rpr = etree.SubElement(el, '{%s}rPr' % NS_A)
            rpr.set('lang', 'en-US'); rpr.set('dirty', '0')
            if bold: rpr.set('b', '1')
            if size: rpr.set('sz', str(size))
            t = etree.SubElement(el, '{%s}t' % NS_A)
            t.text = str(text)
            return el
        
        def bullet_p(text, bold_label=None):
            el = etree.Element('{%s}p' % NS_A)
            pPr = etree.SubElement(el, '{%s}pPr' % NS_A)
            pPr.set('marL', '342900'); pPr.set('indent', '-342900')
            bu = etree.SubElement(pPr, '{%s}buChar' % NS_A)
            bu.set('char', '•')
            if bold_label:
                el.append(r(bold_label + ': ', bold=True))
                el.append(r(text))
            else:
                el.append(r(text))
            return el
        
        def empty_p():
            return etree.Element('{%s}p' % NS_A)

        # ── SLIDE 1: Cover ──
        t1 = parse(1)
        sp_title = get_sp(t1, ph_type='title')
        sp_body  = get_sp(t1, ph_idx=1)
        set_sp_text(sp_title, [p(r('Learn Like a GEM', bold=True))])
        from datetime import date
        today = date.today().strftime('%d %B %Y')
        set_sp_text(sp_body, [
            p(r('Subject: ', bold=True), r('Physical Education')),
            p(r('Lesson: ', bold=True), r(lesson.get('title',''))),
            p(r('Grade: ', bold=True), r(f'{grade} | Week {week} | {lesson_type}')),
            p(r('Date: ', bold=True), r(today)),
        ])
        write_slide(1, t1)

        # ── SLIDE 2: Do Now ──
        t2 = parse(2)
        sp_body2 = get_sp(t2, ph_idx=1)
        dn = lesson.get('doNow', {})
        qs = dn.get('questions', [])
        paras = [p(r('Retrieval Questions', bold=True))]
        for i, q in enumerate(qs):
            paras.append(bullet_p(q, f'Q{i+1}'))
        paras.append(empty_p())
        paras.append(p(r('Teacher will: ', bold=True), r(dn.get('teacherWill',''))))
        paras.append(p(r('Students will: ', bold=True), r(dn.get('studentsWill',''))))
        set_sp_text(sp_body2, paras)
        write_slide(2, t2)

        # ── SLIDE 3: Learning Outcome + To Know ──
        t3 = parse(3)
        sp_lo   = get_sp(t3, ph_idx=1)
        sp_know = get_sp(t3, ph_idx=10)
        sp_sub  = get_sp(t3, ph_idx=11)
        
        lo_text = lesson.get('lo', '')
        set_sp_text(sp_lo, [
            p(r('Learning Outcome', bold=True)),
            p(r(lo_text)),
            empty_p(),
            p(r(f'Standard: {lesson.get("standard","")}'))
        ])
        
        tk = lesson.get('toKnow', {})
        tier2 = ', '.join(tk.get('tier2', []))
        tier3 = ', '.join(tk.get('tier3', []))
        concepts = tk.get('concepts', [])
        know_paras = [p(r('Tier 2: ', bold=True), r(tier2)),
                      p(r('Tier 3: ', bold=True), r(tier3)),
                      empty_p(),
                      p(r('Key Concepts', bold=True))]
        for c in concepts:
            know_paras.append(bullet_p(c))
        set_sp_text(sp_know, know_paras)
        if sp_sub:
            set_sp_text(sp_sub, [p(r(lesson.get('unit',''), bold=True))])
        write_slide(3, t3)

        # ── SLIDE 4: I Do, We Do, You Do overview ──
        t4 = parse(4)
        sp_body4 = get_sp(t4, ph_idx=1)
        ido = lesson.get('iDo', {}); wdo = lesson.get('weDo', {}); ydo = lesson.get('youDo', {})
        set_sp_text(sp_body4, [
            p(r('I Do', bold=True)),
            bullet_p(ido.get('teacherWill', '')),
            empty_p(),
            p(r('We Do', bold=True)),
            bullet_p(wdo.get('activity', '')),
            empty_p(),
            p(r('You Do', bold=True)),
            bullet_p(ydo.get('task', '')),
        ])
        write_slide(4, t4)

        # ── SLIDE 5: I Do ──
        t5 = parse(5)
        sp_body5 = get_sp(t5, ph_idx=1)
        steps = ido.get('steps', [])
        paras5 = []
        for s in steps:
            paras5.append(p(r(s.get('label',''), bold=True), r(': ' + s.get('text',''))))
        paras5.append(empty_p())
        paras5.append(p(r('Think Aloud: ', bold=True), r(f'"{ido.get("thinkAloud","")}"')))
        paras5.append(empty_p())
        paras5.append(p(r('Teacher will: ', bold=True), r(ido.get('teacherWill',''))))
        paras5.append(p(r('Students will: ', bold=True), r(ido.get('studentsWill',''))))
        set_sp_text(sp_body5, paras5)
        write_slide(5, t5)

        # ── SLIDE 6: We Do ──
        t6 = parse(6)
        sp_body6 = get_sp(t6, ph_idx=1)
        bullets = wdo.get('bullets', [])
        cfuQs   = wdo.get('cfuQ', [])
        paras6  = [p(r(wdo.get('activity',''), bold=True))]
        for b in bullets:
            paras6.append(bullet_p(b))
        paras6.append(empty_p())
        paras6.append(p(r('CFU Questions', bold=True)))
        for q in cfuQs:
            paras6.append(bullet_p(q))
        paras6.append(empty_p())
        paras6.append(p(r('Teacher will: ', bold=True), r(wdo.get('teacherWill',''))))
        paras6.append(p(r('Students will: ', bold=True), r(wdo.get('studentsWill',''))))
        set_sp_text(sp_body6, paras6)
        write_slide(6, t6)

        # ── SLIDE 7: You Do ──
        t7 = parse(7)
        sp_body7 = get_sp(t7, ph_idx=1)
        ybullets = ydo.get('bullets', [])
        dqs      = ydo.get('decisionQ', [])
        paras7   = [p(r(ydo.get('task',''), bold=True))]
        for b in ybullets:
            paras7.append(bullet_p(b))
        paras7.append(empty_p())
        paras7.append(p(r('Decision-Making Questions', bold=True)))
        for q in dqs:
            paras7.append(bullet_p(q))
        paras7.append(empty_p())
        paras7.append(p(r('Teacher will: ', bold=True), r(ydo.get('teacherWill',''))))
        paras7.append(p(r('Students will: ', bold=True), r(ydo.get('studentsWill',''))))
        set_sp_text(sp_body7, paras7)
        write_slide(7, t7)

        # ── SLIDE 8: Affirmative Checking ──
        t8 = parse(8)
        sp_body8 = get_sp(t8, ph_idx=1)
        af = lesson.get('affirmative', {})
        aqs = af.get('questions', [])
        paras8 = []
        for aq in aqs:
            paras8.append(p(r('Q: ', bold=True), r(aq.get('q',''))))
            paras8.append(p(r('Look for: ', bold=True), r(aq.get('lookFor',''))))
            paras8.append(empty_p())
        paras8.append(p(r('Most understood: ', bold=True), r(af.get('mostUnderstood',''))))
        paras8.append(p(r('Some struggled: ', bold=True), r(af.get('someStruggled',''))))
        paras8.append(p(r('Most confused: ', bold=True), r(af.get('mostConfused',''))))
        set_sp_text(sp_body8, paras8)
        write_slide(8, t8)

        # ── SLIDE 9: Exit Ticket ──
        t9 = parse(9)
        sp_body9 = get_sp(t9, ph_idx=1)
        et = lesson.get('exitTicket', {})
        wb = ' | '.join(et.get('wordBank', []))
        paras9 = [
            p(r('Q1 — Identify', bold=True)),
            p(r(et.get('q1',''))),
            p(r('Word Bank: ', bold=True), r(wb)),
            empty_p(),
            p(r('Q2 — Explain', bold=True)),
            p(r(et.get('q2',''))),
        ]
        set_sp_text(sp_body9, paras9)
        write_slide(9, t9)

        # ── SLIDE 10: Thank You — keep as is ──

        # Repack to PPTX
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root, dirs, files in os.walk(tmp_dir):
                for file in files:
                    full = os.path.join(root, file)
                    arc  = os.path.relpath(full, tmp_dir)
                    zout.write(full, arc)
    finally:
        shutil.rmtree(tmp_dir)

