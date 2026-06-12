import copy, os, shutil, zipfile, io
from lxml import etree

# Both namespaces needed
NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

def _r(text, bold=False):
    """Make a text run"""
    r = etree.Element('{%s}r' % NS_A)
    rPr = etree.SubElement(r, '{%s}rPr' % NS_A)
    rPr.set('lang', 'en-US'); rPr.set('dirty', '0')
    if bold: rPr.set('b', '1')
    t = etree.SubElement(r, '{%s}t' % NS_A)
    t.text = str(text)
    return r

def _p(*runs):
    """Make a paragraph with runs"""
    p = etree.Element('{%s}p' % NS_A)
    for run in runs: p.append(copy.deepcopy(run))
    return p

def _bullet(text, label=None):
    """Make a bulleted paragraph"""
    p = etree.Element('{%s}p' % NS_A)
    pPr = etree.SubElement(p, '{%s}pPr' % NS_A)
    pPr.set('marL', '342900'); pPr.set('indent', '-342900')
    bu = etree.SubElement(pPr, '{%s}buChar' % NS_A)
    bu.set('char', '•')
    if label:
        p.append(_r(label + ': ', bold=True))
    p.append(_r(text))
    return p

def _empty():
    return etree.Element('{%s}p' % NS_A)

def set_text(sp, paragraphs):
    """Replace txBody paragraphs — handles both NS_P and NS_A txBody"""
    # txBody is in NS_P in this template
    txBody = sp.find('{%s}txBody' % NS_P)
    if txBody is None:
        txBody = sp.find('.//{%s}txBody' % NS_A)
    if txBody is None:
        return
    # Remove existing <a:p> elements
    for p in txBody.findall('{%s}p' % NS_A):
        txBody.remove(p)
    for para in paragraphs:
        txBody.append(copy.deepcopy(para))

def get_sp(tree, ph_type=None, ph_idx=None):
    """Find shape by placeholder type or idx"""
    for sp in tree.findall('.//{%s}sp' % NS_P):
        ph = sp.find('.//{%s}ph' % NS_P)
        if ph is None: continue
        if ph_type and ph.get('type') == ph_type:
            return sp
        if ph_idx is not None and ph.get('idx') == str(ph_idx):
            return sp
    return None

def build_pptx(lesson, grade, week, lesson_type, template_path, output_path):
    from datetime import date

    buf = io.BytesIO()
    with zipfile.ZipFile(template_path, 'r') as zin:
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                
                if item.filename.startswith('ppt/slides/slide') and item.filename.endswith('.xml'):
                    # Extract slide number
                    fname = item.filename.split('/')[-1]  # slide1.xml
                    num = int(fname.replace('slide','').replace('.xml',''))
                    
                    tree = etree.fromstring(data)
                    today = date.today().strftime('%d %B %Y')
                    
                    if num == 1:
                        # Cover slide
                        sp_t = get_sp(tree, ph_type='title')
                        sp_b = get_sp(tree, ph_idx=1)
                        if sp_t: set_text(sp_t, [_p(_r('Learn Like a GEM', bold=True))])
                        if sp_b: set_text(sp_b, [
                            _p(_r('Subject: ', bold=True), _r('Physical Education')),
                            _p(_r('Lesson: ', bold=True), _r(lesson.get('title',''))),
                            _p(_r('Grade: ', bold=True), _r(f"{grade} | Week {week} | {lesson_type}")),
                            _p(_r('Date: ', bold=True), _r(today)),
                        ])

                    elif num == 2:
                        # Do Now
                        dn = lesson.get('doNow', {})
                        qs = dn.get('questions', [])
                        sp_b = get_sp(tree, ph_idx=1)
                        paras = [_p(_r('Retrieval Questions', bold=True))]
                        for i, q in enumerate(qs):
                            paras.append(_bullet(q, f'Q{i+1}'))
                        paras += [_empty(),
                            _p(_r('Teacher will: ', bold=True), _r(dn.get('teacherWill',''))),
                            _p(_r('Students will: ', bold=True), _r(dn.get('studentsWill','')))]
                        if sp_b: set_text(sp_b, paras)

                    elif num == 3:
                        # Learning Outcome + To Know
                        tk = lesson.get('toKnow', {})
                        lo_text = lesson.get('lo', '')
                        sp_lo   = get_sp(tree, ph_idx=1)
                        sp_know = get_sp(tree, ph_idx=10)
                        sp_sub  = get_sp(tree, ph_idx=11)
                        if sp_lo: set_text(sp_lo, [
                            _p(_r('Learning Outcome', bold=True)),
                            _p(_r(lo_text)),
                            _empty(),
                            _p(_r(f"Standard: {lesson.get('standard','')}"))
                        ])
                        tier2 = ', '.join(tk.get('tier2', []))
                        tier3 = ', '.join(tk.get('tier3', []))
                        know_paras = [
                            _p(_r('Tier 2 Vocabulary: ', bold=True), _r(tier2)),
                            _p(_r('Tier 3 Vocabulary: ', bold=True), _r(tier3)),
                            _empty(),
                            _p(_r('Key Concepts', bold=True))
                        ]
                        for c in tk.get('concepts', []):
                            know_paras.append(_bullet(c))
                        if sp_know: set_text(sp_know, know_paras)
                        if sp_sub:  set_text(sp_sub, [_p(_r(lesson.get('unit',''), bold=True))])

                    elif num == 4:
                        # I Do / We Do / You Do overview
                        ido = lesson.get('iDo', {}); wdo = lesson.get('weDo', {}); ydo = lesson.get('youDo', {})
                        sp_b = get_sp(tree, ph_idx=1)
                        if sp_b: set_text(sp_b, [
                            _p(_r('I Do', bold=True)),
                            _bullet(ido.get('teacherWill', '')),
                            _empty(),
                            _p(_r('We Do', bold=True)),
                            _bullet(wdo.get('activity', '')),
                            _empty(),
                            _p(_r('You Do', bold=True)),
                            _bullet(ydo.get('task', '')),
                        ])

                    elif num == 5:
                        # I Do detail
                        ido = lesson.get('iDo', {})
                        sp_b = get_sp(tree, ph_idx=1)
                        paras = []
                        for s in ido.get('steps', []):
                            paras.append(_p(_r(s.get('label',''), bold=True), _r(': ' + s.get('text',''))))
                        paras += [_empty(),
                            _p(_r('Think Aloud: ', bold=True), _r(f'"{ido.get("thinkAloud","")}"')),
                            _empty(),
                            _p(_r('Teacher will: ', bold=True), _r(ido.get('teacherWill',''))),
                            _p(_r('Students will: ', bold=True), _r(ido.get('studentsWill','')))]
                        if sp_b: set_text(sp_b, paras)

                    elif num == 6:
                        # We Do
                        wdo = lesson.get('weDo', {})
                        sp_b = get_sp(tree, ph_idx=1)
                        paras = [_p(_r(wdo.get('activity',''), bold=True))]
                        for b in wdo.get('bullets', []): paras.append(_bullet(b))
                        paras.append(_empty())
                        paras.append(_p(_r('CFU Questions', bold=True)))
                        for q in wdo.get('cfuQ', []): paras.append(_bullet(q))
                        paras += [_empty(),
                            _p(_r('Teacher will: ', bold=True), _r(wdo.get('teacherWill',''))),
                            _p(_r('Students will: ', bold=True), _r(wdo.get('studentsWill','')))]
                        if sp_b: set_text(sp_b, paras)

                    elif num == 7:
                        # You Do
                        ydo = lesson.get('youDo', {})
                        sp_b = get_sp(tree, ph_idx=1)
                        paras = [_p(_r(ydo.get('task',''), bold=True))]
                        for b in ydo.get('bullets', []): paras.append(_bullet(b))
                        paras.append(_empty())
                        paras.append(_p(_r('Decision-Making Questions', bold=True)))
                        for q in ydo.get('decisionQ', []): paras.append(_bullet(q))
                        paras += [_empty(),
                            _p(_r('Teacher will: ', bold=True), _r(ydo.get('teacherWill',''))),
                            _p(_r('Students will: ', bold=True), _r(ydo.get('studentsWill','')))]
                        if sp_b: set_text(sp_b, paras)

                    elif num == 8:
                        # Affirmative Checking
                        af = lesson.get('affirmative', {})
                        sp_b = get_sp(tree, ph_idx=1)
                        paras = []
                        for aq in af.get('questions', []):
                            paras.append(_p(_r('Q: ', bold=True), _r(aq.get('q',''))))
                            paras.append(_p(_r('Look for: ', bold=True), _r(aq.get('lookFor',''))))
                            paras.append(_empty())
                        paras += [
                            _p(_r('Most understood: ', bold=True), _r(af.get('mostUnderstood',''))),
                            _p(_r('Some struggled: ', bold=True), _r(af.get('someStruggled',''))),
                            _p(_r('Most confused: ', bold=True), _r(af.get('mostConfused','')))
                        ]
                        if sp_b: set_text(sp_b, paras)

                    elif num == 9:
                        # Exit Ticket
                        et = lesson.get('exitTicket', {})
                        sp_b = get_sp(tree, ph_idx=1)
                        wb = ' | '.join(et.get('wordBank', []))
                        paras = [
                            _p(_r('Q1 — Identify', bold=True)),
                            _p(_r(et.get('q1',''))),
                            _p(_r('Word Bank: ', bold=True), _r(wb)),
                            _empty(),
                            _p(_r('Q2 — Explain', bold=True)),
                            _p(_r(et.get('q2','')))
                        ]
                        if sp_b: set_text(sp_b, paras)

                    data = etree.tostring(tree, xml_declaration=True, encoding='UTF-8')
                
                zout.writestr(item, data)

    with open(output_path, 'wb') as f:
        f.write(buf.getvalue())
