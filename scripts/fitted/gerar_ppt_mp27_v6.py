"""PPT MP'27 Fitted Units – v6 (apresentação ao CEO, padrão premium do CP v8).

Uso:
    python gerar_ppt_mp27_v6.py <deck_template.pptx> <basebook.xlsx> <mens_v4.xlsx> <rentab_cliente_ibi_v3.xlsx>
                                <variacao_mo.xlsx> <volumes_mp27.xlsx> <pasta_logos> <saida.pptx>

O deck-template (a v5 salva por ela) só empresta a capa e o layout 'D. Title Only' (marca Pirelli, linha de título,
tema). Todos os demais slides são refeitos do zero com formas e tabelas nativas, direto das fontes:
  Base Book (Month_FY_LC + abas de unidade + Tabela Variantes), MENS v4 (TOTAL, IBI, Impacto JLR),
  Rentabilidade por Cliente IBI v3, Variação MO, Volumes MP'27_ADJ.
Pasta de logos: fiat.png, iveco.png, cnh.png, nissan.png, vw.png.
"""
import copy
import os
import sys

import openpyxl
from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

DECK, BB, MENS, CLI, MO, VOL, LOGOS, OUT = sys.argv[1:9]

# ------------------------------------------------------------------ paleta (igual ao CP v8)
NAVY = RGBColor(0x1F, 0x3A, 0x5F)
MID = RGBColor(0x2E, 0x5A, 0xA8)
HEAD = RGBColor(0x1F, 0x4E, 0x79)
SKY = RGBColor(0x9D, 0xC3, 0xE6)
LIGHT = RGBColor(0xD9, 0xDF, 0xEC)
PANEL = RGBColor(0xF3, 0xF5, 0xFA)
EDGE = RGBColor(0x8A, 0x99, 0xB5)
GREY = RGBColor(0x80, 0x80, 0x80)
DGREY = RGBColor(0x40, 0x40, 0x40)
BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
RED = RGBColor(0xC0, 0x39, 0x2B)
RULE = RGBColor(0xD0, 0xD5, 0xE0)
UCOL = {'IBI': NAVY, 'SJP': MID, 'GOI': RGBColor(0x7F, 0xB2, 0xE5), 'RES': RGBColor(0xC9, 0xD6, 0xEA)}
MINUS = '−'
LM, RM = 450000, 11742000          # margens de conteúdo
W_ALL = RM - LM


# ------------------------------------------------------------------ formatos pt-BR
def num(x, d=1):
    return f"{abs(x):,.{d}f}".replace(',', 'X').replace('.', ',').replace('X', '.')


def f1(x, d=1):
    s = num(x, d)
    return (MINUS + s) if x < 0 and round(abs(x), d) != 0 else s


def sg(x, d=1):
    if round(abs(x), d) == 0:
        return num(0, d)
    return (MINUS if x < 0 else '+') + num(x, d)


def par(x, d=1):
    return f"({num(x, d)})" if x < 0 and round(abs(x), d) != 0 else num(x, d)


def pc(x, d=1, signed=False):
    s = num(x * 100, d) + '%'
    return ((MINUS if x < 0 else '+') + s) if signed else (MINUS + s if x < 0 else s)


def fp(x, d=0):
    return num(x, d)


# ------------------------------------------------------------------ dados: Base Book
bb = openpyxl.load_workbook(BB, data_only=True)


def load_bridge(sheet, hr, dr):
    ws = bb[sheet]
    items, end = [], None
    for col in range(3, 20):
        h, v = ws.cell(hr, col).value, ws.cell(dr, col).value
        if h is None:
            continue
        h = str(h).strip()
        if h == "MP'27" or h.startswith('Delta') or h.startswith('%'):
            continue
        if h == 'EBIT':
            end = float(v)
            break
        items.append((h.replace('\n', ' '), float(v)))
    assert abs(items[0][1] + sum(v for _, v in items[1:]) - end) < 0.003, (sheet, dr)
    return {'start': items[0][1], 'deltas': items[1:], 'end': end, 'd': dict(items[1:])}


BR_SRC = {'BR': ('Tabela Variantes s targ', 2, 4), 'IBI': ('Tabela Variantes s targ', 6, 8),
          'SJP': ('Tabela Variantes CM s targ', 3, 5), 'GOI': ('Tabela Variantes CM s targ', 8, 10),
          'RES': ('Tabela Variantes CM s targ', 13, 15)}
BRG = {u: load_bridge(*s) for u, s in BR_SRC.items()}
PL_SHEET = {'BR': 'Month_FY_LC - MP27', 'IBI': 'IBIRITE', 'SJP': 'SJP', 'GOI': 'GOIANA', 'RES': 'RESENDE'}
PL_ROWS = [('pcs', 42), ('ns', 44), ('nspc', 45), ('cm', 47), ('cmp', 48), ('fix', 51), ('oi', 53),
           ('ebit', 55), ('ebpc', 56), ('ros', 58)]
COLS, DCOLS = [39, 40, 41, 42, 43], [45, 46]          # Act24, Act25, MP26, R9, MP27 | vs R9, vs MP26


def load_pl(u):
    ws = bb[PL_SHEET[u]]
    assert ws.cell(39, 43).value == "MP'27" and ws.cell(39, 39).value == 'Actual 2024', u
    d = {}
    for k, r in PL_ROWS:
        d[k] = [float(ws.cell(r, c).value or 0) for c in COLS]
        d['d_' + k] = [float(ws.cell(r, c).value or 0) for c in DCOLS]
    return d


PL = {u: load_pl(u) for u in PL_SHEET}
UNITS = ['IBI', 'SJP', 'GOI', 'RES']
for u in BRG:
    assert abs(BRG[u]['end'] - PL[u]['ebit'][4]) < 0.002 and abs(BRG[u]['start'] - PL[u]['ebit'][3]) < 0.002, u
assert abs(sum(PL[u]['ebit'][4] for u in UNITS) - PL['BR']['ebit'][4]) < 0.002
R = lambda u, k, i=4: PL[u][k][i]         # i: 2 = MP26, 3 = R9, 4 = MP27
DD = lambda u, k: BRG[u]['d'][k]

# ------------------------------------------------------------------ dados: MENS v4
mens = openpyxl.load_workbook(MENS, data_only=True)
jl = mens['Impacto JLR']
JLR = {}
for r in range(1, 60):
    k = jl.cell(r, 3).value
    if k in ('Fiat', 'Iveco', 'CNH') and k not in JLR:
        JLR[k] = [jl.cell(r, c).value for c in range(4, 14)]   # vol R9, vol 27, preço R9, preço 27, fat R9, fat 27, ...
FIAT_P9, FIAT_P27, FIAT_REAJ = JLR['Fiat'][2], JLR['Fiat'][3], JLR['Fiat'][9] / 1000   # efeito preço (mi)
FIAT_PCT = FIAT_P27 / FIAT_P9 - 1
IVECO_V9, IVECO_V27 = JLR['Iveco'][0], JLR['Iveco'][1]
mt = mens['TOTAL']
trow = {mt.cell(r, 3).value: r for r in range(1, 60) if mt.cell(r, 3).value}
tv = lambda k: float(mt.cell(trow[k], 17).value)
SENS_VOL = (tv('NET SALES') + tv('Direct Materials') + tv('Transportation') + tv('Other Variable')) / 1000 / 100
SENS_PRECO = JLR['Fiat'][1] * 0.10 / 1e6            # +R$ 0,10/pç em todo o volume Fiat (mi)
mi_ibi = mens['IBI']
OTHER_FIXED_V4 = float(mi_ibi.cell(37, 17).value) / 1000

# ------------------------------------------------------------------ dados: rentabilidade por cliente
wbc = openpyxl.load_workbook(CLI, data_only=True)
wc = wbc['Rentabilidade por Cliente']
crow = lambda r: [float(wc.cell(r, c).value or 0) for c in (2, 3, 4, 5)]
assert str(wc['A23'].value).startswith('Realocacao') and str(wc['A24'].value) == 'EBIT'
C_PCS, C_NS, C_NSPC, C_CM, C_FIX, C_REAL, C_EBIT, C_ROS = (crow(5), crow(6), crow(7), crow(14), crow(16), crow(23),
                                                           crow(24), crow(25))
assert abs(C_EBIT[3] / 1000 - R('IBI', 'ebit')) < 0.002, (C_EBIT[3], R('IBI', 'ebit'))
ROS_META = float(wbc['Premissas']['B23'].value)
wcd = wbc['Dados IBI']
assert wcd['A4'].value.startswith('Volume') and wcd['B4'].value == 'CNH'
CNH_VOL = [float(wcd.cell(4, c).value or 0) for c in range(3, 15)]
CNH_ZERO = [m for m in range(12) if CNH_VOL[m] == 0]
sh_ns = [C_NS[i] / C_NS[3] for i in range(3)]
sh_eb = [C_EBIT[i] / C_EBIT[3] for i in range(3)]

# ------------------------------------------------------------------ dados: mão de obra (R$ mil)
wm = openpyxl.load_workbook(MO, data_only=True)['Resumo (4 grupos)']
MOD = {}
for r in range(5, 40):
    u, g = wm.cell(r, 1).value, wm.cell(r, 2).value
    if u in ('SJP', 'IBI', 'GOI', 'RES', 'GER') and g and g != 'Total MO':
        MOD.setdefault(u, {})[g] = tuple(float(wm.cell(r, c).value) for c in (3, 4, 5))
GR = ['Salarios e PPR', 'Encargos', 'Assistencia medica', 'Outros beneficios']


def mo(u, g=None, i=2):                # i: 0 = R9 ano, 1 = run-rate, 2 = MP27
    us = [u] if u != 'ALL' else list(MOD)
    gs = [g] if g else GR
    return sum(MOD[x][y][i] for x in us for y in gs) / 1000


assert abs((mo('ALL') - mo('ALL', i=0)) + DD('BR', 'Labour')) < 0.05
GR_SAL = ['Salarios e PPR', 'Encargos', 'Outros beneficios']
sal = lambda u, i: sum(mo(u, x, i) for x in GR_SAL)
med = lambda u, i: mo(u, 'Assistencia medica', i)
ANO_CHEIO = {u: mo(u, i=1) - mo(u, i=0) for u in UNITS + ['GER']}
REAJ_SAL, REAJ_MED = sal('ALL', 2) - sal('ALL', 1), med('ALL', 2) - med('ALL', 1)
MO9, MORR, MO27 = mo('ALL', i=0), mo('ALL', i=1), mo('ALL')

# ------------------------------------------------------------------ dados: volumes por projeto
wv = openpyxl.load_workbook(VOL, data_only=True)


def vol_row(sheet, label, after=None):
    ws = wv[sheet]
    start = 1
    if after:
        start = next(r for r in range(1, ws.max_row + 1) if str(ws.cell(r, 2).value or '').strip() == after)
    for r in range(start, ws.max_row + 1):
        if str(ws.cell(r, 1).value or '').strip().lower() == label.lower():
            return float(ws.cell(r, 14).value)
    raise KeyError((sheet, label))


TUKAN = vol_row('SJP - VW', 'Produção (pç Tukan)', "MP'27") / 1000
VIRTUS = vol_row('SJP - VW', 'Produção (pç Virtus)', "MP'27") / 1000
NISSAN = (vol_row('Resende - Nissan', 'Produção pç Kait', "MP'27")
          + vol_row('Resende - Nissan', 'Produção pç Kicks', "MP'27")) / 1000
assert abs(NISSAN - R('RES', 'pcs')) < 1, (NISSAN, R('RES', 'pcs'))
NB_PCS = NISSAN + TUKAN + C_PCS[2]

# ------------------------------------------------------------------ números-chave
BRp = PL['BR']
e9, e27, e26 = BRp['ebit'][3], BRp['ebit'][4], BRp['ebit'][2]
ns9, ns27 = BRp['ns'][3], BRp['ns'][4]
ros9, ros27 = BRp['ros'][3], BRp['ros'][4]
pcs9, pcs27 = BRp['pcs'][3], BRp['pcs'][4]
ANTT = -DD('IBI', 'ANTT')
SOR = e9 - sum(R(u, 'ebit', 3) for u in UNITS)
RISCO_MO = 0.95 + 0.46 + 0.36             # Auditoria Labour (achados 1, 3 e 7)
FIAT_EBIT_SHARE = C_EBIT[0] / 1000 / e27
CAMPOS_MENS = {'other_fixed_v4': OTHER_FIXED_V4}

# ================================================================ helpers de desenho
prs = Presentation(DECK)
LAYOUT = [l for l in prs.slide_layouts if l.name == 'D. Title Only'][0]
# mantém só a capa
sldIdLst = prs.slides._sldIdLst
for sldId in list(sldIdLst)[1:]:
    prs.part.drop_rel(sldId.rId)
    sldIdLst.remove(sldId)


def rich(shapes, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP, name=None, wrap=True, margins=(0, 0, 0, 0), shape=None):
    """paras: lista de parágrafos; cada um = dict(runs=[(texto, {size,bold,color,italic})], align, space_after, bullet)
    ou lista de runs."""
    tb = shape or shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = [Emu(m) for m in margins]
    tf.vertical_anchor = anchor
    first = True
    for para in paras:
        if isinstance(para, list):
            para = {'runs': para}
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = para.get('align', PP_ALIGN.LEFT)
        if 'space_after' in para:
            p.space_after = Pt(para['space_after'])
        if 'space_before' in para:
            p.space_before = Pt(para['space_before'])
        if 'line' in para:
            p.line_spacing = para['line']
        if para.get('bullet'):
            pPr = p._p.get_or_add_pPr()
            pPr.set('marL', '150000')
            pPr.set('indent', '-150000')
            pPr.append(pPr.makeelement(qn('a:buFont'), {'typeface': 'Arial'}))
            pPr.append(pPr.makeelement(qn('a:buChar'), {'char': '•'}))
        for t, f in para['runs']:
            r = p.add_run()
            r.text = t
            r.font.size = Pt(f.get('size', 11))
            r.font.bold = f.get('bold', False)
            r.font.italic = f.get('italic', False)
            r.font.color.rgb = f.get('color', BLACK)
    if name:
        tb.name = name
    return tb


def T(text, size=11, bold=False, color=BLACK, italic=False):
    return (text, {'size': size, 'bold': bold, 'color': color, 'italic': italic})


def txt(shapes, x, y, w, h, text, size=11, bold=False, color=BLACK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE,
        name=None, italic=False, wrap=False):
    paras = [{'runs': [T(ln, size, bold, color, italic)], 'align': align} for ln in text.split('\n')]
    return rich(shapes, x, y, w, h, paras, anchor, name, wrap)


def box(shapes, x, y, w, h, fill, name='Forma', shape=MSO_SHAPE.RECTANGLE, line_rgb=None, lw=0.75, radius=None):
    b = shapes.add_shape(shape, Emu(x), Emu(y), Emu(w), Emu(h))
    if fill is None:
        b.fill.background()
    else:
        b.fill.solid()
        b.fill.fore_color.rgb = fill
    if line_rgb is None:
        b.line.fill.background()
    else:
        b.line.color.rgb = line_rgb
        b.line.width = Pt(lw)
    b.shadow.inherit = False
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        b.adjustments[0] = radius
    b.name = name
    return b


def line(shapes, x1, y1, x2, y2, color=GREY, w=0.75, dash=None, arrow=False, name=None):
    c = shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(x1), Emu(y1), Emu(x2), Emu(y2))
    c.line.color.rgb = color
    c.line.width = Pt(w)
    ln = c.line._get_or_add_ln()
    if dash:
        ln.append(ln.makeelement(qn('a:prstDash'), {'val': dash}))
    if arrow:
        ln.append(ln.makeelement(qn('a:tailEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'}))
    if name:
        c.name = name
    return c


def new_slide(title, headline, tag="(+) ganho  |  R$ milhões", notes=None):
    s = prs.slides.add_slide(LAYOUT)
    s.shapes.title.text = title
    tidx = s.shapes.title.placeholder_format.idx
    for ph in list(s.placeholders):
        if ph.placeholder_format.idx != tidx:
            ph.element.getparent().remove(ph.element)
    if tag:
        txt(s.shapes, LM, 520000, 4000000, 190000, tag, 9, False, DGREY, name='Unidade', italic=True)
    rich(s.shapes, LM, 740000, W_ALL, 600000, [[T(headline, 18, True, HEAD)]], MSO_ANCHOR.TOP, 'Mensagem principal')
    if notes:
        s.notes_slide.notes_text_frame.text = notes
    return s


def footnote(s, text, w=9700000, y=6480000):
    txt(s.shapes, LM, y, w, 260000, text, 8, False, GREY, anchor=MSO_ANCHOR.TOP, name='Nota de rodape', italic=True, wrap=True)


def card(s, x, y, w, h, name='Card', accent=MID):
    box(s.shapes, x, y, w, h, PANEL, name)
    box(s.shapes, x, y, 45000, h, accent, name + ' barra')


def side_panel(s, x, y, w, h, title, blocks, size=11.5, gap=9):
    """Painel 'Resumo executivo' do CP: barra azul à esquerda, blocos com título azul + texto, régua entre blocos."""
    card(s, x, y, w, h, 'Painel')
    paras = [{'runs': [T(title, 15, True, NAVY, italic=True)], 'space_after': 8}]
    for i, (head, body) in enumerate(blocks):
        paras.append({'runs': [T(head, size + 0.5, True, HEAD)], 'space_after': 2, 'space_before': 0 if i == 0 else gap})
        for ln in ([body] if isinstance(body, str) else body):
            paras.append({'runs': [T(ln, size, False, DGREY)], 'space_after': 2, 'line': 1.0})
    rich(s.shapes, x + 190000, y + 130000, w - 330000, h - 220000, paras, name='Painel texto')


# ---------------------------------------------------------------- tabelas nativas
def _no_style(gt):
    tbl = gt.table._tbl
    pr = tbl.tblPr
    for a in ('firstRow', 'bandRow', 'firstCol', 'lastRow', 'lastCol', 'bandCol'):
        if a in pr.attrib:
            del pr.attrib[a]
    for e in pr.findall(qn('a:tableStyleId')):
        pr.remove(e)
    sid = etree.SubElement(pr, qn('a:tableStyleId'))
    sid.text = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'          # sem estilo, sem grade


def _borders(cell, spec):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
        for e in tcPr.findall(qn(tag)):
            tcPr.remove(e)
    els = []
    for side in ('L', 'R', 'T', 'B'):
        ln = etree.Element(qn('a:ln' + side))
        if side in spec:
            col, w = spec[side]
            ln.set('w', str(int(w * 12700)))
            sf = etree.SubElement(ln, qn('a:solidFill'))
            etree.SubElement(sf, qn('a:srgbClr')).set('val', col)
        else:
            ln.set('w', '0')
            etree.SubElement(ln, qn('a:noFill'))
        els.append(ln)
    for k, e in enumerate(els):
        tcPr.insert(k, e)


def table(s, x, y, colw, rows, rowh, name='Tabela', font=10.5, head_rows=1, head_fill=NAVY, bold_rows=(), fill_rows=None,
          aligns=None, col_fill=None, grey_cols=(), cell_style=None, head_h=None):
    """rows: lista de listas de texto. cell_style(i,j)-> dict(color,bold,fill,italic,size) opcional."""
    n, m = len(rows), len(rows[0])
    hh = head_h or rowh
    total = sum((hh if i < head_rows else rowh) for i in range(n))
    gt = s.shapes.add_table(n, m, Emu(x), Emu(y), Emu(sum(colw)), Emu(total))
    _no_style(gt)
    tb = gt.table
    for j, wd in enumerate(colw):
        tb.columns[j].width = Emu(wd)
    for i, row in enumerate(rows):
        tb.rows[i].height = Emu(hh if i < head_rows else rowh)
        for j, val in enumerate(row):
            c = tb.cell(i, j)
            c.margin_left = c.margin_right = Emu(70000)
            c.margin_top = c.margin_bottom = Emu(15000)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            st = (cell_style(i, j) if cell_style else None) or {}
            is_head = i < head_rows
            fill = st.get('fill', head_fill if is_head else (fill_rows or {}).get(i))
            if fill is None:
                c.fill.background()
            else:
                c.fill.solid()
                c.fill.fore_color.rgb = fill
            p = c.text_frame.paragraphs[0]
            c.text_frame.word_wrap = True
            al = (aligns[j] if aligns else (PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT))
            p.alignment = PP_ALIGN.CENTER if is_head and j > 0 and not (aligns and aligns[j] == PP_ALIGN.LEFT) else al
            r = p.add_run()
            r.text = str(val)
            r.font.size = Pt(st.get('size', font))
            r.font.bold = st.get('bold', is_head or i in bold_rows)
            r.font.italic = st.get('italic', False)
            r.font.color.rgb = st.get('color', WHITE if is_head else (GREY if j in grey_cols else BLACK))
            _borders(c, {} if is_head else {'B': ('D5D9E3', 0.5)})
    gt.name = name
    return gt


# ---------------------------------------------------------------- bridge (cascata) com formas nativas
def waterfall(s, x, y, w, h, start_lbl, start, deltas, end_lbl, end, pill=None, label_h=470000, head=620000,
              name='Bridge', val_size=11, tot_size=14, bad_up=False):
    assert abs((start or 0) + sum(v for _, v in deltas) - end) < 0.006, (name, start, deltas, end)
    cols = ([(start_lbl, 'tot', start)] if start is not None else []) + [(k, 'd', v) for k, v in deltas] + [(end_lbl, 'tot', end)]
    bars, cum = [], 0.0
    for _, kind, v in cols:
        if kind == 'tot':
            lo, hi, cum = min(0, v), max(0, v), v
        else:
            lo, hi = min(cum, cum + v), max(cum, cum + v)
            cum += v
        bars.append((lo, hi, cum))
    vmax = max(max(b[1] for b in bars), 0)
    vmin = min(min(b[0] for b in bars), 0)
    top, bot = y + head, y + h - label_h
    sc = (bot - top - (120000 if vmin < 0 else 0)) / (vmax - vmin)
    yv = lambda v: int(top + (vmax - v) * sc)
    n = len(cols)
    slot = w / n
    bw = int(slot * 0.62)
    sh = s.shapes
    line(sh, x, yv(0), x + w, yv(0), RULE, 0.75, name=name + ' eixo')
    xc = [int(x + slot * (i + 0.5)) for i in range(n)]
    for i, ((lbl, kind, v), (lo, hi, cum)) in enumerate(zip(cols, bars)):
        y1, y2 = yv(hi), yv(lo)
        hgt = max(y2 - y1, 14000)
        if kind == 'tot':
            b = box(sh, xc[i] - bw // 2, y1, bw, hgt, NAVY, f'{name} barra {lbl}')
        elif (v >= 0) != bad_up:
            b = box(sh, xc[i] - bw // 2, y1, bw, hgt, SKY, f'{name} barra {lbl}', line_rgb=RGBColor(0x6F, 0xA0, 0xD6), lw=0.5)
        else:
            b = box(sh, xc[i] - bw // 2, y1, bw, hgt, LIGHT, f'{name} barra {lbl}', line_rgb=EDGE, lw=0.5)
        # rótulo de valor
        if kind == 'tot':
            lab, col, sz, bd = f1(v), NAVY, tot_size, True
        else:
            lab, col, sz, bd = sg(v), (GREEN if (v > 0) != bad_up else RED), val_size, True
        txt(sh, xc[i] - int(slot / 2), y1 - 255000, int(slot), 235000, lab, sz, bd, col, PP_ALIGN.CENTER, MSO_ANCHOR.BOTTOM,
            name=f'{name} valor {lbl}')
        txt(sh, xc[i] - int(slot / 2) + 10000, bot + 70000 + (120000 if vmin < 0 else 0), int(slot) - 20000, label_h - 70000, lbl, 10,
            kind == 'tot', DGREY, PP_ALIGN.CENTER, MSO_ANCHOR.TOP, name=f'{name} categoria {lbl}', wrap=True)
        if i < n - 1:
            lvl = cum if kind != 'tot' else v
            line(sh, xc[i] + bw // 2, yv(lvl), xc[i + 1] - bw // 2, yv(lvl), GREY, 0.5, dash='dash', name=f'{name} conector')
    if pill:
        py = y + 20000
        x1, x2 = xc[0] - bw // 2, xc[-1] + bw // 2
        pw, ph = 1000000, 270000
        xm = (x1 + x2) // 2
        line(sh, x1, py + ph // 2, x2, py + ph // 2, GREY, 0.75, name=name + ' colchete')
        line(sh, x1, py + ph // 2, x1, yv(bars[0][1]) - 275000, GREY, 0.75, name=name + ' colchete esq')
        line(sh, x2, py + ph // 2, x2, yv(bars[-1][1]) - 275000, GREY, 0.75, arrow=True, name=name + ' colchete dir')
        pl_ = box(sh, xm - pw // 2, py, pw, ph, WHITE, name + ' pill', MSO_SHAPE.ROUNDED_RECTANGLE, line_rgb=GREY, lw=0.75, radius=0.5)
        rich(sh, 0, 0, 0, 0, [{'runs': [T(pill, 12, True, BLACK)], 'align': PP_ALIGN.CENTER}], MSO_ANCHOR.MIDDLE, shape=pl_)


def logo(s, key, x, y, w, h):
    p = os.path.join(LOGOS, key + '.png')
    iw, ih = Image.open(p).size
    sc = min(w / iw, h / ih)
    pw, ph = int(iw * sc), int(ih * sc)
    return s.shapes.add_picture(p, Emu(x + (w - pw) // 2), Emu(y + (h - ph) // 2), Emu(pw), Emu(ph))


# ================================================================ SLIDE 2 – resumo executivo
def kpi(s, x, y, w, h, label, value, sub, sub_col, accent=MID):
    card(s, x, y, w, h, 'KPI ' + label, accent)
    txt(s.shapes, x + 190000, y + 85000, w - 260000, 220000, label.upper(), 9.5, True, GREY, name='KPI rotulo')
    txt(s.shapes, x + 190000, y + 300000, w - 260000, 480000, value, 28, True, NAVY, name='KPI valor')
    txt(s.shapes, x + 190000, y + 800000, w - 260000, 230000, sub, 10.5, True, sub_col, name='KPI delta')


def slide_resumo():
    s = new_slide("Fitted Units Economics MP'27 – Resumo executivo",
                  f"EBIT de R$ {f1(e27)} mi e ROS de {pc(ros27)}: {sg(e27 - e9)} mi vs R9, mas {sg(e27 - e26)} mi vs MP'26 e acima do target de 13%",
                  notes=("Fontes: Base Book Fitted Units MP27 (Month_FY_LC e abas de unidade), MENS FITTED MP27 v4 (TOTAL, IBI, Impacto JLR), "
                         "Rentabilidade por Cliente IBI v3, Variação MO. R9 Brasil (13,0) inclui 0,2 da SOR encerrada; no perímetro das 4 unidades o R9 é 12,8."))
    kw, kg, ky, kh = 2680000, 190000, 1400000, 1090000
    d_pcs, d_ns, d_e = pcs27 / pcs9 - 1, ns27 / ns9 - 1, e27 / e9 - 1
    kp = [("Volume", f"{fp(pcs27)} mil pç", f"{pc(d_pcs, 1, True)} vs R9", GREEN),
          ("Net Sales", f"R$ {f1(ns27)} mi", f"{pc(d_ns, 1, True)} vs R9", GREEN),
          ("EBIT", f"R$ {f1(e27)} mi", f"{pc(d_e, 1, True)} vs R9  ·  {pc(e27 / e26 - 1, 1, True)} vs MP'26", RED),
          ("ROS", pc(ros27), f"{sg((ros27 - ros9) * 100)} p.p. vs R9 ({pc(ros9)})", RED)]
    for i, (a, b, c, col) in enumerate(kp):
        kpi(s, LM + i * (kw + kg), ky, kw, kh, a, b, c, col)
    # tabela por unidade
    y0 = 2650000
    txt(s.shapes, LM, y0, 6000000, 260000, "EBIT por unidade – R$ mi", 13, True, NAVY)
    msgs = {'IBI': 'Preço Fiat e fixos compensam MO, ANTT e Jaguar', 'SJP': 'Desconto ao cliente e MO pressionam a margem',
            'GOI': 'Preço compensa o menor volume', 'RES': 'Ano cheio, ROS de 10%'}
    rows = [["Unidade", "Peças (k)", "Net Sales", "EBIT R9", "EBIT MP'27", "Δ vs R9", "ROS", "Leitura"]]
    for u in UNITS:
        rows.append([u, fp(R(u, 'pcs')), f1(R(u, 'ns')), par(R(u, 'ebit', 3)), par(R(u, 'ebit')),
                     par(R(u, 'ebit') - R(u, 'ebit', 3)), pc(R(u, 'ros')), msgs[u]])
    rows.append(["Brasil", fp(pcs27), f1(ns27), par(e9), par(e27), par(e27 - e9), pc(ros27), "R9 inclui 0,2 da SOR (encerrada)"])
    cw = [820000, 800000, 800000, 780000, 900000, 780000, 700000]
    cw.append(7500000 - sum(cw))
    last = len(rows) - 1

    def cs(i, j):
        if i == 0:
            return {'size': 10}
        st = {'fill': LIGHT, 'bold': True} if i == last else {}
        if j == 5:
            v = float(rows[i][j].replace('(', '-').replace(')', '').replace(',', '.'))
            st.update({'color': GREEN if v > 0.04 else (RED if v < -0.04 else DGREY), 'bold': True})
        if j == 7:
            st.update({'size': 10, 'color': DGREY, 'bold': False})
        return st
    table(s, LM, y0 + 300000, cw, rows, 300000, 'Tabela unidades', font=11,
          aligns=[PP_ALIGN.LEFT] + [PP_ALIGN.RIGHT] * 6 + [PP_ALIGN.LEFT], cell_style=cs)
    # blocos: sustenta / pressiona
    yb, hb = 4950000, 1430000
    bw_ = (7500000 - 180000) // 2
    sust = [f"Volume {sg(DD('BR', 'Volume'))}: RES em ano cheio e IBI",
            f"Preço/Mix {sg(DD('BR', 'Price/ Mix'))}: reajuste Fiat IBI ({sg(FIAT_REAJ)})",
            f"Custos fixos {sg(DD('BR', 'Cost Fix.'))}: base fixa de IBI revisada"]
    pres = [f"Mão de obra {f1(DD('BR', 'Labour'))} ({pc(MO27 / MO9 - 1, 1, True)}): 3º turno IBI e RES",
            f"ANTT {f1(DD('BR', 'ANTT'))}  ·  Jaguar {f1(DD('BR', 'Encerr. Jaguar'))}",
            f"Custos variáveis {f1(DD('BR', 'Cost Var.'))}"]
    for k, (ttl, items, acc) in enumerate((("O que sustenta o plano", sust, GREEN), ("O que pressiona", pres, RED))):
        x = LM + k * (bw_ + 180000)
        card(s, x, yb, bw_, hb, 'Bloco ' + ttl, acc)
        rich(s.shapes, x + 190000, yb + 110000, bw_ - 260000, hb - 150000,
             [{'runs': [T(ttl, 13, True, HEAD)], 'space_after': 6}] +
             [{'runs': [T(t, 11.5, False, DGREY)], 'bullet': True, 'space_after': 5} for t in items], name='Bloco texto')
    # decisões
    xd = LM + 7500000 + 220000
    wd = RM - xd
    side_panel(s, xd, y0, wd, 6380000 - y0, "Decisões-chave", [
        ("1  Fechar o reajuste Fiat", f"R$ {num(FIAT_P27, 2)}/pç ({pc(FIAT_PCT, 1, True)} vs R9) = {sg(FIAT_REAJ)} mi de EBIT"),
        ("2  Repassar a ANTT ao cliente", f"Nova regra de frete mínimo: custo de {f1(ANTT)} mi no plano"),
        ("3  Validar riscos de MO com o RH", f"Provisão de férias horista, SJP CC 8289 e GOI: até {f1(RISCO_MO, 1)} mi"),
        ("4  Sensibilidade a volume", f"Cada 1% de peças ≈ {num(SENS_VOL, 1)} mi de EBIT")], size=11.5, gap=14)
    footnote(s, "Fonte: Base Book Fitted Units MP27, MENS FITTED MP27 v4. ROS = EBIT / Net Sales. Target open book ≈ 13%.")
    return s


# ================================================================ SLIDES de P&L + bridge (Brasil e unidades)
def pl_rows(u):
    p = PL[u]
    cm_pct = [c / n if n else 0 for c, n in zip(p['cm'], p['ns'])]
    z = lambda vals, f: [('–' if abs(v) < 1e-9 else f(v)) for v in vals]
    rows = [["", "Actual '24", "Actual '25", "MP'26", "Fcst R9", "MP'27", "vs R9", "vs MP'26"],
            ["Peças (mil)"] + z(p['pcs'], fp) + [par(p['d_pcs'][0], 0), par(p['d_pcs'][1], 0)],
            ["Net Sales"] + z(p['ns'], f1) + [par(p['d_ns'][0]), par(p['d_ns'][1])],
            ["   NS / peça (R$)"] + z(p['nspc'], lambda v: num(v, 2)) + [par(p['d_nspc'][0], 2), par(p['d_nspc'][1], 2)],
            ["Margem de contribuição"] + z(p['cm'], f1) + [par(p['d_cm'][0]), par(p['d_cm'][1])],
            ["   % s/ NS"] + z(cm_pct, pc) + [sg((cm_pct[4] - cm_pct[3]) * 100) + ' pp', sg((cm_pct[4] - cm_pct[2]) * 100) + ' pp'],
            ["Custos fixos"] + z(p['fix'], par) + [par(p['d_fix'][0]), par(p['d_fix'][1])]]
    has_oi = any(abs(v) > 1e-6 for v in p['oi'])
    if has_oi:
        rows.append(["Outras receitas"] + z(p['oi'], par) + [par(p['d_oi'][0]), par(p['d_oi'][1])])
    rows.append(["EBIT"] + z(p['ebit'], par) + [par(p['d_ebit'][0]), par(p['d_ebit'][1])])
    rows.append(["   ROS"] + z(p['ros'], pc) + [sg((p['ros'][4] - p['ros'][3]) * 100) + ' pp', sg((p['ros'][4] - p['ros'][2]) * 100) + ' pp'])
    return rows


def pl_table(s, u, x, y, w, rowh=205000):
    rows = pl_rows(u)
    n = len(rows)
    ebit_i = n - 2
    first = 330000
    cw = [1780000] + [int((w - 1780000) / 7)] * 7
    cw[-1] = w - sum(cw[:-1])

    def cs(i, j):
        if i == 0:
            return {'fill': (SKY if j >= 6 else NAVY), 'color': (NAVY if j >= 6 else WHITE), 'size': 10, 'bold': True}
        lbl = rows[i][0]
        base = {}
        if not lbl.startswith('   '):
            base['bold'] = True
        else:
            base.update({'color': GREY, 'size': 9.5, 'italic': j == 0})
        if i == ebit_i:
            base.update({'fill': LIGHT, 'bold': True, 'size': 11.5})
        if j == 5 and not lbl.startswith('   '):
            base['color'] = NAVY
        if j >= 6 and i != 0 and lbl.strip() not in ('Peças (mil)', 'NS / peça (R$)', '% s/ NS', 'ROS'):
            pass
        return base
    return table(s, x, y, cw, rows, rowh, f'P&L {u}', font=10.5, head_h=first, cell_style=cs), y + first + rowh * (n - 1)


def unit_slide(u, title, headline, blocks, notes, tag="(+) ganho  |  R$ milhões", pl_w=7550000, extra_foot=None):
    s = new_slide(title, headline, tag, notes)
    gt, ybot = pl_table(s, u, LM, 1430000, pl_w)
    yb = ybot + 90000
    b = BRG[u]
    pct = b['end'] / b['start'] - 1
    pill = f"{pc(pct, 1, True)}" if b['start'] > 0.5 else f"Δ {sg(b['end'] - b['start'])} mi"
    labels = {'Price/ Mix': 'Price /\nMix', 'Encerr. Jaguar': 'Encerr.\nJaguar', 'Cost Var.': 'Cost\nVar.', 'Cost Fix.': 'Cost\nFix.',
              'Volume': 'Volume', 'Labour': 'Labour', 'ANTT': 'ANTT'}
    deltas = [(labels.get(k, k), v) for k, v in b['deltas']]
    waterfall(s, LM - 60000, yb, pl_w + 60000, 6400000 - yb, 'R9', b['start'], deltas, "MP'27", b['end'], pill=pill,
              name=f'Bridge {u}')
    xp = LM + pl_w + 230000
    side_panel(s, xp, 1430000, RM - xp, 6380000 - 1430000, "Principais mensagens", blocks, size=11.5, gap=14)
    footnote(s, "Fonte: Base Book Fitted Units MP27 (abas de unidade e Tabela Variantes)." + (' ' + extra_foot if extra_foot else ''),
             w=pl_w)
    return s


def slide_brasil():
    sal_ = sg(DD('BR', 'Volume'))
    return unit_slide('BR', "Fitted Units Economics MP'27 – Brasil",
                      f"Receita cresce {pc(ns27 / ns9 - 1, 1)}, mas MO, ANTT e Jaguar superam o ganho de preço e fixos: EBIT {sg(e27 - e9)} mi vs R9",
                      [(f"EBIT {f1(e27)}  |  ROS {pc(ros27)}", f"{sg(e27 - e9)} vs R9 ({pc(e27 / e9 - 1, 1, True)})  ·  {sg(e27 - e26)} vs MP'26"),
                       (f"Receita {pc(ns27 / ns9 - 1, 1, True)}", [f"Volume {sal_}: RES em ano cheio e IBI",
                                                                  f"Price/Mix {sg(DD('BR', 'Price/ Mix'))}: reajuste Fiat IBI"]),
                       ("Custos e eventos pressionam", [f"Mão de obra {f1(DD('BR', 'Labour'))}: 3º turno IBI e RES",
                                                       f"ANTT {f1(DD('BR', 'ANTT'))}  ·  Jaguar {f1(DD('BR', 'Encerr. Jaguar'))}  ·  variáveis {f1(DD('BR', 'Cost Var.'))}"]),
                       (f"Custos fixos {sg(DD('BR', 'Cost Fix.'))}", "Base fixa de IBI revisada; efeito favorável vs R9"),
                       ("Ponto de atenção", f"Sem o reajuste Fiat ({sg(FIAT_REAJ)}), o EBIT cai para ~{f1(e27 - FIAT_REAJ)}")],
                      f"Bridge Brasil (Base Book, Tabela Variantes s targ), R$ mi: R9 {b_str('BR')}",
                      extra_foot="R9 Brasil (13,0) inclui 0,2 da SOR encerrada.")


def b_str(u):
    b = BRG[u]
    return f"{f1(b['start'], 3)} | " + " | ".join(f"{k} {f1(v, 3)}" for k, v in b['deltas']) + f" | MP'27 {f1(b['end'], 3)}"


# ================================================================ mix do EBIT
def donut(s, x, y, dsz, order, vals, big, small):
    tot = sum(vals)
    cd = CategoryChartData()
    cd.categories = order
    cd.add_series('EBIT', [max(v, 0) for v in vals])
    gf = s.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Emu(x), Emu(y), Emu(dsz), Emu(dsz), cd)
    gf.name = 'Pizza EBIT'
    ch = gf.chart
    ch.has_legend = False
    ch.has_title = False
    ser = ch.plots[0].series[0]
    for i, u in enumerate(order):
        pt = ser.points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = UCOL[u]
        pt.format.line.color.rgb = WHITE
        pt.format.line.width = Pt(2.25)
        share = vals[i] / tot
        tf = pt.data_label.text_frame
        tf.text = f"{round(share * 100)}%" if share >= 0.09 else ''
        if tf.paragraphs[0].runs:
            r = tf.paragraphs[0].runs[0]
            r.font.size = Pt(13)
            r.font.bold = True
            r.font.color.rgb = NAVY if u in ('GOI', 'RES') else WHITE
    for el in ch._chartSpace.xpath('.//c:holeSize'):
        el.set('val', '64')
    cx, cy = x + dsz // 2, y + dsz // 2
    txt(s.shapes, cx - 600000, cy - 290000, 1200000, 330000, big, 24, True, NAVY, PP_ALIGN.CENTER, MSO_ANCHOR.BOTTOM, 'Centro valor')
    txt(s.shapes, cx - 600000, cy + 50000, 1200000, 200000, small, 10, False, GREY, PP_ALIGN.CENTER, MSO_ANCHOR.TOP, 'Centro legenda')


def slide_mix():
    r9 = [R(u, 'ebit', 3) for u in UNITS]
    mp = [R(u, 'ebit') for u in UNITS]
    t9, tm = sum(r9), sum(mp)
    ibi_sjp = (mp[0] + mp[1]) / tm
    s = new_slide("Fitted Units Economics MP'27 – Mix do EBIT por unidade",
                  f"IBI e SJP concentram {pc(ibi_sjp, 0)} do EBIT do MP'27; GOI ganha peso e RES passa a contribuir",
                  notes="Pizzas consideram as 4 unidades (R9: 12,8; MP'27: 11,5); R9 Brasil (13,0) inclui 0,2 da SOR. RES no R9 = −0,1, sem fatia.")
    dsz = 2300000
    xs = [LM + 150000, LM + 150000 + dsz + 800000]
    for x, ttl, vals, big, sub in ((xs[0], "Fcst R9 – 2026", r9, f1(t9), "EBIT 4 unidades"), (xs[1], "MP'27 – 2027", mp, f1(tm), "EBIT 4 unidades")):
        txt(s.shapes, x, 1380000, dsz, 280000, ttl, 13, True, NAVY, PP_ALIGN.CENTER)
        donut(s, x, 1680000, dsz, UNITS, vals, big, sub)
    xa = xs[0] + dsz + 80000
    line(s.shapes, xa, 2830000, xa + 640000, 2830000, GREY, 1.25, arrow=True, name='Seta')
    txt(s.shapes, xa - 80000, 2480000, 800000, 280000, pc(tm / t9 - 1, 1, True), 14, True, RED, PP_ALIGN.CENTER)
    txt(s.shapes, xa - 80000, 2900000, 800000, 220000, f"Δ {sg(tm - t9)} mi", 10, False, GREY, PP_ALIGN.CENTER)
    # painel
    xp = xs[1] + dsz + 380000
    fiat_share = FIAT_EBIT_SHARE
    side_panel(s, xp, 1380000, RM - xp, 2800000, "Leitura do mix", [
        (f"IBI: {pc(mp[0] / tm, 0)} do EBIT", f"Âncora: {pc(R('IBI', 'ns') / sum(R(u, 'ns') for u in UNITS), 0)} do Net Sales do Brasil"),
        (f"SJP: {pc(mp[1] / tm, 0)}", "Perde peso: desconto ao cliente e MO"),
        (f"GOI {pc(mp[2] / tm, 0)}  ·  RES {pc(mp[3] / tm, 0)}", "GOI estável em 1,4; RES passa a contribuir"),
        ("Concentração", f"Fiat = {pc(fiat_share, 0)} do EBIT do Brasil")], size=11, gap=9)
    # tabela
    rows = [["Unidade", "% Peças", "% Net Sales", "EBIT R9", "% R9", "EBIT MP'27", "% MP'27", "ROS MP'27", "Δ EBIT"]]
    for u in UNITS:
        rows.append([u, pc(R(u, 'pcs') / sum(R(x, 'pcs') for x in UNITS), 0), pc(R(u, 'ns') / sum(R(x, 'ns') for x in UNITS), 0),
                     par(R(u, 'ebit', 3)), pc(R(u, 'ebit', 3) / t9, 0), par(R(u, 'ebit')), pc(R(u, 'ebit') / tm, 0), pc(R(u, 'ros')),
                     par(R(u, 'ebit') - R(u, 'ebit', 3))])
    rows.append(["SOR (encerrada)", "–", "–", par(SOR), "–", "0,0", "–", "–", par(-SOR)])
    rows.append(["Brasil", "100%", "100%", par(e9), "", par(e27), "", pc(ros27), par(e27 - e9)])
    cw = [1900000, 1100000, 1250000, 1150000, 1000000, 1350000, 1100000, 1300000, 1100000]
    cw[-1] = W_ALL - sum(cw[:-1])

    def cs(i, j):
        if j == 0 and 0 < i <= 4:
            return {'color': UCOL[UNITS[i - 1]] if UNITS[i - 1] != 'RES' else EDGE, 'bold': True}
        if i == len(rows) - 1:
            return {'bold': True, 'fill': LIGHT}
        if j == 8 and i > 0:
            v = rows[i][8]
            return {'bold': True, 'color': RED if v.startswith('(') else GREEN}
        return {}
    table(s, LM, 4420000, cw, rows, 255000, 'Tabela mix', font=10.5, head_h=280000, cell_style=cs, aligns=[PP_ALIGN.LEFT] + [PP_ALIGN.RIGHT] * 8)
    footnote(s, "Fonte: Base Book Fitted Units MP27. Pizzas consideram as 4 unidades (R9: 12,8; MP'27: 11,5); o R9 do Brasil (13,0) inclui 0,2 de outras "
                "receitas da SOR. RES no R9 = −0,1, sem fatia na pizza.")
    return s


# ================================================================ clientes IBI
def slide_clientes():
    f_ros = C_ROS[0]
    s = new_slide("Fitted Units Economics MP'27 – IBI | Rentabilidade por cliente",
                  f"Fiat traz {pc(sh_ns[0], 0)} da receita e {pc(sh_eb[0], 0)} do EBIT de Ibirité, com ROS de {pc(f_ros)} (meta {pc(ROS_META)}); "
                  "Iveco e CNH têm margem alta em volume pequeno",
                  tag=None,
                  notes="Fonte: Rentabilidade por Cliente IBI MP27 v3 (21/10, ligada à MENS v4). Faturamento e material de consumo diretos; frete 100% Fiat; "
                        "demais custos rateados por volume; custos dedicados de CNH e Iveco realocados do total. Realocação de 0,3 mi da Fiat para Iveco/CNH para atingir a meta de ROS.")
    names = [('fiat', 'Fiat'), ('iveco', 'Iveco'), ('cnh', 'CNH')]
    cw_, cg_, cy_, ch_ = 3640000, 190000, 1430000, 2560000
    for k, (lg, nm) in enumerate(names):
        x = LM + k * (cw_ + cg_)
        card(s, x, cy_, cw_, ch_, 'Cliente ' + nm)
        logo(s, lg, x + 230000, cy_ + 130000, 900000, 560000)
        rich(s.shapes, x + 1300000, cy_ + 110000, cw_ - 1450000, 620000,
             [{'runs': [T('ROS ', 12, False, DGREY), T(pc(C_ROS[k]), 24, True, NAVY)], 'align': PP_ALIGN.RIGHT},
              {'runs': [T((f"meta {pc(ROS_META)}  ·  " if k == 0 else '') + f"{fp(C_PCS[k])}k pçs", 10, False, GREY)], 'align': PP_ALIGN.RIGHT}],
             MSO_ANCHOR.MIDDLE)
        for j, (lab, v, col) in enumerate((("% da receita IBI", sh_ns[k], EDGE), ("% do EBIT IBI", sh_eb[k], NAVY))):
            yy = cy_ + 900000 + j * 760000
            txt(s.shapes, x + 230000, yy, 2200000, 230000, lab, 10.5, False, DGREY)
            txt(s.shapes, x + cw_ - 1230000, yy - 20000, 1000000, 260000, pc(v, 0), 15, True, NAVY, PP_ALIGN.RIGHT)
            box(s.shapes, x + 230000, yy + 290000, cw_ - 460000, 150000, RGBColor(0xDD, 0xE1, 0xEA), 'Trilha')
            box(s.shapes, x + 230000, yy + 290000, max(int((cw_ - 460000) * v), 20000), 150000, col, 'Barra')
    # tabela
    yt = cy_ + ch_ + 170000
    rows = [["R$ mi", "Fiat", "Iveco", "CNH", "Total IBI"],
            ["Peças (mil)"] + [fp(v) for v in C_PCS],
            ["Net Sales"] + [f1(v / 1000) for v in C_NS],
            ["   NS / peça (R$)"] + [num(v, 2) for v in C_NSPC],
            ["Margem de contribuição"] + [f1(v / 1000) for v in C_CM],
            ["Custos fixos"] + [par(v / 1000) for v in C_FIX],
            ["Realocação p/ meta de ROS da Fiat"] + [par(v / 1000, 2) if abs(v) > 1e-9 else '–' for v in C_REAL],
            ["EBIT"] + [par(v / 1000) for v in C_EBIT],
            ["   ROS"] + [pc(v) for v in C_ROS]]
    tw = 7550000
    cw = [2750000] + [int((tw - 2750000) / 4)] * 4
    cw[-1] = tw - sum(cw[:-1])

    def cs(i, j):
        if rows[i][0].startswith('   ') and i > 0:
            return {'color': GREY, 'size': 9.5, 'italic': j == 0}
        if rows[i][0] == 'EBIT':
            return {'fill': LIGHT, 'bold': True}
        return {}
    table(s, LM, yt, cw, rows, 225000, 'Tabela clientes', font=10.5, cell_style=cs, head_h=270000)
    nz = list(CNH_ZERO)
    MES = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']
    xp = LM + tw + 230000
    side_panel(s, xp, yt - 30000, RM - xp, 6400000 - yt + 30000, "Leitura", [
        ("Fiat: escala do negócio", f"ROS {pc(f_ros)}, acima da meta de {pc(ROS_META)}"),
        ("Iveco: margem alta, volume menor", f"ROS {pc(C_ROS[1])}; volume {pc(IVECO_V27 / IVECO_V9 - 1, 1, True)} vs R9"),
        ("CNH: margem alta, volume irregular",
         f"ROS {pc(C_ROS[2])}, com custo dedicado" + (f"; {' e '.join(MES[m] for m in nz)} sem volume" if nz else '')),
        ("Cuidado", "Rateio por volume: validar Iveco/CNH antes de usar em preço")], size=10, gap=5)
    footnote(s, "Fonte: Rentabilidade por Cliente IBI MP27 v3 (ligada à MENS v4). Faturamento e material direto; frete 100% Fiat; demais custos por volume. "
                "Realocação de 0,3 mi (Fiat → Iveco/CNH) para a meta de ROS da Fiat.", w=7550000)
    return s


# ================================================================ mão de obra
def slide_mo():
    estrutura = sum(ANO_CHEIO.values())
    reaj = REAJ_SAL + REAJ_MED
    s = new_slide("Fitted Units Economics MP'27 – Mão de obra",
                  f"Mão de obra {sg(MO27 - MO9)} mi ({pc(MO27 / MO9 - 1, 1, True)}): {f1(estrutura)} é a estrutura aberta em 2026 em ano cheio "
                  f"e {f1(reaj)} o reajuste de 2027 ({pc(MO27 / MORR - 1, 1, True)})",
                  tag="(+) ganho  |  R$ milhões (custo direto)",
                  notes="Fonte: Labour Cost FITTED MP27 v2 e Fcst R9. Estrutura ano cheio = custo mensal set–dez/26 × 12 menos o R9. Reajuste 2027 = MP'27 menos esse custo anualizado.")
    deltas = [("IBI\n3º turno", ANO_CHEIO['IBI']), ("RES\nano cheio", ANO_CHEIO['RES']),
              ("Demais\nunidades", ANO_CHEIO['SJP'] + ANO_CHEIO['GOI'] + ANO_CHEIO['GER']),
              ("Reajuste sal.\n+ encargos", REAJ_SAL), ("Assist.\nmédica", REAJ_MED)]
    txt(s.shapes, LM, 1400000, 7500000, 260000, f"Aumento da mão de obra vs R9 (R$ {f1(MO9)} → {f1(MO27)} mi)  –  aumento de custo em vermelho", 12.5, True, NAVY)
    waterfall(s, LM - 60000, 1700000, 7600000, 2500000, None, None, deltas, f"Total\n{pc(MO27 / MO9 - 1, 1, True)}", MO27 - MO9,
              name='Bridge MO', head=330000, bad_up=True)
    rows = [["R$ mi", "Fcst R9", "Estrutura ano cheio", "Reajuste 2027", "MP'27", "Δ total", "Δ %"]]
    for u in UNITS + ['GER', 'ALL']:
        a, b_, c = mo(u, i=0), mo(u, i=1), mo(u)
        rows.append(['Total Fitted' if u == 'ALL' else ('Gerência' if u == 'GER' else u), f1(a), sg(b_ - a), sg(c - b_), f1(c),
                     sg(c - a), pc(c / a - 1, 1, True) if a > 1 else 'n.a.'])
    cw = [1400000, 950000, 1500000, 1200000, 950000, 800000, 800000]
    tw = sum(cw)
    table(s, LM, 4300000, cw, rows, 255000, 'Tabela MO', font=10.5, bold_rows=(len(rows) - 1,),
          fill_rows={len(rows) - 1: LIGHT}, head_h=300000)
    xp = LM + tw + 250000
    side_panel(s, xp, 1430000, RM - xp, 6380000 - 1430000, "Leitura", [
        (f"Estrutura 2026 em ano cheio {sg(estrutura)}", [f"IBI {sg(ANO_CHEIO['IBI'])}: 3º turno iniciado em 2026", f"RES {sg(ANO_CHEIO['RES'])}: unidade aberta em 2026"]),
        (f"Reajuste 2027 {sg(reaj)} ({pc(MO27 / MORR - 1, 1, True)})",
         [f"Salários e encargos {pc(sal('ALL', 2) / sal('ALL', 1) - 1, 1, True)}", f"Assistência médica {pc(med('ALL', 2) / med('ALL', 1) - 1, 1, True)}"]),
        ("Riscos a validar com o RH", ["Férias do horista em ~5,5% (estatutário 11,1%): até −0,95",
                                       "SJP CC 8289: horista fixo +35% sem premissa visível (−0,46)",
                                       "GOI: assistência médica 3,4% do salário vs 31% em IBI (até −0,36)"]),
        ("Sem a estrutura nova", f"A MO sobe {pc(MO27 / MORR - 1, 1)} vs run-rate de set–dez/26")], size=10.5, gap=7)
    footnote(s, "Fonte: Labour Cost FITTED MP27 v2 e Fcst R9. Estrutura ano cheio = custo mensal de set–dez/26 × 12 menos o R9. "
                "Reajuste 2027 = MP'27 menos esse custo anualizado.", w=7550000)
    return s


# ================================================================ new business
def slide_nb():
    s = new_slide("Fitted Units Economics MP'27 – New Business",
                  f"New business soma {fp(NB_PCS)}k peças ({pc(NB_PCS / pcs27, 0)} do volume): Nissan em RES, Tukan em SJP e CNH em IBI",
                  tag=None, notes="Fonte: Volumes MP'27 (PO's\\Volume\\MP'27_ADJ), Base Book Fitted Units MP27 e Rentabilidade por Cliente IBI v3.")
    NB = [('nissan', "Nissan – Resende", "Kait e Kicks · 1º ano cheio", f"{fp(NISSAN)}k pçs", f"{pc(NISSAN / R('RES', 'pcs', 3) - 1, 0, True)} vs R9",
           [f"Net Sales R$ {f1(R('RES', 'ns'))} mi", f"EBIT R$ {f1(R('RES', 'ebit'))} mi · ROS {pc(R('RES', 'ros'), 0)}"]),
          ('vw', "VW – São José dos Pinhais", "Lançamento do Tukan", f"{fp(TUKAN)}k pçs", f"{pc(TUKAN / R('SJP', 'pcs'), 0)} do volume de SJP",
           [f"Substitui o Virtus ({fp(VIRTUS)}k pçs, até fev/27)", f"Volume de SJP estável: {fp(R('SJP', 'pcs'))}k pçs"]),
          ('cnh', "CNH – Ibirité", "Cliente iniciado em 2026", f"{fp(C_PCS[2])}k pçs", f"{pc(JLR['CNH'][1] / JLR['CNH'][0] - 1, 0, True)} vs R9",
           [f"Net Sales R$ {f1(C_NS[2] / 1000)} mi", f"EBIT R$ {f1(C_EBIT[2] / 1000)} mi · ROS {pc(C_ROS[2], 0)}"])]
    W3, G3, Y3, H3 = 3640000, 190000, 1430000, 2800000
    for k, (lg, tit, sub, big, bsub, lines) in enumerate(NB):
        x = LM + k * (W3 + G3)
        card(s, x, Y3, W3, H3, f'Card NB {k + 1}')
        logo(s, lg, x + 230000, Y3 + 170000, 700000, 520000)
        rich(s.shapes, x + 1100000, Y3 + 140000, W3 - 1250000, 600000,
             [[T(tit, 13.5, True, HEAD)], [T(sub, 10.5, False, DGREY)]], MSO_ANCHOR.MIDDLE)
        rich(s.shapes, x + 230000, Y3 + 930000, W3 - 460000, 780000,
             [[T(big, 30, True, NAVY)], [T(bsub, 11.5, True, GREEN)]], MSO_ANCHOR.TOP)
        line(s.shapes, x + 230000, Y3 + 1850000, x + W3 - 230000, Y3 + 1850000, RULE, 0.75)
        rich(s.shapes, x + 230000, Y3 + 2000000, W3 - 460000, 900000,
             [{'runs': [T(t, 12, False, BLACK)], 'bullet': True, 'space_after': 6} for t in lines], MSO_ANCHOR.TOP)
    YS = Y3 + H3 + 200000
    txt(s.shapes, LM, YS, 6000000, 260000, "Saídas do perímetro", 13, True, NAVY)
    assert jl.cell(24, 3).value == 'JLR'
    jag_fat = float(jl.cell(24, 8).value) / 1000          # faturamento Jaguar no R9 (aba Impacto JLR, R$ mi)
    SAI = [("Jaguar – Ibirité", f"Encerrado em 2026: EBIT {f1(DD('IBI', 'Encerr. Jaguar'))} e R$ {f1(jag_fat)} mi de receita"),
           ("Toyota – Sorocaba", f"Unidade desmobilizada: R$ {f1(SOR)} mi de outras receitas no R9")]
    sw = (W_ALL - 190000) // 2
    for k, (tit, t) in enumerate(SAI):
        x = LM + k * (sw + 190000)
        card(s, x, YS + 320000, sw, 950000, f'Card saida {k + 1}', RED)
        rich(s.shapes, x + 250000, YS + 320000, sw - 400000, 950000, [[T(tit, 13, True, HEAD)], [T(t, 11.5, False, BLACK)]], MSO_ANCHOR.MIDDLE)
    footnote(s, "Fonte: Volumes MP'27 (PO's\\Volume\\MP'27_ADJ), Base Book Fitted Units MP27 e Rentabilidade por Cliente IBI v3. "
                "% vs R9 da Nissan reflete a abertura de RES ao longo de 2026.", w=9700000)
    return s


# ================================================================ riscos e oportunidades
def slide_ro():
    RISKS = [("Reajuste Fiat não aprovado", f"Preço R$ {num(FIAT_P27, 2)}/pç ({pc(FIAT_PCT, 1, True)} vs R9) ainda em negociação", -FIAT_REAJ, ''),
             ("Provisão de férias do horista", "Premissa ~5,5% vs 11,1% do estatutário", -0.95, 'até '),
             ("Horista fixo SJP (CC 8289)", "+35% sem premissa visível", -0.46, ''),
             ("Assistência médica GOI", "3,4% do salário vs 31% em IBI", -0.36, 'até '),
             ("Volume", "Cada 1% a menos de peças (margem variável)", -SENS_VOL, 'sens. ')]
    OPPS = [("Repasse da ANTT ao cliente", "Frete mínimo absorvido no plano", ANTT, ''),
            ("Preço Fiat acima do plano", "Cada +R$ 0,10/pç no preço", SENS_PRECO, 'sens. '),
            ("Volume", "Cada 1% a mais de peças (margem variável)", SENS_VOL, 'sens. ')]
    tot_r = sum(v for *_, v, pre in RISKS if not pre.startswith('sens'))
    tot_o = sum(v for *_, v, pre in OPPS if not pre.startswith('sens'))
    net = tot_r + tot_o
    s = new_slide("Fitted Units Riscos e Oportunidades MP'27",
                  f"Riscos quantificados de {f1(tot_r, 2)} mi, concentrados no reajuste Fiat e na MO; oportunidades de {sg(tot_o, 2)} mi",
                  tag="(+) ganho  |  EBIT, R$ milhões",
                  notes="Fonte: Base Book e MENS FITTED MP27 v4 (reajuste Fiat, ANTT, sensibilidades); Auditoria Labour FITTED MP27 (riscos de MO, hipóteses a confirmar com o RH). "
                        "Itens de sensibilidade (volume e preço) não entram nos totais.")
    CW2, GAP, YH = (W_ALL - 250000) // 2, 250000, 1430000
    ih, ig = 640000, 60000
    for k, (ttl, col, items, tot) in enumerate((("Riscos", RED, RISKS, tot_r), ("Oportunidades", GREEN, OPPS, tot_o))):
        x = LM + k * (CW2 + GAP)
        hb = box(s.shapes, x, YH, CW2, 430000, col, 'Cabecalho ' + ttl)
        rich(s.shapes, 0, 0, 0, 0, [{'runs': [T(ttl, 15, True, WHITE)]}], MSO_ANCHOR.MIDDLE, shape=hb, margins=(180000, 0, 0, 0))
        txt(s.shapes, x + CW2 - 1900000, YH, 1750000, 430000, "EBIT (R$ mi)", 10.5, False, WHITE, PP_ALIGN.RIGHT)
        for i, (t1, t2, v, pre) in enumerate(items):
            yy = YH + 500000 + i * (ih + ig)
            sens = pre.startswith('sens')
            card(s, x, yy, CW2, ih, f'{ttl} item {i + 1}', GREY if sens else col)
            rich(s.shapes, x + 200000, yy, CW2 - 1900000, ih, [[T(t1, 12, True, HEAD)], [T(t2, 10.5, False, DGREY)]], MSO_ANCHOR.MIDDLE)
            lab = pre + (sg(v, 2) if v > 0 else f1(v, 2))
            txt(s.shapes, x + CW2 - 1800000, yy, 1650000, ih, lab, 12 if sens else 17, True, GREY if sens else col, PP_ALIGN.RIGHT)
        yy = YH + 500000 + len(items) * (ih + ig) + 10000
        line(s.shapes, x, yy, x + CW2, yy, GREY, 1)
        txt(s.shapes, x + 100000, yy + 40000, CW2 - 1800000, 300000, 'Total quantificado (sem sensibilidades)', 11.5, True, BLACK)
        txt(s.shapes, x + CW2 - 1800000, yy + 40000, 1650000, 300000, f1(tot, 2) if k == 0 else sg(tot, 2), 17, True, col, PP_ALIGN.RIGHT)
    # exposição líquida (coluna de oportunidades, abaixo dos cards)
    xo = LM + CW2 + GAP
    yn = YH + 500000 + len(OPPS) * (ih + ig) + 430000
    hn = YH + 500000 + len(RISKS) * (ih + ig) - yn
    box(s.shapes, xo, yn, CW2, hn, NAVY, 'Exposicao liquida')
    rich(s.shapes, xo + 250000, yn, CW2 - 500000, hn,
         [[T('Exposição líquida (sem sensibilidades)', 11.5, True, SKY)],
          [T(f"{f1(net, 2)} mi", 28, True, WHITE)],
          [T(f"EBIT de R$ {f1(e27 + tot_r)} a {f1(e27 + tot_o)} mi (plano: R$ {f1(e27)} mi)", 11.5, False, WHITE)]], MSO_ANCHOR.MIDDLE)
    footnote(s, "Fonte: Base Book e MENS FITTED MP27 v4 (reajuste Fiat, ANTT, sensibilidades); Auditoria Labour FITTED MP27 (riscos de MO, "
                "hipóteses a confirmar com o RH). Sensibilidades de volume e preço não entram nos totais.", w=9700000)
    return s


# ================================================================ decisões
def slide_decisoes():
    tot_r = -(FIAT_REAJ + RISCO_MO)
    s = new_slide("Fitted Units Economics MP'27 – Decisões e próximos passos",
                  f"Quatro decisões protegem o EBIT de R$ {f1(e27)} mi: preço, ANTT, mão de obra e rentabilidade por cliente",
                  tag=None)
    items = [("1", "Fechar o reajuste Fiat", f"Preço de R$ {num(FIAT_P27, 2)}/pç ({pc(FIAT_PCT, 1, True)} vs R9) vale {sg(FIAT_REAJ)} mi de EBIT. "
                                              f"Sem ele, o EBIT do Brasil cai para ~{f1(e27 - FIAT_REAJ)}.", "Comercial / Fiat"),
             ("2", "Repassar a ANTT ao cliente", f"Nova regra de frete mínimo gera custo de {f1(ANTT)} mi já absorvido no plano: o repasse é upside direto.", "Comercial / Logística"),
             ("3", "Validar riscos de mão de obra", f"Provisão de férias do horista, horista fixo SJP e assistência médica GOI somam até {f1(RISCO_MO, 1)} mi.", "RH / Controladoria"),
             ("4", "Rentabilidade por cliente em IBI", f"Fiat ROS {pc(C_ROS[0])} (meta {pc(ROS_META)}); Iveco e CNH com margem alta em base rateada: validar custos específicos antes de usar em preço.", "Controladoria / Custos")]
    cw, gap, y0, h = (W_ALL - 190000) // 2, 190000, 1430000, 1500000
    for i, (n_, t1, t2, who) in enumerate(items):
        x = LM + (i % 2) * (cw + gap)
        y = y0 + (i // 2) * (h + 130000)
        card(s, x, y, cw, h, f'Decisao {n_}')
        c = box(s.shapes, x + 230000, y + 200000, 500000, 500000, NAVY, 'Numero', MSO_SHAPE.OVAL)
        rich(s.shapes, 0, 0, 0, 0, [{'runs': [T(n_, 20, True, WHITE)], 'align': PP_ALIGN.CENTER}], MSO_ANCHOR.MIDDLE, shape=c)
        rich(s.shapes, x + 900000, y + 130000, cw - 1080000, h - 200000,
             [{'runs': [T(t1, 15, True, HEAD)], 'space_after': 5}, {'runs': [T(t2, 12, False, DGREY)], 'space_after': 6},
              {'runs': [T('Responsável: ', 10.5, True, GREY), T(who, 10.5, False, GREY)]}], MSO_ANCHOR.TOP)
    # cenários de EBIT
    ys = y0 + 2 * h + 130000 + 230000
    txt(s.shapes, LM, ys, 6000000, 260000, "Faixa de EBIT do Brasil – R$ mi", 13, True, NAVY)
    sc = [("Riscos materializados", e27 + tot_r, f"{f1(tot_r, 2)} vs plano: sem reajuste Fiat e com riscos de MO", RED),
          ("Plano MP'27", e27, f"ROS {pc(ros27)}  ·  {sg(e27 - e9)} vs R9", NAVY),
          ("Com repasse da ANTT", e27 + ANTT, f"{sg(ANTT)} vs plano: ANTT repassada ao cliente", GREEN)]
    sw = (W_ALL - 2 * 190000) // 3
    for k, (a, v, d, col) in enumerate(sc):
        x = LM + k * (sw + 190000)
        card(s, x, ys + 300000, sw, 1050000, 'Cenario ' + a, col)
        rich(s.shapes, x + 200000, ys + 300000, sw - 300000, 1050000,
             [[T(a.upper(), 9.5, True, GREY)], [T(f"R$ {f1(v)} mi", 24, True, col)], [T(d, 10, False, DGREY)]], MSO_ANCHOR.MIDDLE)
    footnote(s, "Fonte: Base Book Fitted Units MP27, MENS FITTED MP27 v4, Rentabilidade por Cliente IBI v3, Auditoria Labour FITTED MP27.")
    return s


# ================================================================ montagem
def main():
    slide_resumo()
    slide_brasil()
    slide_mix()
    # IBI
    ib = BRG['IBI']
    unit_slide('IBI', "Fitted Units Economics MP'27 – Ibirité (IBI)",
               f"IBI entrega EBIT de R$ {f1(R('IBI', 'ebit'))} mi (ROS {pc(R('IBI', 'ros'))}): a saída da Jaguar, a MO e a ANTT são cobertas por preço e fixos",
               [(f"EBIT {f1(R('IBI', 'ebit'))}  |  ROS {pc(R('IBI', 'ros'))}",
                 f"{sg(R('IBI', 'ebit') - R('IBI', 'ebit', 3))} vs R9 ({pc(R('IBI', 'ebit') / R('IBI', 'ebit', 3) - 1, 1, True)})  ·  {sg(R('IBI', 'ebit') - R('IBI', 'ebit', 2))} vs MP'26"),
                (f"Saída da Jaguar {f1(DD('IBI', 'Encerr. Jaguar'))}", "R$ 1,5 mi de receita saem do perímetro"),
                ("3º turno Fiat e ANTT", [f"Mão de obra {f1(DD('IBI', 'Labour'))}: 3º turno (MO de 2,5 mi/ano)",
                                          f"ANTT {f1(DD('IBI', 'ANTT'))}: nova regra de frete mínimo"]),
                (f"Preço Fiat R$ {num(FIAT_P27, 2)}/pç", f"{pc(FIAT_PCT, 1, True)} vs R9 = {sg(FIAT_REAJ)} mi; fixos {sg(DD('IBI', 'Cost Fix.'))} com a base fixa revisada"),
                ("Ponto de atenção", f"Sem o reajuste Fiat, o EBIT de IBI cai para ~{f1(R('IBI', 'ebit') - FIAT_REAJ)}")],
               f"Bridge IBI (Base Book): {b_str('IBI')}")
    # SJP
    sj = BRG['SJP']
    unit_slide('SJP', "Fitted Units Economics MP'27 – São José dos Pinhais (SJP)",
               f"SJP mantém ROS de {pc(R('SJP', 'ros'))}, mas perde {f1(R('SJP', 'ebit', 3) - R('SJP', 'ebit'))} mi vs R9 com desconto ao cliente e mão de obra",
               [(f"EBIT {f1(R('SJP', 'ebit'))}  |  ROS {pc(R('SJP', 'ros'))}",
                 f"{sg(R('SJP', 'ebit') - R('SJP', 'ebit', 3))} vs R9 ({pc(R('SJP', 'ebit') / R('SJP', 'ebit', 3) - 1, 1, True)})  ·  {sg(R('SJP', 'ebit') - R('SJP', 'ebit', 2))} vs MP'26"),
                ("Desconto de −3,5% ao cliente", ["R$ 23,92 → 23,05/pç: ≈ −0,7 mi/ano", f"Parte compensada: Price/Mix só {f1(DD('SJP', 'Price/ Mix'))} vs R9"]),
                (f"Mão de obra {f1(DD('SJP', 'Labour'))} (+{num((mo('SJP') / mo('SJP', i=0) - 1) * 100)}%)", ["86% do aumento: salários e encargos", "14%: assistência médica"]),
                ("Volume e demais custos estáveis", f"{fp(R('SJP', 'pcs'))}k pçs, em linha com o R9")],
               f"Bridge SJP (Base Book): {b_str('SJP')}")
    for u, nm in (('GOI', 'Goiana (GOI)'), ('RES', 'Resende (RES)')):
        if u == 'GOI':
            hl = f"GOI mantém o EBIT de R$ {f1(R('GOI', 'ebit'))} mi: preço e menores fixos compensam volume e custos"
            blocks = [(f"EBIT {f1(R('GOI', 'ebit'))}  |  ROS {pc(R('GOI', 'ros'))}",
                       f"{sg(R('GOI', 'ebit') - R('GOI', 'ebit', 3))} vs R9 ({pc(R('GOI', 'ebit') / R('GOI', 'ebit', 3) - 1, 1, True)})  ·  {sg(R('GOI', 'ebit') - R('GOI', 'ebit', 2))} vs MP'26"),
                      ("Preço compensa o volume", [f"Preço por peça {pc(R('GOI', 'nspc') / R('GOI', 'nspc', 3) - 1, 1, True)}: Price/Mix {sg(DD('GOI', 'Price/ Mix'))}",
                                                   f"Volume {pc(R('GOI', 'pcs') / R('GOI', 'pcs', 3) - 1, 1, True)}: {sg(DD('GOI', 'Volume'))}"]),
                      ("Custos fixos e variáveis equalizados", [f"Fixos {sg(DD('GOI', 'Cost Fix.'))} vs R9",
                                                                f"Compensam MO {f1(DD('GOI', 'Labour'))} e variáveis {f1(DD('GOI', 'Cost Var.'))}"])]
        else:
            hl = f"RES passa a contribuir: EBIT de R$ {f1(R('RES', 'ebit'))} mi e ROS de {pc(R('RES', 'ros'), 0)} já no primeiro ano cheio"
            blocks = [(f"EBIT {f1(R('RES', 'ebit'))}  |  ROS {pc(R('RES', 'ros'))}",
                       f"{sg(R('RES', 'ebit') - R('RES', 'ebit', 3))} vs R9  ·  {sg(R('RES', 'ebit') - R('RES', 'ebit', 2))} vs MP'26"),
                      ("Delta de perímetro", ["Unidade não operou o ano todo em 2026", f"{fp(R('RES', 'pcs'))}k pçs no MP'27 vs {fp(R('RES', 'pcs', 3))}k no R9"]),
                      ("ROS de 10% já no 1º ano cheio", [f"Preço R$ {num(R('RES', 'nspc'), 2)}/pç (R9: {num(R('RES', 'nspc', 3), 2)})",
                                                         f"Cobre MO {f1(DD('RES', 'Labour'))} e fixos {f1(DD('RES', 'Cost Fix.'))}"])]
        unit_slide(u, f"Fitted Units Economics MP'27 – {nm}", hl, blocks, f"Bridge {u} (Base Book): {b_str(u)}")
    slide_clientes()
    slide_mo()
    slide_nb()
    slide_ro()
    slide_decisoes()
    prs.save(OUT)
    print('salvo', OUT, len(prs.slides), 'slides')


main()
