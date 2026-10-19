"""PPT MP27 Fitted Units - revisao 2 (pos Base Book de 19/10, sem target).

Uso: python gerar_ppt_mp27_rev2.py <deck_base.pptx> <basebook.xlsx> <mens.xlsx> <rentab_cliente_ibi.xlsx>
                                   <variacao_mo.xlsx> <detalhe_despesas.xlsx> <saida.pptx>

Parte do deck aprovado (v5 renomeado pela Juliana) e:
- refaz as tabelas de P&L como tabelas nativas a partir do Base Book (abas Month_FY_LC / unidades, AK38:AT60);
- refaz os bridges (abas 'Tabela Variantes s targ' / 'CM s targ');
- reescreve os comentarios (MENS v3, rentabilidade por cliente v3, variacao da MO, Detalhe de Despesas);
- recria Mix do EBIT e Rentabilidade por cliente IBI; adiciona Resumo executivo, Mao de obra e Perguntas-chave;
- remove a secao duplicada (o Base Book novo nao tem mais versao 'com target').
"""
import copy
import sys
from decimal import Decimal, ROUND_HALF_UP

import openpyxl
from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

SRC, BASE, MENS, CLI, MO, DET, OUT = sys.argv[1:8]

DARK = RGBColor(0x36, 0x4D, 0x6E)
LIGHT = RGBColor(0xDF, 0xE5, 0xEF)
ACC = RGBColor(0x9B, 0xC2, 0xE6)
GREY = RGBColor(0x80, 0x80, 0x80)
DGREY = RGBColor(0x40, 0x40, 0x40)
BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)
HEAD = RGBColor(0x1F, 0x4E, 0x79)
NAVY = RGBColor(0x1F, 0x38, 0x64)
PANEL = RGBColor(0xF2, 0xF2, 0xF2)
GREEN = RGBColor(0x3A, 0x9A, 0x5B)
RED = RGBColor(0xC0, 0x39, 0x2B)
YEL = RGBColor(0xFF, 0xDD, 0x00)
UCOL = {'IBI': DARK, 'SJP': RGBColor(0x5B, 0x8D, 0xB8), 'GOI': ACC, 'RES': YEL}
CCOL = {'Fiat': DARK, 'Iveco': ACC, 'CNH': YEL}


# ================================================================ formatos
def _q(x, d):
    return Decimal(str(abs(x))).quantize(Decimal(1).scaleb(-d), ROUND_HALF_UP)


def f1(x, d=1):
    """-1,3 / 0,0 (bridge e texto)."""
    v = _q(x, d)
    s = f"{v:,.{d}f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    return s if (x >= 0 or v == 0) else f"−{s}"


def sg(x, d=1):
    """com sinal: +1,1 / −2,0."""
    t = f1(x, d)
    return t if t.startswith('−') else f"+{t}"


def fp(x, d=1):
    """tabela: negativos entre parenteses."""
    v = _q(x, d)
    t = f"{v:,.{d}f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    return f"({t})" if x < 0 and v != 0 else t


def pc(x, d=1, signed=False):
    v = _q(x * 100, d)
    t = f"{v:.{d}f}".replace('.', ',')
    if x < 0 and v != 0:
        return f"−{t}%"
    return f"+{t}%" if signed and v != 0 else f"{t}%"


# ================================================================ dados
bb = openpyxl.load_workbook(BASE, data_only=True)
mens = openpyxl.load_workbook(MENS, data_only=True)

BR_SRC = {'BR': ('Tabela Variantes s targ', 2, 4), 'IBI': ('Tabela Variantes s targ', 6, 8),
          'SJP': ('Tabela Variantes CM s targ', 3, 5), 'GOI': ('Tabela Variantes CM s targ', 8, 10),
          'RES': ('Tabela Variantes CM s targ', 13, 15)}


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
        items.append((h, float(v)))
    start, deltas = items[0][1], items[1:]
    assert abs(start + sum(v for _, v in deltas) - end) < 0.003, (sheet, dr)
    return {'start': start, 'deltas': deltas, 'end': end, 'd': {k.replace('\n', ' '): v for k, v in deltas}}


BRG = {u: load_bridge(*s) for u, s in BR_SRC.items()}

PL_SHEET = {'BR': 'Month_FY_LC - MP27', 'IBI': 'IBIRITE', 'SJP': 'SJP', 'GOI': 'GOIANA', 'RES': 'RESENDE'}
PL_ROWS = [('pcs', 42), ('ns', 44), ('nspc', 45), ('cm', 47), ('cmp', 48), ('fix', 51), ('oi', 53),
           ('ebit', 55), ('ebpc', 56), ('ros', 58)]
COLS = [39, 40, 41, 42, 43]          # AM..AQ: Act24, Act25, MP26, R9, MP27
DCOLS = [45, 46]                      # AS, AT: MP27 vs R9, MP27 vs MP26


def load_pl(u):
    ws = bb[PL_SHEET[u]]
    assert ws.cell(39, 43).value == "MP'27" and ws.cell(39, 39).value == 'Actual 2024', u
    d = {}
    for k, r in PL_ROWS:
        d[k] = [float(ws.cell(r, c).value or 0) for c in COLS]
        d['d_' + k] = [float(ws.cell(r, c).value or 0) for c in DCOLS]
    d['debit'] = [float(ws.cell(60, c).value or 0) for c in (40, 41, 42, 43)]
    return d


PL = {u: load_pl(u) for u in PL_SHEET}
for u in BRG:   # bridge x tabela
    assert abs(BRG[u]['end'] - PL[u]['ebit'][4]) < 0.002 and abs(BRG[u]['start'] - PL[u]['ebit'][3]) < 0.002, u
UNITS = ['IBI', 'SJP', 'GOI', 'RES']
assert abs(sum(PL[u]['ebit'][4] for u in UNITS) - PL['BR']['ebit'][4]) < 0.002

R = lambda u, k, i=4: PL[u][k][i]          # i: 3 = R9, 4 = MP27, 2 = MP26

# MENS (R$ mil)
mt = mens['TOTAL']


def mrow(ws, label, col):
    for r in range(1, 60):
        if str(ws.cell(r, 3).value or '').strip() == label:   # rotulos na coluna C
            return float(ws.cell(r, col).value or 0)
    raise KeyError(label)


TOT, R9C = 17, 19   # colunas Q (TOTAL MP'27) e S (R9) da MENS
assert abs(mrow(mt, 'EBIT', TOT) / 1000 - R('BR', 'ebit')) < 0.002
SENS = (mrow(mt, 'NET SALES', TOT) + mrow(mt, 'Direct Materials', TOT) + mrow(mt, 'Transportation', TOT)
        + mrow(mt, 'Other Variable', TOT)) / 1000 / 100          # mi de EBIT por 1% de volume
GOI_FRETE_MES = -mrow(mens['GO'], 'Transportation', TOT) / 12
GOI_FRETE_D = (mrow(mens['GO'], 'Transportation', TOT) - mrow(mens['GO'], 'Transportation', R9C)) / 1000
GOI_OFIX = (-mrow(mens['GO'], 'Other Fixed', TOT) / 1000, -mrow(mens['GO'], 'Other Fixed', R9C) / 1000)
ROS_REF = float(mt['Z44'].value)                     # referencia de ROS da MENS (EBIT ESPERADO = NS x 13,8%)
EBIT_REF = float(mt['Z43'].value) / 1000
assert str(mt['Y43'].value).strip() == 'EBIT ESPERADO'

jl = mens['Impacto JLR']
JLR = {}
for r in range(1, 60):
    k = jl.cell(r, 3).value   # rotulos na coluna C, R9 na D
    if k in ('Fiat', 'Iveco', 'CNH', 'JLR') and k not in JLR:
        JLR[k] = [jl.cell(r, c).value for c in range(4, 14)]   # vol R9, vol 27, pr R9, pr 27, fat R9, fat 27, dFat, vol, mix, preco
    if k == 'Margem direta da JLR (R$ mil)':
        JLR['margem'] = jl.cell(r, 4).value
    if k == 'Faturamento JLR (R$ mil)':
        JLR['fat'] = jl.cell(r, 4).value
    if k == 'Volume JLR (peças)':
        JLR['vol'] = jl.cell(r, 4).value
FIAT_PR9, FIAT_PR27, FIAT_PRECO = JLR['Fiat'][2], JLR['Fiat'][3], JLR['Fiat'][9] / 1000
IVECO_V9, IVECO_V27, IVECO_DF = JLR['Iveco'][0], JLR['Iveco'][1], JLR['Iveco'][6] / 1000
REDUCOES_IBI = sum(float(mens['Ajustes IBI'].cell(r, 4).value or 0) for r in range(4, 9)) / 1000

SJP_PRECO_ATUAL = float(bb['SJP'].cell(45, 53).value)          # BA45 = Actual 2026 (Set YTD) NS/pc
SJP_PRECO_27 = R('SJP', 'nspc')
SJP_DESC = SJP_PRECO_27 / SJP_PRECO_ATUAL - 1
SJP_DESC_MI = (SJP_PRECO_27 - SJP_PRECO_ATUAL) * R('SJP', 'pcs') / 1000

# Detalhe de Despesas - servicos do 3o turno IBI (aba Comparativo, coluna Delta vs R9)
det = openpyxl.load_workbook(DET, data_only=True, read_only=True)
T3 = {}
for row in det['Comparativo'].iter_rows(values_only=True):
    vals = [v for v in row if v not in (None, '')]
    if len(vals) == 15 and vals[0] in ('Limpeza', 'Locação', 'Alimentação', 'Transporte'):
        T3[vals[0]] = vals[-1] / 1e6
T3_TOT = sum(T3.values())
assert len(T3) == 4, T3

# Rentabilidade por cliente IBI (v3)
wc = openpyxl.load_workbook(CLI, data_only=True)['Rentabilidade por Cliente']
CL = ['Fiat', 'Iveco', 'CNH']
crow = lambda r: [float(wc.cell(r, c).value) for c in (2, 3, 4, 5)]
C_PCS, C_NS, C_NSPC, C_CM, C_FIX, C_EBIT, C_ROS, C_EBPC = crow(5), crow(6), crow(7), crow(14), crow(16), crow(24), crow(25), crow(26)
assert abs(C_EBIT[3] / 1000 - R('IBI', 'ebit')) < 0.002, (C_EBIT[3], R('IBI', 'ebit'))
wcp = openpyxl.load_workbook(CLI, data_only=True)['Premissas']
CNH_ESPEC = (float(wcp['D19'].value) + float(wcp['D20'].value)) / 1000
ROS_META = float(wcp['B23'].value)
wce = openpyxl.load_workbook(CLI, data_only=True)['EBIT Mensal']
CNH_MES = [float(wce.cell(4, c).value) for c in range(2, 14)]
wcd = openpyxl.load_workbook(CLI, data_only=True)['Dados IBI']
assert wcd['A4'].value.startswith('Volume') and wcd['B4'].value == 'CNH'
CNH_VOL = [float(wcd.cell(4, c).value or 0) for c in range(3, 15)]

# Mao de obra (R$ mil): {unidade: {grupo: (R9, run-rate, MP27)}}
wm = openpyxl.load_workbook(MO, data_only=True)['Resumo (4 grupos)']
MOD = {}
for r in range(5, 40):
    u, g = wm.cell(r, 1).value, wm.cell(r, 2).value
    if u in ('SJP', 'IBI', 'GOI', 'RES', 'GER') and g and g != 'Total MO':
        MOD.setdefault(u, {})[g] = tuple(float(wm.cell(r, c).value) for c in (3, 4, 5))
GR = ['Salarios e PPR', 'Encargos', 'Assistencia medica', 'Outros beneficios']


def mo(u, g=None, i=2):
    us = [u] if u != 'ALL' else list(MOD)
    gs = [g] if g else GR
    return sum(MOD[x][y][i] for x in us for y in gs) / 1000


assert abs((mo('ALL') - mo('ALL', i=0)) + BRG['BR']['d']['Labour']) < 0.05

# ================================================================ helpers de desenho
prs = Presentation(SRC)
slides = list(prs.slides)
LAYOUT = slides[1].slide_layout
FLAG = [s for s in slides[1].shapes if s.name == 'Picture 163'][0]


def txt(shapes, x, y, w, h, text, size, bold=False, color=BLACK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM,
        name=None, italic=False, wrap=False):
    tb = shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    for i, ln in enumerate(text.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
    if name:
        tb.name = name
    return tb


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


def box(shapes, x, y, w, h, fill, name, shape=MSO_SHAPE.RECTANGLE, line_rgb=None):
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
        b.line.width = Pt(0.75)
    b.shadow.inherit = False
    b.name = name
    return b


def set_bullet(p, char='•', indent=150000):
    pPr = p._p.get_or_add_pPr()
    pPr.set('marL', str(indent))
    pPr.set('indent', str(-indent))
    for t in ('a:buNone', 'a:buChar', 'a:buAutoNum'):
        for e in pPr.findall(qn(t)):
            pPr.remove(e)
    bf = pPr.makeelement(qn('a:buFont'), {'typeface': 'Arial'})
    bc = pPr.makeelement(qn('a:buChar'), {'char': char})
    pPr.append(bf)
    pPr.append(bc)


def no_bullet(p):
    pPr = p._p.get_or_add_pPr()
    pPr.append(pPr.makeelement(qn('a:buNone'), {}))


# ---------------------------------------------------------------- bridge (formas nativas, mesmo padrao da v5)
X0, X1 = 300000, 8450000
YT, YB = 4950000, 6420000
Y_BR = 4480000
LBL_D, LBL_T = 150000, 210000


def draw_bridge(slide, spec):
    start, deltas, end = spec['start'], spec['deltas'], spec['end']
    cols = [('Fcst R9', 'tot', start)] + [(lab, 'd', v) for lab, v in deltas] + [("MP’27", 'tot', end)]
    n = len(cols)
    slot = (X1 - X0) / n
    bw = int(slot * 0.62)
    bars, cum = [], 0.0
    for lab, kind, v in cols:
        if kind == 'tot':
            lo, hi, cum = min(0, v), max(0, v), v
        else:
            lo, hi = min(cum, cum + v), max(cum, cum + v)
            cum += v
        bars.append((lo, hi, cum))
    vmax = max(max(b[1] for b in bars), 0)
    vmin = min(min(b[0] for b in bars), 0)
    yb_eff = YB - (200000 if vmin < 0 else 0)
    sc = (yb_eff - YT) / (vmax - vmin)
    y = lambda v: int(YT + (vmax - v) * sc)
    sh = slide.shapes
    cx = lambda i: int(X0 + slot * (i + 0.5))
    new = [line(sh, X0, y(0), X1, y(0), GREY, 0.75, name='Eixo')]
    for i, (lab, kind, v) in enumerate(cols):
        lo, hi, lvl = bars[i]
        x = cx(i) - bw // 2
        top, h = y(hi), max(y(lo) - y(hi), 12700)
        r = sh.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(top), Emu(bw), Emu(h))
        r.fill.solid()
        r.fill.fore_color.rgb = DARK if kind == 'tot' else (LIGHT if v < 0 else ACC)
        r.line.color.rgb = BLACK
        r.line.width = Pt(0.75)
        r.shadow.inherit = False
        r.name = f'Barra {lab.replace(chr(10), " ")}'
        new.append(r)
        if kind == 'tot':
            below = v < 0
            ly = (y(0) + 25000) if below else (top - 25000 - LBL_T)
            new.append(txt(sh, cx(i) - int(slot / 2), ly, int(slot), LBL_T, f1(v), 14, True,
                           anchor=MSO_ANCHOR.TOP if below else MSO_ANCHOR.BOTTOM, name=f'Valor {lab}'))
        else:
            ly = top - 25000 - LBL_D
            new.append(txt(sh, cx(i) - int(slot / 2), ly, int(slot), LBL_D, sg(v), 10, name=f'Valor {lab.replace(chr(10), " ")}'))
        new.append(txt(sh, cx(i) - int(slot / 2), 6500000, int(slot), 300000, lab, 9, anchor=MSO_ANCHOR.TOP,
                       name=f'Categoria {lab.replace(chr(10), " ")}'))
        if i < n - 1:
            new.append(line(sh, x + bw, y(lvl), cx(i + 1) - bw // 2, y(lvl), BLACK, 0.25, 'lgDash', name='Conector'))
    xl, xr = cx(0), cx(n - 1)

    def label_top(i):
        v = cols[i][2]
        return (y(0) - 20000) if v < 0 else (y(bars[i][1]) - 25000 - LBL_T - 30000)

    new.append(line(sh, xl, Y_BR, xl, label_top(0), GREY, 0.75, name='Colchete esq'))
    new.append(line(sh, xl, Y_BR, xr, Y_BR, GREY, 0.75, name='Colchete topo'))
    new.append(line(sh, xr, Y_BR, xr, label_top(n - 1), GREY, 0.75, arrow=True, name='Colchete dir'))
    if start > 0:
        ptxt = pc((end - start) / start, signed=True)
    else:
        ptxt = f"Δ {sg(end - start)}"
    ew, eh = (1010000 if len(ptxt) > 6 else 800000), 300000
    el = box(sh, (xl + xr) // 2 - ew // 2, Y_BR - eh // 2, ew, eh, WHITE, 'Variacao', MSO_SHAPE.OVAL, GREY)
    tf = el.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = ptxt
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = BLACK
    new.append(el)
    grp = sh.add_group_shape(new)
    grp.name = 'Bridge MP27 (Base Book)'
    return grp


# ---------------------------------------------------------------- painel de comentarios
PX, PW, PY, PH = 8720000, 3230000, 520000, 5900000


def draw_comments(slide, title, blocks, size=10.5):
    """blocks: lista de (titulo, [bullets]) ou (titulo, 'texto corrido')."""
    sh = slide.shapes
    box(sh, PX, PY, PW, PH, PANEL, 'Painel comentarios')
    hd = box(sh, PX, PY, PW, 330000, HEAD, 'Painel titulo')
    tf = hd.text_frame
    tf.margin_left = Emu(90000)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = f"Principais mensagens – {title}"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = WHITE
    tb = sh.add_textbox(Emu(PX + 90000), Emu(PY + 410000), Emu(PW - 180000), Emu(PH - 470000))
    tb.name = 'Comentarios'
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    first = True
    for lead, body in blocks:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        if not first:
            p.space_before = Pt(8)
        first = False
        no_bullet(p)
        r1 = p.add_run()
        r1.text = lead
        r1.font.size = Pt(size + 1)
        r1.font.bold = True
        r1.font.color.rgb = HEAD
        items = body if isinstance(body, list) else [body]
        for it in items:
            q = tf.add_paragraph()
            q.space_before = Pt(2)
            if isinstance(body, list):
                set_bullet(q)
            else:
                no_bullet(q)
            r2 = q.add_run()
            r2.text = it
            r2.font.size = Pt(size)
            r2.font.color.rgb = BLACK
    return tb


# ---------------------------------------------------------------- tabela de P&L nativa (layout do Base Book)
LN_TAGS = {'L': 'a:lnL', 'R': 'a:lnR', 'T': 'a:lnT', 'B': 'a:lnB'}


def cell_lines(cell, spec):
    """spec: {'L': (rgb_hex, w_pt) ou None, ...} para os 4 lados."""
    tcPr = cell._tc.get_or_add_tcPr()
    for t in LN_TAGS.values():
        for e in tcPr.findall(qn(t)):
            tcPr.remove(e)
    for i, side in enumerate('LRTB'):
        v = spec.get(side)
        ln = etree.Element(qn(LN_TAGS[side]), w=str(int((v[1] if v else 0) * 12700)), cap='flat', cmpd='sng', algn='ctr')
        if v:
            sf = etree.SubElement(ln, qn('a:solidFill'))
            etree.SubElement(sf, qn('a:srgbClr'), val=v[0])
            etree.SubElement(ln, qn('a:prstDash'), val='solid')
        else:
            etree.SubElement(ln, qn('a:noFill'))
        tcPr.insert(i, ln)


def strip_table_style(gt):
    tbl = gt._element.graphic.graphicData.tbl
    tblPr = tbl.tblPr
    for a in ('firstRow', 'bandRow', 'firstCol', 'lastRow', 'lastCol', 'bandCol'):
        tblPr.set(a, '0')
    for e in tblPr.findall(qn('a:tableStyleId')):
        tblPr.remove(e)


def put(cell, runs, align=PP_ALIGN.RIGHT, fill=None, anchor=MSO_ANCHOR.MIDDLE, ml=40000, mr=60000):
    """runs: lista de (texto, size, bold, italic, rgb)."""
    cell.margin_left, cell.margin_right = Emu(ml), Emu(mr)
    cell.margin_top = cell.margin_bottom = Emu(0)
    cell.vertical_anchor = anchor
    if fill is None:
        cell.fill.background()
    else:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    tf = cell.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = align
    for t, s, b, it, c in runs:
        r = p.add_run()
        r.text = t
        r.font.size = Pt(s)
        r.font.bold = b
        r.font.italic = it
        r.font.color.rgb = c
    # fim de paragrafo com a menor fonte (celula vazia nao herda 18pt e nao estica a linha)
    p._p.get_or_add_endParaRPr().set('sz', str(int(min(x[1] for x in runs) * 100)))


def pl_table(slide, u):
    d = PL[u]
    big = 13
    small = 9.5
    rows = [  # (rotulo, chave, grande?, formato)
        ('Pieces', 'pcs', True, 'int'), ('Net Sales', 'ns', True, 'v'), ('NS/pc', 'nspc', False, 'v'),
        ('CM', 'cm', True, 'v'), ('%', 'cmp', False, 'p'), ('Fixed Cost', 'fix', True, 'v'),
        ('Other Income', 'oi', True, 'v'), ('EBIT', 'ebit', True, 'v'), ('EBIT/pc', 'ebpc', False, 'v'),
        ('Ros%', 'ros', True, 'p')]
    colw = [1.45, 1.03, 1.03, 1.03, 1.03, 1.03, 0.14, 1.06, 1.06]
    rowh = [0.30, 0.27] + [0.33 if g else 0.22 for _, _, g, _ in rows] + [0.07, 0.27]
    rowh[-3] = 0.26   # Ros%
    nr, nc = len(rowh), len(colw)
    gt = slide.shapes.add_table(nr, nc, Inches(0.25), Inches(0.56), Inches(sum(colw)), Inches(sum(rowh)))
    gt.name = 'Tabela P&L (Base Book)'
    strip_table_style(gt)
    tb = gt.table
    for j, w in enumerate(colw):
        tb.columns[j].width = Inches(w)
    for i, h in enumerate(rowh):
        tb.rows[i].height = Inches(h)
    OUTER, INNER = ('000000', 0.75), ('7F7F7F', 0.5)
    hdr = ['Actual 2024', 'Actual 2025', "MP'26", 'Fcst R9', "MP'27"]
    dhdr = ["MP'27 vs R9", "MP'27 vs MP'26"]
    # cabecalho
    tb.cell(0, 1).merge(tb.cell(0, 5))
    put(tb.cell(0, 0), [('(+) gain  |  mln BRL', 8, False, False, DGREY)], PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
    put(tb.cell(0, 1), [('Full Year', 12, True, False, WHITE)], PP_ALIGN.CENTER, NAVY, MSO_ANCHOR.BOTTOM)
    for j in range(1, 6):
        put(tb.cell(1, j), [(hdr[j - 1], 10, False, False, WHITE)], PP_ALIGN.CENTER, NAVY, MSO_ANCHOR.TOP)
    for k, j in enumerate((7, 8)):
        put(tb.cell(0, j), [('Δ', 12, True, False, NAVY)], PP_ALIGN.CENTER, ACC, MSO_ANCHOR.BOTTOM)
        put(tb.cell(1, j), [(dhdr[k], 9.5, False, False, NAVY)], PP_ALIGN.CENTER, ACC, MSO_ANCHOR.TOP)
    put(tb.cell(1, 0), [('', 8, False, False, BLACK)], fill=None)
    for i in (0, 1):
        put(tb.cell(i, 6), [('', 8, False, False, BLACK)], fill=None)
    # corpo
    for k, (lab, key, g, fmt) in enumerate(rows):
        i = k + 2
        if lab == 'Pieces':
            lr = [('Pieces ', big, True, False, BLACK), ('(/000)', 10, True, False, BLACK)]
        elif g:
            lr = [(lab, big if lab != 'Fixed Cost' else 11.5, True, False, BLACK)]
        else:
            lr = [(lab, small, False, True, DGREY)]
        put(tb.cell(i, 0), lr, PP_ALIGN.RIGHT, mr=90000)
        for j in range(5):
            v = d[key][j]
            t = fp(v, 0) if fmt == 'int' else (pc(v) if fmt == 'p' else fp(v))
            if u == 'RES' and j < 3 and key not in ('oi', 'ebit') and abs(v) < 1e-9:
                t = '–'
            put(tb.cell(i, j + 1), [(t, big if g else small, g, False, BLACK)], PP_ALIGN.CENTER)
        for k2, j in enumerate((7, 8)):
            v = d['d_' + key][k2]
            t = fp(v, 0) if fmt == 'int' else (pc(v) if fmt == 'p' else fp(v))
            put(tb.cell(i, j), [(t, big - 1 if g else small, g, False, DGREY)], PP_ALIGN.CENTER)
        put(tb.cell(i, 6), [('', 8, False, False, BLACK)])
    # espacador + linha de delta EBIT
    sp = nr - 2
    for j in range(nc):
        put(tb.cell(sp, j), [('', 4, False, False, BLACK)])
    last = nr - 1
    put(tb.cell(last, 0), [('Δ EBIT', 12, True, False, BLACK)], PP_ALIGN.RIGHT, mr=90000)
    for j in (1, 6, 7, 8):
        put(tb.cell(last, j), [('', 8, False, False, BLACK)])
    for k, j in enumerate((2, 3, 4, 5)):
        v = d['debit'][k]
        put(tb.cell(last, j), [('●  ', 11, False, False, GREEN if v >= 0 else RED), (fp(v), 10.5, False, False, BLACK)],
            PP_ALIGN.CENTER)
    # bordas
    body = range(2, 2 + len(rows))
    for i in range(nr):
        for j in range(nc):
            s = {}
            in_fy = 1 <= j <= 5 and i in body
            in_d = j in (7, 8) and i in body
            if in_fy or in_d:
                lo, hi = (1, 5) if in_fy else (7, 8)
                s['L'] = OUTER if j == lo else INNER
                s['R'] = OUTER if j == hi else INNER
                s['T'] = OUTER if i == 2 else None
                s['B'] = OUTER if i == 2 + len(rows) - 1 else None
            if i == last and 2 <= j <= 5:
                s = {'L': OUTER if j == 2 else INNER, 'R': OUTER if j == 5 else INNER, 'T': OUTER, 'B': OUTER}
            cell_lines(tb.cell(i, j), s)
    return gt


# ================================================================ comentarios por unidade
def dd(u, k):
    return BRG[u]['d'].get(k, 0.0)


def head_kpi(u):
    e9, e27, e26 = R(u, 'ebit', 3), R(u, 'ebit'), R(u, 'ebit', 2)
    vs9 = f"{sg(e27 - e9)} vs Fcst R9" + (f" ({pc((e27 - e9) / e9, signed=True)})" if e9 > 0 else '')
    return (f"EBIT {f1(e27)}  |  ROS {pc(R(u, 'ros'))}", f"{vs9} · {sg(e27 - e26)} vs MP'26")


mo_ibi = (mo('IBI', i=0), mo('IBI', i=1), mo('IBI'))
mo_sjp = (mo('SJP', i=0), mo('SJP', i=1), mo('SJP'))
d_sjp = mo_sjp[2] - mo_sjp[0]
sjp_sal = (mo('SJP', 'Salarios e PPR') + mo('SJP', 'Encargos') - mo('SJP', 'Salarios e PPR', 0) - mo('SJP', 'Encargos', 0)) / d_sjp
sjp_med = (mo('SJP', 'Assistencia medica') - mo('SJP', 'Assistencia medica', 0)) / d_sjp
sjp_med_g = mo('SJP', 'Assistencia medica') / mo('SJP', 'Assistencia medica', 0) - 1
ibi_g = lambda g, i: mo('IBI', g) / mo('IBI', g, i) - 1
ebit_sem_fiat_ibi = R('IBI', 'ebit') - FIAT_PRECO
ebit_sem_fiat_br = R('BR', 'ebit') - FIAT_PRECO
ros_sem_fiat_br = ebit_sem_fiat_br / (R('BR', 'ns') - FIAT_PRECO)

COMMENTS = {
    'BR': ("Brasil", [
        head_kpi('BR'),
        ('Receita +8,6%: preço e perímetro', [
            f"Price/Mix {sg(dd('BR', 'Price/ Mix'))}: reajuste da Fiat em IBI ({pc(FIAT_PR27 / FIAT_PR9 - 1, signed=True)}/pç) para cobrir o 3º turno.",
            f"Volume {sg(dd('BR', 'Volume'))}: RES em ano cheio ({fp(R('RES', 'pcs'), 0)}k pçs vs {fp(R('RES', 'pcs', 3), 0)}k no R9)."]),
        (f"MO {f1(dd('BR', 'Labour'))}: maior pressão de custo", [
            f"MO {pc(mo('ALL') / mo('ALL', i=0) - 1, signed=True)} vs R9, mas {pc(mo('ALL') / mo('ALL', i=1) - 1, signed=True)} vs run-rate: 3º turno IBI e RES em ano cheio."]),
        ('Eventos específicos', [
            f"ANTT {f1(dd('BR', 'ANTT'))} (frete mínimo, Fiat) e saída da Jaguar {f1(dd('BR', 'Encerr. Jaguar'))}."]),
        (f"Cost Fix. {f1(dd('BR', 'Cost Fix.'))}", [
            f"IBI {f1(dd('IBI', 'Cost Fix.'))} (serviços do 3º turno), RES {f1(dd('RES', 'Cost Fix.'))} (perímetro), "
            f"GOI {sg(dd('GOI', 'Cost Fix.'))}; SOR {f1(dd('BR', 'Cost Fix.') - sum(dd(x, 'Cost Fix.') for x in UNITS))} (encerrada)."]),
        ('Leitura', f"Sem o reajuste da Fiat, o EBIT do Brasil seria ~{f1(ebit_sem_fiat_br)} (ROS ~{pc(ros_sem_fiat_br, 0)})."),
    ]),
    'IBI': ("Ibirité (IBI)", [
        head_kpi('IBI'),
        (f"Saída da Jaguar {f1(dd('IBI', 'Encerr. Jaguar'))}",
         f"O R9 tinha {f1(JLR['vol'] / 1000)}k pçs e R$ {f1(JLR['fat'] / 1000)} mi de faturamento (margem direta ~{f1(JLR['margem'] / 1000)})."),
        ('Início do 3º turno', [
            f"MO {pc(mo_ibi[2] / mo_ibi[0] - 1, signed=True)} vs R9 ({sg(mo_ibi[2] - mo_ibi[0])} mi): {f1(mo_ibi[1] - mo_ibi[0])} é efeito ano cheio "
            f"e {f1(mo_ibi[2] - mo_ibi[1])} reajuste 2027 ({pc(mo_ibi[2] / mo_ibi[1] - 1, signed=True)} vs run-rate).",
            f"Reajuste vs run-rate: salários {pc(ibi_g('Salarios e PPR', 1), signed=True)}, assist. médica {pc(ibi_g('Assistencia medica', 1), signed=True)}.",
            f"Serviços {sg(T3_TOT)} mi: empilhadeiras {sg(T3['Locação'])}, alimentação {sg(T3['Alimentação'])}, "
            f"transporte {sg(T3['Transporte'])}, limpeza {sg(T3['Limpeza'])} (já líquidos de {f1(REDUCOES_IBI)} de reduções)."]),
        ('Preço adicional cobre o 3º turno', [
            f"Fiat R$ {f1(FIAT_PR9, 2)} → {f1(FIAT_PR27, 2)}/pç ({pc(FIAT_PR27 / FIAT_PR9 - 1, signed=True)}) = {sg(FIAT_PRECO)} mi; absorve também a ANTT ({f1(dd('IBI', 'ANTT'))}).",
            f"Sem o reajuste, o EBIT de IBI seria ~{f1(ebit_sem_fiat_ibi)}: fechar o preço com a Fiat é crítico."]),
    ]),
    'SJP': ("SJP", [
        head_kpi('SJP'),
        ("Desconto de −3,5% repassado ao cliente", [
            f"Preço R$ {f1(SJP_PRECO_ATUAL, 2)} → {f1(SJP_PRECO_27, 2)}/pç (≈ {f1(SJP_DESC_MI)} mi/ano vs preço atual).",
            f"Conseguimos compensar parte: Price/Mix vs R9 de só {f1(dd('SJP', 'Price/ Mix'))}, com volume estável ({fp(R('SJP', 'pcs'), 0)}k pçs)."]),
        (f"MO: maior impacto ({f1(dd('SJP', 'Labour'))})", [
            f"MO {pc(mo_sjp[2] / mo_sjp[0] - 1, signed=True)} vs R9 ({f1(mo_sjp[0])} → {f1(mo_sjp[2])} mi; {pc(mo_sjp[2] / mo_sjp[1] - 1, signed=True)} vs run-rate).",
            f"Do aumento: {pc(sjp_sal, 0)} salários e encargos e {pc(sjp_med, 0)} assistência médica (que sobe {pc(sjp_med_g, 0, True)} vs R9)."]),
        ('Demais custos estáveis',
         f"Cost Var. {sg(dd('SJP', 'Cost Var.'))} e Cost Fix. {sg(dd('SJP', 'Cost Fix.'))} vs R9."),
    ]),
    'GOI': ("Goiana (GOI)", [
        head_kpi('GOI'),
        ('Melhora de preço compensa o volume', [
            f"NS/pç R$ {f1(R('GOI', 'nspc'), 2)} ({pc(R('GOI', 'nspc') / R('GOI', 'nspc', 3) - 1, signed=True)} vs R9): Price/Mix {sg(dd('GOI', 'Price/ Mix'))}.",
            f"Volume {pc(R('GOI', 'pcs') / R('GOI', 'pcs', 3) - 1, signed=True)} ({fp(R('GOI', 'pcs'), 0)}k pçs): {f1(dd('GOI', 'Volume'))}."]),
        ('Custos fixos e variáveis equalizados', [
            f"Cost Fix. {sg(dd('GOI', 'Cost Fix.'))}: outros fixos {f1(GOI_OFIX[0])} vs {f1(GOI_OFIX[1])} no R9.",
            f"Compensa MO {f1(dd('GOI', 'Labour'))} e Cost Var. {f1(dd('GOI', 'Cost Var.'))} (frete ~R$ {f1(GOI_FRETE_MES, 0)}k/mês, {f1(GOI_FRETE_D)} vs R9)."]),
    ]),
    'RES': ("Resende (RES)", [
        head_kpi('RES'),
        ('Delta de perímetro', [
            f"A unidade não esteve completa e aberta durante todo 2026: {fp(R('RES', 'pcs'), 0)}k pçs no MP'27 vs {fp(R('RES', 'pcs', 3), 0)}k no R9 (Volume {sg(dd('RES', 'Volume'))}).",
            f"MO de {f1(mo('RES'))} mi = run-rate set–dez/26: só efeito ano cheio."]),
        (f"ROS de {pc(R('RES', 'ros'), 0)} composto", [
            f"NS/pç R$ {f1(R('RES', 'nspc'), 2)} (R9: {f1(R('RES', 'nspc', 3), 2)}): Price/Mix {sg(dd('RES', 'Price/ Mix'))} cobre MO {f1(dd('RES', 'Labour'))} e estrutura fixa {f1(dd('RES', 'Cost Fix.'))}."]),
        ("vs MP'26", "No MP'26 só havia EBIT em Other Income (sem volume fechado); por isso o delta de volume vs MP'26 é elevado."),
    ]),
}

# ================================================================ slides novos
TITLE_Y = None


def new_slide(title):
    s = prs.slides.add_slide(LAYOUT)
    s.shapes.title.text = title
    tidx = s.shapes.title.placeholder_format.idx
    for ph in list(s.placeholders):
        if ph.placeholder_format.idx != tidx:
            ph.element.getparent().remove(ph.element)
    el = copy.deepcopy(FLAG._element)
    blip = el.xpath('.//a:blip')[0]
    blip.set(qn('r:embed'), s.part.relate_to(slides[1].part.related_part(blip.get(qn('r:embed'))), RT.IMAGE))
    s.shapes._spTree.append(el)
    return s


def headline(s, text, w=8150000):
    tb = s.shapes.add_textbox(Emu(300000), Emu(560000), Emu(w), Emu(520000))
    tb.name = 'Mensagem principal'
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    r = tf.paragraphs[0].add_run()
    r.text = text
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = HEAD
    return tb


def footnote(s, text, y=6440000, w=8150000):
    tb = s.shapes.add_textbox(Emu(300000), Emu(y), Emu(w), Emu(330000))
    tb.name = 'Nota de rodape'
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    r = tf.paragraphs[0].add_run()
    r.text = text
    r.font.size = Pt(8)
    r.font.color.rgb = GREY
    return tb


def simple_table(s, x, y, colw, rows, rowh, font=10, bold_last=False, chips=None, aligns=None, wrap_cols=()):
    n, m = len(rows), len(rows[0])
    gt = s.shapes.add_table(n, m, Emu(x), Emu(y), Emu(sum(colw)), Emu(rowh * n))
    strip_table_style(gt)
    tb = gt.table
    for j, wd in enumerate(colw):
        tb.columns[j].width = Emu(wd)
    for i, row in enumerate(rows):
        tb.rows[i].height = Emu(rowh)
        for j, val in enumerate(row):
            c = tb.cell(i, j)
            c.margin_left = c.margin_right = Emu(60000)
            c.margin_top = c.margin_bottom = Emu(10000)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.fill.solid()
            last = bold_last and i == n - 1
            c.fill.fore_color.rgb = HEAD if i == 0 else (LIGHT if last else (PANEL if i % 2 == 0 else WHITE))
            c.text_frame.word_wrap = j in wrap_cols
            p = c.text_frame.paragraphs[0]
            p.alignment = aligns[j] if aligns else (PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT)
            if chips and i > 0 and j == 0 and chips.get(i):
                rc = p.add_run()
                rc.text = '● '
                rc.font.size = Pt(font)
                rc.font.color.rgb = chips[i]
            r = p.add_run()
            r.text = str(val)
            r.font.size = Pt(font)
            r.font.bold = (i == 0) or last
            r.font.color.rgb = WHITE if i == 0 else BLACK
            cell_lines(c, {'B': ('D9D9D9', 0.5)} if 0 < i < n - 1 else {})
    return gt


def kpi_tile(s, x, y, w, h, label, value, sub, sub_color, name):
    box(s.shapes, x, y, w, h, PANEL, name, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.08
    txt(s.shapes, x + 120000, y + 90000, w - 240000, 230000, label, 11, False, DGREY, PP_ALIGN.LEFT, MSO_ANCHOR.TOP, name + ' rotulo')
    txt(s.shapes, x + 120000, y + 330000, w - 240000, 430000, value, 26, True, HEAD, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE, name + ' valor')
    tb = txt(s.shapes, x + 120000, y + 790000, w - 240000, 330000, '', 10.5, False, sub_color, PP_ALIGN.LEFT, MSO_ANCHOR.TOP, name + ' delta', wrap=True)
    tf = tb.text_frame
    tf.paragraphs[0].runs[0].text = sub[0]
    tf.paragraphs[0].runs[0].font.bold = True
    if len(sub) > 1:
        r = tf.paragraphs[0].add_run()
        r.text = sub[1]
        r.font.size = Pt(10.5)
        r.font.color.rgb = DGREY


def summary_slide():
    s = new_slide("Fitted Units Economics Brasil – Resumo executivo MP'27")
    b = PL['BR']
    e9, e27, e26 = b['ebit'][3], b['ebit'][4], b['ebit'][2]
    headline(s, f"MP'27: EBIT de R$ {f1(e27)} mi (ROS {pc(b['ros'][4])}), {sg(e27 - e26)} vs MP'26 e {sg(e27 - e9)} vs Fcst R9. "
                f"O preço compensa parte da pressão de MO, 3º turno e ANTT", w=11590000)
    tiles = [
        ('Volume (mil peças)', fp(b['pcs'][4], 0), (pc(b['pcs'][4] / b['pcs'][3] - 1, signed=True), ' vs Fcst R9'), GREEN),
        ('Net Sales (mln BRL)', f1(b['ns'][4]), (pc(b['ns'][4] / b['ns'][3] - 1, signed=True), f" vs R9 · NS/pç {f1(b['nspc'][4], 2)}"), GREEN),
        ('EBIT (mln BRL)', f1(e27), (pc((e27 - e9) / e9, signed=True), f" vs R9 · {pc((e27 - e26) / e26, signed=True)} vs MP'26"), RED),
        ('ROS', pc(b['ros'][4]), (f"{f1((b['ros'][4] - b['ros'][3]) * 100)} p.p.", f" vs R9 ({pc(b['ros'][3])})"), RED),
    ]
    W, G, X, Y, H = 2747500, 200000, 300000, 1230000, 1180000
    for k, (lab, val, sub, col) in enumerate(tiles):
        kpi_tile(s, X + k * (W + G), Y, W, H, lab, val, sub, col, f'KPI {k + 1}')
    # tabela por unidade
    msg = {'IBI': "Reajuste Fiat cobre o 3º turno; saída Jaguar",
           'SJP': "Desconto −3,5% ao cliente; MO +14%",
           'GOI': "Preço compensa volume; custos equalizados",
           'RES': "Perímetro: ano cheio, ROS de 10%"}
    rows = [["Unidade", "Peças (k)", "Net Sales", "EBIT R9", "EBIT MP'27", "Δ vs R9", "ROS", "% EBIT", "Mensagem"]]
    for u in UNITS:
        rows.append([u, fp(R(u, 'pcs'), 0), f1(R(u, 'ns')), fp(R(u, 'ebit', 3)), fp(R(u, 'ebit')), fp(R(u, 'ebit') - R(u, 'ebit', 3)),
                     pc(R(u, 'ros')), pc(R(u, 'ebit') / e27, 0), msg[u]])
    rows.append(["Brasil", fp(b['pcs'][4], 0), f1(b['ns'][4]), fp(e9), fp(e27), fp(e27 - e9), pc(b['ros'][4]), '100%', "R9 inclui 0,2 da SOR (encerrada)"])
    L_, Rt = PP_ALIGN.LEFT, PP_ALIGN.RIGHT
    simple_table(s, 300000, 2650000, [820000, 820000, 820000, 780000, 900000, 760000, 700000, 700000, 5290000], rows, 330000, 10.5,
                 bold_last=True, chips={1 + k: UCOL[u] for k, u in enumerate(UNITS)}, aligns=[L_] + [Rt] * 7 + [L_])
    footnote(s, "Fonte: Base Book Fitted Units MP27 (19/10) e MENS FITTED MP27 v3. Valores em mln BRL. ROS = EBIT / Net Sales.", y=4720000, w=11590000)
    # painel direito: o que sustenta / o que ameaca
    items = [
        ('O que sustenta o plano', [
            f"Preço: Price/Mix {sg(BRG['BR']['d']['Price/ Mix'])}, puxado pelo reajuste Fiat ({sg(FIAT_PRECO)} mi).",
            f"Perímetro RES: Volume {sg(dd('RES', 'Volume'))} e ROS de {pc(R('RES', 'ros'), 0)}.",
            f"GOI estável: custos fixos {sg(dd('GOI', 'Cost Fix.'))} compensam MO e frete."]),
        ('O que pressiona', [
            f"MO {f1(dd('BR', 'Labour'))} ({pc(mo('ALL') / mo('ALL', i=0) - 1, signed=True)} vs R9).",
            f"3º turno IBI: serviços +{f1(T3_TOT)} mi; ANTT {f1(dd('BR', 'ANTT'))}; Jaguar {f1(dd('BR', 'Encerr. Jaguar'))}.",
            f"Concentração: a Fiat responde por {pc(C_EBIT[0] / 1000 / e27, 0)} do EBIT do Brasil."]),
        ('Para decidir', [
            f"Fechar o reajuste com a Fiat: sem ele o EBIT seria ~{f1(ebit_sem_fiat_br)}.",
            f"ROS de referência {pc(ROS_REF)} (EBIT {f1(EBIT_REF)}): faltam {f1(EBIT_REF - e27)} mi.",
            f"Cada 1% de volume ≈ R$ {f1(SENS)} mi de EBIT."]),
    ]
    sh = s.shapes
    y0 = 5080000
    colw3 = 3730000
    for k, (lead, bl) in enumerate(items):
        x = 300000 + k * (colw3 + 200000)
        box(sh, x, y0, colw3, 1420000, PANEL, f'Bloco {k + 1}', MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.06
        tb = sh.add_textbox(Emu(x + 120000), Emu(y0 + 90000), Emu(colw3 - 240000), Emu(1260000))
        tb.name = f'Bloco {k + 1} texto'
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        no_bullet(p)
        r = p.add_run()
        r.text = lead
        r.font.size = Pt(12)
        r.font.bold = True
        r.font.color.rgb = HEAD
        for it in bl:
            q = tf.add_paragraph()
            q.space_before = Pt(3)
            set_bullet(q)
            r = q.add_run()
            r.text = it
            r.font.size = Pt(10.5)
    s.notes_slide.notes_text_frame.text = ("Resumo executivo. Fontes: Base Book (Month_FY_LC e abas de unidade, AK38:AT60), "
                                           "MENS FITTED MP27 v3 (TOTAL, IBI, Impacto JLR), Detalhe de Despesas V3 (Comparativo 3T).")
    return s


def set_hole(chart, pct):
    for el in chart._chartSpace.xpath('.//c:holeSize'):
        el.set('val', str(pct))


def donut(s, x, y, dsz, order, vals, big, small):
    tot = sum(vals)   # mesma base da tabela (inclui RES negativo no R9)
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
            r.font.size = Pt(14)
            r.font.bold = True
            r.font.color.rgb = BLACK if u in ('GOI', 'RES') else WHITE
    set_hole(ch, 62)
    cx, cy = x + dsz // 2, y + dsz // 2
    txt(s.shapes, cx - 700000, cy - 300000, 1400000, 330000, big, 24, True, HEAD, anchor=MSO_ANCHOR.BOTTOM, name='Centro valor')
    txt(s.shapes, cx - 700000, cy + 40000, 1400000, 200000, small, 10, False, GREY, anchor=MSO_ANCHOR.TOP, name='Centro legenda')


def mix_slide():
    s = new_slide("Fitted Units Economics Brasil – Mix do EBIT por unidade")
    r9 = [R(u, 'ebit', 3) for u in UNITS]
    mp = [R(u, 'ebit') for u in UNITS]
    t9, tm = sum(r9), sum(mp)
    sh9, shm = [v / t9 for v in r9], [v / tm for v in mp]
    br9, brm = R('BR', 'ebit', 3), R('BR', 'ebit')
    top2 = shm[0] + shm[1]
    headline(s, f"IBI e SJP concentram {round(top2 * 100)}% do EBIT do MP'27; GOI ganha peso e a RES passa a contribuir")
    D, X9, XM, Y = 2850000, 700000, 4900000, 1480000
    txt(s.shapes, X9, 1130000, D, 300000, "Fcst R9 – 2026", 14, True, HEAD, anchor=MSO_ANCHOR.MIDDLE, name='Titulo R9')
    txt(s.shapes, XM, 1130000, D, 300000, "MP'27 – 2027", 14, True, HEAD, anchor=MSO_ANCHOR.MIDDLE, name='Titulo MP27')
    donut(s, X9, Y, D, UNITS, r9, f1(br9), 'EBIT Brasil')
    donut(s, XM, Y, D, UNITS, mp, f1(brm), 'EBIT Brasil')
    ar = box(s.shapes, 3800000, 2650000, 950000, 520000, ACC, 'Seta R9 para MP27', MSO_SHAPE.RIGHT_ARROW)
    txt(s.shapes, 3700000, 3230000, 1150000, 260000, pc((brm - br9) / br9, signed=True), 16, True, RED, anchor=MSO_ANCHOR.MIDDLE, name='Variacao EBIT')
    txt(s.shapes, 3700000, 3480000, 1150000, 200000, f"Δ {sg(brm - br9)} mln BRL", 10, False, GREY, anchor=MSO_ANCHOR.TOP, name='Variacao absoluta')
    pcs_t, ns_t = R('BR', 'pcs'), R('BR', 'ns')
    rows = [["Unidade", "% Peças", "% Net Sales", "EBIT R9", "% R9", "EBIT MP'27", "% MP'27", "ROS MP'27", "Δ EBIT"]]
    for k, u in enumerate(UNITS):
        rows.append([u, pc(R(u, 'pcs') / pcs_t, 0), pc(R(u, 'ns') / ns_t, 0), fp(r9[k]), pc(sh9[k], 0), fp(mp[k]), pc(shm[k], 0),
                     pc(R(u, 'ros')), fp(mp[k] - r9[k])])
    sor = br9 - t9
    rows.append(["SOR (encerrada)", '–', '–', fp(sor), '–', fp(0.0), '–', '–', fp(-sor)])
    rows.append(["Brasil", '100%', '100%', fp(br9), '', fp(brm), '', pc(R('BR', 'ros')), fp(brm - br9)])
    simple_table(s, 300000, 4520000, [1500000, 800000, 950000, 850000, 750000, 950000, 800000, 900000, 650000], rows, 235000, 10,
                 bold_last=True, chips={1 + k: UCOL[u] for k, u in enumerate(UNITS)})
    footnote(s, f"Fonte: Base Book Fitted Units MP27 (19/10). Pizzas consideram as 4 unidades (R9: {f1(t9)}; MP'27: {f1(tm)}); o R9 do Brasil ({f1(br9)}) "
                f"inclui {f1(sor)} de outras receitas da SOR. RES no R9 = {f1(r9[3])} ({pc(sh9[3])} do EBIT), sem fatia na pizza.", 6330000)
    fiat_br = C_EBIT[0] / 1000 / brm
    draw_comments(s, "Mix do EBIT", [
        ('IBI segue como âncora', f"{round(sh9[0] * 100)}% → {round(shm[0] * 100)}% do EBIT ({f1(r9[0])} → {f1(mp[0])}); {pc(R('IBI', 'ns') / ns_t, 0)} do Net Sales do Brasil."),
        ('SJP perde peso', f"{round(sh9[1] * 100)}% → {round(shm[1] * 100)}% ({f1(r9[1])} → {f1(mp[1])}): desconto de preço negociado com o cliente e MO."),
        ('GOI estável e com mais peso', f"EBIT de {f1(mp[2])} (= R9); participação de {round(sh9[2] * 100)}% para {round(shm[2] * 100)}%."),
        ('RES passa a contribuir', f"EBIT de {f1(r9[3])} para {f1(mp[3])} ({round(shm[3] * 100)}%), com ROS de {pc(R('RES', 'ros'), 0)}."),
        ('Concentração de risco', f"IBI + SJP = {round(top2 * 100)}% do EBIT. Só a Fiat responde por {pc(fiat_br, 0)} do EBIT do Brasil."),
    ])
    return s


def client_slide():
    s = new_slide("Fitted Units Economics Brasil – IBI | Rentabilidade por cliente")
    sh_pc = [C_PCS[i] / C_PCS[3] for i in range(3)]
    sh_ns = [C_NS[i] / C_NS[3] for i in range(3)]
    sh_eb = [C_EBIT[i] / C_EBIT[3] for i in range(3)]
    headline(s, f"Com o reajuste, a Fiat tem ROS de {pc(C_ROS[0])} e gera {round(sh_eb[0] * 100)}% do EBIT de Ibirité; "
                f"a CNH (ROS {pc(C_ROS[2])}) é o ponto fraco da carteira")
    txt(s.shapes, 300000, 1130000, 4000000, 260000, "Peso de cada cliente (% do total IBI)", 11, True, HEAD, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE, name='Titulo grafico peso')
    txt(s.shapes, 4500000, 1130000, 3950000, 260000, f"ROS por cliente (meta mínima Fiat: {pc(ROS_META)})", 11, True, HEAD, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE, name='Titulo grafico ROS')
    cd = CategoryChartData()
    cd.categories = ['Peças', 'Net Sales', 'EBIT']
    for i, c in enumerate(CL):
        cd.add_series(c, [sh_pc[i], sh_ns[i], sh_eb[i]])
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED_100, Emu(300000), Emu(1400000), Emu(4000000), Emu(2550000), cd)
    gf.name = 'Peso por cliente'
    ch = gf.chart
    ch.has_title = False
    ch.font.size = Pt(10)
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.RIGHT
    ch.legend.include_in_layout = False
    ch.legend.font.size = Pt(10)
    ch.value_axis.visible = False
    ch.value_axis.has_major_gridlines = False
    ch.category_axis.format.line.color.rgb = GREY
    ch.plots[0].gap_width = 55
    shares = [sh_pc, sh_ns, sh_eb]
    for i, c in enumerate(CL):
        ser = ch.plots[0].series[i]
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = CCOL[c]
        ser.format.line.color.rgb = WHITE
        dl = ser.data_labels
        dl.show_value = True
        dl.number_format = '0%'
        dl.number_format_is_linked = False
        dl.font.size = Pt(11)
        dl.font.bold = True
        dl.font.color.rgb = WHITE if c == 'Fiat' else BLACK
        dl.position = XL_LABEL_POSITION.CENTER
        for k in range(3):
            if shares[k][i] < 0.05:
                ser.points[k].data_label.text_frame.text = ''
    cd2 = CategoryChartData()
    cd2.categories = ['Fiat', 'Iveco', 'CNH', 'Total IBI']
    cd2.add_series('ROS', C_ROS)
    gf2 = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Emu(4500000), Emu(1400000), Emu(3950000), Emu(2550000), cd2)
    gf2.name = 'ROS por cliente'
    c2 = gf2.chart
    c2.has_title = False
    c2.has_legend = False
    c2.font.size = Pt(10)
    c2.value_axis.visible = False
    c2.value_axis.has_major_gridlines = False
    c2.value_axis.maximum_scale = 0.22
    c2.value_axis.minimum_scale = 0
    c2.category_axis.format.line.color.rgb = GREY
    c2.plots[0].gap_width = 60
    c2.plots[0].vary_by_categories = False
    ser = c2.plots[0].series[0]
    for k, col in enumerate([CCOL['Fiat'], CCOL['Iveco'], CCOL['CNH'], GREY]):
        pt = ser.points[k]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = col
    c2.plots[0].has_data_labels = True
    dl = c2.plots[0].data_labels
    dl.show_value = True
    dl.number_format = '0.0%'
    dl.number_format_is_linked = False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size = Pt(12)
    dl.font.bold = True
    # linha da meta (plot area aproximada: 8% a 92% da altura do grafico)
    gy0, gh = 1400000, 2550000
    ymeta = int(gy0 + gh * 0.08 + gh * 0.80 * (1 - ROS_META / 0.22))
    line(s.shapes, 4600000, ymeta, 8400000, ymeta, RED, 1.0, 'dash', name='Meta ROS Fiat')
    txt(s.shapes, 6480000, ymeta - 190000, 900000, 170000, f"meta {pc(ROS_META)}", 9, True, RED, PP_ALIGN.CENTER, name="Meta rotulo")
    k1 = lambda v: fp(v / 1000.0)
    rows = [["mln BRL", "Fiat", "Iveco", "CNH", "Total IBI"],
            ["Peças (mil)"] + [fp(v, 0) for v in C_PCS],
            ["Net Sales"] + [k1(v) for v in C_NS],
            ["NS/pç (R$)"] + [fp(v, 2) for v in C_NSPC],
            ["CM"] + [k1(v) for v in C_CM],
            ["Fixed Cost"] + [k1(v) for v in C_FIX],
            ["EBIT"] + [k1(v) for v in C_EBIT],
            ["ROS %"] + [pc(v) for v in C_ROS],
            ["% do EBIT IBI"] + [pc(v) for v in sh_eb] + ['100,0%']]
    simple_table(s, 300000, 4080000, [2150000, 1500000, 1500000, 1500000, 1500000], rows, 235000, 10)
    footnote(s, "Fonte: Rentabilidade por Cliente IBI – MP27 v3 (ligada à MENS v3, aba IBI). Faturamento e material de consumo diretos; frete 100% Fiat; "
                f"demais custos rateados por faturamento. Custos dedicados CNH ({f1(CNH_ESPEC, 2)} mi/ano) e Iveco realocados do total IBI.", 6250000)
    MES = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']
    cnh_zero = [k for k in range(12) if CNH_VOL[k] == 0]
    cnh_neg = [MES[k] for k in cnh_zero]
    cnh_perda = -sum(CNH_MES[k] for k in cnh_zero) / len(cnh_zero)
    draw_comments(s, "Clientes IBI", [
        ('Fiat: escala e margem recuperada', f"{pc(sh_pc[0], 0)} das peças, {pc(sh_ns[0], 0)} do Net Sales e {pc(sh_eb[0], 0)} do EBIT. "
                                             f"ROS {pc(C_ROS[0])} (meta {pc(ROS_META)}) com o preço de R$ {f1(FIAT_PR27, 2)}/pç."),
        ('Iveco: boa margem, menos volume', f"ROS {pc(C_ROS[1])}, mas volume {pc(IVECO_V27 / IVECO_V9 - 1, signed=True)} vs R9 "
                                            f"({fp(IVECO_V9 / 1000, 0)}k → {fp(IVECO_V27 / 1000, 0)}k pçs; {f1(IVECO_DF)} mi de faturamento)."),
        ('CNH: o ponto fraco', [
            f"ROS {pc(C_ROS[2])}: custos dedicados de {f1(CNH_ESPEC, 2)} mi/ano (2.470 m², empilhadeira e 2 pessoas).",
            f"Sem volume em {' e '.join(cnh_neg)}, mas com custo dedicado: EBIT negativo de R$ {f1(cnh_perda, 0)}k em cada mês."]),
        ('Próximo passo', "Rever preço da CNH (R$ 115,46/pç) ou dimensionar o custo dedicado; validar Iveco/CNH com Custos."),
    ])
    return s


def mo_slide():
    s = new_slide("Fitted Units Economics Brasil – Mão de obra MP'27")
    t9, trr, t27 = mo('ALL', i=0), mo('ALL', i=1), mo('ALL')
    ano_cheio, reaj = trr - t9, t27 - trr
    headline(s, f"MO sobe {pc(t27 / t9 - 1)} vs R9 ({sg(t27 - t9)} mi), mas só {pc(t27 / trr - 1)} vs run-rate: "
                f"{pc(ano_cheio / (t27 - t9), 0)} do aumento é efeito ano cheio de 2026 (3º turno IBI e RES)")
    txt(s.shapes, 300000, 1130000, 8150000, 260000, "MO por unidade – R$ mi (custo direto, sem rateio da Gerência)", 11, True, HEAD, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE, name='Titulo grafico MO')
    cd = CategoryChartData()
    cd.categories = UNITS
    for nm, i in (('Fcst R9 (ano)', 0), ('Run-rate R9 (set–dez × 12)', 1), ("MP'27", 2)):
        cd.add_series(nm, [round(mo(u, i=i), 2) for u in UNITS])
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Emu(300000), Emu(1400000), Emu(8150000), Emu(2500000), cd)
    gf.name = 'MO por unidade'
    ch = gf.chart
    ch.has_title = False
    ch.font.size = Pt(10)
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.TOP
    ch.legend.include_in_layout = False
    ch.legend.font.size = Pt(10)
    ch.value_axis.visible = False
    ch.value_axis.has_major_gridlines = False
    ch.category_axis.format.line.color.rgb = GREY
    ch.category_axis.tick_labels.font.size = Pt(11)
    ch.category_axis.tick_labels.font.bold = True
    ch.plots[0].gap_width = 70
    ch.plots[0].overlap = -10
    for k, col in enumerate((LIGHT, ACC, DARK)):
        ser = ch.plots[0].series[k]
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = col
        ser.format.line.color.rgb = DARK if k == 0 else col
        ser.format.line.width = Pt(0.5)
        dl = ser.data_labels
        dl.show_value = True
        dl.number_format = '#,##0.0'
        dl.number_format_is_linked = False
        dl.font.size = Pt(10)
        dl.font.bold = k == 2
        dl.position = XL_LABEL_POSITION.OUTSIDE_END
    rows = [["R$ mi", "R9 ano", "Run-rate", "MP'27", "Δ vs R9", "Δ % R9", "Δ % run-rate", "Sal. + enc. vs RR", "Assist. méd. vs RR"]]
    for u in UNITS + ['GER', 'ALL']:
        a, b_, c = mo(u, i=0), mo(u, i=1), mo(u)
        sal = (mo(u, 'Salarios e PPR') + mo(u, 'Encargos')) / (mo(u, 'Salarios e PPR', i=1) + mo(u, 'Encargos', i=1)) - 1
        med = mo(u, 'Assistencia medica') / mo(u, 'Assistencia medica', i=1) - 1
        rows.append(['Total Fitted' if u == 'ALL' else u, f1(a), f1(b_), f1(c), sg(c - a), pc(c / a - 1, signed=True),
                     pc(c / b_ - 1, signed=True), pc(sal, signed=True), pc(med, signed=True)])
    simple_table(s, 300000, 4120000, [1300000, 800000, 850000, 800000, 800000, 850000, 950000, 950000, 850000], rows, 255000, 10, bold_last=True)
    footnote(s, "Fonte: Variação MO FITTED MP27 v2 x R9 x run-rate (Labour Cost FITTED MP27 v2 e R9). Run-rate = média set–dez/26 × 12. "
                "Δ % de RES vs R9 reflete a rampa de 2026. RR = run-rate.", 6000000)
    med_all = mo('ALL', 'Assistencia medica') / mo('ALL', 'Assistencia medica', i=0) - 1
    draw_comments(s, "Mão de obra", [
        (f"Efeito ano cheio {sg(ano_cheio)} mi", "3º turno de IBI e RES operando o ano todo: já está no run-rate de set–dez/26."),
        (f"Reajuste 2027 {sg(reaj)} mi ({pc(t27 / trr - 1, signed=True)})", [
            f"Salários e PPR {pc(mo('ALL', 'Salarios e PPR') / mo('ALL', 'Salarios e PPR', i=1) - 1, signed=True)}, "
            f"encargos {pc(mo('ALL', 'Encargos') / mo('ALL', 'Encargos', i=1) - 1, signed=True)}.",
            f"Assist. médica {pc(mo('ALL', 'Assistencia medica') / mo('ALL', 'Assistencia medica', i=1) - 1, signed=True)}: a linha que mais cresce "
            f"({pc(med_all, 0, True)} vs R9 ano)."]),
        ('Riscos a validar com o RH', [
            "Provisão de férias do horista em ~5,5% (estatutário 11,1%): até +0,95 mi.",
            "SJP CC 8289 (horista fixo) +35% vs run-rate (+0,46 mi) sem premissa visível.",
            "Assist. médica do horista GOI = 3,4% do salário (IBI 31%): possível +0,36 mi."]),
    ], size=10)
    return s


def qa_slide():
    s = new_slide("Fitted Units Economics Brasil – Perguntas-chave do MP'27")
    headline(s, "O que o plano precisa responder: preço, custo, volume e risco", w=11590000)
    b = PL['BR']
    cost_push = -(dd('BR', 'Labour') + dd('BR', 'Cost Fix.') + dd('BR', 'ANTT') + dd('BR', 'Encerr. Jaguar') + dd('BR', 'Cost Var.'))
    qa = [
        (f"Por que o EBIT cai se o faturamento sobe {pc(b['ns'][4] / b['ns'][3] - 1, 1)}?",
         f"Preço e volume somam {sg(dd('BR', 'Price/ Mix') + dd('BR', 'Volume'))}, mas os custos crescem ≈ {f1(cost_push)}: "
         f"MO {f1(-dd('BR', 'Labour'))}, fixos {f1(-dd('BR', 'Cost Fix.'))} (3º turno), ANTT {f1(-dd('BR', 'ANTT'))}, Jaguar {f1(-dd('BR', 'Encerr. Jaguar'))} "
         f"e variáveis {f1(-dd('BR', 'Cost Var.'))}."),
        ("Quanto do plano depende de preço novo?",
         f"Price/Mix de {sg(dd('BR', 'Price/ Mix'))}. Só o reajuste Fiat ({pc(FIAT_PR27 / FIAT_PR9 - 1, signed=True)}) vale {f1(FIAT_PRECO)} mi: "
         f"sem ele o EBIT seria ~{f1(ebit_sem_fiat_br)} (ROS ~{pc(ros_sem_fiat_br, 0)})."),
        ("Qual a sensibilidade a volume?",
         f"MO e handling são praticamente fixos no plano: cada 1% de volume ≈ R$ {f1(SENS)} mi de EBIT "
         f"({pc(SENS / b['ebit'][4], 0)} do EBIT)."),
        ("A mão de obra está sob controle?",
         f"{pc(mo('ALL') / mo('ALL', i=0) - 1, signed=True)} vs R9, mas {pc(mo('ALL') / mo('ALL', i=1) - 1, signed=True)} vs run-rate. "
         f"Assistência médica é o item mais inflacionado ({pc(mo('ALL', 'Assistencia medica') / mo('ALL', 'Assistencia medica', i=1) - 1, signed=True)})."),
        ("Onde está o risco de margem por cliente?",
         f"CNH com ROS de {pc(C_ROS[2])} (custo dedicado de {f1(CNH_ESPEC, 2)} mi); Iveco perde {pc(-(IVECO_V27 / IVECO_V9 - 1), 0)} de volume. "
         f"A Fiat responde por {pc(C_EBIT[0] / 1000 / b['ebit'][4], 0)} do EBIT do Brasil."),
        ("Há custos fora do plano?",
         "Provisão de férias do horista (até −0,95 mi), assistência médica GOI (até −0,36 mi) e CC 8289 SJP (0,46 mi a validar)."),
    ]
    sh = s.shapes
    W, H, G = 3730000, 2350000, 200000
    for k, (q, a) in enumerate(qa):
        col, row = k % 3, k // 3
        x, y = 300000 + col * (W + G), 1350000 + row * (H + G)
        box(sh, x, y, W, H, PANEL, f'Card {k + 1}', MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.05
        c = box(sh, x + 140000, y + 150000, 420000, 420000, HEAD, f'Card {k + 1} numero', MSO_SHAPE.OVAL)
        tf = c.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = str(k + 1)
        r.font.size = Pt(16)
        r.font.bold = True
        r.font.color.rgb = WHITE
        tq = txt(sh, x + 680000, y + 130000, W - 820000, 520000, q, 14, True, HEAD, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE, f'Card {k + 1} pergunta', wrap=True)
        ta = txt(sh, x + 180000, y + 760000, W - 360000, H - 880000, a, 13.5, False, BLACK, PP_ALIGN.LEFT, MSO_ANCHOR.TOP, f'Card {k + 1} resposta', wrap=True)
    s.notes_slide.notes_text_frame.text = ("Perguntas prováveis do presidente. Sensibilidade = (Net Sales − materiais − frete − outros variáveis)/100, "
                                           "MENS TOTAL; MO e handling tratados como fixos (elasticidade MO/peças 0,01–0,12 no plano). "
                                           "Riscos de MO: Auditoria Labour FITTED MP27 (achados 1, 3 e 7) – hipóteses a confirmar com o RH.")
    return s


# ================================================================ montagem
UNIT_SLIDES = {1: 'BR', 3: 'IBI', 5: 'SJP', 6: 'GOI', 7: 'RES'}   # indice 0-based no deck base


def clean_unit_slide(s):
    rem = []
    for sh in s.shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE and sh.width > Inches(5):
            rem.append(sh)
        elif sh.shape_type == MSO_SHAPE_TYPE.GROUP and sh.name.startswith('Bridge'):
            rem.append(sh)
        elif sh.name in ('Painel comentarios', 'Painel titulo', 'Comentarios'):
            rem.append(sh)
    for sh in rem:
        el = sh._element
        rids = set(el.xpath('.//@r:id'))
        el.getparent().remove(el)
        for rid in rids:
            try:
                s.part.drop_rel(rid)
            except KeyError:
                pass
    return len(rem)


for idx, u in UNIT_SLIDES.items():
    s = slides[idx]
    n = clean_unit_slide(s)
    for shp in s.shapes:
        if shp.has_text_frame and shp.text_frame.text.strip() == 'Result Variant Fitted Units':
            shp.left = Emu(int((X0 + X1) / 2 - shp.width / 2))
            shp.top = Emu(3980000)
    pl_table(s, u)
    draw_bridge(s, BRG[u])
    title, blocks = COMMENTS[u]
    draw_comments(s, title, blocks)
    spec = BRG[u]
    rows = [('Fcst R9', spec['start'])] + [(l.replace('\n', ' '), v) for l, v in spec['deltas']] + [("MP'27", spec['end'])]
    s.notes_slide.notes_text_frame.text = (
        "Dados do bridge (Base_Book_Fitted Units_MP27.xlsx de 19/10, abas 'Tabela Variantes s targ' / 'CM s targ'), mln BRL:\n"
        + "\n".join(f"{l}\t{v:.3f}" for l, v in rows)
        + "\nTabela: valores do Base Book (AK38:AT60 da aba da unidade / Month_FY_LC para o Brasil).")
    print(idx + 1, u, 'removidos', n, [(l, f1(v)) for l, v in rows])

sumS, mixS, cliS, moS, qaS = summary_slide(), mix_slide(), client_slide(), mo_slide(), qa_slide()

lst = prs.slides._sldIdLst
ids = list(lst)
o, new = ids[:16], ids[16:]
final = [o[0], new[0], o[1], new[1], o[3], new[2], o[5], o[6], o[7], new[3], new[4]]
drop = [e for e in ids if e not in final]
for el in ids:
    lst.remove(el)
for el in final:
    lst.append(el)
for el in drop:   # solta as relacoes dos slides removidos
    prs.part.drop_rel(el.get(qn('r:id')))

prs.save(OUT)
print('salvo', OUT, len(final), 'slides')
print('SENS', round(SENS, 3), 'T3', T3, 'FIAT', FIAT_PR9, FIAT_PR27, FIAT_PRECO, 'SJP', SJP_PRECO_ATUAL, SJP_DESC, SJP_DESC_MI)
