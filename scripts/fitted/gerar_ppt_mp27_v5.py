"""Gera a v2 do PPT 'Fitted Units Meeting Budget'27': refaz o bridge (a partir do Base Book) e adiciona comentarios.
Uso: python build_v2.py <orig.pptx> <basebook.xlsx> <saida.pptx>
As imagens das tabelas e o restante dos slides nao sao tocados.
"""
import re
import sys
from decimal import Decimal, ROUND_HALF_UP

import openpyxl
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE, MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

SRC, BASE, OUT, CLI = sys.argv[1:5]
CORTE_EIXO = False   # corte de eixo nas barras R9/MP27 (v4 tinha; a Juliana pediu para retirar)

DARK = RGBColor(0x36, 0x4D, 0x6E)
LIGHT = RGBColor(0xDF, 0xE5, 0xEF)
TARGET = RGBColor(0x9B, 0xC2, 0xE6)
GREY = RGBColor(0x80, 0x80, 0x80)
BLACK = RGBColor(0, 0, 0)
HEAD = RGBColor(0x1F, 0x4E, 0x79)
PANEL = RGBColor(0xF2, 0xF2, 0xF2)

# ---------------------------------------------------------------- dados do Base Book
wb = openpyxl.load_workbook(BASE, data_only=True)

# (aba, linha do cabecalho, linha dos dados)
SRC_T = {'BR': ('Tabela Variantes targ', 2, 4), 'IBI': ('Tabela Variantes targ', 7, 9),
         'SJP': ('Tabela Variantes CM targ', 3, 5), 'GOI': ('Tabela Variantes CM targ', 8, 10),
         'RES': ('Tabela Variantes CM targ', 13, 15)}
SRC_N = {'BR': ('Tabela Variantes s targ', 2, 4), 'IBI': ('Tabela Variantes s targ', 6, 8),
         'SJP': ('Tabela Variantes CM s targ', 3, 5), 'GOI': ('Tabela Variantes CM s targ', 8, 10),
         'RES': ('Tabela Variantes CM s targ', 13, 15)}


def load(sheet, hr, dr):
    ws = wb[sheet]
    items, end = [], None
    for col in range(3, 20):
        h, v = ws.cell(hr, col).value, ws.cell(dr, col).value
        if h is None:
            continue
        h = str(h).strip()
        if h == "MP'27":
            continue
        if h == 'EBIT':
            end = v
            break
        if h.startswith('Delta') or h.startswith('%'):
            continue
        items.append((h, float(v)))
    start = items[0][1]
    deltas = items[1:]
    assert abs(start + sum(v for _, v in deltas) - end) < 0.003, (sheet, dr, start, deltas, end)
    return {'start': start, 'deltas': deltas, 'end': float(end)}


DATA = {m: {u: load(*src[u]) for u in src} for m, src in (('T', SRC_T), ('N', SRC_N))}


def r1(x):
    return Decimal(str(abs(x))).quantize(Decimal('0.1'), ROUND_HALF_UP) * (1 if x >= 0 else -1)


def f1(x):
    v = r1(x)
    return f"{v:.1f}".replace('.', ',').replace('-0,0', '0,0')


# ---------------------------------------------------------------- comentarios (Base Book + MENS FITTED MP27 v3)
def bullets(u, mode):
    t = mode == 'T'
    tg = dict(DATA['T'][u]['deltas']).get('Target', 0)
    fix_n = dict(DATA['N'][u]['deltas'])['Cost\nFix.']
    ebit = DATA['T'][u]['end']
    sem = ebit - tg
    pre = dict(DATA['T'][u]['deltas'])['Cost\nFix.']

    if t:
        fix = ('Target', f"{f1(tg)} ({round(tg / ebit * 100)}% do EBIT). Sem ele, o EBIT do MP'27 seria {f1(sem)}. "
                         f"A entrega depende do plano de redução de custos.")
    else:
        fix = ('Cost Fix. (líquido)', f"{f1(fix_n)} líquido: inclui o target de eficiência de {f1(tg)} sobre um efeito subjacente de {f1(pre)}. "
                                      f"Sem o target, o EBIT do MP'27 seria {f1(sem)}.")
    if u == 'BR':
        return ("Brasil", [
            ('EBIT 9,2 | ROS 12,6%', "-3,8 vs Fcst R9 (-29%) e -0,6 vs MP'26. Volume estável em 4.971k pçs (-3k vs R9; +422k vs MP'26), "
                                     "mas o NS/pç cai para 14,6 (R9: 15,4)."),
            ('Volume -0,3 | Price/Mix -1,1', "O ramp-up da RES (+1,5) compensa o volume de SJP, IBI e GOI (-1,8). Price/Mix: IBI -0,6, SJP -0,7, GOI -0,3, RES +0,6."),
            ('Labour -5,2 (principal driver)', "-3,9 variável e -1,3 fixo: IBI -2,5, SJP -1,2, RES -1,2, GOI -0,3. Considera reajuste de ~7%; "
                                               "dados oficiais do RH ainda pendentes."),
            ('Eventos específicos -2,0', "Saída da Jaguar (JLR) na IBI -0,6; nova regra de frete mínimo da ANTT -1,5 (IBI). Cost Var. -1,0."),
            ('Cost Fix. antes do target -1,7', "IBI -1,7 e RES -0,9, parcialmente compensados pela GOI +1,1. A SOR (encerrada em 2027) deixa de trazer 0,2 de outras receitas."),
            fix])
    if u == 'IBI':
        return ("Ibirité (IBI)", [
            ('EBIT 5,4 | ROS 13,5%', "-2,0 vs Fcst R9 (-27%) e -0,2 vs MP'26. 2.805k pçs (-118k vs R9; +126k vs MP'26); NS/pç 14,2 (R9: 14,6)."),
            ('Volume -0,7 | Price/Mix -0,6', "Volume Fiat -3,6% vs R9 (2.750k pçs) e sem volume de JLR em 2027."),
            ('Saída da Jaguar -0,6', "O R9 incluía 1,1k pçs de JLR e R$ 1,5 mi de faturamento líquido (margem direta de ~R$ 0,5 mi)."),
            ('Labour -2,5 | ANTT -1,5', "Labour: -2,2 variável e -0,3 fixo (reajuste de ~7%, dados do RH pendentes). "
                                         "Frete maior pela nova regra de frete mínimo da ANTT."),
            ('Cost Var. -0,2 | Cost Fix. -1,7', "Reduções de ~R$ 0,9 mi já consideradas (alimentação, empilhadeiras, limpeza, energia e transporte de volta a R9 + inflação)."),
            fix])
    if u == 'SJP':
        return ("SJP", [
            ('EBIT 2,0 | ROS 12,5%', "-2,1 vs Fcst R9 (-51%) e estável vs MP'26 (2,0). 711k pçs (-72k vs R9; +17k vs MP'26); NS/pç 22,2 (R9: 23,2)."),
            ('Volume -1,0 | Price/Mix -0,7', "Volume 9% abaixo do R9 (783k pçs); price/mix menor, com NS/pç mais baixo."),
            ('Labour -1,2', "-0,8 variável e -0,3 fixo; reajuste de ~7% aplicado, dados oficiais do RH pendentes."),
            ('Cost Var. 0,0 | Cost Fix. 0,0', "Sem variação subjacente de custos variáveis ou fixos antes do target."),
            fix])
    if u == 'GOI':
        return ("Goiana (GOI)", [
            ('EBIT 1,4 | ROS 10,3%', "Em linha com o Fcst R9 (+0,0) e -0,4 vs MP'26. 1.097k pçs (-26k vs R9; -78k vs MP'26); NS/pç 12,7 (R9: 13,0)."),
            ('Volume -0,2 | Price/Mix -0,3', "Montagens Jeep; outubro reflete as férias coletivas da Jeep (menos dias úteis)."),
            ('Labour -0,3', "-0,03 variável e -0,3 fixo (reajuste de ~7%, dados do RH pendentes)."),
            ('Cost Var. -0,7', "Frete reajustado para ~R$ 60k/mês (+0,4 vs R9), além de materiais, manuseio e outros custos variáveis."),
            ('Cost Fix. antes do target +1,1', "Outros custos fixos de 1,2 vs 2,6 no R9."),
            fix])
    return ("Resende (RES)", [
        ('EBIT 0,4 | ROS 11,9%', "+0,5 vs Fcst R9 (-0,1) e +0,0 vs MP'26 (0,3, só outras receitas). 357k pçs vs 144k no R9; Net Sales 3,2."),
        ('Volume +1,5 | Price/Mix +0,6', "Operação no ano cheio; NS/pç 8,9 vs 7,3 (preço de lista – a unidade ainda precisa identificar seus ganhos)."),
        ('Labour -1,2', "-0,85 variável e -0,34 fixo, com o reajuste de ~7% agora aplicado também à RES."),
        ('Cost Var. -0,1 | Cost Fix. -0,9', "Estrutura fixa de 1,7 vs 1,1 no R9."),
        fix])


# ---------------------------------------------------------------- desenho
X0, X1 = 300000, 8450000
YT, YB = 4950000, 6420000
Y_BR = 4480000          # linha do colchete
LBL_D, LBL_T = 150000, 210000


def clean_slide(slide):
    to_remove = []
    for sh in slide.shapes:
        top = sh.top or 0
        if sh.shape_type in (MSO_SHAPE_TYPE.LINE, MSO_SHAPE_TYPE.CHART):
            to_remove.append(sh)
        elif sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and top >= 3900000:
            to_remove.append(sh)
    for sh in to_remove:
        el = sh._element
        rids = set(el.xpath('.//@r:id'))
        el.getparent().remove(el)
        for rid in rids:
            try:
                slide.part.drop_rel(rid)
            except KeyError:
                pass
    return len(to_remove)


def txt(shapes, x, y, w, h, text, size, bold=False, color=BLACK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM, name=None):
    tb = shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.word_wrap = False
    tf.vertical_anchor = anchor
    lines = text.split('\n')
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
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
        d = ln.makeelement(qn('a:prstDash'), {'val': dash})
        ln.append(d)
    if arrow:
        t = ln.makeelement(qn('a:tailEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'})
        ln.append(t)
    if name:
        c.name = name
    return c


def break_band(sh, x, bw, ym, d=40000, bh=90000, ov=7000):
    xl, xr = x - ov, x + bw + ov
    fb = sh.build_freeform(xl, ym + d)
    fb.add_line_segments([(xr, ym - d), (xr, ym - d + bh), (xl, ym + d + bh)], close=True)
    band = fb.convert_to_shape()
    band.fill.solid()
    band.fill.fore_color.rgb = RGBColor(255, 255, 255)
    band.line.fill.background()
    band.shadow.inherit = False
    band.name = 'Corte do eixo'
    l1 = line(sh, x, ym + d, x + bw, ym - d, BLACK, 0.75, name='Corte linha 1')
    l2 = line(sh, x, ym + d + bh, x + bw, ym - d + bh, BLACK, 0.75, name='Corte linha 2')
    return [band, l1, l2]


def draw_bridge(slide, spec, pct_text):
    start, deltas, end = spec['start'], spec['deltas'], spec['end']
    cols = [('Fcst R9', 'tot', start)] + [(lab, 'tgt' if lab == 'Target' else 'd', v) for lab, v in deltas] + [("MP’27", 'tot', end)]
    n = len(cols)
    slot = (X1 - X0) / n
    bw = int(slot * 0.62)

    # niveis
    bars, cum = [], 0.0
    for i, (lab, kind, v) in enumerate(cols):
        if kind == 'tot':
            lo, hi = min(0, v), max(0, v)
            cum = v
        else:
            lo, hi = min(cum, cum + v), max(cum, cum + v)
            cum += v
        bars.append((lo, hi, cum))
    lows = [b[0] for k, b in enumerate(bars) if cols[k][1] != 'tot']
    vtop = max(b[1] for b in bars)
    brk = None
    if CORTE_EIXO and lows and min(lows) > 0.5 and min(lows) > 0.1 * vtop:
        cand = round(min(lows) * 0.75, 2)
        mint = min(c[2] for c in cols if c[1] == 'tot')
        if (mint - cand) / (vtop - cand) * (YB - YT) >= 300000:
            brk = cand
    if brk:   # eixo truncado: totais comecam em 'brk' e levam o simbolo de corte
        for k, c in enumerate(cols):
            if c[1] == 'tot':
                bars[k] = (brk, bars[k][1], bars[k][2])
    vmax = max(max(b[1] for b in bars), 0)
    vmin = brk if brk else min(min(b[0] for b in bars), 0)
    yb_eff = YB - (200000 if vmin < 0 else 0)
    sc = (yb_eff - YT) / (vmax - vmin)
    y = lambda v: int(YT + (vmax - v) * sc)

    new = []
    sh = slide.shapes
    cx = lambda i: int(X0 + slot * (i + 0.5))
    # linha de base
    yz = y(brk) if brk else y(0)
    new.append(line(sh, X0, yz, X1, yz, GREY, 0.75, name='Eixo'))
    ym = None
    if brk:
        ym = int(yz - min(yz - y(bars[k][1]) for k, c in enumerate(cols) if c[1] == 'tot') / 2)
    prev_level = None
    for i, (lab, kind, v) in enumerate(cols):
        lo, hi, lvl = bars[i]
        x = cx(i) - bw // 2
        top, h = y(hi), max(y(lo) - y(hi), 12700)
        r = sh.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(top), Emu(bw), Emu(h))
        r.fill.solid()
        r.fill.fore_color.rgb = DARK if kind == 'tot' else (TARGET if kind == 'tgt' else LIGHT)
        r.line.color.rgb = BLACK
        r.line.width = Pt(0.75)
        r.shadow.inherit = False
        r.name = f'Barra {lab.replace(chr(10), " ")}'
        new.append(r)
        if brk and kind == 'tot':
            new.extend(break_band(sh, x, bw, ym))
        # rotulo do valor
        if kind == 'tot':
            below = v < 0
            ly = (y(0) + 25000) if below else (top - 25000 - LBL_T)
            new.append(txt(sh, cx(i) - int(slot / 2), ly, int(slot), LBL_T, f1(v), 14, True,
                           anchor=MSO_ANCHOR.TOP if below else MSO_ANCHOR.BOTTOM, name=f'Valor {lab}'))
        else:
            new.append(txt(sh, cx(i) - int(slot / 2), top - 25000 - LBL_D, int(slot), LBL_D, f1(v), 10, name=f'Valor {lab.replace(chr(10), " ")}'))
        # categoria
        new.append(txt(sh, cx(i) - int(slot / 2), 6500000, int(slot), 300000, lab, 9, anchor=MSO_ANCHOR.TOP,
                       name=f'Categoria {lab.replace(chr(10), " ")}'))
        # conector tracejado ate a proxima barra
        if i < n - 1:
            nxt_lvl = lvl
            new.append(line(sh, x + bw, y(nxt_lvl), cx(i + 1) - bw // 2, y(nxt_lvl), BLACK, 0.25, 'lgDash', name='Conector'))
    # colchete + bolha
    xl, xr = cx(0), cx(n - 1)

    def label_top(i):
        lo, hi, lvl = bars[i]
        v = cols[i][2]
        return (y(0) - 20000) if v < 0 else (y(hi) - 25000 - LBL_T - 30000)

    new.append(line(sh, xl, Y_BR, xl, label_top(0), GREY, 0.75, name='Colchete esq'))
    new.append(line(sh, xl, Y_BR, xr, Y_BR, GREY, 0.75, name='Colchete topo'))
    new.append(line(sh, xr, Y_BR, xr, label_top(n - 1), GREY, 0.75, arrow=True, name='Colchete dir'))
    ew, eh = (1010000 if len(pct_text) > 6 else 800000), 300000
    el = sh.add_shape(MSO_SHAPE.OVAL, Emu((xl + xr) // 2 - ew // 2), Emu(Y_BR - eh // 2), Emu(ew), Emu(eh))
    el.fill.solid()
    el.fill.fore_color.rgb = RGBColor(255, 255, 255)
    el.line.color.rgb = GREY
    el.line.width = Pt(0.75)
    el.shadow.inherit = False
    tf = el.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = pct_text
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = BLACK
    el.name = 'Variacao'
    new.append(el)

    grp = sh.add_group_shape(new)
    grp.name = 'Bridge MP27 (Base Book)'
    return grp


def mnus(t):
    return re.sub(r'(?<!\w)-(?=\d)', '−', t)


def draw_comments(slide, title, items):
    sh = slide.shapes
    px, pw = 8720000, 3230000
    py, ph = 520000, 5900000
    box = sh.add_shape(MSO_SHAPE.RECTANGLE, Emu(px), Emu(py), Emu(pw), Emu(ph))
    box.fill.solid()
    box.fill.fore_color.rgb = PANEL
    box.line.fill.background()
    box.shadow.inherit = False
    box.name = 'Painel comentarios'
    hd = sh.add_shape(MSO_SHAPE.RECTANGLE, Emu(px), Emu(py), Emu(pw), Emu(330000))
    hd.fill.solid()
    hd.fill.fore_color.rgb = HEAD
    hd.line.fill.background()
    hd.shadow.inherit = False
    tf = hd.text_frame
    tf.margin_left = Emu(90000)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = f"Principais mensagens – {title}"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = RGBColor(255, 255, 255)
    hd.name = 'Painel titulo'
    tb = sh.add_textbox(Emu(px + 90000), Emu(py + 400000), Emu(pw - 180000), Emu(ph - 450000))
    tb.name = 'Comentarios'
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    first = True
    for lead, body in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(9)
        r1_ = p.add_run()
        r1_.text = mnus(lead)
        r1_.font.size = Pt(11.5)
        r1_.font.bold = True
        r1_.font.color.rgb = HEAD
        p.add_line_break()
        r2 = p.add_run()
        r2.text = mnus(body)
        r2.font.size = Pt(11)
        r2.font.color.rgb = BLACK
    return [box, hd, tb]


def pct_text(spec, unit):
    s, e = spec['start'], spec['end']
    if s <= 0:
        return f"Δ +{f1(e - s)}" if e - s >= 0 else f"Δ {f1(e - s)}"
    p = (e - s) / abs(s) * 100
    v = Decimal(str(abs(p))).quantize(Decimal('0.1'), ROUND_HALF_UP)
    return f"{'+' if p >= 0 else '-'}{v}%".replace('.', ',')


UNITS = ['BR', 'IBI', 'SJP', 'GOI', 'RES']
prs = Presentation(SRC)
slides = list(prs.slides)
plan = {1: ('T', 0), 2: ('T', 1), 3: ('T', 2), 4: ('T', 3), 5: ('T', 4),
        7: ('N', 0), 8: ('N', 1), 9: ('N', 2), 10: ('N', 3), 11: ('N', 4)}   # indice 0-based do slide -> (modo, unidade)
for idx, (mode, ui) in plan.items():
    s = slides[idx]
    u = UNITS[ui]
    spec = DATA[mode][u]
    n_rem = clean_slide(s)
    # titulo "Result Variant" centralizado sobre o novo bridge, mesma altura em todos
    for shp in s.shapes:
        if shp.has_text_frame and shp.text_frame.text.strip() == 'Result Variant Fitted Units':
            shp.left = Emu(int((X0 + X1) / 2 - shp.width / 2))
            shp.top = Emu(3950000)
    draw_bridge(s, spec, pct_text(spec, u))
    title, items = bullets(u, mode)
    draw_comments(s, title, items)
    # dados do bridge nas notas (para colar no think-cell)
    tf = s.notes_slide.notes_text_frame
    rows = [('Fcst R9', spec['start'])] + [(l.replace('\n', ' '), v) for l, v in spec['deltas']] + [("MP'27", spec['end'])]
    extra = "Dados do bridge (Base_Book_Fitted Units_MP27.xlsx, " + ('com' if mode == 'T' else 'sem') + " target), mln BRL:\n" + \
            "\n".join(f"{l}\t{v:.3f}" for l, v in rows)
    tf.text = (tf.text + "\n\n" if tf.text.strip() else "") + extra
    print(idx + 1, u, mode, 'removidos', n_rem, [(l.replace('\n', ' '), f1(v)) for l, v in rows], pct_text(spec, u))

# ---------------------------------------------------------------- slides novos: mix do EBIT (pizzas) e rentabilidade por cliente IBI
import copy

from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

WHITE = RGBColor(255, 255, 255)
UCOL = {'IBI': RGBColor(0x36, 0x4D, 0x6E), 'SJP': RGBColor(0x5B, 0x8D, 0xB8), 'GOI': RGBColor(0x9B, 0xC2, 0xE6), 'RES': RGBColor(0xFF, 0xDD, 0x00)}
CCOL = {'Fiat': RGBColor(0x36, 0x4D, 0x6E), 'Iveco': RGBColor(0x9B, 0xC2, 0xE6), 'CNH': RGBColor(0xFF, 0xDD, 0x00)}
RED = RGBColor(0xC0, 0x00, 0x00)
LAYOUT = slides[1].slide_layout
FLAG = [s_ for s_ in slides[1].shapes if s_.name == 'Picture 163'][0]


def fp(x, d=1):
    """Formato de tabela da apresentacao: negativos entre parenteses, virgula decimal."""
    v = Decimal(str(abs(x))).quantize(Decimal('0.1') if d == 1 else Decimal('1'), ROUND_HALF_UP)
    t = f"{v:,.{d}f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    return f"({t})" if x < 0 and v != 0 else t


def pc(x, d=1):
    v = Decimal(str(abs(x * 100))).quantize(Decimal('0.1') if d == 1 else Decimal('1'), ROUND_HALF_UP)
    t = f"{v}".replace('.', ',')
    return f"-{t}%" if x < 0 and v != 0 else f"{t}%"


def new_slide(title):
    s_ = prs.slides.add_slide(LAYOUT)
    s_.shapes.title.text = title
    tidx = s_.shapes.title.placeholder_format.idx
    for ph in list(s_.placeholders):
        if ph.placeholder_format.idx != tidx:
            ph.element.getparent().remove(ph.element)
    el = copy.deepcopy(FLAG._element)
    blip = el.xpath('.//a:blip')[0]
    old = blip.get(qn('r:embed'))
    blip.set(qn('r:embed'), s_.part.relate_to(slides[1].part.related_part(old), RT.IMAGE))
    s_.shapes._spTree.append(el)
    return s_


def headline(s_, text):
    tb = s_.shapes.add_textbox(Emu(300000), Emu(560000), Emu(8150000), Emu(520000))
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


def footnote(s_, text, y=6420000):
    tb = s_.shapes.add_textbox(Emu(300000), Emu(y), Emu(8150000), Emu(330000))
    tb.name = 'Nota de rodape'
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    r = tf.paragraphs[0].add_run()
    r.text = text
    r.font.size = Pt(8)
    r.font.color.rgb = GREY
    return tb


def table(s_, x, y, colw, rows, rowh, font=10, bold_last=False, chips=None):
    n, m = len(rows), len(rows[0])
    gt = s_.shapes.add_table(n, m, Emu(x), Emu(y), Emu(sum(colw)), Emu(rowh * n))
    tb = gt.table
    tbl = gt._element.graphic.graphicData.tbl
    tbl.tblPr.set('bandRow', '0')
    tbl.tblPr.set('firstRow', '0')
    for j, wd in enumerate(colw):
        tb.columns[j].width = Emu(wd)
    for i, row in enumerate(rows):
        tb.rows[i].height = Emu(rowh)
        for j, val in enumerate(row):
            c = tb.cell(i, j)
            c.margin_left = c.margin_right = Emu(60000)
            c.margin_top = c.margin_bottom = Emu(8000)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.fill.solid()
            c.fill.fore_color.rgb = HEAD if i == 0 else (PANEL if i % 2 == 0 else WHITE)
            p = c.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT
            if chips and i > 0 and j == 0 and chips.get(i):
                rc = p.add_run()
                rc.text = '● '
                rc.font.size = Pt(font)
                rc.font.color.rgb = chips[i]
            r = p.add_run()
            r.text = str(val)
            r.font.size = Pt(font)
            r.font.bold = (i == 0) or (bold_last and i == n - 1)
            r.font.color.rgb = WHITE if i == 0 else BLACK
    return gt


def set_hole(chart, pct):
    for el in chart._chartSpace.xpath('.//c:holeSize'):
        el.set('val', str(pct))


def donut(s_, x, y, d, order, vals, big, small):
    tot = sum(vals)
    cd = CategoryChartData()
    cd.categories = order
    cd.add_series('EBIT', [max(v, 0) for v in vals])
    gf = s_.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Emu(x), Emu(y), Emu(d), Emu(d), cd)
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
        if share >= 0.09:
            tf = pt.data_label.text_frame
            tf.text = f"{round(share * 100)}%"
            r = tf.paragraphs[0].runs[0]
            r.font.size = Pt(14)
            r.font.bold = True
            r.font.color.rgb = BLACK if u in ('GOI', 'RES') else WHITE
    set_hole(ch, 62)
    cx = x + d // 2
    cy = y + d // 2
    txt(s_.shapes, cx - 700000, cy - 300000, 1400000, 330000, big, 24, True, color=HEAD, anchor=MSO_ANCHOR.BOTTOM, name='Centro valor')
    txt(s_.shapes, cx - 700000, cy + 40000, 1400000, 200000, small, 10, False, color=GREY, anchor=MSO_ANCHOR.TOP, name='Centro legenda')
    return gf


def pie_slide():
    s_ = new_slide("Fitted Units Economics Brasil – Mix do EBIT por unidade")
    order = ['IBI', 'SJP', 'GOI', 'RES']
    r9 = [DATA['T'][u]['start'] for u in order]
    mp = [DATA['T'][u]['end'] for u in order]
    t9, tm = sum(r9), sum(mp)
    sh9 = [v / t9 for v in r9]
    shm = [v / tm for v in mp]
    br9, brm = DATA['T']['BR']['start'], DATA['T']['BR']['end']
    top2 = shm[0] + shm[1]
    headline(s_, f"IBI e SJP concentram {round(top2 * 100)}% do EBIT do MP'27; a SJP perde {round((sh9[1] - shm[1]) * 100)} p.p. "
                 f"de peso e a RES passa a contribuir")
    D = 2850000
    X9, XM, Y = 700000, 4900000, 1480000
    txt(s_.shapes, X9, 1130000, D, 300000, "Fcst R9 – 2026", 14, True, color=HEAD, anchor=MSO_ANCHOR.MIDDLE, name='Titulo R9')
    txt(s_.shapes, XM, 1130000, D, 300000, "MP'27 – 2027", 14, True, color=HEAD, anchor=MSO_ANCHOR.MIDDLE, name='Titulo MP27')
    donut(s_, X9, Y, D, order, r9, f1(br9), 'EBIT Brasil')
    donut(s_, XM, Y, D, order, mp, f1(brm), 'EBIT Brasil')
    ar = s_.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Emu(3800000), Emu(2650000), Emu(950000), Emu(520000))
    ar.fill.solid()
    ar.fill.fore_color.rgb = TARGET
    ar.line.fill.background()
    ar.shadow.inherit = False
    ar.name = 'Seta R9 para MP27'
    txt(s_.shapes, 3700000, 3230000, 1150000, 260000, pct_text(DATA['T']['BR'], 'BR').replace('-', '−'), 16, True, color=RED, anchor=MSO_ANCHOR.MIDDLE, name='Variacao EBIT')
    txt(s_.shapes, 3700000, 3480000, 1150000, 200000, f"Δ {f1(brm - br9)} mln BRL".replace('-', '−'), 10, False, color=GREY, anchor=MSO_ANCHOR.TOP, name='Variacao absoluta')
    rows = [["EBIT (mln BRL)", "Fcst R9", "% R9", "MP'27", "% MP'27", "Δ EBIT", "Δ p.p."]]
    for k, u in enumerate(order):
        rows.append([u, fp(r9[k]), pc(sh9[k]), fp(mp[k]), pc(shm[k]), fp(mp[k] - r9[k]), f"{(shm[k] - sh9[k]) * 100:+.1f}".replace('.', ',').replace('-', '−')])
    sor = br9 - t9
    rows.append(["SOR (outras receitas)", fp(sor), '–', fp(0.0), '–', fp(-sor), '–'])
    rows.append(["Brasil", fp(br9), '', fp(brm), '', fp(brm - br9), ''])
    chips = {1 + k: UCOL[u] for k, u in enumerate(order)}
    table(s_, 300000, 4650000, [1900000, 1000000, 950000, 1000000, 1050000, 1150000, 1100000], rows, 235000, 10, bold_last=True, chips=chips)
    footnote(s_, "Fonte: Base Book Fitted Units MP27 e MENS FITTED MP27 v3. Pizzas consideram as 4 unidades (R9: 12,8; MP'27: 9,2); o R9 do Brasil (13,0) inclui "
                 f"{f1(sor)} de outras receitas da SOR, encerrada em 2027. RES no R9 = {f1(r9[3])} ({pc(sh9[3])} do EBIT), sem fatia na pizza.", 6430000)
    items = [
        ('IBI segue como âncora', f"{round(sh9[0] * 100)}% → {round(shm[0] * 100)}% do EBIT, mesmo com o EBIT caindo de {f1(r9[0])} para {f1(mp[0])}."),
        ('SJP perde peso', f"{round(sh9[1] * 100)}% → {round(shm[1] * 100)}% do EBIT ({f1(r9[1])} → {f1(mp[1])}): a maior queda relativa do Brasil ({pc((mp[1] - r9[1]) / r9[1], 0)})."),
        ('GOI estável e com mais peso', f"EBIT de {f1(mp[2])} (igual ao R9); participação sobe de {round(sh9[2] * 100)}% para {round(shm[2] * 100)}%."),
        ('RES passa a contribuir', f"EBIT de {f1(r9[3])} para {f1(mp[3])}; chega a {round(shm[3] * 100)}% do EBIT."),
        ('Concentração de risco', f"IBI + SJP = {round(top2 * 100)}% do EBIT. O target (7,6) sustenta 82% do EBIT do Brasil; sem ele o EBIT seria 1,6."),
    ]
    draw_comments(s_, "Mix do EBIT", items)
    return s_


def client_slide():
    wbc = openpyxl.load_workbook(CLI, data_only=True)
    ws = wbc['Rentabilidade por Cliente']
    cl = ['Fiat', 'Iveco', 'CNH']

    def row(r, k=1000.0):
        return [ws.cell(r, c).value for c in (2, 3, 4, 5)]
    pcs, ns, nspc, cm, fix, ebit, ros, ebpc = row(5), row(6), row(7), row(14), row(16), row(23), row(24), row(25)
    assert abs(ebit[3] - 5391.16) < 1 and abs(sum(ebit[:3]) - ebit[3]) < 0.5
    sh_pc = [pcs[i] / pcs[3] for i in range(3)]
    sh_ns = [ns[i] / ns[3] for i in range(3)]
    sh_eb = [ebit[i] / ebit[3] for i in range(3)]
    frete = -ws.cell(12, 2).value
    s_ = new_slide("Fitted Units Economics Brasil – IBI | Rentabilidade por cliente")
    headline(s_, f"A Fiat é {round(sh_ns[0] * 100)}% do faturamento de Ibirité, mas gera só {round(sh_eb[0] * 100)}% do EBIT; "
                 f"Iveco e CNH, com {round((sh_ns[1] + sh_ns[2]) * 100)}% do faturamento, entregam {round((sh_eb[1] + sh_eb[2]) * 100)}%")
    txt(s_.shapes, 300000, 1130000, 4000000, 260000, "Peso de cada cliente (% do total IBI)", 11, True, color=HEAD, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, name='Titulo grafico peso')
    txt(s_.shapes, 4500000, 1130000, 3950000, 260000, "ROS por cliente (EBIT / Net Sales)", 11, True, color=HEAD, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, name='Titulo grafico ROS')
    # 100% empilhado
    cd = CategoryChartData()
    cd.categories = ['Pieces', 'Net Sales', 'EBIT']
    for i, c in enumerate(cl):
        cd.add_series(c, [sh_pc[i], sh_ns[i], sh_eb[i]])
    gf = s_.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED_100, Emu(300000), Emu(1400000), Emu(4000000), Emu(2550000), cd)
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
    for i, c in enumerate(cl):
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
    # ROS
    cd2 = CategoryChartData()
    cd2.categories = ['Fiat', 'Iveco', 'CNH', 'Total IBI']
    cd2.add_series('ROS', ros)
    gf2 = s_.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Emu(4500000), Emu(1400000), Emu(3950000), Emu(2550000), cd2)
    gf2.name = 'ROS por cliente'
    c2 = gf2.chart
    c2.has_title = False
    c2.has_legend = False
    c2.font.size = Pt(10)
    c2.value_axis.visible = False
    c2.value_axis.has_major_gridlines = False
    c2.value_axis.maximum_scale = 1.1
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
    # tabela
    k1 = lambda v: fp(v / 1000.0)
    rows = [["mln BRL", "Fiat", "Iveco", "CNH", "Total IBI"],
            ["Pieces (/000)"] + [fp(v, 0) for v in pcs],
            ["Net Sales"] + [k1(v) for v in ns],
            ["NS/pc (R$)"] + [fp(v) for v in nspc],
            ["CM"] + [k1(v) for v in cm],
            ["Fixed Cost"] + [k1(v) for v in fix],
            ["EBIT"] + [k1(v) for v in ebit],
            ["ROS %"] + [pc(v) for v in ros],
            ["EBIT/pc (R$)"] + [fp(v) for v in ebpc],
            ["% do EBIT"] + [pc(v) for v in sh_eb] + ['100,0%']]
    table(s_, 300000, 4080000, [2150000, 1500000, 1500000, 1500000, 1500000], rows, 222000, 10, bold_last=False)
    footnote(s_, "Fonte: MENS FITTED MP27 v3 (aba IBI), MP'27 em mln BRL. Faturamento e material de consumo são diretos por cliente; demais custos rateados por volume "
                 "(frete 100% Fiat). O EBIT inclui o target de eficiência (5,8), rateado por volume. Margens de Iveco e CNH são indicativas: custos específicos ainda não abertos.", 6340000)
    items = [
        ('Fiat: volume, não margem', f"{round(sh_pc[0] * 100)}% das peças e {round(sh_ns[0] * 100)}% do Net Sales, mas só {round(sh_eb[0] * 100)}% do EBIT: ROS de {pc(ros[0])} e EBIT/pç de R$ {fp(ebpc[0])} (média IBI: R$ {fp(ebpc[3])})."),
        ('Iveco e CNH: nichos de alto valor', f"{round((sh_ns[1] + sh_ns[2]) * 100)}% do Net Sales e {round((sh_eb[1] + sh_eb[2]) * 100)}% do EBIT (ROS de {round(ros[1] * 100)}% e {round(ros[2] * 100)}%). NS/pç de R$ {fp(nspc[1])} e R$ {fp(nspc[2])} vs R$ {fp(nspc[0])} da Fiat."),
        ('Frete pressiona a Fiat', f"Todo o frete (R$ {fp(frete / 1000)} mi) é Fiat: {round(frete / ns[0] * 100)}% do Net Sales do cliente. A nova regra da ANTT (-1,5 no bridge) atinge o cliente de menor margem."),
        ('Jaguar fora da base', "Saída da JLR já refletida: o R9 tinha 1,1k pçs e R$ 1,5 mi de faturamento (margem direta de ~R$ 0,5 mi)."),
        ('Leitura com cautela', "Rateio por volume: Iveco e CNH podem ter custos específicos fora da base. Validar com a área de custos antes de usar como preço-alvo."),
    ]
    draw_comments(s_, "Clientes IBI", items)
    return s_


pieT, cliT, pieN, cliN = pie_slide(), client_slide(), pie_slide(), client_slide()
lst = prs.slides._sldIdLst
ids = list(lst)
o, new = ids[:12], ids[12:]
final = [o[0], o[1], new[0], o[2], new[1], o[3], o[4], o[5], o[6], o[7], new[2], o[8], new[3], o[9], o[10], o[11]]
for el in ids:
    lst.remove(el)
for el in final:
    lst.append(el)


prs.save(OUT)
print('salvo', OUT)
