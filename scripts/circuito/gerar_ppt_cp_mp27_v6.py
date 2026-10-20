"""Deck MP'27 Circuito Panamericano (v6) – resumo para o presidente, receita por origem, custo fixo,
postergações de 2026 e perguntas-chave.

Uso:
    python gerar_ppt_cp_mp27_v6.py <Bridge CP_v2.pptx> <Base_Book_CP_.xlsx> <Faturamento CP - MP'27.xlsx>
                                   <Base de Custo MP'27_v2.xlsx> <saida.pptx>

v6 (pedido da Juliana, 19/10, rascunho "CLAUDE OLHAR AQUI"):
- slide 1: tabela refeita (nativa, valores do Base Book atual) + bridge por origem da receita e driver de custo;
- slide 2: R&D Tests (intercompany) | Terceiros (eventos + OEM + aluguéis) | Total; mensal de eventos e mix;
- slide 3: drivers do custo fixo pela aba Sheet1 da Base de Custo (sem o título "Leitura");
- slide 4: postergações de 2026 = só os 5 itens que continuam no MP'27 (411);
- slide 5: perguntas-chave do CEO.
"""
import sys
from lxml import etree
import openpyxl
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_AXIS_CROSSES

SRC, BB, FAT, BC, OUT = sys.argv[1:6]

NAVY = RGBColor(0x1F, 0x3A, 0x5F)
MIDBLUE = RGBColor(0x2E, 0x5A, 0xA8)
LIGHT = RGBColor(0xDC, 0xE3, 0xF0)
PANEL = RGBColor(0xF3, 0xF5, 0xFA)
EDGE = RGBColor(0x3C, 0x4B, 0x66)
SKY = RGBColor(0x9D, 0xC3, 0xE6)
GREY = RGBColor(0x80, 0x80, 0x80)
DGREY = RGBColor(0x40, 0x40, 0x40)
BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
RED = RGBColor(0xC6, 0x28, 0x28)
HIST = [RGBColor(0xBF, 0xBF, 0xBF), RGBColor(0xA6, 0xA6, 0xA6), RGBColor(0xC9, 0xD6, 0xEA),
        RGBColor(0x8E, 0xA9, 0xCF), NAVY]                       # FY24, FY25, MP26, R9, MP27
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"


def r0(v):
    """arredondamento comercial (meio para cima); 9.619,49998 do base book vira 9.620."""
    v = round(v, 2)
    return int(v + 0.5) if v >= 0 else -int(-v + 0.5)


# ------------------------------------------------------------------ dados: Base Book
bb = openpyxl.load_workbook(BB, data_only=True)
m = bb["Month_FY_LC_"]
SCOL = "DEFGH"                                   # Actual 2024, Actual 2025, MP26, Fcst R9, MP'27


def row(r):
    return [m[c + str(r)].value or 0 for c in SCOL]


net, ev, oem, rent, rd, evg = row(11), row(12), row(13), row(14), row(15), row(16)
fix, gm, adm, ops, mkt, csa, da = row(21), row(22), row(23), row(24), row(25), row(26), row(27)
ebit = row(31)
r9, mp, mp26, fy25 = 3, 4, 2, 1
assert abs(sum(x[mp] for x in (ev, oem, rent, rd, evg)) - net[mp]) < 0.5
assert abs(sum(x[mp] for x in (gm, adm, ops, mkt, csa, da)) - fix[mp]) < 0.5
assert abs(net[mp] + fix[mp] - ebit[mp]) < 0.5

# ------------------------------------------------------------------ dados: faturamento mensal (bruto)
fat = openpyxl.load_workbook(FAT, data_only=True)["Volume MP27"]
MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
ev_mp_m = [(fat.cell(4, c).value or 0) / 1e3 for c in range(2, 14)]
ev_r9_m = [(fat.cell(13, c).value or 0) / 1e3 for c in range(2, 14)]
rd_mp_m = (fat.cell(7, 2).value or 0) / 1e3
rd_r9_m = (fat.cell(16, 2).value or 0) / 1e3

# ------------------------------------------------------------------ dados: justificativas (Base de Custo, Sheet1)
sh = openpyxl.load_workbook(BC, data_only=True)["Sheet1"]
just = {}
for r in range(13, 26):
    k, v, t = sh.cell(r, 3).value, sh.cell(r, 4).value, sh.cell(r, 5).value
    if k and isinstance(v, (int, float)):
        just[str(k).strip()] = (v / 1e3, (t or "").strip().lower())
POST = [("Operacional", "Destinação do resíduo (Sansão)", just["Cercamento | Manutenções - Sansão"][0]),
        ("Manut. Geral", "Restauração APP", just["Restauração APP"][0]),
        ("Operacional", "Manutenção corretiva elétrica", just["Elétrica"][0]),
        ("Marketing", "Ações de marketing", just["Ações de marketing"][0]),
        ("Operacional", "Análise qualidade do ar (ar-condicionado)",
         just["Análise qualidade do ar (ar-condicionado)"][0])]
post = sum(v for _, _, v in POST)
assert abs(post - sum(v for v, t in just.values() if "posterga" in t)) < 0.01
HSE = [("Renovação licenças ambientais de operação + outorga",
        just["RENOVAÇÃO DAS LICENÇAS DE AMBIENTAL DE OPERAÇÃO + Outorga"][0]),
       ("Coleta de lixo orgânico | biodigestor", just["Coleta de Lixo - Orgânico |Biodigestor"][0])]
hse = sum(v for _, v in HSE)
portaria = just["Portaria + Projeto"][0]
infl_ops = just["Reajuste de inflação"][0]

d_fix = -(fix[mp] - fix[r9])
d_gm, d_adm, d_ops = -(gm[mp] - gm[r9]), -(adm[mp] - adm[r9]), -(ops[mp] - ops[r9])
d_mkt, d_csa, d_da = -(mkt[mp] - mkt[r9]), -(csa[mp] - csa[r9]), -(da[mp] - da[r9])
ops_expl = post + hse + portaria + infl_ops                  # Sheet1 explica a estrutura operacional
ops_res = d_ops - ops_expl                                    # resíduo da estrutura operacional
infl = infl_ops + d_gm + d_csa                                # "Inflação" na Sheet1: contratos, manut. geral, CSC
outros = d_fix - post - hse - infl - d_adm - portaria         # marketing (revisão), D&A e resíduo
assert abs(outros - (d_mkt + d_da + ops_res)) < 0.5
DRIVERS = [("Postergações\nde 2026", post), ("Obrigações\nHSE", hse), ("Inflação", infl), ("RH", d_adm),
           ("Projeto\nportaria", portaria), ("Outros", outros)]
assert abs(sum(v for _, v in DRIVERS) - d_fix) < 0.5

# ------------------------------------------------------------------ números-chave
terc = [a + b + c + d for a, b, c, d in zip(ev, oem, rent, evg)]      # terceiros (tudo menos R&D)
ev_tot = [a + b for a, b in zip(ev, evg)]
oem_tot = [a + b for a, b in zip(oem, rent)]
rd_chart = list(rd)
rd_chart[r9] = rd[mp26]                    # R9 = MP26 (9.620); base book 9.619,5 (arredondamento)
d_net = net[mp] - net[r9]
d_rd, d_terc = rd[mp] - rd[r9], terc[mp] - terc[r9]
ros_r9, ros_mp = ebit[r9] / net[r9] * 100, ebit[mp] / net[mp] * 100
ebit_sem_post = ebit[mp] + post
ros_sem = ebit_sem_post / net[mp] * 100
ebitda = ebit[mp] - da[mp]
pontual = post + hse
recorrente = d_fix - pontual
sens_ev = ev_tot[mp] * 0.10                # eventos: margem de contribuição de 100%
ebit_sem_reaj = ebit[mp] - d_rd
r9_2s = sum(ev_r9_m[8:]) / sum(ev_r9_m) * 100


def fmt(v, sign=False, paren=False):
    s = f"{abs(r0(v)):,}".replace(",", ".")
    if r0(v) < 0:
        return f"({s})" if paren else "-" + s
    return ("+" if sign and r0(v) > 0 else "") + s


def pct(a, b):
    return (a / b - 1) * 100


def p1(x, sign=True):
    return (f"{x:+.1f}" if sign else f"{x:.1f}").replace(".", ",")


# ------------------------------------------------------------------ helpers de desenho
def line(sl, x1, y1, x2, y2, color=GREY, w=0.75, dash=None, arrow=False):
    c = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Pt(x1), Pt(y1), Pt(x2), Pt(y2))
    c.line.color.rgb = color
    c.line.width = Pt(w)
    if dash:
        c.line.dash_style = dash
    if arrow:
        etree.SubElement(c.line._get_or_add_ln(), A + "tailEnd").set("type", "triangle")
    return c


def text(sl, x, y, w, h, runs, size=9, color=BLACK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         bold=False, italic=False, spacing=1.0, after=0):
    """runs: str (\\n = novo parágrafo) ou lista de parágrafos; parágrafo = lista de (texto, {fmt})."""
    tb = sl.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    paras = [[(ln, {})] for ln in runs.split("\n")] if isinstance(runs, str) else runs
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        p.space_before = Pt(0)
        p.space_after = Pt(after)
        for t, f in para:
            r = p.add_run()
            r.text = t
            r.font.size = Pt(f.get("size", size))
            r.font.bold = f.get("bold", bold)
            r.font.italic = f.get("italic", italic)
            r.font.color.rgb = f.get("color", color)
    return tb


def rect(sl, x, y, w, h, fill, edge=None, ew=0.75, shape=MSO_SHAPE.RECTANGLE):
    r = sl.shapes.add_shape(shape, Pt(x), Pt(y), Pt(w), Pt(h))
    if fill is None:
        r.fill.background()
    else:
        r.fill.solid()
        r.fill.fore_color.rgb = fill
    if edge is None:
        r.line.fill.background()
    else:
        r.line.color.rgb = edge
        r.line.width = Pt(ew)
    r.shadow.inherit = False
    return r


def headline(sl, main, accent, y=62, size=16):
    text(sl, 36, y, 900, 40, [[(main, {"bold": True, "color": NAVY, "size": size}),
                               (accent, {"bold": True, "color": MIDBLUE, "size": size})]])


def break_mark(sl, c, bw, axis):
    top = axis - 11
    f = sl.shapes.build_freeform(Pt(c - bw / 2 - 3), Pt(top + 5))
    f.add_line_segments([(Pt(c + bw / 2 + 3), Pt(top - 3)), (Pt(c + bw / 2 + 3), Pt(top + 1)),
                         (Pt(c - bw / 2 - 3), Pt(top + 9))], close=True)
    b = f.convert_to_shape()
    b.fill.solid()
    b.fill.fore_color.rgb = WHITE
    b.line.fill.background()
    b.shadow.inherit = False


def waterfall(sl, x0, x1, axis, room, start, steps, end_label, start_label, highlight=(), tot_h=18,
              lab_size=9, good_up=True, bracket=None):
    """Bridge em formas nativas. steps = [(rótulo, delta)]; totais com eixo truncado."""
    n = len(steps) + 2
    dx = (x1 - x0) / n
    bw = min(46.0, dx * 0.66)
    cx = [x0 + dx * (i + 0.5) for i in range(n)]
    cum, lo, hi = 0, 0, 0
    for _, v in steps:
        cum += v
        lo, hi = min(lo, cum), max(hi, cum)
    sc = (room - tot_h - 24) / max(hi - lo, 1)
    base = axis - tot_h + lo * sc if lo < 0 else axis - tot_h            # nível do total inicial
    line(sl, x0, axis, x1, axis, color=BLACK, w=0.75)
    rect(sl, cx[0] - bw / 2, base, bw, axis - base, NAVY)
    text(sl, cx[0] - 30, base - 19, 60, 16, fmt(start), size=13, bold=True, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.BOTTOM)
    level = base
    min_top = base
    for i, (lab, v) in enumerate(steps, start=1):
        new = level - v * sc
        top, bot = min(level, new), max(level, new)
        min_top = min(min_top, top)
        hl = lab in highlight
        rect(sl, cx[i] - bw / 2, top, bw, max(bot - top, 1.2), LIGHT, MIDBLUE if hl else EDGE, 1.75 if hl else 0.75)
        line(sl, cx[i - 1] + bw / 2, level, cx[i] - bw / 2, level, color=BLACK, w=0.25, dash=MSO_LINE.LONG_DASH)
        text(sl, cx[i] - 30, top - 14, 60, 12, fmt(v, sign=True), size=lab_size, bold=True,
             color=BLACK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM)
        level = new
    end = start + sum(v for _, v in steps)
    rect(sl, cx[-1] - bw / 2, level, bw, axis - level, NAVY)
    line(sl, cx[-2] + bw / 2, level, cx[-1] - bw / 2, level, color=BLACK, w=0.25, dash=MSO_LINE.LONG_DASH)
    text(sl, cx[-1] - 30, level - 19, 60, 16, fmt(end), size=13, bold=True, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.BOTTOM)
    break_mark(sl, cx[0], bw, axis)
    break_mark(sl, cx[-1], bw, axis)
    labels = [start_label] + [l for l, _ in steps] + [end_label]
    for c, lab in zip(cx, labels):
        hl = lab in highlight
        text(sl, c - dx / 2, axis + 4, dx, 26, lab, size=8.5, bold=hl or lab in (start_label, end_label),
             color=MIDBLUE if hl else BLACK, align=PP_ALIGN.CENTER)
    if bracket:
        ytop = min(min_top, level) - 32
        line(sl, cx[0], ytop, cx[0], base - 22)
        line(sl, cx[0], ytop, cx[-1], ytop)
        line(sl, cx[-1], ytop, cx[-1], level - 22, arrow=True)
        ov = rect(sl, (cx[0] + cx[-1]) / 2 - 32, ytop - 10, 64, 20, WHITE, GREY, shape=MSO_SHAPE.OVAL)
        tf = ov.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = bracket
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = BLACK
    return end


def style_chart(ch, size=8):
    ch.font.size = Pt(size)
    ch.font.color.rgb = DGREY
    va = ch.value_axis
    va.visible = False
    va.has_major_gridlines = False
    ca = ch.category_axis
    ca.tick_labels.font.size = Pt(size)
    ca.format.line.color.rgb = GREY
    ca.has_major_gridlines = False


def badge(sl, x, y, d, label="vs R9"):
    rect(sl, x, y, 64, 18, LIGHT, None, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(sl, x, y, 64, 18, f"{p1(d)}% {label}", size=8.5, bold=True,
         color=GREEN if d >= 0 else RED, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def small_multiple(sl, x, y, w, h, title, vals, sub, ymax):
    text(sl, x, y, w - 66, 16, title, size=11, bold=True, color=NAVY)
    text(sl, x, y + 15, w, 12, sub, size=8, italic=True, color=GREY)
    cd = CategoryChartData()
    cd.categories = ["FY24", "FY25", "MP26", "R9", "MP’27"]
    cd.add_series(title, [r0(v) for v in vals])
    gf = sl.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Pt(x), Pt(y + 28), Pt(w), Pt(h - 28), cd)
    ch = gf.chart
    ch.has_legend = False
    ch.has_title = False
    style_chart(ch)
    plot = ch.plots[0]
    plot.gap_width = 45
    plot.has_data_labels = True
    dl = plot.series[0].data_labels
    dl.show_value = True
    dl.number_format = '#,##0'
    dl.number_format_is_linked = False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size = Pt(8)
    for i, pt in enumerate(plot.series[0].points):
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = HIST[i]
    ch.value_axis.minimum_scale = 0
    ch.value_axis.maximum_scale = ymax
    badge(sl, x + w - 64, y, pct(vals[4], vals[3]))


def table(sl, x, y, col_w, rows, row_h=19, header=None, size=9, total_last=False, col_align=None):
    nr = len(rows) + (1 if header else 0)
    gt = sl.shapes.add_table(nr, len(col_w), Pt(x), Pt(y), Pt(sum(col_w)), Pt(row_h * nr))
    tbl = gt.table
    pr = gt._element.graphic.graphicData.tbl.tblPr
    pr.set("firstRow", "0")
    pr.set("bandRow", "0")
    sid = pr.find(A + "tableStyleId")
    if sid is not None:
        sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
    for i, w in enumerate(col_w):
        tbl.columns[i].width = Pt(w)
    col_align = col_align or [PP_ALIGN.LEFT] * (len(col_w) - 1) + [PP_ALIGN.RIGHT]
    allrows = ([header] if header else []) + rows
    for r_, vals in enumerate(allrows):
        tbl.rows[r_].height = Pt(row_h)
        is_h = header and r_ == 0
        is_t = total_last and r_ == nr - 1
        for c_, v in enumerate(vals):
            ce = tbl.cell(r_, c_)
            ce.margin_top = ce.margin_bottom = Pt(1)
            ce.margin_left = ce.margin_right = Pt(5)
            ce.vertical_anchor = MSO_ANCHOR.MIDDLE
            ce.text = ""
            p = ce.text_frame.paragraphs[0]
            p.alignment = col_align[c_]
            t, f = (v if isinstance(v, tuple) else (v, {}))
            rr = p.add_run()
            rr.text = str(t)
            rr.font.size = Pt(f.get("size", size))
            rr.font.bold = f.get("bold", bool(is_h or is_t))
            rr.font.italic = f.get("italic", False)
            rr.font.color.rgb = f.get("color", WHITE if is_h else BLACK)
            if is_h:
                ce.fill.solid()
                ce.fill.fore_color.rgb = NAVY
            elif is_t:
                ce.fill.solid()
                ce.fill.fore_color.rgb = LIGHT
            else:
                ce.fill.background()
        if not (is_h or is_t):
            for c_ in range(len(col_w)):
                tcPr = tbl.cell(r_, c_)._tc.get_or_add_tcPr()
                ln = etree.SubElement(tcPr, A + "lnB", w="3175")
                sf = etree.SubElement(ln, A + "solidFill")
                etree.SubElement(sf, A + "srgbClr", val="D9D9D9")
    return gt


def cell_border(ce, side, color="FFFFFF", w=12700):
    tcPr = ce._tc.get_or_add_tcPr()
    for old in tcPr.findall(A + side):
        tcPr.remove(old)
    ln = etree.SubElement(tcPr, A + side, w=str(w))
    sf = etree.SubElement(ln, A + "solidFill")
    etree.SubElement(sf, A + "srgbClr", val=color)


def new_slide(prs, layout, title):
    s = prs.slides.add_slide(layout)
    for ph in list(s.placeholders):
        if "TITLE" in str(ph.placeholder_format.type):
            ph.text_frame.text = title
        else:
            ph._element.getparent().remove(ph._element)
    return s


def gain_label(sl):
    text(sl, 10, 52, 200, 14, "(+) gain  |  BRL ’000", size=9, italic=True)


def panel_blocks(sl, x, y, w, h, title, items, size=10):
    """Painel com blocos (título em negrito + texto) distribuídos igualmente na altura – sem sobra."""
    rect(sl, x, y, w, h, PANEL, LIGHT)
    rect(sl, x, y, 4, h, MIDBLUE)
    text(sl, x + 16, y + 10, w - 28, 20, title, size=13, bold=True, color=NAVY)
    top = y + 38
    bh = (y + h - 8 - top) / len(items)
    for i, (head, body) in enumerate(items):
        by = top + i * bh
        if i:
            line(sl, x + 16, by, x + w - 14, by, color=RGBColor(0xC9, 0xD3, 0xE6), w=0.5)
        text(sl, x + 16, by + 4, w - 30, bh - 8,
             [[(head, {"bold": True, "color": NAVY, "size": size + 0.5})],
              [(body, {"color": DGREY, "size": size})]],
             anchor=MSO_ANCHOR.MIDDLE, spacing=1.08, after=3)


# ================================================================== slide 1 – resumo
prs = Presentation(SRC)
s1 = prs.slides[0]
keep = {"Object 4", "Title 2"}
for i, shp in enumerate(list(s1.shapes)):
    # mantém o objeto think-cell, o título, "(+) gain" e o texto vertical; sai bridge antigo, ovais e imagem da tabela
    if shp.name in keep or (shp.has_text_frame and shp.text_frame.text.startswith(("(+) gain", "Result Variant"))):
        continue
    shp._element.getparent().remove(shp._element)

# --- tabela P&L (nativa, Base Book)
LBL = [("Net Sales", net, "main"), ("Events sales", ev, "sub"), ("OEMs’ technical tests", oem, "sub"),
       ("Long-term rents", rent, "sub"), ("R&D Tests", rd, "sub"), ("Events sales group", evg, "sub"),
       ("CM", net, "cm"),
       ("Fixed Cost", fix, "main"), ("General maintenance", gm, "sub"), ("Administrative costs", adm, "sub"),
       ("Operational structure", ops, "sub"), ("Marketing and promotions", mkt, "sub"),
       ("Costs CSA/CSC", csa, "sub"), ("D&A", da, "sub"),
       ("EBIT", ebit, "main"), ("Ros", None, "ros")]
CW = [118, 72, 72, 72, 72, 72, 10, 64, 64, 64]
HDR = ["", "Actual 2024", "Actual 2025", "MP26", "Fcst R9", "MP’27", "", "vs. R9", "vs. MP’26", "vs. PY’25"]
nrows = 2 + len(LBL) + 1
TX, TY = 14, 40
heights = [18, 15] + [17 if k in ("main", "cm") else (12.5 if k == "sub" else 14) for _, _, k in LBL] + [17]
gt = s1.shapes.add_table(nrows, len(CW), Pt(TX), Pt(TY), Pt(sum(CW)), Pt(sum(heights)))
tbl = gt.table
pr = gt._element.graphic.graphicData.tbl.tblPr
pr.set("firstRow", "0")
pr.set("bandRow", "0")
sid = pr.find(A + "tableStyleId")
if sid is not None:
    sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
for i, w in enumerate(CW):
    tbl.columns[i].width = Pt(w)
for i, h in enumerate(heights):
    tbl.rows[i].height = Pt(h)


def put(r_, c_, t, size=8, bold=False, color=BLACK, align=PP_ALIGN.RIGHT, fill=None, italic=False):
    ce = tbl.cell(r_, c_)
    ce.margin_top = ce.margin_bottom = Pt(0)
    ce.margin_left = Pt(3)
    ce.margin_right = Pt(6)
    ce.vertical_anchor = MSO_ANCHOR.MIDDLE
    ce.text = ""
    p = ce.text_frame.paragraphs[0]
    p.alignment = align
    rr = p.add_run()
    rr.text = t
    rr.font.size = Pt(size)
    rr.font.bold = bold
    rr.font.italic = italic
    rr.font.color.rgb = color
    etree.SubElement(p._p, A + "endParaRPr").set("sz", str(int(size * 100)))
    if fill is None:
        ce.fill.background()
    else:
        ce.fill.solid()
        ce.fill.fore_color.rgb = fill
    return ce


for c_ in range(len(CW)):
    put(0, c_, "", fill=None)
    put(1, c_, "", fill=None)
c = tbl.cell(0, 1)
c.merge(tbl.cell(0, 5))
put(0, 1, "Full Year", size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, fill=NAVY)
c = tbl.cell(0, 7)
c.merge(tbl.cell(0, 9))
put(0, 7, "Δ", size=11, bold=True, color=NAVY, align=PP_ALIGN.CENTER, fill=SKY)
for c_, h in enumerate(HDR):
    if c_ in (0, 6):
        continue
    put(1, c_, h, size=8.5, color=WHITE if c_ <= 5 else NAVY, align=PP_ALIGN.CENTER,
        fill=NAVY if c_ <= 5 else SKY, bold=(c_ == 5))
MPFILL = RGBColor(0xEE, 0xF2, 0xF9)
for i, (lab, vals, kind) in enumerate(LBL):
    r_ = 2 + i
    big = kind in ("main", "cm")
    sz = 10 if kind == "main" else (9.5 if kind == "cm" else 7.5)
    put(r_, 0, lab, size=sz, bold=kind == "main", align=PP_ALIGN.RIGHT)
    put(r_, 6, "")
    if kind == "ros":
        for k in range(5):
            put(r_, 1 + k, p1(ebit[k] / net[k] * 100, False) + "%", size=8.5, align=PP_ALIGN.CENTER,
                fill=MPFILL if k == 4 else None, bold=k == 4)
        d_ros = ebit[mp] / net[mp] * 100
        for k, ref in enumerate((r9, mp26, fy25)):
            put(r_, 7 + k, p1(d_ros - ebit[ref] / net[ref] * 100) + " p.p.", size=7.5, color=GREY,
                align=PP_ALIGN.CENTER)
        continue
    for k in range(5):
        put(r_, 1 + k, fmt(vals[k], paren=True), size=sz, bold=kind == "main" or (k == 4 and big),
            align=PP_ALIGN.CENTER, fill=MPFILL if k == 4 else None)
    for k, ref in enumerate((r9, mp26, fy25)):
        dv = vals[mp] - vals[ref]
        put(r_, 7 + k, fmt(dv, paren=True), size=sz, bold=kind == "main",
            color=(NAVY if kind == "main" else GREY), align=PP_ALIGN.CENTER)
# linha Δ EBIT (variação ano a ano, com sinal colorido)
rl = nrows - 1
put(rl, 0, "Δ EBIT", size=10, bold=True)
put(rl, 1, "")
for k in range(1, 5):
    dv = ebit[k] - ebit[k - 1]
    ce = put(rl, 1 + k, "", align=PP_ALIGN.CENTER)
    p = ce.text_frame.paragraphs[0]
    r1 = p.add_run()
    r1.text = "● "
    r1.font.size = Pt(9)
    r1.font.color.rgb = GREEN if dv >= 0 else RED
    r2 = p.add_run()
    r2.text = fmt(dv, paren=True)
    r2.font.size = Pt(9)
    r2.font.bold = k == 4
    r2.font.color.rgb = BLACK
for c_ in range(6, 10):
    put(rl, c_, "")
# bordas: linhas brancas entre colunas do cabeçalho; linha fina sob Net Sales/CM/Fixed Cost/D&A/EBIT
for r_ in range(nrows):
    for c_ in range(len(CW)):
        ce = tbl.cell(r_, c_)
        for side in ("lnL", "lnR", "lnT", "lnB"):
            cell_border(ce, side, "FFFFFF", 0)
SEP = {2 + 6, 2 + 7, 2 + 14, 2 + 15}          # antes de CM, Fixed Cost, EBIT, Ros
for r_ in range(2, nrows - 1):
    for c_ in list(range(1, 6)) + list(range(7, 10)):
        if r_ in SEP:
            cell_border(tbl.cell(r_, c_), "lnT", "A6A6A6", 6350)
for c_ in range(1, 6):
    cell_border(tbl.cell(rl, c_), "lnT", "1F3A5F", 9525)

# --- bridge por origem da receita e driver de custo
steps = [("R&D\n(intercompany)", d_rd), ("Terceiros", d_terc), ("Postergações\nde 2026", -post),
         ("Obrigações\nHSE", -hse), ("RH", -d_adm), ("Inflação e\noutros", -(infl + portaria + outros))]
assert abs(ebit[r9] + sum(v for _, v in steps) - ebit[mp]) < 0.5
waterfall(s1, 66, 700, 492, 150, ebit[r9], steps, "MP’27", "R9",
          highlight=("Postergações\nde 2026", "Obrigações\nHSE"), bracket=f"{p1(pct(ebit[mp], ebit[r9]))}%")
text(s1, 66, 528, 640, 10, f"Terceiros = eventos (externos e grupo), testes OEM e aluguéis de longo prazo. "
     f"Inflação e outros = reajuste de contratos ({fmt(infl)}), projeto da portaria ({fmt(portaria)}) e demais "
     f"({fmt(outros, sign=True)}).", size=7, italic=True, color=GREY)

panel_blocks(s1, 722, 70, 226, 450, "Mensagens-chave", [
    (f"EBIT {fmt(ebit[mp])}  |  {p1(pct(ebit[mp], ebit[r9]))}% vs R9",
     f"ROS de {p1(ros_mp, False)}% ({p1(ros_mp - ros_r9)} p.p.). A receita cresce mais que o custo fixo."),
    (f"Receita {fmt(net[mp])}  |  {p1(pct(net[mp], net[r9]))}% vs R9",
     f"R&D intercompany +{fmt(d_rd)} (contrato +{p1(pct(rd_mp_m, rd_r9_m), False)}%) e terceiros "
     f"+{fmt(d_terc)} com a retomada de eventos após um 2026 atípico."),
    (f"Custo fixo {fmt(-fix[mp])}  |  {p1(pct(fix[mp], fix[r9]))}% vs R9",
     f"{pontual / d_fix * 100:.0f}% do aumento é pontual ou obrigatório: postergações de 2026 ({fmt(post)}) e "
     f"licenças HSE ({fmt(hse)}). O restante cresce {p1(recorrente / -fix[r9] * 100, False)}%."),
    (f"Geração de caixa: EBITDA {fmt(ebitda)}",
     f"{p1(ebitda / net[mp] * 100, False)}% da receita; D&A estável em {fmt(-da[mp])}."),
    (f"Sem as postergações: EBIT {fmt(ebit_sem_post)}",
     f"ROS de {p1(ros_sem, False)}%. Gasto pontual que protegeu 2026."),
], size=9.5)
s1.notes_slide.notes_text_frame.text = (
    "Tabela: Base_Book_CP_.xlsx / Month_FY_LC_ (valores atuais).\n"
    "Bridge EBIT (R$ mil): " + " | ".join(f"{l.replace(chr(10), ' ')} {v:,.1f}" for l, v in
                                          [("R9", ebit[r9])] + steps + [("MP'27", ebit[mp])]) +
    f"\nTerceiros = Events sales + Events sales group + OEM + Long-term rents. Base Book (Tabela Variantes) "
    f"usa Eventos/OEM/R&D/Cost Fix.; aqui o custo fixo foi aberto pelos drivers da Base de Custo MP'27_v2 (Sheet1).")

layout = s1.slide_layout
TITLE = "Circuito Panamericano Economics MP’27"

# ================================================================== slide 2 – receita por origem
s2 = new_slide(prs, layout, TITLE)
gain_label(s2)
headline(s2, f"Receita líquida de BRL {p1(net[mp] / 1e3, False)} mi ({p1(pct(net[mp], net[r9]))}% vs R9)",
         " — R&D garante a base e terceiros retomam após 2026 atípico")

W, Y0, H = 196, 100, 190
ymax = max(rd_chart) * 1.18
small_multiple(s2, 36, Y0, W, H, "R&D Tests", rd_chart, "intercompany – contrato mensal", ymax)
small_multiple(s2, 36 + W + 14, Y0, W, H, "Terceiros", terc, "eventos + OEM + aluguéis", ymax * 0.5)

# total = R&D + terceiros (barras empilhadas, total no topo)
xt = 36 + 2 * (W + 14)
text(s2, xt, Y0, W - 66, 16, "Total", size=11, bold=True, color=NAVY)
text(s2, xt, Y0 + 15, W, 12, "receita líquida = R&D + terceiros", size=8, italic=True, color=GREY)
cd = CategoryChartData()
cd.categories = ["FY24", "FY25", "MP26", "R9", "MP’27"]
rd_r = [r0(v) for v in rd_chart]
tot_r = [r0(v) for v in net]
cd.add_series("R&D", rd_r)
cd.add_series("Terceiros", [r0(v) for v in terc])
cd.add_series("Total", [0] * 5)
gf = s2.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Pt(xt), Pt(Y0 + 28), Pt(W), Pt(H - 28), cd)
ch = gf.chart
style_chart(ch)
ch.has_title = False
ch.has_legend = False
plot = ch.plots[0]
plot.gap_width = 45
plot.overlap = 100
for k, (se, col) in enumerate(zip(plot.series, (NAVY, SKY, None))):
    if col is None:
        se.format.fill.background()
        for j, pt in enumerate(se.points):
            dl = pt.data_label
            dl.position = XL_LABEL_POSITION.INSIDE_BASE
            dl.text_frame.text = fmt(tot_r[j])
            dl.text_frame.paragraphs[0].runs[0].font.size = Pt(8)
            dl.text_frame.paragraphs[0].runs[0].font.bold = True
            dl.text_frame.paragraphs[0].runs[0].font.color.rgb = BLACK
        continue
    se.format.fill.solid()
    se.format.fill.fore_color.rgb = col
    dl = se.data_labels
    dl.show_value = True
    dl.number_format = '#,##0'
    dl.number_format_is_linked = False
    dl.position = XL_LABEL_POSITION.CENTER
    dl.font.size = Pt(7.5)
    dl.font.color.rgb = WHITE if col == NAVY else NAVY
ch.value_axis.minimum_scale = 0
ch.value_axis.maximum_scale = max(net) * 1.13
badge(s2, xt + W - 64, Y0, pct(net[mp], net[r9]))
# legenda manual
for k, (lab, col) in enumerate((("R&D", NAVY), ("Terceiros", SKY))):
    rect(s2, xt + 70 + k * 62, Y0 + H + 2, 8, 8, col)
    text(s2, xt + 81 + k * 62, Y0 + H, 50, 12, lab, size=8, color=DGREY)

# eventos mês a mês (bruto) – linhas
yb = Y0 + H + 22
text(s2, 36, yb, 420, 16, "Eventos externos – faturamento bruto mensal (BRL ’000)", size=11, bold=True, color=NAVY)
text(s2, 36, yb + 15, 420, 12, f"R9 {fmt(sum(ev_r9_m))}  →  MP’27 {fmt(sum(ev_mp_m))} no ano "
     f"({p1(pct(sum(ev_mp_m), sum(ev_r9_m)))}%)", size=8, italic=True, color=GREY)
cd = CategoryChartData()
cd.categories = MESES
cd.add_series("2026 (R9)", [r0(v) for v in ev_r9_m])
cd.add_series("MP’27", [r0(v) for v in ev_mp_m])
gf = s2.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Pt(30), Pt(yb + 26), Pt(432), Pt(520 - yb - 26), cd)
ch = gf.chart
style_chart(ch, size=7.5)
ch.has_title = False
ch.has_legend = True
ch.legend.position = XL_LEGEND_POSITION.TOP
ch.legend.include_in_layout = False
ch.legend.font.size = Pt(8)
ch.value_axis.minimum_scale = -90
ch.value_axis.maximum_scale = 560
ch.value_axis.crosses = XL_AXIS_CROSSES.MINIMUM
R9COL = RGBColor(0x8E, 0xA9, 0xCF)
for k, (se, col) in enumerate(zip(ch.plots[0].series, (R9COL, NAVY))):
    se.smooth = False
    se.format.line.color.rgb = col
    se.format.line.width = Pt(1.75 if k == 0 else 2.5)
    if k == 0:
        se.format.line.dash_style = MSO_LINE.DASH
    se.marker.style = 8
    se.marker.size = 5 if k == 0 else 6
    se.marker.format.fill.solid()
    se.marker.format.fill.fore_color.rgb = col if k else WHITE
    se.marker.format.line.color.rgb = col
    lab_col = NAVY if k else RGBColor(0x5B, 0x7D, 0xB1)
    for j in range(12):                        # rótulo da linha mais alta acima, da mais baixa abaixo
        mine, other = (ev_r9_m[j], ev_mp_m[j]) if k == 0 else (ev_mp_m[j], ev_r9_m[j])
        pl = se.points[j].data_label
        pl.has_text_frame = False
        pl.show_value = True
        pl.number_format = '#,##0'
        pl.number_format_is_linked = False
        pl.font.size = Pt(7.5)
        pl.font.bold = bool(k)
        pl.font.color.rgb = lab_col
        pl.position = XL_LABEL_POSITION.ABOVE if mine > other or (mine == other and k == 1) else XL_LABEL_POSITION.BELOW

# mix da receita MP'27 (rosca)
xm = 476
text(s2, xm, yb, 180, 16, "Mix da receita MP’27", size=11, bold=True, color=NAVY)
mix = [("R&D", rd[mp], NAVY), ("Eventos", ev_tot[mp], MIDBLUE), ("OEM + aluguéis", oem_tot[mp], SKY)]
cd = CategoryChartData()
cd.categories = [l for l, _, _ in mix]
cd.add_series("Mix", [v for _, v, _ in mix])
gf = s2.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Pt(xm - 6), Pt(yb + 22), Pt(104), Pt(104), cd)
ch = gf.chart
ch.has_legend = False
ch.has_title = False
plot = ch.plots[0]
for (_, _, col), pt in zip(mix, plot.series[0].points):
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = col
    pt.format.line.color.rgb = WHITE
    pt.format.line.width = Pt(1.5)
dn = plot._element
hs = dn.find(C + "holeSize")
if hs is None:
    hs = etree.SubElement(dn, C + "holeSize")
hs.set("val", "58")
fsa = dn.find(C + "firstSliceAng")
if fsa is not None:
    fsa.set("val", "0")
text(s2, xm - 6 + 27, yb + 22 + 40, 50, 24, [[(f"{fmt(net[mp])}", {"bold": True, "size": 9, "color": NAVY})],
                                             [("BRL ’000", {"size": 6.5, "color": GREY})]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
for k, (lab, v, col) in enumerate(mix):
    yy = yb + 34 + k * 28
    rect(s2, xm + 104, yy + 3, 8, 8, col)
    text(s2, xm + 116, yy, 64, 26, [[(f"{v / net[mp] * 100:.0f}%", {"bold": True, "size": 10, "color": NAVY}),
                                     (f"  {fmt(v)}", {"size": 7.5, "color": GREY})],
                                    [(lab, {"size": 7.5, "color": DGREY})]], spacing=0.95)
text(s2, xm, yb + 132, 180, 60, [[("Terceiros: ", {"bold": True, "size": 8, "color": NAVY}),
                                  (f"{(terc[mp]) / net[mp] * 100:.0f}% da receita ({fmt(terc[mp])}); "
                                   f"eventos com margem de contribuição de 100%.", {"size": 8, "color": DGREY})]],
     spacing=1.05)

panel_blocks(s2, 670, 100, 278, 420, "Destaques", [
    (f"R&D intercompany  |  {rd[mp] / net[mp] * 100:.0f}% da receita",
     f"Contrato mensal passa de {fmt(rd_r9_m)} para {fmt(rd_mp_m)} mil brutos "
     f"(+{p1(pct(rd_mp_m, rd_r9_m), False)}%): +{fmt(d_rd)} de receita líquida. Base estável do Circuito."),
    (f"Terceiros {fmt(terc[mp])}  |  {p1(pct(terc[mp], terc[r9]))}% vs R9",
     f"Eventos +{fmt(ev_tot[mp] - ev_tot[r9])} após um 2026 atípico (Copa do Mundo, entre outros fatores); "
     f"OEM e aluguéis estáveis ({fmt(oem_tot[mp])})."),
    ("Premissa conservadora",
     f"Terceiros ainda {p1(pct(terc[mp], terc[fy25]))}% vs FY25 e {p1(pct(terc[mp], terc[mp26]))}% vs MP26."),
    ("Agenda de eventos mais estável",
     f"MP’27 com {fmt(min(ev_mp_m))}–{fmt(max(ev_mp_m[:8]))} mil/mês até agosto e {fmt(ev_mp_m[-1])} mil/mês "
     f"de set a dez; em 2026, {p1(r9_2s, False)}% do ano ficou concentrado de set a dez."),
    ("Sensibilidade",
     f"Eventos não têm custo variável: cada 10% de eventos = ±{fmt(sens_ev)} de EBIT. "
     "OEM segue sob demanda das montadoras."),
], size=9.3)
s2.notes_slide.notes_text_frame.text = (
    "Fonte: Base_Book_CP_.xlsx (receita líquida) e Faturamento CP - MP'27.xlsx (bruto mensal de eventos).\n"
    f"Terceiros inclui Events sales group ({fmt(evg[mp])} no MP'27), que são eventos do grupo Pirelli.\n"
    f"R&D R9 mostrado como {fmt(rd_chart[r9])} (igual MP26); base book {rd[r9]:,.1f}.")

# ================================================================== slide 3 – custo fixo
s3 = new_slide(prs, layout, TITLE)
gain_label(s3)
headline(s3, f"Custo fixo de BRL {p1(-fix[mp] / 1e3, False)} mi ({p1(pct(fix[mp], fix[r9]))}% vs R9)",
         f" — {pontual / d_fix * 100:.0f}% do aumento vem de postergações de 2026 e obrigações HSE")
text(s3, 36, 100, 520, 16, "Custo fixo R9 → MP’27 (BRL ’000, custo = positivo)", size=11, bold=True, color=NAVY)
waterfall(s3, 36, 560, 440, 290, -fix[r9], DRIVERS, "MP’27", "R9",
          highlight=("Postergações\nde 2026", "Obrigações\nHSE"), good_up=False,
          bracket=f"{p1(pct(fix[mp], fix[r9]))}%")
text(s3, 36, 500, 530, 22, f"Inflação = reajuste de contratos da estrutura operacional ({fmt(infl_ops)}), "
     f"manutenção geral ({fmt(d_gm)}) e CSA/CSC ({fmt(d_csa)}). Outros = marketing ({fmt(d_mkt, sign=True)}), "
     f"D&A (estável) e demais itens da estrutura operacional ({fmt(ops_res, sign=True)}).",
     size=7, italic=True, color=GREY)

text(s3, 590, 100, 358, 16, "Composição do aumento (BRL ’000)", size=11, bold=True, color=NAVY)
rows = [["Despesas postergadas de 2026", fmt(post)],
        ["Renovação licenças ambientais + outorga (HSE)", fmt(HSE[0][1])],
        ["Coleta de lixo orgânico | biodigestor (HSE)", fmt(HSE[1][1])],
        ["Reajuste de contratos – estrutura operacional", fmt(infl_ops)],
        ["Reajuste – manutenção geral e CSA/CSC", fmt(d_gm + d_csa)],
        ["RH – reajuste salarial e provisões", fmt(d_adm)],
        ["Projeto da portaria", fmt(portaria)],
        ["Outros (marketing, D&A e demais)", fmt(outros)],
        [("Total", {"bold": True}), (fmt(d_fix), {"bold": True, "color": MIDBLUE})]]
table(s3, 590, 120, [298, 60], rows, row_h=16, header=["Driver", "BRL ’000"], size=8.5, total_last=True)

panel_blocks(s3, 590, 290, 358, 230, "O que está por trás do aumento", [
    (f"{pontual / d_fix * 100:.0f}% pontual ou obrigatório  |  {fmt(pontual)}",
     f"Postergações de 2026 ({fmt(post)}) e obrigações de HSE ({fmt(hse)}): renovação das licenças "
     "ambientais + outorga e biodigestor."),
    (f"Custo recorrente: {p1(recorrente / -fix[r9] * 100)}%  |  {fmt(recorrente)}",
     f"Inflação de contratos ({fmt(infl)}), RH ({fmt(d_adm)}: dissídio e provisões) e projeto da portaria "
     f"({fmt(portaria)})."),
    (f"D&A {fmt(-da[mp])}  |  {-da[mp] / -fix[mp] * 100:.0f}% do custo fixo",
     "Estável e sem efeito caixa."),
], size=9)
s3.notes_slide.notes_text_frame.text = (
    "Drivers: Base de Custo MP'27_v2.xlsx / Sheet1 (justificativas) + Base_Book_CP_.xlsx.\n"
    + " | ".join(f"{l.replace(chr(10), ' ')} {v:,.1f}" for l, v in DRIVERS) + f" | Total {d_fix:,.1f}\n"
    f"Estrutura operacional: delta {d_ops:,.1f}; explicado na Sheet1 {ops_expl:,.1f}; resíduo {ops_res:,.1f}.")

# ================================================================== slide 4 – postergações
s4 = new_slide(prs, layout, TITLE)
gain_label(s4)
headline(s4, f"BRL {fmt(post)} mil de despesas postergadas de 2026 para 2027",
         " — gasto pontual que protegeu 2026 e volta no custo do MP’27")

text(s4, 36, 104, 440, 16, "Despesas postergadas (2026 → MP’27)", size=11, bold=True, color=NAVY)
rows = [[(g, {"color": GREY, "size": 8.5}), it, fmt(v), f"{v / post * 100:.0f}%"] for g, it, v in POST]
rows.append(["", "TOTAL", (fmt(post), {"bold": True, "color": MIDBLUE}), "100%"])
table(s4, 36, 124, [86, 262, 62, 56], rows, row_h=28, header=["Área", "Item", "BRL ’000", "% total"], size=9.5,
      total_last=True, col_align=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT])

grp = {}
for g, _, v in POST:
    grp[g] = grp.get(g, 0) + v
order = sorted(grp, key=lambda g: -grp[g])
yg = 124 + 28 * 7 + 26
text(s4, 36, yg, 440, 16, "Por área", size=11, bold=True, color=NAVY)
rows = [[g, fmt(grp[g]), f"{grp[g] / post * 100:.0f}%"] for g in order]
rows.append(["TOTAL", (fmt(post), {"bold": True, "color": MIDBLUE}), "100%"])
table(s4, 36, yg + 20, [348, 62, 56], rows, row_h=22, header=["Área", "BRL ’000", "% total"], size=9.5,
      total_last=True, col_align=[PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT])

top2 = POST[0][2] + POST[1][2]
panel_blocks(s4, 530, 104, 418, 416, "Mensagens-chave", [
    ("2026: alavanca para proteger o resultado",
     "Com a queda de eventos (ano de Copa do Mundo, entre outros fatores), estes itens foram adiados sem "
     "comprometer a operação e a segurança do Circuito."),
    (f"{top2 / post * 100:.0f}% em dois itens",
     f"Cercamento e manutenções do Sansão ({fmt(POST[0][2])}) e restauração da APP ({fmt(POST[1][2])}): "
     "manutenção da área e compromisso ambiental."),
    (f"{post / d_fix * 100:.0f}% do aumento do custo fixo",
     f"{fmt(post)} dos {fmt(d_fix)} de aumento vs R9; {p1(post / -fix[mp] * 100, False)}% do custo fixo do MP’27."),
    (f"Sem as postergações: EBIT {fmt(ebit_sem_post)}",
     f"ROS de {p1(ros_sem, False)}% vs {p1(ros_mp, False)}% no MP’27. Gasto pontual, não recorrente."),
], size=10)
s4.notes_slide.notes_text_frame.text = (
    "Itens e valores: Base de Custo MP'27_v2.xlsx / Sheet1 (justificativa 'postergação de 2026').\n"
    "Lista original de 2026 (611) – itens que NÃO entram no MP'27: reajustes de contratos 74, pintura box/"
    "sinalizações 70, plantio bambuzal 50, retirada de bambu 45, viagens/reclassificações 35, sanitários/"
    "vestiários 21, custos I.T. 20, pluvial 2 (= 317). Valores atualizados: Sansão 170→180, restauração 72→167 "
    "(Restauração APP), elétrica 23→35.")

# ================================================================== slide 5 – perguntas-chave
s5 = new_slide(prs, layout, TITLE)
gain_label(s5)
headline(s5, "Perguntas-chave", " — o que sustenta o MP’27 do Circuito Panamericano")
QA = [
    ("O que explica a melhora do EBIT?",
     f"Receita +{fmt(d_net)} (R&D +{fmt(d_rd)}; terceiros +{fmt(d_terc)}) mais que compensa o custo fixo "
     f"+{fmt(d_fix)}. EBIT {fmt(ebit[mp])} ({p1(pct(ebit[mp], ebit[r9]))}% vs R9)."),
    ("Qual a dependência do contrato de R&D?",
     f"{rd[mp] / net[mp] * 100:.0f}% da receita é intercompany. Sem o reajuste do contrato "
     f"({p1(pct(rd_mp_m, rd_r9_m), False)}%), o EBIT seria {fmt(ebit_sem_reaj)}, abaixo do R9 "
     f"({fmt(ebit[r9])})."),
    ("A retomada de eventos é realista?",
     f"Bruto de {fmt(sum(ev_mp_m))} vs {fmt(sum(ev_r9_m))} no R9, ainda abaixo de FY25 e do MP26. "
     f"Cada 10% de eventos = ±{fmt(sens_ev)} de EBIT."),
    ("O aumento do custo fixo é recorrente?",
     f"Não na maior parte: {fmt(post)} são postergações pontuais e {fmt(hse)} obrigações HSE. "
     f"O custo recorrente cresce {p1(recorrente / -fix[r9] * 100, False)}%."),
    ("E a geração de caixa?",
     f"EBITDA de {fmt(ebitda)} ({p1(ebitda / net[mp] * 100, False)}% da receita). "
     f"D&A estável em {fmt(-da[mp])}."),
    ("Por que abaixo do MP26?",
     f"EBIT {fmt(ebit[mp] - ebit[mp26], sign=True)} vs MP26: eventos {fmt(ev_tot[mp] - ev_tot[mp26], sign=True)} "
     f"e custo fixo {fmt(fix[mp] - fix[mp26], sign=True)}, compensados em parte pelo R&D "
     f"{fmt(rd[mp] - rd[mp26], sign=True)}."),
]
CWID, CH, GX, GY = 450, 104, 12, 10
for i, (q, a) in enumerate(QA):
    x = 36 + (i % 2) * (CWID + GX)
    y = 104 + (i // 2) * (CH + GY)
    rect(s5, x, y, CWID, CH, PANEL, LIGHT)
    rect(s5, x, y, 4, CH, MIDBLUE)
    text(s5, x + 16, y + 10, CWID - 30, 20, q, size=12, bold=True, color=NAVY)
    text(s5, x + 16, y + 34, CWID - 30, CH - 40, a, size=10.5, color=DGREY, spacing=1.1)
yr = 104 + 3 * (CH + GY) + 4
rect(s5, 36, yr, 2 * CWID + GX, 520 - yr, NAVY)
for k, (head, body) in enumerate([
        ("Riscos", f"Agenda de eventos no 2º semestre ({fmt(ev_mp_m[-1])} mil/mês brutos) e demanda de OEM, "
                   "que segue sob demanda das montadoras."),
        ("Alavancas", "Mesmo racional de 2026: escalonar itens não obrigatórios ao longo do ano se eventos não "
                      f"se confirmarem; eventos acima da premissa somam {fmt(sens_ev)} de EBIT a cada 10%.")]):
    x = 36 + k * (CWID + GX)
    text(s5, x + 16, yr + 10, CWID - 30, 76, [[(head, {"bold": True, "size": 12, "color": WHITE})],
                                              [(body, {"size": 10.5, "color": WHITE})]], spacing=1.1, after=4)
s5.notes_slide.notes_text_frame.text = (
    "Slide de apoio para perguntas do presidente. Números do Base Book, Faturamento e Base de Custo MP'27_v2.")

prs.save(OUT)
print("ok", OUT)
print("EBIT R9/MP27", round(ebit[r9], 1), round(ebit[mp], 1), "net", round(net[mp], 1), "fix", round(fix[mp], 1))
print("drivers", [(l.replace(chr(10), ' '), round(v, 1)) for l, v in DRIVERS], "total", round(d_fix, 1))
print("bridge", [(l.replace(chr(10), ' '), round(v, 1)) for l, v in steps])
print(f"post {post:.1f} hse {hse:.1f} pontual {pontual / d_fix * 100:.1f}% recorrente {recorrente / -fix[r9] * 100:.2f}%"
      f" ebitda {ebitda:.1f} sem_post {ebit_sem_post:.1f} ros_sem {ros_sem:.2f} sem_reaj {ebit_sem_reaj:.1f}")
