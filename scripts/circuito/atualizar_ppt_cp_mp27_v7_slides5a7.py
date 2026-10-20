"""Monta os slides 5 a 7 do deck MP'27 do Circuito Panamericano (Bridge CP_v7) sobre a versão salva pela Juliana.

Uso:
    python atualizar_ppt_cp_mp27_v7_slides5a7.py <Bridge CP_v7.pptx> <Base de Custo MP'27_v2.xlsx>
                                                 <MENS CP 2027 - MP'27.xlsx> <saida.pptx>

Pedido de 20/10: slides 1 a 4 NÃO podem mudar (aqui só se mexe em slides 5, 6 e 7; os demais não são tocados).
- slide 5: principais linhas de custo dentro de cada voz do custo fixo (aba Base Reajustes da Base de Custo);
- slide 6: riscos e oportunidades do CP (mesma estrutura do slide de Fitted que ela deixou como modelo);
- slide 7: estrutura vazia de investimentos 2027 (o time preenche os itens).
"""
import re
import sys
import copy
from lxml import etree
import openpyxl
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC, BC, MENS, OUT = sys.argv[1:5]

NAVY = RGBColor(0x1F, 0x3A, 0x5F)
MIDBLUE = RGBColor(0x2E, 0x5A, 0xA8)
LIGHT = RGBColor(0xDC, 0xE3, 0xF0)
PANEL = RGBColor(0xF3, 0xF5, 0xFA)
SKY = RGBColor(0x9D, 0xC3, 0xE6)
GREY = RGBColor(0x80, 0x80, 0x80)
DGREY = RGBColor(0x40, 0x40, 0x40)
BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def r0(v):
    v = round(v, 2)
    return int(v + 0.5) if v >= 0 else -int(-v + 0.5)


def fmt(v, sign=False):
    s = f"{abs(r0(v)):,}".replace(",", ".")
    if r0(v) < 0:
        return "−" + s
    return ("+" if sign and r0(v) > 0 else "") + s


def num(x):
    return x if isinstance(x, (int, float)) else 0


# ------------------------------------------------------------------ dados: MENS (totais do MP'27, R9) e Base de Custo
loc = openpyxl.load_workbook(MENS, data_only=True)["LOCAL"]


def mens(row, col="Q"):
    return (loc[f"{col}{row}"].value or 0) / 1e3


net_mp, fix_mp, ebit_mp, ebit_r9 = mens(21), -mens(30), mens(40), mens(40, "S")
ev_mp = mens(22) + mens(26)                       # eventos líquidos (externos + grupo)
oem_mp = mens(23) + mens(24)                      # OEM + aluguéis
rd_mp, rd_r9 = mens(25), mens(25, "S")
ev_gross = mens(8)
adm_mp = -mens(32)
TOT_MENS = {"General maintenance": -mens(31), "Administrative costs": -mens(32),
            "Operational structure Circuito Panamericano": -mens(33), "Marketing and promotions": -mens(34),
            "Costs CSA/CSC": -mens(35), "D&A": -mens(36)}

wb = openpyxl.load_workbook(BC, data_only=True)
ws = wb["Base Reajustes"]
assert ws.cell(4, 43).value == "TOTAL MP27"
linhas = []                                       # (voz, descrição, MP'27 em BRL mil)
for r in range(5, 84):
    voz = ws.cell(r, 1).value
    if voz:
        linhas.append((voz, (ws.cell(r, 2).value or "").strip(), num(ws.cell(r, 43).value) / 1e3))
tot_voz = {}
for voz, _, v in linhas:
    tot_voz[voz] = tot_voz.get(voz, 0) + v
for voz, t in tot_voz.items():                    # a Base de Custo tem que fechar com a MENS
    assert abs(t - TOT_MENS[voz]) < 1.0, (voz, t, TOT_MENS[voz])
just = {}
for r in range(13, 26):
    k, v, t = wb["Sheet1"].cell(r, 3).value, wb["Sheet1"].cell(r, 4).value, wb["Sheet1"].cell(r, 5).value
    if k and isinstance(v, (int, float)):
        just[str(k).strip()] = (v / 1e3, (t or "").strip().lower())
sansao_post = just["Cercamento | Manutenções - Sansão"][0]


def agrupar(voz, regras, mostrar):
    """regras = [(rótulo, regex, tag)]; devolve as `mostrar` maiores linhas + 'Demais linhas'."""
    ac, tags = {}, {}
    for v_, desc, val in linhas:
        if v_ != voz:
            continue
        for rot, rx, tag in regras:
            if re.search(rx, desc, re.I):
                ac[rot] = ac.get(rot, 0) + val
                tags[rot] = tag
                break
        else:
            ac["Demais linhas"] = ac.get("Demais linhas", 0) + val
    top = sorted((k for k in ac if k != "Demais linhas"), key=lambda k: -ac[k])[:mostrar]
    out = [(k, ac[k], tags[k]) for k in top]
    resto = sum(v for k, v in ac.items() if k not in top)
    out.append(("Demais linhas", resto, ""))
    assert abs(sum(v for _, v, _ in out) - tot_voz[voz]) < 0.5
    return out


GM = agrupar("General maintenance", [
    ("Vigilância e segurança", r"^Security", ""),
    ("Jardinagem", r"^Jardinagem", ""),
    ("Plantio, restauração de APP e ambiental",
     r"^Manutenção [12] fase|^Plantio|^Retirada de Bambus|^Delta SP|^Laudo Fauna|^Relatorios APP|^Acompanhamento Ambiental|^Irrigação$", ""),
    ("Limpeza", r"^Limpeza", ""),
    ("Off Road comercial (manutenção mensal)", r"^Off Road", ""),
    ("Locação de veículos", r"Locação Carro", ""),
], 6)
OPS = agrupar("Operational structure Circuito Panamericano", [
    ("Energia elétrica", r"^Energia", ""),
    ("Licenças ambientais, outorga e biodigestor", r"^RENOVAÇÃO DAS LICENÇAS|Biodigestor", "novo · HSE"),
    ("Manutenção predial, pintura e tendas", r"^Predial|^Pintura prédio|^Tendas", ""),
    ("Cercamento e manutenções (Sansão)", r"^Cercamento", f"{fmt(sansao_post)} de postergação"),
    ("Irrigação (peças, bombeamento, lagos)", r"^Irriga", ""),
    ("IPTU", r"^IPTU", ""),
    ("Coleta de lixo e resíduos", r"^Coleta de Lixo - Res", ""),
    ("Contratos e locações de apoio",
     r"^Gerador|^Climatização|^Poços|^Elevador|^Container|^Rádios|^Purificadores|^LOCAÇÃO MÁQUINA|^Compressores|^Combate", ""),
], 8)
mkt_lin = [(d, v) for v_, d, v in linhas if v_ == "Marketing and promotions" and abs(v) > 0.05]
mkt_pub = sum(v for d, v in mkt_lin if d.startswith("Mkt"))
mkt_via = sum(v for d, v in mkt_lin if d.startswith("Viagens"))
assert abs(mkt_pub + mkt_via - tot_voz["Marketing and promotions"]) < 0.5

# ------------------------------------------------------------------ riscos e oportunidades (sensibilidades do plano)
d_rd = rd_mp - rd_r9
sens_ev = ev_mp * 0.10                            # eventos: margem de contribuição de 100%
sens_oem = oem_mp * 0.10
infl_plano, infl_pct = 213.0, 4.4                 # abertura do custo fixo informada por ela (20/10)
sens_infl = infl_plano / infl_pct                 # 1 p.p. de inflação, proporcional ao plano
sens_labour = adm_mp * 0.01                       # 1 p.p. sobre a folha (RH)
post_total = just["Cercamento | Manutenções - Sansão"][0] + just["Restauração APP"][0] + just["Elétrica"][0] \
    + just["Ações de marketing"][0] + just["Análise qualidade do ar (ar-condicionado)"][0]
RISCOS = [("Eventos abaixo da premissa", f"−10% sobre {fmt(ev_gross)} brutos; margem de contribuição de 100%", -sens_ev),
          ("Reajuste do contrato de R&D não confirmado", f"Contrato mensal fica no valor do R9 (sem o +8,5%)", -d_rd),
          ("Inflação acima de 4,4%", "Cada +1 p.p. sobre contratos e reajustes, proporcional ao plano", -sens_infl),
          ("OEM abaixo da premissa", f"−10% sobre {fmt(oem_mp)} líquidos; demanda sob encomenda das montadoras", -sens_oem),
          ("Labour acima do plano", f"Cada +1 p.p. sobre a folha de {fmt(adm_mp)}", -sens_labour)]
OPORT = [("Eventos acima da premissa", f"+10% sobre {fmt(ev_gross)} brutos; sem custo variável", sens_ev),
         ("Escalonar itens não obrigatórios", f"Mesmo racional de 2026: postergações de {fmt(post_total)} se eventos não vierem", post_total),
         ("OEM acima da premissa", f"+10% sobre {fmt(oem_mp)} líquidos", sens_oem),
         ("Inflação abaixo de 4,4%", "Cada −1 p.p. sobre contratos e reajustes", sens_infl)]
tot_r, tot_o = sum(v for *_, v in RISCOS), sum(v for *_, v in OPORT)


# ------------------------------------------------------------------ helpers de desenho
def line(sl, x1, y1, x2, y2, color=GREY, w=0.75):
    c = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Pt(x1), Pt(y1), Pt(x2), Pt(y2))
    c.line.color.rgb = color
    c.line.width = Pt(w)
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


def rect(sl, x, y, w, h, fill, edge=None, ew=0.75):
    r = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))
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
            etree.SubElement(p._p, A + "endParaRPr").set("sz", str(int(f.get("size", size) * 100)))
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
                ln = etree.Element(A + "lnB", w="6350")
                tcPr.insert(0, ln)                       # lnB tem que vir antes do preenchimento da célula
                sf = etree.SubElement(ln, A + "solidFill")
                etree.SubElement(sf, A + "srgbClr", val="BFBFBF")
    return gt


def panel_blocks(sl, x, y, w, h, title, items, size=10):
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
              [(body, {"color": DGREY, "size": size, "italic": body.startswith("[")})]],
             anchor=MSO_ANCHOR.MIDDLE, spacing=1.08, after=3)


def limpar(sl, manter):
    for shp in list(sl.shapes):
        if not manter(shp):
            shp._element.getparent().remove(shp._element)


def set_paras(shape, textos):
    """troca o texto mantendo a formatação do 1º trecho de cada parágrafo (os demais trechos saem)."""
    for para, t in zip(shape.text_frame.paragraphs, textos):
        runs = para.runs
        runs[0].text = t
        for r in runs[1:]:
            r._r.getparent().remove(r._r)


def card(sl, x, y, w, h, nome, total, share, rows, bar_w=105, row_h=21):
    """cartão de voz do custo fixo: total, % do custo fixo e principais linhas com barra proporcional."""
    rect(sl, x, y, w, h, PANEL, LIGHT)
    rect(sl, x, y, 4, h, MIDBLUE)
    text(sl, x + 16, y + 9, w - 130, 16, nome, size=11.5, bold=True, color=NAVY)
    text(sl, x + 16, y + 25, w - 130, 12, f"{share:.0f}% do custo fixo", size=8, italic=True, color=GREY)
    text(sl, x + w - 112, y + 8, 98, 22, fmt(total), size=15, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
    line(sl, x + 16, y + 42, x + w - 14, y + 42, color=RGBColor(0xC9, 0xD3, 0xE6), w=0.5)
    mx = max(abs(v) for _, v, _ in rows) or 1
    for i, (lab, v, tag) in enumerate(rows):
        yy = y + 46 + i * row_h
        resto = lab == "Demais linhas"
        runs = [(lab, {"size": 9, "color": GREY if resto else BLACK, "italic": resto})]
        if tag:
            runs.append(("  " + tag, {"size": 7.5, "color": MIDBLUE, "italic": True}))
        text(sl, x + 16, yy, 250, row_h, [runs], anchor=MSO_ANCHOR.MIDDLE)
        bx = x + 16 + 256
        if v > 0:
            rect(sl, bx, yy + 6, max(bar_w * v / mx, 1.5), row_h - 12, SKY if resto else NAVY)
        text(sl, x + w - 62, yy, 48, row_h, fmt(v), size=9.5, bold=not resto, color=GREY if resto else BLACK,
             align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def small_card(sl, x, y, w, h, nome, total, share, rows, nota):
    rect(sl, x, y, w, h, PANEL, LIGHT)
    rect(sl, x, y, 4, h, MIDBLUE)
    text(sl, x + 14, y + 9, w - 90, 16, nome, size=11, bold=True, color=NAVY)
    text(sl, x + 14, y + 25, w - 90, 12, f"{share:.0f}% do custo fixo", size=8, italic=True, color=GREY)
    text(sl, x + w - 80, y + 8, 68, 22, fmt(total), size=15, bold=True, color=NAVY, align=PP_ALIGN.RIGHT)
    line(sl, x + 14, y + 42, x + w - 12, y + 42, color=RGBColor(0xC9, 0xD3, 0xE6), w=0.5)
    yy = y + 48
    for lab, v in rows:
        text(sl, x + 14, yy, w - 76, 28, lab, size=9, anchor=MSO_ANCHOR.TOP, spacing=1.05)
        text(sl, x + w - 62, yy, 50, 14, fmt(v), size=9.5, bold=True, align=PP_ALIGN.RIGHT)
        yy += 30
    text(sl, x + 14, y + h - 40, w - 26, 34, nota, size=8, italic=True, color=GREY, anchor=MSO_ANCHOR.BOTTOM,
         spacing=1.05)


prs = Presentation(SRC)
s5, s6, s7 = prs.slides[4], prs.slides[5], prs.slides[6]

# ================================================================== slide 5 – principais linhas de custo por voz
limpar(s5, lambda shp: shp.is_placeholder or (shp.has_text_frame and shp.text_frame.text.startswith("(+) gain")))
cinco = (tot_voz["D&A"] + sum(v for k, v, _ in GM[:1]) + tot_voz["Administrative costs"]
         + sum(v for k, v, _ in GM if k == "Jardinagem") + sum(v for k, v, _ in OPS if k == "Energia elétrica"))
headline(s5, f"Custo fixo de BRL {fix_mp / 1e3:.1f} mi: ".replace(".", ","),
         f"D&A, vigilância, RH, jardinagem e energia somam {cinco / fix_mp * 100:.0f}%")
share = {v: t / fix_mp * 100 for v, t in tot_voz.items()}
CW1, GAP = 450, 12
Y1, H1 = 100, 240
card(s5, 36, Y1, CW1, H1, "Manutenção geral", tot_voz["General maintenance"], share["General maintenance"], GM)
card(s5, 36 + CW1 + GAP, Y1, CW1, H1, "Estrutura operacional",
     tot_voz["Operational structure Circuito Panamericano"], share["Operational structure Circuito Panamericano"], OPS)
Y2, H2 = Y1 + H1 + 10, 170
SW = (912 - 3 * GAP) / 4
small_card(s5, 36, Y2, SW, H2, "Administrativo", tot_voz["Administrative costs"], share["Administrative costs"],
           [("RH – ordenados, dissídio e provisões", tot_voz["Administrative costs"])],
           "Única linha da voz: folha da estrutura do Circuito.")
small_card(s5, 36 + (SW + GAP), Y2, SW, H2, "Marketing e promoções", tot_voz["Marketing and promotions"],
           share["Marketing and promotions"], [("Publicidade (contrato mensal)", mkt_pub), ("Viagens", mkt_via)],
           "Sem reserva de patrocínio no MP’27.")
small_card(s5, 36 + 2 * (SW + GAP), Y2, SW, H2, "CSA / CSC", tot_voz["Costs CSA/CSC"], share["Costs CSA/CSC"],
           [("Repasse CSC (centro de serviços do grupo)", tot_voz["Costs CSA/CSC"])],
           "Rateio do grupo, sem gestão local.")
small_card(s5, 36 + 3 * (SW + GAP), Y2, SW, H2, "D&A", tot_voz["D&A"], share["D&A"],
           [("Depreciação – edifícios e construções", tot_voz["D&A"])],
           "Sem efeito caixa; estável vs R9.")
text(s5, 36, 524, 800, 10, "Fonte: Base de Custo MP’27 (aba Base Reajustes), valores do MP’27 em BRL ’000; linhas agrupadas por natureza. "
     "O total de cada voz fecha com a MENS.", size=7, italic=True, color=GREY)
s5.notes_slide.notes_text_frame.text = (
    "Fonte: Base de Custo MP'27_v2.xlsx, aba Base Reajustes (coluna TOTAL MP27), salva em 20/10 11:25. Cada voz fecha com a "
    "MENS CP 2027 (aba LOCAL).\n"
    "Sem comparação com o R9 por linha: o R9 da Base de Custo (9.791) difere do R9 da MENS (9.740), então os deltas "
    "por linha não fecham com o bridge do slide 3.\n"
    + "\n".join(f"{n}: " + " | ".join(f"{l} {v:,.1f}" for l, v, _ in rows)
                for n, rows in (("Manutenção geral", GM), ("Estrutura operacional", OPS))))

# ================================================================== slide 6 – riscos e oportunidades
tit = [p for p in s6.shapes if p.name == "Title 1"][0]
set_paras(tit, ["Circuito Panamericano Riscos e Oportunidades MP’27"])
shp = {p.name: p for p in s6.shapes}
set_paras(shp["Mensagem principal"], [f"Riscos quantificados de {fmt(tot_r)}, concentrados no contrato de R&D; "
                                      f"oportunidades de {fmt(tot_o, sign=True)}"])
for k in ("Cabecalho Riscos unidade", "Cabecalho Oportunidades unidade"):
    set_paras(shp[k], ["EBIT (BRL ’000)"])
for i, (tit_, sub, v) in enumerate(RISCOS, start=1):
    set_paras(shp[f"Riscos item {i} texto"], [tit_, sub])
    set_paras(shp[f"Riscos item {i} valor"], [fmt(v)])
set_paras(shp["Total Riscos valor"], [fmt(tot_r)])
set_paras(shp["Total Riscos rotulo"], ["Total quantificado"])
# a lista de oportunidades tem 3 linhas no modelo: acrescenta a 4ª copiando a 3ª, uma linha abaixo
dy = shp["Oportunidades item 3"].top - shp["Oportunidades item 2"].top
novo_id = max(p.shape_id for p in s6.shapes) + 1
for nome in ("Oportunidades item 3", "Oportunidades item 3 texto", "Oportunidades item 3 valor"):
    el = copy.deepcopy(shp[nome]._element)
    el.xpath(".//p:cNvPr")[0].set("id", str(novo_id))
    el.xpath(".//p:cNvPr")[0].set("name", nome.replace("item 3", "item 4"))
    novo_id += 1
    shp[nome]._element.addnext(el)
    for p in s6.shapes:
        if p._element is el:
            p.top = p.top + dy
shp = {p.name: p for p in s6.shapes}
for i, (tit_, sub, v) in enumerate(OPORT, start=1):
    set_paras(shp[f"Oportunidades item {i} texto"], [tit_, sub])
    set_paras(shp[f"Oportunidades item {i} valor"], [fmt(v, sign=True)])
set_paras(shp["Total Oportunidades valor"], [fmt(tot_o, sign=True)])
set_paras(shp["Total Oportunidades rotulo"], ["Total quantificado"])
set_paras(shp["Nota de rodape"], [
    "Fonte: MENS CP 2027 – MP’27 e Base de Custo MP’27. Sensibilidades de ±10% sobre eventos e OEM (margem de contribuição de 100%), "
    f"±1 p.p. de inflação (plano: {fmt(infl_plano)} para {infl_pct:.1f}%, proporcional) e ±1 p.p. na folha; "
    "itens simétricos, não somam um cenário único. Premissas a validar."])
s6.notes_slide.notes_text_frame.text = (
    "Riscos e oportunidades do CP calculados como sensibilidades do plano (proposta a validar com a Juliana).\n"
    + " | ".join(f"{a}: {v:,.1f}" for a, _, v in RISCOS + OPORT))

# ================================================================== slide 7 – investimentos (estrutura)
tit = [p for p in s7.shapes if p.name == "Title 1"][0]
set_paras(tit, ["Circuito Panamericano Investimentos MP’27"])
limpar(s7, lambda shp_: shp_.is_placeholder or (shp_.has_text_frame and shp_.text_frame.text.startswith("(+) gain")))
headline(s7, "Investimentos de 2027", " — [mensagem principal a definir pelo time]")
KPI = [("Total de investimentos MP’27", "[ ]", "BRL ’000"), ("vs 2026", "[ ]", "BRL ’000 e %"),
       ("Obrigatórios (HSE e segurança)", "[ ]", "% do total")]
kw = (912 - 2 * 12) / 3
for i, (nome, val, sub) in enumerate(KPI):
    x = 36 + i * (kw + 12)
    rect(s7, x, 100, kw, 58, PANEL, LIGHT)
    rect(s7, x, 100, 4, 58, MIDBLUE)
    text(s7, x + 16, 108, kw - 30, 14, nome, size=10, bold=True, color=NAVY)
    text(s7, x + 16, 124, kw - 30, 26, [[(val, {"size": 18, "bold": True, "color": GREY}),
                                         ("   " + sub, {"size": 8.5, "italic": True, "color": GREY})]])
G = {"color": GREY, "italic": True}
rows = [[("[item de investimento]", G), ("[categoria]", G), ("[justificativa]", G), ("[ ]", G)]]
rows += [["", "", "", ""] for _ in range(8)]
rows.append(["Total", "", "", "[ ]"])
table(s7, 36, 174, [210, 110, 190, 70], rows, row_h=29.5,
      header=["Item", "Categoria", "Justificativa", "BRL ’000"], size=9, total_last=True,
      col_align=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT])
panel_blocks(s7, 640, 174, 308, 346, "Mensagens-chave", [
    ("[Mensagem 1]", "[o que o time quer destacar: total e prioridades]"),
    ("[Mensagem 2]", "[itens obrigatórios e prazo]"),
    ("[Mensagem 3]", "[retorno esperado ou risco de não fazer]"),
], size=9.5)
s7.notes_slide.notes_text_frame.text = "Slide só com a estrutura (pedido da Juliana, 20/10): o time preenche os itens de investimento de 2027."

prs.save(OUT)
print("ok", OUT)
print("fix", round(fix_mp, 1), "5 linhas", round(cinco, 1), f"{cinco / fix_mp * 100:.1f}%")
print("GM", [(a, round(b, 1)) for a, b, _ in GM])
print("OPS", [(a, round(b, 1)) for a, b, _ in OPS])
print("riscos", [round(v, 1) for *_, v in RISCOS], round(tot_r, 1), "oport", [round(v, 1) for *_, v in OPORT], round(tot_o, 1))
