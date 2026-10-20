"""PPT MP27 Fitted Units - ajustes na v3 (pedido de 19/10 noite), sobre a v3 salva pela Juliana.

Uso: python gerar_ppt_mp27_v3_ajustes.py <deck_v3.pptx> <basebook.xlsx> <mens.xlsx> <rentab_cliente_ibi.xlsx>
                                         <variacao_mo.xlsx> <detalhe_despesas.xlsx> <volumes_mp27.xlsx> <pasta_logos> <saida.pptx>

- Resumo: bloco de decisoes renomeado, com 3 itens;
- IBI: 3o turno Fiat em valor (MO 2,5 + demais 2,0, premissa dela) e ANTT como custo extra;
- Rentabilidade por cliente: nova versao (rateio por volume), colunas por cliente com logo;
- Mao de obra: bridge por driver (estrutura 2026 em ano cheio x reajuste 2027);
- novos slides New Business e Risks & Opportunities (antes das Perguntas-chave).
Pasta de logos: fiat.png, iveco.png, cnh.png, nissan.png, vw.png.
"""
import copy
import os
import sys

import openpyxl
from PIL import Image

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
VOL, LOGOS = sys.argv[7], sys.argv[8]
sys.argv = sys.argv[:7] + sys.argv[9:]          # rev2 espera 7 argumentos
import gerar_ppt_mp27_rev2 as g  # noqa: E402
from gerar_ppt_mp27_rev2 import (BLACK, DARK, DGREY, GREEN, GREY, HEAD, LIGHT, PANEL, RED, WHITE, Emu,  # noqa: E402
                                 MSO_ANCHOR, MSO_SHAPE, PP_ALIGN, Pt, R, RGBColor, dd, f1, fp, pc, sg)

prs = g.prs
mo = g.mo
UNITS = g.UNITS

# premissas informadas pela Juliana (19/10): custo anual do 3o turno Fiat em IBI
T3_MO, T3_DESP = 2.5, 2.0


def by_title(text):
    for s in prs.slides:
        t = s.shapes.title
        if t is not None and t.text_frame.text.strip() == text:
            return s
    raise KeyError(text)


def shape(slide, name):
    return [x for x in slide.shapes if x.name == name][0]


def drop(slide, pred):
    for sh in [x for x in slide.shapes if pred(x)]:
        sh._element.getparent().remove(sh._element)


def set_text(sh, text):
    p = sh.text_frame.paragraphs[0]
    p.runs[0].text = text
    for r in p.runs[1:]:
        r._r.getparent().remove(r._r)


def comments(slide, title, blocks, size=10.5):
    drop(slide, lambda x: x.name in ('Painel comentarios', 'Painel titulo', 'Comentarios'))
    g.draw_comments(slide, title, blocks, size)


def logo(slide, name, cx, y, h, max_w):
    path = os.path.join(LOGOS, f'{name}.png')
    w0, h0 = Image.open(path).size
    w = int(h * w0 / h0)
    if w > max_w:
        w, h = max_w, int(max_w * h0 / w0)
    pic = slide.shapes.add_picture(path, Emu(cx - w // 2), Emu(y), Emu(w), Emu(h))
    pic.name = f'Logo {name}'
    return pic


def rich(slide, x, y, w, h, parts, name, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    """parts: lista de paragrafos; cada paragrafo = lista de (texto, size, bold, rgb)."""
    tb = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for i, par in enumerate(parts):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if i:
            p.space_before = Pt(3)
        for t, sz, b, c in par:
            r = p.add_run()
            r.text = t
            r.font.size = Pt(sz)
            r.font.bold = b
            r.font.color.rgb = c
    return tb


# ================================================================ numeros
BR = g.PL['BR']
e27 = BR['ebit'][4]
fiat_pct = g.FIAT_PR27 / g.FIAT_PR9 - 1
antt = -dd('IBI', 'ANTT')
cobertura = g.FIAT_PRECO / (T3_MO + T3_DESP + antt)
mo9, mo27 = mo('ALL', i=0), mo('ALL')
GR_SAL = ['Salarios e PPR', 'Encargos', 'Outros beneficios']
sal = lambda u, i: sum(mo(u, x, i) for x in GR_SAL)
med = lambda u, i: mo(u, 'Assistencia medica', i)
ano_cheio = {u: mo(u, i=1) - mo(u, i=0) for u in UNITS + ['GER']}
reaj_sal = sal('ALL', 2) - sal('ALL', 1)
reaj_med = med('ALL', 2) - med('ALL', 1)
estrutura = sum(ano_cheio.values())
risco_mo = 0.95 + 0.46 + 0.36           # Auditoria Labour (achados 1, 3 e 7)
CL = g.CL
sh_ns = [g.C_NS[i] / g.C_NS[3] for i in range(3)]
sh_eb = [g.C_EBIT[i] / g.C_EBIT[3] for i in range(3)]
fiat_br = g.C_EBIT[0] / 1000 / e27

# volumes MP'27 por projeto (PO's\Volume\MP'27_ADJ)
wv = openpyxl.load_workbook(VOL, data_only=True)


def vol_row(sheet, label, after=None):
    ws = wv[sheet]
    start = 1
    if after:
        start = next(r for r in range(1, ws.max_row + 1) if str(ws.cell(r, 2).value or '').strip() == after)
    for r in range(start, ws.max_row + 1):
        if str(ws.cell(r, 1).value or '').strip().lower() == label.lower():
            return float(ws.cell(r, 14).value)   # coluna N = total do ano
    raise KeyError((sheet, label))


TUKAN = vol_row('SJP - VW', 'Produção (pç Tukan)', "MP'27") / 1000
VIRTUS = vol_row('SJP - VW', 'Produção (pç Virtus)', "MP'27") / 1000
NISSAN = (vol_row('Resende - Nissan', 'Produção pç Kait', "MP'27") + vol_row('Resende - Nissan', 'Produção pç Kicks', "MP'27")) / 1000
assert abs(NISSAN - R('RES', 'pcs')) < 1, (NISSAN, R('RES', 'pcs'))
assert abs(g.C_PCS[2] - g.JLR['CNH'][1] / 1000) < 0.5
NB_PCS = NISSAN + TUKAN + g.C_PCS[2]

# ================================================================ slide 4 - resumo: decisoes
s = by_title("Fitted Units Economics Brasil – Resumo executivo MP'27")
tb = shape(s, 'Bloco 3 texto')
tb.height = Emu(1260000)
items = ['Decisões necessárias', f"Fechar o reajuste Fiat ({f1(g.FIAT_PRECO)} mi)",
         f"Repassar a ANTT ao cliente ({f1(antt)} mi)", f"Validar riscos de MO com o RH (até {f1(risco_mo)} mi)"]
body = tb.text_frame._txBody
for p in tb.text_frame.paragraphs[2:]:          # sobra vazia da edicao manual
    body.remove(p._p)
bullet = tb.text_frame.paragraphs[1]._p
for _ in items[2:]:
    body.append(copy.deepcopy(bullet))
for p, t in zip(tb.text_frame.paragraphs, items):
    p.runs[0].text = t
    for r in p.runs[1:]:
        r._r.getparent().remove(r._r)

# ================================================================ slide 5 - IBI
s = by_title('Fitted Units Economics Brasil – IBI')
comments(s, "Ibirité (IBI)", [
    g.head_kpi('IBI'),
    (f"Saída da Jaguar {f1(dd('IBI', 'Encerr. Jaguar'))}", [f"R$ {f1(g.JLR['fat'] / 1000)} mi de receita saem do perímetro"]),
    (f"3º turno Fiat: {f1(T3_MO + T3_DESP)} mi/ano", [f"Mão de obra: {f1(T3_MO)} mi", f"Demais despesas: {f1(T3_DESP)} mi"]),
    (f"ANTT: custo extra de {f1(antt)} mi", ["Nova regra de frete mínimo"]),
    (f"Preço adicional Fiat {pc(fiat_pct, signed=True)}/pç", [
        f"{sg(g.FIAT_PRECO)} mi: cobre {pc(cobertura, 0)} do 3º turno e ANTT",
        f"Sem o reajuste, EBIT de IBI ~{f1(g.ebit_sem_fiat_ibi)}"]),
])

# ================================================================ slide 6 - rentabilidade por cliente
s = by_title('Fitted Units Economics Brasil – IBI | Rentabilidade por cliente')
keep = {'Title 1', 'Picture 163', 'Mensagem principal'}
drop(s, lambda x: x.name not in keep)
set_text(shape(s, 'Mensagem principal'),
         f"A Fiat traz {pc(sh_ns[0], 0)} da receita e {pc(sh_eb[0], 0)} do EBIT de Ibirité, com ROS de {pc(g.C_ROS[0])} "
         f"(meta {pc(g.ROS_META)}); Iveco e CNH têm margem alta em volume pequeno")
g.txt(s.shapes, 300000, 1150000, 5000000, 240000, "Participação no total de IBI e ROS por cliente", 11, True, HEAD,
      PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE, name='Titulo grafico')
CW, CG, CY, CH = 2600000, 175000, 1480000, 2700000
TRACK = RGBColor(0xD9, 0xD9, 0xD9)
for k, c in enumerate(CL):
    x0 = 300000 + k * (CW + CG)
    cx = x0 + CW // 2
    g.box(s.shapes, x0, CY, CW, CH, PANEL, f'Card {c}', MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.05
    logo(s, c.lower(), cx, CY + 140000, 420000, 1500000)
    rich(s, x0, CY + 650000, CW, 480000, [[('ROS ', 13, False, DGREY), (pc(g.C_ROS[k]), 24, True, HEAD)]],
         f'ROS {c}', PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    meta = f"meta {pc(g.ROS_META)}" if c == 'Fiat' else f"{fp(g.C_PCS[k], 0)}k pçs"
    g.txt(s.shapes, x0, CY + 1130000, CW, 220000, meta if c != 'Fiat' else f"{meta} · {fp(g.C_PCS[k], 0)}k pçs", 10, False, DGREY,
          anchor=MSO_ANCHOR.TOP, name=f'Sub {c}')
    # barras de participacao (trilho = 100% do total IBI)
    tx, tw = x0 + 180000, CW - 360000
    for j, (v, col, lab) in enumerate(((sh_ns[k], LIGHT, '% da receita IBI'), (sh_eb[k], DARK, '% do EBIT IBI'))):
        y = CY + 1520000 + j * 560000
        g.txt(s.shapes, tx, y, tw - 600000, 200000, lab, 10, False, DGREY, PP_ALIGN.LEFT, MSO_ANCHOR.BOTTOM, f'Rotulo {c} {lab}')
        g.txt(s.shapes, tx + tw - 600000, y - 20000, 600000, 220000, pc(v, 0), 14, True, HEAD, PP_ALIGN.RIGHT,
              MSO_ANCHOR.BOTTOM, f'Valor {c} {lab}')
        g.box(s.shapes, tx, y + 240000, tw, 150000, TRACK, f'Trilho {c} {lab}')
        b = g.box(s.shapes, tx, y + 240000, max(int(tw * v), 25400), 150000, col, f'Barra {c} {lab}')
        b.line.color.rgb = DARK
        b.line.width = Pt(0.5)
k1 = lambda v: fp(v / 1000.0)
rows = [["mln BRL", "Fiat", "Iveco", "CNH", "Total IBI"],
        ["Peças (mil)"] + [fp(v, 0) for v in g.C_PCS],
        ["Net Sales"] + [k1(v) for v in g.C_NS],
        ["NS/pç (R$)"] + [fp(v, 2) for v in g.C_NSPC],
        ["CM"] + [k1(v) for v in g.C_CM],
        ["Fixed Cost"] + [k1(v) for v in g.C_FIX],
        ["EBIT"] + [k1(v) for v in g.C_EBIT],
        ["ROS %"] + [pc(v) for v in g.C_ROS]]
g.simple_table(s, 300000, 4330000, [2150000, 1500000, 1500000, 1500000, 1500000], rows, 230000, 10)
g.footnote(s, "Fonte: Rentabilidade por Cliente IBI – MP27 v3 (19/10, ligada à MENS v3). Faturamento e material de consumo diretos; "
              f"frete 100% Fiat; demais custos rateados por volume. Inclui o custo dedicado da CNH ({f1(g.CNH_ESPEC, 2)} mi/ano).", 6230000)
cnh_zero = [m for m in range(12) if g.CNH_VOL[m] == 0]
MES = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']
cnh_perda = -sum(g.CNH_MES[m] for m in cnh_zero) / len(cnh_zero)
comments(s, "Clientes IBI", [
    ('Fiat: escala do negócio', [f"{pc(sh_ns[0], 0)} da receita e {pc(sh_eb[0], 0)} do EBIT",
                                 f"ROS {pc(g.C_ROS[0])}, acima da meta de {pc(g.ROS_META)}"]),
    ('Iveco: margem alta, volume menor', [f"ROS {pc(g.C_ROS[1])}",
                                          f"Volume {pc(g.IVECO_V27 / g.IVECO_V9 - 1, signed=True)} vs R9"]),
    ('CNH: margem alta, volume irregular', [f"ROS {pc(g.C_ROS[2])}, com custo dedicado",
                                            f"{' e '.join(MES[m] for m in cnh_zero).capitalize()} sem volume: −R$ {f1(cnh_perda, 0)}k/mês"]),
    ('Leitura', ["Iveco e CNH: 7% da receita, 22% do EBIT"]),
])
assert round((sh_ns[1] + sh_ns[2]) * 100) == 7 and round((sh_eb[1] + sh_eb[2]) * 100) == 22

# ================================================================ slide 10 - mao de obra
s = by_title("Fitted Units Economics Brasil – Mão de obra MP'27")
keep = {'Title 1', 'Picture 163', 'Mensagem principal'}
drop(s, lambda x: x.name not in keep)
set_text(shape(s, 'Mensagem principal'),
         f"Mão de obra {sg(mo27 - mo9)} mi ({pc(mo27 / mo9 - 1, signed=True)}): {f1(estrutura)} é a estrutura aberta em 2026 "
         f"em ano cheio e {f1(reaj_sal + reaj_med)} o reajuste de 2027 ({pc(mo27 / mo('ALL', i=1) - 1, signed=True)})")
g.txt(s.shapes, 300000, 1110000, 8150000, 220000, "Bridge da mão de obra – R$ mi (custo direto)", 11, True, HEAD,
      PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE, name='Titulo bridge MO')
spec = {'start': mo9, 'end': mo27, 'deltas': [
    ('IBI\n3º turno', ano_cheio['IBI']), ('RES\nano cheio', ano_cheio['RES']),
    ('Demais\nunidades', ano_cheio['SJP'] + ano_cheio['GOI'] + ano_cheio['GER']),
    ('Reajuste sal.\n+ encargos', reaj_sal), ('Assist.\nmédica', reaj_med)]}
assert abs(spec['start'] + sum(v for _, v in spec['deltas']) - spec['end']) < 1e-6
g.X0, g.X1, g.Y_BR, g.YT, g.YB = 300000, 8450000, 1520000, 1900000, 3600000
OFF = 22.0      # eixo truncado: as barras de total comecam em 22 mi
grp = g.draw_bridge(s, dict(spec, start=spec['start'] - OFF, end=spec['end'] - OFF))
grp.name = 'Bridge MO'
for sh in grp.shapes:
    if sh.name == 'Valor Fcst R9':
        set_text(sh, f1(mo9))
    elif sh.name == 'Valor MP’27':
        set_text(sh, f1(mo27))
    elif sh.name == 'Variacao':
        set_text(sh, pc(mo27 / mo9 - 1, signed=True))
    elif sh.name in ('Barra Fcst R9', 'Barra MP’27'):   # corte de eixo
        yb = sh.top + sh.height - 260000
        g.box(s.shapes, sh.left - 30000, yb, sh.width + 60000, 70000, WHITE, 'Corte de eixo')
        for dy in (0, 70000):
            g.line(s.shapes, sh.left - 60000, yb + dy + 25000, sh.left + sh.width + 60000, yb + dy - 25000, DGREY, 0.75,
                   name='Corte de eixo linha')
g.box(s.shapes, 300000 + int((8450000 - 300000) / 7 * 1) + 60000, 4120000, int((8450000 - 300000) / 7 * 3) - 120000, 30000, GREY, 'Chave estrutura')
g.box(s.shapes, 300000 + int((8450000 - 300000) / 7 * 4) + 60000, 4120000, int((8450000 - 300000) / 7 * 2) - 120000, 30000, GREY, 'Chave reajuste')
g.txt(s.shapes, 300000 + int((8450000 - 300000) / 7 * 1), 4170000, int((8450000 - 300000) / 7 * 3), 200000,
      f"Estrutura 2026 em ano cheio {sg(estrutura)}", 9.5, True, DGREY, anchor=MSO_ANCHOR.TOP, name='Rotulo estrutura')
g.txt(s.shapes, 300000 + int((8450000 - 300000) / 7 * 4), 4170000, int((8450000 - 300000) / 7 * 2), 200000,
      f"Reajuste 2027 {sg(reaj_sal + reaj_med)}", 9.5, True, DGREY, anchor=MSO_ANCHOR.TOP, name='Rotulo reajuste')
rows = [["R$ mi", "Fcst R9", "Estrutura ano cheio", "Reajuste 2027", "MP'27", "Δ total", "Δ %"]]
for u in UNITS + ['GER', 'ALL']:
    a, rr, c = mo(u, i=0), mo(u, i=1), mo(u)
    rows.append(['Total Fitted' if u == 'ALL' else ('Gerência' if u == 'GER' else u), f1(a), sg(rr - a), sg(c - rr), f1(c),
                 sg(c - a), pc(c / a - 1, signed=True)])
g.simple_table(s, 300000, 4500000, [1550000, 1000000, 1500000, 1250000, 1000000, 950000, 900000], rows, 235000, 10, bold_last=True)
g.footnote(s, "Fonte: Labour Cost FITTED MP27 v2 e Fcst R9. Estrutura ano cheio = custo mensal de set–dez/26 × 12 menos o R9 "
              "(3º turno IBI e RES abertos ao longo de 2026). Reajuste 2027 = MP'27 menos esse custo anualizado.", 6230000)
comments(s, "Mão de obra", [
    (f"Estrutura 2026 em ano cheio {sg(estrutura)}", [
        f"IBI {sg(ano_cheio['IBI'])}: 3º turno iniciado em 2026",
        f"RES {sg(ano_cheio['RES'])}: unidade aberta em 2026"]),
    (f"Reajuste 2027 {sg(reaj_sal + reaj_med)} ({pc(mo27 / mo('ALL', i=1) - 1, signed=True)})", [
        f"Salários e encargos {pc(sal('ALL', 2) / sal('ALL', 1) - 1, signed=True)}",
        f"Assistência médica {pc(med('ALL', 2) / med('ALL', 1) - 1, signed=True)}"]),
    ('Leitura', [f"Sem a estrutura nova, a MO sobe {pc(mo27 / mo('ALL', i=1) - 1)}",
                 f"3º turno Fiat: MO de {f1(T3_MO)} mi/ano"]),
])

# ================================================================ mix e perguntas: participacao da Fiat (nova rentabilidade)
s = by_title('Fitted Units Economics Brasil – Mix do EBIT por unidade')
for p in shape(s, 'Comentarios').text_frame.paragraphs:
    for r in p.runs:
        if r.text.startswith('Fiat = '):
            r.text = f"Fiat = {pc(fiat_br, 0)} do EBIT do Brasil"
s = by_title("Fitted Units Economics Brasil – Perguntas-chave do MP'27")
set_text(shape(s, 'Card 5 resposta'),
         f"Fiat: ROS {pc(g.C_ROS[0])}, perto da meta de {pc(g.ROS_META)}, com {pc(sh_ns[0], 0)} da receita de IBI. "
         f"Iveco: volume {pc(g.IVECO_V27 / g.IVECO_V9 - 1, signed=True)}. CNH: meses sem volume.")
qa = s


# ================================================================ novo slide: New Business
def card(sl, x, y, w, h, name):
    g.box(sl.shapes, x, y, w, h, PANEL, name, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.05


nb = g.new_slide("Fitted Units Economics Brasil – New Business")
g.headline(nb, f"New business soma {fp(NB_PCS, 0)}k peças em 2027 ({pc(NB_PCS / R('BR', 'pcs'), 0)} do volume): "
               f"Nissan em RES, Tukan em SJP e CNH em IBI", w=11590000)
NB = [
    ('nissan', 'Nissan – Resende', 'Kait e Kicks · 1º ano cheio', f"{fp(NISSAN, 0)}k pçs",
     f"{pc(NISSAN / R('RES', 'pcs', 3) - 1, 0, True)} vs R9",
     [f"Net Sales R$ {f1(R('RES', 'ns'))} mi", f"EBIT R$ {f1(R('RES', 'ebit'))} mi · ROS {pc(R('RES', 'ros'), 0)}"]),
    ('vw', 'VW – São José dos Pinhais', 'Lançamento do Tukan', f"{fp(TUKAN, 0)}k pçs",
     f"{pc(TUKAN / R('SJP', 'pcs'), 0)} do volume de SJP",
     [f"Substitui o Virtus ({fp(VIRTUS, 0)}k pçs, até fev/27)", f"Volume de SJP estável: {fp(R('SJP', 'pcs'), 0)}k pçs"]),
    ('cnh', 'CNH – Ibirité', 'Cliente iniciado em 2026', f"{fp(g.C_PCS[2], 0)}k pçs",
     f"{pc(g.JLR['CNH'][1] / g.JLR['CNH'][0] - 1, 0, True)} vs R9",
     [f"Net Sales R$ {f1(g.C_NS[2] / 1000)} mi", f"EBIT R$ {f1(g.C_EBIT[2] / 1000)} mi · ROS {pc(g.C_ROS[2], 0)}"]),
]
W3, G3, Y3, H3 = 3730000, 200000, 1300000, 2950000
for k, (lg, tit, sub, big, bsub, lines) in enumerate(NB):
    x = 300000 + k * (W3 + G3)
    card(nb, x, Y3, W3, H3, f'Card NB {k + 1}')
    logo(nb, lg, x + 520000, Y3 + 170000, 560000, 800000)
    rich(nb, x + 1000000, Y3 + 170000, W3 - 1150000, 560000,
         [[(tit, 14, True, HEAD)], [(sub, 11, False, DGREY)]], f'Card NB {k + 1} titulo', anchor=MSO_ANCHOR.MIDDLE)
    rich(nb, x + 180000, Y3 + 950000, W3 - 360000, 700000,
         [[(big, 28, True, HEAD)], [(bsub, 11, True, GREEN)]], f'Card NB {k + 1} volume')
    rich(nb, x + 180000, Y3 + 1850000, W3 - 360000, 900000,
         [[('•  ' + t, 12, False, BLACK)] for t in lines], f'Card NB {k + 1} texto')
YS = 4500000
g.txt(nb.shapes, 300000, YS, 6000000, 260000, "Saídas do perímetro", 12, True, HEAD, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE, 'Titulo saidas')
SAI = [("Jaguar – Ibirité", f"Encerrado em 2026: EBIT {f1(dd('IBI', 'Encerr. Jaguar'))} e R$ {f1(g.JLR['fat'] / 1000)} mi de receita"),
       ("Toyota – Sorocaba", f"Unidade desmobilizada: {f1(R('BR', 'ebit', 3) - sum(R(u, 'ebit', 3) for u in UNITS))} de outras receitas no R9")]
for k, (tit, txt_) in enumerate(SAI):
    x = 300000 + k * (5795000 + 0)
    w = 5695000
    card(nb, x, YS + 330000, w, 780000, f'Card saida {k + 1}')
    b = g.box(nb.shapes, x + 150000, YS + 330000 + 230000, 320000, 320000, RED, f'Icone saida {k + 1}', MSO_SHAPE.OVAL)
    tf = b.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    r = tf.paragraphs[0].add_run()
    r.text = '−'
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = WHITE
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    rich(nb, x + 600000, YS + 330000, w - 750000, 780000, [[(tit, 13, True, HEAD)], [(txt_, 11, False, BLACK)]],
         f'Card saida {k + 1} texto', anchor=MSO_ANCHOR.MIDDLE)
g.footnote(nb, "Fonte: Volumes MP'27 (PO's\\Volume\\MP'27_ADJ, 19/10), Base Book Fitted Units MP27 e Rentabilidade por Cliente IBI v3. "
               "% vs R9 da Nissan reflete a abertura de RES ao longo de 2026.", y=6000000, w=11590000)

# ================================================================ novo slide: Risks & Opportunities
ro = g.new_slide("Fitted Units Economics Brasil – Risks & Opportunities")
RISKS = [("Reajuste Fiat não aprovado", f"Preço {pc(fiat_pct, signed=True)}/pç ainda em negociação", -g.FIAT_PRECO, ''),
         ("Provisão de férias do horista", "Premissa ~5,5% vs 11,1% do estatutário", -0.95, 'até '),
         ("Horista fixo SJP (CC 8289)", "+35% sem premissa visível", -0.46, ''),
         ("Assistência médica GOI", "3,4% do salário vs 31% em IBI", -0.36, 'até ')]
OPPS = [("Repasse da ANTT ao cliente", "Frete mínimo absorvido no plano", antt, ''),
        ("Completar o repasse do 3º turno", f"Preço Fiat cobre {pc(cobertura, 0)} de 3º turno + ANTT",
         T3_MO + T3_DESP + antt - g.FIAT_PRECO, '')]
tot_r, tot_o = sum(v for *_, v, _ in RISKS), sum(v for *_, v, _ in OPPS)
g.headline(ro, f"Riscos quantificados de {f1(tot_r)} mi, concentrados no reajuste Fiat; oportunidades de {sg(tot_o)} mi",
           w=11590000)
CW2, XR, XO, YH = 5695000, 300000, 6195000, 1300000
for x, title, col, lst, tot in ((XR, 'Riscos', RED, RISKS, tot_r), (XO, 'Oportunidades', GREEN, OPPS, tot_o)):
    hd = g.box(ro.shapes, x, YH, CW2, 380000, col, f'Cabecalho {title}')
    tf = hd.text_frame
    tf.margin_left = Emu(150000)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    r = tf.paragraphs[0].add_run()
    r.text = title
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = WHITE
    g.txt(ro.shapes, x + CW2 - 1900000, YH, 1750000, 380000, 'EBIT (mln BRL)', 10, False, WHITE, PP_ALIGN.RIGHT,
          MSO_ANCHOR.MIDDLE, f'Cabecalho {title} unidade')
    y = YH + 460000
    for k, (it, det, v, pre) in enumerate(lst + [("Volume", "Cada 1% de peças", (1 if col == GREEN else -1) * g.SENS, '')]):
        is_vol = it == 'Volume'
        card(ro, x, y, CW2, 640000, f'{title} item {k + 1}')
        rich(ro, x + 150000, y, CW2 - 2000000, 640000, [[(it, 12.5, True, HEAD)], [(det, 10.5, False, DGREY)]],
             f'{title} item {k + 1} texto', anchor=MSO_ANCHOR.MIDDLE)
        val = pre + (f"{sg(v)} / 1%" if is_vol else sg(v, 2 if abs(v) < 1 else 1))
        g.txt(ro.shapes, x + CW2 - 1900000, y, 1750000, 640000, val, 16, True, col, PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE,
              f'{title} item {k + 1} valor')
        y += 720000
    yt = YH + 460000 + 5 * 720000
    g.line(ro.shapes, x, yt - 30000, x + CW2, yt - 30000, GREY, 1.0, name=f'Linha total {title}')
    g.txt(ro.shapes, x + 150000, yt, CW2 - 2000000, 330000, 'Total quantificado (sem volume)', 12, True, BLACK,
          PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE, f'Total {title} rotulo')
    g.txt(ro.shapes, x + CW2 - 1900000, yt, 1750000, 330000, sg(tot), 16, True, col, PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE,
          f'Total {title} valor')
g.footnote(ro, "Fonte: Base Book e MENS FITTED MP27 v3 (reajuste Fiat, ANTT, sensibilidade a volume = margem variável por 1% de peças); "
               "Auditoria Labour FITTED MP27 (riscos de MO, hipóteses a confirmar com o RH); 3º turno Fiat = MO 2,5 + demais 2,0 mi/ano.",
           y=6230000, w=11590000)

# ordem: New Business e R&O antes das Perguntas-chave
lst = prs.slides._sldIdLst
ids = list(lst)
qa_id = next(e for e in ids if prs.slides.get(int(e.get('id'))) is qa)
for el in ids[-2:]:
    lst.remove(el)
pos = list(lst).index(qa_id)
for el in reversed(ids[-2:]):
    lst.insert(pos, el)

prs.save(g.OUT)
print('salvo', g.OUT, len(prs.slides), 'slides | NB pcs', round(NB_PCS), 'Tukan', TUKAN, 'Nissan', NISSAN,
      '| MO estrutura', round(estrutura, 2), 'reaj', round(reaj_sal, 2), round(reaj_med, 2), '| riscos', round(tot_r, 2), 'opps', round(tot_o, 2))
