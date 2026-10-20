"""PPT MP27 Fitted Units - v3 (sobre a v2 salva pela Juliana, com as tabelas recoladas por ela).

Uso: python gerar_ppt_mp27_v3.py <deck_v2_juliana.pptx> <basebook.xlsx> <mens.xlsx> <rentab_cliente_ibi.xlsx>
                                 <variacao_mo.xlsx> <detalhe_despesas.xlsx> <saida.pptx>

Nao mexe nas imagens de tabela coladas por ela. Altera:
- comentarios de todos os slides (curtos, uma ideia por linha);
- barras de variacao dos bridges na cor neutra (sem azul nas positivas);
- grafico de peso por cliente (formas nativas; o grafico 100% empilhado renderizava fora de proporcao);
- slide de Mao de obra sem run-rate (R9 x MP'27 e composicao do aumento);
- resumo executivo e perguntas-chave alinhados aos novos textos.
Os dados e helpers vem de gerar_ppt_mp27_rev2.py (mesmos asserts de conciliacao).
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gerar_ppt_mp27_rev2 as g  # noqa: E402  (carrega dados e abre o deck de entrada)
from gerar_ppt_mp27_rev2 import (ACC, BLACK, CCOL, CL, DARK, GREY, HEAD, LIGHT, WHITE, Emu, MSO_ANCHOR,  # noqa: E402
                                 MSO_SHAPE, MSO_SHAPE_TYPE, PP_ALIGN, Pt, R, dd, f1, fp, pc, sg)

prs = g.prs
mo = g.mo
UNITS = g.UNITS


def by_title(text):
    for s in prs.slides:
        t = s.shapes.title
        if t is not None and t.text_frame.text.strip() == text:
            return s
    raise KeyError(text)


def remove(slide, names):
    for sh in [x for x in slide.shapes if x.name in names]:
        sh._element.getparent().remove(sh._element)


def shape(slide, name):
    return [x for x in slide.shapes if x.name == name][0]


def set_text(sh, text):
    """troca o texto mantendo a formatacao do primeiro run."""
    p = sh.text_frame.paragraphs[0]
    p.runs[0].text = text
    for r in p.runs[1:]:
        r._r.getparent().remove(r._r)


def comments(slide, title, blocks, size=10.5):
    remove(slide, ('Painel comentarios', 'Painel titulo', 'Comentarios'))
    g.draw_comments(slide, title, blocks, size)


# ================================================================ numeros
BR = g.PL['BR']
e9, e27, e26 = BR['ebit'][3], BR['ebit'][4], BR['ebit'][2]
fiat_pct = g.FIAT_PR27 / g.FIAT_PR9 - 1
mo9, mo27 = mo('ALL', i=0), mo('ALL')
d_mo = mo27 - mo9
sal_enc = lambda u, i=2: mo(u, 'Salarios e PPR', i) + mo(u, 'Encargos', i)
med = lambda u, i=2: mo(u, 'Assistencia medica', i)
sh_sal = (sal_enc('ALL') - sal_enc('ALL', 0)) / d_mo
sh_med = (med('ALL') - med('ALL', 0)) / d_mo
g_med = med('ALL') / med('ALL', 0) - 1
d_mo_u = {u: mo(u) - mo(u, i=0) for u in UNITS + ['GER']}
mo_ibi_pct = mo('IBI') / mo('IBI', i=0) - 1
mo_sjp_pct = mo('SJP') / mo('SJP', i=0) - 1
ebit_t = sum(R(u, 'ebit') for u in UNITS)
shm = {u: R(u, 'ebit') / ebit_t for u in UNITS}
fiat_br = g.C_EBIT[0] / 1000 / e27
cl_eb = [g.C_EBIT[i] / g.C_EBIT[3] for i in range(3)]

# ================================================================ 1. bridges: barras de variacao na cor neutra
n_rec = 0
for s in prs.slides:
    for grp in s.shapes:
        if grp.shape_type == MSO_SHAPE_TYPE.GROUP and grp.name.startswith('Bridge'):
            for b in grp.shapes:
                if b.name.startswith('Barra') and b.fill.type is not None and b.fill.fore_color.rgb == ACC:
                    b.fill.fore_color.rgb = LIGHT
                    n_rec += 1

# ================================================================ 2. comentarios dos slides de unidade
UNIT_TXT = {
    'Fitted Units Economics Brasil': ("Brasil", [
        g.head_kpi('BR'),
        (f"Receita {pc(BR['ns'][4] / BR['ns'][3] - 1, signed=True)}", [
            f"Reajuste Fiat em IBI: Price/Mix {sg(dd('BR', 'Price/ Mix'))}",
            f"RES em ano cheio: Volume {sg(dd('BR', 'Volume'))}"]),
        ('Custos e eventos pressionam', [
            f"Mão de obra {f1(dd('BR', 'Labour'))}: 3º turno IBI e RES",
            f"Fixos {f1(dd('BR', 'Cost Fix.'))}: serviços do 3º turno IBI",
            f"ANTT {f1(dd('BR', 'ANTT'))} · Jaguar {f1(dd('BR', 'Encerr. Jaguar'))} · variáveis {f1(dd('BR', 'Cost Var.'))}"]),
        ('Ponto de atenção', [
            f"Sem o reajuste Fiat, EBIT de ~{f1(g.ebit_sem_fiat_br)}"]),
    ]),
    'Fitted Units Economics Brasil – IBI': ("Ibirité (IBI)", [
        g.head_kpi('IBI'),
        (f"Saída da Jaguar {f1(dd('IBI', 'Encerr. Jaguar'))}", [
            f"R$ {f1(g.JLR['fat'] / 1000)} mi de receita saem do perímetro"]),
        ('Início do 3º turno', [
            f"Mão de obra {sg(d_mo_u['IBI'])} mi ({pc(mo_ibi_pct, signed=True)} vs R9)",
            f"Salários e encargos {sg(sal_enc('IBI') - sal_enc('IBI', 0))} · assist. médica {sg(med('IBI') - med('IBI', 0))}",
            f"Serviços de apoio ao turno: {sg(g.T3_TOT)} mi"]),
        ('Preço adicional cobre o 3º turno', [
            f"Fiat {pc(fiat_pct, signed=True)}/pç: {sg(g.FIAT_PRECO)} mi de receita",
            f"Sem o reajuste, EBIT de IBI ~{f1(g.ebit_sem_fiat_ibi)}"]),
    ]),
    'Fitted Units Economics Brasil – SJP': ("SJP", [
        g.head_kpi('SJP'),
        ('Desconto de −3,5% ao cliente', [
            f"R$ {f1(g.SJP_PRECO_ATUAL, 2)} → {f1(g.SJP_PRECO_27, 2)}/pç: ≈ {f1(g.SJP_DESC_MI)} mi/ano",
            f"Parte compensada: Price/Mix só {f1(dd('SJP', 'Price/ Mix'))} vs R9"]),
        (f"Mão de obra {f1(dd('SJP', 'Labour'))} ({pc(mo_sjp_pct, signed=True)})", [
            f"{pc(g.sjp_sal, 0)} do aumento: salários e encargos",
            f"{pc(g.sjp_med, 0)}: assistência médica ({pc(g.sjp_med_g, 0, True)} vs R9)"]),
        ('Volume e demais custos estáveis', [
            f"{fp(R('SJP', 'pcs'), 0)}k pçs, em linha com o R9"]),
    ]),
    'Fitted Units Economics Brasil – GOI': ("Goiana (GOI)", [
        g.head_kpi('GOI'),
        ('Preço compensa o volume', [
            f"Preço por peça {pc(R('GOI', 'nspc') / R('GOI', 'nspc', 3) - 1, signed=True)}: Price/Mix {sg(dd('GOI', 'Price/ Mix'))}",
            f"Volume {pc(R('GOI', 'pcs') / R('GOI', 'pcs', 3) - 1, signed=True)}: {f1(dd('GOI', 'Volume'))}"]),
        ('Custos fixos e variáveis equalizados', [
            f"Fixos {sg(dd('GOI', 'Cost Fix.'))} vs R9",
            f"Compensam MO {f1(dd('GOI', 'Labour'))} e variáveis {f1(dd('GOI', 'Cost Var.'))}"]),
    ]),
    'Fitted Units Economics Brasil – RES': ("Resende (RES)", [
        g.head_kpi('RES'),
        ('Delta de perímetro', [
            "Unidade não operou o ano todo em 2026",
            f"{fp(R('RES', 'pcs'), 0)}k pçs no MP'27 vs {fp(R('RES', 'pcs', 3), 0)}k no R9"]),
        (f"ROS de {pc(R('RES', 'ros'), 0)} já no 1º ano cheio", [
            f"Preço R$ {f1(R('RES', 'nspc'), 2)}/pç (R9: {f1(R('RES', 'nspc', 3), 2)})",
            f"Cobre MO {f1(dd('RES', 'Labour'))} e fixos {f1(dd('RES', 'Cost Fix.'))}"]),
    ]),
}
for title, (ptitle, blocks) in UNIT_TXT.items():
    comments(by_title(title), ptitle, blocks)

# ================================================================ 3. mix do EBIT
s = by_title('Fitted Units Economics Brasil – Mix do EBIT por unidade')
comments(s, "Mix do EBIT", [
    (f"IBI: {pc(shm['IBI'], 0)} do EBIT", [f"Âncora: {pc(R('IBI', 'ns') / BR['ns'][4], 0)} do Net Sales do Brasil"]),
    (f"SJP: {pc(shm['SJP'], 0)}", ["Perde peso: desconto ao cliente e MO"]),
    (f"GOI: {pc(shm['GOI'], 0)}", [f"EBIT estável em {f1(R('GOI', 'ebit'))}"]),
    (f"RES: {pc(shm['RES'], 0)}", [f"Passa a contribuir, com ROS de {pc(R('RES', 'ros'), 0)}"]),
    ('Concentração', [f"Fiat = {pc(fiat_br, 0)} do EBIT do Brasil"]),
])

# ================================================================ 4. resumo executivo
s = by_title("Fitted Units Economics Brasil – Resumo executivo MP'27")
set_text(shape(s, 'Mensagem principal'),
         f"EBIT de R$ {f1(e27)} mi (ROS {pc(BR['ros'][4])}): o preço compensa parte da pressão de MO, 3º turno e ANTT")
MSG = {'IBI': "Preço Fiat cobre 3º turno e Jaguar", 'SJP': "Desconto −3,5% ao cliente; MO +14%",
       'GOI': "Preço compensa volume", 'RES': "Ano cheio, ROS de 10%", 'Brasil': "R9 inclui 0,2 da SOR (encerrada)"}
tb = shape(s, 'Table 74').table
for i in range(1, len(tb.rows)):
    key = tb.cell(i, 0).text_frame.text.replace('●', '').strip()
    set_text(tb.cell(i, 8), MSG[key])
BLOCOS = {
    'Bloco 1 texto': ('O que sustenta o plano', [
        f"Preço: {sg(dd('BR', 'Price/ Mix'))} (reajuste Fiat)",
        f"RES em ano cheio: ROS de {pc(R('RES', 'ros'), 0)}",
        f"GOI: EBIT estável em {f1(R('GOI', 'ebit'))}"]),
    'Bloco 2 texto': ('O que pressiona', [
        f"Mão de obra: {f1(dd('BR', 'Labour'))} ({pc(mo27 / mo9 - 1, signed=True)})",
        f"3º turno IBI: serviços +{f1(g.T3_TOT)} mi",
        f"ANTT {f1(dd('BR', 'ANTT'))} · Jaguar {f1(dd('BR', 'Encerr. Jaguar'))}"]),
    'Bloco 3 texto': ('Para decidir', [
        f"Fechar o reajuste Fiat ({f1(g.FIAT_PRECO)} mi)",
        f"Gap de {f1(g.EBIT_REF - e27)} mi para ROS de {pc(g.ROS_REF)}",
        f"1% de volume ≈ {f1(g.SENS)} mi de EBIT"]),
}
for nm, (lead, items) in BLOCOS.items():
    tf = shape(s, nm).text_frame
    ps = tf.paragraphs
    assert len(ps) == 1 + len(items), nm
    for p, it in zip(ps, [lead] + items):
        p.runs[0].text = it
        for r in p.runs[1:]:
            r._r.getparent().remove(r._r)

# ================================================================ 5. rentabilidade por cliente: peso em formas nativas
s = by_title('Fitted Units Economics Brasil – IBI | Rentabilidade por cliente')
remove(s, ('Peso por cliente',))
shares = [('Peças', [g.C_PCS[i] / g.C_PCS[3] for i in range(3)]),
          ('Net Sales', [g.C_NS[i] / g.C_NS[3] for i in range(3)]),
          ('EBIT', cl_eb)]
LX, BX, BW, BH = 300000, 1150000, 3150000, 380000
for k, c in enumerate(CL):    # legenda
    x = BX + k * 900000
    g.box(s.shapes, x, 1480000, 130000, 130000, CCOL[c], f'Legenda {c}')
    g.txt(s.shapes, x + 180000, 1440000, 650000, 210000, c, 10, False, BLACK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE, f'Legenda {c} texto')
for j, (lab, v) in enumerate(shares):
    y = 1800000 + j * 720000
    g.txt(s.shapes, LX, y, 800000, BH, lab, 11, True, BLACK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE, f'Peso {lab} rotulo')
    x = BX
    for i, c in enumerate(CL):
        w = int(round(BW * v[i]))
        b = g.box(s.shapes, x, y, max(w, 12700), BH, CCOL[c], f'Peso {lab} {c}')
        b.line.color.rgb = WHITE
        b.line.width = Pt(0.75)
        if i == 0:
            tf = b.text_frame
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = pc(v[0], 0)
            r.font.size = Pt(12)
            r.font.bold = True
            r.font.color.rgb = WHITE
        x += w
    g.txt(s.shapes, BX, y + BH + 30000, BW, 200000, f"Iveco {pc(v[1], 0)} · CNH {pc(v[2], 0)}", 9, False, GREY,
          PP_ALIGN.RIGHT, MSO_ANCHOR.TOP, f'Peso {lab} demais')
cnh_zero = [k for k in range(12) if g.CNH_VOL[k] == 0]
MES = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']
cnh_perda = -sum(g.CNH_MES[k] for k in cnh_zero) / len(cnh_zero)
comments(s, "Clientes IBI", [
    ('Fiat: escala e margem', [f"ROS {pc(g.C_ROS[0])}, acima da meta de {pc(g.ROS_META)}",
                               f"{pc(cl_eb[0], 0)} do EBIT de Ibirité"]),
    ('Iveco: margem boa, volume menor', [f"ROS {pc(g.C_ROS[1])}; volume {pc(g.IVECO_V27 / g.IVECO_V9 - 1, signed=True)} vs R9"]),
    ('CNH: ponto fraco', [f"ROS {pc(g.C_ROS[2])}: custo dedicado {f1(g.CNH_ESPEC, 2)} mi/ano",
                          f"{' e '.join(MES[k] for k in cnh_zero).capitalize()} sem volume: −R$ {f1(cnh_perda, 0)}k/mês"]),
    ('Próximo passo', ["Rever preço da CNH ou o custo dedicado"]),
])

# ================================================================ 6. mao de obra (sem run-rate)
s = by_title("Fitted Units Economics Brasil – Mão de obra MP'27")
remove(s, ('MO por unidade', 'Table 60', 'Nota de rodape', 'Titulo grafico MO'))
set_text(shape(s, 'Mensagem principal'),
         f"Mão de obra {sg(d_mo)} mi ({pc(mo27 / mo9 - 1, signed=True)}) vs R9: 3º turno de IBI e RES em ano cheio")
g.txt(s.shapes, 300000, 1130000, 8150000, 260000, "Mão de obra por unidade – R$ mi (custo direto)", 11, True, HEAD,
      PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE, name='Titulo grafico MO')
cd = g.CategoryChartData()
cd.categories = UNITS
cd.add_series('Fcst R9', [round(mo(u, i=0), 2) for u in UNITS])
cd.add_series("MP'27", [round(mo(u), 2) for u in UNITS])
gf = s.shapes.add_chart(g.XL_CHART_TYPE.COLUMN_CLUSTERED, Emu(300000), Emu(1400000), Emu(8150000), Emu(2500000), cd)
gf.name = 'MO por unidade'
ch = gf.chart
ch.has_title = False
ch.font.size = Pt(10)
ch.has_legend = True
ch.legend.position = g.XL_LEGEND_POSITION.TOP
ch.legend.include_in_layout = False
ch.legend.font.size = Pt(10)
ch.value_axis.visible = False
ch.value_axis.has_major_gridlines = False
ch.category_axis.format.line.color.rgb = GREY
ch.category_axis.tick_labels.font.size = Pt(11)
ch.category_axis.tick_labels.font.bold = True
ch.plots[0].gap_width = 90
ch.plots[0].overlap = -10
for k, col in enumerate((LIGHT, DARK)):
    ser = ch.plots[0].series[k]
    ser.format.fill.solid()
    ser.format.fill.fore_color.rgb = col
    ser.format.line.color.rgb = DARK
    ser.format.line.width = Pt(0.5)
    dl = ser.data_labels
    dl.show_value = True
    dl.number_format = '#,##0.0'
    dl.number_format_is_linked = False
    dl.font.size = Pt(10)
    dl.font.bold = k == 1
    dl.position = g.XL_LABEL_POSITION.OUTSIDE_END
rows = [["R$ mi", "Fcst R9", "MP'27", "Δ", "Δ %", "Sal. + encargos", "Assist. médica"]]
for u in UNITS + ['GER', 'ALL']:
    rows.append(['Total Fitted' if u == 'ALL' else ('Gerência' if u == 'GER' else u), f1(mo(u, i=0)), f1(mo(u)),
                 sg(mo(u) - mo(u, i=0)), pc(mo(u) / mo(u, i=0) - 1, signed=True),
                 sg(sal_enc(u) - sal_enc(u, 0)), sg(med(u) - med(u, 0))])
g.simple_table(s, 300000, 4120000, [1650000, 1050000, 1050000, 1050000, 1050000, 1200000, 1100000], rows, 255000, 10, bold_last=True)
g.footnote(s, "Fonte: Labour Cost FITTED MP27 v2 e Fcst R9 (custo direto por unidade). Δ % da RES reflete a abertura da unidade em 2026.", 6000000)
comments(s, "Mão de obra", [
    ('Onde cresce', [f"IBI {sg(d_mo_u['IBI'])}: início do 3º turno",
                     f"RES {sg(d_mo_u['RES'])}: unidade em ano cheio",
                     f"SJP {sg(d_mo_u['SJP'])} · GOI {sg(d_mo_u['GOI'])}"]),
    ('O que cresce', [f"Salários e encargos: {pc(sh_sal, 0)} do aumento",
                      f"Assist. médica: {pc(sh_med, 0)} ({pc(g_med, 0, True)} vs R9)"]),
    ('Riscos a validar com o RH', ["Férias do horista: até +0,95 mi",
                                   "Assist. médica GOI: até +0,36 mi",
                                   "CC 8289 SJP: +0,46 mi"]),
])

# ================================================================ 7. perguntas-chave
s = by_title("Fitted Units Economics Brasil – Perguntas-chave do MP'27")
CARDS = [
    f"Preço e volume somam {sg(dd('BR', 'Price/ Mix') + dd('BR', 'Volume'))}, mas os custos sobem mais: MO {f1(dd('BR', 'Labour'))}, "
    f"fixos {f1(dd('BR', 'Cost Fix.'))}, ANTT {f1(dd('BR', 'ANTT'))} e Jaguar {f1(dd('BR', 'Encerr. Jaguar'))}.",
    f"Price/Mix {sg(dd('BR', 'Price/ Mix'))}. O reajuste Fiat ({pc(fiat_pct, signed=True)}) vale {f1(g.FIAT_PRECO)} mi; "
    f"sem ele, EBIT de ~{f1(g.ebit_sem_fiat_br)} (ROS ~{pc(g.ros_sem_fiat_br, 0)}).",
    f"Cada 1% de volume ≈ R$ {f1(g.SENS)} mi de EBIT ({pc(g.SENS / e27, 0)} do EBIT): MO e handling são fixos no plano.",
    f"{sg(d_mo)} mi ({pc(mo27 / mo9 - 1, signed=True)}) vs R9: 3º turno de IBI e RES em ano cheio. "
    f"Assistência médica é o item que mais cresce ({pc(g_med, 0, True)}).",
    f"CNH: ROS {pc(g.C_ROS[2])} com custo dedicado de {f1(g.CNH_ESPEC, 2)} mi. Iveco: volume {pc(g.IVECO_V27 / g.IVECO_V9 - 1, signed=True)}. "
    f"Fiat = {pc(fiat_br, 0)} do EBIT do Brasil.",
    "A validar com o RH: férias do horista (até 0,95 mi), assist. médica GOI (até 0,36 mi) e CC 8289 SJP (0,46 mi).",
]
for k, t in enumerate(CARDS):
    set_text(shape(s, f'Card {k + 1} resposta'), t)

# ================================================================ conferencias e saida
assert n_rec >= 5, n_rec
txt_all = ' '.join(sh.text_frame.text for s in prs.slides for sh in s.shapes if sh.has_text_frame)
assert 'run-rate' not in txt_all.lower(), 'run-rate ainda aparece no deck'
prs.save(g.OUT)
print('salvo', g.OUT, '| barras recoloridas:', n_rec, '| MO', round(mo9, 2), '->', round(mo27, 2),
      '| sal+enc', round(sh_sal, 3), 'med', round(sh_med, 3), round(g_med, 3))
