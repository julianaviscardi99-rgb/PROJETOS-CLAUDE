"""PPT MP27 Fitted v9 = v8 (dela, 21/10 14:17) com ajustes pontuais (pedido da Juliana):

- slide 2: layout igual aos demais ("D. Title Only": titulo, bandeira, logo Pirelli e numero de pagina); remove o rodape
  (fonte + miniatura) que vinha do layout MASTER; conteudo nao muda;
- slides 3 (Brasil) e 7 (IBI): bridge refeito com o MENS FITTED v4 (mesmas formulas da 'Tabela Variantes' do Base Book);
  tabelas e comentarios nao mudam;
- slide 4 (Riscos e Oportunidades): comentario abaixo de cada item e 5 quadros dos dois lados;
- slide 8 (Rentabilidade por cliente IBI): numeros da Rentabilidade v3 (14:17) + acabamento visual;
- demais slides intocados.

Uso: python ajustar_ppt_mp27_v9.py <v8.pptx> <mens_v4.xlsx> <rentab_cliente_ibi_v3.xlsx> <saida.pptx>
"""
import copy
import io
import sys

import openpyxl
from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Emu

DECK, MENS, RENT, OUT = sys.argv[1:5]
MINUS = '−'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS = {'a': A, 'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}


def num(x, d=1):
    return f"{abs(x):,.{d}f}".replace(',', 'X').replace('.', ',').replace('X', '.')


def f1(x, d=1):
    return (MINUS if round(x, d) < 0 else '') + num(x, d)


def sg(x, d=1):
    return num(0, d) if round(abs(x), d) == 0 else (MINUS if x < 0 else '+') + num(x, d)


def pc(x, d=1):
    return num(x * 100, d) + '%'


prs = Presentation(DECK)
S = lambda n: prs.slides[n - 1]


def bynames(slide):
    out = {}
    for sh in slide.shapes:
        out[sh.name] = sh
    return out


def settext(sh, new, run=0):
    """troca o texto mantendo a formatacao do run indicado (demais runs do 1o paragrafo ficam vazios)."""
    p = sh.text_frame.paragraphs[0]
    rs = p.runs
    rs[run].text = new
    for i, r in enumerate(rs):
        if i != run and run == 0:
            r.text = ''


def new_id(slide):
    return max(int(x) for x in slide.shapes._spTree.xpath('.//p:cNvPr/@id')) + 1


def clone(slide, sh, name=None):
    el = copy.deepcopy(sh._element)
    c = el.xpath('.//p:cNvPr')[0]
    c.set('id', str(new_id(slide)))
    for ext in c.findall('{%s}extLst' % A):
        c.remove(ext)
    if name:
        c.set('name', name)
    slide.shapes._spTree.append(el)
    return slide.shapes[-1]


# =========================================================================================== MENS -> bridges
wb = openpyxl.load_workbook(MENS, data_only=True)
JAGUAR, ANTT = -0.55, -1.497       # mesmas premissas fixas da 'Tabela Variantes' do Base Book


def bridge(sheet):
    ws = wb[sheet]
    T = lambda r: float(ws.cell(r, 20).value or 0)
    start, end = float(ws.cell(43, 19).value) / 1000, float(ws.cell(43, 17).value) / 1000
    d = [('Volume', T(8) / 1000),
         ('Price/ Mix', (T(13) + T(16)) / 1000 - JAGUAR),
         ('Encerr. Jaguar', JAGUAR),
         ('Labour', (T(19) + T(32)) / 1000),
         ('Cost Var.', sum(T(r) for r in range(20, 24)) / 1000 - ANTT),
         ('ANTT', ANTT),
         ('Cost Fix.', sum(T(r) for r in range(33, 42)) / 1000)]
    assert abs(start + sum(v for _, v in d) - end) < 0.003, (sheet, start, d, end)
    return {'start': start, 'deltas': d, 'end': end}


X0, X1, YT, YB, Y_BR = 300000, 8450000, 4950000, 6420000, 4480000
LBL_D, LBL_T = 150000, 210000


def upd_bridge(n, spec):
    s = S(n)
    grp = [x for x in s.shapes if x.shape_type == 6 and x.name.startswith('Bridge MP27')][0]
    g = {x.name: x for x in grp.shapes}
    con = [x for x in grp.shapes if x.name == 'Conector']
    cols = [('Fcst R9', 'tot', spec['start'])] + [(k, 'd', v) for k, v in spec['deltas']] + [('MP’27', 'tot', spec['end'])]
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
    sc = (YB - (200000 if vmin < 0 else 0) - YT) / (vmax - vmin)
    y = lambda v: int(YT + (vmax - v) * sc)
    for i, (lab, kind, v) in enumerate(cols):
        lo, hi, lvl = bars[i]
        b = g[f'Barra {lab}']
        b.top, b.height = Emu(y(hi)), Emu(max(y(lo) - y(hi), 12700))
        t = g[f'Valor {lab}']
        if kind == 'tot':
            t.top = Emu((y(0) + 25000) if v < 0 else (y(hi) - 25000 - LBL_T))
            txt = f1(v)
        else:
            t.top = Emu(y(hi) - 25000 - LBL_D)
            txt = sg(v)
        p = t.text_frame.paragraphs[0]
        p.runs[0].text = txt
        for r in p.runs[1:]:
            r.text = ''
        if i < len(cols) - 1:
            con[i].top = Emu(y(lvl))
    lt = lambda i: (y(0) - 20000) if cols[i][2] < 0 else (y(bars[i][1]) - 25000 - LBL_T - 30000)
    g['Colchete esq'].height = Emu(lt(0) - Y_BR)
    g['Colchete dir'].height = Emu(lt(len(cols) - 1) - Y_BR)
    if s.has_notes_slide:       # dados do grafico guardados nas notas
        notes = s.notes_slide.notes_text_frame
        vals = {lab.replace('’', "'"): v for lab, _, v in cols}
        out = []
        for ln in notes.text.split('\n'):
            k = ln.split('\t')[0]
            out.append(f"{k}\t{vals[k]:.3f}" if '\t' in ln and k in vals else ln)
        notes.text = '\n'.join(out)
    print('slide', n, 'bridge', [(k, round(v, 3)) for k, v in spec['deltas']], round(spec['start'], 3), '->', round(spec['end'], 3))


# =========================================================================================== slide 2: layout
s2 = S(2)
assert 'Bridge consolidado' in s2.shapes[0].text_frame.text
ref = S(3)
lay = ref.slide_layout
assert lay.name == 'D. Title Only'
for rel in s2.part.rels.values():
    if rel.reltype == RT.SLIDE_LAYOUT:
        rel._target = lay.part
b2 = bynames(s2)
title_txt = b2['Text 0'].text_frame.text
tree2 = s2.shapes._spTree
tree2.remove(b2['Text 0']._element)
tree2.remove(b2['Slide Number Placeholder 0']._element)
ttl = copy.deepcopy([x for x in ref.shapes if x.name == 'Title 2'][0]._element)
ttl.xpath('.//p:cNvPr')[0].set('id', str(new_id(s2)))
for ext in ttl.xpath('.//p:cNvPr/a:extLst'):
    ttl.xpath('.//p:cNvPr')[0].remove(ext)
paras = ttl.xpath('.//a:p')
runs = paras[0].findall('{%s}r' % A)
runs[0].find('{%s}t' % A).text = title_txt
for r in runs[1:]:
    paras[0].remove(r)
tree2.insert(2, ttl)
flag = [x for x in ref.shapes if x.name == 'Picture 163'][0]
pic = s2.shapes.add_picture(io.BytesIO(flag.image.blob), flag.left, flag.top, flag.width, flag.height)
pic.name = 'Picture 163'
print('slide 2: layout', s2.slide_layout.name)

# =========================================================================================== slides 3 e 7: bridges
assert 'Economics Brasil' in S(3).shapes.title.text_frame.text and 'IBI' in S(7).shapes.title.text_frame.text
upd_bridge(3, bridge('TOTAL'))
upd_bridge(7, bridge('IBI'))

# =========================================================================================== slide 4: R&O
s4 = S(4)
assert 'Risks' in s4.shapes.title.text_frame.text
b4 = bynames(s4)
GREY = '595959'


def comentario(box, label, txt, left, width, nome):
    """label sobe um pouco e o comentario entra abaixo, dentro do quadro."""
    label.top = Emu(box.top + 95000)
    c = clone(s4, label, nome)
    c.left, c.top, c.width, c.height = Emu(left), Emu(box.top + 335000), Emu(width), Emu(230000)
    bp = c.text_frame._txBody.find('{%s}bodyPr' % A)
    bp.set('anchor', 't')
    for ch in list(bp):
        bp.remove(ch)
    p = c.text_frame.paragraphs[0]
    r = p.runs[0]
    r.text = txt
    for x in p.runs[1:]:
        x.text = ''
    rp = r._r.find('{%s}rPr' % A)
    rp.set('sz', '950')
    rp.set('b', '0')
    rp.set('i', '1')
    rp.find('.//{%s}srgbClr' % A).set('val', GREY)
    end = p._p.find('{%s}endParaRPr' % A)
    if end is not None:
        p._p.remove(end)


# 5 quadros tambem em Oportunidades (copias do item 3, mesmo passo vertical)
for k, top in ((4, 3920000), (5, 4640000)):
    box = clone(s4, b4['Oportunidades item 3'], f'Oportunidades item {k}')
    box.top = Emu(top)
    val = clone(s4, b4['Oportunidades item 3 valor'], f'Oportunidades item {k} valor')
    val.top = Emu(top + 196889)
    settext(val, 'xx')
b4['Oportunidades item 1 valor'].top = Emu(2480000 + 196889)    # valor do item 2 estava ~37 mil EMU acima do alinhamento
b4['Oportunidades item 1 valor'].name = 'Oportunidades item 2 valor'
b4 = bynames(s4)

TXT = {
    'Riscos item 1': "Reajuste Fiat de +3,6% (R$ 13,64/pç) em negociação; plano considera +R$ 1,4 mi vs R09",
    'Riscos item 2': "Mudança de escala exige reforço de efetivo e eleva o custo de MO; impacto a quantificar",
    'Oportunidades item 1': "Ganho potencial de headcount entre unidades, ainda não refletido no plano; a quantificar",
    'Oportunidades item 2': "Potencial de redução do custo de energia elétrica em Ibirité, fora do plano; a quantificar",
    'Oportunidades item 3': "Reajuste de preço junto à Tukan ainda não refletido no plano; a quantificar",
}
for key, txt in TXT.items():
    lab = b4[key + ' texto']
    comentario(b4[key], lab, txt, lab.left, 5050000 if key.startswith('Riscos') else 5050000, key + ' comentario')

# =========================================================================================== slide 8: rentabilidade
s8 = S(8)
assert 'Rentabilidade por cliente' in s8.shapes.title.text_frame.text
rw = openpyxl.load_workbook(RENT, data_only=True)['Rentabilidade por Cliente']
assert rw['A24'].value == 'EBIT' and rw['A25'].value == 'ROS %'
COL = {'Fiat': 2, 'Iveco': 3, 'CNH': 4, 'Total': 5}
R = lambda row, c: float(rw.cell(row, COL[c]).value or 0)
ns = {c: R(6, c) for c in COL}
ebit = {c: R(24, c) for c in COL}
ros = {c: R(25, c) for c in COL}
cm = {c: R(14, c) for c in COL}
fx = {c: R(16, c) for c in COL}
pcs = {c: R(5, c) for c in COL}
rev_sh = {c: ns[c] / ns['Total'] for c in ('Fiat', 'Iveco', 'CNH')}
ebit_sh = {c: ebit[c] / ebit['Total'] for c in ('Fiat', 'Iveco', 'CNH')}
META = float(rw['B29'].value)
assert abs(ebit['Total'] - 6597.93) < 1 and abs(sum(ebit[c] for c in ('Fiat', 'Iveco', 'CNH')) - ebit['Total']) < 1
b8 = bynames(s8)
pct0 = lambda x: f"{round(x * 100)}%"

settext(b8['Mensagem principal'],
        f"A Fiat traz {pct0(rev_sh['Fiat'])} da receita e {pct0(ebit_sh['Fiat'])} do EBIT de Ibirité, com ROS de {pc(ros['Fiat'])} "
        f"(meta {pc(META)}); Iveco e CNH têm margem alta em volume pequeno")
GREEN = '3A9A5B'
for c in ('Fiat', 'Iveco', 'CNH'):
    settext(b8[f'ROS {c}'], pc(ros[c]), run=1)
    sub = b8[f'Sub {c}']
    sub.left, sub.width = b8[f'Card {c}'].left, b8[f'Card {c}'].width
    sub.text_frame._txBody.find('{%s}bodyPr' % A).set('wrap', 'square')
    p = sub.text_frame.paragraphs[0]
    p.runs[0].text = f"{pcs[c]:,.0f}".replace(',', '.') + f"k pçs  |  EBIT R$ {num(ebit[c] / 1000)} mi"
    for r in p.runs[1:]:
        r.text = ''
    # chip de comparacao logo abaixo
    chip = clone(s8, b8['Sub Iveco'], f'Chip {c}')
    chip.left, chip.top, chip.width, chip.height = b8[f'Card {c}'].left, Emu(2800000), b8[f'Card {c}'].width, Emu(170000)
    cp = chip.text_frame.paragraphs[0]
    if c == 'Fiat':
        t = f"Meta ROS ≥ {pc(META)}: atende ({sg((ros['Fiat'] - META) * 100)} p.p.)"
    else:
        t = f"{sg((ros[c] - ros['Total']) * 100)} p.p. vs ROS total IBI ({pc(ros['Total'])})"
    cp.runs[0].text = t
    for r in cp.runs[1:]:
        r.text = ''
    rp = cp.runs[0]._r.find('{%s}rPr' % A)
    rp.set('sz', '950')
    rp.set('b', '1')
    rp.find('.//{%s}srgbClr' % A).set('val', GREEN)
    b8[f'Valor {c} % da receita IBI'].text_frame.paragraphs[0].runs[0].text = pct0(rev_sh[c])
    b8[f'Valor {c} % do EBIT IBI'].text_frame.paragraphs[0].runs[0].text = pct0(ebit_sh[c])
    b8[f'Barra {c} % da receita IBI'].width = Emu(int(2240000 * rev_sh[c]))
    b8[f'Barra {c} % do EBIT IBI'].width = Emu(int(2240000 * ebit_sh[c]))

# ---- tabela: mais linhas (CM %, EBIT/pc), zebra refeita, EBIT e ROS em destaque
gf = b8['Table 95']
tbl = gf.table._tbl
trs = tbl.findall('{%s}tr' % A)
by_lab = {tr.findall('{%s}tc' % A)[0].xpath('string(.//a:t)'): tr for tr in trs}
tr_cm, tr_ebit = by_lab['CM'], by_lab['EBIT']


def newrow(src, after):
    n = copy.deepcopy(src)
    after.addnext(n)
    return n


tr_cmp = newrow(tr_cm, tr_cm)
tr_epc = newrow(tr_ebit, tr_ebit)
trs = tbl.findall('{%s}tr' % A)


def set_row(tr, vals):
    for tc, v in zip(tr.findall('{%s}tc' % A), vals):
        ts = tc.xpath('.//a:t')
        ts[0].text = v
        for t in ts[1:]:
            t.text = ''


def fill(tc, rgb):
    pr = tc.find('{%s}tcPr' % A)
    for tag in ('solidFill', 'noFill'):
        for e in pr.findall('{%s}%s' % (A, tag)):
            pr.remove(e)
    sf = etree.SubElement(pr, '{%s}solidFill' % A)
    etree.SubElement(sf, '{%s}srgbClr' % A).set('val', rgb)


def style(tc, bold, color):
    for rp in tc.xpath('.//a:rPr'):
        rp.set('b', '1' if bold else '0')
        sf = rp.find('{%s}solidFill' % A)
        if sf is None:
            sf = etree.SubElement(rp, '{%s}solidFill' % A)
            etree.SubElement(sf, '{%s}srgbClr' % A)
        sf.find('{%s}srgbClr' % A).set('val', color)


cs = ('Fiat', 'Iveco', 'CNH', 'Total')
paren = lambda x: f"({num(x / 1000)})" if x < 0 else num(x / 1000)
rows = {
    'Peças (mil)': [f"{pcs[c]:,.0f}".replace(',', '.') for c in cs],
    'Net Sales': [num(ns[c] / 1000) for c in cs],
    'NS/pç (R$)': [num(ns[c] / pcs[c], 2) for c in cs],
    'CM': [num(cm[c] / 1000) for c in cs],
    'Fixed Cost': [paren(fx[c]) for c in cs],
    'EBIT': [num(ebit[c] / 1000) for c in cs],
    'ROS %': [pc(ros[c]) for c in cs],
}
for tr in trs:
    lab = tr.findall('{%s}tc' % A)[0].xpath('string(.//a:t)')
    if lab in rows:
        set_row(tr, [lab] + rows[lab])
set_row(tr_cmp, ['CM % da receita'] + [pc(cm[c] / ns[c]) for c in cs])
set_row(tr_epc, ['EBIT/pç (R$)'] + [num(ebit[c] / pcs[c] / 1000 * 1000, 2) for c in cs])
trs = tbl.findall('{%s}tr' % A)
RH = 215000
for i, tr in enumerate(trs):
    tr.set('h', str(RH))
    if i == 0:
        continue
    lab = tr.findall('{%s}tc' % A)[0].xpath('string(.//a:t)')
    key = lab in ('EBIT', 'ROS %')
    for tc in tr.findall('{%s}tc' % A):
        fill(tc, 'DCE6F1' if key else ('FFFFFF' if i % 2 else 'F2F2F2'))
        style(tc, key, '1F4E79' if key else '000000')
    if lab in ('CM % da receita', 'EBIT/pç (R$)'):
        for tc in tr.findall('{%s}tc' % A):
            for rp in tc.xpath('.//a:rPr'):
                rp.set('i', '1')
gf.height = Emu(RH * len(trs))

# ---- comentarios (mesma estrutura de paragrafos, texto novo)
tx = b8['Comentarios'].text_frame
P = tx.paragraphs


def ptxt(i, t):
    rs = P[i].runs
    rs[0].text = t
    for r in rs[1:]:
        r.text = ''


ptxt(1, f"{pct0(rev_sh['Fiat'])} da receita e {pct0(ebit_sh['Fiat'])} do EBIT (R$ {num(ebit['Fiat'] / 1000)} mi)")
ptxt(2, f"ROS {pc(ros['Fiat'])}, acima da meta de {pc(META)} ({sg((ros['Fiat'] - META) * 100)} p.p.)")
ptxt(4, f"ROS {pc(ros['Iveco'])} e EBIT de R$ {num(ebit['Iveco'] / 1000)} mi")
ptxt(7, f"ROS {pc(ros['CNH'])} e EBIT de R$ {num(ebit['CNH'] / 1000)} mi, com custo dedicado")
rs_iv = (ns['Iveco'] + ns['CNH']) / ns['Total']
eb_iv = (ebit['Iveco'] + ebit['CNH']) / ebit['Total']
ptxt(10, f"Iveco e CNH: {pct0(rs_iv)} da receita e {pct0(eb_iv)} do EBIT; margem de contribuição de "
         f"{round(cm['CNH'] / ns['CNH'] * 100)}% a {round(cm['Iveco'] / ns['Iveco'] * 100)}% da receita, contra {round(cm['Fiat'] / ns['Fiat'] * 100)}% da Fiat")
import copy as _cp
last = P[10]._p
for t in (f"Fiat: cada 1 p.p. de ROS vale R$ {num(ns['Fiat'] / 1000 * 0.01)} mi de EBIT",
          f"Total IBI: EBIT de R$ {num(ebit['Total'] / 1000)} mi e ROS de {pc(ros['Total'])}"):
    np_ = _cp.deepcopy(P[10]._p)
    last.addnext(np_)
    last = np_
    rs_ = np_.findall('{%s}r' % A)
    rs_[0].find('{%s}t' % A).text = t
    for r_ in rs_[1:]:
        np_.remove(r_)
print('slide 8:', {c: (pc(ros[c]), pct0(rev_sh[c]), pct0(ebit_sh[c])) for c in ('Fiat', 'Iveco', 'CNH')})

prs.save(OUT)
print('salvo', OUT, len(prs.slides), 'slides; ocultos:', [i + 1 for i, s in enumerate(prs.slides) if s._element.get('show') == '0'])
