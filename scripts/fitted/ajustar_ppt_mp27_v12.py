"""PPT MP27 Fitted v12 = v11 (dela) atualizada com MENS FITTED v4 e Rentabilidade IBI v3 de 22/10 19:16 (precos recuperados).
So muda: barras/rotulos dos bridges (slides 2, 3, 6, 7, 9, 10), KPIs soltos do slide 2 e % de variacao dos bridges.
NAO mexe em tabelas, estrutura nem comentarios (paineis 'Comentarios', tabela do slide 2, tabela do slide 8)
Uso extra: segundo argumento extra = Rentab IBI v3.
Uso: python ajustar_ppt_mp27_v12.py <v11.pptx> <mens_v4.xlsx> <saida.pptx> <rentab_ibi_v3.xlsx>
"""
import re
import sys

import openpyxl
from pptx import Presentation
from pptx.util import Emu

DECK, MENS, OUT, RENT = sys.argv[1:5]
MINUS = '−'
JAGUAR = -0.55
wb = openpyxl.load_workbook(MENS, data_only=True)
prs = Presentation(DECK)
S = lambda n: prs.slides[n - 1]


def f1(x):
    return (MINUS if round(x, 1) < 0 else '') + f"{abs(x):.1f}".replace('.', ',')


def sg(x):
    return '0,0' if round(abs(x), 1) == 0 else (MINUS if x < 0 else '+') + f"{abs(x):.1f}".replace('.', ',')


def settext(sh, txt):
    if sh.text_frame.text.replace('+', '') == txt.replace('+', ''):
        return                      # mesmo valor: preserva o texto original
    p = sh.text_frame.paragraphs[0]
    p.runs[0].text = txt
    for r in p.runs[1:]:
        r.text = ''


def repl(sh, rx, new, count=1):
    """troca por regex dentro dos runs, preservando formatacao"""
    for p in sh.text_frame.paragraphs:
        for r in p.runs:
            if rx.search(r.text):
                r.text = rx.sub(new, r.text, count=count)
                return
    raise AssertionError((sh.name, sh.text_frame.text))


def bridge(sheet, jaguar):
    ws = wb[sheet]
    T = lambda r: float(ws.cell(r, 20).value or 0)
    start, end = float(ws.cell(43, 19).value) / 1000, float(ws.cell(43, 17).value) / 1000
    pm = (T(13) + T(16)) / 1000
    d = [('Volume', T(8) / 1000)]
    if jaguar:
        d += [('Price/ Mix', pm - JAGUAR), ('Encerr. Jaguar', JAGUAR)]
    else:
        d += [('Price/ Mix', pm)]
    d += [('Labour', (T(19) + T(32)) / 1000), ('Cost Var.', sum(T(r) for r in range(20, 24)) / 1000),
          ('Cost Fix.', sum(T(r) for r in range(33, 42)) / 1000)]
    assert abs(start + sum(v for _, v in d) - end) < 0.003, (sheet, start, d, end)
    return {'start': start, 'deltas': d, 'end': end}


YT, YB, Y_BR, LBL_D, LBL_T = 4950000, 6420000, 4480000, 150000, 210000


def upd_bridge(n, spec):
    s = S(n)
    grp = [x for x in s.shapes if x.shape_type == 6 and x.name.startswith('Bridge MP27')][0]
    g = {x.name: x for x in grp.shapes}
    con = [x for x in grp.shapes if x.name == 'Conector']
    cols = [('Fcst R9', 'tot', spec['start'])] + [(k, 'd', v) for k, v in spec['deltas']] + [('MP’27', 'tot', spec['end'])]
    assert all(f'Barra {c[0]}' in g for c in cols), (n, [c[0] for c in cols])
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
        settext(t, txt)
        if i < len(cols) - 1:
            con[i].top = Emu(y(lvl))
    lt = lambda i: (y(0) - 20000) if cols[i][2] < 0 else (y(bars[i][1]) - 25000 - LBL_T - 30000)
    g['Colchete esq'].height = Emu(lt(0) - Y_BR)
    g['Colchete dir'].height = Emu(lt(len(cols) - 1) - Y_BR)
    a, b_ = spec['start'], spec['end']
    pct = (b_ - a) / a * 100 if a > 0.5 else None
    settext(g['Variacao'], ('Δ ' + sg(b_ - a)) if pct is None else '0,0%' if round(abs(pct), 1) == 0
            else f"{'+' if pct > 0 else MINUS}{abs(pct):.1f}%".replace('.', ','))
    if s.has_notes_slide:       # dados do grafico guardados nas notas
        notes = s.notes_slide.notes_text_frame
        vals = {lab.replace('’', "'"): v for lab, _, v in cols}
        out = []
        for ln in notes.text.split('\n'):
            k = ln.split('\t')[0]
            out.append(f"{k}\t{vals[k]:.3f}" if '\t' in ln and k in vals else ln)
        notes.text = '\n'.join(out)
    print('slide', n, [(k, round(v, 3)) for k, v in spec['deltas']], round(spec['start'], 3), '->', round(spec['end'], 3))


upd_bridge(3, bridge('TOTAL', True))
upd_bridge(6, bridge('SJP', False))
upd_bridge(7, bridge('IBI', True))
upd_bridge(9, bridge('GO', False))
upd_bridge(10, bridge('RES', False))

# ------------------------------------------------------------------ slide 2: bridge por unidade + KPIs soltos
s2 = S(2)
ws = wb['TOTAL']
E = lambda sh, c: float(wb[sh].cell(43, c).value) / 1000
r9, mp = E('TOTAL', 19), E('TOTAL', 17)
units = [('IBI', 'Ibirité'), ('SJP', 'SJP'), ('GO', 'Goiana'), ('RES', 'Resende')]
dl = [E(u, 17) - E(u, 19) for u, _ in units]
dl.append(mp - r9 - sum(dl))                      # Sorocaba (encerrada) = residuo do Brasil
assert abs(dl[4] - (-0.188)) < 0.01, dl
nsm, nsr = float(ws.cell(10, 17).value) / 1000, float(ws.cell(10, 19).value) / 1000
ros, ros9 = mp / nsm * 100, r9 / nsr * 100
Y0, K = 5204968, 2535277 / r9
y = lambda v: int(Y0 - v * K)
byid = {x.shape_id: x for x in s2.shapes}
# shape_id: barras IBI 26, SJP 29, GOI 32, RES 35, SOR 87, final 39; conectores 27 30 33 36 90 (+38 final)
bar_ids, lab_ids, con_ids = [26, 29, 32, 35, 87], [28, 31, 34, 37, 88], [27, 30, 33, 36, 90]
cum = r9
for i, d in enumerate(dl):
    hi = max(cum, cum + d)
    bar, lab = byid[bar_ids[i]], byid[lab_ids[i]]
    bar.top, bar.height = Emu(y(hi)), Emu(max(int(abs(d) * K), 12700))
    lab.top = Emu(y(hi) - 256032)
    settext(lab, sg(d))
    cum += d
    byid[con_ids[i]].top = Emu(y(cum))
byid[38].top = Emu(y(mp))           # conector final
fin = byid[39]
fin.top, fin.height = Emu(y(mp)), Emu(Y0 - y(mp))
byid[40].top = Emu(y(mp) - 256032)
settext(byid[40], f1(mp))
settext(byid[25], f1(r9))
byid[25].top = Emu(y(r9) - 256032)
tx = lambda i, t: settext(byid[i], t)
dmi = (f"{MINUS if mp < r9 else '+'}" if round(abs(mp - r9), 1) else '') + f"R$ {abs(mp - r9):.1f} mi".replace('.', ',')
vol, vol9 = float(ws.cell(8, 17).value), float(ws.cell(8, 19).value)
tx(8, f"{vol:,.0f} mil peças".replace(',', '.'))
tx(9, f"{'+' if vol >= vol9 else MINUS}{abs(vol / vol9 - 1) * 100:.1f}% vs R09".replace('.', ','))
tx(12, f"R$ {f1(nsm)} mi")
tx(13, f"{'+' if nsm >= nsr else MINUS}{abs(nsm / nsr - 1) * 100:.1f}% vs R09".replace('.', ','))
tx(5, f"EBIT de R$ {f1(mp)} mi (ROS {f1(ros)}%): {dmi} vs R09")
tx(16, f"R$ {f1(mp)} mi")
tx(17, dmi + ' ')
tx(20, f"{f1(ros)}%")
tx(21, f"{MINUS if ros < ros9 else '+'}{abs(ros - ros9):.1f}".replace('.', ',') + " p.p. vs R09 Brasil")
tx(48, f"Δ EBIT: {dmi}")
tx(49, f"Δ EBIT: {dmi}")
conc = (E('IBI', 17) + E('SJP', 17)) / mp * 100
repl(byid[81], re.compile(r'\d+%'), f'{conc:.0f}%')

# ------------------------------------------------------------------ slide 8: cards/barras soltos (tabela NAO e tocada)
s8 = S(8)
sh8 = {x.name: x for x in s8.shapes}
R = openpyxl.load_workbook(RENT, data_only=True)['Rentabilidade por Cliente']
row = {R.cell(r, 1).value: r for r in range(1, 40)}
col = {'Fiat': 2, 'Iveco': 3, 'CNH': 4}
tot_ns, tot_eb = R.cell(row['NET SALES'], 5).value, R.cell(row['EBIT'], 5).value
pct = {}
for c, k in col.items():
    pcs, ns, eb = (R.cell(row[x], k).value for x in ('Pieces (mil)', 'NET SALES', 'EBIT'))
    pr, pe = ns / tot_ns, eb / tot_eb
    pct[c] = (pr, pe, eb / ns)
    repl(sh8[f'ROS {c}'], re.compile(r'[\d,]+%'), f"{eb / ns * 100:.1f}%".replace('.', ','))
    repl(sh8[f'Sub {c}'], re.compile(r'^[\d.]+k'), f"{pcs:,.0f}k".replace(',', '.'))
    repl(sh8[f'Sub {c}'], re.compile(r'R\$ [\d,]+ mi'), f"R$ {eb / 1000:.1f} mi".replace('.', ','))
    for lab, v in (('receita', pr), ('EBIT', pe)):
        key = f'% da receita IBI' if lab == 'receita' else '% do EBIT IBI'
        settext(sh8[f'Valor {c} {key}'], f"{v * 100:.0f}%")
        sh8[f'Barra {c} {key}'].width = Emu(int(sh8[f'Trilho {c} {key}'].width * v))
fr, fe, fros = pct['Fiat']
repl(sh8['Mensagem principal'], re.compile(r'\d+(?=% do EBIT)'), f"{fe * 100:.0f}")
repl(sh8['Mensagem principal'], re.compile(r'\d+(?=% da)'), f"{fr * 100:.0f}")
repl(sh8['Mensagem principal'], re.compile(r'(?<=ROS de )[\d,]+(?=%)'), f"{fros * 100:.1f}".replace('.', ','))
print('slide 8', {c: tuple(round(x, 3) for x in v) for c, v in pct.items()})
print('slide 2', [round(d, 3) for d in dl], round(mp, 3), f'ROS {ros:.2f} vs {ros9:.2f}', f'conc {conc:.1f}%')
prs.save(OUT)
print('salvo', OUT, len(prs.slides))
