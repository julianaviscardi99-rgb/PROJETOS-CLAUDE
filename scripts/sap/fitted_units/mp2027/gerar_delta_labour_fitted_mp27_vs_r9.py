"""Delta Labour Cost FITTED: MP'27 vs R9 2026 (por unidade, categoria e conta SAP).

Uso: python gerar_delta_labour_fitted_mp27_vs_r9.py <Labour FITTED MP27.xlsx> <Labour FITTED R09.xlsx> <saida.xlsx>
Somente leitura nas fontes; grava apenas o arquivo de saida.
O R9 nao tem coluna de unidade: a unidade vem do centro de custo no MP'27.
Centros de custo que so existem no R9 ficam como ENC (sem MP'27).
"""
import sys
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

UNIDADES = ["IBI", "SJP", "GOI", "RES", "GER", "ENC"]
NOME_UN = {"IBI": "IBI - Ibirite", "SJP": "SJP - Sao Jose dos Pinhais", "GOI": "GOI - Goiana",
           "RES": "RES - Resende", "GER": "GER - Gerencia", "ENC": "ENC - CCs so no R9 (sem MP'27)"}
MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

COMENTARIOS = {
    ("IBI", "4201000P"): "Rampa de 2026: R9 sai de 308k (jan) p/ ~450k (dez); MP'27 ja parte de ~455k. Efeito de base, nao so aumento.",
    ("IBI", "4201160P"): "Encargos s/ HE: no R9 so existem de set em diante (~55k/mes); MP'27 tem os 12 meses.",
    ("IBI", "4245100P"): "R9 ~100k/mes ate ago e 134k desde set; MP'27 136k e 151k desde abr (reajuste do plano?).",
    ("IBI", "4201130P"): "PLR horista: R9 61-76k/mes (jan com estorno de -79k); MP'27 83k fixo.",
    ("IBI", "4201110P"): "13o horista acompanha a rampa de salarios.",
    ("IBI", "4201150P"): "Unica conta relevante de IBI que cai: encargos s/ salarios menores no MP'27 que no R9.",
    ("IBI", "4201170P"): "Encargos s/ ferias: R9 tem estornos/valores baixos no inicio do ano.",
    ("SJP", "4201130P"): "PLR horista: R9 ~74k/mes (jan -32k); MP'27 88k fixo (+19%).",
    ("SJP", "4201000P"): "Jan/26 baixo no R9 (SJP total 238k) e +reajuste no MP'27.",
    ("RES", "4201000P"): "RES comecou em ago/26 no R9 (60k) e so tem run-rate cheio (~158k/mes) em set; MP'27 tem o ano cheio. Efeito de calendario.",
    ("RES", "4200000P"): "Idem: mensalistas de RES so aparecem no R9 a partir de ago.",
    ("GOI", "4201100P"): "HE: R9 43-59k/mes ate jul e ~19k de ago em diante; MP'27 segue o nivel recente (~22k).",
    ("GOI", "4201000P"): "Reajuste sobre o nivel do R9.",
    ("ENC", "4201130P"): "PLR provisionado so em jan/26 (223k) em CCs 8231/8269/8292 sem MP'27: nao recorrente.",
}


def ler(path, ini, col_cc, col_ac, col_cat=1, cc2u=None):
    ws = openpyxl.load_workbook(path, data_only=True)["Labour Cost"]
    out = defaultdict(lambda: [0.0] * 12)
    desc, un = {}, {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[col_cat]:
            continue
        cc = str(r[col_cc])
        u = r[3] if cc2u is None else cc2u.get(cc, "ENC")
        un[cc] = u
        desc[r[col_ac]] = r[col_ac + 1]
        for i, v in enumerate(r[ini:ini + 12]):
            out[(u, r[col_cat], r[col_ac])][i] += v or 0
    return out, desc, un


def main(f27, f9, saida):
    m27, d27, cc2u = ler(f27, 9, 2, 4)
    m9, d9, _ = ler(f9, 8, 2, 3, cc2u=cc2u)
    desc = {**d9, **d27}
    chaves = set(m27) | set(m9)
    z = [0.0] * 12
    linhas = sorted(chaves, key=lambda k: -abs(sum(m27.get(k, z)) - sum(m9.get(k, z))))

    wb = openpyxl.Workbook()
    azul = Font(color="0000FF"); neg = Font(bold=True)
    hdr = PatternFill("solid", fgColor="1F3864")
    num = '#,##0;[Red](#,##0);"-"'; pct = '0.0%;[Red]-0.0%'

    def cab(ws, row, cols):
        for i, c in enumerate(cols, 1):
            x = ws.cell(row, i, c)
            x.font = Font(bold=True, color="FFFFFF"); x.fill = hdr
            x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # --- base: unidade x categoria x conta
    wb_ = wb.active; wb_.title = "Base unidade x conta"
    wb_["A1"] = "Labour Cost FITTED - MP'27 vs R9 2026 (R$) - base por unidade, categoria e conta"
    wb_["A1"].font = Font(bold=True, size=14)
    wb_["A2"] = "R9 = 2026.01-12 de 'Labour Cost FITTED - R09 2026'; MP'27 = 2027.01-12 de 'Labour Cost FITTED - MP27'. Azul = valor das fontes; preto = formula."
    wb_["A2"].font = Font(italic=True, color="666666")
    cab(wb_, 4, ["Unidade", "Categoria", "Conta SAP", "Descricao", "R9 2026", "MP'27", "Delta R$", "Delta %", "% do delta total", "Comentario"])
    r0 = 5; rt = r0 + len(linhas); vistos = set()
    for n, k in enumerate(linhas):
        r = r0 + n; u, cat, ac = k
        for c, v in enumerate([u, cat, ac, (desc.get(ac) or "").strip()], 1):
            wb_.cell(r, c, v)
        wb_.cell(r, 5, round(sum(m9.get(k, z)), 2)).font = azul
        wb_.cell(r, 6, round(sum(m27.get(k, z)), 2)).font = azul
        wb_.cell(r, 7, f"=F{r}-E{r}")
        wb_.cell(r, 8, f'=IF(E{r}=0,"n/a",G{r}/ABS(E{r}))')
        wb_.cell(r, 9, f"=G{r}/$G${rt}")
        if (u, ac) in COMENTARIOS and (u, ac) not in vistos:
            wb_.cell(r, 10, COMENTARIOS[(u, ac)]); vistos.add((u, ac))
        for col in "EFG": wb_[f"{col}{r}"].number_format = num
        wb_[f"H{r}"].number_format = wb_[f"I{r}"].number_format = pct
        wb_[f"H{r}"].alignment = Alignment(horizontal="right")
    wb_.cell(rt, 1, "TOTAL")
    for col in "EFG": wb_[f"{col}{rt}"] = f"=SUM({col}{r0}:{col}{rt-1})"; wb_[f"{col}{rt}"].number_format = num
    wb_[f"H{rt}"] = f"=G{rt}/ABS(E{rt})"; wb_[f"I{rt}"] = f"=SUM(I{r0}:I{rt-1})"
    wb_[f"H{rt}"].number_format = wb_[f"I{rt}"].number_format = pct
    for c in range(1, 11):
        wb_.cell(rt, c).font = neg; wb_.cell(rt, c).border = Border(top=Side(style="medium"))
    for col, w in zip("ABCDEFGHIJ", [9, 12, 11, 34, 14, 14, 14, 10, 14, 90]):
        wb_.column_dimensions[col].width = w
    wb_.freeze_panes = "A5"
    B = f"'Base unidade x conta'!"
    rng = lambda c: f"{B}${c}${r0}:${c}${rt-1}"

    # --- mensal por unidade
    wm = wb.create_sheet("Mensal por unidade")
    wm["A1"] = "Total mensal por unidade (R$)"; wm["A1"].font = Font(bold=True, size=14)
    cab(wm, 3, ["Unidade", "Cenario"] + MESES + ["Total"])
    pos = {}; r = 4
    for u in UNIDADES:
        for cen, src in (("R9 2026", m9), ("MP'27", m27)):
            wm.cell(r, 1, u); wm.cell(r, 2, cen)
            for i in range(12):
                v = sum(vals[i] for k, vals in src.items() if k[0] == u)
                c = wm.cell(r, 3 + i, round(v, 2)); c.font = azul; c.number_format = num
            wm.cell(r, 15, f"=SUM(C{r}:N{r})").number_format = num
            pos[(u, cen)] = r; r += 1
        wm.cell(r, 1, u); wm.cell(r, 2, "Delta")
        for i in range(13):
            L = chr(ord("C") + i)
            wm[f"{L}{r}"] = f"={L}{r-1}-{L}{r-2}"; wm[f"{L}{r}"].number_format = num; wm[f"{L}{r}"].font = neg
        r += 2
    wm.column_dimensions["A"].width = 9; wm.column_dimensions["B"].width = 10
    for i in range(13): wm.column_dimensions[chr(ord("C") + i)].width = 11
    wm.freeze_panes = "C4"

    # --- resumo
    wr = wb.create_sheet("Resumo", 0)
    wr["A1"] = "Labour Cost FITTED - MP'27 vs R9 2026 (R$)"; wr["A1"].font = Font(bold=True, size=14)
    wr["A2"] = "Por unidade. 'Run-rate R9' = media set-dez/26 x 12: separa efeito de base/rampa de 2026 do aumento real."
    wr["A2"].font = Font(italic=True, color="666666")
    cab(wr, 4, ["Unidade", "R9 2026", "MP'27", "Delta R$", "Delta %", "% do delta total", "Run-rate R9 (set-dez x12)", "MP'27 vs run-rate R$", "MP'27 vs run-rate %"])
    for j, u in enumerate(UNIDADES):
        r = 5 + j
        wr.cell(r, 1, NOME_UN[u])
        wr[f"B{r}"] = f'=SUMIF({rng("A")},"{u}",{rng("E")})'
        wr[f"C{r}"] = f'=SUMIF({rng("A")},"{u}",{rng("F")})'
        wr[f"D{r}"] = f"=C{r}-B{r}"
        wr[f"E{r}"] = f'=IF(B{r}=0,"n/a",D{r}/ABS(B{r}))'
        wr[f"F{r}"] = f"=D{r}/$D${5+len(UNIDADES)}"
        wr[f"G{r}"] = f"=AVERAGE('Mensal por unidade'!K{pos[(u,'R9 2026')]}:N{pos[(u,'R9 2026')]})*12"
        wr[f"H{r}"] = f"=C{r}-G{r}"
        wr[f"I{r}"] = f'=IF(G{r}=0,"n/a",H{r}/ABS(G{r}))'
    rt2 = 5 + len(UNIDADES)
    wr.cell(rt2, 1, "TOTAL")
    for col in "BCDGH": wr[f"{col}{rt2}"] = f"=SUM({col}5:{col}{rt2-1})"
    wr[f"E{rt2}"] = f"=D{rt2}/ABS(B{rt2})"; wr[f"F{rt2}"] = f"=SUM(F5:F{rt2-1})"; wr[f"I{rt2}"] = f"=H{rt2}/ABS(G{rt2})"
    for r in range(5, rt2 + 1):
        for col in "BCDGH": wr[f"{col}{r}"].number_format = num
        for col in "EFI": wr[f"{col}{r}"].number_format = pct; wr[f"{col}{r}"].alignment = Alignment(horizontal="right")
    for c in range(1, 10):
        wr.cell(rt2, c).font = neg; wr.cell(rt2, c).border = Border(top=Side(style="medium"))

    rc = rt2 + 3
    wr.cell(rc, 1, "Por categoria").font = Font(bold=True, size=12)
    cab(wr, rc + 1, ["Categoria", "R9 2026", "MP'27", "Delta R$", "Delta %", "% do delta total"])
    cats = sorted({k[1] for k in chaves})
    for j, cat in enumerate(cats):
        r = rc + 2 + j
        wr.cell(r, 1, cat)
        wr[f"B{r}"] = f'=SUMIF({rng("B")},A{r},{rng("E")})'
        wr[f"C{r}"] = f'=SUMIF({rng("B")},A{r},{rng("F")})'
        wr[f"D{r}"] = f"=C{r}-B{r}"; wr[f"E{r}"] = f"=D{r}/ABS(B{r})"; wr[f"F{r}"] = f"=D{r}/$D${rt2}"
        for col in "BCD": wr[f"{col}{r}"].number_format = num
        wr[f"E{r}"].number_format = wr[f"F{r}"].number_format = pct

    rn = rc + 3 + len(cats)
    wr.cell(rn, 1, "Notas").font = Font(bold=True)
    notas = [
        "Comentarios sao leitura dos padroes mensais (hipoteses), nao premissas documentadas: confirmar com o RH.",
        "RES: operacao comecou em ago/26 no R9; o delta e quase todo efeito de calendario (run-rate set-dez x12 ~ MP'27).",
        "ENC: centros de custo 8231, 8269 e 8292 so existem no R9 (PLR provisionado em jan/26, 263k) e nao tem unidade no MP'27.",
        "O R9 nao traz a unidade: ela foi obtida pelo centro de custo no arquivo MP'27.",
    ]
    for i, t in enumerate(notas, 1): wr.cell(rn + i, 1, t)
    wr.column_dimensions["A"].width = 34
    for col in "BCDEFGHI": wr.column_dimensions[col].width = 16
    wr.row_dimensions[4].height = 32

    # --- por conta
    wc = wb.create_sheet("Delta por conta", 1)
    wc["A1"] = "Delta por conta SAP (todas as unidades)"; wc["A1"].font = Font(bold=True, size=14)
    cab(wc, 3, ["Conta SAP", "Descricao", "R9 2026", "MP'27", "Delta R$", "Delta %", "% do delta total"])
    contas = sorted({k[2] for k in chaves}, key=lambda a: -abs(sum(sum(m27[k]) for k in m27 if k[2] == a) - sum(sum(m9[k]) for k in m9 if k[2] == a)))
    for j, ac in enumerate(contas):
        r = 4 + j
        wc.cell(r, 1, ac); wc.cell(r, 2, (desc.get(ac) or "").strip())
        wc[f"C{r}"] = f'=SUMIF({rng("C")},A{r},{rng("E")})'
        wc[f"D{r}"] = f'=SUMIF({rng("C")},A{r},{rng("F")})'
        wc[f"E{r}"] = f"=D{r}-C{r}"; wc[f"F{r}"] = f'=IF(C{r}=0,"n/a",E{r}/ABS(C{r}))'
        wc[f"G{r}"] = f"=E{r}/$E${4+len(contas)}"
        for col in "CDE": wc[f"{col}{r}"].number_format = num
        wc[f"F{r}"].number_format = wc[f"G{r}"].number_format = pct; wc[f"F{r}"].alignment = Alignment(horizontal="right")
    rt3 = 4 + len(contas)
    wc.cell(rt3, 1, "TOTAL")
    for col in "CDE": wc[f"{col}{rt3}"] = f"=SUM({col}4:{col}{rt3-1})"; wc[f"{col}{rt3}"].number_format = num
    wc[f"F{rt3}"] = f"=E{rt3}/ABS(C{rt3})"; wc[f"G{rt3}"] = f"=SUM(G4:G{rt3-1})"
    wc[f"F{rt3}"].number_format = wc[f"G{rt3}"].number_format = pct
    for c in range(1, 8):
        wc.cell(rt3, c).font = neg; wc.cell(rt3, c).border = Border(top=Side(style="medium"))
    for col, w in zip("ABCDEFG", [12, 36, 15, 15, 15, 10, 14]): wc.column_dimensions[col].width = w
    wc.freeze_panes = "A4"
    wb.save(saida)


if __name__ == "__main__":
    main(*sys.argv[1:4])
