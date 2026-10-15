"""Delta Labour Cost CP: MP'27 vs R9 2026, por conta SAP.

Uso: python gerar_delta_labour_mp27_vs_r9.py <Labour MP27.xlsx> <Labour R09.xlsx> <saida.xlsx>
Somente leitura nas fontes; grava apenas o arquivo de saida.
"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

COMENTARIOS = {
    ("GERENTE", "4200130P"): "R9 provisiona ~5,7k/mes ate ago (estorno -2,5k em abr) e sobe p/ 8,8-9,3k em set-dez; MP'27 ja parte de 9,3k em todos os meses.",
    ("GERENTE", "4200000P"): "Jan/27 +7,3% e jun/27 +4,8% (dissidio?). R9 tem jul em 11k e set em 39k (evento pontual).",
    ("GERENTE", "4200170P"): "Distorcao do R9: estorno de -12,1k em jul reduz o ano a 1,7k. Sem ele o R9 seria ~16k.",
    ("GERENTE", "4200120P"): "Provisao = 13,9% dos ordenados no MP'27 vs 11,1% no R9 (1/12 + 1/3). Conferir premissa/formula.",
    ("HORISTA", "4201000P"): "R9 comeca em 850 (jan) e 4,4k (fev): entrada tardia. MP'27 tem o ano cheio.",
    ("GERENTE", "4200150P"): "Acompanha ordenados (reajuste); R9 tem pico em jul (12,4k).",
    ("HORISTA", "4201100P"): "R9 quase zero de jan-ago e 1,8k desde set; MP'27 fixa 1,4k/mes.",
    ("HORISTA", "4201160P"): "Acompanha horas extras (R9 so a partir de set).",
    ("HORISTA", "4201130P"): "Unica linha que cai: R9 tem picos em jan e abr.",
}


def ler(path):
    ws = openpyxl.load_workbook(path, data_only=True)["Labour Cost"]
    out = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[1]:
            continue
        vals = [v for v in r[8:20] if isinstance(v, (int, float))]
        out[(r[1], r[3])] = (r[4], sum(vals))
    return out


def main(f27, f9, saida):
    a, b = ler(f27), ler(f9)
    linhas = []
    for k, (desc, v27) in a.items():
        linhas.append((k, desc, b.get(k, (desc, 0))[1], v27))
    linhas.sort(key=lambda x: -abs(x[3] - x[2]))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Delta por conta"
    azul = Font(color="0000FF")
    hdr = PatternFill("solid", fgColor="1F3864")
    fino = Side(style="thin", color="BBBBBB")

    ws["A1"] = "Labour Cost CP - MP'27 vs R9 2026 (R$)"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "R9 = 2026.01-12 do arquivo 'Labour Cost CP - R09 2026'; MP'27 = 2027.01-12 de 'Labour Cost CP - MP27'. Azul = valor das fontes; preto = formula."
    ws["A2"].font = Font(italic=True, color="666666")

    cab = ["Categoria", "Conta SAP", "Descricao", "R9 2026", "MP'27", "Delta R$", "Delta %", "% do delta total", "Comentario"]
    for i, c in enumerate(cab, 1):
        x = ws.cell(4, i, c)
        x.font = Font(bold=True, color="FFFFFF")
        x.fill = hdr
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    r0 = 5
    rt = r0 + len(linhas)
    for n, (k, desc, v9, v27) in enumerate(linhas):
        r = r0 + n
        ws.cell(r, 1, k[0]); ws.cell(r, 2, k[1]); ws.cell(r, 3, desc.strip())
        ws.cell(r, 4, round(v9, 2)).font = azul
        ws.cell(r, 5, round(v27, 2)).font = azul
        ws.cell(r, 6, f"=E{r}-D{r}")
        ws.cell(r, 7, f'=IF(D{r}=0,"n/a",F{r}/ABS(D{r}))')
        ws.cell(r, 8, f"=F{r}/$F${rt}")
        ws.cell(r, 9, COMENTARIOS.get(k, ""))
    ws.cell(rt, 1, "TOTAL").font = Font(bold=True)
    for col in "DEF":
        ws[f"{col}{rt}"] = f"=SUM({col}{r0}:{col}{rt-1})"
    ws[f"G{rt}"] = f"=F{rt}/ABS(D{rt})"
    ws[f"H{rt}"] = f"=SUM(H{r0}:H{rt-1})"
    for c in range(1, 10):
        ws.cell(rt, c).font = Font(bold=True)
        ws.cell(rt, c).border = Border(top=Side(style="medium"))
    for r in range(r0, rt + 1):
        for col in "DEF":
            ws[f"{col}{r}"].number_format = '#,##0;[Red](#,##0);"-"'
        ws[f"G{r}"].number_format = '0.0%;[Red]-0.0%'
        ws[f"H{r}"].number_format = '0.0%;[Red]-0.0%'
        ws[f"G{r}"].alignment = Alignment(horizontal="right")
        ws[f"I{r}"].alignment = Alignment(wrap_text=True, vertical="top")

    # resumo por categoria
    rs = rt + 3
    ws.cell(rs, 1, "Resumo por categoria").font = Font(bold=True, size=12)
    for i, c in enumerate(["Categoria", "", "", "R9 2026", "MP'27", "Delta R$", "Delta %", "% do delta total"], 1):
        x = ws.cell(rs + 1, i, c)
        x.font = Font(bold=True, color="FFFFFF"); x.fill = hdr
    for j, cat in enumerate(["GERENTE", "HORISTA"]):
        r = rs + 2 + j
        ws.cell(r, 1, cat)
        ws[f"D{r}"] = f'=SUMIF($A${r0}:$A${rt-1},$A{r},D${r0}:D${rt-1})'
        ws[f"E{r}"] = f'=SUMIF($A${r0}:$A${rt-1},$A{r},E${r0}:E${rt-1})'
        ws[f"F{r}"] = f"=E{r}-D{r}"
        ws[f"G{r}"] = f"=F{r}/ABS(D{r})"
        ws[f"H{r}"] = f"=F{r}/$F${rt}"
        for col in "DEF":
            ws[f"{col}{r}"].number_format = '#,##0;[Red](#,##0);"-"'
        ws[f"G{r}"].number_format = ws[f"H{r}"].number_format = '0.0%;[Red]-0.0%'

    nota = rs + 5
    ws.cell(nota, 1, "Notas").font = Font(bold=True)
    notas = [
        "Comentarios sao leitura dos padroes mensais (hipoteses), nao premissas documentadas: confirmar com o RH.",
        "Pontos a confirmar: eventos de jul e set/2026 no R9; 13,9% de provisao de ferias no MP'27; premissa de PLR; reajuste de jan/27 (7,3%).",
        "Linha de total do R9 na fonte so tem set-dez preenchido; aqui o total e a soma das contas.",
    ]
    for i, t in enumerate(notas, 1):
        ws.cell(nota + i, 1, t)

    for col, w in zip("ABCDEFGHI", [12, 12, 34, 14, 14, 14, 10, 14, 80]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"
    wb.save(saida)


if __name__ == "__main__":
    main(*sys.argv[1:4])
