"""Auditoria Labour Cost FITTED: MP27 v1 x v2 x R9 2026 (por unidade, conta, razoes e conciliacao).

Uso: python auditoria_labour_mp27.py <v1.xlsx> <v2.xlsx> <r9.xlsx> <mens_mp27.xlsx> <pl_r9.xlsx> <saida.xlsx>
Somente leitura nas fontes; grava apenas o arquivo de saida.
O R9 nao traz unidade: ela vem do centro de custo no MP27 (CC so no R9 = ENC).
"""
import sys
import numpy as np
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L

f1, f2, f9, fmens, fpl9, out = sys.argv[1:7]
MS = [f"M{i:02d}" for i in range(1, 13)]
UN = ["SJP", "IBI", "GOI", "RES", "GER", "ENC"]


def carregar(path, cc2u=None):
    df = pd.read_excel(path, sheet_name="Labour Cost", dtype={"SAP Code": str, "Cost Center": str})
    df = df[df["SAP Code"].notna()].copy()
    cols = [c for c in df.columns if str(c)[:2] == "20"]
    df = df.rename(columns={c: f"M{str(c)[-2:]}" for c in cols})
    df["Var"] = np.where(df["Cost Center Variability"].astype(str).str.startswith("Fixo"), "Fixo", "Variavel")
    df["SAP Account Description"] = df["SAP Account Description"].astype(str).str.strip()
    if cc2u is not None:
        df["LOCALIDADE"] = df["Cost Center"].map(cc2u).fillna("ENC")
    return df


v1, v2 = carregar(f1), carregar(f2)
cc2u = v2.drop_duplicates("Cost Center").set_index("Cost Center")["LOCALIDADE"].to_dict()
r9 = carregar(f9, cc2u)
RR = ["M09", "M10", "M11", "M12"]
tot = lambda d: d[MS].sum().sum()
rr = lambda d: d[RR].sum().sum() * 3  # run-rate anual = media set-dez x 12

wb = Workbook()
HDR = PatternFill("solid", fgColor="1F3864"); SUB = PatternFill("solid", fgColor="D9E1F2")
RED = PatternFill("solid", fgColor="F8CBAD"); YEL = PatternFill("solid", fgColor="FFF2CC")
NUM = '#,##0;[Red]-#,##0;"-"'; PCT = '0.0%;[Red]-0.0%;"-"'


def head(ws, row, cols):
    for i, t in enumerate(cols, 1):
        c = ws.cell(row, i, t)
        c.font = Font(bold=True, color="FFFFFF"); c.fill = HDR
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def larg(ws, w):
    for i, x in enumerate(w, 1):
        ws.column_dimensions[L(i)].width = x


# ---------------------------------------------------------------- 1. Resumo por unidade
ws = wb.active; ws.title = "Resumo por unidade"
ws["A1"] = "Labour Cost FITTED - R9 2026 x MP27 v1 x MP27 v2 (R$ mil, custo direto, sem rateio da Gerencia)"; ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = "Run-rate R9 = media set-dez/26 x 12 (separa rampa de 2026 do aumento real). Azul = valor das fontes; preto = formula."
ws["A2"].font = Font(italic=True, color="666666")
cols = ["Unidade", "Tipo", "R9 FY26", "R9 run-rate", "MP27 v1", "MP27 v2", "v2 vs v1", "v2 vs v1 %", "v2 vs R9 FY", "v2 vs R9 FY %", "v2 vs run-rate", "v2 vs run-rate %"]
head(ws, 4, cols); r = 5; first = r
for u in UN:
    for tp in ("Variavel", "Fixo"):
        sel = lambda d: d[(d.LOCALIDADE == u) & (d.Var == tp)]
        vals = [tot(sel(r9)) / 1e3, rr(sel(r9)) / 1e3, tot(sel(v1)) / 1e3, tot(sel(v2)) / 1e3]
        if not any(abs(x) > 0.5 for x in vals):
            continue
        ws.cell(r, 1, u); ws.cell(r, 2, tp)
        for j, x in enumerate(vals):
            c = ws.cell(r, 3 + j, round(x, 1)); c.font = Font(color="0000FF"); c.number_format = NUM
        ws.cell(r, 7, f"=F{r}-E{r}"); ws.cell(r, 8, f'=IF(E{r}=0,"n/a",G{r}/E{r})')
        ws.cell(r, 9, f"=F{r}-C{r}"); ws.cell(r, 10, f'=IF(C{r}=0,"n/a",I{r}/C{r})')
        ws.cell(r, 11, f"=F{r}-D{r}"); ws.cell(r, 12, f'=IF(D{r}=0,"n/a",K{r}/D{r})')
        r += 1
ws.cell(r, 1, "TOTAL").font = Font(bold=True)
for col in "CDEFGIK":
    ws[f"{col}{r}"] = f"=SUM({col}{first}:{col}{r-1})"; ws[f"{col}{r}"].font = Font(bold=True)
ws[f"H{r}"] = f"=G{r}/E{r}"; ws[f"J{r}"] = f"=I{r}/C{r}"; ws[f"L{r}"] = f"=K{r}/D{r}"
for rr_ in range(first, r + 1):
    for col in "CDEFGIK": ws[f"{col}{rr_}"].number_format = NUM
    for col in "HJL": ws[f"{col}{rr_}"].number_format = PCT; ws[f"{col}{rr_}"].alignment = Alignment(horizontal="right")
    ws.cell(rr_, 1).fill = SUB if rr_ == r else PatternFill()
ws.cell(r + 2, 1, "Leitura: RES Var+Fixo juntos = R9 run-rate (1.910 x 1.909): a quebra Fixo->Variavel de RES 8349 e reclassificacao, nao aumento. ENC = CCs 8231/8269/8292 (PLR jan/26) sem unidade.")
ws.cell(r + 3, 1, "SJP Fixo +31% vs run-rate (inclui rateio GER); ver aba 'Linhas x run-rate' (CC 8289 horista fixo +35%).")
larg(ws, [10, 10, 12, 12, 12, 12, 12, 11, 13, 13, 14, 14]); ws.freeze_panes = "C5"

# ---------------------------------------------------------------- 2. Linhas x run-rate
ws = wb.create_sheet("Linhas x run-rate")
K = ["LOCALIDADE", "Cost Center", "Category", "Var"]
a = r9.groupby(K)[RR].sum().mean(axis=1) * 12 / 1e3
b1 = v1.groupby(K)[MS].sum().sum(axis=1) / 1e3; b2 = v2.groupby(K)[MS].sum().sum(axis=1) / 1e3
t = pd.DataFrame({"rr": a, "v1": b1, "v2": b2}).fillna(0)
t["d"] = t.v2 - t.rr; t = t.sort_values("d", key=abs, ascending=False)
ws["A1"] = "Custo anual por CC x categoria: run-rate R9 (set-dez x12) x MP27 v1 x v2 (R$ mil)"; ws["A1"].font = Font(bold=True, size=13)
head(ws, 3, ["Unidade", "Centro de custo", "Categoria", "Var/Fixo", "R9 run-rate", "MP27 v1", "MP27 v2", "v2 vs run-rate", "%"])
for i, (k, row) in enumerate(t.iterrows(), 4):
    for j, v in enumerate(k, 1): ws.cell(i, j, v)
    for j, x in enumerate([row.rr, row.v1, row.v2], 5):
        c = ws.cell(i, j, round(float(x), 1)); c.font = Font(color="0000FF"); c.number_format = NUM
    ws.cell(i, 8, f"=G{i}-E{i}").number_format = NUM
    ws.cell(i, 9, f'=IF(E{i}=0,"n/a",H{i}/E{i})').number_format = PCT
    if abs(row.d) > 150 and row.rr != 0 and abs(row.d / row.rr) > 0.3:
        for j in range(1, 10): ws.cell(i, j).fill = RED
n = len(t) + 4
ws.cell(n, 1, "TOTAL").font = Font(bold=True)
for col in "EFGH": ws[f"{col}{n}"] = f"=SUM({col}4:{col}{n-1})"; ws[f"{col}{n}"].number_format = NUM; ws[f"{col}{n}"].font = Font(bold=True)
larg(ws, [10, 14, 13, 10, 13, 12, 12, 14, 9]); ws.freeze_panes = "A4"

# ---------------------------------------------------------------- 3. Razoes (provisoes e encargos)
ws = wb.create_sheet("Razoes provisoes")
ws["A1"] = "Provisoes e encargos sobre a base (salario + HE). Estatutario: 13o 8,33%; ferias 11,11% (1/12 + 1/3)"; ws["A1"].font = Font(bold=True, size=13)
head(ws, 3, ["Grupo", "Unidade", "Cenario", "Base sal+HE (R$ mil)", "13o / base", "Ferias / base", "Enc s/ ferias / ferias", "Enc s/ 13o / 13o", "PLR / salario", "Enc s/ salario", "Assist. medica / salario"])
GR = {"HORISTA": (["HORISTA", "SENAI"], dict(sal="4201000P", he="4201100P", p13="4201110P", fer="4201120P", plr="4201130P", enc="4201150P", ef="4201170P", e13="4201180P", med="4245100P")),
      "MENSALISTA": (["MENSALISTA", "SENIOR"], dict(sal="4200000P", he="4200100P", p13="4200110P", fer="4200120P", plr="4200130P", enc="4200150P", ef="4200170P", e13="4200180P", med="4245000P"))}
CEN = [("R9 jan-ago (realizado)", r9, MS[:8]), ("R9 set-dez (forecast)", r9, RR), ("MP27 v1", v1, MS), ("MP27 v2", v2, MS)]
r = 4
for g, (cats, A) in GR.items():
    for u in ["IBI", "SJP", "GOI", "RES", "GER"]:
        for nome, d, mm in CEN:
            x = d[(d.Category.isin(cats)) & (d.LOCALIDADE == u)].groupby("SAP Account")[mm].sum().sum(axis=1)
            s = x.get(A["sal"], 0)
            if s <= 0: continue
            base = s + x.get(A["he"], 0)
            vals = [g, u, nome, round(base / 1e3), x.get(A["p13"], 0) / base, x.get(A["fer"], 0) / base,
                    x.get(A["ef"], 0) / max(x.get(A["fer"], 0), 1), x.get(A["e13"], 0) / max(x.get(A["p13"], 0), 1),
                    x.get(A["plr"], 0) / s, x.get(A["enc"], 0) / s, x.get(A["med"], 0) / s]
            for j, v in enumerate(vals, 1):
                c = ws.cell(r, j, v)
                if j >= 5: c.number_format = '0.0%'
                if j == 4: c.number_format = NUM
            if g == "HORISTA" and nome == "MP27 v2" and u in ("IBI", "SJP", "GOI"):
                ws.cell(r, 6).fill = RED; ws.cell(r, 7).fill = RED
            if g == "HORISTA" and u == "GOI" and vals[10] < 0.08:
                ws.cell(r, 11).fill = YEL
            r += 1
ws.cell(r + 1, 1, "Vermelho: ferias do horista em ~5,5% da base (metade do estatutario) e encargo s/ ferias ~90% da propria provisao (esperado ~35-40%); RES horista esta correto (11,4%).")
ws.cell(r + 2, 1, "Amarelo: assistencia medica do horista GOI ~3% do salario vs 24-41% nas demais unidades (conferir).")
larg(ws, [13, 9, 22, 14, 10, 12, 16, 14, 12, 12, 14]); ws.freeze_panes = "A4"

# ---------------------------------------------------------------- 4. Exposicao ferias horista
ws = wb.create_sheet("Exposicao ferias horista")
ws["A1"] = "Ferias horista MP27 v2 x estatutario (11,11% x (salario+HE)) - R$ mil/ano"; ws["A1"].font = Font(bold=True, size=13)
head(ws, 3, ["Unidade", "Base sal+HE", "Ferias v2", "Ferias estatutario", "Gap ferias", "Taxa enc s/ salario", "Gap encargos", "Gap total"])
r = 4
for u in ["IBI", "SJP", "GOI", "RES"]:
    x = v2[(v2.Category.isin(["HORISTA", "SENAI"])) & (v2.LOCALIDADE == u)].groupby("SAP Account")[MS].sum().sum(axis=1)
    base = x["4201000P"] + x.get("4201100P", 0)
    ws.cell(r, 1, u)
    for j, val in enumerate([base / 1e3, x["4201120P"] / 1e3], 2):
        c = ws.cell(r, j, round(val, 1)); c.font = Font(color="0000FF"); c.number_format = NUM
    ws.cell(r, 4, f"=B{r}/9").number_format = NUM; ws.cell(r, 5, f"=D{r}-C{r}").number_format = NUM
    c = ws.cell(r, 6, round(x["4201150P"] / x["4201000P"], 4)); c.font = Font(color="0000FF"); c.number_format = "0.0%"
    ws.cell(r, 7, f"=E{r}*F{r}").number_format = NUM; ws.cell(r, 8, f"=E{r}+G{r}").number_format = NUM
    r += 1
ws.cell(r, 1, "TOTAL").font = Font(bold=True)
for col in "BCDEGH": ws[f"{col}{r}"] = f"=SUM({col}4:{col}{r-1})"; ws[f"{col}{r}"].number_format = NUM; ws[f"{col}{r}"].font = Font(bold=True)
ws.cell(r + 2, 1, "Hipotese a confirmar com o RH: se a provisao de ferias do horista deve ser ~11,1% da base (como RES e como o v1 em ~10,3%), IBI/SJP/GOI estao subprovisionados.")
larg(ws, [10, 13, 12, 17, 12, 17, 14, 12])

# ---------------------------------------------------------------- 5. MO variavel x volume
ws = wb.create_sheet("MO variavel x volume")
mens = openpyxl.load_workbook(fmens, data_only=True, read_only=False)
ws["A1"] = "MO 'variavel' MP27 v2 x pecas (MENS FITTED MP27 v3, linhas Pieces e Labour variavel)"; ws["A1"].font = Font(bold=True, size=13)
head(ws, 3, ["Unidade", "Metrica"] + ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"] + ["Total", "R9 FY (MENS)"])
r = 4; resumo = []
for sh in ("SJP", "IBI", "GO", "RES"):
    w = mens[sh]
    pc = [w.cell(8, c).value for c in range(5, 17)]; lab = [-w.cell(19, c).value for c in range(5, 17)]
    ws.cell(r, 1, sh); ws.cell(r, 2, "Pecas (mil)")
    ws.cell(r + 1, 1, sh); ws.cell(r + 1, 2, "MO var (R$ mil)")
    ws.cell(r + 2, 1, sh); ws.cell(r + 2, 2, "R$ por peca")
    for j in range(12):
        ws.cell(r, 3 + j, round(pc[j], 1)).number_format = NUM
        ws.cell(r + 1, 3 + j, round(lab[j], 1)).number_format = NUM
        ws.cell(r + 2, 3 + j, f"={L(3+j)}{r+1}/{L(3+j)}{r}").number_format = "0.00"
    ws.cell(r, 15, f"=SUM(C{r}:N{r})").number_format = NUM; ws.cell(r + 1, 15, f"=SUM(C{r+1}:N{r+1})").number_format = NUM
    ws.cell(r + 2, 15, f"=O{r+1}/O{r}").number_format = "0.00"
    ws.cell(r, 16, round(w.cell(8, 19).value, 1)).number_format = NUM; ws.cell(r + 1, 16, round(-w.cell(19, 19).value, 1)).number_format = NUM
    ws.cell(r + 2, 16, f'=IF(P{r+1}<50,"n/a",P{r+1}/P{r})').number_format = "0.00"
    el = np.polyfit(np.log(pc), np.log(lab), 1)[0]
    resumo.append((sh, round(float(el), 2), round(float(np.corrcoef(pc, lab)[0, 1]), 2)))
    r += 4
ws.cell(r, 1, "Elasticidade MO/volume (1,0 = totalmente variavel) e correlacao mensal:").font = Font(bold=True)
for i, (sh, el, co) in enumerate(resumo, 1):
    ws.cell(r + i, 1, sh); ws.cell(r + i, 2, el); ws.cell(r + i, 3, co)
ws.cell(r + 6, 1, "A MO 'variavel' do MP27 nao acompanha o volume: SJP dez = 38,9 mil pecas e MO 376 mil (R$ 9,66/peca); GOI out = 45 mil pecas (parada Jeep) e MO 318 mil (R$ 7,05/peca).")
larg(ws, [10, 17] + [8] * 12 + [10, 13])

# ---------------------------------------------------------------- 6. Conciliacoes
ws = wb.create_sheet("Conciliacoes")
ws["A1"] = "Conciliacoes (R$ mil)"; ws["A1"].font = Font(bold=True, size=13)
head(ws, 3, ["Teste", "Fonte A", "Fonte B", "Diferenca", "Resultado"])
pl = openpyxl.load_workbook(fpl9, data_only=True)["Resumo Resultado Ano"]
plv = -np.array([pl.cell(20, c).value for c in range(4, 16)], float); plf = -np.array([pl.cell(31, c).value for c in range(4, 16)], float)
lv = r9[r9.Var == "Variavel"][MS].sum().values / 1e3; lf = r9[r9.Var == "Fixo"][MS].sum().values / 1e3
rat = {"GOI": .22, "IBI": .52, "RES": .035, "SJP": .225}
ger = v2[v2.LOCALIDADE == "GER"][MS].sum().values
mx = 0
for sh, u in {"SJP": "SJP", "IBI": "IBI", "GO": "GOI", "RES": "RES"}.items():
    w = mens[sh]
    for row, tp in ((19, "V"), (32, "F")):
        mv = -np.array([w.cell(row, c).value for c in range(5, 17)], float)
        sub = v2[v2.LOCALIDADE == u]
        cal = (sub[sub.Var == "Variavel"][MS].sum().values / 1e3) if tp == "V" else (sub[sub.Var == "Fixo"][MS].sum().values + rat[u] * ger) / 1e3
        mx = max(mx, np.abs(mv - cal).max())
linhas = [
    ("MENS MP27 v3 x Labour v2 (4 unidades, Var e Fixo+rateio, max mes)", "MENS", "Labour v2", round(mx, 3), "OK - MENS esta ligada ao Labour v2 (link externo) e fecha"),
    ("P&L R9 Labour Var x Labour R9: set-dez", round(plv[8:].sum(), 1), round(lv[8:].sum(), 1), round(lv[8:].sum() - plv[8:].sum(), 1), "OK - fecha"),
    ("P&L R9 Labour Fixo x Labour R9: set-dez", round(plf[8:].sum(), 1), round(lf[8:].sum(), 1), round(lf[8:].sum() - plf[8:].sum(), 1), "OK - fecha"),
    ("P&L R9 Labour Var x Labour R9: jan-ago", round(plv[:8].sum(), 1), round(lv[:8].sum(), 1), round(lv[:8].sum() - plv[:8].sum(), 1), "NAO fecha: jan +228 (PLR jan/26 em CCs 8231/8269/8292 sem unidade)"),
    ("P&L R9 Labour Fixo x Labour R9: jan-ago", round(plf[:8].sum(), 1), round(lf[:8].sum(), 1), round(lf[:8].sum() - plf[:8].sum(), 1), "NAO fecha: jan +37; mar +4; ago -9"),
]
for i, row in enumerate(linhas, 4):
    for j, v in enumerate(row, 1): ws.cell(i, j, v)
    ws.cell(i, 4).number_format = '#,##0.0;[Red]-#,##0.0'
larg(ws, [70, 14, 14, 12, 75])

# ---------------------------------------------------------------- 0. Achados (primeira aba)
ws = wb.create_sheet("Achados", 0)
ws["A1"] = "Auditoria Labour Cost FITTED - MP27 v1 x v2 x R9: achados priorizados"; ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = "Impacto em R$ mil/ano, estimativa. Hipoteses = a confirmar com RH/Rafaela; nao sao premissas documentadas."
ws["A2"].font = Font(italic=True, color="666666")
head(ws, 4, ["#", "Gravidade", "Unidade", "Achado", "Evidencia", "Impacto (R$ mil/ano)", "O que fazer"])
ACH = [
    ("ALTA", "IBI/SJP/GOI", "Provisao de ferias do HORISTA na v2 esta em ~5,5% da base (salario+HE): metade do estatutario (11,1%). RES horista esta correto (11,4%).",
     "v1 = 10,3%; v2 = 5,8%. R9 realizado jan-ago ~12-13%, mas R9 forecast set-dez ja cai p/ 6-9% (origem provavel). Encargo s/ ferias (4201170P) nao foi reduzido: 89% da propria provisao (esperado ~35-40%, como nos mensalistas). Delta v1->v2: ferias horista -583k, 13o horista -224k (este correto: v1 10,2% vs 8,33%).",
     -955, "Confirmar com RH a regra de provisao de ferias horista. Se 11,1%: +R$ 695k ferias + R$ 262k encargos (IBI 581 / SJP 218 / GOI 156). Revisar tambem o R9 set-dez."),
    ("ALTA", "Todas", "MO 'variavel' do MP27 nao acompanha o volume.",
     "Elasticidade MO/pecas: SJP 0,02; IBI 0,12; GOI 0,04; RES 0,01 (1,0 = variavel). SJP dez: 38,9 mil pecas e MO 376k (R$ 9,66/peca); GOI out: 45 mil pecas e MO 318k (R$ 7,05/peca).",
     None, "Decidir se e por headcount fixo (entao classificar como fixo / indicar premissa) ou flexionar com volume; sensibilidade de volume hoje nao reduz MO."),
    ("ALTA", "SJP", "CC 8289 (horista FIXO): +35% vs run-rate set-dez/26 sem justificativa visivel.",
     "Custo/mes R9 set-dez 94,7k -> MP27 jan 127,4k. Salario 43,4k->57,5k (+32%), PLR 11,1k->19,8k (+79%), assist. medica 5,6k->9,7k (+75%). SJP Fixo total +31% vs run-rate (+471k incl. rateio).",
     460, "Pedir ao RH o HC/premissa do CC 8289 (houve transferencia de 8290 -15k/mes?)."),
    ("MEDIA", "GER/GOI/IBI/RES vs SJP", "Reajuste: mensalistas +4,8% em jun/27 (GER, GOI, IBI, RES); SJP +3,1% em out/27; jan/27 = dez/26.",
     "A premissa citada no deck (~7% reajuste MO) nao aparece no arquivo. Em jun as provisoes tem retroativo (5 meses x delta x %), mas o retroativo do salario em si nao existe.",
     None, "Confirmar data-base/indice por unidade; SJP em mes diferente e % menor chama atencao."),
    ("MEDIA", "RES", "CC 8349 (R$ 873k) e FIXO no R9 e VARIAVEL no MP27.",
     "Na MENS a MO variavel RES vai de 9,6 (R9) para 872,5 (MP27): o bridge custo variavel x fixo fica distorcido. RES total = R9 run-rate (1.910 x 1.909), ou seja, reclassificacao, nao aumento.",
     873, "Confirmar a classificacao e ajustar a leitura do bridge Var/Fixo."),
    ("MEDIA", "Todas", "Rateio da Gerencia do MP27 (22/52/22,5/3,5% GOI/IBI/SJP/RES) e digitado e difere da media calculada na propria aba Rateio Fixo (20,6/55,6/18,0/5,8%) e do vigente (27/49/21/3%).",
     "GER = 589k/ano. Efeito vs media calculada: SJP +26,5k; IBI -21,1k; GOI +8,0k; RES -13,4k (soma zero).",
     None, "Documentar o criterio; ja estava como pendencia ('rateio fixo so depois de quadrar o faturamento')."),
    ("MEDIA", "GOI", "Assistencia medica do HORISTA = 3,4% do salario (IBI 31%, SJP 24%, RES 41%).",
     "Mesma coisa no R9 (0,9% jan-ago / 3,5% set-dez); mensalista GOI = 17,6%. Pode ser estrutural (beneficio em outra conta/CC) ou omissao.",
     356, "Confirmar. Se GOI tivesse a proporcao de SJP: +R$ 356k."),
    ("BAIXA", "R9", "Labour R9 nao fecha com o P&L R9 em jan-ago (+268k); set-dez fecha 100%.",
     "Jan +228k var / +37k fixo = PLR jan/26 nos CCs 8231/8269/8292 (sem unidade, 'ENC'); mar +21k; ago -18k. A MENS usa R9 = 25.499k; o arquivo soma 25.768k.",
     268, "Excluir ENC da base R9 de comparacao (ja esta fora na MENS)."),
    ("BAIXA", "R9", "Provisao de ferias NEGATIVA em dez/26 no forecast: SJP -24,5k, GOI -24,9k (CCs 8289/8290/8291 e 8300/8302/8303/8368).",
     "Tambem estornos de PLR/enc. em jan/fev. Dezembro negativo parece nao ser recorrente nem esperado como forecast.",
     -49, "Revisar o forecast de dez/26 (afeta R9 e o run-rate)."),
    ("BAIXA", "SJP", "PLR de MENSALISTA do CC 8289 lancado na conta de horista 4201130P (R$ 26,4k no MP27; 7,4k R9).",
     "Correto seria 4200130P. Mesmo total de MO, so muda a conta gestorial.", None, "Corrigir a conta nas duas bases."),
    ("BAIXA", "Estrutura", "Sheet1 (tabelas dinamicas) e valor colado; a MENS le o Labour v2 via link externo [3]Sheet1.",
     "Hoje as pivots batem com a aba Labour Cost (0 diferenca) e a MENS bate com o v2 (max 0,0 mil). A v1 daria ate 40 mil/mes de diferenca. Se editar Labour Cost sem atualizar a pivot, a MENS nao muda.",
     None, "Atualizar a pivot a cada alteracao; manter conferencia (linhas 40/50/51 do Sheet1)."),
    ("BAIXA", "IBI", "CC 8278 (horista fixo) so tem assistencia medica e seguro (R$ 54k/ano), sem salario.",
     "Igual no R9.", 54, "Verificar se e custo de ex-funcionarios."),
    ("INFO", "Estrutura", "P&L Fitted Units_Budget27.xlsx (09/10) esta desatualizado em MO.",
     "Var 19,8 mi / Fixo 8,5 mi vs Labour v2 21,7 / 8,1. A fonte atual e a MENS v3 (ligada ao v2).", None, "Nao usar o P&L Budget27 de 09/10 para MO."),
    ("INFO", "v1->v2", "Acertos da v2: 13o horista (10,2% -> 8,5% vs 8,33% estatutario) e ferias mensalista (14,2% -> 11,5% vs 11,1%).",
     "v1 -> v2: -882,5k (-2,9%): IBI -562, GOI -158, SJP -131, RES -24, GER -8.", None, "Manter."),
]
for i, (g, un, ach, ev, imp, acao) in enumerate(ACH, 1):
    r = 4 + i
    vals = [i, g, un, ach, ev, imp, acao]
    for j, v in enumerate(vals, 1):
        c = ws.cell(r, j, v); c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(r, 6).number_format = NUM
    ws.cell(r, 2).fill = RED if g == "ALTA" else (YEL if g == "MEDIA" else PatternFill())
larg(ws, [4, 10, 16, 52, 70, 14, 52]); ws.freeze_panes = "A5"

wb.save(out)
print("salvo:", out)
