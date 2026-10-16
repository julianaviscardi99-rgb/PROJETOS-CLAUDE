#!/usr/bin/env python3
"""TLM MP27 - de-para da mao de obra do Pyplan para as linhas do P&L do Resultado.

Problema (achado em 2026-10-16): o Resultado TLM puxa cada linha com SUMIFS em
Intermediaria[Desc.Resultado] (+ WH + Variabilidade). Na Base Total MP27 do Caetano,
as linhas do Labour Cost do Pyplan vieram com a descricao crua do Pyplan
("SALARIOS", "ENC S/ SALRIOS  HORISTA"...) e nao com o rotulo do P&L
("Salarios", "Enc.Salarios Hor."...) -> ~R$ 44,8 mi de MO ficavam fora do resultado.

Regra do de-para: chave = Conta Gestorial; rotulo = Desc.Resultado mais usado para
essa conta na base oficial de 2026 (Base Total_Act_Flash_Fct.xlsx). Variabilidade:
contas que em 2026 sao sempre F (ex.: 4247100 previdencia privada) ficam F.

So LEITURA na rede. Saida: data/processed/tlm/mp2027/De-para Pyplan MP27 - AAAAMMDD.xlsx
(nunca sobrescreve: _v2, _v3...). Simula a puxada do Resultado antes x depois.
"""
import collections
import datetime as dt
import re
import sys
import warnings
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

warnings.filterwarnings("ignore")

REDE_TLM = Path(r"\\SRV054-01BR.group.pirelli.com\DAC_TLM")
PASTA_MP27 = REDE_TLM / "MP27" / "1ª Versão MP27 vs R09" / "TLM_Automation"
BASE_MP27 = PASTA_MP27 / "Base Total MP27.xlsx"
# Resultado ligado ao MP27 (08/10) - usado so para ler os rotulos/abas que o SUMIFS procura
RESULTADO_REF = PASTA_MP27 / "_Versoes_anteriores" / "20261008_draft_por_indice" / "Resultado TLM - MP27 vs R09.xlsm"
BASE_2026 = REDE_TLM / "2026" / "Base Total_Act_Flash_Fct.xlsx"
SAIDA = Path(__file__).resolve().parents[2] / "data" / "processed" / "tlm" / "mp2027"

FONTE_PYPLAN = "RH via Pyplan"
MESES = ["January", "February", "March", "April", "May", "June", "July",
         "August", "September", "October", "November", "December"]


def nome_com_versao(pasta: Path, nome: str) -> Path:
    alvo = pasta / nome
    n = 2
    while alvo.exists():
        alvo = pasta / f"{Path(nome).stem}_v{n}{Path(nome).suffix}"
        n += 1
    return alvo


def ler_tabela(caminho: Path, coluna_cabecalho: str):
    ws = openpyxl.load_workbook(caminho, read_only=True, data_only=True).worksheets[0]
    it = ws.iter_rows(values_only=True)
    for r in it:
        if r and coluna_cabecalho in r:
            cab = r
            break
    ix = {c: i for i, c in enumerate(cab) if isinstance(c, str)}
    return ix, [r for r in it if r and r[ix["Conta Gestorial"]] is not None]


def chaves_do_resultado(caminho: Path):
    """(WH, Variabilidade, rotulo) que o SUMIFS de cada aba procura - em casefold (Excel ignora maiusculas)."""
    wb = openpyxl.load_workbook(caminho, read_only=True)
    chaves = set()
    for ws in wb.worksheets:
        linhas = list(ws.iter_rows(max_row=200, max_col=2, values_only=True))
        if len(linhas) < 3 or not isinstance(linhas[2][0], str):
            continue
        wh = linhas[2][0].strip()
        for a, b in linhas:
            if isinstance(a, str) and isinstance(b, str) and "Intermediaria[Desc.Resultado]" in b:
                m = re.search(r"Intermediaria\[Variabilidade\],\$A\$(\d+)", b)
                var = linhas[int(m.group(1)) - 1][0] if m else None
                chaves.add((wh.casefold(), str(var).strip().casefold(), a.strip().casefold()))
    return chaves


def main():
    print("Lendo Resultado (rotulos)...")
    chaves = chaves_do_resultado(RESULTADO_REF)
    print("Lendo base 2026...")
    ix26, base26 = ler_tabela(BASE_2026, "Conta Gestorial")
    desc26 = collections.defaultdict(collections.Counter)
    var26 = collections.defaultdict(collections.Counter)
    for r in base26:
        cg = str(r[ix26["Conta Gestorial"]])
        desc26[cg][r[ix26["Desc.Resultado"]]] += 1
        var26[cg][r[ix26["Variabilidade"]]] += 1
    print("Lendo Base Total MP27...")
    ix, base = ler_tabela(BASE_MP27, "Source.Name")

    # de-para por conta gestorial (so Pyplan)
    grupos = collections.defaultdict(lambda: {"valor": 0.0, "linhas": 0, "var": collections.Counter()})
    for r in base:
        if not str(r[ix["Source.Name"]]).startswith(FONTE_PYPLAN):
            continue
        k = (str(r[ix["Conta Gestorial"]]), r[ix["Desc.Resultado"]])
        grupos[k]["valor"] += r[ix["Total Ano"]] or 0
        grupos[k]["linhas"] += 1
        grupos[k]["var"][r[ix["Variabilidade"]]] += 1
    depara, regra_var, linhas_depara = {}, {}, []
    for (cg, d), g in sorted(grupos.items(), key=lambda x: -abs(x[1]["valor"])):
        novo = desc26[cg].most_common(1)[0][0] if desc26.get(cg) else None
        so_fixo_2026 = set(var26[cg]) == {"F"}
        tem_v = g["var"].get("V", 0) > 0
        if novo and novo != d:
            depara[cg] = novo
        if so_fixo_2026 and tem_v:
            regra_var[cg] = "F"
        alternativas = "; ".join(f"{k} ({c} lin)" for k, c in desc26[cg].most_common()[1:3]) if desc26.get(cg) else ""
        acao = ("troca descrição" if cg in depara else "já estava certo") + (" + variável→fixo" if cg in regra_var else "")
        linhas_depara.append([cg, d, novo or "(sem correspondência)", acao, g["linhas"],
                              "/".join(sorted(k for k in g["var"] if k)), round(g["valor"], 2), alternativas])

    # simulacao antes x depois
    def simula(aplicar):
        entra, fora = collections.Counter(), collections.defaultdict(collections.Counter)
        for r in base:
            fonte = str(r[ix["Source.Name"]]).split("|")[0].strip()
            cg = str(r[ix["Conta Gestorial"]])
            d = (r[ix["Desc.Resultado"]] or "").strip()
            var = r[ix["Variabilidade"]]
            if aplicar and fonte.startswith(FONTE_PYPLAN):
                d = depara.get(cg, d)
                var = regra_var.get(cg, var)
            v = r[ix["Total Ano"]] or 0
            ok = (str(r[ix["WH"]]).casefold(), str(var).casefold(), d.casefold()) in chaves
            entra[(fonte, ok)] += v
            if not ok:
                fora[fonte][(r[ix["WH"]], var, cg, d or "(vazio)")] += v
        return entra, fora

    antes, _ = simula(False)
    depois, fora_depois = simula(True)
    fontes = sorted({k[0] for k in antes})

    # ---- Excel ----
    wb = openpyxl.Workbook()
    neg, amar = Font(bold=True), PatternFill("solid", fgColor="FFF2CC")
    titulo = Font(bold=True, size=13)

    ws = wb.active
    ws.title = "Resumo"
    ws.append(["TLM MP27 - de-para da mão de obra do Pyplan"])
    ws["A1"].font = titulo
    ws.append([f"gerado em {dt.datetime.now():%d/%m/%Y %H:%M} por scripts/tlm/gerar_depara_pyplan.py - só leitura da rede"])
    ws.append([])
    ws.append(["Simulação da puxada do Resultado (SUMIFS por WH + Variabilidade + Desc.Resultado), R$ ano"])
    ws.append(["Fonte", "Total na base", "Entra no Resultado HOJE", "Fica fora HOJE",
               "Entra COM o de-para", "Fica fora COM o de-para"])
    for c in ws[5]:
        c.font = neg
    for f in fontes:
        tot = antes[(f, True)] + antes[(f, False)]
        ws.append([f, tot, antes[(f, True)], antes[(f, False)], depois[(f, True)], depois[(f, False)]])
    ult = ws.max_row
    ws.append(["TOTAL"] + [f"=SUM({col}6:{col}{ult})" for col in "BCDEF"])
    for c in ws[ws.max_row]:
        c.font = neg
    ws.append([])
    ws.append(["Leitura"]); ws[ws.max_row][0].font = neg
    for t in [
        "Chave do de-para = Conta Gestorial; o rótulo novo é o Desc.Resultado que a base oficial de 2026 usa para a mesma conta.",
        "Só as linhas do Pyplan mudam; base das áreas (Rafaela) e RH fixo (Iana) ficam como vieram.",
        "4247100 Previdência privada: em 2026 é sempre Fixo (488 de 488 linhas); no Pyplan veio parte como Variável -> proposta: Fixo.",
        "O que sobra fora são as 19 linhas da Iana sem conta gestorial (classe fora do plano 0981) - corrigir com a Iana, não é de-para.",
        "A simulação usa os rótulos do Resultado MP27 de 08/10 (_Versoes_anteriores), a única versão ligada à base MP27.",
    ]:
        ws.append([t])
    for col, w in zip("ABCDEF", [44, 18, 22, 18, 20, 22]):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=6, max_row=ws.max_row, min_col=2, max_col=6):
        for c in row:
            c.number_format = "#,##0;-#,##0"

    ws = wb.create_sheet("De-para")
    cab = ["Conta Gestorial", "Desc.Resultado como veio (Pyplan)", "Desc.Resultado proposto (base 2026)",
           "Ação", "Linhas", "Variabilidade no Pyplan", "R$ ano", "Outros nomes usados em 2026"]
    ws.append(cab)
    for c in ws[1]:
        c.font = neg
    for l in linhas_depara:
        ws.append(l)
        if "troca" in l[3] or "fixo" in l[3]:
            for c in ws[ws.max_row]:
                c.fill = amar
    for col, w in zip("ABCDEFGH", [16, 34, 34, 30, 8, 12, 16, 40]):
        ws.column_dimensions[col].width = w
    for c in ws["G"][1:]:
        c.number_format = "#,##0;-#,##0"
    ws.freeze_panes = "A2"

    ws = wb.create_sheet("Fica fora")
    ws.append(["Fonte", "WH", "Variabilidade", "Conta Gestorial", "Desc.Resultado", "R$ ano", "Motivo"])
    for c in ws[1]:
        c.font = neg
    for f, itens in fora_depois.items():
        for (wh, var, cg, d), v in itens.most_common():
            motivo = "sem conta gestorial (classe fora do plano 0981)" if d == "(vazio)" else "rótulo não existe nessa aba/bloco do Resultado"
            ws.append([f, wh, var, cg, d, round(v, 2), motivo])
    for col, w in zip("ABCDEFG", [34, 10, 12, 16, 30, 14, 46]):
        ws.column_dimensions[col].width = w
    for c in ws["F"][1:]:
        c.number_format = "#,##0;-#,##0"

    SAIDA.mkdir(parents=True, exist_ok=True)
    destino = nome_com_versao(SAIDA, f"De-para Pyplan MP27 - {dt.date.today():%Y%m%d}.xlsx")
    wb.save(destino)
    print("Salvo:", destino)
    for f in fontes:
        print(f"  {f[:40]:40} hoje fora R$ {antes[(f, False)]:>14,.0f} | com de-para fora R$ {depois[(f, False)]:>12,.0f}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
