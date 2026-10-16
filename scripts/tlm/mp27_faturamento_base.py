# -*- coding: utf-8 -*-
"""mp27_faturamento_base.py -- planilha de preenchimento do faturamento MP27 da TLM (0981).

OBJETIVO: dar a quem preenche o MP27 a base do R09 (Actual jan-ago + Fcst set-dez) por unidade x BU
          e, ao lado, as colunas do MP27 em branco para digitar tarifa, toneladas, ecolabel e fixo.
OPERACAO: le a `Analise Fat. R09 vs R07.xlsx` da pasta do R09 (blocos Barueri / Campinas / Cabreuva /
          Feira de Santana x Car / Moto + Outras Atividades), grava na pasta `Volume` do MP27 no padrao
          azul, com celulas de entrada em amarelo e Valor Liq.Var = Tarifa x Tons por formula.
ALINHAMENTO: e o que o Caetano pediu em 28/09/2026 ("dados de R09 do faturamento como base e colunas
          para preenchimento do MP27"). Cabreuva e MP Barueri saem em 2027: ficam na base como referencia,
          marcados "sai", sem colunas de entrada. Nada e escrito no SAP.
"""
import os
import sys
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(BASE_DIR, "scripts", "antidumping"))
from caderno import LINHA_CABECALHO, escrever_caderno                  # noqa: E402
sys.path.insert(0, os.path.join(BASE_DIR, "scripts", "tlm"))
import mp27_tarifa                                                     # noqa: E402

REDE = r"\\SRV054-01BR.group.pirelli.com\DAC_TLM"
ANO_MP = datetime.now().year + 1
ANO_BASE = ANO_MP - 1
AA_MP = ANO_MP % 100
ROUND = "R09"
FONTE = os.path.join(REDE, str(ANO_BASE), "1. FCST", ROUND, "An\u00e1lise Fat. %s vs R07.xlsx" % ROUND)
ABA_FONTE = "Analise Faturamento vs Fct"
PASTA_VOLUME = os.path.join(REDE, "MP%d" % AA_MP, "1\u00aa Vers\u00e3o MP%d vs R09" % AA_MP, "Volume")
RESULTADO_MP = os.path.join(os.path.dirname(PASTA_VOLUME), "TLM_Automation", "Resultado TLM - MP%d vs R09.xlsm" % AA_MP)
RESULTADO_R09 = os.path.join(REDE, str(ANO_BASE), "1. FCST", ROUND, "Resultado TLM - %s '%d.xlsm" % (ROUND, ANO_BASE % 100))

MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
LINHAS = ["Tarifa", "Tons", "Valor Liq.Var", "Ecolabel", "Valor Liq.Fix"]
LINHAS_VALOR = ["Valor Liq.Var", "Ecolabel", "Valor Liq.Fix"]
UNIDADES_FORA = {"Cabre\u00fava": "sai - Cabre\u00fava n\u00e3o existe em %d" % ANO_MP}
OUTRAS_FORA = {"MP BA": "sai - faturamento MP Barueri n\u00e3o existe em %d" % ANO_MP}
# Caetano 08/10/2026: enquanto a tecnica nao entrega o Volume MP27, simular com as toneladas do R09 mes a mes
SIMULAR_TONS_R09 = True
LARANJA = PatternFill("solid", fgColor="F8CBAD")
AMARELO = PatternFill("solid", fgColor="FFF2CC")
CINZA = PatternFill("solid", fgColor="EDEDED")


def ler_fonte():
    """Devolve as linhas do bloco Log (unidade x BU x linha) e das Outras Atividades, em R$ mil / k ton."""
    df = pd.read_excel(FONTE, sheet_name=ABA_FONTE, header=None)
    log, outras = [], []
    unidade, bu, cliente, em_outras = None, None, None, False
    for _, r in df.iterrows():
        a, b = r[0], r[1]
        b = str(b).strip() if pd.notna(b) else ""
        if b.startswith("Total TLM"):
            break
        if pd.notna(a):
            a = str(a).strip()
            if a in ("Comercial", "Pneus"):
                cliente, em_outras = a, True
            elif a not in ("Car", "Moto"):
                unidade = a
        if not b:
            continue
        if em_outras:
            if b.startswith("Total"):
                continue
            outras.append({"Cliente": cliente, "Atividade": b, "meses": r[2:14].tolist(), "total": r[14]})
            continue
        if b.startswith("Total"):
            unidade = None                  # fecha o bloco; os consolidados "Total CAR/Moto" que vem depois sao ignorados
        elif b in ("Car", "Moto"):
            bu = b                          # o cabecalho da BU pode vir ANTES do nome da unidade (Campinas)
        elif unidade is None:
            continue
        elif b in LINHAS:
            log.append({"Unidade": unidade, "BU": bu, "Linha": b, "meses": r[2:14].tolist(), "total": r[14]})
    return log, outras


def montar():
    log, outras = ler_fonte()
    c26 = ["%d %s" % (ANO_BASE, m) for m in MESES]
    cmp_ = ["MP%d %s" % (AA_MP, m) for m in MESES]
    t26, tmp = "Total %d" % ANO_BASE, "Total MP%d" % AA_MP
    colunas = ["Bloco", "Unidade / Cliente", "BU / Atividade", "Linha", "Per\u00edmetro %d" % ANO_MP] + c26 + [t26] + cmp_ + [tmp, "Delta", "Como preencher"]

    linhas, tipos = [], []        # tipos: entrada | formula_var | total | sai | outras

    def add(bloco, u, bu, linha, per, meses, total, tipo, dica):
        d = {"Bloco": bloco, "Unidade / Cliente": u, "BU / Atividade": bu, "Linha": linha, "Per\u00edmetro %d" % ANO_MP: per,
             "Como preencher": dica}
        for c, v in zip(c26, meses):
            d[c] = float(v) if pd.notna(v) else 0.0
        d[t26] = float(total) if pd.notna(total) else sum(d[c] for c in c26)
        linhas.append(d)
        tipos.append(tipo)

    for u in dict.fromkeys(x["Unidade"] for x in log):
        per = UNIDADES_FORA.get(u, "fica")
        for bu in ("Car", "Moto"):
            bloco = [x for x in log if x["Unidade"] == u and x["BU"] == bu]
            if not bloco:
                continue
            for x in bloco:
                li = x["Linha"]
                if per != "fica":
                    tipo, dica = "sai", "n\u00e3o preencher - unidade fora do MP%d" % AA_MP
                elif li == "Valor Liq.Var":
                    tipo, dica = "formula_var", "calculado = Tarifa x Tons"
                elif li == "Tarifa":
                    tipo, dica = "entrada", "R$/ton - digitar (refer\u00eancia: infla\u00e7\u00e3o MP2027 4,4%% sobre %s)" % ROUND
                elif li == "Tons":
                    tipo, dica = "entrada", "k ton - digitar (vem do Volume Movimentado MP%d)" % AA_MP
                else:
                    tipo, dica = "entrada", "R$ mil - digitar"
                add("Log", u, bu, li, per, x["meses"], x["total"], tipo, dica)
        add("Log", u, "Total", "Total %s" % u, per, [0.0] * 12, 0.0, "total", "soma dos valores Car + Moto")
    for x in outras:
        per = OUTRAS_FORA.get(x["Atividade"], "fica")
        tipo = "sai" if per != "fica" else "outras"
        dica = "n\u00e3o preencher - fora do MP%d" % AA_MP if tipo == "sai" else "R$ mil - digitar"
        add("Outras Atividades", x["Cliente"], x["Atividade"], "Valor Liq.", per, x["meses"], x["total"], tipo, dica)
    add("Outras Atividades", "Total", "Total", "Total Outras Atividades", "fica", [0.0] * 12, 0.0, "total", "soma")
    add("TLM", "Total", "Total", "Total TLM", "fica", [0.0] * 12, 0.0, "total", "Log + Outras Atividades")

    df = pd.DataFrame(linhas, columns=colunas)
    # totais de 2026 recalculados a partir das linhas de valor (a fonte tem os mesmos numeros)
    for i, tp in enumerate(tipos):
        if tp != "total":
            continue
        if df.at[i, "Bloco"] == "Log":
            u = df.at[i, "Unidade / Cliente"]
            sel = (df["Bloco"] == "Log") & (df["Unidade / Cliente"] == u) & (df["Linha"].isin(LINHAS_VALOR))
        elif df.at[i, "Linha"].startswith("Total Outras"):
            sel = (df["Bloco"] == "Outras Atividades") & (df["Linha"] == "Valor Liq.")
        else:
            sel = (df["Linha"].isin(LINHAS_VALOR)) | (df["Linha"] == "Valor Liq.")
        for c in c26 + [t26]:
            df.at[i, c] = df.loc[sel, c].sum()

    premissas = pd.DataFrame([
        ("Base", "%s %d: Actual jan-ago + Fcst set-dez, da 'An\u00e1lise Fat. %s vs R07.xlsx' (R$ mil e k ton)" % (ROUND, ANO_BASE, ROUND)),
        ("Como preencher", "c\u00e9lulas AMARELAS = digitar (Tons, Ecolabel e Outras Atividades); VERDES = calculadas pela aba Tarifa (Tarifa e Valor Liq.Fix); Valor Liq.Var, totais e Delta s\u00e3o f\u00f3rmula"),
        ("Per\u00edmetro", "Cabre\u00fava n\u00e3o existe em %d - fica na base como refer\u00eancia, sem entrada" % ANO_MP),
        ("Per\u00edmetro", "Faturamento MP Barueri (MP BA) n\u00e3o existe em %d - idem" % ANO_MP),
        ("Refer\u00eancia macro", "infla\u00e7\u00e3o m\u00e9dia Brasil MP2027 4,4% - BRL/USD 5,25 - BRL/EUR 6,14 (MP27\\Macroeconomics + ForEx 1st draft 10_09)"),
        ("Toneladas", ("SIMULA\u00c7\u00c3O (c\u00e9lulas LARANJA): toneladas, Ecolabel e Outras Atividades do %s repetidos m\u00eas a m\u00eas no MP%d - trocar pelo Volume Movimentado MP%d da t\u00e9cnica quando chegar" % (ROUND, AA_MP, AA_MP))
         if SIMULAR_TONS_R09 else "vir\u00e3o do Volume Movimentado MP%d da \u00e1rea t\u00e9cnica (mesma pasta) - at\u00e9 l\u00e1 ficam em branco" % AA_MP),
        ("Tarifa - regra", "a do MP26 (abas Tarifa / Tarifa por B.U do Resultado TLM - MP26 vs R09): tarifa a custo, unidade por unidade - a conta passo a passo esta na aba Tarifa"),
        ("Tarifa variável", "custo variável da WH (P&L linha 21) x % BU / toneladas da BU x (1 + markup) - markup Barueri 2%, Campinas e Feira 0%"),
        ("Tarifa fixa", "(custo fixo da WH (linha 50) - questões trabalhistas (linha 83) - receita Ecolabel) / m2 utilizados no ano - a mesma para Car e Moto"),
        ("De onde vem a despesa", "abas de WH do Resultado TLM - MP%d vs R09 (Barueri = BALog + BAADM, Campinas = CPLog, Feira = FSLog), copiadas na aba Custos com a célula de origem" % AA_MP),
        ("% BU", "aba BU: coeficiente de empenho de M.O do MP26 (Barueri 76/24, Campinas 50/50) ou %% pela tonelada carregada MP%d - critério em BU!C4" % AA_MP),
        ("Área", "aba Area_m2: m2 do R09 por mês e MP%d em amarelo (começa igual a dezembro do R09)" % AA_MP),
        ("Comparativo", "aba Comparativo_R09: R09 praticado x R09 pela mesma conta x MP%d (toneladas, m2, custo, tarifa, faturamento, margem)" % AA_MP),
        ("Células verdes", "Tarifa e Valor Liq.Fix MP%d da aba Faturamento são calculados pela aba Tarifa; sobrescrever só se houver tarifa de contrato" % AA_MP),
        ("Bruto", "líquido / (1 - impostos): Barueri 11,25%, Campinas e Feira 14,25%"),
        ("Fora da conta", "Outras Atividades seguem digitadas (MP26: SB.Transp = custo fixo / 3 clientes; CP e FS Fab = custo / 0,98; R&D CB e CP Check = R09 x 1,044)"),
        ("Gerado", "%s por scripts/tlm/mp27_faturamento_base.py" % datetime.now().strftime("%d/%m/%Y %H:%M")),
    ], columns=["Item", "Valor"])

    os.makedirs(PASTA_VOLUME, exist_ok=True)
    destino = os.path.join(PASTA_VOLUME, "Faturamento MP%d - base %s - %s.xlsx" % (AA_MP, ROUND, datetime.now().strftime("%Y%m%d")))
    escrever_caderno(destino, [
        ("Premissas", "Faturamento MP%d TLM - como usar" % AA_MP, "", premissas),
        ("Faturamento", "Faturamento MP%d - base %s %d x preenchimento MP%d" % (AA_MP, ROUND, ANO_BASE, AA_MP),
         "R$ mil e k ton - amarelo = digitar - Valor Liq.Var = Tarifa x Tons", df),
    ])
    linha_de = aplicar_formulas(destino, df, tipos, c26, cmp_, t26, tmp)
    custos, _ = mp27_tarifa.acrescentar(destino, linha_de, RESULTADO_MP, RESULTADO_R09, col_mp0=len(colunas) - len(cmp_) - 3 + 1)
    print("Tarifa: custos lidos de %s (%d series)" % (os.path.basename(RESULTADO_MP), len(custos)))
    erros, total = mp27_tarifa.recalcular(destino)
    print("Recalculado no Excel - Total TLM MP%d R$ mil: %s | celulas de erro: %s" % (AA_MP, total, erros))
    print("\nSaida:", destino)
    print("Linhas:", len(df), "| unidades:", sorted(df.loc[df.Bloco == "Log", "Unidade / Cliente"].unique()))
    print("Total %d (Log + Outras) R$ mil: %.0f" % (ANO_BASE, df.loc[df["Linha"] == "Total TLM", t26].iloc[0]))
    return destino


def aplicar_formulas(destino, df, tipos, c26, cmp_, t26, tmp):
    wb = load_workbook(destino)
    ws = wb["Faturamento"]
    col = {c: i + 1 for i, c in enumerate(df.columns)}
    L = get_column_letter
    r0 = LINHA_CABECALHO + 1
    linha_de = {}                                   # (unidade, bu, linha) -> numero da linha na aba
    for i in range(len(df)):
        linha_de[(df.at[i, "Unidade / Cliente"], df.at[i, "BU / Atividade"], df.at[i, "Linha"])] = r0 + i

    for i, tp in enumerate(tipos):
        r = r0 + i
        u, bu = df.at[i, "Unidade / Cliente"], df.at[i, "BU / Atividade"]
        # Tons, Ecolabel e Outras Atividades = curva do R09 mes a mes (Caetano 08/10: "deixar carregado para o Resultado pegar")
        simula = SIMULAR_TONS_R09 and tp in ("entrada", "outras")
        for c, c_base in zip(cmp_, c26):
            cell = ws.cell(row=r, column=col[c])
            cell.number_format = "#,##0.00"
            if simula:
                cell.value = float(df.at[i, c_base])
                cell.fill = LARANJA
            elif tp in ("entrada", "outras"):
                cell.fill = AMARELO
            elif tp == "sai":
                cell.value = 0
                cell.fill = CINZA
            elif tp == "formula_var":
                ct = L(col[c])
                cell.value = "=%s%d*%s%d" % (ct, linha_de[(u, bu, "Tarifa")], ct, linha_de[(u, bu, "Tons")])
            elif tp == "total":
                cell.value = _soma_total(df, i, r0, L(col[c]))
        if simula:
            fonte = ("trocar pelo Volume Movimentado MP%d da técnica" % AA_MP) if df.at[i, "Linha"] == "Tons" \
                else "trocar pela proposta MP%d da atividade quando houver" % AA_MP
            ws.cell(row=r, column=col["Como preencher"],
                    value="SIMULAÇÃO: curva do %s mês a mês - %s" % (ROUND, fonte))
        # total MP27 e delta
        if df.at[i, "Linha"] == "Tarifa":
            tons_r = linha_de[(u, bu, "Tons")]
            var_r = linha_de[(u, bu, "Valor Liq.Var")]
            for tcol, first, last in ((t26, c26[0], c26[-1]), (tmp, cmp_[0], cmp_[-1])):
                ws.cell(row=r, column=col[tcol]).value = "=IFERROR(%s%d/%s%d,0)" % (L(col[tcol]), var_r, L(col[tcol]), tons_r)
        else:
            ws.cell(row=r, column=col[tmp]).value = "=SUM(%s%d:%s%d)" % (L(col[cmp_[0]]), r, L(col[cmp_[-1]]), r)
        ws.cell(row=r, column=col[tmp]).number_format = "#,##0.00"
        ws.cell(row=r, column=col["Delta"]).value = "=%s%d-%s%d" % (L(col[tmp]), r, L(col[t26]), r)
        ws.cell(row=r, column=col["Delta"]).number_format = "#,##0.00"
        if tp == "total":
            for j in range(1, len(df.columns) + 1):
                ws.cell(row=r, column=j).font = Font(bold=True)
    wb.save(destino)
    return linha_de


def _soma_total(df, i, r0, letra):
    """Formula do total MP27 de uma linha 'total': soma das linhas de valor do seu grupo."""
    if df.at[i, "Bloco"] == "Log":
        u = df.at[i, "Unidade / Cliente"]
        sel = (df["Bloco"] == "Log") & (df["Unidade / Cliente"] == u) & (df["Linha"].isin(LINHAS_VALOR))
    elif df.at[i, "Linha"].startswith("Total Outras"):
        sel = (df["Bloco"] == "Outras Atividades") & (df["Linha"] == "Valor Liq.")
    else:
        sel = (df["Linha"].isin(LINHAS_VALOR)) | (df["Linha"] == "Valor Liq.")
    refs = ["%s%d" % (letra, r0 + k) for k in df.index[sel]]
    return "=" + "+".join(refs) if refs else 0


if __name__ == "__main__":
    montar()
