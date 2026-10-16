# -*- coding: utf-8 -*-
"""mp27_resultado.py -- monta o `Resultado TLM - MP27 vs R09.xlsm` com as formulas lendo a base do MP27.

OBJETIVO: o P&L do MP27 no mesmo esqueleto do Resultado mensal, sem redigitar nada -- quem muda o
          draft de despesas ou preenche o faturamento ve o efeito no P&L.
OPERACAO:
  [1] `Base Total MP27.xlsx` -- tabela `Intermediaria` (mesmas colunas que os 1.994 SUMIFS de cada
      aba de WH consomem: WH, Variabilidade, Desc.Resultado, Cenario, Mes/Ano, January..December),
      alimentada pela aba Draft_MP27 do `Draft MP27 Despesas - <data>.xlsx`. Cenario = Mes/Ano = "MP27".
  [2] copia o `Resultado TLM - Flash Set '26.xlsm` (ja com o LF R09) para TLM_Automation e, via Excel:
      troca o link [1] Base Total_Act_Flash_Fct -> Base Total MP27; poe "MP27" nas linhas 5 e 6 do
      'P&L - TLM' (as abas de WH apontam para la); liga as linhas Tons / Tarifa / Fixo / Ecolabel e as
      Outras Atividades da aba Faturamento a planilha `Faturamento MP27 - base R09` da pasta Volume;
      zera Cabreuva e MP Barueri (nao existem em 2027); recalcula e grava.
ALINHAMENTO: e o que o MP26 fez a mao (`Resultado TLM - MP26 vs R09.xlsx` ligado ao `Base Total
          MP26.XLSX`), agora reproduzivel. Nada e escrito no SAP; o original do Flash nao e tocado.

Pede Excel instalado (win32com) -- a gravacao com openpyxl quebraria os 24 links externos.
"""
import glob
import os
import shutil
import sys
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(BASE_DIR, "scripts", "antidumping"))
from caderno import LINHA_CABECALHO, escrever_caderno                  # noqa: E402

sys.path.insert(0, os.path.join(BASE_DIR, "scripts", "tlm"))
import mp27_tarifa                                                     # noqa: E402

REDE = r"\\SRV054-01BR.group.pirelli.com\DAC_TLM"
ANO_MP = datetime.now().year + 1
ANO_BASE = ANO_MP - 1
AA_MP = ANO_MP % 100
CENARIO = "MP%d" % AA_MP
PASTA_MP = os.path.join(REDE, "MP%d" % AA_MP, "1\u00aa Vers\u00e3o MP%d vs R09" % AA_MP)
PASTA_SAIDA = os.path.join(PASTA_MP, "TLM_Automation")
PASTA_VOLUME = os.path.join(PASTA_MP, "Volume")
FLASH_ORIGEM = os.path.join(REDE, str(ANO_BASE), "2. Flash", "09.Setembro", "Resultado TLM - Flash Set '%d.xlsm" % (ANO_BASE % 100))
BASE_ACT_FLASH = os.path.join(REDE, str(ANO_BASE), "Base Total_Act_Flash_Fct.xlsx")
BASE_MP = os.path.join(PASTA_SAIDA, "Base Total MP%d.xlsx" % AA_MP)
RESULTADO = os.path.join(PASTA_SAIDA, "Resultado TLM - MP%d vs R09.xlsm" % AA_MP)

MESES = ["January", "February", "March", "April", "May", "June",
         "July", "August", "September", "October", "November", "December"]
COLS_TABELA = ["Source.Name", "Conta Gestorial", "Descri\u00e7\u00e3o", "Conta Fiscal", "Descri\u00e7\u00e3o_1",
               "Centro de Custo", "Mini-F\u00e1brica", "Descri\u00e7\u00e3o_2", "Tp.Custo (Pr\u00f3prio /  Aprop.)"] + MESES + [
               "Total Ano", "Usu\u00e1rio", "Data", "Hora", "Variabilidade", "WH", "Gestorial 2", "Cenario",
               "Desc.Resultado", "M\u00eas Despesa", "M\u00eas/Ano", "Chave", "Check Mdo"]

# Aba Faturamento do Resultado <- planilha de preenchimento (Unidade/Cliente, BU/Atividade, Linha).
# fator 1000: a planilha esta em R$ mil / k ton; a aba Faturamento em R$ / ton.
LIGACOES_FATURAMENTO = [
    (4, ("Barueri", "Car", "Tons"), 1000), (5, ("Barueri", "Car", "Tarifa"), 1),
    (12, ("Barueri", "Car", "Valor Liq.Fix"), 1000), (22, ("Barueri", "Car", "Ecolabel"), 1000),
    (33, ("Barueri", "Moto", "Tons"), 1000), (34, ("Barueri", "Moto", "Tarifa"), 1),
    (41, ("Barueri", "Moto", "Valor Liq.Fix"), 1000),
    (56, ("Campinas", "Car", "Tons"), 1000), (57, ("Campinas", "Car", "Tarifa"), 1),
    (64, ("Campinas", "Car", "Valor Liq.Fix"), 1000),
    (67, ("Campinas", "Moto", "Tons"), 1000), (68, ("Campinas", "Moto", "Tarifa"), 1),
    (75, ("Campinas", "Moto", "Valor Liq.Fix"), 1000),
    (137, ("Feira de Santana", "Car", "Tons"), 1000), (138, ("Feira de Santana", "Car", "Tarifa"), 1),
    (145, ("Feira de Santana", "Car", "Valor Liq.Fix"), 1000), (150, ("Feira de Santana", "Car", "Ecolabel"), 1000),
    (162, ("Comercial", "SB.Transp", "Valor Liq."), 1000), (166, ("Comercial", "Campneus", "Valor Liq."), 1000),
    (170, ("Pneus", "SB.Transp", "Valor Liq."), 1000), (189, ("Pneus", "Campneus BA", "Valor Liq."), 1000),
    (195, ("Pneus", "HEVEA", "Valor Liq."), 1000), (199, ("Pneus", "Q.A", "Valor Liq."), 1000),
    (203, ("Pneus", "FS.Check", "Valor Liq."), 1000), (212, ("Pneus", "Fitted", "Valor Liq."), 1000),
]
# Cabreuva (82-127) e MP Barueri (180-184) nao existem em 2027: zera o que alimenta os blocos.
# (181/184 e nao 182: a Area zerada vira divisor das formulas de custo e faturamento do MP BA)
ZERAR_FATURAMENTO = [82, 87, 92, 97, 106, 111, 119, 124, 181, 184]
FIRST_COL_MP = 19          # coluna S = MP27 Jan na planilha de preenchimento (5 chaves + 12 meses de 2026 + total)



def contar_erros(ws):
    """Celulas com erro (#VALUE!, #REF!, #DIV/0!...) lendo o intervalo usado de uma vez: celula a celula via COM
    levava horas no Resultado (ida e volta Python x Excel para cada uma das centenas de milhares de celulas)."""
    v = ws.UsedRange.Value
    if not isinstance(v, tuple):
        v = ((v,),)
    return sum(1 for linha in v for x in linha if isinstance(x, int) and x < -2146826000)

def _ultimo(pasta, padrao):
    arqs = sorted(glob.glob(os.path.join(pasta, padrao)))
    if not arqs:
        raise FileNotFoundError("nenhum %s em %s" % (padrao, pasta))
    return arqs[-1]


# ----------------------------------------------------------------------------- [1] Base Total MP27
def montar_base(draft):
    d = pd.read_excel(draft, sheet_name="Draft_%s" % CENARIO, header=LINHA_CABECALHO - 1,
                      dtype={"Centro de Custo": str, "Conta Gestorial": str, "Conta Fiscal": str})
    mapa = pd.read_excel(BASE_ACT_FLASH, sheet_name="Parametros", usecols=[0, 1], dtype=str).dropna()
    mapa = dict(zip(mapa.iloc[:, 0].str.strip(), mapa.iloc[:, 1]))
    desc_res = d["Conta Fiscal"].map(mapa)
    sem_mapa = d.loc[desc_res.isna(), "Conta Fiscal"].unique()
    desc_res = desc_res.fillna(d["Conta Gestorial Desc"])
    agora = datetime.now()
    t = pd.DataFrame({
        "Source.Name": (d["Enviado por"] + " | " + d["Fonte"]) if "Fonte" in d else os.path.basename(draft),
        "Conta Gestorial": d["Conta Gestorial"], "Descri\u00e7\u00e3o": d["Conta Gestorial Desc"],
        "Conta Fiscal": d["Conta Fiscal"], "Descri\u00e7\u00e3o_1": d["Classe Desc"],
        "Centro de Custo": d["Centro de Custo"], "Mini-F\u00e1brica": "",
        "Descri\u00e7\u00e3o_2": d["Mini-F\u00e1brica Desc"], "Tp.Custo (Pr\u00f3prio /  Aprop.)": d["ORIG."],
    })
    for m in MESES:
        t[m] = d[m].astype(float).round(2)
    t["Total Ano"] = t[MESES].sum(axis=1)
    t["Usu\u00e1rio"] = os.environ.get("USERNAME", "").upper()
    t["Data"] = agora.strftime("%d/%m/%Y")
    t["Hora"] = agora.strftime("%H:%M:%S")
    t["Variabilidade"] = d["VAR."].where(d["Conta Fiscal"] != "N14301S056", "V")   # mesma excecao da base oficial
    t["WH"] = d["WH"]
    t["Gestorial 2"] = ""
    t["Cenario"] = CENARIO
    t["Desc.Resultado"] = desc_res
    t["M\u00eas Despesa"] = t["Total Ano"]
    t["M\u00eas/Ano"] = CENARIO
    t["Chave"] = t["Conta Fiscal"] + t["Centro de Custo"]
    t["Check Mdo"] = t["Conta Fiscal"].str[:4]
    t = t[COLS_TABELA]

    escrever_caderno(BASE_MP, [("Intermediaria", "Base Total %s" % CENARIO,
                                "tabela Intermediaria lida pelo Resultado TLM - %s vs R09 - fonte: %s" % (CENARIO, os.path.basename(draft)), t)])
    wb = load_workbook(BASE_MP)
    ws = wb["Intermediaria"]
    ws.auto_filter.ref = None                       # a tabela traz o proprio filtro; os dois juntos corrompem o arquivo
    ref = "A%d:%s%d" % (LINHA_CABECALHO, get_column_letter(len(COLS_TABELA)), LINHA_CABECALHO + len(t))
    tab = Table(displayName="Intermediaria", ref=ref)
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
    ws.add_table(tab)
    wb.save(BASE_MP)
    print("  Base Total: %d linhas, total %s R$ %.0f, conta fiscal sem Desc.Resultado no mapa: %d %s"
          % (len(t), CENARIO, t["Total Ano"].sum(), len(sem_mapa), list(sem_mapa)[:8]))
    return t


# ----------------------------------------------------------------------------- [2] Resultado
def linhas_da_planilha(fat):
    """(Unidade/Cliente, BU/Atividade, Linha) -> numero da linha na aba Faturamento da planilha de preenchimento."""
    f = pd.read_excel(fat, sheet_name="Faturamento", header=LINHA_CABECALHO - 1, usecols=[0, 1, 2, 3])
    return {(r["Unidade / Cliente"], r["BU / Atividade"], r["Linha"]): LINHA_CABECALHO + 1 + i for i, r in f.iterrows()}


def montar_resultado(fat):
    import win32com.client
    shutil.copy2(FLASH_ORIGEM, RESULTADO)
    linhas = linhas_da_planilha(fat)
    ref_fat = "'%s\\[%s]Faturamento'" % (os.path.dirname(fat), os.path.basename(fat))

    xl = mp27_tarifa.abrir_excel()
    xl.EnableEvents = False
    try:
        base = xl.Workbooks.Open(BASE_MP, UpdateLinks=0, ReadOnly=True)     # aberta: referencia estruturada externa so resolve com a fonte aberta
        wb = xl.Workbooks.Open(RESULTADO, UpdateLinks=0)
        antigo = [s for s in (wb.LinkSources(1) or []) if os.path.basename(str(s)).lower() == "base total_act_flash_fct.xlsx"]
        if not antigo:
            raise RuntimeError("link da Base Total_Act_Flash_Fct nao encontrado no Resultado")
        wb.ChangeLink(antigo[0], BASE_MP, 1)
        print("  link trocado:", os.path.basename(antigo[0]), "->", os.path.basename(BASE_MP))

        pl = wb.Worksheets("P&L - TLM")
        pl.Range("B5:M5").Value = CENARIO
        pl.Range("B6:M6").Value = CENARIO
        wb.Worksheets("Parametros").Range("J3").Value = CENARIO

        ws = wb.Worksheets("Faturamento")
        ligadas = 0
        for linha, chave, fator in LIGACOES_FATURAMENTO:
            if chave not in linhas:
                print("  AVISO: %s nao existe na planilha de preenchimento - linha %d mantida" % (chave, linha))
                continue
            for j in range(12):
                origem = "%s!%s%d" % (ref_fat, get_column_letter(FIRST_COL_MP + j), linhas[chave])
                ws.Cells(linha, 2 + j).Formula = "=%s*%d" % (origem, fator) if fator != 1 else "=%s" % origem
            ligadas += 1
        for linha in ZERAR_FATURAMENTO:
            ws.Range(ws.Cells(linha, 2), ws.Cells(linha, 13)).Value = 0
        print("  Faturamento: %d linhas ligadas a planilha de preenchimento, %d zeradas (Cabreuva / MP BA)" % (ligadas, len(ZERAR_FATURAMENTO)))

        # SUMIF/SUMIFS em pasta externa FECHADA devolve #VALUE! (HC [21], KSB1 [23], faturamentos mensais):
        # abre as fontes que existem antes de recalcular, como estavam quando o Flash foi salvo.
        abertas = []
        for fonte in (wb.LinkSources(1) or []):
            fonte = str(fonte)
            if os.path.normcase(fonte) in (os.path.normcase(BASE_MP), os.path.normcase(fat)) or not os.path.exists(fonte):
                continue
            # outro Resultado (MP26 do bloco de comparacao) fica FECHADO: so ha referencia direta a celula, que le o valor
            # gravado; aberto, ele se recalcula com as fontes DELE fechadas e devolve #VALUE! em todo o bloco AD:AP
            if os.path.basename(fonte).lower().startswith("resultado tlm"):
                continue
            try:
                abertas.append(xl.Workbooks.Open(fonte, UpdateLinks=0, ReadOnly=True))
            except Exception as e:                                       # noqa: BLE001 - fonte velha/corrompida nao trava o resto
                print("  AVISO: nao abriu %s (%s)" % (os.path.basename(fonte), str(e)[:60]))
        modelo = xl.Workbooks.Open(fat, UpdateLinks=0, ReadOnly=True)
        print("  fontes abertas para o recalculo: %d" % (len(abertas) + 2))

        xl.CalculateFull()
        checks = {
            "Faturamento liquido FY (N13)": pl.Range("N13").Value,
            "Custos variaveis FY (N21)": pl.Range("N21").Value,
            "Custos fixos FY (N50)": pl.Range("N50").Value,
            "TOTAL COST FY (N135)": pl.Range("N135").Value,
            "PBIT FY (N124)": pl.Range("N124").Value,
        }
        erros = {}
        for nome in ("P&L - TLM", "Faturamento", "BALog", "CPLog", "P&L Log"):
            erros[nome] = contar_erros(wb.Worksheets(nome))
        wb.Save()
        wb.Close(SaveChanges=False)
        base.Close(SaveChanges=False)
        modelo.Close(SaveChanges=False)
        for a in abertas:
            if a is not None:               # Workbooks.Open devolve None para fonte que ja estava aberta (xlsm com macro)
                a.Close(SaveChanges=False)
    finally:
        xl.Quit()
    return checks, erros


def main():
    draft = _ultimo(PASTA_SAIDA, "Draft %s Despesas - *.xlsx" % CENARIO)
    fat = _ultimo(PASTA_VOLUME, "Faturamento %s - base R09 - *.xlsx" % CENARIO)
    print("[1] Base Total %s <- %s" % (CENARIO, os.path.basename(draft)))
    t = montar_base(draft)
    print("[2] Resultado <- %s" % os.path.basename(FLASH_ORIGEM))
    checks, erros = montar_resultado(fat)
    print("\nSaida:", RESULTADO)
    for k, v in checks.items():
        print("  %-32s %14s" % (k, ("%.0f" % v) if isinstance(v, (int, float)) else v))
    print("  celulas de erro por aba:", erros)
    print("  base %s (R$ mil, custo positivo): %.0f" % (CENARIO, t["Total Ano"].sum() / 1000))


if __name__ == "__main__":
    main()
