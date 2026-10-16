# -*- coding: utf-8 -*-
"""mp27_resumo.py -- `Resumo Resultado TLM - MP27 vs R09.xlsx`, o mesmo resumo do Flash / FCT para o MP27.

OBJETIVO: o MP27 sai no mesmo formato de resumo que a area ja usa nos Flash e nos rounds (Resumo FY /
          Month / YTD / YTG, P&L - TLM, Resumo de Resultado), comparando MP27 x R09 '26 x MP26.
OPERACAO: copia o `Resumo Resultado TLM - R09 '26.xlsx` da pasta do R09 para TLM_Automation e, via Excel:
          troca o link do `Resultado TLM - R09 '26.xlsm` pelo `Resultado TLM - MP27 vs R09.xlsm`;
          Parametros B3 = MP27 e C3 = R09 (as colunas C / G do resumo seguem); titulo "MP27 vs R09";
          limpa o que era digitado a mao para o R09 (coluna Flash, ganhos/linhas avulsas, KSB1 do mes);
          recalcula (Resultado fechado: le os valores gravados nele) e grava.
ALINHAMENTO: e o resumo que a area ja le, so que com o cenario MP27 -- nenhuma formula reinventada.
          Nada e escrito no SAP; o resumo do R09 nao e tocado.

Pede Excel instalado (win32com): o resumo tem links externos que o openpyxl quebraria.
"""
import os
import shutil
import sys
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(BASE_DIR, "scripts", "tlm"))
import mp27_tarifa                                                     # noqa: E402

REDE = r"\\SRV054-01BR.group.pirelli.com\DAC_TLM"
ANO_MP = datetime.now().year + 1
ANO_BASE = ANO_MP - 1
AA_MP, AA_BASE = ANO_MP % 100, ANO_BASE % 100
CENARIO, ROUND = "MP%d" % AA_MP, "R09"
PASTA_SAIDA = os.path.join(REDE, "MP%d" % AA_MP, "1\u00aa Vers\u00e3o MP%d vs R09" % AA_MP, "TLM_Automation")
RESUMO_ORIGEM = os.path.join(REDE, str(ANO_BASE), "1. FCST", ROUND, "Resumo Resultado TLM - %s '%d.xlsx" % (ROUND, AA_BASE))
RESULTADO_ROUND = "Resultado TLM - %s '%d.xlsm" % (ROUND, AA_BASE)
RESULTADO_MP = os.path.join(PASTA_SAIDA, "Resultado TLM - %s vs %s.xlsm" % (CENARIO, ROUND))
RESUMO = os.path.join(PASTA_SAIDA, "Resumo Resultado TLM - %s vs %s.xlsx" % (CENARIO, ROUND))

TITULO = '="Resultado TLM - "&Parametros!B3&" vs "&Parametros!C3&" \'%d"' % AA_BASE
ABAS_RESUMO = ("Resumo FY", "Resumo Month", "Resumo YTD")
# digitado a mao no resumo do R09 -- nao vale para o MP27
LIMPAR_FY = ["C41", "G41", "C47:C54", "G47:G54", "C67:C70", "L71:L72"]



def contar_erros(ws):
    """Celulas com erro (#VALUE!, #REF!, #DIV/0!...) lendo o intervalo usado de uma vez: celula a celula via COM
    levava horas no Resultado (ida e volta Python x Excel para cada uma das centenas de milhares de celulas)."""
    v = ws.UsedRange.Value
    if not isinstance(v, tuple):
        v = ((v,),)
    return sum(1 for linha in v for x in linha if isinstance(x, int) and x < -2146826000)

def montar():
    import win32com.client
    shutil.copy2(RESUMO_ORIGEM, RESUMO)
    xl = mp27_tarifa.abrir_excel()
    xl.EnableEvents = False
    try:
        # o Resultado fica FECHADO: o resumo so tem referencia direta a celula (le o valor gravado); aberto, o
        # CalculateFull o recalcula com as fontes DELE fechadas e a coluna MP27 vira #REF!/#VALUE!
        wb = xl.Workbooks.Open(RESUMO, UpdateLinks=0)
        antigo = [s for s in (wb.LinkSources(1) or []) if os.path.basename(str(s)).lower() == RESULTADO_ROUND.lower()]
        if not antigo:
            raise RuntimeError("link do %s nao encontrado no resumo" % RESULTADO_ROUND)
        wb.ChangeLink(antigo[0], RESULTADO_MP, 1)
        print("  link trocado:", os.path.basename(str(antigo[0])), "->", os.path.basename(RESULTADO_MP))

        par = wb.Worksheets("Parametros")
        par.Range("B3").Value = CENARIO
        par.Range("C3").Value = ROUND
        par.Range("D3").Value = 12

        for nome in ABAS_RESUMO + ("Resumo YTG",):
            ws = wb.Worksheets(nome)
            ws.Range("B1").Formula = TITULO
            ws.Range("E3").Value = "-"
            if nome in ABAS_RESUMO:                          # coluna Flash: valores do mes colados a mao
                for c in ws.Range("E6:E63").Cells:
                    if not c.HasFormula:
                        c.Value = None
        # abertura por WH do YTD/YTG usa INDIRECT com o nome do Resultado escrito como TEXTO (o ChangeLink nao alcanca);
        # como no Flash/FCT, so resolve com o Resultado aberto junto
        for nome in ("Resumo YTD", "Resumo YTG", "Resumo Month", "Resumo FY"):
            wb.Worksheets(nome).UsedRange.Replace("[%s]" % RESULTADO_ROUND.replace("'", "''"),
                                                  "[%s]" % os.path.basename(RESULTADO_MP), 2)
        fy = wb.Worksheets("Resumo FY")
        for rng in LIMPAR_FY:
            for c in fy.Range(rng).Cells:
                if not c.HasFormula or "Sheet2" in str(c.Formula):
                    c.Value = None
        fy.Range("C110").Value = CENARIO
        fy.Range("E110").Value = "-"
        fy.Range("G110").Value = ROUND

        xl.CalculateFull()
        checks = {k: fy.Range(c).Value for k, c in (
            ("Net sales MP27 (C6)", "C6"), ("Net sales R09 (G6)", "G6"), ("Net sales MP26 (I6)", "I6"),
            ("Tot. var MP27 (C11)", "C11"), ("Tot. fixed MP27 (C19)", "C19"),
            ("PBIT MP27 (C27)", "C27"), ("PBIT R09 (G27)", "G27"), ("PBIT MP26 (I27)", "I27"))}
        erros = {}
        for nome in ABAS_RESUMO + ("Resumo YTG", "P&L - TLM", "Resumo de Resultado"):
            ws = wb.Worksheets(nome)
            erros[nome] = contar_erros(ws)
        links = [os.path.basename(str(s)) for s in (wb.LinkSources(1) or [])]
        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()
    return checks, erros, links


def main():
    print("Resumo <- %s" % os.path.basename(RESUMO_ORIGEM))
    checks, erros, links = montar()
    print("\nSaida:", RESUMO)
    for k, v in checks.items():
        print("  %-26s %14s" % (k, ("%.0f" % v) if isinstance(v, (int, float)) else v))
    print("  celulas de erro por aba:", erros)
    print("  links:", links)


if __name__ == "__main__":
    main()
