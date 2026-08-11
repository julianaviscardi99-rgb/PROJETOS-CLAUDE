#!/usr/bin/env python3
"""
Analisa a ZLFIB (Pesquisa genérica de Notas Fiscais) em busca de Notas Fiscais
duplicadas — projeto "Fitted Recuperação".

Regras confirmadas pela Juliana em 2026-08-11:
- "Lista por Item" (não "Lista por Nota", que não funciona bem nessa transação).
  Cada NF aparece repetida uma vez por item — precisa agrupar por Nr Documento
  (DOCNUM) para voltar ao nível de NF antes de procurar duplicidade.
- Filiais da Fitted Units: 0031=SJP, 0032=IBI, 0053=SOR, 0054=GOI.
- "Buscar chave de acesso" (ACKEY) marcado: notas de mercadoria têm chave de
  acesso (identificador único da NF de verdade — se a mesma chave aparece em
  mais de um Nr Documento, é duplicidade certa). Notas de serviço NÃO têm
  chave de acesso — nesses casos, duplicidade é inferida por
  Parceiro + Nota Fiscal + Série + Valor total repetidos em mais de um
  Nr Documento.

Lê a grid ALV direto via COM (GetCellValue) em vez de exportar para Excel —
mais rápido e mais simples que negociar o menu de exportação dessa tela.
"""
import time
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

from atualizar_ksb1_gui import connect_session

FILIAIS = {"0031": "SJP", "0032": "IBI", "0053": "SOR", "0054": "GOI"}
DATA_DE = "01.01.2026"
DATA_ATE = "31.07.2026"

COLS = ["BRANCH", "DOCNUM", "NFNUM", "SERIES", "NFTYPE", "PSTDAT", "DOCDAT", "PARID", "NAME1", "NFTOT", "ACKEY"]

PASTA_DESTINO = Path(
    r"\\FSS024-01BR.group.pirelli.com\GFU_DAC\Custos Fitted Units\Estudos\Estudo Duplicidade Pagamento"
)


def abrir_zlfib(session, log):
    log("Abrindo a transação ZLFIB...")
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nZLFIB"
    session.FindById("wnd[0]").SendVKey(0)
    time.sleep(1)
    if session.Info.Transaction != "ZLFIB":
        raise RuntimeError(f"Não consegui abrir a ZLFIB (tela atual: '{session.Info.Transaction}').")


def buscar_filial(session, filial: str, log):
    wnd = session.FindById("wnd[0]")
    wnd.FindById("usr/ctxtP_BUKRS").Text = "0580"
    wnd.FindById("usr/ctxtS_BRANCH-LOW").Text = filial
    wnd.FindById("usr/ctxtS_PSTDAT-LOW").Text = DATA_DE
    wnd.FindById("usr/ctxtS_PSTDAT-HIGH").Text = DATA_ATE
    wnd.FindById("usr/radP_ITEM").Select()
    wnd.FindById("usr/chkP_ACKEY").Selected = True
    log(f"Executando ZLFIB — Filial {filial} ({FILIAIS[filial]})...")
    wnd.SendVKey(8)
    time.sleep(2)


def ler_grid(session, filial: str, log):
    grid = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")
    total = grid.RowCount
    log(f"  {total} linha(s) de item na grid, lendo...")
    linhas = []
    for r in range(total):
        valores = [grid.GetCellValue(r, c) for c in COLS]
        linhas.append(dict(zip(COLS, valores)))
    return linhas


def parse_valor(txt: str) -> float:
    if not txt:
        return 0.0
    return float(txt.replace(".", "").replace(",", "."))


def coletar_todas_filiais(log=print):
    session = connect_session()
    todas_linhas = []
    for filial in FILIAIS:
        abrir_zlfib(session, log)
        buscar_filial(session, filial, log)
        linhas = ler_grid(session, filial, log)
        todas_linhas.extend(linhas)
    return todas_linhas


def colapsar_por_nf(linhas_item: list) -> list:
    """Uma NF vem repetida uma linha por item (mesmo DOCNUM); reduz a uma
    linha por Nr Documento, já que os campos de cabeçalho (NFNUM, SERIES,
    NFTOT, ACKEY etc.) se repetem idênticos em todas as linhas do item."""
    por_doc = {}
    qtd_itens = defaultdict(int)
    for l in linhas_item:
        qtd_itens[l["DOCNUM"]] += 1
        por_doc.setdefault(l["DOCNUM"], l)
    notas = []
    for docnum, cab in por_doc.items():
        nota = dict(cab)
        nota["QTD_ITENS"] = qtd_itens[docnum]
        nota["valor"] = parse_valor(cab["NFTOT"])
        notas.append(nota)
    return notas


def encontrar_duplicidades(notas: list):
    com_chave = [n for n in notas if n["ACKEY"]]
    sem_chave = [n for n in notas if not n["ACKEY"]]

    grupos_chave = defaultdict(list)
    for n in com_chave:
        grupos_chave[n["ACKEY"]].append(n)
    dup_chave = {k: v for k, v in grupos_chave.items() if len({n["DOCNUM"] for n in v}) > 1}

    grupos_sem_chave = defaultdict(list)
    for n in sem_chave:
        chave = (n["PARID"], n["NFNUM"], n["SERIES"], round(n["valor"], 2))
        grupos_sem_chave[chave].append(n)
    dup_sem_chave = {
        k: v for k, v in grupos_sem_chave.items() if len({n["DOCNUM"] for n in v}) > 1
    }

    return dup_chave, dup_sem_chave


def montar_planilha(wb, titulo, grupos, log=print):
    ws = wb.create_sheet(titulo[:31])
    ws.append([
        "Filial", "Nr Documento", "Nota Fiscal", "Série", "Tipo NF", "Data lançamento",
        "Parceiro", "Nome", "Valor total NF", "Qtd. itens", "Chave de acesso",
    ])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for chave, grupo in sorted(grupos.items(), key=lambda kv: -kv[1][0]["valor"] * len(kv[1])):
        for n in grupo:
            ws.append([
                n["BRANCH"], n["DOCNUM"], n["NFNUM"], n["SERIES"], n["NFTYPE"], n["PSTDAT"],
                n["PARID"], n["NAME1"], n["valor"], n["QTD_ITENS"], n["ACKEY"],
            ])
    return ws


def analisar(log=print) -> Path:
    linhas_item = coletar_todas_filiais(log)
    notas = colapsar_por_nf(linhas_item)
    log(f"\n{len(linhas_item)} linha(s) de item no total -> {len(notas)} Nota(s) Fiscal(is) únicas (por Nr Documento).")

    dup_chave, dup_sem_chave = encontrar_duplicidades(notas)

    total_dup_chave = sum(n["valor"] for g in dup_chave.values() for n in g)
    total_dup_sem_chave = sum(n["valor"] for g in dup_sem_chave.values() for n in g)

    wb = Workbook()
    ws_resumo = wb.active
    ws_resumo.title = "Resumo"
    ws_resumo["A1"] = "Análise de Notas Fiscais duplicadas (ZLFIB) — Fitted Units"
    ws_resumo["A1"].font = Font(bold=True, size=13)
    ws_resumo.append(["Período analisado", f"{DATA_DE} a {DATA_ATE}"])
    ws_resumo.append(["Filiais", ", ".join(f"{k} ({v})" for k, v in FILIAIS.items())])
    ws_resumo.append([])
    ws_resumo.append(["Linhas de item lidas (todas as filiais)", len(linhas_item)])
    ws_resumo.append(["Notas Fiscais únicas (após agrupar por Nr Documento)", len(notas)])
    ws_resumo.append([])
    ws_resumo.append(["Duplicidade por Chave de acesso (notas de mercadoria — mais confiável)"])
    ws_resumo.append(["Grupos duplicados", len(dup_chave)])
    ws_resumo.append(["Notas envolvidas", sum(len(g) for g in dup_chave.values())])
    ws_resumo.append(["Valor total envolvido", total_dup_chave])
    ws_resumo.append([])
    ws_resumo.append(["Duplicidade por Parceiro+NF+Série+Valor (notas de serviço, sem chave de acesso)"])
    ws_resumo.append(["Grupos duplicados", len(dup_sem_chave)])
    ws_resumo.append(["Notas envolvidas", sum(len(g) for g in dup_sem_chave.values())])
    ws_resumo.append(["Valor total envolvido", total_dup_sem_chave])

    montar_planilha(wb, "Dup. por Chave de Acesso", dup_chave, log)
    montar_planilha(wb, "Dup. sem Chave (serviço)", dup_sem_chave, log)

    nome_base = "Análise Duplicidade NF (ZLFIB).xlsx"
    candidato = PASTA_DESTINO / nome_base
    versao = 2
    while candidato.exists():
        candidato = PASTA_DESTINO / f"Análise Duplicidade NF (ZLFIB)_v{versao}.xlsx"
        versao += 1
    wb.save(candidato)
    log(f"\nArquivo gerado: {candidato}")
    return candidato


if __name__ == "__main__":
    analisar()
