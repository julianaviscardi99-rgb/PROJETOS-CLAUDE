#!/usr/bin/env python3
"""
Automatiza o passo 3 do processo recorrente de Fitted Units Despesas: monta
o arquivo mensal "KSB1 <Mes> <Ciclo> <Ano>.xlsx" a partir do arquivo Actual
do mes anterior, replicando o processo manual real (BASE_KSB1 + Pivot nativo
- ver memory/DECISOES.md, reversao de arquitetura de 2026-08-11):

1. Copia o KSB1 do Actual do mes anterior (ja acumulado) -> nova copia
   versionada (nunca sobrescreve o arquivo original).
2. Decide Gestoriais vs Sem Agrupamento (mesma regra do check_agrupamentos)
   e cola as linhas novas do extrato bruto do mes no fim da aba BASE_KSB1
   (colunas A-R, mapeamento 1:1 confirmado com o extrato).
3. Arrasta (AutoFill, equivalente a "arrastar a formula") as colunas S:AI
   da ultima linha existente para as linhas novas.
4. Da refresh nas Pivot Tables nativas (Pivot_Inter., Pivot_Detalhes) via
   automacao COM do Excel.
5. Salva a copia.

Links externos do BASE_KSB1 (Contas / RHFitted) NAO sao atualizados
(UpdateLinks=0) - usa o mesmo cache que ja estava no arquivo do mes
anterior, para nao misturar "mudanca na base de contas" com "erro na
automacao" ao validar contra um mes ja fechado manualmente.

Depende de Excel instalado (pywin32) - abre uma instancia oculta e isolada
(DispatchEx), nao interfere com o Excel que a usuaria tiver aberto.
"""
import re
import shutil
import sys
import zipfile
from pathlib import Path

import win32com.client
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "_shared"))
from ksb1_core import BU, MESES_PASTA, REDE_BASE, nome_com_versao  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
from check_agrupamentos_ksb1 import (  # noqa: E402
    eh_conta_ignorada,
    encontrar_arquivo,
    linha_e_subtotal,
)

MESES_INGLES = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December",
}

N_COLS_BRUTO = 18  # colunas A-R do BASE_KSB1, 1:1 com o extrato bruto da KSB1
COL_CONTA = 4
COL_VALOR = 17

XL_UP = -4162
XL_FILL_DEFAULT = 0


def _abrev(mes: int) -> str:
    return MESES_PASTA[mes].split(" - ")[1]


def decidir_fonte_e_ler_linhas(mes: int, ano: int, log=print):
    """Decide Gestoriais vs Sem Agrupamento (mesma regra do check_agrupamentos_ksb1)
    e devolve as linhas de detalhe completas (18 colunas, A-R do BASE_KSB1),
    ja filtradas (sem subtotal, sem conta em branco, excluindo contas do
    Check 1 se for usado o Sem Agrupamento)."""
    pasta_mes = REDE_BASE / str(ano) / "00.Extração Base KSB1" / MESES_PASTA[mes]
    arquivo_gest = encontrar_arquivo(pasta_mes, f"KSB1 - {BU['nome']} {mes:02d}.{ano} - Gestoriais")
    arquivo_sem = encontrar_arquivo(pasta_mes, f"KSB1 - {BU['nome']} {mes:02d}.{ano} - Sem Agrupamento")

    def linhas_completas(caminho):
        wb = load_workbook(caminho, data_only=True)
        ws = wb.active
        linhas = []
        for r in range(2, ws.max_row + 1):
            conta = ws.cell(row=r, column=COL_CONTA).value
            if conta in (None, ""):
                continue
            if linha_e_subtotal(ws, r):
                continue
            linha = [ws.cell(row=r, column=c).value for c in range(1, N_COLS_BRUTO + 1)]
            valor = float(ws.cell(row=r, column=COL_VALOR).value or 0)
            linhas.append((str(conta).strip(), valor, linha))
        return linhas

    linhas_gest = linhas_completas(arquivo_gest)
    linhas_sem = linhas_completas(arquivo_sem)

    total_gest = sum(v for _, v, _ in linhas_gest)
    linhas_sem_filtradas = [(c, v, l) for c, v, l in linhas_sem if not eh_conta_ignorada(c)]
    total_sem_filtrado = sum(v for _, v, _ in linhas_sem_filtradas)

    if abs(total_gest - total_sem_filtrado) < 0.01:
        log(f"Gestoriais bate com Sem Agrupamento (total={total_gest:,.2f}) — usando Gestoriais ({arquivo_gest.name}).")
        return [l for _, _, l in linhas_gest], "Gestoriais", arquivo_gest
    else:
        log(
            f"Gestoriais ({total_gest:,.2f}) NÃO bate com Sem Agrupamento filtrado "
            f"({total_sem_filtrado:,.2f}) — usando Sem Agrupamento (excluindo contas do Check 1), "
            f"conforme regra confirmada em 2026-08-11 ({arquivo_sem.name})."
        )
        return [l for _, _, l in linhas_sem_filtradas], "Sem Agrupamento", arquivo_sem


def localizar_ksb1_actual_anterior(mes: int, ano: int) -> Path:
    mes_ant, ano_ant = (mes - 1, ano) if mes > 1 else (12, ano - 1)
    pasta = REDE_BASE / str(ano_ant) / MESES_PASTA[mes_ant] / f"{mes_ant:02d}_{_abrev(mes_ant)}_Actual"
    nome = f"KSB1 {MESES_INGLES[mes_ant]} Actual {ano_ant}.xlsx"
    caminho = pasta / nome
    if not caminho.exists():
        raise FileNotFoundError(f"Não encontrei o KSB1 Actual do mês anterior: {caminho}")
    return caminho


def copiar_para_teste(caminho_origem: Path, pasta_destino: Path, nome_base: str, log=print) -> Path:
    pasta_destino.mkdir(parents=True, exist_ok=True)
    nome_saida = nome_com_versao(pasta_destino, nome_base)
    caminho_saida = pasta_destino / nome_saida
    log(f"Copiando {caminho_origem.name} -> {caminho_saida} ...")
    shutil.copy2(caminho_origem, caminho_saida)
    return caminho_saida


def remover_flag_somente_leitura_recomendada(caminho: Path, log=print):
    """O KSB1 Actual/Flash tem a flag interna 'Somente leitura recomendada'
    (<fileSharing readOnlyRecommended="1"/> em xl/workbook.xml) gravada no
    arquivo. Com DisplayAlerts=False, o Excel abre silenciosamente em modo
    leitura mesmo com IgnoreReadOnlyRecommended=True no Open() (o parâmetro
    não se mostrou confiável via COM), e o Save() vira um no-op sem erro.
    Como esta função só roda na NOSSA cópia de teste (nunca no arquivo
    original), removemos a flag direto do XML antes de abrir no Excel."""
    with zipfile.ZipFile(caminho, "r") as zin:
        itens = {n: zin.read(n) for n in zin.namelist()}
        infos = {n: zin.getinfo(n) for n in zin.namelist()}

    wb_xml = itens["xl/workbook.xml"].decode("utf-8")
    nova_xml, n_removidas = re.subn(r"<fileSharing[^/]*/>", "", wb_xml)
    if n_removidas == 0:
        log("Arquivo não tinha a flag 'Somente leitura recomendada' — nada a remover.")
        return
    itens["xl/workbook.xml"] = nova_xml.encode("utf-8")

    with zipfile.ZipFile(caminho, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, data in itens.items():
            zout.writestr(infos[n], data)
    log(f"Removida a flag 'Somente leitura recomendada' da cópia de teste ({n_removidas} ocorrência(s)).")


def colar_linhas_e_atualizar_pivots(caminho_copia: Path, linhas_novas: list, log=print):
    log("Abrindo Excel (instância isolada, oculta)...")
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    try:
        # IgnoreReadOnlyRecommended=True: o BASE_KSB1 tem a flag interna
        # "Somente leitura recomendada" (fileSharing readOnlyRecommended="1")
        # - sem isso, com DisplayAlerts=False, o Excel abre o arquivo em modo
        # leitura silenciosamente e o Save() vira um no-op sem erro nenhum.
        wb = excel.Workbooks.Open(
            str(caminho_copia), UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True
        )
        if wb.ReadOnly:
            raise RuntimeError(
                "O arquivo abriu em modo somente leitura mesmo com IgnoreReadOnlyRecommended=True "
                "— provavelmente já está aberto/travado por outro processo do Excel."
            )
        ws = wb.Worksheets("BASE_KSB1")
        last_row = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
        n = len(linhas_novas)
        log(f"BASE_KSB1: última linha existente = {last_row}. Colando {n} linha(s) nova(s) (linhas {last_row + 1}-{last_row + n})...")

        destino_dados = ws.Range(ws.Cells(last_row + 1, 1), ws.Cells(last_row + n, N_COLS_BRUTO))
        destino_dados.Value = linhas_novas

        last_row_pos_escrita = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
        if last_row_pos_escrita != last_row + n:
            raise RuntimeError(
                f"Depois de colar, a última linha em memória é {last_row_pos_escrita}, "
                f"esperava {last_row + n} — a escrita não aconteceu como esperado."
            )

        log("Arrastando fórmulas das colunas S:AI para as linhas novas (AutoFill)...")
        origem_formula = ws.Range(ws.Cells(last_row, 19), ws.Cells(last_row, 35))
        destino_formula = ws.Range(ws.Cells(last_row, 19), ws.Cells(last_row + n, 35))
        origem_formula.AutoFill(destino_formula, XL_FILL_DEFAULT)

        log("Recalculando a planilha inteira...")
        excel.CalculateFullRebuild()

        log("Atualizando as Pivot Tables (Pivot_Inter., Pivot_Detalhes)...")
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()

        log("Salvando...")
        wb.Save()
        if not wb.Saved:
            raise RuntimeError("wb.Save() retornou mas wb.Saved ainda é False — o arquivo pode não ter sido gravado.")
        wb.Close(SaveChanges=False)
        log("Concluído.")
    finally:
        excel.Quit()


def gerar_ksb1_mensal(mes: int, ano: int, ciclo: str, pasta_saida: Path, sufixo_nome: str = "", log=print) -> Path:
    linhas_novas, fonte, arquivo_fonte = decidir_fonte_e_ler_linhas(mes, ano, log)

    caminho_origem = localizar_ksb1_actual_anterior(mes, ano)
    log(f"Partindo do Actual do mês anterior: {caminho_origem.name}")

    nome_base = f"KSB1 {MESES_INGLES[mes]} {ciclo} {ano}{sufixo_nome}.xlsx"
    caminho_copia = copiar_para_teste(caminho_origem, pasta_saida, nome_base, log)

    colar_linhas_e_atualizar_pivots(caminho_copia, linhas_novas, log)

    log(f"\nArquivo gerado: {caminho_copia}")
    log(f"Fonte usada: {fonte} ({arquivo_fonte.name}), {len(linhas_novas)} linha(s) colada(s).")
    return caminho_copia


if __name__ == "__main__":
    import sys as _sys

    if len(_sys.argv) != 4:
        print("Uso: python gerar_ksb1_mensal.py <mes> <ano> <Actual|Flash>")
        _sys.exit(1)
    _mes, _ano, _ciclo = int(_sys.argv[1]), int(_sys.argv[2]), _sys.argv[3]
    _pasta_saida = Path(__file__).resolve().parent.parent.parent.parent.parent / "data" / "processed" / "fitted_units_despesas" / "base_ksb1_teste"
    gerar_ksb1_mensal(_mes, _ano, _ciclo, _pasta_saida, sufixo_nome=" - TESTE VALIDAÇÃO")
