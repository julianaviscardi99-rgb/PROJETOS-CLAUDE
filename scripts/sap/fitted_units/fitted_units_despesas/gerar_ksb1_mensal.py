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

import pywintypes
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "_shared"))
from ksb1_core import (  # noqa: E402
    BU,
    MESES_PASTA,
    REDE_BASE,
    abrir_excel_isolado,
    com_retry,
    encontrar_arquivo_mais_recente,
    localizar_extracao_ksb1,
    nome_com_versao,
    resolver_pasta_ciclo,
)

sys.path.insert(0, str(Path(__file__).resolve().parent))
from check_agrupamentos_ksb1 import eh_conta_ignorada, linha_e_subtotal  # noqa: E402

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
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105


def _abrev(mes: int) -> str:
    return MESES_PASTA[mes].split(" - ")[1]


def decidir_fonte_e_ler_linhas(mes: int, ano: int, ciclo: str, log=print):
    """Decide Gestoriais vs Sem Agrupamento (mesma regra do check_agrupamentos_ksb1)
    e devolve as linhas de detalhe completas (18 colunas, A-R do BASE_KSB1),
    ja filtradas (sem subtotal, sem conta em branco, excluindo contas do
    Check 1 se for usado o Sem Agrupamento). As extracoes brutas usadas sao
    sempre do Ciclo pedido (ver ksb1_core.encontrar_arquivo_ksb1) - nao mais
    "a mais recente por data de modificacao", que podia pegar por engano a
    extracao de outro Ciclo do mesmo mes."""
    pasta_mes = REDE_BASE / str(ano) / "00.Extração Base KSB1" / MESES_PASTA[mes]
    arquivo_gest = localizar_extracao_ksb1(pasta_mes, BU["nome"], mes, ano, "Gestoriais", ciclo)
    arquivo_sem = localizar_extracao_ksb1(pasta_mes, BU["nome"], mes, ano, "Sem Agrupamento", ciclo)

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
    linhas_sem_filtradas = [(c, v, l) for c, v, l in linhas_sem if not eh_conta_ignorada(c, mes, ano)]
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
    pasta_mes = REDE_BASE / str(ano_ant) / MESES_PASTA[mes_ant]
    pasta = resolver_pasta_ciclo(pasta_mes, mes_ant, "Actual")
    nome = f"KSB1 {MESES_INGLES[mes_ant]} Actual {ano_ant}.xlsx"
    caminho = encontrar_arquivo_mais_recente(pasta, nome)
    if caminho is None:
        raise FileNotFoundError(f"Não encontrei o KSB1 Actual do mês anterior: {pasta / nome}")
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


def colar_linhas_e_atualizar_pivots(caminho_copia: Path, linhas_novas: list, log=print, pid_callback=None):
    excel = abrir_excel_isolado(log, pid_callback)
    try:
        # IgnoreReadOnlyRecommended=True: o BASE_KSB1 tem a flag interna
        # "Somente leitura recomendada" (fileSharing readOnlyRecommended="1")
        # - sem isso, com DisplayAlerts=False, o Excel abre o arquivo em modo
        # leitura silenciosamente e o Save() vira um no-op sem erro nenhum.
        wb = com_retry(
            excel.Workbooks.Open,
            str(caminho_copia), UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True,
            log=log,
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

        # Calculo manual durante a colagem/AutoFill: com calculo automatico
        # (padrao do Excel), CADA uma das milhares de escritas linha-a-linha
        # abaixo dispara recalculo do arquivo inteiro (formulas + Pivot) -
        # e' o principal motivo da operacao levar 10+ minutos num mes com
        # muitas linhas (achado real, 2026-09-01, usuaria reportou "mais de
        # 12 minutos"). Nao afeta a protecao contra o bug de corrupcao
        # #N/A (que e' sobre GRANULARIDADE da escrita, nao sobre o modo de
        # calculo) - o recalculo completo continua acontecendo, so' que UMA
        # vez so' (CalculateFullRebuild, abaixo) em vez de milhares de vezes.
        # Restaurado pra automatico antes de salvar, pra nao mudar o modo de
        # calculo padrao do arquivo pra quem abrir depois.
        com_retry(setattr, excel, "Calculation", XL_CALCULATION_MANUAL, log=log)
        com_retry(setattr, excel, "ScreenUpdating", False, log=log)

        # Colar linha por linha, NUNCA o bloco inteiro de uma vez: escrever um
        # array grande via Range.Value em uma unica chamada COM pode corromper
        # aleatoriamente algumas celulas em erro #N/A (bug de marshalling do
        # pywin32/COM, sem relacao com o conteudo - achado e confirmado
        # testando gerar_base_intermediaria.py ao vivo em 2026-08-21: 166
        # erros colando tudo de uma vez, 25 em blocos de 50, ZERO linha por
        # linha). Nenhum arquivo de producao foi afetado ate agora (confirmado
        # que este script nunca rodou de verdade contra a rede), mas a mesma
        # proteção foi aplicada aqui preventivamente.
        for i, linha in enumerate(linhas_novas):
            r = last_row + 1 + i
            ws.Range(ws.Cells(r, 1), ws.Cells(r, N_COLS_BRUTO)).Value = [linha]

        last_row_pos_escrita = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
        if last_row_pos_escrita != last_row + n:
            raise RuntimeError(
                f"Depois de colar, a última linha em memória é {last_row_pos_escrita}, "
                f"esperava {last_row + n} — a escrita não aconteceu como esperado."
            )

        log("Verificando se alguma célula foi gravada como erro (#N/A) durante a colagem...")
        conferencia = ws.Range(ws.Cells(last_row + 1, 1), ws.Cells(last_row + n, N_COLS_BRUTO)).Value
        celulas_com_erro = sum(
            1 for linha in conferencia for v in linha if isinstance(v, int) and v < 0
        )
        if celulas_com_erro:
            raise RuntimeError(
                f"{celulas_com_erro} célula(s) gravada(s) como erro (#N/A) mesmo colando linha por linha — "
                "abortando sem salvar."
            )

        log("Arrastando fórmulas das colunas S:AI para as linhas novas (AutoFill)...")
        origem_formula = ws.Range(ws.Cells(last_row, 19), ws.Cells(last_row, 35))
        destino_formula = ws.Range(ws.Cells(last_row, 19), ws.Cells(last_row + n, 35))
        com_retry(origem_formula.AutoFill, destino_formula, XL_FILL_DEFAULT, log=log)

        log("Recalculando a planilha inteira...")
        com_retry(excel.CalculateFullRebuild, log=log)

        # Restaura o modo padrao (automatico) antes de salvar/mexer nas
        # Pivots - o CalculateFullRebuild acima ja forcou o recalculo
        # completo uma vez, entao nao perde nada; so' evita salvar o arquivo
        # com o app deixado em modo manual.
        com_retry(setattr, excel, "Calculation", XL_CALCULATION_AUTOMATIC, log=log)
        com_retry(setattr, excel, "ScreenUpdating", True, log=log)

        log("Atualizando as Pivot Tables (Pivot_Inter., Pivot_Detalhes)...")
        com_retry(wb.RefreshAll, log=log)
        com_retry(excel.CalculateUntilAsyncQueriesDone, log=log)

        log("Salvando...")
        com_retry(wb.Save, log=log)
        if not wb.Saved:
            raise RuntimeError("wb.Save() retornou mas wb.Saved ainda é False — o arquivo pode não ter sido gravado.")
        com_retry(wb.Close, SaveChanges=False, log=log)
        log("Concluído.")
    finally:
        # Retentativa tambem no Quit - se o Excel "ocupado" derrubar essa
        # chamada, ela e' engolida sem retentar, deixa o processo orfao
        # rodando pra sempre (achado ao vivo, 2026-09-01: sobrou um
        # EXCEL.EXE ocioso depois de uma rodada bem-sucedida deste script).
        try:
            com_retry(excel.Quit, log=log)
        except pywintypes.com_error:
            log("AVISO: não consegui confirmar o encerramento do Excel (pode ter ficado um processo órfão).")


def gerar_ksb1_mensal(
    mes: int, ano: int, ciclo: str, pasta_saida: Path, sufixo_nome: str = "", log=print, pid_callback=None
) -> Path:
    linhas_novas, fonte, arquivo_fonte = decidir_fonte_e_ler_linhas(mes, ano, ciclo, log)

    caminho_origem = localizar_ksb1_actual_anterior(mes, ano)
    log(f"Partindo do Actual do mês anterior: {caminho_origem.name}")

    nome_base = f"KSB1 {MESES_INGLES[mes]} {ciclo} {ano}{sufixo_nome}.xlsx"
    caminho_copia = copiar_para_teste(caminho_origem, pasta_saida, nome_base, log)
    remover_flag_somente_leitura_recomendada(caminho_copia, log)

    colar_linhas_e_atualizar_pivots(caminho_copia, linhas_novas, log, pid_callback)

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
