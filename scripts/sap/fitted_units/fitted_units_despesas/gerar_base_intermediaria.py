#!/usr/bin/env python3
"""
Passo 4 do processo recorrente (Fitted Units Despesas, Ciclo Actual): monta a
aba "Intermediária" do arquivo "Base Intermediária Fitted <Mes> <Ciclo> <Ano>.xlsx"
a partir do Pivot_Inter. do "arquivo gigante" (BASE_KSB1) já gerado pelo Passo 3.

Logica confirmada com a Juliana em 2026-08-21:
1. Copia a Base Intermediária Actual do mes anterior -> nova copia versionada
   (nunca sobrescreve o arquivo original), igual ao Passo 3.
2. Le o Pivot_Inter. do BASE_KSB1 do mes/ciclo atual (ja gerado pelo Passo 3).
3. Na aba Intermediária, as linhas 2 ate a primeira linha SEM cor de fundo sao
   reservadas pra reclassificacoes/provisoes manuais (só existem de verdade no
   Flash) - no Actual ficam sempre em branco e NUNCA sao tocadas.
4. Apaga toda a area de dados a partir da primeira linha sem cor (rotulos,
   meses, Total Ano, colunas Y:AJ) e cola tudo de novo do Pivot_Inter. - full
   rebuild a cada rodada, em vez de atualizacao incremental celula a celula,
   pra nao acumular erro de alinhamento mes a mes.
5. Unidades ENCERRADAS (ver ontology/fitted_units.json ->
   centros_de_custo_por_unidade, status "encerrada"): a linha continua
   aparecendo na Intermediária (com os rotulos), mas o valor do mes fica em
   branco - o residual retroativo delas nao entra no EBIT. O que ficou de
   fora e' salvo num arquivo separado (colunas A-AA), so com os dados do mes,
   pra usuaria mandar por e-mail pra contabilidade.
6. Arrasta a formula da coluna U (Total Ano) e das colunas Y:AJ (Gestorial II
   ate Conta Geral) da primeira linha pra todas as linhas coladas.
7. Da refresh nas Pivot Tables da aba "Pivot" (wb.RefreshAll()), que nao se
   atualizam sozinhas so porque os dados da Intermediária mudaram.
8. Traz da Base Intermediária FLASH do mesmo mes os valores de Despesas e
   Mao de Obra pro quadro "(+) gain" da aba Pivot (linhas 18/19) - e' o que
   a usuaria hoje faz manualmente (abre o Flash, copia, cola aqui). So roda
   se o ciclo sendo gerado for Actual (nao teria sentido comparar Flash
   contra Flash). Nao e' fatal se o arquivo Flash do mes nao existir.
   Tambem corrige as formulas de Custos (H26/I26) do quadro amarelo
   "Month/Flash/Actual/delta", que ficavam travadas na coluna do ultimo mes
   editado a mao - passam a apontar sempre pra coluna do mes atual.
   Faturamento (linha 25) fica de fora, ainda e' manual.

Ciclo Flash ainda NAO implementado (regra das linhas coloridas e' diferente -
aguardando a usuaria detalhar).

Depende de Excel instalado (pywin32) - abre uma instancia oculta e isolada
(DispatchEx), igual ao gerar_ksb1_mensal.py.
"""
import json
import shutil
import sys
from pathlib import Path

import win32com.client
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "_shared"))
from ksb1_core import MESES_PASTA, REDE_BASE, nome_com_versao, resolver_pasta_ciclo  # noqa: E402

MESES_INGLES = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December",
}

ONTOLOGY_PATH = Path(__file__).resolve().parents[4] / "ontology" / "fitted_units.json"

N_COLS_LABEL = 8         # A-H: rotulos comuns entre Pivot_Inter. e Intermediária
COL_CENTRO_CUSTO = 5     # E, tanto no Pivot_Inter. quanto na Intermediária
COL_TOTAL_ANO = 21       # U
COL_FORMULA_INICIO = 25  # Y (Gestorial II)
COL_FORMULA_FIM = 36     # AJ (Conta Geral)
COL_HISTORICO_FIM = 27   # AA - ate onde vai o arquivo de historico de unidades encerradas

# Aba "Pivot", quadro "(+) gain" (layout proprio, diferente da Intermediária:
# coluna C = Janeiro). Linha 15/16 = Despesas/Mão de Obra do PRÓPRIO ciclo do
# arquivo (formula, ja se atualiza sozinha); linha 18/19 = mesma coisa, mas do
# ciclo OPOSTO (Flash, quando o arquivo e' Actual) - sempre um valor fixo,
# colado a mao todo mes ate agora.
COL_PIVOT_MES_JANEIRO = 3       # C
LINHA_PIVOT_DESPESAS_PROPRIO = 15
LINHA_PIVOT_MAO_DE_OBRA_PROPRIO = 16
LINHA_PIVOT_DESPESAS_COMPARACAO = 18
LINHA_PIVOT_MAO_DE_OBRA_COMPARACAO = 19
LINHA_PIVOT_GRAND_TOTAL = 11

# Quadro amarelo "Month / Flash / Actual / delta" (linhas 24-29): celulas
# fixas (nao mudam de mes pra mes), mas as formulas de Custos (H26/I26)
# precisam apontar pra coluna do MES ATUAL no quadro "(+) gain" acima -
# ficavam travadas na coluna do ultimo mes que alguem editou a mao.
COL_QUADRO_CUSTOS_FLASH = 8    # H
COL_QUADRO_CUSTOS_ACTUAL = 9   # I
LINHA_QUADRO_CUSTOS = 26

# Cambio (L25): so e' alterado no Flash - o Actual sempre puxa o mesmo valor
# de la, celula por celula (nao depende da coluna do mes, e' fixo).
LINHA_CAMBIO = 25
COL_CAMBIO = 12  # L

XL_UP = -4162
XL_TO_LEFT = -4159
XL_FILL_DEFAULT = 0
XL_NONE = -4142


def _normalizar_centro(v) -> str:
    s = str(v).strip()
    try:
        return str(int(float(s)))
    except ValueError:
        return s


def _letra_coluna_mes(mes: int) -> str:
    """Letra da coluna do mes no quadro '(+) gain' da aba Pivot (C=Jan...N=Dec)."""
    return chr(ord("A") - 1 + COL_PIVOT_MES_JANEIRO + mes - 1)


def carregar_centros_encerrados() -> dict:
    """Retorna {centro_de_custo (str): nome_da_unidade} para toda unidade com status 'encerrada'."""
    dados = json.loads(ONTOLOGY_PATH.read_text(encoding="utf-8"))
    grupos = dados["centros_de_custo_por_unidade"]["grupos"]
    mapa = {}
    for nome_grupo, info in grupos.items():
        if info.get("status") == "encerrada":
            for c in info["centros"]:
                mapa[str(c)] = nome_grupo
    return mapa


def localizar_base_ksb1_do_mes(mes: int, ano: int, ciclo: str) -> Path:
    pasta = resolver_pasta_ciclo(REDE_BASE / str(ano) / MESES_PASTA[mes], mes, ciclo)
    nome = f"KSB1 {MESES_INGLES[mes]} {ciclo} {ano}.xlsx"
    caminho = pasta / nome
    if not caminho.exists():
        raise FileNotFoundError(
            f"Não encontrei '{nome}' em {pasta} — rode o Passo 3 (Atualizar Pivot KSB1) "
            "pra esse mês/Ciclo antes de rodar a Base Intermediária."
        )
    return caminho


def localizar_base_intermediaria_flash_do_mes(mes: int, ano: int) -> Path | None:
    """Acha a Base Intermediária Flash do MESMO mês (não do mês anterior) -
    usada só pra trazer os valores de comparação na aba Pivot (linhas 18/19,
    quadro "(+) gain"). Devolve None se não existir (log de aviso, não é
    erro fatal - o resto da Base Intermediária já foi gerado igual)."""
    pasta = resolver_pasta_ciclo(REDE_BASE / str(ano) / MESES_PASTA[mes], mes, "Flash")
    nome = f"Base Intermediária Fitted {MESES_INGLES[mes]} Flash {ano}.xlsx"
    caminho = pasta / nome
    return caminho if caminho.exists() else None


def atualizar_comparacao_flash(excel, wb, mes: int, ano: int, log):
    """Traz da Base Intermediária Flash do mesmo mês os valores de Despesas
    e Mão de Obra pro quadro "(+) gain" da aba Pivot (linhas 18/19) - o que
    a usuária hoje faz manualmente (abre o arquivo Flash, copia, cola aqui).
    Não é fatal se o Flash não existir (avisa e segue sem preencher)."""
    caminho_flash = localizar_base_intermediaria_flash_do_mes(mes, ano)
    if caminho_flash is None:
        log(
            f"AVISO: não encontrei a Base Intermediária Flash de {MESES_INGLES[mes]}/{ano} — "
            "o quadro de comparação Flash x Actual (aba Pivot) não foi atualizado."
        )
        return

    log(f"Trazendo Despesas/Mão de Obra do Flash pro quadro de comparação: {caminho_flash.name}...")
    col = COL_PIVOT_MES_JANEIRO + mes - 1

    wb_flash = excel.Workbooks.Open(str(caminho_flash), ReadOnly=True, UpdateLinks=0)
    try:
        excel.CalculateFullRebuild()
        ws_flash = wb_flash.Worksheets("Pivot")
        valor_despesas = ws_flash.Cells(LINHA_PIVOT_DESPESAS_PROPRIO, col).Value
        valor_mao_de_obra = ws_flash.Cells(LINHA_PIVOT_MAO_DE_OBRA_PROPRIO, col).Value
        valor_cambio = ws_flash.Cells(LINHA_CAMBIO, COL_CAMBIO).Value
    finally:
        wb_flash.Close(SaveChanges=False)

    ws = wb.Worksheets("Pivot")
    ws.Cells(LINHA_PIVOT_DESPESAS_COMPARACAO, col).Value = valor_despesas
    ws.Cells(LINHA_PIVOT_MAO_DE_OBRA_COMPARACAO, col).Value = valor_mao_de_obra
    log(f"  Despesas={valor_despesas:,.2f} | Mão de Obra={valor_mao_de_obra:,.2f}")

    # Cambio (L25) so e' alterado no Flash - o Actual sempre puxa o mesmo
    # valor de la, celula por celula (confirmado com a usuaria em 2026-08-21).
    ws.Cells(LINHA_CAMBIO, COL_CAMBIO).Value = valor_cambio
    log(f"  Câmbio (L25) = {valor_cambio} (puxado do Flash).")

    # Quadro amarelo "Month/Flash/Actual/delta": H26 (Custos, Flash) e I26
    # (Custos, Actual) ficavam travados na coluna do ultimo mes editado a mao
    # - reescreve as duas formulas apontando pra coluna do mes atual (ex: em
    # julho, H26 = "=(I18+I19)/1000", I26 = "=I11/1000"). Faturamento (linha
    # 25) fica de fora de proposito - a usuaria confirmou que ainda vai
    # automatizar isso em outro momento.
    letra_mes = _letra_coluna_mes(mes)
    ws.Cells(LINHA_QUADRO_CUSTOS, COL_QUADRO_CUSTOS_FLASH).Formula = (
        f"=({letra_mes}{LINHA_PIVOT_DESPESAS_COMPARACAO}+{letra_mes}{LINHA_PIVOT_MAO_DE_OBRA_COMPARACAO})/1000"
    )
    ws.Cells(LINHA_QUADRO_CUSTOS, COL_QUADRO_CUSTOS_ACTUAL).Formula = f"={letra_mes}{LINHA_PIVOT_GRAND_TOTAL}/1000"
    log(f"  Fórmulas de Custos (linha {LINHA_QUADRO_CUSTOS}) apontando pra coluna {letra_mes} ({MESES_INGLES[mes]}).")


def localizar_base_intermediaria_mes_anterior(mes: int, ano: int) -> Path:
    mes_ant, ano_ant = (mes - 1, ano) if mes > 1 else (12, ano - 1)
    pasta = resolver_pasta_ciclo(REDE_BASE / str(ano_ant) / MESES_PASTA[mes_ant], mes_ant, "Actual")
    nome = f"Base Intermediária Fitted {MESES_INGLES[mes_ant]} Actual {ano_ant}.xlsx"
    caminho = pasta / nome
    if not caminho.exists():
        raise FileNotFoundError(f"Não encontrei a Base Intermediária Actual do mês anterior: {caminho}")
    return caminho


def ler_pivot_inter(caminho_base_ksb1: Path, excel, log):
    log(f"Lendo Pivot_Inter. de {caminho_base_ksb1.name}...")
    wb = excel.Workbooks.Open(str(caminho_base_ksb1), ReadOnly=True, UpdateLinks=0)
    try:
        # Sem isso, algumas celulas do Pivot_Inter podem ser gravadas como
        # "#N/A" no arquivo final (o calculo assincrono ligado ao link externo
        # da base de contas nao tinha terminado antes do Save) - achado
        # testando ao vivo em 2026-08-21.
        excel.CalculateFullRebuild()
        excel.CalculateUntilAsyncQueriesDone()
        ws = wb.Worksheets("Pivot_Inter.")
        last_row = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
        last_col = ws.Cells(4, ws.Columns.Count).End(XL_TO_LEFT).Column
        valores = ws.Range(ws.Cells(4, 1), ws.Cells(last_row, last_col)).Value
    finally:
        wb.Close(SaveChanges=False)

    cabecalho = list(valores[0])
    n_meses = 0
    for v in cabecalho[N_COLS_LABEL:]:
        if isinstance(v, (int, float)):
            n_meses += 1
        else:
            break
    if n_meses == 0:
        raise RuntimeError(f"Não consegui identificar as colunas de mês no cabeçalho do Pivot_Inter.: {cabecalho}")

    linhas = [list(l) for l in valores[1:] if l[0] != "Grand Total"]
    log(f"  {len(linhas)} linha(s) de dados, {n_meses} mês(es) de valor.")
    return cabecalho, linhas, n_meses


def encontrar_primeira_linha_sem_cor(ws) -> int:
    r = 2
    while r < 5000:
        interior = ws.Cells(r, 1).Interior
        if interior.Pattern == XL_NONE:
            return r
        r += 1
    raise RuntimeError("Não encontrei uma linha sem cor nas primeiras 5000 linhas da Intermediária — layout inesperado.")


def gerar_historico_unidades_encerradas(mes, ano, ciclo, cabecalho_historico, linhas_historico, pasta_saida, log) -> Path | None:
    if not linhas_historico:
        log("Nenhuma unidade encerrada com valor no mês — não gerei arquivo de histórico.")
        return None

    wb = Workbook()
    ws = wb.active
    ws.title = "Histórico"
    ws.append([f"Resíduo de unidades encerradas — {MESES_INGLES[mes]}/{ano} ({ciclo})"])
    ws.append([f"Não entram na Base Intermediária oficial — ficam em branco lá. Gerado para envio à contabilidade."])
    ws.append([])
    ws.append(cabecalho_historico)
    for linha in linhas_historico:
        ws.append(linha)

    nome_base = f"Histórico Unidades Encerradas - {MESES_INGLES[mes]} {ciclo} {ano}.xlsx"
    nome_saida = nome_com_versao(pasta_saida, nome_base)
    caminho_saida = pasta_saida / nome_saida
    wb.save(caminho_saida)
    log(f"Arquivo de histórico de unidades encerradas gerado: {caminho_saida}")
    return caminho_saida


def atualizar_base_intermediaria(mes: int, ano: int, ciclo: str, pasta_saida: Path, sufixo_nome: str = "", log=print):
    if ciclo != "Actual":
        raise NotImplementedError("Só o Ciclo Actual está implementado por enquanto — Flash ainda não.")

    centros_encerrados = carregar_centros_encerrados()

    caminho_base_ksb1 = localizar_base_ksb1_do_mes(mes, ano, ciclo)
    caminho_origem_inter = localizar_base_intermediaria_mes_anterior(mes, ano)

    log(f"Partindo da Base Intermediária Actual do mês anterior: {caminho_origem_inter.name}")
    pasta_saida.mkdir(parents=True, exist_ok=True)
    nome_base = f"Base Intermediária Fitted {MESES_INGLES[mes]} {ciclo} {ano}{sufixo_nome}.xlsx"
    nome_saida = nome_com_versao(pasta_saida, nome_base)
    caminho_saida = pasta_saida / nome_saida
    log(f"Copiando {caminho_origem_inter.name} -> {caminho_saida.name} ...")
    shutil.copy2(caminho_origem_inter, caminho_saida)

    log("Abrindo Excel (instância isolada, oculta)...")
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    try:
        cabecalho_pivot, linhas_pivot, n_meses = ler_pivot_inter(caminho_base_ksb1, excel, log)

        wb = excel.Workbooks.Open(str(caminho_saida), UpdateLinks=0, ReadOnly=False)
        if wb.ReadOnly:
            raise RuntimeError(
                "A cópia abriu em modo somente leitura — provavelmente já está aberta por outro processo."
            )
        ws = wb.Worksheets("Intermediária")

        primeira_sem_cor = encontrar_primeira_linha_sem_cor(ws)
        last_row_antiga = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
        n_linhas_novas = len(linhas_pivot)
        if n_linhas_novas == 0:
            raise RuntimeError("Pivot_Inter. não trouxe nenhuma linha de dados — abortando sem mexer no arquivo.")
        last_row_final = primeira_sem_cor + n_linhas_novas - 1
        last_row_limpar = max(last_row_antiga, last_row_final)

        log(
            f"Primeira linha sem cor: {primeira_sem_cor}. Linhas atuais até {last_row_antiga}. "
            f"{n_linhas_novas} linha(s) nova(s) do Pivot_Inter."
        )

        formula_total_ano = ws.Cells(primeira_sem_cor, COL_TOTAL_ANO).Formula
        formulas_extra = [
            ws.Cells(primeira_sem_cor, c).Formula for c in range(COL_FORMULA_INICIO, COL_FORMULA_FIM + 1)
        ]
        if not formula_total_ano or not all(formulas_extra):
            raise RuntimeError(
                f"Não consegui capturar as fórmulas de referência (Total Ano / colunas Y:AJ) da linha "
                f"{primeira_sem_cor} antes de apagar — abortando pra não perder o padrão de fórmula."
            )

        log("Apagando a área de dados antiga...")
        ws.Range(ws.Cells(primeira_sem_cor, 1), ws.Cells(last_row_limpar, COL_FORMULA_FIM)).ClearContents()

        log(f"Colando {n_linhas_novas} linha(s) do Pivot_Inter. (todas, unidades encerradas incluídas por enquanto)...")
        # Colar linha por linha, NUNCA o bloco inteiro de uma vez: escrever um
        # array grande (testado com 601 linhas) via Range.Value em uma unica
        # chamada COM corrompe aleatoriamente algumas celulas em erro #N/A
        # (bug de marshalling do pywin32/COM, nao tem relacao com o conteudo -
        # achado e confirmado testando ao vivo em 2026-08-21: 166 erros colando
        # tudo de uma vez, 25 em blocos de 50, ZERO colando linha por linha).
        linhas_para_colar = [linha[: N_COLS_LABEL + n_meses] for linha in linhas_pivot]
        for i, linha in enumerate(linhas_para_colar):
            r = primeira_sem_cor + i
            ws.Range(ws.Cells(r, 1), ws.Cells(r, N_COLS_LABEL + n_meses)).Value = [linha]

        log("Verificando se alguma célula foi gravada como erro (#N/A) durante a colagem...")
        conferencia = ws.Range(
            ws.Cells(primeira_sem_cor, 1), ws.Cells(last_row_final, N_COLS_LABEL + n_meses)
        ).Value
        celulas_com_erro = sum(
            1 for linha in conferencia for v in linha if isinstance(v, int) and v < 0
        )
        if celulas_com_erro:
            raise RuntimeError(
                f"{celulas_com_erro} célula(s) gravada(s) como erro (#N/A) mesmo colando linha por linha — "
                "abortando sem salvar. Rode de novo; se persistir, é um problema novo, não o bug conhecido."
            )

        log("Arrastando a fórmula da coluna U (Total Ano)...")
        ws.Cells(primeira_sem_cor, COL_TOTAL_ANO).Formula = formula_total_ano
        ws.Cells(primeira_sem_cor, COL_TOTAL_ANO).AutoFill(
            ws.Range(ws.Cells(primeira_sem_cor, COL_TOTAL_ANO), ws.Cells(last_row_final, COL_TOTAL_ANO)),
            XL_FILL_DEFAULT,
        )

        log("Arrastando as fórmulas das colunas Y:AJ...")
        for i, c in enumerate(range(COL_FORMULA_INICIO, COL_FORMULA_FIM + 1)):
            ws.Cells(primeira_sem_cor, c).Formula = formulas_extra[i]
        ws.Range(
            ws.Cells(primeira_sem_cor, COL_FORMULA_INICIO), ws.Cells(primeira_sem_cor, COL_FORMULA_FIM)
        ).AutoFill(
            ws.Range(
                ws.Cells(primeira_sem_cor, COL_FORMULA_INICIO), ws.Cells(last_row_final, COL_FORMULA_FIM)
            ),
            XL_FILL_DEFAULT,
        )

        log("Recalculando a planilha inteira...")
        excel.CalculateFullRebuild()
        excel.CalculateUntilAsyncQueriesDone()

        log("Identificando linhas de unidades encerradas pra separar o histórico e zerar o valor oficial...")
        cabecalho_historico = cabecalho_pivot[:COL_HISTORICO_FIM]
        linhas_historico = []
        idx_col_mes_ini = N_COLS_LABEL + 1
        idx_col_mes_fim = N_COLS_LABEL + n_meses
        for r in range(primeira_sem_cor, last_row_final + 1):
            centro = _normalizar_centro(ws.Cells(r, COL_CENTRO_CUSTO).Value)
            if centro not in centros_encerrados:
                continue
            valores_mes = [ws.Cells(r, c).Value for c in range(idx_col_mes_ini, idx_col_mes_fim + 1)]
            if not any(v not in (None, 0, "") for v in valores_mes):
                continue  # unidade encerrada sem valor nenhum no mes - nada a registrar
            linha_completa = [ws.Cells(r, c).Value for c in range(1, COL_HISTORICO_FIM + 1)]
            linhas_historico.append(linha_completa)
            for c in range(idx_col_mes_ini, idx_col_mes_fim + 1):
                ws.Cells(r, c).ClearContents()

        log(f"  {len(linhas_historico)} linha(s) de unidade encerrada com valor — zeradas na oficial, separadas pro histórico.")

        log("Atualizando as Pivot Tables da aba 'Pivot'...")
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()

        if ciclo == "Actual":
            atualizar_comparacao_flash(excel, wb, mes, ano, log)

        log("Recalculando de novo (Total Ano das linhas zeradas)...")
        excel.CalculateFullRebuild()
        excel.CalculateUntilAsyncQueriesDone()

        log("Salvando...")
        wb.Save()
        if not wb.Saved:
            raise RuntimeError("wb.Save() retornou mas wb.Saved ainda é False — o arquivo pode não ter sido gravado.")
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()

    caminho_historico = gerar_historico_unidades_encerradas(
        mes, ano, ciclo, cabecalho_historico, linhas_historico, pasta_saida, log
    )

    log(f"\nBase Intermediária atualizada: {caminho_saida}")
    return caminho_saida, caminho_historico


if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Uso: python gerar_base_intermediaria.py <mes> <ano> <Actual>")
        sys.exit(1)
    _mes, _ano, _ciclo = int(sys.argv[1]), int(sys.argv[2]), sys.argv[3]
    _pasta_saida = (
        Path(__file__).resolve().parent.parent.parent.parent.parent
        / "data" / "processed" / "fitted_units_despesas" / "base_intermediaria_teste"
    )
    atualizar_base_intermediaria(_mes, _ano, _ciclo, _pasta_saida, sufixo_nome=" - TESTE VALIDAÇÃO")
