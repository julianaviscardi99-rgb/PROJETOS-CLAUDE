#!/usr/bin/env python3
"""
Passo 3 (Provisões, so' Flash) e Passo 4 (Base Intermediária) do processo
recorrente (Fitted Units Despesas). Passo 4 monta a aba "Intermediária" do
arquivo "Base Intermediária Fitted <Mes> <Ciclo> <Ano>.xlsx" a partir do
Pivot_Inter. do "arquivo gigante" (BASE_KSB1) já gerado pelo Passo 4 (botão
"Atualizar Pivot KSB1", gerar_ksb1_mensal.py).

Logica confirmada com a Juliana em 2026-08-21 (renumerado nesta mesma sessao:
o antigo Passo 3 - Atualizar Pivot KSB1 + Finalização - virou Passo 4; o
preenchimento de provisoes do Flash ganhou seu proprio Passo 3):
1. Copia a Base Intermediária Actual do mes anterior -> nova copia versionada
   (nunca sobrescreve o arquivo original), igual ao Passo 4.
2. Le o Pivot_Inter. do BASE_KSB1 do mes/ciclo atual (ja gerado pelo Passo 4).
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
8. Traz os valores de Despesas e Mao de Obra pro quadro "(+) gain" da aba
   Pivot (linhas 18/19) - e' o que a usuaria hoje faz manualmente. Se o
   ciclo sendo gerado for Actual, a fonte e' a Base Intermediária FLASH do
   mesmo mes (linhas 15/16 da aba Pivot dela). Se for Flash, nao faria
   sentido comparar Flash com Flash - a fonte passa a ser o Forecast mais
   recente (R<mes>, com fallback pro Forecast do mes anterior quando R<mes>
   nao existe - R1/R8/R12 nunca sao feitos - avisando qual foi usado; ver
   atualizar_comparacao_forecast). Nao e' fatal se a fonte nao existir.
   Tambem corrige as formulas de Custos (H26/I26) do quadro amarelo
   "Month/Flash/Actual/delta", que ficavam travadas na coluna do ultimo mes
   editado a mao - passam a apontar sempre pra coluna do mes atual - e, so
   no caso Flash, os rotulos de texto herdados do template (linha 15/18 e
   cabecalho da linha 24) que precisam virar "Flash"/"Forecast".
   Faturamento (linha 25) fica de fora, ainda e' manual.

Ciclo Flash (confirmado com a usuaria em 2026-08-21): antes do passo 3 acima,
preenche as linhas coloridas (provisoes/reclassificacoes) a partir do "Fast
Provisao" do mes (pasta "Provisões e Reclassificações" dentro do Flash do
mes, sempre a versao mais alta) - cada linha da aba "Ficha de Solicitação"
(a partir da linha 13: Conta=col C, Centro=col D, Valor=col H) vira uma linha
colorida na Intermediária (linha 2, 3, 4...), arrastando pra cada uma a
formula VLOOKUP (colunas A, B, D, F, G) que ja existe na ultima linha
colorida do arquivo. As provisoes SEMPRE comecam do zero a cada mes (a
contabilidade estorna a provisao do mes anterior, entao nao ha nada
acumulado pra herdar - o arquivo de origem, sempre o Actual do mes anterior,
ja vem com essas linhas em branco). Se as provisoes nao couberem nas linhas
amarelas existentes, insere linhas amarelas novas automaticamente logo antes
da primeira linha verde (nunca mexe no conteudo das verdes/roxas - ver
inserir_linhas_amarelas_novas, 2026-08-22). Linhas verdes (reclassificacoes)
estao DEPRECADAS por decisao da usuaria (2026-08-22) - reclassificacoes ja
acontecem direto no SAP antes do fechamento. O quadro de comparacao (passo 8)
roda normalmente pro Flash
tambem, mas comparando contra o Forecast em vez de outro Flash - ver
atualizar_comparacao_forecast.

Depende de Excel instalado (pywin32) - abre uma instancia oculta e isolada
(DispatchEx), igual ao gerar_ksb1_mensal.py.
"""
import json
import re
import shutil
import sys
import time

import pywintypes
from pathlib import Path

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "_shared"))
from ksb1_core import (  # noqa: E402
    MESES_PASTA,
    REDE_BASE,
    abrir_excel_isolado,
    com_retry,
    encontrar_arquivo_mais_recente,
    nome_com_versao,
    resolver_pasta_ciclo,
)

MESES_INGLES = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December",
}

ONTOLOGY_PATH = Path(__file__).resolve().parents[4] / "ontology" / "fitted_units.json"

N_COLS_LABEL = 8         # A-H: rotulos comuns entre Pivot_Inter. e Intermediária
COL_CENTRO_CUSTO = 5     # E, tanto no Pivot_Inter. quanto na Intermediária
COL_MINI_FABRICA = 6     # F - codigo da unidade, TEXTO com zero a esquerda ("0499")
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

# Cabecalho do quadro amarelo (linha 24, mesmas colunas H/I de LINHA_QUADRO_
# CUSTOS) - rotulo texto que precisa virar "Forecast"/"Flash" quando o
# arquivo sendo gerado e' Flash (herda "Flash"/"Actual" do template copiado,
# que e' sempre o Actual do mes anterior).
LINHA_QUADRO_HEADER = 24

# Aba "Resumo Resultado Ano" do arquivo de Forecast (P&L Fitted Units_
# Forecast_<Mes>_<AA>_.xlsx - versao com "_" no final, sem formula, pedido
# explicito da usuaria em 2026-08-22 pra nao perder a base se alguem mexer
# por engano). Estrutura confirmada inspecionando ao vivo o arquivo real de
# julho/2026 (R7): coluna D = Janeiro, avancando 1 coluna por mes. "Total
# Costs" (linha 38) sempre bateu exatamente com Variable Cost (19) + Fixed
# Cost (30) nos 12 meses testados - mas o check roda em runtime mesmo assim,
# a usuaria pediu pra sempre confirmar antes de confiar.
ABA_FORECAST_RESUMO = "Resumo Resultado Ano"
COL_FORECAST_JANEIRO = 4        # D
LINHA_FORECAST_VARIABLE_COST = 19
LINHA_FORECAST_LABOUR_VARIAVEL = 20
LINHA_FORECAST_FIXED_COST = 30
LINHA_FORECAST_LABOUR_FIXO = 31
LINHA_FORECAST_TOTAL_COSTS = 38

# Budget/MP - usado so' no fechamento de Janeiro (R1 nunca e' feito, ver
# localizar_forecast_para_comparacao). Fica numa area de rede totalmente
# separada da dos outros ciclos (Management Plan, nao Resultados Fitted),
# preparado ANTES do ano comecar mas arquivado sob o ano que ele cobre -
# confirmado pela usuaria em 2026-08-22: fechamento de Janeiro/2026 usa a
# pasta "MP 2026" (mesmo ano, nao o anterior). Mesma aba/linhas do Forecast
# (Resumo Resultado Ano, 19/20/30/31/38) - confirmado inspecionando ao vivo.
PASTA_BUDGET_BASE = Path(r"\\FSS024-01BR.group.pirelli.com\GFU_DAC\Management Plan")

XL_UP = -4162
XL_TO_LEFT = -4159
XL_FILL_DEFAULT = 0
XL_NONE = -4142
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105


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
    caminho = encontrar_arquivo_mais_recente(pasta, nome)
    if caminho is None:
        raise FileNotFoundError(
            f"Não encontrei '{nome}' em {pasta} — rode o Passo 4 (Atualizar Pivot KSB1) "
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
    return encontrar_arquivo_mais_recente(pasta, nome)


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
        com_retry(excel.CalculateFullRebuild, log=log)
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


def localizar_arquivo_forecast(mes: int, ano: int) -> Path | None:
    """Acha o arquivo de Forecast (R<mes>) do mes, versao sem formula (com
    "_" no final do nome - recomendado pela usuaria em 2026-08-22, protege
    contra alteracao acidental). Devolve None se a pasta ou o arquivo nao
    existir (R1, R8 e R12 nunca sao feitos; R9-R11 do ano corrente podem
    ainda nao existir se o mes nao chegou)."""
    pasta = resolver_pasta_ciclo(REDE_BASE / str(ano) / MESES_PASTA[mes], mes, "Forecast")
    nome = f"{mes:02d}_P&L Fitted Units_Forecast_{MESES_INGLES[mes]}_{str(ano)[-2:]}_.xlsx"
    caminho = pasta / nome
    return caminho if caminho.exists() else None


def localizar_arquivo_budget(ano: int) -> Path | None:
    """Acha o arquivo de Budget/MP do ano (fonte de comparação so' pro
    fechamento de Janeiro, que nao tem R1). Devolve None se a pasta ou o
    arquivo nao existir."""
    pasta = PASTA_BUDGET_BASE / f"MP {ano}"
    nome = f"P&L Fitted Units_Budget{str(ano)[-2:]}_.xlsx"
    caminho = pasta / nome
    return caminho if caminho.exists() else None


def localizar_forecast_para_comparacao(mes: int, ano: int, log):
    """Acha a fonte de comparacao pro fechamento do Flash do mes:
    - Janeiro (mes=1): usa direto o Budget/MP do ano (R1 nunca e' feito) -
      ver localizar_arquivo_budget.
    - Outros meses: tenta R<mes> (o Forecast feito com o mes em aberto,
      cobrindo Janeiro-Dezembro); se nao existir (caso conhecido de
      Agosto/R8 e Dezembro/R12, que a Pirelli nao produz), cai pro Forecast
      do mes anterior (R<mes-1>), que tambem projeta o mes atual.

    Devolve (caminho, aviso) - aviso e' None quando achou a fonte esperada
    sem ressalva (R<mes> direto), ou uma string explicando o que foi usado
    (Budget/MP, ou fallback de Forecast) pra virar popup na GUI. Devolve
    None (sem tupla) se nao achou nenhuma fonte - nao e' erro fatal, o
    quadro de comparacao so fica sem preencher nesta rodada, mesmo
    tratamento ja usado quando o Flash do mes nao existe pro caso do
    Actual."""
    if mes == 1:
        caminho = localizar_arquivo_budget(ano)
        if caminho is None:
            log(
                f"AVISO: não encontrei o Budget/MP de {ano} "
                f"({PASTA_BUDGET_BASE / f'MP {ano}'}) — o quadro de comparação não foi preenchido."
            )
            return None
        aviso = (
            f"Fechamento de {MESES_INGLES[mes]}/{ano}: não existe Forecast R1 — usei o "
            f"Budget/MP {ano} como comparação."
        )
        return caminho, aviso

    caminho = localizar_arquivo_forecast(mes, ano)
    if caminho is not None:
        return caminho, None

    mes_ant = mes - 1
    caminho_ant = localizar_arquivo_forecast(mes_ant, ano)
    if caminho_ant is None:
        log(
            f"AVISO: não encontrei o Forecast de {MESES_INGLES[mes]}/{ano} nem o de "
            f"{MESES_INGLES[mes_ant]}/{ano} — o quadro de comparação Forecast não foi preenchido."
        )
        return None

    aviso = (
        f"Fechamento de {MESES_INGLES[mes]}/{ano}: não existe Forecast R{mes} — usei o Forecast de "
        f"{MESES_INGLES[mes_ant]} (R{mes_ant}) como comparação."
    )
    return caminho_ant, aviso


def ler_forecast_despesas_mao_de_obra(excel, caminho_forecast: Path, mes_coluna: int, log):
    """Le a aba 'Resumo Resultado Ano' do arquivo de Forecast, na coluna do
    MES QUE ESTA SENDO FECHADO (nao necessariamente o mes do proprio R -
    no fallback, ex: fechando Agosto com o R7, a coluna lida e' Agosto
    mesmo, que dentro do R7 e' a coluna de Forecast pra frente).

    Devolve (despesas, mao_de_obra) em BRL absoluto positivo (o arquivo de
    Forecast guarda em '000 BRL, custo como numero negativo - precisa
    multiplicar por 1000 e inverter o sinal pra bater com o padrao das
    linhas 18/19 da aba Pivot, que sao sempre positivas e em BRL cheio).

    Antes de confiar no valor, confere se Total Costs bate com Variable
    Cost + Fixed Cost (pedido explicito da usuaria - avisar sempre que nao
    bater, nao so na implementacao)."""
    col = COL_FORECAST_JANEIRO + mes_coluna - 1
    wb = excel.Workbooks.Open(str(caminho_forecast), ReadOnly=True, UpdateLinks=0)
    try:
        com_retry(excel.CalculateFullRebuild, log=log)
        ws = wb.Worksheets(ABA_FORECAST_RESUMO)
        variable_cost = ws.Cells(LINHA_FORECAST_VARIABLE_COST, col).Value
        labour_variavel = ws.Cells(LINHA_FORECAST_LABOUR_VARIAVEL, col).Value
        fixed_cost = ws.Cells(LINHA_FORECAST_FIXED_COST, col).Value
        labour_fixo = ws.Cells(LINHA_FORECAST_LABOUR_FIXO, col).Value
        total_costs = ws.Cells(LINHA_FORECAST_TOTAL_COSTS, col).Value
    finally:
        wb.Close(SaveChanges=False)

    soma_var_fixo = variable_cost + fixed_cost
    if abs(soma_var_fixo - total_costs) > 0.01:
        log(
            f"AVISO: no Forecast ({caminho_forecast.name}), Total Costs ({total_costs:,.2f}) não bate "
            f"com Variable Cost + Fixed Cost ({soma_var_fixo:,.2f}) — confira o arquivo antes de confiar "
            "no quadro de comparação."
        )

    mao_de_obra = -(labour_variavel + labour_fixo) * 1000
    despesas = -total_costs * 1000 - mao_de_obra
    return despesas, mao_de_obra


def atualizar_comparacao_forecast(excel, wb, mes: int, ano: int, log) -> str | None:
    """Equivalente ao atualizar_comparacao_flash, mas pro ciclo Flash: nao
    faria sentido comparar Flash com Flash, entao traz do Forecast mais
    recente (ver localizar_forecast_para_comparacao) os valores de Despesas/
    Mão de Obra pro quadro "(+) gain" (linhas 18/19) e recalcula as fórmulas
    de Custos (H26/I26) - mesma mecânica já usada pro Actual, que é genérica
    o bastante pra servir os dois casos sem mudança.

    Também corrige os rótulos de texto que o arquivo herda do template
    copiado (sempre o Actual do mês anterior): linha 15/coluna A "Actual" ->
    "Flash" (é o próprio ciclo do arquivo agora), linha 18/coluna A "Flash"
    -> "Forecast", cabeçalho do quadro amarelo (linha 24, colunas H/I)
    "Flash"/"Actual" -> "Forecast"/"Flash". Confirmado em 2026-08-22
    inspecionando o arquivo real 'Base Intermediária Fitted July Flash
    2026.xlsx' - é exatamente assim que a usuária já preenchia à mão.

    Devolve uma string de aviso (fallback de Forecast usado, ou fonte não
    encontrada) pra virar popup na GUI, ou None se não houve nenhuma
    ressalva."""
    rotulo_fonte = "Budget" if mes == 1 else "Forecast"
    ws = wb.Worksheets("Pivot")
    ws.Cells(LINHA_PIVOT_DESPESAS_PROPRIO, 1).Value = "Flash"
    ws.Cells(LINHA_PIVOT_DESPESAS_COMPARACAO, 1).Value = rotulo_fonte
    ws.Cells(LINHA_QUADRO_HEADER, COL_QUADRO_CUSTOS_FLASH).Value = rotulo_fonte
    ws.Cells(LINHA_QUADRO_HEADER, COL_QUADRO_CUSTOS_ACTUAL).Value = "Flash"

    resultado = localizar_forecast_para_comparacao(mes, ano, log)
    if resultado is None:
        return (
            f"Não encontrei Forecast/Budget pra comparar no fechamento de {MESES_INGLES[mes]}/{ano} — "
            "o quadro de comparação (linhas 18/19 da aba Pivot) ficou vazio nesta rodada."
        )
    caminho_fonte, aviso = resultado

    log(f"Trazendo Despesas/Mão de Obra pro quadro de comparação: {caminho_fonte.name}...")
    despesas, mao_de_obra = ler_forecast_despesas_mao_de_obra(excel, caminho_fonte, mes, log)

    col = COL_PIVOT_MES_JANEIRO + mes - 1
    ws.Cells(LINHA_PIVOT_DESPESAS_COMPARACAO, col).Value = despesas
    ws.Cells(LINHA_PIVOT_MAO_DE_OBRA_COMPARACAO, col).Value = mao_de_obra
    log(f"  Despesas={despesas:,.2f} | Mão de Obra={mao_de_obra:,.2f}")

    letra_mes = _letra_coluna_mes(mes)
    ws.Cells(LINHA_QUADRO_CUSTOS, COL_QUADRO_CUSTOS_FLASH).Formula = (
        f"=({letra_mes}{LINHA_PIVOT_DESPESAS_COMPARACAO}+{letra_mes}{LINHA_PIVOT_MAO_DE_OBRA_COMPARACAO})/1000"
    )
    ws.Cells(LINHA_QUADRO_CUSTOS, COL_QUADRO_CUSTOS_ACTUAL).Formula = f"={letra_mes}{LINHA_PIVOT_GRAND_TOTAL}/1000"
    log(f"  Fórmulas de Custos (linha {LINHA_QUADRO_CUSTOS}) apontando pra coluna {letra_mes} ({MESES_INGLES[mes]}).")

    return aviso


def localizar_base_intermediaria_mes_anterior(mes: int, ano: int) -> Path:
    mes_ant, ano_ant = (mes - 1, ano) if mes > 1 else (12, ano - 1)
    pasta = resolver_pasta_ciclo(REDE_BASE / str(ano_ant) / MESES_PASTA[mes_ant], mes_ant, "Actual")
    nome = f"Base Intermediária Fitted {MESES_INGLES[mes_ant]} Actual {ano_ant}.xlsx"
    caminho = encontrar_arquivo_mais_recente(pasta, nome)
    if caminho is None:
        raise FileNotFoundError(f"Não encontrei a Base Intermediária Actual do mês anterior: {pasta / nome}")
    return caminho


def _celulas_com_erro(valores, n_cols_rotulo):
    """Devolve [(linha_relativa, coluna, rotulo_A_H)] pra toda celula
    gravada como erro (#N/A etc. vem do COM como int negativo) numa tupla
    de tuplas lida via Range.Value.

    Checa TAMBEM as colunas de rotulo (A-H), nao so' as de valor: um #N/A
    ali e' o caso real de unidade nova que ficou fora do de-para de MF da
    base de contas (Resende/MF 0483 em Agosto/2026, achado em 2026-09-01 -
    a coluna G, 'Centro de Montagem', vinha #N/A em 40 linhas). Como essas
    colunas eram puladas aqui, o erro passava batido na leitura e so'
    estourava depois, ja' colado na Base Intermediaria, com a mensagem
    generica de 'bug conhecido de marshalling' - despistando o diagnostico."""
    erros = []
    for i, linha in enumerate(valores):
        linha = list(linha)
        if linha and linha[0] == "Grand Total":
            continue
        for j, v in enumerate(linha):
            if isinstance(v, int) and v < 0:
                erros.append((i, j, linha[:n_cols_rotulo]))
    return erros


def _corrigir_celulas_com_erro(ws, primeira_sem_cor, linhas_para_colar, log, tentativas=5, espera_s=2):
    """Reconfere a área recém-colada na Intermediária e, pra qualquer célula
    que ainda veio como erro (#N/A) mesmo colando linha por linha (a
    mitigação já em uso contra o bug de marshalling do pywin32/COM — reduz a
    corrupção, mas não zera por completo), reescreve SÓ a célula ruim (não a
    linha inteira, pra não reintroduzir o mesmo risco em escala maior) e
    reconfere de novo. Repete até `tentativas` vezes antes de desistir.
    Devolve a lista de células (linha_absoluta, coluna) que continuam
    erradas ao final — vazia se tudo foi corrigido. Achado ao vivo em
    2026-09-01: a checagem antiga abortava direto na primeira vez, sem
    tentar se auto-corrigir, exigindo que a usuária rodasse tudo de novo à
    mão."""
    n_linhas = len(linhas_para_colar)
    n_cols = len(linhas_para_colar[0])
    faixa = ws.Range(ws.Cells(primeira_sem_cor, 1), ws.Cells(primeira_sem_cor + n_linhas - 1, n_cols))
    for tentativa in range(tentativas):
        conferencia = faixa.Value
        ruins = [
            (i, j) for i, linha in enumerate(conferencia)
            for j, v in enumerate(linha) if isinstance(v, int) and v < 0
        ]
        if not ruins:
            return []
        log(
            f"  {len(ruins)} célula(s) ainda gravada(s) como erro depois da colagem — "
            f"reescrevendo individualmente (tentativa {tentativa + 1}/{tentativas})..."
        )
        for i, j in ruins:
            ws.Cells(primeira_sem_cor + i, j + 1).Value = linhas_para_colar[i][j]
        time.sleep(espera_s)

    conferencia = faixa.Value
    return [
        (primeira_sem_cor + i, j + 1) for i, linha in enumerate(conferencia)
        for j, v in enumerate(linha) if isinstance(v, int) and v < 0
    ]


def ler_pivot_inter(caminho_base_ksb1: Path, excel, log, tentativas=4, espera_s=5):
    log(f"Lendo Pivot_Inter. de {caminho_base_ksb1.name}...")
    wb = com_retry(
        excel.Workbooks.Open, str(caminho_base_ksb1), ReadOnly=True, UpdateLinks=0, log=log
    )
    try:
        ws = wb.Worksheets("Pivot_Inter.")
        last_row = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
        last_col = ws.Cells(4, ws.Columns.Count).End(XL_TO_LEFT).Column

        # Sem recalcular antes de ler, algumas celulas do Pivot_Inter podem
        # ser gravadas como "#N/A" no arquivo final (o calculo assincrono
        # ligado ao link externo da base de contas nao tinha terminado antes
        # do Save) - achado testando ao vivo em 2026-08-21. Mesmo com o
        # CalculateFullRebuild/CalculateUntilAsyncQueriesDone abaixo, o link
        # externo pode levar mais alguns segundos pra "assentar" de verdade
        # (rede) - por isso reconfere e tenta de novo (com espera) antes de
        # desistir, em vez de devolver o #N/A direto pra usuaria decidir
        # manualmente se roda tudo de novo (achado ao vivo em 2026-09-01: os
        # mesmos ~40 valores vinham como erro em tentativas sucessivas do
        # botao, cada uma com um Excel/leitura novos, ate sumirem sozinhos
        # alguns minutos depois - ou seja, e' a origem assentando, nao o bug
        # de marshalling da colagem, que ja e' protegido a parte).
        for tentativa in range(tentativas):
            com_retry(excel.CalculateFullRebuild, log=log)
            com_retry(excel.CalculateUntilAsyncQueriesDone, log=log)
            valores = ws.Range(ws.Cells(4, 1), ws.Cells(last_row, last_col)).Value
            erros = _celulas_com_erro(valores[1:], N_COLS_LABEL)
            if not erros:
                break
            log(
                f"  {len(erros)} célula(s) do Pivot_Inter. ainda vieram como erro "
                f"(link externo provavelmente ainda assentando) — esperando {espera_s}s "
                f"e recalculando de novo ({tentativa + 1}/{tentativas})..."
            )
            time.sleep(espera_s)
        else:
            colunas = sorted({chr(ord("A") + j) if j < 26 else str(j + 1) for _, j, _ in erros})
            rotulos = "\n  ".join(str(rotulo) for _, _, rotulo in erros[:5])
            raise RuntimeError(
                f"{len(erros)} célula(s) do Pivot_Inter. de '{caminho_base_ksb1.name}' continuam "
                f"gravadas como erro (#N/A) mesmo após {tentativas} recálculos — não é o link externo "
                "assentando nem bug do COM, é um problema real nos dados: alguma conta, centro de "
                "custo ou MF sem correspondência na base de contas "
                "(Base_Contas_Contábeis_Fitted_22.xlsx). Se o erro estiver numa coluna de rótulo "
                "(A-H), o suspeito nº 1 é uma unidade/MF nova cadastrada fora do range da fórmula "
                f"de-para. Coluna(s) afetada(s): {', '.join(colunas)}. Primeiras linhas (rótulo A-H):"
                f"\n  {rotulos}"
            )
    finally:
        com_retry(wb.Close, SaveChanges=False, log=log)

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


COL_PROVISAO_CONTA = 3    # C, na Ficha de Solicitação
COL_PROVISAO_CENTRO = 4   # D
COL_PROVISAO_VALOR = 8    # H
LINHA_PROVISAO_INICIO = 13  # primeira linha de dado na Ficha de Solicitação

LINHA_INTER_PROVISAO_INICIO = 2   # primeira linha colorida da Intermediária
COL_INTER_CONTA_FISCAL = 3        # C
COL_INTER_CENTRO_CUSTO = 5        # E
COL_FORMULA_MODELO = [1, 2, 4, 6, 7]  # A, B, D, F, G - colunas com formula VLOOKUP


def arquivo_esta_aberto(caminho: Path) -> bool:
    """O Excel cria um arquivo de lock '~$<nome>' na mesma pasta enquanto
    alguém (a usuária ou outra pessoa) está com o arquivo aberto. Usado pra
    garantir que a leitura do Fast Provisão pega a versão salva de verdade,
    não um estado intermediário ainda sendo editado."""
    return (caminho.parent / f"~${caminho.name}").exists()


def localizar_fast_provisao(mes: int, ano: int) -> Path:
    """Acha o arquivo 'Fast Provisão_<Mês>[_v#].xlsx' de versão mais alta na
    pasta 'Provisões e Reclassificações' do Flash do mês. Ignora arquivos de
    lock do Excel (começam com '~$'). Levanta erro claro se o arquivo
    escolhido estiver aberto no momento (confirmado com a usuária em
    2026-08-21: precisa garantir que está fechado e salvo antes de ler)."""
    pasta = resolver_pasta_ciclo(REDE_BASE / str(ano) / MESES_PASTA[mes], mes, "Flash") / "Provisões e Reclassificações"
    if not pasta.exists():
        raise FileNotFoundError(f"Pasta de Provisões e Reclassificações não encontrada: {pasta}")

    candidatos = [
        f for f in pasta.glob(f"Fast Provisão_{MESES_INGLES[mes]}*.xlsx") if not f.name.startswith("~$")
    ]
    if not candidatos:
        raise FileNotFoundError(f"Nenhum arquivo 'Fast Provisão_{MESES_INGLES[mes]}*.xlsx' encontrado em {pasta}")

    def versao(caminho: Path) -> int:
        m = re.search(r"_v(\d+)\.xlsx$", caminho.name)
        return int(m.group(1)) if m else 1

    escolhido = max(candidatos, key=versao)
    if arquivo_esta_aberto(escolhido):
        raise RuntimeError(
            f"O arquivo '{escolhido.name}' está aberto no Excel agora — feche e salve antes de rodar de novo, "
            "pra garantir que as provisões lidas são a versão final, não uma edição em andamento."
        )
    return escolhido


COR_AMARELA_PROVISAO = 65535  # Interior.Color do amarelo puro (RGB 255,255,0) - so' linhas de provisao
XL_SHIFT_DOWN = -4121
XL_FORMAT_FROM_ABOVE = 0


def encontrar_ultima_linha_amarela(ws) -> int:
    """Acha a última linha amarela (provisão) a partir de
    LINHA_INTER_PROVISAO_INICIO — distinta das verdes/roxas que vêm depois
    (mesma área "colorida", cores diferentes). Usada pra calcular a
    capacidade REAL de provisões (só as amarelas contam, não a área
    verde/roxa) e pra saber onde inserir uma linha nova, sem nunca tocar
    nas verdes/roxas (confirmado pela usuária em 2026-08-22). Usa
    Interior.Color direto (não Pattern, que é igual pras 3 cores, nem
    ColorIndex, que não discrimina bem cores customizadas/tema — confirmado
    inspecionando ao vivo: verde e roxo têm o mesmo ColorIndex mas Color
    diferente)."""
    r = LINHA_INTER_PROVISAO_INICIO
    while ws.Cells(r, 1).Interior.Color == COR_AMARELA_PROVISAO:
        r += 1
    return r - 1


def inserir_linhas_amarelas_novas(ws, n_linhas: int, ultima_linha_amarela: int, log):
    """Insere n_linhas novas linhas amarelas (provisão) logo antes da
    primeira linha verde, empurrando verde/roxa/branca (área de dados) pra
    baixo — nunca mexe no CONTEÚDO das linhas verdes/roxas, só desloca a
    posição delas (efeito colateral inevitável de inserir linha acima).
    Cada linha nova recebe o mesmo formato (cor/borda) da última linha
    amarela existente, via CopyOrigin (sem usar área de transferência).
    Confirmado com a usuária: sempre amarela por padrão, nunca verde/roxa."""
    ponto_insercao = ultima_linha_amarela + 1
    for _ in range(n_linhas):
        ws.Rows(ponto_insercao).Insert(Shift=XL_SHIFT_DOWN, CopyOrigin=XL_FORMAT_FROM_ABOVE)
    log(
        f"  AVISO: {n_linhas} provisão(ões) a mais do que as linhas amarelas disponíveis — "
        f"inserida(s) {n_linhas} linha(s) amarela(s) nova(s) (linhas {ponto_insercao}-"
        f"{ponto_insercao + n_linhas - 1}), empurrando as linhas verdes/roxas/brancas pra baixo. "
        "Nenhum conteúdo das linhas verdes/roxas foi alterado."
    )


def preencher_provisoes_flash(wb, mes: int, ano: int, log):
    """Preenche as linhas coloridas (provisões/reclassificações) da aba
    Intermediária a partir do Fast Provisão do mês (aba 'Ficha de
    Solicitação', linha 13 em diante: coluna C=Conta, D=Centro, H=Valor).
    Cada linha da ficha vira uma linha colorida (linha 13 -> Intermediária
    linha 2, linha 14 -> linha 3, ...). Preenche só C, E e a coluna do mês
    atual; arrasta pra cada linha nova a fórmula VLOOKUP (A, B, D, F, G) que
    já existe na última linha COLORIDA do arquivo (a roxa — mantida de
    propósito só como "molde" de fórmula, confirmado pela usuária). Se as
    provisões não couberem nas linhas amarelas já existentes, insere linhas
    amarelas novas automaticamente (ver inserir_linhas_amarelas_novas) —
    nunca toca no conteúdo das linhas verdes/roxas. As linhas amarelas
    dentro da capacidade que sobrarem sem provisão este mês são limpas por
    completo no final (A até AJ, incluindo Y:AJ) — senão a fórmula herdada
    de Y:AJ fica pendurada e sempre resolve em #N/A (achado ao vivo em
    2026-09-01)."""
    caminho_provisao = localizar_fast_provisao(mes, ano)
    log(f"Lendo provisões de {caminho_provisao.name}...")

    wb_prov = wb.Application.Workbooks.Open(str(caminho_provisao), ReadOnly=True, UpdateLinks=0)
    try:
        ws_prov = wb_prov.Worksheets("Ficha de Solicitação")
        last_row_prov = ws_prov.Cells(ws_prov.Rows.Count, COL_PROVISAO_CONTA).End(XL_UP).Row
        provisoes = []
        for r in range(LINHA_PROVISAO_INICIO, last_row_prov + 1):
            conta = ws_prov.Cells(r, COL_PROVISAO_CONTA).Value
            if conta in (None, ""):
                continue
            centro = ws_prov.Cells(r, COL_PROVISAO_CENTRO).Value
            valor = ws_prov.Cells(r, COL_PROVISAO_VALOR).Value
            provisoes.append((conta, centro, valor))
    finally:
        wb_prov.Close(SaveChanges=False)

    log(f"  {len(provisoes)} provisão(ões)/reclassificação(ões) encontrada(s).")

    ws = wb.Worksheets("Intermediária")

    # Captura ANTES de qualquer insercao: a fórmula "molde" (última linha
    # colorida = roxa) precisa do número de linha ORIGINAL (a regex troca
    # a referência interna da fórmula, ex: C67 -> C<nova linha>) - se
    # capturasse depois de inserir, o número já teria mudado.
    ultima_linha_colorida = encontrar_primeira_linha_sem_cor(ws) - 1
    formulas_modelo = {
        c: ws.Cells(ultima_linha_colorida, c).Formula for c in COL_FORMULA_MODELO
    }

    ultima_linha_amarela = encontrar_ultima_linha_amarela(ws)
    capacidade = ultima_linha_amarela - LINHA_INTER_PROVISAO_INICIO + 1
    if len(provisoes) > capacidade:
        inserir_linhas_amarelas_novas(ws, len(provisoes) - capacidade, ultima_linha_amarela, log)
        ultima_linha_amarela = encontrar_ultima_linha_amarela(ws)

    col_mes = N_COLS_LABEL + mes
    for i, (conta, centro, valor) in enumerate(provisoes):
        r = LINHA_INTER_PROVISAO_INICIO + i
        ws.Cells(r, COL_INTER_CONTA_FISCAL).Value = conta
        ws.Cells(r, COL_INTER_CENTRO_CUSTO).Value = centro
        ws.Cells(r, col_mes).Value = valor
        for c, formula in formulas_modelo.items():
            ws.Cells(r, c).Formula = re.sub(
                rf"([A-Z]){ultima_linha_colorida}\b", rf"\g<1>{r}", formula
            )

    log(f"  {len(provisoes)} linha(s) colorida(s) preenchida(s) (linhas {LINHA_INTER_PROVISAO_INICIO}-{LINHA_INTER_PROVISAO_INICIO + len(provisoes) - 1}).")

    # Linhas amarelas dentro da capacidade mas sem provisão neste mês (ex:
    # mês com menos provisões que o mês anterior, ou capacidade herdada de
    # um mês antigo com mais linhas): limpa por completo, incluindo as
    # fórmulas de Y:AJ herdadas do template/inserção de linha, senão ficam
    # penduradas referenciando uma linha em branco e sempre resolvem em
    # #N/A — achado ao vivo em 2026-09-01 pela usuária, olhando o arquivo
    # direto no Excel (linhas 27+ da Intermediária de Agosto/Flash).
    primeira_linha_sobrando = LINHA_INTER_PROVISAO_INICIO + len(provisoes)
    if primeira_linha_sobrando <= ultima_linha_amarela:
        ws.Range(
            ws.Cells(primeira_linha_sobrando, 1), ws.Cells(ultima_linha_amarela, COL_FORMULA_FIM)
        ).ClearContents()
        log(
            f"  Linhas amarelas sem provisão este mês ({primeira_linha_sobrando}-{ultima_linha_amarela}) "
            "limpas por completo (inclusive fórmulas Y:AJ herdadas), pra não sobrar #N/A."
        )


def limpar_provisoes(ws, log):
    """Apaga o conteúdo das linhas AMARELAS (provisão) POR COMPLETO —
    rótulos/valores (A-T) e as fórmulas herdadas de Y:AJ (Gestorial II até
    Conta Geral) — sem tocar na formatação/cor nem nas linhas verdes/roxas
    (confirmado explicitamente pela usuária em 2026-08-22 — bug real
    corrigido nesta data: a versão anterior apagava até a última linha
    COLORIDA, incluindo verde/roxa, e isso também destruía a fórmula
    "molde" da roxa antes dela ser copiada).

    Limpar só até a coluna T (sem incluir Y:AJ) deixava a fórmula de Y:AJ
    de linhas que ficaram sem uso (ex: mês com menos provisões que o
    anterior) pendurada referenciando uma linha agora em branco — sempre
    resolvendo em #N/A. Achado ao vivo em 2026-09-01 pela usuária, olhando
    o arquivo direto no Excel.

    Usado pelo 'Atualizar Provisões' antes de preencher de novo, mesmo
    espírito do full-rebuild da área branca (evita sobrar provisão antiga
    se a lista nova tiver menos linhas que a anterior)."""
    ultima_linha_amarela = encontrar_ultima_linha_amarela(ws)
    ws.Range(
        ws.Cells(LINHA_INTER_PROVISAO_INICIO, 1), ws.Cells(ultima_linha_amarela, COL_FORMULA_FIM)
    ).ClearContents()
    log(f"  Linhas amarelas (2-{ultima_linha_amarela}) limpas antes de preencher de novo (verdes/roxas não são tocadas).")


def localizar_base_intermediaria_flash_existente(mes: int, ano: int, pasta_saida: Path) -> Path:
    """Acha a Base Intermediária Flash do mês já criada por 'Lançar
    Provisões' (versão mais alta salva em pasta_saida) — usada por
    'Atualizar Provisões' e pela Finalização (Passo 4) quando o ciclo é
    Flash, que não criam mais uma cópia nova do zero."""
    prefixo = f"Base Intermediária Fitted {MESES_INGLES[mes]} Flash {ano}"
    candidatos = list(pasta_saida.glob(f"{prefixo}*.xlsx"))
    if not candidatos:
        raise FileNotFoundError(
            f"Não encontrei nenhuma Base Intermediária Flash de {MESES_INGLES[mes]}/{ano} em {pasta_saida} — "
            "rode 'Lançar Provisões' primeiro."
        )
    return max(candidatos, key=lambda p: p.stat().st_mtime)


def lancar_provisoes(
    mes: int, ano: int, pasta_saida: Path, sufixo_nome: str = "", log=print, pid_callback=None
) -> Path:
    """Passo 3 (Provisões), botão 'Lançar Provisões': cria a Base
    Intermediária Flash do mês (cópia do Actual do mês anterior) e preenche
    as linhas coloridas pela primeira vez a partir do Fast Provisão."""
    caminho_origem = localizar_base_intermediaria_mes_anterior(mes, ano)
    log(f"Partindo da Base Intermediária Actual do mês anterior: {caminho_origem.name}")
    pasta_saida.mkdir(parents=True, exist_ok=True)
    nome_base = f"Base Intermediária Fitted {MESES_INGLES[mes]} Flash {ano}{sufixo_nome}.xlsx"
    nome_saida = nome_com_versao(pasta_saida, nome_base)
    caminho_saida = pasta_saida / nome_saida
    log(f"Copiando {caminho_origem.name} -> {caminho_saida.name} ...")
    shutil.copy2(caminho_origem, caminho_saida)

    excel = abrir_excel_isolado(log, pid_callback)
    try:
        wb = com_retry(
            excel.Workbooks.Open, str(caminho_saida), UpdateLinks=0, ReadOnly=False, log=log
        )
        if wb.ReadOnly:
            raise RuntimeError(
                "A cópia abriu em modo somente leitura — provavelmente já está aberta por outro processo."
            )
        preencher_provisoes_flash(wb, mes, ano, log)
        com_retry(excel.CalculateFullRebuild, log=log)
        log("Salvando...")
        com_retry(wb.Save, log=log)
        com_retry(wb.Close, SaveChanges=False, log=log)
    finally:
        excel.Quit()

    log(f"\nProvisões lançadas: {caminho_saida}")
    return caminho_saida


def atualizar_provisoes(mes: int, ano: int, pasta_saida: Path, log=print, pid_callback=None) -> Path:
    """Passo 3 (Provisões), botão 'Atualizar Provisões': relê o Fast
    Provisão (versão mais alta no momento) e atualiza as linhas coloridas
    da Base Intermediária Flash já existente (não cria cópia nova)."""
    caminho_saida = localizar_base_intermediaria_flash_existente(mes, ano, pasta_saida)
    log(f"Atualizando provisões em: {caminho_saida.name}")

    excel = abrir_excel_isolado(log, pid_callback)
    try:
        wb = com_retry(
            excel.Workbooks.Open, str(caminho_saida), UpdateLinks=0, ReadOnly=False, log=log
        )
        if wb.ReadOnly:
            raise RuntimeError(
                "O arquivo abriu em modo somente leitura — provavelmente já está aberto por outro processo."
            )
        ws = wb.Worksheets("Intermediária")
        limpar_provisoes(ws, log)
        preencher_provisoes_flash(wb, mes, ano, log)
        com_retry(excel.CalculateFullRebuild, log=log)
        log("Salvando...")
        com_retry(wb.Save, log=log)
        com_retry(wb.Close, SaveChanges=False, log=log)
    finally:
        excel.Quit()

    log(f"\nProvisões atualizadas: {caminho_saida}")
    return caminho_saida


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


def atualizar_base_intermediaria(
    mes: int, ano: int, ciclo: str, pasta_saida: Path, sufixo_nome: str = "", log=print, pid_callback=None
):
    if ciclo not in ("Actual", "Flash"):
        raise ValueError(f"Ciclo '{ciclo}' desconhecido — só Actual e Flash são suportados.")

    centros_encerrados = carregar_centros_encerrados()

    caminho_base_ksb1 = localizar_base_ksb1_do_mes(mes, ano, ciclo)
    pasta_saida.mkdir(parents=True, exist_ok=True)

    if ciclo == "Actual":
        caminho_origem_inter = localizar_base_intermediaria_mes_anterior(mes, ano)
        log(f"Partindo da Base Intermediária Actual do mês anterior: {caminho_origem_inter.name}")
        nome_base = f"Base Intermediária Fitted {MESES_INGLES[mes]} {ciclo} {ano}{sufixo_nome}.xlsx"
        nome_saida = nome_com_versao(pasta_saida, nome_base)
        caminho_saida = pasta_saida / nome_saida
        log(f"Copiando {caminho_origem_inter.name} -> {caminho_saida.name} ...")
        shutil.copy2(caminho_origem_inter, caminho_saida)
    else:
        # Flash: o arquivo ja foi criado pelo Passo 3 (botao "Lançar
        # Provisões"), com as linhas coloridas ja preenchidas - so continua
        # nele, nao cria copia nova nem repete as provisoes.
        caminho_saida = localizar_base_intermediaria_flash_existente(mes, ano, pasta_saida)
        log(f"Continuando na Base Intermediária Flash já criada pelo Passo 3: {caminho_saida.name}")

    excel = abrir_excel_isolado(log, pid_callback)
    try:
        cabecalho_pivot, linhas_pivot, n_meses = ler_pivot_inter(caminho_base_ksb1, excel, log)

        wb = com_retry(
            excel.Workbooks.Open, str(caminho_saida), UpdateLinks=0, ReadOnly=False, log=log
        )
        if wb.ReadOnly:
            raise RuntimeError(
                "A cópia abriu em modo somente leitura — provavelmente já está aberta por outro processo."
            )
        ws = wb.Worksheets("Intermediária")

        # Calculo manual durante a colagem linha a linha + AutoFill abaixo:
        # com calculo automatico (padrao do Excel), cada uma das
        # centenas/milhares de escritas dispara recalculo do arquivo inteiro
        # - motivo real da operacao levar 10+ minutos num mes com muitas
        # linhas (mesmo achado de gerar_ksb1_mensal.py, 2026-09-01). Nao
        # afeta a protecao contra o bug de corrupcao #N/A (que e' sobre a
        # GRANULARIDADE da escrita, nao o modo de calculo) - o recalculo
        # completo abaixo (CalculateFullRebuild) continua garantindo o valor
        # certo antes de qualquer leitura de celula. Precisa vir DEPOIS de
        # abrir uma pasta de trabalho - setar Application.Calculation sem
        # nenhum workbook aberto (ex: logo apos abrir_excel_isolado, antes
        # do ler_pivot_inter acima) derruba com "Unable to set the
        # Calculation property of the Application class" (achado ao vivo,
        # 2026-09-01, gerando erro real no Passo 3 pra usuaria).
        com_retry(setattr, excel, "Calculation", XL_CALCULATION_MANUAL, log=log)
        com_retry(setattr, excel, "ScreenUpdating", False, log=log)

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
        com_retry(
            ws.Range(ws.Cells(primeira_sem_cor, 1), ws.Cells(last_row_limpar, COL_FORMULA_FIM)).ClearContents,
            log=log,
        )

        log(f"Colando {n_linhas_novas} linha(s) do Pivot_Inter. (todas, unidades encerradas incluídas por enquanto)...")
        # Colar linha por linha, NUNCA o bloco inteiro de uma vez: escrever um
        # array grande (testado com 601 linhas) via Range.Value em uma unica
        # chamada COM corrompe aleatoriamente algumas celulas em erro #N/A
        # (bug de marshalling do pywin32/COM, nao tem relacao com o conteudo -
        # achado e confirmado testando ao vivo em 2026-08-21: 166 erros colando
        # tudo de uma vez, 25 em blocos de 50, ZERO colando linha por linha).
        linhas_para_colar = [linha[: N_COLS_LABEL + n_meses] for linha in linhas_pivot]

        # A coluna F (Mini-Fábrica) vem do Pivot_Inter. como TEXTO com zero
        # à esquerda ("0499", "0491", "0483"...). Escrever essa string via
        # COM numa célula de formato Geral faz o Excel converter pra NÚMERO
        # (499) e perder o zero - e aí o Passo 4 (Rateio de Custos), que
        # casa a unidade por esse código, não reconhece mais nada: a
        # Gerência ("0499") deixa de existir e nada é rateado (achado ao
        # vivo em 2026-09-01, no fechamento de Agosto - a usuária percebeu
        # que "não está rateando nada da Gerência"). Forçar formato de
        # texto na coluna F da área nova preserva o zero, igual aos
        # arquivos montados à mão (conferido em Julho/2026: sempre '0499').
        com_retry(
            setattr,
            ws.Range(ws.Cells(primeira_sem_cor, COL_MINI_FABRICA), ws.Cells(last_row_final, COL_MINI_FABRICA)),
            "NumberFormat", "@", log=log,
        )

        for i, linha in enumerate(linhas_para_colar):
            r = primeira_sem_cor + i
            ws.Range(ws.Cells(r, 1), ws.Cells(r, N_COLS_LABEL + n_meses)).Value = [linha]

        log("Verificando se alguma célula foi gravada como erro (#N/A) durante a colagem...")
        celulas_com_erro = _corrigir_celulas_com_erro(ws, primeira_sem_cor, linhas_para_colar, log)
        if celulas_com_erro:
            linhas_afetadas = sorted({r for r, _ in celulas_com_erro})
            trecho = linhas_afetadas[:15]
            reticencias = "..." if len(linhas_afetadas) > 15 else ""
            raise RuntimeError(
                f"{len(celulas_com_erro)} célula(s) continuam gravadas como erro (#N/A) mesmo depois de "
                "colar linha por linha e reescrever célula a célula várias vezes — abortando sem salvar. "
                "Não é mais o bug de marshalling conhecido (esse já é corrigido automaticamente agora); "
                f"é um problema novo. Linhas afetadas na Intermediária: {trecho}{reticencias}"
            )

        log("Arrastando a fórmula da coluna U (Total Ano)...")
        ws.Cells(primeira_sem_cor, COL_TOTAL_ANO).Formula = formula_total_ano
        com_retry(
            ws.Cells(primeira_sem_cor, COL_TOTAL_ANO).AutoFill,
            ws.Range(ws.Cells(primeira_sem_cor, COL_TOTAL_ANO), ws.Cells(last_row_final, COL_TOTAL_ANO)),
            XL_FILL_DEFAULT,
            log=log,
        )

        log("Arrastando as fórmulas das colunas Y:AJ...")
        for i, c in enumerate(range(COL_FORMULA_INICIO, COL_FORMULA_FIM + 1)):
            ws.Cells(primeira_sem_cor, c).Formula = formulas_extra[i]
        com_retry(
            ws.Range(
                ws.Cells(primeira_sem_cor, COL_FORMULA_INICIO), ws.Cells(primeira_sem_cor, COL_FORMULA_FIM)
            ).AutoFill,
            ws.Range(
                ws.Cells(primeira_sem_cor, COL_FORMULA_INICIO), ws.Cells(last_row_final, COL_FORMULA_FIM)
            ),
            XL_FILL_DEFAULT,
            log=log,
        )

        log("Recalculando a planilha inteira...")
        com_retry(excel.CalculateFullRebuild, log=log)
        com_retry(excel.CalculateUntilAsyncQueriesDone, log=log)

        # Restaura o modo padrao (automatico) - a colagem/AutoFill lentos ja
        # terminaram e o recalculo completo acima ja garantiu os valores
        # certos; o resto da funcao (leitura de unidades encerradas,
        # comparacao Flash/Forecast) segue em modo automatico, como sempre.
        com_retry(setattr, excel, "Calculation", XL_CALCULATION_AUTOMATIC, log=log)
        com_retry(setattr, excel, "ScreenUpdating", True, log=log)

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
        com_retry(wb.RefreshAll, log=log)
        com_retry(excel.CalculateUntilAsyncQueriesDone, log=log)

        aviso_comparacao = None
        if ciclo == "Actual":
            atualizar_comparacao_flash(excel, wb, mes, ano, log)
        else:
            aviso_comparacao = atualizar_comparacao_forecast(excel, wb, mes, ano, log)

        log("Recalculando de novo (Total Ano das linhas zeradas)...")
        com_retry(excel.CalculateFullRebuild, log=log)
        com_retry(excel.CalculateUntilAsyncQueriesDone, log=log)

        log("Salvando...")
        com_retry(wb.Save, log=log)
        if not wb.Saved:
            raise RuntimeError("wb.Save() retornou mas wb.Saved ainda é False — o arquivo pode não ter sido gravado.")
        com_retry(wb.Close, SaveChanges=False, log=log)
    finally:
        # Retentativa tambem no Quit - se o Excel "ocupado" derrubar essa
        # chamada, ela e' engolida sem retentar, deixa o processo orfao
        # rodando pra sempre (achado ao vivo, 2026-09-01: sobrou um
        # EXCEL.EXE ocioso depois de uma rodada bem-sucedida).
        try:
            com_retry(excel.Quit, log=log)
        except pywintypes.com_error:
            log("AVISO: não consegui confirmar o encerramento do Excel (pode ter ficado um processo órfão).")

    caminho_historico = gerar_historico_unidades_encerradas(
        mes, ano, ciclo, cabecalho_historico, linhas_historico, pasta_saida, log
    )

    log(f"\nBase Intermediária atualizada: {caminho_saida}")
    return caminho_saida, caminho_historico, aviso_comparacao


if __name__ == "__main__":
    if len(sys.argv) not in (4, 5):
        print("Uso: python gerar_base_intermediaria.py <mes> <ano> <Actual|Flash> [lancar_provisoes|atualizar_provisoes|finalizar]")
        sys.exit(1)
    _mes, _ano, _ciclo = int(sys.argv[1]), int(sys.argv[2]), sys.argv[3]
    _acao = sys.argv[4] if len(sys.argv) == 5 else "finalizar"
    _pasta_saida = (
        Path(__file__).resolve().parent.parent.parent.parent.parent
        / "data" / "processed" / "fitted_units_despesas" / "base_intermediaria_teste"
    )
    if _acao == "lancar_provisoes":
        lancar_provisoes(_mes, _ano, _pasta_saida, sufixo_nome=" - TESTE VALIDAÇÃO")
    elif _acao == "atualizar_provisoes":
        atualizar_provisoes(_mes, _ano, _pasta_saida)
    elif _acao == "finalizar":
        atualizar_base_intermediaria(_mes, _ano, _ciclo, _pasta_saida, sufixo_nome=" - TESTE VALIDAÇÃO")
    else:
        print(f"Ação desconhecida: {_acao}")
        sys.exit(1)
