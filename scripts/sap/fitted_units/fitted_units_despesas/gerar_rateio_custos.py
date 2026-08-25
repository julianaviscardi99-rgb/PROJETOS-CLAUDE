#!/usr/bin/env python3
"""
Passo 5 (Rateio de Custos) do processo recorrente (Fitted Units Despesas).

Substitui o arquivo manual "_Abertura custos Fitted Units <Mes> <Ciclo>
<Ano>.xlsx" (link externo trocado a mao todo mes + celula de mes/Ciclo
trocada a mao) por um arquivo novo, gerado direto a partir da Base
Intermediaria do mes/Ciclo (ja gerada pelo Passo 4) - sem link nenhum.

Desenhado com a Juliana em 2026-08-25 (ver memory/BRIEFING.md e
memory/DECISOES.md dessa data pro historico completo da conversa):

1. Le a aba "Intermediaria" da Base Intermediaria do mes/Ciclo (mesma pra
   Actual e Flash, sem excecao) e classifica cada linha por Mini-Fabrica
   (unidade) e por conta+tipo (Variavel/Fixo) numa das subcategorias do
   MAPA_CATEGORIAS (extraido linha a linha do arquivo antigo em 2026-08-25).
2. Unidades ativas: SJP (0490), IBI (0491), GOI (0481), RES (0483).
   Gerencia (0499) e' tratada como uma "unidade" a parte (coluna GER no
   quadro sem rateio), nao recebe nem manda rateio pra si mesma.
3. Unidades ENCERRADAS (ontology/fitted_units.json ->
   centros_de_custo_por_unidade, status "encerrada", chave por Centro de
   Custo - mesma fonte que gerar_base_intermediaria.py ja usa) nunca entram
   no quadro nem recebem rateio (sem faturamento). Se aparecer residuo de
   custo nelas:
   - Sorocaba: fica de fora do rateio (esta em reclassificacao pra custo
     nao-recorrente, tratada em outro lugar) - so' aparece no aviso.
   - Qualquer outra unidade encerrada: o residuo soma no total da Gerencia
     antes de aplicar o rateio.
   Em ambos os casos, um aviso com as linhas (conta/descricao/unidade/valor)
   aparece no arquivo gerado - nunca fica escondido so' no log.
4. O rateio (% por unidade, muda geralmente em Janeiro mas pode mudar fora
   de epoca) fica num arquivo de configuracao separado
   (ontology/rateio_gerencia.json), NUNCA hardcoded aqui - ver
   carregar_rateio_vigente.
5. Gera 3 blocos no arquivo de saida: tabela do rateio vigente, quadro "sem
   rateio" (unidades ativas + GER, cada uma com seu proprio custo) e quadro
   "com rateio" (so' unidades ativas, com uma linha propria "Rateio
   Gerencia" e um Check provando que o TOTAL dos dois quadros bate).
6. Salva no mesmo racional de sempre (resolver_pasta_ciclo), nunca
   sobrescreve (nome_com_versao).

NAO mexe em nada da Base Intermediaria - so' le (openpyxl, ReadOnly,
data_only=True - nao precisa de Excel/COM, os valores ja estao calculados
no arquivo).
"""
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "_shared"))
from ksb1_core import (  # noqa: E402
    MESES_PASTA,
    REDE_BASE,
    encontrar_arquivo_mais_recente,
    nome_com_versao,
    resolver_pasta_ciclo,
)

MESES_INGLES = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December",
}
MESES_ABREV = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
}

ONTOLOGY_DIR = Path(__file__).resolve().parents[4] / "ontology"
ONTOLOGY_UNIDADES_PATH = ONTOLOGY_DIR / "fitted_units.json"
RATEIO_CONFIG_PATH = ONTOLOGY_DIR / "rateio_gerencia.json"

# Arquivo mestre (mantido pela controladoria central, nao pela usuaria) com
# a classificacao Variavel/Fixo por Conta Gestorial - ver
# ontology/fitted_units.json -> fonte_contas_gestoriais_e_centros. Usado so'
# como FALLBACK quando o campo "Tp.Custo" da Base Intermediaria vem em
# branco (achado real em 2026-08-25: a Base Intermediaria le esse campo por
# formula, e o cache pode ficar em branco se a formula nao recalculou pra
# aquela unidade - o SJP de julho/2026 veio assim). A Base Intermediaria
# continua sendo a referencia principal (pedido explicito da usuaria); esse
# arquivo so' entra quando ela nao tem a informacao.
BASE_CONTAS_PATH = Path(
    r"\\FSS024-01BR.group.pirelli.com\GFU_DAC\Custos Fitted Units\Custos Efetivos"
    r"\_Acompanhamentos e Controles\Base_Contas_Contábeis_Fitted_22.xlsx"
)

# Mini-Fabrica -> sigla, so' as unidades ATIVAS que entram no quadro/rateio.
# Confirmado com a usuaria em 2026-08-25 (RES ainda sem custo real ate essa
# data, codigo confirmado por ela: "0483 e' de Resende").
UNIDADES_ATIVAS = {
    "0490": "SJP",
    "0491": "IBI",
    "0481": "GOI",
    "0483": "RES",
}
GERENCIA_MINIFABRICA = "0499"
SIGLA_GERENCIA = "GER"

# Ordem de exibicao das unidades no quadro (GER so' aparece no quadro "sem
# rateio", nunca no "com rateio").
ORDEM_UNIDADES_ATIVAS = ["SJP", "IBI", "GOI", "RES"]

ORDEM_VARIAVEL = ["Labour", "Handling", "Direct Materials", "Transportation", "Other Variable"]
ORDEM_FIXO = ["Labour", "Depreciation", "IFRS16", "Rents", "Condominio", "Other Fixed"]

# Mapeamento conta gestorial + tipo (V/F) -> subcategoria. Extraido linha a
# linha do arquivo real "_Abertura custos Fitted Units July Actual 2026.xlsx"
# (aba "Sao J. dos Pinhais") em 2026-08-25 - ver memory/DECISOES.md dessa
# data pro detalhe de como foi extraido. Se aparecer uma conta nova que nao
# esta aqui, cai no fallback ("Other Variable"/"Other Fixed" conforme o
# tipo) E gera um aviso no arquivo, pra usuaria decidir a categoria certa.
_CONTAS_VARIAVEL = {
    "Labour": [
        4201000, 4201100, 4201110, 4201120, 4201130, 4201140, 4201150,
        4201160, 4201170, 4201180, 4201190, 4254110, 4201400, 4245100,
        4201410, 4201135,
    ],
    "Handling": [4201250, 4257000, 4257001],
    "Direct Materials": [4236300, 4240000],
    "Transportation": [4211000],
    "Other Variable": [
        4203050, 4203055, 4203400, 4203700, 4205250, 4205350, 4205450,
        4205500, 4222100, 4236100, 4260000, 4212000, 4263000, 4266000,
        4258012,
    ],
}
_CONTAS_FIXO = {
    "Labour": [
        4201000, 4201100, 4201110, 4201120, 4201130, 4201140, 4201150,
        4201160, 4201170, 4201180, 4201190, 4254110, 4201410, 4245100,
        4200000, 4200100, 4200110, 4200120, 4200130, 4200140, 4200150,
        4200160, 4200170, 4200180, 4200190, 4247100, 4254100, 4200400,
        4200410, 4245000, 4200135, 4201135,
    ],
    "Depreciation": [4255000, 4255200],
    "IFRS16": [4255002],
    "Rents": [4257000, 4257001],
    "Condominio": [42570001],
    "Other Fixed": [
        4203050, 4203055, 4203400, 4203700, 4205250, 4205350, 4205450,
        4205500, 4222100, 4202000, 4202100, 4204000, 4205100, 4211000,
        4212000, 4222000, 4230000, 4230100, 4232000, 4232001, 4233000,
        4234000, 4235000, 4235100, 4236020, 4236050, 4236100, 4236200,
        4237000, 4237100, 4238000, 4240000, 4243000, 4243100, 4244000,
        4244100, 4246000, 4246100, 4247200, 4253000, 4254000, 4222001,
        4258001, 4258002, 4258021, 4258000, 4258010, 4258020, 4258012,
        4258100, 4260000, 4248000, 4262100, 4262300, 4230101, 4263000,
        4266000,
    ],
}


def _construir_mapa_categorias() -> dict:
    mapa = {}
    for subcat, contas in _CONTAS_VARIAVEL.items():
        for conta in contas:
            mapa[("V", conta)] = subcat
    for subcat, contas in _CONTAS_FIXO.items():
        for conta in contas:
            mapa[("F", conta)] = subcat
    return mapa


MAPA_CATEGORIAS = _construir_mapa_categorias()


def carregar_centros_encerrados() -> dict:
    """Retorna {centro_de_custo (str): nome_da_unidade} pra toda unidade com
    status 'encerrada' - mesma fonte/formato que gerar_base_intermediaria.py
    ja usa (ontology/fitted_units.json), pra nao duplicar a lista."""
    dados = json.loads(ONTOLOGY_UNIDADES_PATH.read_text(encoding="utf-8"))
    grupos = dados["centros_de_custo_por_unidade"]["grupos"]
    mapa = {}
    for nome_grupo, info in grupos.items():
        if info.get("status") == "encerrada":
            for c in info["centros"]:
                mapa[str(c)] = nome_grupo
    return mapa


def carregar_rateio_vigente(mes: int, ano: int) -> tuple[dict, str]:
    """Le ontology/rateio_gerencia.json e devolve (percentuais, vigente_desde)
    da entrada mais recente cujo 'vigente_desde' (formato 'AAAA-MM') seja
    <= o mes/ano pedido. NUNCA hardcoda a %, porque ela muda (geralmente em
    Janeiro, mas as vezes fora de epoca - ex: quando a Resende entrou)."""
    dados = json.loads(RATEIO_CONFIG_PATH.read_text(encoding="utf-8"))
    alvo = f"{ano:04d}-{mes:02d}"
    candidatas = [e for e in dados["entradas"] if e["vigente_desde"] <= alvo]
    if not candidatas:
        raise RuntimeError(
            f"Não encontrei nenhuma % de rateio vigente pra {alvo} em "
            f"{RATEIO_CONFIG_PATH} — confira o arquivo de configuração."
        )
    escolhida = max(candidatas, key=lambda e: e["vigente_desde"])
    return escolhida["percentuais"], escolhida["vigente_desde"]


def localizar_base_intermediaria(mes: int, ano: int, ciclo: str) -> Path:
    pasta = resolver_pasta_ciclo(REDE_BASE / str(ano) / MESES_PASTA[mes], mes, ciclo)
    nome = f"Base Intermediária Fitted {MESES_INGLES[mes]} {ciclo} {ano}.xlsx"
    caminho = encontrar_arquivo_mais_recente(pasta, nome)
    if caminho is None:
        raise FileNotFoundError(
            f"Não encontrei '{nome}' em {pasta} — rode o Passo 4 (Base "
            "Intermediária) pra esse mês/Ciclo antes de rodar o Rateio de Custos."
        )
    return caminho


def carregar_var_fallback_por_conta() -> dict:
    """Le Base_Contas_Contábeis_Fitted_22.xlsx (aba 'Contas') e devolve
    {conta_gestorial (int): 'F'|'V'} - so' pras contas com classificacao
    SEM ambiguidade (ignora 'F/V' e em branco, ja que nesses casos nao da'
    pra decidir sozinho; melhor deixar cair no aviso de "nao classificado"
    do que adivinhar errado)."""
    import openpyxl as _openpyxl

    wb = _openpyxl.load_workbook(BASE_CONTAS_PATH, read_only=True, data_only=True, keep_links=False)
    ws = wb["Contas"]
    mapa = {}
    for r in range(3, ws.max_row + 1):
        conta_gestorial = ws.cell(row=r, column=3).value
        var = ws.cell(row=r, column=8).value
        if conta_gestorial is None or var not in ("F", "V"):
            continue
        try:
            mapa[int(conta_gestorial)] = var
        except (TypeError, ValueError):
            continue
    return mapa


def _fmt_moeda(v):
    if v is None:
        v = 0
    v = round(v, 1)
    if v < 0:
        return f"({abs(v):,.1f})".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{v:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")


def ler_e_classificar(caminho_base_intermediaria: Path, mes: int, log):
    """Le a aba Intermediaria e devolve:
    - totais_ativos: {sigla_unidade: {subcategoria: valor}} (SJP/IBI/GOI/RES/GER,
      subcategoria pode ser "Não Classificado" quando nem a Base Intermediária
      nem o fallback do Base_Contas souberam dizer Variável/Fixo)
    - residuos_encerradas: lista de dicts (unidade, conta, descricao, valor)
      pra toda unidade encerrada com valor != 0 no mes
    - contas_nao_mapeadas: set de (tipo, conta, desc) que caiu no fallback
      de subcategoria (Other Variable/Fixed)
    Valores em '000 BRL, custo positivo (mesma convencao do arquivo antigo:
    o valor bruto do SAP vem negativo pra custo, aqui ja inverte o sinal).

    Prioridade da classificação Variável/Fixo (2026-08-25, achado real: o
    SJP de julho veio com "Tp.Custo" em branco na Base Intermediária, porque
    esse campo é uma fórmula que às vezes não recalcula):
    1. Campo "Tp.Custo" da própria Base Intermediária, se vier preenchido
       (V ou F) — é a referência principal, pedido explícito da usuária.
    2. Se vier em branco: cai no arquivo mestre Base_Contas_Contábeis
       (mantido pela controladoria central), buscando pela Conta Gestorial.
    3. Se nem assim resolver (conta não encontrada, ou o mestre marca ela
       como ambígua "F/V"): entra na subcategoria "Não Classificado" — nunca
       some o valor, sempre visível no arquivo."""
    wb = __import__("openpyxl").load_workbook(
        caminho_base_intermediaria, read_only=True, data_only=True, keep_links=False
    )
    ws = wb["Intermediária"]
    col_mes = 8 + mes  # A=1...H=8, I(Jan)=9 -> mes=1 vira coluna 9

    centros_encerrados = carregar_centros_encerrados()
    var_fallback = carregar_var_fallback_por_conta()

    totais_ativos = {sigla: {} for sigla in list(UNIDADES_ATIVAS.values()) + [SIGLA_GERENCIA]}
    residuos_encerradas = []
    contas_nao_mapeadas = set()
    contas_fallback_usado = set()
    contas_nao_classificadas = set()

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=max(col_mes, 8)):
        conta = row[0].value  # A - Conta Gestorial
        desc_conta = row[1].value  # B
        centro_custo = row[4].value  # E
        mini_fabrica = row[5].value  # F
        tipo = row[7].value  # H - 'V' ou 'F'
        if conta is None:
            continue
        valor_bruto = row[col_mes - 1].value if col_mes - 1 < len(row) else None
        if not valor_bruto:
            continue
        valor = -valor_bruto / 1000  # mesma convencao do arquivo antigo

        try:
            conta_int = int(conta)
        except (TypeError, ValueError):
            conta_int = conta

        if tipo not in ("V", "F"):
            fallback = var_fallback.get(conta_int) if isinstance(conta_int, int) else None
            if fallback is not None:
                tipo = fallback
                contas_fallback_usado.add((conta_int, desc_conta))
            else:
                contas_nao_classificadas.add((conta_int, desc_conta))

        mini_fabrica_str = str(mini_fabrica).strip() if mini_fabrica is not None else ""
        centro_custo_str = str(centro_custo).strip() if centro_custo is not None else ""

        if mini_fabrica_str in UNIDADES_ATIVAS:
            sigla = UNIDADES_ATIVAS[mini_fabrica_str]
        elif mini_fabrica_str == GERENCIA_MINIFABRICA:
            sigla = SIGLA_GERENCIA
        elif centro_custo_str in centros_encerrados:
            residuos_encerradas.append(
                {
                    "unidade": centros_encerrados[centro_custo_str],
                    "conta": conta_int,
                    "descricao": desc_conta,
                    "valor": valor,
                }
            )
            continue
        else:
            # Nem unidade ativa, nem Gerencia, nem unidade encerrada
            # conhecida (ex: FATURAMENTO, que ja e' pra ignorar) - fora de
            # escopo, nao entra no quadro nem no aviso.
            continue

        if tipo not in ("V", "F"):
            chave = "Não Classificado"
        else:
            subcat = MAPA_CATEGORIAS.get((tipo, conta_int))
            if subcat is None:
                contas_nao_mapeadas.add((tipo, conta_int, desc_conta))
                subcat = "Other Variable" if tipo == "V" else "Other Fixed"
            # Chave = (tipo, subcategoria), NUNCA so' a subcategoria: "Labour"
            # existe tanto em Variable Cost quanto em Fixed Cost - se a chave
            # fosse so' o nome, os dois se misturariam num balde so' e o
            # Total Costs contaria o mesmo valor 2x (bug real encontrado e
            # corrigido em 2026-08-25, testando contra Julho/Actual real).
            chave = (tipo, subcat)

        totais_ativos[sigla][chave] = totais_ativos[sigla].get(chave, 0) + valor

    if contas_fallback_usado:
        for conta_int, desc in sorted(contas_fallback_usado, key=lambda x: str(x[0])):
            log(
                f"INFO: conta {conta_int} ({desc}) veio sem Tp.Custo na Base Intermediária "
                "— usei a classificação do Base_Contas_Contábeis como fallback."
            )
    if contas_nao_classificadas:
        for conta_int, desc in sorted(contas_nao_classificadas, key=lambda x: str(x[0])):
            log(
                f"AVISO: conta {conta_int} ({desc}) veio sem Tp.Custo na Base Intermediária "
                "e também não encontrei uma classificação clara no Base_Contas_Contábeis "
                "(ambígua ou não cadastrada) — entrou em 'Não Classificado', confira o valor."
            )
    if contas_nao_mapeadas:
        for tipo, conta_int, desc in sorted(contas_nao_mapeadas, key=lambda x: (x[0], str(x[1]))):
            log(
                f"AVISO: conta {conta_int} ({desc}, tipo {tipo}) não estava mapeada — "
                "caiu no 'Other Variable/Fixed' por padrão. Confirme a categoria certa."
            )

    return totais_ativos, residuos_encerradas, contas_nao_mapeadas


def calcular_rateio(totais_ativos: dict, residuos_encerradas: list, percentuais: dict, log):
    """Devolve (rateio_por_unidade, gerencia_total, residuo_somado_gerencia,
    residuo_sorocaba_fora). Gerencia_total = soma de todas as subcategorias
    da Gerencia + residuo de unidades encerradas (exceto Sorocaba)."""
    ger = totais_ativos.get(SIGLA_GERENCIA, {})
    gerencia_total = sum(ger.values())

    residuo_somado_gerencia = 0.0
    residuo_sorocaba_fora = 0.0
    for r in residuos_encerradas:
        if r["unidade"].upper() == "SOROCABA":
            residuo_sorocaba_fora += r["valor"]
        else:
            residuo_somado_gerencia += r["valor"]
            log(
                f"AVISO: resíduo de {r['unidade']} (conta {r['conta']} - {r['descricao']}: "
                f"{_fmt_moeda(r['valor'])}) somado ao custo da Gerência antes do rateio."
            )
    if residuo_sorocaba_fora:
        log(
            f"AVISO: resíduo de Sorocaba no mês ({_fmt_moeda(residuo_sorocaba_fora)}) — "
            "NÃO entra no rateio (unidade em reclassificação para custo não recorrente)."
        )

    base_rateio = gerencia_total + residuo_somado_gerencia

    rateio_por_unidade = {}
    for sigla in ORDEM_UNIDADES_ATIVAS:
        pct = percentuais.get(sigla, 0)
        rateio_por_unidade[sigla] = base_rateio * pct

    return rateio_por_unidade, gerencia_total, residuo_somado_gerencia, residuo_sorocaba_fora


# ---------------------------------------------------------------------------
# Geracao do arquivo Excel
# ---------------------------------------------------------------------------

AZUL_ESCURO = "1F3864"
AZUL_CLARO_TOTAL = "D9E2F3"
CINZA_CLARO = "F2F2F2"
BRANCO = "FFFFFF"
FONTE_CABECALHO = Font(name="Calibri", size=11, bold=True, color=BRANCO)
FONTE_CATEGORIA = Font(name="Calibri", size=10, bold=True)
FONTE_ITEM = Font(name="Calibri", size=10, italic=True)
FONTE_TOTAL = Font(name="Calibri", size=10, bold=True)
FONTE_NOTA = Font(name="Calibri", size=9, italic=True, color="555555")
BORDA_FINA = Border(bottom=Side(style="thin", color="BFBFBF"))


def _escrever_quadro(ws, linha_inicio, col_inicio, titulo, colunas, dados, ordem_var, ordem_fixo, coluna_total=True):
    """Escreve um quadro (Variable Cost + Fixed Cost + Total Costs) a
    partir da linha/coluna dadas. `dados` = {coluna: {subcategoria: valor}}.
    Devolve a linha seguinte livre depois do quadro."""
    n_cols = len(colunas) + (1 if coluna_total else 0)

    # Cabecalho
    ws.merge_cells(start_row=linha_inicio, start_column=col_inicio, end_row=linha_inicio, end_column=col_inicio + n_cols)
    cel = ws.cell(row=linha_inicio, column=col_inicio, value=titulo)
    cel.font = Font(name="Calibri", size=13, bold=True)
    cel.alignment = Alignment(horizontal="center")

    linha_nota = linha_inicio + 1
    ws.cell(row=linha_nota, column=col_inicio, value="(+) gain   |   '000 BRL").font = FONTE_NOTA

    linha_cab = linha_inicio + 2
    ws.cell(row=linha_cab, column=col_inicio, value="")
    for i, colname in enumerate(colunas):
        c = ws.cell(row=linha_cab, column=col_inicio + 1 + i, value=colname)
        c.font = FONTE_CABECALHO
        c.fill = PatternFill("solid", fgColor=AZUL_ESCURO)
        c.alignment = Alignment(horizontal="center")
    if coluna_total:
        c = ws.cell(row=linha_cab, column=col_inicio + 1 + len(colunas), value="TOTAL")
        c.font = FONTE_CABECALHO
        c.fill = PatternFill("solid", fgColor=AZUL_ESCURO)
        c.alignment = Alignment(horizontal="center")

    linha = linha_cab + 1
    linhas_dados_totais = []  # (linha, e' linha de "Total Costs"?)

    def _linha_categoria(nome):
        nonlocal linha
        c = ws.cell(row=linha, column=col_inicio, value=nome)
        c.font = FONTE_CATEGORIA
        primeira_linha_cat = linha
        return primeira_linha_cat

    def _linha_item(nome, tipo_macro):
        # tipo_macro = "V" ou "F" - "Labour" (e outros nomes) existem nos
        # dois macros com valores DIFERENTES, por isso a chave em `dados` e'
        # sempre (tipo, nome), nunca so' o nome.
        nonlocal linha
        c = ws.cell(row=linha, column=col_inicio, value=nome)
        c.font = FONTE_ITEM
        for i, colname in enumerate(colunas):
            valor = dados.get(colname, {}).get((tipo_macro, nome), 0)
            cc = ws.cell(row=linha, column=col_inicio + 1 + i, value=round(valor, 1))
            cc.font = FONTE_ITEM
            cc.number_format = "#,##0.0;(#,##0.0)"
        linha += 1

    # Variable Cost
    linha_var_categoria = _linha_categoria("Variable Cost")
    linha += 1
    linha_var_inicio_itens = linha
    for item in ordem_var:
        _linha_item(item, "V")
    linha_var_fim_itens = linha - 1

    # Fixed Cost
    linha_fix_categoria = _linha_categoria("Fixed Cost")
    linha += 1
    linha_fix_inicio_itens = linha
    for item in ordem_fixo:
        _linha_item(item, "F")
    linha_fix_fim_itens = linha - 1

    # Preenche a soma da linha "Variable Cost"/"Fixed Cost" (soma dos itens dela)
    for i, colname in enumerate(colunas):
        col_letra = get_column_letter(col_inicio + 1 + i)
        ws.cell(
            row=linha_var_categoria, column=col_inicio + 1 + i,
            value=f"=SUM({col_letra}{linha_var_inicio_itens}:{col_letra}{linha_var_fim_itens})",
        ).font = FONTE_CATEGORIA
        ws.cell(row=linha_var_categoria, column=col_inicio + 1 + i).number_format = "#,##0.0;(#,##0.0)"
        ws.cell(
            row=linha_fix_categoria, column=col_inicio + 1 + i,
            value=f"=SUM({col_letra}{linha_fix_inicio_itens}:{col_letra}{linha_fix_fim_itens})",
        ).font = FONTE_CATEGORIA
        ws.cell(row=linha_fix_categoria, column=col_inicio + 1 + i).number_format = "#,##0.0;(#,##0.0)"

    return {
        "linha_var_categoria": linha_var_categoria,
        "linha_fix_categoria": linha_fix_categoria,
        "linha_apos_itens": linha,
        "colunas": colunas,
        "col_inicio": col_inicio,
        "n_cols": n_cols,
        "coluna_total": coluna_total,
    }


def _finalizar_quadro_com_total(ws, info, linhas_extra_antes_do_total=None):
    """Adiciona (se pedido) linhas extras entre Fixed Cost e Total Costs
    (ex: 'Rateio Gerência'), depois a linha Total Costs (= Variable + Fixed
    + extras), e a coluna TOTAL (soma das unidades) em toda a área de dados."""
    linha = info["linha_apos_itens"]
    col_inicio = info["col_inicio"]
    colunas = info["colunas"]

    linhas_extras_num = []
    if linhas_extra_antes_do_total:
        for nome, valores in linhas_extra_antes_do_total:
            c = ws.cell(row=linha, column=col_inicio, value=nome)
            c.font = FONTE_ITEM
            for i, colname in enumerate(colunas):
                cc = ws.cell(row=linha, column=col_inicio + 1 + i, value=round(valores.get(colname, 0), 1))
                cc.font = FONTE_ITEM
                cc.number_format = "#,##0.0;(#,##0.0)"
            linhas_extras_num.append(linha)
            linha += 1

    linha_total_costs = linha
    ws.cell(row=linha_total_costs, column=col_inicio, value="Total Costs").font = FONTE_TOTAL
    for i, colname in enumerate(colunas):
        col_letra = get_column_letter(col_inicio + 1 + i)
        partes = [f"{col_letra}{info['linha_var_categoria']}", f"{col_letra}{info['linha_fix_categoria']}"]
        partes += [f"{col_letra}{r}" for r in linhas_extras_num]
        cc = ws.cell(row=linha_total_costs, column=col_inicio + 1 + i, value="=" + "+".join(partes))
        cc.font = FONTE_TOTAL
        cc.number_format = "#,##0.0;(#,##0.0)"
        cc.border = Border(top=Side(style="thin", color="000000"))

    # Coluna TOTAL: soma das unidades, linha a linha, da primeira linha de
    # categoria ate a linha Total Costs.
    if info["coluna_total"]:
        col_total_letra = get_column_letter(col_inicio + 1 + len(colunas))
        primeira = info["linha_var_categoria"]
        for r in range(primeira, linha_total_costs + 1):
            letras = [get_column_letter(col_inicio + 1 + i) for i in range(len(colunas))]
            formula = "=" + "+".join(f"{l}{r}" for l in letras)
            cc = ws.cell(row=r, column=col_inicio + 1 + len(colunas), value=formula)
            cc.number_format = "#,##0.0;(#,##0.0)"
            cc.fill = PatternFill("solid", fgColor=AZUL_CLARO_TOTAL)
            if r == linha_total_costs:
                cc.font = FONTE_TOTAL
                cc.border = Border(top=Side(style="thin", color="000000"))

    return linha_total_costs, linha + 1


def gerar_arquivo_rateio_custos(mes: int, ano: int, ciclo: str, pasta_saida: Path, log=print) -> Path:
    caminho_base = localizar_base_intermediaria(mes, ano, ciclo)
    log(f"Lendo Base Intermediária: {caminho_base.name}...")
    totais_ativos, residuos_encerradas, _ = ler_e_classificar(caminho_base, mes, log)

    percentuais, vigente_desde = carregar_rateio_vigente(mes, ano)
    log(f"Rateio vigente desde {vigente_desde}: {percentuais}")

    rateio_por_unidade, gerencia_total, residuo_somado, residuo_sorocaba = calcular_rateio(
        totais_ativos, residuos_encerradas, percentuais, log
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "Rateio de Custos"
    ws.sheet_view.showGridLines = False

    titulo_periodo = f"{MESES_ABREV[mes]} {ano} - {ciclo}"

    # --- Bloco 1: tabela do rateio vigente ---
    ws.cell(row=1, column=1, value=f"Rateio de Custos — {titulo_periodo}").font = Font(size=14, bold=True)
    ws.cell(row=2, column=1, value=f"Rateio Gerência vigente desde {vigente_desde}").font = FONTE_NOTA

    linha = 4
    ws.cell(row=linha, column=1, value="CM").font = FONTE_CABECALHO
    ws.cell(row=linha, column=1).fill = PatternFill("solid", fgColor=AZUL_ESCURO)
    ws.cell(row=linha, column=2, value="Rateio Atual").font = FONTE_CABECALHO
    ws.cell(row=linha, column=2).fill = PatternFill("solid", fgColor=AZUL_ESCURO)
    linha += 1
    linha_pct_inicio = linha
    for sigla in ORDEM_UNIDADES_ATIVAS:
        ws.cell(row=linha, column=1, value=sigla)
        cc = ws.cell(row=linha, column=2, value=percentuais.get(sigla, 0))
        cc.number_format = "0.0%"
        linha += 1
    linha_pct_fim = linha - 1
    ws.cell(row=linha, column=1, value="TOTAL").font = FONTE_TOTAL
    cc = ws.cell(row=linha, column=2, value=f"=SUM(B{linha_pct_inicio}:B{linha_pct_fim})")
    cc.number_format = "0.0%"
    cc.font = FONTE_TOTAL

    # --- Avisos de resíduo (unidades encerradas) ---
    linha_aviso = linha + 3
    if residuos_encerradas:
        ws.cell(row=linha_aviso, column=1, value="⚠ Resíduo de unidade(s) encerrada(s) neste mês").font = Font(bold=True, color="9C0006")
        linha_aviso += 1
        cab = ["Unidade", "Conta", "Descrição", "Valor ('000 BRL)", "Entrou no rateio?"]
        for i, c in enumerate(cab):
            ws.cell(row=linha_aviso, column=1 + i, value=c).font = Font(bold=True)
        linha_aviso += 1
        for r in residuos_encerradas:
            entrou = "Não (Sorocaba — reclassificação não recorrente)" if r["unidade"].upper() == "SOROCABA" else "Sim (somado à Gerência)"
            ws.cell(row=linha_aviso, column=1, value=r["unidade"])
            ws.cell(row=linha_aviso, column=2, value=r["conta"])
            ws.cell(row=linha_aviso, column=3, value=r["descricao"])
            cc = ws.cell(row=linha_aviso, column=4, value=round(r["valor"], 1))
            cc.number_format = "#,##0.0;(#,##0.0)"
            ws.cell(row=linha_aviso, column=5, value=entrou)
            linha_aviso += 1
        linha_prox_bloco = linha_aviso + 2
    else:
        linha_prox_bloco = linha_aviso

    # --- Bloco 2: quadro sem rateio (SJP | IBI | GOI | RES | GER | TOTAL) ---
    colunas_sem_rateio = ORDEM_UNIDADES_ATIVAS + [SIGLA_GERENCIA]
    tem_nao_classificado = any(
        totais_ativos.get(c, {}).get("Não Classificado", 0) for c in colunas_sem_rateio
    )
    valores_nao_classificado = {
        c: totais_ativos.get(c, {}).get("Não Classificado", 0) for c in colunas_sem_rateio
    }
    extras_quadro1 = []
    if tem_nao_classificado:
        extras_quadro1.append(("Não Classificado", valores_nao_classificado))
    if residuo_somado:
        # Soma na coluna GER pra o TOTAL deste quadro bater com o TOTAL do
        # quadro "com rateio" (que ja' inclui esse residuo na base do
        # rateio) - sem isso o Check nunca fecharia em 0.
        extras_quadro1.append((
            "Resíduo Unidades Encerradas",
            {c: (residuo_somado if c == SIGLA_GERENCIA else 0) for c in colunas_sem_rateio},
        ))
    extras_quadro1 = extras_quadro1 or None

    info1 = _escrever_quadro(
        ws, linha_prox_bloco, 1,
        f"{titulo_periodo} — por unidade (sem rateio)",
        colunas_sem_rateio, totais_ativos, ORDEM_VARIAVEL, ORDEM_FIXO,
    )
    linha_total1, linha_prox_bloco = _finalizar_quadro_com_total(ws, info1, linhas_extra_antes_do_total=extras_quadro1)
    linha_prox_bloco += 2

    # --- Bloco 3: quadro com rateio (SJP | IBI | GOI | RES | TOTAL) ---
    dados_com_rateio = {sigla: totais_ativos.get(sigla, {}) for sigla in ORDEM_UNIDADES_ATIVAS}
    extras_quadro2 = []
    if tem_nao_classificado:
        extras_quadro2.append((
            "Não Classificado",
            {c: totais_ativos.get(c, {}).get("Não Classificado", 0) for c in ORDEM_UNIDADES_ATIVAS},
        ))
    extras_quadro2.append(("Rateio Gerência", rateio_por_unidade))

    info2 = _escrever_quadro(
        ws, linha_prox_bloco, 1,
        f"{titulo_periodo} — por unidade (com rateio)",
        ORDEM_UNIDADES_ATIVAS, dados_com_rateio, ORDEM_VARIAVEL, ORDEM_FIXO,
    )
    linha_total2, linha_prox_bloco = _finalizar_quadro_com_total(
        ws, info2, linhas_extra_antes_do_total=extras_quadro2
    )

    # --- Check: TOTAL do quadro 1 (coluna TOTAL, linha Total Costs) deve
    # bater com o TOTAL do quadro 2 (coluna TOTAL, linha Total Costs) ---
    col_total1 = get_column_letter(info1["col_inicio"] + 1 + len(info1["colunas"]))
    col_total2 = get_column_letter(info2["col_inicio"] + 1 + len(info2["colunas"]))
    ws.cell(row=linha_prox_bloco + 1, column=1, value="Check (deve ser 0,00):").font = Font(bold=True)
    cc = ws.cell(
        row=linha_prox_bloco + 1, column=2,
        value=f"={col_total1}{linha_total1}-{col_total2}{linha_total2}",
    )
    cc.number_format = "#,##0.00;(#,##0.00)"
    cc.font = Font(bold=True)

    for col in range(1, 8):
        ws.column_dimensions[get_column_letter(col)].width = 16 if col > 1 else 22

    pasta_saida.mkdir(parents=True, exist_ok=True)
    nome_arquivo = nome_com_versao(pasta_saida, f"Rateio de Custos Fitted Units {MESES_INGLES[mes]} {ciclo} {ano}.xlsx")
    caminho_saida = pasta_saida / nome_arquivo
    wb.save(caminho_saida)
    log(f"Arquivo de Rateio de Custos gerado: {caminho_saida}")
    return caminho_saida


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Gera o arquivo de Rateio de Custos (Passo 5)")
    parser.add_argument("--mes", type=int, required=True)
    parser.add_argument("--ano", type=int, required=True)
    parser.add_argument("--ciclo", choices=["Actual", "Flash"], required=True)
    parser.add_argument("--pasta-saida", type=Path, required=True, help="Pasta de saída (teste local ou rede)")
    args = parser.parse_args()

    gerar_arquivo_rateio_custos(args.mes, args.ano, args.ciclo, args.pasta_saida)
