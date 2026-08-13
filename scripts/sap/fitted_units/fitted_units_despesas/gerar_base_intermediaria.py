#!/usr/bin/env python3
"""
Gera os valores do mes para colar na aba "Intermediaria" do arquivo
"Base Intermediaria Fitted <Mes> <Ciclo> <Ano>.xlsx" (Fitted Units Despesas),
direto dos extratos brutos da KSB1 (Gestoriais / Sem Agrupamento) — sem passar
pelo arquivo acumulado BASE_KSB1 (~45 mil linhas) nem mexer em Tabela Dinamica.

Por que da pra pular o BASE_KSB1 (confirmado com a Juliana em 2026-08-11):
o BASE_KSB1 so tem dois usos — alimentar a Base Intermediaria e fazer o
"check do agrupamento" (conferir se toda conta esta vinculada a um
gestorial). O segundo ja e feito hoje direto dos extratos brutos por
check_agrupamentos_ksb1.py, sem depender do BASE_KSB1. E a classificacao
Fixo/Variavel de cada (Conta Fiscal, Centro de Custo) ja esta gravada na
propria aba Intermediaria (coluna H) — nao precisa ser recalculada. Entao so
falta somar o Valor/MR do mes, agrupado por (Centro de Custo, Conta Fiscal),
e casar com a linha correspondente da Intermediaria.

Por seguranca, antes de gerar os valores este script exige que o "Check de
agrupamentos" do mes exista (gera se ainda nao existir) — e essa checagem
que garante o vinculo/classificacao de cada conta, papel que antes dependia
do BASE_KSB1.

Nunca sobrescreve o arquivo de trabalho: sempre salva uma copia nova
("... - gerado.xlsx", versionada com nome_com_versao) com os valores do mes
preenchidos e uma aba extra "Pendencias" para o que nao pode ser casado
automaticamente (chave ambigua ou sem linha correspondente).
"""
from pathlib import Path

from openpyxl import load_workbook

from atualizar_ksb1_gui import BU, MESES_PASTA, REDE_BASE, nome_com_versao
from check_agrupamentos_ksb1 import (
    eh_conta_ignorada,
    encontrar_arquivo,
    linha_e_subtotal,
)

MESES_INGLES = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December",
}

COL_CENTRO = 3  # "Centro custo"
COL_CONTA = 4   # "Classe de custo"
COL_VALOR = 17  # "Valor/MR"

COL_INTER_GESTORIAL = 1
COL_INTER_CONTA_FISCAL = 3
COL_INTER_CENTRO = 5


def _normalizar_centro(v) -> str:
    s = str(v).strip()
    try:
        return str(int(float(s)))
    except ValueError:
        return s


def _normalizar_conta(v) -> str:
    return str(v).strip()


def ler_linhas_centro_conta_valor(caminho: Path):
    """Le (centro, conta, valor) de cada linha de detalhe, pulando subtotais e linhas sem conta."""
    wb = load_workbook(caminho, data_only=True)
    ws = wb.active
    linhas = []
    for r in range(2, ws.max_row + 1):
        conta = ws.cell(row=r, column=COL_CONTA).value
        if conta in (None, ""):
            continue
        if linha_e_subtotal(ws, r):
            continue
        centro = ws.cell(row=r, column=COL_CENTRO).value
        valor = ws.cell(row=r, column=COL_VALOR).value or 0
        linhas.append((_normalizar_centro(centro), _normalizar_conta(conta), float(valor)))
    return linhas


def garantir_check_de_agrupamentos(mes: int, ano: int, log) -> Path:
    """Garante que o Check de agrupamentos do mes existe (gera se nao existir).

    E essa checagem que confirma o vinculo/classificacao gestorial de cada
    conta antes de somarmos qualquer valor — o mesmo papel que o BASE_KSB1
    cumpria antes.
    """
    pasta_mes = REDE_BASE / str(ano) / "00.Extração Base KSB1" / MESES_PASTA[mes]
    try:
        caminho = encontrar_arquivo(pasta_mes, f"Check de agrupamentos - {mes:02d}.{ano}")
        log(f"Check de agrupamentos já existe: {caminho.name}")
        return caminho
    except FileNotFoundError:
        log("Check de agrupamentos ainda não existe para este mês, gerando agora...")
        from check_agrupamentos_ksb1 import gerar_check

        return gerar_check(mes, ano, log=log)


def escolher_linhas_do_mes(mes: int, ano: int, log):
    """Decide Gestoriais vs Sem Agrupamento (regra confirmada em 2026-08-11) e
    retorna a lista (centro, conta, valor) já filtrada, pronta para agrupar."""
    garantir_check_de_agrupamentos(mes, ano, log)

    pasta_mes = REDE_BASE / str(ano) / "00.Extração Base KSB1" / MESES_PASTA[mes]
    arquivo_gest = encontrar_arquivo(pasta_mes, f"KSB1 - {BU['nome']} {mes:02d}.{ano} - Gestoriais")
    arquivo_sem = encontrar_arquivo(pasta_mes, f"KSB1 - {BU['nome']} {mes:02d}.{ano} - Sem Agrupamento")

    linhas_gest = ler_linhas_centro_conta_valor(arquivo_gest)
    linhas_sem = ler_linhas_centro_conta_valor(arquivo_sem)

    total_gest = sum(v for _, _, v in linhas_gest)
    linhas_sem_filtradas = [(c, k, v) for c, k, v in linhas_sem if not eh_conta_ignorada(k)]
    total_sem_filtrado = sum(v for _, _, v in linhas_sem_filtradas)

    if abs(total_gest - total_sem_filtrado) < 0.01:
        log(f"Gestoriais bate com Sem Agrupamento (total={total_gest:,.2f}) — usando Gestoriais.")
        return linhas_gest
    else:
        log(
            f"Gestoriais ({total_gest:,.2f}) NÃO bate com Sem Agrupamento filtrado "
            f"({total_sem_filtrado:,.2f}) — usando Sem Agrupamento (excluindo contas do Check 1), "
            "conforme regra confirmada em 2026-08-11."
        )
        return linhas_sem_filtradas


def agrupar_por_centro_conta(linhas):
    soma = {}
    for centro, conta, valor in linhas:
        chave = (centro, conta)
        soma[chave] = soma.get(chave, 0.0) + valor
    return soma


def localizar_base_intermediaria(mes: int, ano: int, ciclo: str) -> Path:
    mes_pasta_abrev = MESES_PASTA[mes].split(" - ")[1]  # ex: "Jun"
    pasta_ciclo = REDE_BASE / str(ano) / MESES_PASTA[mes] / f"{mes:02d}_{mes_pasta_abrev}_{ciclo}"
    nome = f"Base Intermediária Fitted {MESES_INGLES[mes]} {ciclo} {ano}.xlsx"
    caminho = pasta_ciclo / nome
    if not caminho.exists():
        raise FileNotFoundError(f"Não encontrei '{nome}' em {pasta_ciclo}")
    return caminho


def atualizar_base_intermediaria(mes: int, ano: int, ciclo: str, log=print) -> Path:
    somas = agrupar_por_centro_conta(escolher_linhas_do_mes(mes, ano, log))

    caminho_origem = localizar_base_intermediaria(mes, ano, ciclo)
    log(f"Abrindo {caminho_origem.name}...")
    wb = load_workbook(caminho_origem, data_only=False)
    ws = wb["Intermediária"]

    col_mes = None
    for c in range(1, ws.max_column + 1):
        if str(ws.cell(row=1, column=c).value or "").strip() == MESES_INGLES[mes]:
            col_mes = c
            break
    if col_mes is None:
        raise RuntimeError(f"Não achei a coluna '{MESES_INGLES[mes]}' no cabeçalho da aba Intermediária.")

    linhas_por_chave = {}
    for r in range(2, ws.max_row + 1):
        conta = ws.cell(row=r, column=COL_INTER_CONTA_FISCAL).value
        centro = ws.cell(row=r, column=COL_INTER_CENTRO).value
        if conta in (None, "") or centro in (None, ""):
            continue
        chave = (_normalizar_centro(centro), _normalizar_conta(conta))
        linhas_por_chave.setdefault(chave, []).append(r)

    preenchidas = []
    pendencias = []
    total_extrato = sum(somas.values())
    total_preenchido = 0.0

    for chave, valor in somas.items():
        linhas = linhas_por_chave.get(chave, [])
        centro, conta = chave
        if len(linhas) == 1:
            r = linhas[0]
            valor_antigo = ws.cell(row=r, column=col_mes).value
            ws.cell(row=r, column=col_mes).value = valor
            preenchidas.append((r, centro, conta, valor_antigo, valor))
            total_preenchido += valor
        elif len(linhas) == 0:
            pendencias.append(("SEM LINHA CORRESPONDENTE", centro, conta, valor, ""))
        else:
            pendencias.append(("CHAVE AMBÍGUA", centro, conta, valor, ", ".join(str(x) for x in linhas)))

    ws_pend = wb.create_sheet("Pendências")
    ws_pend.append([f"Pendências — {MESES_INGLES[mes]}/{ano} ({ciclo})"])
    ws_pend.append([])
    ws_pend.append(["Total no extrato do mês", total_extrato])
    ws_pend.append(["Total preenchido automaticamente na Intermediária", total_preenchido])
    ws_pend.append(["Total em pendência (não preenchido)", total_extrato - total_preenchido])
    ws_pend.append([])
    ws_pend.append(["Situação", "Centro de Custo", "Conta Fiscal", "Valor", "Linhas candidatas na Intermediária"])
    for linha in pendencias:
        ws_pend.append(list(linha))

    log(f"\n{len(preenchidas)} linha(s) preenchida(s) automaticamente na coluna {MESES_INGLES[mes]}.")
    if pendencias:
        log(f"AVISO: {len(pendencias)} combinação(ões) em pendência (ver aba Pendências) — total {total_extrato - total_preenchido:,.2f}.")
    else:
        log("Nenhuma pendência — todas as combinações do extrato bateram com uma linha única.")

    pasta_ciclo = caminho_origem.parent
    nome_base = f"Base Intermediária Fitted {MESES_INGLES[mes]} {ciclo} {ano} - gerado.xlsx"
    nome_saida = nome_com_versao(pasta_ciclo, nome_base)
    caminho_saida = pasta_ciclo / nome_saida
    wb.save(caminho_saida)
    log(f"\nArquivo gerado: {caminho_saida}")
    return caminho_saida


if __name__ == "__main__":
    import sys

    if len(sys.argv) != 4:
        print("Uso: python gerar_base_intermediaria.py <mes> <ano> <Actual|Flash>")
        sys.exit(1)
    atualizar_base_intermediaria(int(sys.argv[1]), int(sys.argv[2]), sys.argv[3])
