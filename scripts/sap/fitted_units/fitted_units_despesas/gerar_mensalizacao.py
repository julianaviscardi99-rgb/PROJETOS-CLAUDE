#!/usr/bin/env python3
"""
Passo 6 (Mensalizacao) do processo recorrente (Fitted Units Despesas).

Substitui o arquivo manual "MENS FITTED <Ciclo> <Mes>.xls" (hoje copiado a
mao do Forecast/Actual anterior, com varios ajustes manuais de celula - a
"perfumaria" - e colagem manual dos numeros do Rateio de Custos) por uma
geracao automatica, a partir do Passo 5 (gerar_rateio_custos.py).

Desenhado com a Juliana em 2026-08-26 (ver memory/BRIEFING.md dessa data
pro historico completo da conversa e a explicacao passo a passo). Ainda
SO' O CENARIO FLASH, caso NORMAL (existe Forecast pro mes sendo fechado) -
o caso "sem Forecast" (usar o Actual anterior + a perfumaria do ultimo
Forecast) e o cenario Actual em si ainda nao foram implementados.

Racional confirmado pela usuaria:
1. Localizar o Forecast mais recente pro mes sendo fechado (pasta
   "Fcst\\Fcst <ano>\\R<mes> <ano>\\MENS FITTED FORECAST <MES>.xls" - o
   numero da revisao Rn e' sempre igual ao numero do mes, R1=Refresh
   Janeiro ... R12=Refresh Dezembro, confirmado pela usuaria).
2. Copiar esse arquivo pra pasta de saida (rede oficial seria
   "Flash\\<ano>\\<MM> - <Mes em ingles>\\", mas o script recebe a pasta
   como parametro - nunca hardcoda o destino de producao, ver
   gerar_arquivo_mensalizacao) e renomear trocando "FORECAST" por "FLASH"
   no nome.
3. So' mexe na aba de cada unidade ATIVA (SJP, IBI, GO, RES) + a aba TOTAL -
   as outras abas do arquivo (Action Plan Ibirite, MDO, Rateio Fixo,
   Confronto, Sheet1, Proposta) ficam de fora, confirmado pela usuaria
   ("todas as abas que estao abertas" = so' essas 5).
4. Em cada uma dessas 5 abas ("perfumaria"):
   - E5: texto do Ciclo -> "FLASH"
   - S5: numero da revisao -> "R<mes>" (ex: "R7")
   - C47: mesmo numero de revisao -> "R<mes>"
   - E47:Q47 <- copia de E43:Q43
   - S8:S44 <- cola como VALOR o que esta em Q8:Q44
   So' a aba SJP precisa ser editada direto pra E5/S5/C47 (as outras abas
   puxam esse cenario dela por formula) - MAS linha 47 e S8:S44 precisam
   ser repetidos em cada uma das 5 abas (nao propaga via formula).
5. Na coluna do MES QUE ESTA FECHANDO (so' essa, nunca as outras - as
   colunas dos meses seguintes ja vem como formulas vivas, linkadas aos
   arquivos-fonte do proprio Forecast, e NAO sao tocadas):
   - Linhas 19-23 (Labour/Handling/Direct Materials/Transportation/Other
     Variable) e 32-37 (Labour/Depreciation/IFRS16/Rents/Condominio/Other
     Fixed) recebem os valores do Passo 5 (com rateio), mesma ordem/
     categorias exatas - confirmado, mapeamento 1:1 linha a linha.
   - Linhas 18 (Variable Cost, =SUM), 31 (Fixed Cost, =SUM) e 26
     (Variabile/Pc, =IF(...)) sao FORMULAS e NUNCA sao tocadas - decisao
     explicita da usuaria ("o mais facil e nem colar nessas linhas") -
     recalculam sozinhas a partir das linhas de detalhe acima.
6. Check: TOTAL COST (linha 39, =linha31+linha18) da coluna do mes deve
   bater com o Total Costs por unidade que o Passo 5 ja calcula (mesmo
   racional do check por unidade do Passo 5).

NAO mexe em nada do Passo 5 (gerar_rateio_custos.py) - so' importa as
funcoes de leitura/calculo pra reaproveitar os mesmos numeros ja validados.
"""
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "_shared"))
from ksb1_core import MESES_NOMES, abrir_excel_isolado, nome_com_versao  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
from gerar_rateio_custos import (  # noqa: E402
    MESES_INGLES,
    ORDEM_FIXO,
    ORDEM_VARIAVEL,
    calcular_dados_com_rateio,
    carregar_rateio_vigente,
    ler_e_classificar,
    localizar_base_intermediaria,
)

# Rede oficial de producao (Forecast/Flash/Actual) - NUNCA escrita direto
# por este script (a pasta de saida vem sempre por parametro, ver
# gerar_arquivo_mensalizacao) - so' usada pra LOCALIZAR o Forecast/Actual
# de origem (leitura).
REDE_BASE_MENSALIZACAO = Path(
    r"\\FSS024-01BR.group.pirelli.com\EO_FITTED\BU FITTED\Forecast"
)

# Aba do arquivo de Mensalizacao -> sigla usada no Passo 5 (gerar_rateio_
# custos.py). "GO" na Mensalizacao = "GOI" no Passo 5 (nomes diferentes pro
# mesmo Goiana) - confirmado pela usuaria, 2026-08-26.
ABAS_UNIDADES = {
    "SJP": "SJP",
    "IBI": "IBI",
    "GO": "GOI",
    "RES": "RES",
}
ABA_TOTAL = "TOTAL"
ABAS_MENSALIZACAO = list(ABAS_UNIDADES) + [ABA_TOTAL]

# Linhas de detalhe (Excel, 1-based) - mesma ORDEM de gerar_rateio_custos.
# ORDEM_VARIAVEL/ORDEM_FIXO, mapeamento 1:1 confirmado pela usuaria.
LINHA_INICIO_VARIAVEL = 19  # Labour, Handling, Direct Materials, Transportation, Other Variable
LINHA_INICIO_FIXO = 32      # Labour, Depreciation, IFRS16, Rents, Condominio, Other Fixed
LINHA_TOTAL_COST = 39

COL_MES_INICIAL = 4  # coluna E (=5) e' Janeiro -> coluna = COL_MES_INICIAL + mes


def coluna_do_mes(mes: int) -> int:
    return COL_MES_INICIAL + mes


def _achar_arquivo_por_prefixo(pasta: Path, prefixo: str, log=print) -> Path | None:
    """Acha o único arquivo '.xls' na pasta cujo nome começa com `prefixo` -
    NÃO monta o nome do mês (ex: 'MENS FITTED FLASH JULHO.xls') porque o
    nome do mês NÃO é consistente entre os arquivos reais (achado real,
    2026-08-26: Jan-Jun usam inglês - 'JANUARY', 'FEBRUARY', 'MARCH', até
    abreviado - 'APR' - e só Julho usa português 'JULHO'). Avisa (não
    levanta erro) se achar mais de um candidato - quem chama decide."""
    if not pasta.exists():
        return None
    candidatos = sorted(pasta.glob(f"{prefixo}*.xls"))
    if not candidatos:
        return None
    if len(candidatos) > 1:
        log(f"AVISO: mais de um arquivo '{prefixo}*.xls' em {pasta} - usando o primeiro: {candidatos}")
    return candidatos[0]


def localizar_forecast_do_mes(mes: int, ano: int, log=print) -> Path | None:
    """Devolve o caminho do arquivo Forecast pro mes/ano pedido, se a
    revisao Rn (n=mes) existir - None se ainda nao existir (ex: fechamento
    de Agosto/2026, so' existe ate R7)."""
    pasta = REDE_BASE_MENSALIZACAO / "Fcst" / f"Fcst {ano}" / f"R{mes} {ano}"
    return _achar_arquivo_por_prefixo(pasta, "MENS FITTED FORECAST ", log)


def localizar_forecast_mais_recente(mes: int, ano: int, log=print) -> tuple[Path, int] | None:
    """Devolve (caminho, revisao) do Forecast mais recente ANTERIOR ou
    igual ao mes pedido (usado pra buscar a 'perfumaria' quando o mes
    sendo fechado ainda nao tem Forecast proprio - ex: Agosto/2026 usa a
    revisao R7). Percorre de tras pra frente (mes, mes-1, ...) ate achar."""
    for m in range(mes, 0, -1):
        caminho = localizar_forecast_do_mes(m, ano, log)
        if caminho is not None:
            return caminho, m
    return None


def localizar_actual_do_mes(mes: int, ano: int, log=print) -> Path | None:
    # As pastas de Actual/Flash usam o mes por extenso em INGLES (ex:
    # "07 - July"), diferente das pastas do Passo 1-5 (GFU_DAC, que usam
    # abreviação em inglês, "07 - Jul", MESES_PASTA) - confirmado
    # explorando a rede em 2026-08-26. O NOME DO ARQUIVO dentro da pasta
    # não é consistente (ver _achar_arquivo_por_prefixo) - por isso busca
    # por prefixo, não monta o nome do mês.
    pasta = REDE_BASE_MENSALIZACAO / "Actual" / str(ano) / f"{mes:02d} - {MESES_INGLES[mes]}"
    return _achar_arquivo_por_prefixo(pasta, "MENS FITTED ACTUAL ", log)


def localizar_flash_do_mes(mes: int, ano: int, log=print) -> Path | None:
    """Devolve o caminho do arquivo Flash já fechado pro mes/ano pedido -
    é a base pro Ciclo Actual (confirmado pela usuária, 2026-08-26: 'o
    arquivo base que você deve pegar é do flash pra continuar com o
    actual... a perfumaria já deve estar toda ok, porque você já arrumou
    no flash')."""
    pasta = REDE_BASE_MENSALIZACAO / "Flash" / str(ano) / f"{mes:02d} - {MESES_INGLES[mes]}"
    return _achar_arquivo_por_prefixo(pasta, "MENS FITTED FLASH ", log)


def determinar_fonte(mes: int, ano: int, log) -> dict:
    """Decide de onde vem a base (numeros) e de onde vem a perfumaria
    (cenario de comparacao), pro mes/ano/Flash pedido:
    - Caso normal: existe Forecast R<mes> -> base E perfumaria vem dele.
    - Caso sem Forecast: usa o Actual do mes anterior como base, mas a
      perfumaria continua vindo do Forecast mais recente disponivel
      (confirmado pela usuaria, 2026-08-26, com o exemplo real de Agosto/
      2025: 'MENS FITTED 2025 FLASH AGO com JUL EFETIVO.xls', S5='R7 JUL act').
    """
    forecast_do_mes = localizar_forecast_do_mes(mes, ano, log)
    if forecast_do_mes is not None:
        log(f"Forecast R{mes}/{ano} encontrado - caso normal (base e perfumaria vêm dele).")
        return {"caso": "normal", "base": forecast_do_mes, "revisao_perfumaria": mes}

    mes_anterior, ano_anterior = (12, ano - 1) if mes == 1 else (mes - 1, ano)
    actual_anterior = localizar_actual_do_mes(mes_anterior, ano_anterior, log)
    if actual_anterior is None:
        raise FileNotFoundError(
            f"Não achei nem o Forecast R{mes}/{ano} nem o Actual de "
            f"{MESES_NOMES[mes_anterior]}/{ano_anterior} — preciso de um dos dois "
            "pra montar a Mensalização."
        )
    ultimo = localizar_forecast_mais_recente(mes_anterior, ano_anterior if mes == 1 else ano)
    if ultimo is None:
        raise FileNotFoundError(
            f"Achei o Actual de {MESES_NOMES[mes_anterior]}/{ano_anterior}, mas não achei "
            "nenhum Forecast anterior pra usar como perfumaria (cenário de comparação)."
        )
    caminho_forecast, revisao = ultimo
    log(
        f"Forecast R{mes}/{ano} NÃO existe ainda - caso sem Forecast: base = Actual "
        f"{MESES_NOMES[mes_anterior]}/{ano_anterior}, perfumaria = R{revisao}."
    )
    return {"caso": "sem_forecast", "base": actual_anterior, "revisao_perfumaria": revisao}


def _copiar_e_renomear(fonte: Path, pasta_saida: Path, mes: int, ano: int, ciclo: str, log) -> Path:
    pasta_saida.mkdir(parents=True, exist_ok=True)
    nome_novo = nome_com_versao(pasta_saida, f"MENS FITTED {ciclo.upper()} {MESES_NOMES[mes].upper()}.xls")
    caminho_novo = pasta_saida / nome_novo
    shutil.copy2(fonte, caminho_novo)
    log(f"Copiado {fonte.name} -> {caminho_novo}")
    return caminho_novo


def _trocar_flash_por_actual(wb, log):
    """Ciclo Actual: a perfumaria já está toda certa (herdada do Flash do
    mesmo mês, que já foi fechado e corrigido antes) - só precisa trocar
    todo texto 'Flash' por 'Actual' nas 5 abas (pedido explícito da
    usuária, 2026-08-26: 'tudo o que estiver escrito flash, vira actual').
    Usa Find/Replace do próprio Excel (LookAt=xlWhole, só troca célula cujo
    conteúdo INTEIRO é 'Flash' - não mexe em texto que só contém a palavra
    como parte de algo maior, nem em fórmulas com link externo)."""
    xlWhole = 1
    for aba in ABAS_MENSALIZACAO:
        ws = wb.Worksheets(aba)
        trocou = ws.Cells.Replace(
            What="Flash", Replacement="Actual", LookAt=xlWhole, MatchCase=False
        )
        log(f"{aba}: texto 'Flash' trocado por 'Actual' onde encontrado.")


def _aplicar_perfumaria(wb, mes: int, revisao_perfumaria: int, sufixo_texto: str, log):
    """Aplica as edicoes 'cosmeticas' (E5/S5/C47/linha47/Q->S) nas 5 abas
    (SJP/IBI/GO/RES/TOTAL). So' a aba SJP recebe E5/S5/C47 diretamente (as
    outras puxam por formula) - linha 47 e S8:S44 sao repetidos em todas."""
    texto_revisao = f"R{revisao_perfumaria}{sufixo_texto}"
    ws_sjp = wb.Worksheets("SJP")
    ws_sjp.Range("E5").Value = "FLASH"
    ws_sjp.Range("S5").Value = texto_revisao
    ws_sjp.Range("C47").Value = f"R{revisao_perfumaria}"
    log(f"SJP: E5=FLASH, S5={texto_revisao}, C47=R{revisao_perfumaria}")

    for aba in ABAS_MENSALIZACAO:
        ws = wb.Worksheets(aba)
        ws.Range("E47:Q47").Value = ws.Range("E43:Q43").Value
        valores_q = ws.Range("Q8:Q44").Value
        ws.Range("S8:S44").Value = valores_q
        log(f"{aba}: linha 47 <- linha 43, S8:S44 <- Q8:Q44 (colado como valor)")


def _colar_valores_rateio(wb, mes: int, dados_por_unidade: dict, log) -> dict:
    """Cola os valores do Passo 5 (com rateio) nas linhas de detalhe
    (19-23 Variavel, 32-37 Fixo) da coluna do mes sendo fechado, em cada
    uma das 5 abas. NUNCA toca nas linhas 18/31/26 (formulas) - so' os
    detalhes, que ja alimentam essas formulas. Devolve {aba: total_colado}
    pro check final (comparar com TOTAL COST da planilha).

    Convenção de sinal: o Passo 5 guarda custo como NEGATIVO (mesma
    convenção do arquivo antigo '_Abertura custos...'), mas o arquivo de
    Mensalização guarda custo como POSITIVO (confirmado lendo o arquivo
    real - linha 39 'TOTAL COST' mostra 1.561,08, não -1.561,08) - por
    isso inverte o sinal aqui. Achado real testando contra a rede,
    2026-08-26: sem essa inversão o Check batia (comparando negativo com
    negativo), mas o valor gravado na planilha ficava com o sinal errado."""
    col = coluna_do_mes(mes)
    totais_colados = {}
    for aba in ABAS_MENSALIZACAO:
        sigla_rateio = ABAS_UNIDADES.get(aba, "TOTAL")
        dados_unidade = dados_por_unidade.get(sigla_rateio, {})
        ws = wb.Worksheets(aba)
        total = 0.0
        for i, item in enumerate(ORDEM_VARIAVEL):
            valor = -dados_unidade.get(("V", item), 0)
            ws.Cells(LINHA_INICIO_VARIAVEL + i, col).Value = valor
            total += valor
        for i, item in enumerate(ORDEM_FIXO):
            valor = -dados_unidade.get(("F", item), 0)
            ws.Cells(LINHA_INICIO_FIXO + i, col).Value = valor
            total += valor
        totais_colados[aba] = total
        log(f"{aba}: valores do Passo 5 colados na coluna do mês (total {total:,.1f}).")
    return totais_colados


def _conferir_total_cost(wb, mes: int, totais_colados: dict, log) -> dict:
    """Le o TOTAL COST (formula, linha 39) da coluna do mes, ja recalculado
    pelo proprio Excel a partir dos valores colados, e compara com o total
    que o Passo 5 devolveu pra essa unidade - devolve {aba: (excel, passo5,
    diff)}."""
    col = coluna_do_mes(mes)
    checks = {}
    for aba in ABAS_MENSALIZACAO:
        ws = wb.Worksheets(aba)
        valor_excel = ws.Cells(LINHA_TOTAL_COST, col).Value or 0.0
        valor_passo5 = totais_colados.get(aba, 0.0)
        diff = valor_excel - valor_passo5
        checks[aba] = (valor_excel, valor_passo5, diff)
        status = "OK" if abs(diff) < 0.05 else "DIFERENÇA"
        log(f"Check {aba}: TOTAL COST Excel={valor_excel:,.1f}  Passo5={valor_passo5:,.1f}  [{status}]")
    return checks


def gerar_arquivo_mensalizacao(
    mes: int, ano: int, pasta_saida: Path, ciclo: str = "Flash", log=print, pid_callback=None
) -> tuple[Path, dict]:
    """Gera o arquivo de Mensalização pro mês/ano/Ciclo pedido, salvando em
    `pasta_saida` (NUNCA a pasta oficial de rede direto - quem chama decide
    isso, mesmo racional do Passo 5). Devolve (caminho_gerado, checks) -
    checks = {aba: (total_excel, total_passo5, diferença)}.

    Ciclo Flash: base = Forecast do mês (ou Actual anterior + perfumaria do
    último Forecast, se ainda não existir Forecast pro mês) - ver
    determinar_fonte. Aplica a "perfumaria" completa (E5/S5/C47/linha47/
    S8:S44) nas 5 abas.

    Ciclo Actual: base = o Flash DO MESMO MÊS, já fechado - confirmado pela
    usuária, 2026-08-26 ("o arquivo base que você deve pegar é do flash
    pra continuar com o actual"). A perfumaria NÃO é refeita (já veio certa
    do Flash) - só troca todo texto "Flash" por "Actual" nas 5 abas."""
    pasta_saida = pasta_saida.resolve()

    caminho_base_intermediaria = localizar_base_intermediaria(mes, ano, ciclo)
    log(f"Lendo Base Intermediária (Passo 5, Ciclo {ciclo}): {caminho_base_intermediaria.name}...")
    totais_ativos, residuos, _, _, _ = ler_e_classificar(caminho_base_intermediaria, mes, log)
    percentuais, vigente_desde = carregar_rateio_vigente(mes, ano)
    log(f"Rateio vigente desde {vigente_desde}: {percentuais}")
    dados_com_rateio = calcular_dados_com_rateio(totais_ativos, percentuais)

    dados_por_unidade = dict(dados_com_rateio)
    dados_por_unidade["TOTAL"] = {}
    for sigla, dados in dados_com_rateio.items():
        for chave, valor in dados.items():
            dados_por_unidade["TOTAL"][chave] = dados_por_unidade["TOTAL"].get(chave, 0) + valor

    if ciclo == "Actual":
        base = localizar_flash_do_mes(mes, ano, log)
        if base is None:
            raise FileNotFoundError(
                f"Não achei o Flash de {MESES_NOMES[mes]}/{ano} já fechado - preciso dele "
                "como base pra montar o Actual (a perfumaria vem de lá)."
            )
        log(f"Ciclo Actual: base = Flash {MESES_NOMES[mes]}/{ano} já fechado ({base.name}).")
    else:
        fonte = determinar_fonte(mes, ano, log)
        base = fonte["base"]

    caminho_novo = _copiar_e_renomear(base, pasta_saida, mes, ano, ciclo, log)

    excel = abrir_excel_isolado(log, pid_callback)
    excel.AskToUpdateLinks = False
    try:
        wb = excel.Workbooks.Open(str(caminho_novo.resolve()))
        if ciclo == "Actual":
            _trocar_flash_por_actual(wb, log)
        else:
            sufixo_texto = f" {MESES_NOMES[mes][:3].upper()} act" if fonte["caso"] == "sem_forecast" else ""
            _aplicar_perfumaria(wb, mes, fonte["revisao_perfumaria"], sufixo_texto, log)
        totais_colados = _colar_valores_rateio(wb, mes, dados_por_unidade, log)
        wb.Save()
        checks = _conferir_total_cost(wb, mes, totais_colados, log)
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()

    log(f"\nArquivo de Mensalização gerado: {caminho_novo}")
    return caminho_novo, checks


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Gera o arquivo de Mensalização (Passo 6)")
    parser.add_argument("--mes", type=int, required=True)
    parser.add_argument("--ano", type=int, required=True)
    parser.add_argument("--ciclo", choices=["Flash", "Actual"], default="Flash")
    parser.add_argument("--pasta-saida", type=Path, required=True, help="Pasta de saída (teste local ou rede)")
    args = parser.parse_args()

    gerar_arquivo_mensalizacao(args.mes, args.ano, args.pasta_saida, args.ciclo)
