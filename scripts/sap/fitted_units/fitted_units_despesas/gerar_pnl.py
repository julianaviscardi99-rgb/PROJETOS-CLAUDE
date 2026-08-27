#!/usr/bin/env python3
"""
Passo 7 (P&L) do processo recorrente (Fitted Units Despesas).

Substitui a criação manual do arquivo `<MM>_P&L Fitted Units_<Ciclo>_<Mês>-
<AA>.xlsx` (hoje: copiar o arquivo do mesmo Ciclo do mês anterior e trocar
"na mão" os links externos + alguns textos) por uma geração automática.

Desenhado com a Juliana em 2026-08-27 (ver memory/BRIEFING.md e
memory/DECISOES.md dessa data). Mecanismo derivado por COMPARAÇÃO CÉLULA A
CÉLULA entre os arquivos reais de Maio->Junho e Junho->Julho/2026 (ambos os
Ciclos) - não é suposição, é o diff exato do que a usuária edita todo mês
(salvo em data/processed/fitted_units_despesas/pnl_teste/, fora do Git).

v1 (2026-08-27): cobre Actual e Flash. NÃO gera a cópia "congelada" (nome
com "_" no final, só valor) - fica pra uma 2ª rodada, confirmado com a
usuária. NÃO cobre a virada de ano (Dezembro -> Janeiro) - o bloco Jan-Dez
da aba "Resumo Resultado Ano" provavelmente precisa de um reset bem maior
nesse caso, que a usuária ainda não detalhou; rodar isso só a partir de
Fevereiro caso o fechamento de Janeiro seja feito manualmente.

Racional confirmado (ver DECISOES.md 2026-08-27 pra detalhe completo):
1. Ponto de partida = sempre uma cópia do arquivo do MESMO Ciclo do mês
   anterior (nunca template em branco) - vale igual pra Actual e Flash.
2. Os links externos que TROCAM TODO MÊS (mecanismo: Workbook.ChangeLink,
   mantém o link vivo, nunca converte pra valor):
   - Mensalização (sempre) -> `MENS FITTED <CICLO> <Mês-PT>.xls`, saída do
     Passo 6, em EO_FITTED\\BU FITTED\\Forecast\\<Ciclo>\\<ano>\\<MM> - <Mês
     em inglês>\\. NUNCA usar EO_CONSUMER (achado real 2026-08-27: arquivos
     de Maio/Junho usam EO_CONSUMER, mas a usuária confirmou que
     EO_FITTED é o correto/atual - ver memory/errors/2026-08-27_pnl_link_
     flash_forecast_pasta_faltando.md).
   - Forecast (sempre) -> `<MM>_P&L Fitted Units_Forecast_...xlsx` do mês
     fechando; se ainda não existir (mês muito recente), cai pro Forecast
     do mês anterior (mesmo fallback R8->R7 já usado no Passo 6).
   - Flash (só Ciclo Actual) -> o P&L Flash DO MESMO MÊS, já fechado.
3. Os links que só trocam em JANEIRO (fora do escopo de teste desta v1,
   mas já implementados): PY (Actual de Dezembro do ano anterior) e MP
   (Management Plan do ano que está sendo fechado, não do anterior).
4. SEMPRE monta o caminho de destino do ChangeLink com o nível "MM - Mês"
   incluído (nunca replica os caminhos curtos/quebrados que os arquivos
   atuais têm - achado real 2026-08-27, ver memory/errors/ acima).
5. Textos que mudam (linha 4/5 de "Resumo Resultado Ano" e "Resultado
   YTD", linha 4/5 de "Resumo Resultado Mês") - ver `atualizar_textos`.
6. "Resumo Resultado Mês": colunas D/E/F/G(/H) NÃO são link externo (exceto
   a coluna Flash, no Actual) - são referência deslizante dentro do próprio
   arquivo pra 'Resumo Resultado Ano' (reescrita direta da fórmula).
7. "Resultado YTD": ganha uma coluna nova a cada mês, copiada (Copy +
   PasteSpecial de fórmulas) da coluna do mês anterior - o próprio Excel
   desliza as referências relativas, replicando exatamente o que a usuária
   descreveu ("copia a coluna de fórmula do mês anterior").
"""
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "_shared"))
from ksb1_core import (  # noqa: E402
    MESES_NOMES,
    MESES_PASTA,
    REDE_BASE,
    abrir_excel_isolado,
    nome_com_versao,
    resolver_pasta_ciclo,
)

sys.path.insert(0, str(Path(__file__).resolve().parent))
from gerar_rateio_custos import MESES_ABREV, MESES_INGLES  # noqa: E402  (ex: "July" / "Jul")

# Rede oficial da Mensalizacao (saida do Passo 6) - so' usada pra montar o
# link de destino, nunca escrita por este script.
REDE_MENSALIZACAO = Path(r"\\FSS024-01BR.group.pirelli.com\EO_FITTED\BU FITTED\Forecast")

REDE_MP = Path(r"\\FSS024-01BR.group.pirelli.com\GFU_DAC\Management Plan")

# Blocos de 12 colunas (Jan=coluna inicial) da aba "Resumo Resultado Ano" -
# e replicados na mesma posicao em "Resultado YTD" (achado real, 2026-08-27:
# as duas abas compartilham o mesmo layout de colunas largas). Descobertos
# por comparacao celula a celula Mai->Jun e Jun->Jul/2026 reais, NAO e'
# suposicao - ver docstring do modulo.
COL_BLOCO_MES = 4    # D:O -> valor do proprio mes (Actual/Flash/Forecast)
COL_BLOCO_FORECAST = 19   # S:AD -> Forecast completo do ano (pull-through)
COL_BLOCO_MP = 34    # AH:AS -> MP/Budget completo do ano
COL_BLOCO_PY = 49    # AW:BH -> PY (ano anterior) completo

# Linhas de detalhe (Excel, 1-based) que tem ponteiro deslizante em "Resumo
# Resultado Mes" (colunas D e as de comparacao) - mesmo conjunto pros dois
# Ciclos, confirmado no diff Jun->Jul E Mai->Jun (nenhuma linha extra ou
# faltando entre as duas transicoes).
LINHAS_RESUMO_MES = [8, 13, 14, 15, 20, 21, 22, 23, 24, 31, 32, 33, 34, 35, 36, 41, 42]

# Ultima linha com conteudo real na aba "Resultado YTD" (confirmado no
# arquivo real - linhas com formula vao ate 58, resto e' vazio/rotulo).
ULTIMA_LINHA_YTD = 60


def col_letra(col: int) -> str:
    s = ""
    while col > 0:
        col, r = divmod(col - 1, 26)
        s = chr(65 + r) + s
    return s


def nome_arquivo_pnl(mes: int, ano: int, ciclo: str) -> str:
    return f"{mes:02d}_P&L Fitted Units_{ciclo}_{MESES_INGLES[mes]}-{ano % 100:02d}.xlsx"


def mes_anterior(mes: int, ano: int) -> tuple[int, int]:
    return (12, ano - 1) if mes == 1 else (mes - 1, ano)


def _localizar_versao_com_formula(pasta: Path, nome_base: str) -> Path | None:
    """Acha a versão mais recente (nome_com_versao: sem sufixo, depois
    '_v2', '_v3'...) de um arquivo - NUNCA a cópia 'congelada' (sufixo '_'
    puro, só valor, sem fórmula/link) que o processo manual também salva
    na mesma pasta com um nome quase igual. Achado real testando este
    script, 2026-08-27: um glob genérico (`encontrar_arquivo_mais_recente`)
    casava com as duas e podia pegar a congelada por engano (sem links
    externos pra trocar)."""
    base = Path(nome_base)
    stem, ext = base.stem, base.suffix
    melhor = pasta / nome_base if (pasta / nome_base).exists() else None
    versao = 2
    while (pasta / f"{stem}_v{versao}{ext}").exists():
        melhor = pasta / f"{stem}_v{versao}{ext}"
        versao += 1
    return melhor


def localizar_pnl(mes: int, ano: int, ciclo: str) -> Path:
    pasta = resolver_pasta_ciclo(REDE_BASE / str(ano) / MESES_PASTA[mes], mes, ciclo)
    caminho = _localizar_versao_com_formula(pasta, nome_arquivo_pnl(mes, ano, ciclo))
    if caminho is None:
        raise FileNotFoundError(
            f"Não encontrei o P&L de {ciclo} de {MESES_NOMES[mes]}/{ano} em {pasta} "
            "(preciso dele pra montar o mês seguinte)."
        )
    return caminho


def _achar_arquivo_por_prefixo(pasta: Path, prefixo: str) -> Path | None:
    """Acha o P&L de Forecast pra linkar - prefere a versão COM fórmula
    viva (nome sem '_' no final) sobre a 'congelada' quando as duas
    existem, por ser o que o arquivo real de Julho/2026 usa (validado
    2026-08-27, comparando contra o arquivo real - ver `data/processed/
    fitted_units_despesas/pnl_teste/`). Meses mais antigos (Maio/Junho)
    usam a congelada - se isso for de propósito, ajustar aqui."""
    if not pasta.exists():
        return None
    candidatos = sorted(pasta.glob(f"{prefixo}*.xlsx"), key=lambda p: p.stat().st_mtime)
    if not candidatos:
        return None
    vivos = [c for c in candidatos if not c.stem.endswith("_")]
    return vivos[-1] if vivos else candidatos[-1]


def localizar_forecast_pnl(mes: int, ano: int, log=print) -> tuple[Path, int]:
    """Acha o P&L de Forecast do mes/ano pedido pra usar como link de
    comparacao. O NOME desses arquivos varia mes a mes (achado real,
    2026-08-27: Maio usa sufixo '_', Junho usa '_v2_', Julho não usa
    sufixo nenhum) - por isso busca por PREFIXO (glob), igual ao Passo 6
    faz pro Forecast da Mensalizacao, e pega o mais recente por data de
    modificacao entre os candidatos. Se o mes pedido ainda nao tiver
    Forecast (fechamento muito no inicio do mes), cai pro mes anterior -
    mesmo fallback R8->R7 ja usado no Passo 6 - e devolve a revisao
    realmente usada (pra rotular 'Forecast R<n>')."""
    for m in range(mes, 0, -1):
        pasta = resolver_pasta_ciclo(REDE_BASE / str(ano) / MESES_PASTA[m], m, "Forecast")
        caminho = _achar_arquivo_por_prefixo(pasta, f"{m:02d}_P&L Fitted Units_Forecast")
        if caminho is not None:
            if m != mes:
                log(f"Forecast R{mes} ainda não existe - usando R{m} ({caminho.name}) como fallback.")
            return caminho, m
    raise FileNotFoundError(
        f"Não encontrei nenhum P&L de Forecast (nem R{mes} nem anterior) pra {ano}."
    )


def _construir_link_mensalizacao(mes: int, ano: int, ciclo: str) -> str:
    pasta = REDE_MENSALIZACAO / ciclo / str(ano) / f"{mes:02d} - {MESES_INGLES[mes]}"
    nome = f"MENS FITTED {ciclo.upper()} {MESES_NOMES[mes].upper()}.xls"
    caminho = pasta / nome
    if not caminho.exists():
        raise FileNotFoundError(
            f"Não achei a Mensalização de {ciclo} de {MESES_NOMES[mes]}/{ano} em {caminho} "
            "- rode o Passo 6 pra esse mês/Ciclo antes do P&L."
        )
    return str(caminho)


def _construir_link_mp(ano: int) -> str:
    return str(REDE_MP / f"MP {ano}" / f"P&L Fitted Units_Budget{ano % 100:02d}.xlsx")


def _identificar_link(links: list[str], marcador: str) -> str:
    candidatos = [l for l in links if marcador.lower() in l.lower()]
    if not candidatos:
        raise RuntimeError(f"Não achei nenhum link externo contendo '{marcador}' no arquivo.")
    if len(candidatos) > 1:
        raise RuntimeError(f"Mais de um link contendo '{marcador}': {candidatos}")
    return candidatos[0]


def atualizar_links(wb, mes: int, ano: int, ciclo: str, log) -> int:
    """Troca os links externos que precisam trocar (ver docstring do
    módulo) - devolve a revisão de Forecast realmente usada (pra rotular
    'Forecast R<n>')."""
    links = list(wb.LinkSources(1) or [])

    old_mens = _identificar_link(links, "MENS FITTED")
    novo_mens = _construir_link_mensalizacao(mes, ano, ciclo)
    wb.ChangeLink(old_mens, novo_mens, 1)
    log(f"Link Mensalização -> {novo_mens}")

    old_forecast = _identificar_link(links, "P&L Fitted Units_Forecast")
    caminho_forecast, revisao = localizar_forecast_pnl(mes, ano, log)
    wb.ChangeLink(old_forecast, str(caminho_forecast), 1)
    log(f"Link Forecast -> {caminho_forecast}")

    if ciclo == "Actual":
        old_flash = _identificar_link(links, "P&L Fitted Units_Flash")
        novo_flash = localizar_pnl(mes, ano, "Flash")
        wb.ChangeLink(old_flash, str(novo_flash), 1)
        log(f"Link Flash -> {novo_flash}")

    if mes == 1:
        old_py = _identificar_link(links, "P&L Fitted Units_Actual_December")
        novo_py = localizar_pnl(12, ano - 1, "Actual")
        wb.ChangeLink(old_py, str(novo_py), 1)
        log(f"Link PY -> {novo_py}")

        old_mp = _identificar_link(links, "Management Plan")
        novo_mp = _construir_link_mp(ano)
        wb.ChangeLink(old_mp, novo_mp, 1)
        log(f"Link MP -> {novo_mp}")

    wb.Application.CalculateFullRebuild()
    return revisao


def atualizar_coluna_mes_corrente_flash(wb, mes: int, ano: int, log):
    """Ciclo Flash apenas: a coluna do MÊS QUE ESTÁ FECHANDO, no bloco D:O
    de 'Resumo Resultado Ano', muda de FONTE - deixa de puxar do Forecast
    (como fazia quando ainda era um mês futuro, no arquivo do mês anterior)
    e passa a puxar direto da Mensalização Flash do mês. `ChangeLink` não
    serve aqui: ele trocaria TODAS as células que referenciam o link antigo
    de uma vez (inclusive as dos meses já fechados, que devem continuar
    apontando pro Mensalização do MÊS DELES, não do mês atual).

    Mapeamento linha a linha (Ano -> Mensalização TOTAL) não é 1:1 nem
    uniforme (achado real comparando Jun->Jul/2026: linha 13 do Ano mapeia
    pra linha 11 da Mensalização, algumas linhas têm sinal invertido) - em
    vez de tentar recriar essa tabela na mão, copia o padrão exato já
    existente na coluna do MÊS ANTERIOR (que, no arquivo-base copiado,
    ainda tem as fórmulas corretas apontando pra Mensalização do mês
    anterior - o próprio Excel ajusta a coluna relativa ao copiar) e só
    troca o nome do arquivo (mês anterior -> mês atual) via Replace
    escopado a essa coluna."""
    col_novo = col_letra(COL_BLOCO_MES + mes - 1)
    col_anterior = col_letra(COL_BLOCO_MES + mes - 2)
    nome_antigo = f"MENS FITTED FLASH {MESES_NOMES[mes - 1].upper()}.xls"
    nome_novo = f"MENS FITTED FLASH {MESES_NOMES[mes].upper()}.xls"

    ws = wb.Worksheets("Resumo Resultado Ano")
    origem = ws.Range(f"{col_anterior}1:{col_anterior}{ULTIMA_LINHA_YTD}")
    destino = ws.Range(f"{col_novo}1:{col_novo}{ULTIMA_LINHA_YTD}")
    destino.FormulaR1C1 = origem.FormulaR1C1
    destino.Replace(What=nome_antigo, Replacement=nome_novo, LookAt=2, MatchCase=False)  # xlPart
    log(f"Resumo Resultado Ano: coluna {col_novo} (mês fechando) agora puxa de '{nome_novo}'.")


def atualizar_textos(wb, mes: int, ano: int, ciclo: str, revisao_forecast: int, log):
    """Atualiza os rótulos de texto que mudam todo mês (achado por diff
    real Mai->Jun e Jun->Jul/2026, ambos os Ciclos - ver docstring)."""
    col_mes = COL_BLOCO_MES + mes - 1
    col_mes_anterior = COL_BLOCO_MES + mes - 2
    col_forecast_lag = COL_BLOCO_FORECAST + mes - 2

    # Achado validando contra o arquivo real de Julho (2026-08-27): a aba
    # "Resultado YTD" só carrega esses rótulos de cenário (linha 5) no
    # Ciclo Flash - no Actual ela fica em branco (não é replicada da
    # "Resumo Resultado Ano" como eu supunha antes de validar).
    abas_com_rotulo_linha5 = ("Resumo Resultado Ano",) if ciclo == "Actual" else (
        "Resumo Resultado Ano", "Resultado YTD"
    )
    for aba in abas_com_rotulo_linha5:
        ws = wb.Worksheets(aba)
        ws.Cells(5, col_mes).Value = "Actual" if ciclo == "Actual" else "Flash"
        if ciclo == "Flash" and mes >= 2:
            ws.Cells(5, col_mes_anterior).Value = "Actual"
        if mes >= 2:
            ws.Cells(5, col_forecast_lag).Value = "Actual"

    ws_ano = wb.Worksheets("Resumo Resultado Ano")
    ws_ano.Range("Q4").Value = (
        f"Actual {MESES_INGLES[mes]}" if ciclo == "Actual" else f"{MESES_INGLES[mes]} Flash"
    )
    ws_ano.Range("AF4").Value = f"Forecast R{revisao_forecast}"
    log(f"Resumo Resultado Ano: Q4/AF4/linha5 atualizados (mês {col_letra(col_mes)}).")

    ws_ytd = wb.Worksheets("Resultado YTD")
    ws_ytd.Range("Q4").Value = (
        f"YTD {MESES_INGLES[mes]}" if ciclo == "Actual" else f"{MESES_INGLES[mes]} YTD"
    )
    if mes >= 2:
        # Achado validando contra os arquivos reais (2026-08-27): no Actual,
        # a aba YTD só tem 1 bloco (D:O) e NÃO tem rótulo "Forecast R<n>"
        # (AF4 fica em branco). No Flash, ela replica os 4 blocos (D:O/
        # Forecast/MP/PY) e TEM AF4 - por isso as duas regras abaixo só
        # rodam pros blocos que de fato existem em cada Ciclo.
        blocos_ytd = [COL_BLOCO_MES] if ciclo == "Actual" else [
            COL_BLOCO_MES, COL_BLOCO_FORECAST, COL_BLOCO_MP, COL_BLOCO_PY
        ]
        for bloco in blocos_ytd:
            ws_ytd.Cells(4, bloco + mes - 1).Value = MESES_ABREV[mes]
        if ciclo == "Flash":
            ws_ytd.Range("AF4").Value = f"Forecast R{revisao_forecast}"

    ws_mes = wb.Worksheets("Resumo Resultado M\u00eas")
    ws_mes.Range("D4").Value = f"{MESES_INGLES[mes]} Month"
    rotulo_forecast_cel = "F5" if ciclo == "Actual" else "E5"
    ws_mes.Range(rotulo_forecast_cel).Value = f"Forecast R{revisao_forecast}"
    log(f"Resumo Resultado Mês: D4/{rotulo_forecast_cel} atualizados.")


def atualizar_ponteiros_resumo_mes(wb, mes: int, ciclo: str, log):
    """Reescreve as fórmulas deslizantes de 'Resumo Resultado Mês' (colunas
    D/E/F/G ou D/F/G/H, dependendo do Ciclo - NÃO é link externo, exceto a
    coluna Flash do Ciclo Actual, que atualizar_links já trata)."""
    ws = wb.Worksheets("Resumo Resultado M\u00eas")
    col_valor_mes = col_letra(COL_BLOCO_MES + mes - 1)
    col_forecast = col_letra(COL_BLOCO_FORECAST + mes - 1)
    col_mp = col_letra(COL_BLOCO_MP + mes - 1)
    col_py = col_letra(COL_BLOCO_PY + mes - 1)

    if ciclo == "Actual":
        colunas_destino = {"D": col_valor_mes, "F": col_forecast, "G": col_mp, "H": col_py}
    else:
        colunas_destino = {"D": col_valor_mes, "E": col_forecast, "F": col_mp, "G": col_py}

    for col_dest, col_fonte in colunas_destino.items():
        for linha in LINHAS_RESUMO_MES:
            ws.Range(f"{col_dest}{linha}").Formula = f"='Resumo Resultado Ano'!{col_fonte}{linha}"
    log(f"Resumo Resultado Mês: ponteiros ({', '.join(colunas_destino)}) apontando pro mês certo.")


def atualizar_ytd(wb, mes: int, ciclo: str, log):
    """'Resultado YTD' ganha a coluna do novo mês, copiada da coluna do mês
    anterior - exatamente como a usuária descreveu ('copia a coluna de
    fórmula do mês anterior'). Usa `FormulaR1C1` (array por célula) em vez
    de Copy/PasteSpecial: testado 2026-08-27, o Copy/PasteSpecial via COM
    falha em instância isolada/invisível do Excel (sem acesso à área de
    transferência) - atribuir FormulaR1C1 tem o MESMO efeito (referências
    relativas se ajustam pra nova posição, valores literais são mantidos
    como estão), sem depender de clipboard."""
    if mes < 2:
        log("Mês 1 (Janeiro): geração da coluna YTD fora do escopo desta versão - pular.")
        return

    ws = wb.Worksheets("Resultado YTD")
    blocos = [COL_BLOCO_MES] if ciclo == "Actual" else [
        COL_BLOCO_MES, COL_BLOCO_FORECAST, COL_BLOCO_MP, COL_BLOCO_PY
    ]
    for bloco in blocos:
        col_velha = col_letra(bloco + mes - 2)
        col_nova = col_letra(bloco + mes - 1)
        origem = ws.Range(f"{col_velha}1:{col_velha}{ULTIMA_LINHA_YTD}")
        destino = ws.Range(f"{col_nova}1:{col_nova}{ULTIMA_LINHA_YTD}")
        destino.FormulaR1C1 = origem.FormulaR1C1
    log(f"Resultado YTD: coluna(s) do novo mês copiada(s) da coluna anterior ({len(blocos)} bloco(s)).")


def gerar_arquivo_pnl(mes: int, ano: int, ciclo: str, pasta_saida: Path, log=print, pid_callback=None) -> Path:
    pasta_saida = Path(pasta_saida).resolve()
    mes_ant, ano_ant = mes_anterior(mes, ano)
    caminho_anterior = localizar_pnl(mes_ant, ano_ant, ciclo)
    log(f"Base (mês anterior, mesmo Ciclo): {caminho_anterior}")

    pasta_saida.mkdir(parents=True, exist_ok=True)
    nome_novo = nome_com_versao(pasta_saida, nome_arquivo_pnl(mes, ano, ciclo))
    caminho_novo = pasta_saida / nome_novo
    shutil.copy2(caminho_anterior, caminho_novo)
    log(f"Copiado -> {caminho_novo}")

    excel = abrir_excel_isolado(log, pid_callback)
    excel.AskToUpdateLinks = False
    try:
        wb = excel.Workbooks.Open(str(caminho_novo))
        revisao_forecast = atualizar_links(wb, mes, ano, ciclo, log)
        if ciclo == "Flash" and mes >= 2:
            atualizar_coluna_mes_corrente_flash(wb, mes, ano, log)
        # Ordem importa: atualizar_ytd copia a coluna INTEIRA (linhas 1-60)
        # do mês anterior, incluindo rótulo/linha5 - se rodasse depois de
        # atualizar_textos, sobrescreveria o texto recém-corrigido com o
        # valor antigo copiado. Achado real testando este script,
        # 2026-08-27 (J4 voltava pra 'Jun' em vez de 'Jul').
        atualizar_ytd(wb, mes, ciclo, log)
        atualizar_textos(wb, mes, ano, ciclo, revisao_forecast, log)
        atualizar_ponteiros_resumo_mes(wb, mes, ciclo, log)
        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()

    log(f"\nArquivo de P&L gerado: {caminho_novo}")
    return caminho_novo


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Gera o arquivo de P&L (Passo 7)")
    parser.add_argument("--mes", type=int, required=True)
    parser.add_argument("--ano", type=int, required=True)
    parser.add_argument("--ciclo", choices=["Actual", "Flash"], required=True)
    parser.add_argument("--pasta-saida", type=Path, required=True, help="Pasta de saída (teste local ou rede)")
    args = parser.parse_args()

    gerar_arquivo_pnl(args.mes, args.ano, args.ciclo, args.pasta_saida)
