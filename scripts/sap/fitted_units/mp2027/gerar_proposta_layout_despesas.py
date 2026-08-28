"""MP2027 - PROPOSTA de layout pro arquivo "Detalhe_Despesas_Fitted Units"
(hoje 107 colunas, 6 blocos repetidos de "R07 JAN..DEZ" sem nome claro).
Pedido da usuária (2026-08-28): colunas A-X (classificação) ficam iguais;
depois disso, layout novo por mês: Valor Mensal -> % Reajuste -> Valor
Final Previsão (só isso, os outros 3 blocos - quantidade/valor calculado/
valor com incremento - somem, eram intermediários de cálculo, não achado
final). Números vêm do R7 (Forecast Julho) só pra prototipar o layout -
"não estou me importando com os números agora".

Mapeamento dos blocos originais (decodificado via fórmula, não achismo):
- bloco1 (col Y, 24) = Valor Mensal (base, valor estático)
- bloco4 (col BQ, 68) = % Reajuste (índice de reajuste aplicado, estático)
- bloco6 (col CQ, 94) = Valor Final Previsão (número oficial - validado
  batendo com a aba "Resumo Custos" em 2026-08-28)
Blocos 2 (quantidade), 3 (=bloco1*bloco2) e 5 (=bloco1*bloco4+bloco3) ficam
de fora - eram passos intermediários do cálculo, não usados fora dessa aba.

Próxima etapa (depois da usuária aprovar este layout): automatizar a coluna
Efetivo puxando do arquivo oficial "KSB1 <Mês> Actual <Ano>.xlsx" (aba
BASE_KSB1) - ainda não implementado, só o layout do Budget/Forecast por
enquanto.
"""
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "_shared"))
from ksb1_core import nome_com_versao  # noqa: E402

ORIGEM = Path(r"\\FSS024-01BR.group.pirelli.com\GFU_DAC\Custos Fitted Units\Resultados Fitted\2026\07 - Jul\07_Jul_Forecast\Detalhe_Despesas_Fitted Units_Forecast July.xlsx")
SAIDA = Path(r"\\FSS024-01BR.group.pirelli.com\GFU_DAC\Management Plan\MP2027")

MESES = ["JAN", "FEV", "MAR", "ABR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]

# colunas A-X, 0-indexed 0-23 - mantidas exatamente como estão
N_COLS_CLASSIFICACAO = 24

# metadados por linha (nao mudam por mes) - ficam logo apos a classificacao
COLS_METADADOS = {
    "Área Responsável": 37,
    "Índice de Reajuste": 38,
    "Mês de Contrato": 39,
    "Comprador": 40,
    "Premissas": 41,
}

COL_VALOR_MENSAL_INICIO = 24   # bloco1
COL_REAJUSTE_INICIO = 68       # bloco4
COL_VALOR_FINAL_INICIO = 94    # bloco6

AZUL_CLARO = PatternFill("solid", fgColor="D9E6F5")
CINZA_CLARO = PatternFill("solid", fgColor="EFEFEF")
VERDE_CLARO = PatternFill("solid", fgColor="E2F0D9")


def montar_proposta():
    wb_origem = load_workbook(ORIGEM, data_only=True, read_only=True)
    ws_origem = wb_origem["DataBase_Detail"]

    wb = Workbook()
    ws = wb.active
    ws.title = "Proposta Layout"

    header_classificacao = None
    linhas_saida = []

    for i, row in enumerate(ws_origem.iter_rows(values_only=True)):
        if i == 1:
            header_classificacao = row[:N_COLS_CLASSIFICACAO]
            continue
        if i < 2:
            continue

        classificacao = list(row[:N_COLS_CLASSIFICACAO])
        metadados = [row[c] for c in COLS_METADADOS.values()]
        valor_mensal = [row[COL_VALOR_MENSAL_INICIO + m] for m in range(12)]
        reajuste = [row[COL_REAJUSTE_INICIO + m] for m in range(12)]
        valor_final = [row[COL_VALOR_FINAL_INICIO + m] for m in range(12)]

        if all(v is None for v in valor_final):
            continue

        linhas_saida.append((classificacao, metadados, valor_mensal, reajuste, valor_final))

    # cabecalho: classificacao (A-X) + metadados + 3 blocos x 12 meses + total
    header = list(header_classificacao) + list(COLS_METADADOS.keys())
    for metrica in ["Valor Mensal", "% Reajuste", "Valor Final Previsão"]:
        for mes in MESES:
            header.append(f"{metrica} {mes}")
    header.append("TOTAL ANO (Valor Final Previsão)")
    ws.append(header)

    for cel in ws[1]:
        cel.font = Font(bold=True)
    col_inicio_metricas = N_COLS_CLASSIFICACAO + len(COLS_METADADOS) + 1
    ws.freeze_panes = f"{get_column_letter(col_inicio_metricas)}2"  # trava classificacao+metadados, rola só nas metricas

    for classificacao, metadados, valor_mensal, reajuste, valor_final in linhas_saida:
        total_ano = sum(v for v in valor_final if isinstance(v, (int, float)))
        linha = list(classificacao) + list(metadados) + list(valor_mensal) + list(reajuste) + list(valor_final) + [total_ano]
        ws.append(linha)

    n_linhas = len(linhas_saida) + 1
    col_valor_mensal_saida = N_COLS_CLASSIFICACAO + len(COLS_METADADOS) + 1
    col_reajuste_saida = col_valor_mensal_saida + 12
    col_valor_final_saida = col_reajuste_saida + 12
    col_total_saida = col_valor_final_saida + 12

    for row_idx in range(2, n_linhas + 1):
        for c in range(col_valor_mensal_saida, col_valor_mensal_saida + 12):
            ws.cell(row=row_idx, column=c).fill = AZUL_CLARO
            ws.cell(row=row_idx, column=c).number_format = "#,##0"
        for c in range(col_reajuste_saida, col_reajuste_saida + 12):
            ws.cell(row=row_idx, column=c).fill = CINZA_CLARO
            ws.cell(row=row_idx, column=c).number_format = "0.00%"
        for c in range(col_valor_final_saida, col_valor_final_saida + 12):
            ws.cell(row=row_idx, column=c).fill = VERDE_CLARO
            ws.cell(row=row_idx, column=c).number_format = "#,##0"
        ws.cell(row=row_idx, column=col_total_saida).number_format = "#,##0"
        ws.cell(row=row_idx, column=col_total_saida).font = Font(bold=True)

    for c in range(col_valor_mensal_saida, col_valor_mensal_saida + 12):
        ws.cell(row=1, column=c).fill = AZUL_CLARO
    for c in range(col_reajuste_saida, col_reajuste_saida + 12):
        ws.cell(row=1, column=c).fill = CINZA_CLARO
    for c in range(col_valor_final_saida, col_valor_final_saida + 12):
        ws.cell(row=1, column=c).fill = VERDE_CLARO
    ws.cell(row=1, column=col_total_saida).fill = VERDE_CLARO

    for c in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(c)].width = 13
    ws.column_dimensions[get_column_letter(8)].width = 40   # DETALHE SERVIÇO/PRODUTO
    ws.column_dimensions[get_column_letter(10)].width = 30  # Nome Fornecedor

    SAIDA.mkdir(parents=True, exist_ok=True)
    nome_arquivo = nome_com_versao(SAIDA, "MP2027_PROPOSTA_Layout_Detalhe_Despesas.xlsx")
    caminho_final = SAIDA / nome_arquivo
    wb.save(caminho_final)
    return caminho_final, len(linhas_saida)


if __name__ == "__main__":
    caminho, n = montar_proposta()
    print(f"Proposta gerada: {caminho}")
    print(f"{n} linhas.")
    print("\nAzul = Valor Mensal | Cinza = % Reajuste | Verde = Valor Final Previsão")
    print("Colunas A-X (classificação) e os 5 metadados (Área Responsável..Premissas) intactos.")
