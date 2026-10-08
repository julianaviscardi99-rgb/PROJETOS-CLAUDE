#!/usr/bin/env python3
"""
Cria uma COPIA de um arquivo 'KSB1 <Mes> <Ciclo> <Ano>.xlsx' (aba BASE_KSB1 +
Pivot_Inter. + Pivot_Detalhes) mantendo na BASE_KSB1 so os lancamentos dos
meses pedidos (coluna 'Mes', S). O original nunca e' alterado.

Uso: python filtrar_ksb1_por_meses.py <origem.xlsx> <mes_inicial> <mes_final> <pasta_destino>

Como funciona: a BASE_KSB1 vem ordenada por mes (blocos contiguos); o script
confere isso e apaga de uma vez so o bloco de linhas fora do intervalo
(apagar linhas nao sofre do bug de corrupcao por escrita em bloco do COM),
atualiza as tabelas dinamicas e confere contagem/soma contra o original.
"""
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "_shared"))
from ksb1_core import abrir_excel_isolado, com_retry, nome_com_versao  # noqa: E402

ABA_BASE = "BASE_KSB1"
COL_MES = 19      # S
COL_VALOR = 17    # Q  (Valor/MR)
LINHA_CABECALHO = 1


def ler_coluna(ws, col, ultima):
    v = ws.Range(ws.Cells(2, col), ws.Cells(ultima, col)).Value
    return [x[0] for x in v]


def main(origem: Path, mes_ini: int, mes_fim: int, pasta_destino: Path, log=print):
    pasta_destino.mkdir(parents=True, exist_ok=True)
    sufixo = f"{mes_ini:02d}-{mes_fim:02d}"
    nome = nome_com_versao(pasta_destino, f"{origem.stem} - meses {sufixo}.xlsx")
    destino = pasta_destino / nome
    log(f"Copiando para {destino} ...")
    shutil.copy2(origem, destino)

    excel = abrir_excel_isolado(log, None)
    try:
        wb = com_retry(excel.Workbooks.Open, str(destino), UpdateLinks=0, ReadOnly=False, log=log)
        ws = wb.Worksheets(ABA_BASE)
        ultima = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        meses = ler_coluna(ws, COL_MES, ultima)
        valores = ler_coluna(ws, COL_VALOR, ultima)
        if any(m is None or isinstance(m, str) for m in meses):
            raise RuntimeError("Há linhas com 'Mês' vazio ou erro na BASE_KSB1 — abortei sem salvar.")

        antes = len(meses)
        manter = [i for i, m in enumerate(meses) if mes_ini <= m <= mes_fim]
        if not manter:
            raise RuntimeError("Nenhuma linha dentro do intervalo de meses.")
        # Blocos contiguos: tudo que sobra precisa ser um unico bloco no meio/fim.
        ini, fim = manter[0], manter[-1]
        if fim - ini + 1 != len(manter):
            raise RuntimeError("Os meses pedidos não estão em bloco contíguo na BASE_KSB1 — abortei.")
        soma_esperada = sum(v for i, v in enumerate(valores) if i in set(manter) and isinstance(v, (int, float)))

        # Apaga primeiro o bloco de baixo (depois do intervalo), depois o de cima.
        if fim < antes - 1:
            ws.Rows(f"{fim + 3}:{ultima}").Delete()
        if ini > 0:
            ws.Rows(f"2:{ini + 1}").Delete()

        ultima2 = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        meses2 = ler_coluna(ws, COL_MES, ultima2)
        valores2 = ler_coluna(ws, COL_VALOR, ultima2)
        soma2 = sum(v for v in valores2 if isinstance(v, (int, float)))
        if len(meses2) != len(manter) or abs(soma2 - soma_esperada) > 0.01 or any(not (mes_ini <= m <= mes_fim) for m in meses2):
            raise RuntimeError("Conferência pós-exclusão falhou — abortei sem salvar.")

        log("Atualizando tabelas dinâmicas...")
        com_retry(wb.RefreshAll, log=log)
        excel.CalculateUntilAsyncQueriesDone()
        for aba in ("Pivot_Inter.", "Pivot_Detalhes"):
            try:
                for pt in wb.Worksheets(aba).PivotTables():
                    pt.RefreshTable()
            except Exception as e:  # noqa: BLE001
                log(f"  aviso: refresh de {aba}: {e}")

        wb.Save()
        wb.Close(SaveChanges=False)
        log(f"OK: {antes} linhas -> {len(meses2)} linhas (meses {mes_ini}-{mes_fim}), soma Valor/MR = {soma2:,.2f}")
        return destino
    finally:
        excel.Quit()


if __name__ == "__main__":
    if len(sys.argv) != 5:
        sys.exit(__doc__)
    main(Path(sys.argv[1]), int(sys.argv[2]), int(sys.argv[3]), Path(sys.argv[4]))
