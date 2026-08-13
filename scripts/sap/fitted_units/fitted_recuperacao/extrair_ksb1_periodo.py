#!/usr/bin/env python3
"""
Extrai a KSB1 (Fitted Units, Sem Agrupamento) para um periodo arbitrario —
usado no estudo "Fitted Recuperacao" (deteccao de lancamento/pagamento em
duplicidade a fornecedor), que precisa de todo o historico (jan-jul/2026),
nao de um mes fechado por vez como o fluxo mensal recorrente.

Sem Agrupamento (nao Gestoriais) porque aqui o objetivo e pegar TODAS as
contas/lancamentos do periodo, sem filtro de agrupamento gestorial.

Reaproveita as funcoes ja testadas de conexao/navegacao/popup de
fitted_units/_shared/ksb1_core.py em vez de duplicar essa logica.

Pre-requisitos: SAP GUI aberto e logada (script abre a KSB1 sozinho).
"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "_shared"))
from ksb1_core import (  # noqa: E402
    BU,
    abrir_ksb1,
    connect_session,
    nome_com_versao,
    voltar_para_selecao,
)

PASTA_DESTINO = Path(
    r"\\FSS024-01BR.group.pirelli.com\GFU_DAC\Custos Fitted Units\Estudos\Estudo Duplicidade Pagamento"
)


def extrair_periodo(session, data_de: str, data_ate: str, pasta_destino: Path, log=print) -> Path:
    wnd = session.FindById("wnd[0]")
    wnd.FindById("usr/ctxtP_KOKRS").Text = "0580"
    if wnd.FindById("usr/ctxtKSTGR", False) is not None:
        wnd.FindById("usr/ctxtKSTGR").Text = BU["kstgr"]
    wnd.FindById("usr/ctxtKOAGR").Text = ""  # Sem Agrupamento
    wnd.FindById("usr/ctxtR_BUDAT-LOW").Text = data_de
    wnd.FindById("usr/ctxtR_BUDAT-HIGH").Text = data_ate
    wnd.FindById("usr/ctxtP_DISVAR").Text = BU["disvar"]

    log(f"Executando KSB1 Sem Agrupamento, período {data_de} a {data_ate}...")
    session.FindById("wnd[0]").SendVKey(8)

    pasta_destino.mkdir(parents=True, exist_ok=True)
    de_fmt = data_de.replace(".", "")
    ate_fmt = data_ate.replace(".", "")
    nome_arquivo = nome_com_versao(
        pasta_destino, f"KSB1 - {BU['nome']} {de_fmt}-{ate_fmt} - Sem Agrupamento.XLSX"
    )

    session.FindById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select()
    wnd1 = session.FindById("wnd[1]")
    wnd1.FindById("usr/ctxtDY_PATH").Text = str(pasta_destino)
    wnd1.FindById("usr/ctxtDY_FILENAME").Text = nome_arquivo
    wnd1.FindById("tbar[0]/btn[0]").Press()  # Gerar

    arquivo_final = pasta_destino / nome_arquivo
    for _ in range(40):
        if arquivo_final.exists():
            break
        time.sleep(0.5)

    if arquivo_final.exists():
        log(f"Salvo em {arquivo_final}")
    else:
        log("AVISO: não encontrei o arquivo gerado na pasta esperada. Confira manualmente.")

    voltar_para_selecao(session, log)
    return arquivo_final


def main():
    if len(sys.argv) != 3:
        print("Uso: python extrair_ksb1_periodo.py <DD.MM.AAAA> <DD.MM.AAAA>")
        print("Ex:  python extrair_ksb1_periodo.py 01.01.2026 31.07.2026")
        sys.exit(1)

    data_de, data_ate = sys.argv[1], sys.argv[2]

    try:
        session = connect_session()
    except Exception as e:
        print(f"ERRO ao conectar ao SAP GUI: {e}")
        sys.exit(1)

    try:
        abrir_ksb1(session, log=print)
    except RuntimeError as e:
        print(str(e))
        sys.exit(1)

    extrair_periodo(session, data_de, data_ate, PASTA_DESTINO, log=print)
    print("\nConcluído!")


if __name__ == "__main__":
    main()
