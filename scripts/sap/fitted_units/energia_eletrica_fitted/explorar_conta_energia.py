#!/usr/bin/env python3
"""
Exploracao inicial do sub-projeto "Energia Eletrica Fitted" (2026-08-13).

Extrai a KSB1 (Sem Agrupamento) para o ano de 2026, usando sempre o mesmo
layout/parametros ja mapeados pro resto de Fitted Units (BU['kstgr'] e
BU['disvar'], nunca preencher Centro de custo/Classe de custo direto na
tela — isso causa o erro "Selecionar uma das alternativas indicadas",
porque sao campos alternativos aos de grupo).

Salva sempre em data/processed/ (nunca em pasta temporaria nova, pra nao
disparar o popup de seguranca do SAPGUI pedindo autorizacao de novo a
cada pasta diferente).
"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "_shared"))
from ksb1_core import BU, abrir_ksb1, connect_session, nome_com_versao, voltar_para_selecao

PASTA = Path(__file__).resolve().parents[4] / "data" / "processed" / "energia_eletrica_fitted"


def extrair(data_de: str, data_ate: str, log=print) -> Path:
    session = connect_session()
    abrir_ksb1(session, log)

    wnd = session.FindById("wnd[0]")
    # Nunca tocar em KOSTL-LOW/HIGH nem KSTAR-LOW/HIGH (Centro de custo /
    # Classe de custo especificos) -- sao alternativos aos campos de grupo
    # abaixo e causam erro se os dois lados ficarem preenchidos ao mesmo
    # tempo.
    wnd.FindById("usr/ctxtKSTGR").Text = BU["kstgr"]
    wnd.FindById("usr/ctxtKOAGR").Text = ""  # Sem Agrupamento
    wnd.FindById("usr/ctxtR_BUDAT-LOW").Text = data_de
    wnd.FindById("usr/ctxtR_BUDAT-HIGH").Text = data_ate
    wnd.FindById("usr/ctxtP_DISVAR").Text = BU["disvar"]

    log(f"Executando KSB1 Sem Agrupamento, {data_de} a {data_ate}...")
    wnd.SendVKey(8)
    time.sleep(2)

    PASTA.mkdir(parents=True, exist_ok=True)
    nome_arquivo = nome_com_versao(PASTA, "KSB1 - Fitted Units 2026 - Sem Agrupamento (energia).xlsx")

    session.FindById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select()
    time.sleep(1)
    wnd1 = session.FindById("wnd[1]")
    wnd1.FindById("usr/ctxtDY_PATH").Text = str(PASTA)
    wnd1.FindById("usr/ctxtDY_FILENAME").Text = nome_arquivo
    wnd1.FindById("tbar[0]/btn[0]").Press()

    caminho = PASTA / nome_arquivo
    for _ in range(40):
        if caminho.exists():
            break
        time.sleep(0.5)

    if caminho.exists():
        log(f"Salvo em {caminho}")
    else:
        log("AVISO: não encontrei o arquivo gerado na pasta esperada.")

    voltar_para_selecao(session, log)
    return caminho


if __name__ == "__main__":
    extrair("01.01.2026", "31.07.2026")
