"""Gera o 'Net Sales MP27 - Envio Cobranca' da TLM no formato do modelo da Fitted.

Uso: python mp27_net_sales_envio.py <Resultado_TLM_xlsm> <modelo_Net_Sales_xlsx> <saida_xlsx>

Le do Resultado (aba 'P&L - TLM', valores gravados): linha 4 (Tonelada Movimentada) e
linha 13 (Faturamento Liquido), MP27 = B:M e R09 = P:AA. Edita o XML do modelo para
manter a formatacao identica; totais e deltas seguem como formulas.
"""
import re
import sys
import zipfile

import openpyxl

LIN_TON, LIN_NET = 4, 13
COLS_MP27 = range(2, 14)   # B:M
COLS_R09 = range(16, 28)   # P:AA


def ler(resultado):
    ws = openpyxl.load_workbook(resultado, data_only=True, read_only=True)["P&L - TLM"]
    linhas = {r: [c for c in row] for r, row in enumerate(ws.iter_rows(min_row=1, max_row=LIN_NET, values_only=True), 1)}
    def pega(lin, cols):
        v = [linhas[lin][c - 1] for c in cols]
        assert all(isinstance(x, (int, float)) for x in v), (lin, v)
        return v
    return {
        "ton_mp": pega(LIN_TON, COLS_MP27), "net_mp": pega(LIN_NET, COLS_MP27),
        "ton_r9": pega(LIN_TON, COLS_R09), "net_r9": pega(LIN_NET, COLS_R09),
    }


def main(resultado, modelo, saida):
    d = ler(resultado)
    letras = "BCDEFGHIJKLM"
    valores = {}
    for lin, chave in ((5, "ton_mp"), (6, "net_mp"), (11, "ton_r9"), (12, "net_r9")):
        for l, v in zip(letras, d[chave]):
            valores[f"{l}{lin}"] = v
        valores[f"N{lin}"] = sum(d[chave])
    for lin_d, a, b in ((14, 5, 11), (15, 6, 12)):
        for l in letras + "N":
            valores[f"{l}{lin_d}"] = valores[f"{l}{a}"] - valores[f"{l}{b}"]

    zin = zipfile.ZipFile(modelo)
    with zipfile.ZipFile(saida, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            dados = zin.read(it.filename)
            if it.filename == "xl/worksheets/sheet1.xml":
                x = dados.decode("utf-8")
                for ref, v in valores.items():
                    pad, n = re.subn(
                        rf'(<c r="{ref}"[^>]*>(?:<f[^>]*/>|<f[^>]*>[^<]*</f>)?)<v>[^<]*</v>', rf"\g<1><v>{v!r}</v>", x)
                    assert n == 1, ref
                    x = pad
                dados = x.encode("utf-8")
            elif it.filename == "xl/sharedStrings.xml":
                x = dados.decode("utf-8")
                x = x.replace("<t>PEZZI</t>", "<t>TONELADAS</t>").replace("<t>BUDGET'27</t>", "<t>MP27</t>")
                dados = x.encode("utf-8")
            elif it.filename == "xl/workbook.xml":
                x = dados.decode("utf-8")
                x = re.sub(r"<mc:AlternateContent.*?</mc:AlternateContent>", "", x, flags=re.S)
                dados = x.encode("utf-8")
            zout.writestr(it, dados)
    print("ok", saida)
    print("MP27 ton", round(valores["N5"], 1), "net", round(valores["N6"], 1))
    print("R9   ton", round(valores["N11"], 1), "net", round(valores["N12"], 1))
    print("delta ton", round(valores["N14"], 1), "net", round(valores["N15"], 1))


if __name__ == "__main__":
    main(*sys.argv[1:4])
