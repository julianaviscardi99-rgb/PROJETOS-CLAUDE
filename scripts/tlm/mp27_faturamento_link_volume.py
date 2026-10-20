"""MP27 TLM - liga as toneladas do Faturamento MP27 ao Volume Movimentado MP27 (formula com link externo).

Uso: python mp27_faturamento_link_volume.py <faturamento_origem.xlsx> <volume_movimentado.xlsm> <faturamento_destino.xlsx> [aba_fluxos] [sufixo_rotulo]

O volume de 20/10 renomeou a aba para "03. Fluxos_MP27_Ofic" e os rotulos ganharam o sufixo "_Ofic".

- Origem NUNCA e alterada: o destino e um arquivo novo (regra do projeto: _v2, _v3 ou data nova).
- Linha Tons de cada unidade/BU (S:AD, jan-dez MP27) = '03. Fluxos_MP27' do Volume Movimentado, linha
  <WH><BU>TOTALTOTALMP27 (IN + OUT = "Tons Mov" do Resultado), coluna das toneladas (U:AF) / 1000 (k t).
  Mesma logica do Resultado TLM R10, aba Faturamento ('10. Fluxos_R10' linhas TOTAL).
- Precisa de Excel (pywin32), como mp27_resultado.py / mp27_tarifa.py.
"""
import os
import sys

from mp27_tarifa import abrir_excel

ABA_FLUXOS = "03. Fluxos_MP27"
COL_FAT_JAN = 19        # S = MP27 Jan na aba Faturamento
COL_FLUXO_T_JAN = 21    # U = Jan (toneladas) em 03. Fluxos_MP27
AZUL_LINK = 0xF7EBDD    # BGR de DDEBF7

# (linha Tons em Faturamento, rotulo da linha em 03. Fluxos_MP27 coluna B, descricao)
LINHAS = [
    (6, "BACarTOTALTOTALMP27", "Barueri Car"),
    (11, "BAMotoTOTALTOTALMP27", "Barueri Moto"),
    (16, "CPCarTOTALTOTALMP27", "Campinas Car"),
    (21, "CPMotoTOTALTOTALMP27", "Campinas Moto"),
    (38, "FSCarTOTALTOTALMP27", "Feira de Santana Car"),
]


def achar_linha(ws_fluxos, rotulo):  # rotulo ja com sufixo
    achadas = [r for r in range(1, ws_fluxos.UsedRange.Rows.Count + 1) if ws_fluxos.Cells(r, 2).Value == rotulo]
    if len(achadas) != 1:
        raise RuntimeError("rotulo %r aparece %d vezes em %s" % (rotulo, len(achadas), ABA_FLUXOS))
    return achadas[0]


def main(origem, volume, destino, aba=ABA_FLUXOS, sufixo=""):
    origem, volume, destino = (os.path.abspath(p) for p in (origem, volume, destino))
    if os.path.exists(destino):
        raise SystemExit("destino ja existe, nao sobrescrevo: " + destino)
    xl = abrir_excel()
    try:
        xl.AutomationSecurity = 3  # macros desligadas ao abrir o .xlsm
        wv = xl.Workbooks.Open(volume, 0, True)  # sem atualizar links, somente leitura
        wb = xl.Workbooks.Open(origem, 0, False)
        ws = wb.Worksheets("Faturamento")
        wf = wv.Worksheets(aba)
        pasta, arq = os.path.split(volume)
        ref = "'%s\\[%s]%s'" % (pasta, arq, aba)

        for lin_fat, rotulo, nome in LINHAS:
            lin_fl = achar_linha(wf, rotulo + sufixo)
            for j in range(12):
                col_fl = wf.Cells(1, COL_FLUXO_T_JAN + j).Address.replace("$", "").rstrip("1")
                c = ws.Cells(lin_fat, COL_FAT_JAN + j)
                c.Formula = "=%s!%s%d/1000" % (ref, col_fl, lin_fl)
                c.Interior.Color = AZUL_LINK
            ws.Cells(lin_fat, 33).Value = "link: %s!%s (t / 1000)" % (aba, rotulo + sufixo)
            print("  %-22s linha %3d <- %s linha %d" % (nome, lin_fat, aba, lin_fl))

        # Premissas: troca o aviso de simulacao so para Tons
        wp = wb.Worksheets("Premissas")
        wp.Range("B10").Value = ("Tons das 5 linhas de Logistica (Barueri Car/Moto, Campinas Car/Moto, Feira Car) = LINK para o Volume Movimentado MP27 da tecnica "
                                 "(03. Fluxos_MP27, linhas TOTAL = IN + OUT, em t / 1000; celulas AZUIS). Ecolabel e Outras Atividades continuam SIMULACAO "
                                 "(R09 repetido, LARANJA) ate a tecnica / areas mandarem.")
        wp.Range("B21").Value = "Tons ligadas ao Volume Movimentado MP27 por scripts/tlm/mp27_faturamento_link_volume.py"

        xl.CalculateFull()
        wb.SaveAs(destino, 51)
        wb.Close(False)
        wv.Close(False)
        print("Gravado:", destino)
    finally:
        try:
            xl.Quit()
        except Exception as e_quit:             # nao mascarar a excecao original (Excel ocupado: 'Call was rejected by callee')
            print('  AVISO: Excel nao fechou (%s)' % str(e_quit)[:60])


if __name__ == "__main__":
    if len(sys.argv) not in (4, 6):
        raise SystemExit(__doc__)
    main(*sys.argv[1:])
