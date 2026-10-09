#!/usr/bin/env python3
"""Passo 1 do cockpit do Circuito Panamericano: extracao da KSB1 (todas as despesas da area 2281).

Independente da Fitted Units (nao importa nada de scripts/sap/fitted_units), mas segue o mesmo
fluxo de la: SAP GUI Scripting -> KSB1 -> exportar planilha -> pasta temporaria -> mover pra pasta final.

Diferencas em relacao a Fitted (explicadas pela Juliana em 2026-10-09):
- Area de contabilidade de custos 2281 (Fitted = 0580).
- Uma extracao so': TODAS as despesas da empresa, sem grupo de centros de custo e sem
  agrupamento (gestoriais / sem agrupamento nao se aplicam ao CP).
- Salva em <base CP>\\<ano>\\<MM - Mes>\\<02. Flash | 03. Actual>\\KSB1 - Circuito Panamericano MM.YYYY - <Ciclo>.XLSX (mesmo padrao da Fitted)
  (nome padrao dela; se ja existir, _v2, _v3... - nunca sobrescreve).

Pre-requisitos: pywin32; SAP GUI aberto/logado com Scripting habilitado
(Alt+F12 > Opcoes > Acessibilidade e Scripting > Scripting).
"""
import calendar
import shutil
import time
from pathlib import Path

import pythoncom
import win32com.client
import win32con
import win32gui

NOME_BU = "Circuito Panamericano"
AREA_CONTAB = "2281"
VARIANTE = "/DESPFITTED"
CICLOS = ("Flash", "Actual")
PASTA_CICLO = {"Flash": "02. Flash", "Actual": "03. Actual"}

REDE_BASE = Path(
    r"\\FSS024-01BR.group.pirelli.com\CONTROLLING\Reporting"
    r"\Reporting ACT_FCST_MP Cons. e Ind\Circuito Panamericano"
)
# O SAP so' deixa o script gravar sem o popup "Seguranca SAPGUI" (Permitir / Memorizar minha
# decisao) numa pasta que ele ja conhece. O CP exporta pela pasta temporaria fixa que a Fitted ja
# usa (ja autorizada) e depois move o arquivo pra pasta final. Pedido da Juliana (2026-10-09): usar
# SEMPRE a pasta Temporario de 2026, tambem em 2027 em diante (nao muda por ano). Os nomes de
# arquivo nao colidem com os da Fitted ("KSB1 - Circuito Panamericano ..." x "KSB1 - Fitted Units ...").
PASTA_TEMPORARIA = Path(
    r"\\FSS024-01BR.group.pirelli.com\GFU_DAC\Custos Fitted Units\Resultados Fitted"
    r"\2026\00.Extração Base KSB1\Temporario"
)


MESES_NOMES = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
    7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}
MESES_PASTA = {
    1: "01 - Jan", 2: "02 - Feb", 3: "03 - Mar", 4: "04 - Apr", 5: "05 - May", 6: "06 - Jun",
    7: "07 - Jul", 8: "08 - Aug", 9: "09 - Sep", 10: "10 - Oct", 11: "11 - Nov", 12: "12 - Dec",
}


class ErroComTitulo(Exception):
    """Erro com titulo proprio pro messagebox (a GUI mostra na thread principal)."""

    def __init__(self, titulo, mensagem):
        super().__init__(mensagem)
        self.titulo = titulo
        self.mensagem = mensagem


def pasta_destino(mes: int, ano: int, ciclo: str) -> Path:
    return REDE_BASE / str(ano) / MESES_PASTA[mes] / PASTA_CICLO[ciclo]


def nome_arquivo(mes: int, ano: int, ciclo: str) -> str:
    # Mesmo padrao da Fitted ("KSB1 - Fitted Units 09.2026 - Sem Agrupamento - Actual.XLSX"),
    # sem a parte do agrupamento (o CP traz todas as despesas numa extracao so').
    return f"KSB1 - {NOME_BU} {mes:02d}.{ano} - {ciclo}.XLSX"


def nome_com_versao(pasta: Path, nome_base: str) -> str:
    """Nunca sobrescreve: nome.XLSX, nome_v2.XLSX, nome_v3.XLSX ..."""
    base = Path(nome_base)
    candidato = pasta / nome_base
    versao = 2
    while candidato.exists():
        candidato = pasta / f"{base.stem}_v{versao}{base.suffix}"
        versao += 1
    return candidato.name


# ---------------------------------------------------------------- SAP GUI

def connect_session():
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    return application.Children(0).Children(0)


def _buscar_campo_editavel(elemento):
    """Primeiro GuiCTextField dentro do elemento (o campo do popup fica num subscreen)."""
    try:
        if elemento.Type == "GuiCTextField":
            return elemento
    except Exception:
        pass
    try:
        filhos = elemento.Children
    except Exception:
        return None
    for filho in filhos:
        achado = _buscar_campo_editavel(filho)
        if achado is not None:
            return achado
    return None


def tratar_popup_area_contabil(session, log):
    """Popup 'Definir area contab.custos' (aparece na 1a transacao de uma sessao nova)."""
    wnd1 = session.FindById("wnd[1]", False)
    if wnd1 is None:
        return
    log(f"Popup 'Definir área contab.custos' detectado, preenchendo {AREA_CONTAB}...")
    campo = _buscar_campo_editavel(wnd1.FindById("usr"))
    if campo is None:
        raise RuntimeError(
            "Não consegui identificar o campo do popup 'Definir área contab.custos'. "
            f"Preencha {AREA_CONTAB} manualmente no SAP, confirme, e tente de novo."
        )
    campo.Text = AREA_CONTAB
    wnd1.FindById("tbar[0]/btn[0]").Press()


def abrir_ksb1(session, log):
    log("Abrindo a transação KSB1...")
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nKSB1"
    session.FindById("wnd[0]").SendVKey(0)
    tratar_popup_area_contabil(session, log)
    if session.Info.Transaction != "KSB1":
        raise RuntimeError(f"Não consegui abrir a KSB1 (tela atual: '{session.Info.Transaction}').")


def voltar_para_selecao(session, log):
    log("Voltando para a tela de seleção...")
    wnd = session.FindById("wnd[0]")
    wnd.FindById("tbar[0]/btn[3]").Press()
    if wnd.FindById("usr/ctxtP_KOKRS", False) is None:
        abrir_ksb1(session, log)


# ---------------------------------------------------------------- Excel aberto pelo SAP

def fechar_excel_se_aberto(caminho_arquivo: Path, log=print) -> bool:
    """O SAP as vezes abre o arquivo exportado no Excel e trava o arquivo na hora de mover.
    Fecha so' essa pasta de trabalho (achada pelo nome na Running Object Table)."""
    nome_alvo = Path(caminho_arquivo).name.lower()
    try:
        rot = pythoncom.GetRunningObjectTable()
        ctx = pythoncom.CreateBindCtx(0)
    except Exception:
        return False
    for moniker in rot:
        try:
            nome = moniker.GetDisplayName(ctx, None)
        except Exception:
            continue
        if not nome.lower().endswith(nome_alvo):
            continue
        try:
            obj = rot.GetObject(moniker)
            wb = win32com.client.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
            log(f"O SAP abriu '{Path(caminho_arquivo).name}' no Excel - fechando pra liberar o arquivo...")
            wb.Close(SaveChanges=False)
            return True
        except Exception:
            continue
    return False


def limpar_excel_orfao(log=print):
    """Fecha o aviso 'Microsoft Excel' que sobra depois de mover o arquivo e qualquer
    instancia do Excel sem nenhuma pasta de trabalho aberta (nunca mexe num Excel em uso)."""

    def _fechar_aviso(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd).strip() == "Microsoft Excel":
            try:
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                log("Fechei um aviso do Excel que sobrou depois da extração.")
            except Exception:
                pass
        return True

    try:
        win32gui.EnumWindows(_fechar_aviso, None)
    except Exception:
        pass
    try:
        excel = win32com.client.GetObject(Class="Excel.Application")
        if excel.Workbooks.Count == 0:
            excel.Quit()
            log("Fechei uma janela do Excel que ficou vazia depois da extração.")
    except Exception:
        pass


# ---------------------------------------------------------------- Extracao

def extrair_ksb1(session, mes, ano, ciclo, log):
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    data_de = f"01.{mes:02d}.{ano}"
    data_ate = f"{ultimo_dia:02d}.{mes:02d}.{ano}"

    wnd = session.FindById("wnd[0]")
    wnd.FindById("usr/ctxtP_KOKRS").Text = AREA_CONTAB
    # Todas as despesas da empresa: sem grupo de centros de custo e sem agrupamento
    for campo in ("usr/ctxtKSTGR", "usr/ctxtKOAGR"):
        if wnd.FindById(campo, False) is not None:
            wnd.FindById(campo).Text = ""
    wnd.FindById("usr/ctxtR_BUDAT-LOW").Text = data_de
    wnd.FindById("usr/ctxtR_BUDAT-HIGH").Text = data_ate
    wnd.FindById("usr/ctxtP_DISVAR").Text = VARIANTE

    log(f"Executando KSB1 (área {AREA_CONTAB}, {data_de} a {data_ate}, variante {VARIANTE})...")
    session.FindById("wnd[0]").SendVKey(8)  # Executar (F8)

    pasta_final = pasta_destino(mes, ano, ciclo)
    pasta_final.mkdir(parents=True, exist_ok=True)
    nome = nome_com_versao(pasta_final, nome_arquivo(mes, ano, ciclo))
    if nome != nome_arquivo(mes, ano, ciclo):
        log(f"Já existe '{nome_arquivo(mes, ano, ciclo)}' nessa pasta - salvando como '{nome}' (nada é sobrescrito).")

    PASTA_TEMPORARIA.mkdir(parents=True, exist_ok=True)
    arquivo_temp = PASTA_TEMPORARIA / nome
    if arquivo_temp.exists():
        for _ in range(30):  # sobra de tentativa anterior que falhou antes de mover
            try:
                arquivo_temp.unlink()
                break
            except OSError:
                fechar_excel_se_aberto(arquivo_temp, log)
                time.sleep(1)
        else:
            raise ErroComTitulo(
                "Arquivo temporário travado",
                f"'{arquivo_temp.name}' ainda está aberto em algum programa (provavelmente Excel).\n\n"
                "Feche o arquivo e tente de novo.",
            )

    session.FindById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select()  # Lista > Exportar > Planilha
    wnd1 = session.FindById("wnd[1]")
    wnd1.FindById("usr/ctxtDY_PATH").Text = str(PASTA_TEMPORARIA)
    wnd1.FindById("usr/ctxtDY_FILENAME").Text = nome
    wnd1.FindById("tbar[0]/btn[0]").Press()  # Gerar

    for _ in range(40):
        if arquivo_temp.exists():
            break
        time.sleep(0.5)

    arquivo_final = pasta_final / nome
    if not arquivo_temp.exists():
        log("AVISO: não encontrei o arquivo exportado na pasta temporária. Confira manualmente.")
    else:
        movido = False
        for _ in range(30):
            try:
                shutil.move(str(arquivo_temp), str(arquivo_final))
                movido = True
                break
            except OSError:
                fechar_excel_se_aberto(arquivo_temp, log)
                time.sleep(1)
        if movido:
            log(f"Salvo em: {arquivo_final}")
            for _ in range(5):
                time.sleep(1)
                limpar_excel_orfao(log)
        else:
            log(f"AVISO: o arquivo ficou travado na pasta temporária ({arquivo_temp}). "
                f"Feche-o e mova manualmente para {arquivo_final}.")

    voltar_para_selecao(session, log)
    return arquivo_final


def rodar(mes, ano, ciclo, log):
    """Ponto de entrada do Passo 1 (chamado pela GUI, em thread propria)."""
    pythoncom.CoInitialize()
    try:
        try:
            session = connect_session()
        except Exception as e:
            raise ErroComTitulo(
                "Erro de conexão",
                f"Não consegui conectar ao SAP GUI.\n\nDetalhe: {e}\n\n"
                "Verifique se o SAP GUI está aberto e logado, e se o Scripting está habilitado "
                "(Alt+F12 > Opções > Acessibilidade e Scripting > Scripting).",
            ) from e
        try:
            abrir_ksb1(session, log)
        except RuntimeError as e:
            raise ErroComTitulo("Não consegui abrir a KSB1", f"{e}\n\nConfirme o acesso à KSB1 e tente de novo.") from e

        log(f"Extraindo KSB1 - {MESES_NOMES[mes]}/{ano} (Ciclo {ciclo})...")
        try:
            extrair_ksb1(session, mes, ano, ciclo, log)
        except ErroComTitulo:
            raise
        except Exception as e:
            raise ErroComTitulo("Erro durante a extração", str(e)) from e
        log("\nConcluído!")
    finally:
        pythoncom.CoUninitialize()
