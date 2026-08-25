#!/usr/bin/env python3
"""
Codigo compartilhado de conexao/navegacao no SAP GUI (transacao KSB1),
usado tanto por Fitted Units Despesas (fluxo mensal recorrente) quanto por
Fitted Recuperacao (extracao de periodo arbitrario) — movido pra ca em
2026-08-13 quando os dois sub-projetos foram separados em pastas.
"""
import time
from pathlib import Path

import win32com.client
import win32con
import win32gui
import win32process

BU = {"nome": "Fitted Units", "kstgr": "0495", "disvar": "/DESPFITTED"}

CICLOS = ("Actual", "Flash")

REDE_BASE = Path(
    r"\\FSS024-01BR.group.pirelli.com\GFU_DAC\Custos Fitted Units\Resultados Fitted"
)

MESES_NOMES = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}
MESES_PASTA = {
    1: "01 - Jan", 2: "02 - Feb", 3: "03 - Mar", 4: "04 - Apr",
    5: "05 - May", 6: "06 - Jun", 7: "07 - Jul", 8: "08 - Aug",
    9: "09 - Sep", 10: "10 - Oct", 11: "11 - Nov", 12: "12 - Dec",
}


def connect_session():
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(0)
    return connection.Children(0)


def _buscar_campo_editavel(elemento):
    # Busca recursiva pelo primeiro campo de valor editavel (GuiCTextField).
    # Necessario porque no popup "Definir área contab.custos" o campo fica
    # dentro de um subscreen (usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE),
    # nao direto em usr — descoberto rodando diagnosticar_popup.py (ver
    # memory/errors/2026-08-11_popup_area_contabil_ao_reentrar_sap.md).
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
    # Ao sair e voltar a entrar no SAP (nova sessao), a primeira transacao
    # do dia costuma abrir um popup modal "Definir área contab.custos"
    # pedindo a area contabil de custos antes de liberar a tela principal.
    # Se nao for fechado, os campos da tela de selecao da KSB1 (wnd[0])
    # ficam inacessiveis. Preenche com "0580" (fixo para Fitted Units, ver
    # memory/errors/2026-08-10_ksb1_kokrs_vazio.md) e confirma na seta verde.
    wnd1 = session.FindById("wnd[1]", False)
    if wnd1 is None:
        return

    log("Popup 'Definir área contab.custos' detectado, preenchendo 0580...")
    campo = _buscar_campo_editavel(wnd1.FindById("usr"))

    if campo is None:
        # Nao adivinha: clicar em "confirmar" sem preencher o campo faz o
        # proprio SAP abrir "Preencher todos os campos obrigatorios" e deixa
        # a sessao com popups empilhados (foi o que aconteceu em
        # memory/errors/2026-08-11_popup_area_contabil_ao_reentrar_sap.md).
        raise RuntimeError(
            "Não consegui identificar o campo do popup 'Definir área contab.custos' "
            "automaticamente. Feche os popups no SAP, abra o popup de novo e rode "
            "scripts/sap/fitted_units/_shared/ferramentas/diagnosticar_popup.py para "
            "descobrir o Id exato do campo."
        )

    campo.Text = "0580"
    wnd1.FindById("tbar[0]/btn[0]").Press()  # seta verde (confirmar)


def abrir_ksb1(session, log):
    # Sempre reenviamos /nKSB1, mesmo se a transacao atual ja for "KSB1":
    # a transacao permanece "KSB1" tanto na tela de selecao quanto na tela
    # de resultados apos rodar o relatorio, entao checar so o nome da
    # transacao nao garante que estamos na tela de selecao (isso causava
    # "The control could not be found by id" na segunda extracao).
    log("Abrindo a transação KSB1...")
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nKSB1"
    session.FindById("wnd[0]").SendVKey(0)  # Enter

    tratar_popup_area_contabil(session, log)

    if session.Info.Transaction != "KSB1":
        raise RuntimeError(
            f"Não consegui abrir a KSB1 (tela atual: '{session.Info.Transaction}')."
        )


def voltar_para_selecao(session, log):
    # Clica no botao "Voltar" (seta verde) em vez de SendVKey(3): o SendVKey
    # simula a tecla F3, que o SAP pode reportar como desabilitada dependendo
    # do estado da tela ("The virtual key is not enabled"), mesmo com o botao
    # visualmente clicavel. Pressionar o botao direto evita esse problema e
    # mantem os campos da tela de selecao preenchidos (diferente de reabrir a
    # transacao do zero com /nKSB1, que limpa tudo).
    log("Voltando para a tela de seleção...")
    wnd = session.FindById("wnd[0]")
    wnd.FindById("tbar[0]/btn[3]").Press()

    if wnd.FindById("usr/ctxtP_KOKRS", False) is None:
        # Nao caiu na tela de selecao esperada: reabre a transacao do zero
        # como rede de seguranca.
        abrir_ksb1(session, log)


def resolver_pasta_ciclo(pasta_mes: Path, mes: int, ciclo: str) -> Path:
    """Acha a subpasta do Ciclo (Actual/Flash/Forecast) dentro da pasta do mes
    (ex: '.../03 - Mar/'), tolerando o mes por extenso no nome da subpasta em
    vez da abreviacao de 3 letras padrao - inconsistencia real encontrada em
    03_March_Actual/04_April_Actual (deveriam ser 03_Mar_Actual/04_Apr_Actual,
    como todos os outros meses de 2026). Prefere o nome padrao exato; se nao
    existir, cai para qualquer subpasta existente que bata com '<MM>_*_<Ciclo>'.
    Se nada existir (mes/ciclo novo, pasta ainda nao criada), devolve o
    caminho padrao mesmo assim - quem chama decide se cria ou reporta erro."""
    abrev = MESES_PASTA[mes].split(" - ")[1]
    padrao = pasta_mes / f"{mes:02d}_{abrev}_{ciclo}"
    if padrao.exists():
        return padrao

    candidatos = sorted(pasta_mes.glob(f"{mes:02d}_*_{ciclo}")) if pasta_mes.exists() else []
    if len(candidatos) == 1:
        return candidatos[0]
    if len(candidatos) > 1:
        raise RuntimeError(
            f"Mais de uma pasta do Ciclo '{ciclo}' encontrada em {pasta_mes}: "
            f"{[c.name for c in candidatos]}"
        )
    return padrao


def prefixo_arquivo_ksb1(bu_nome: str, mes: int, ano: int, agrup_label: str) -> str:
    return f"KSB1 - {bu_nome} {mes:02d}.{ano} - {agrup_label}"


def nome_arquivo_ksb1(bu_nome: str, mes: int, ano: int, agrup_label: str, ciclo: str) -> str:
    # Ciclo faz parte do nome desde 2026-08-21 (antes disso, o Passo 3 pegava
    # a extracao mais recente por data de modificacao em vez de casar com o
    # Ciclo pedido - se a usuaria extraisse Flash e depois Actual no mesmo
    # mes, regerar o Flash pegava por engano os dados do Actual).
    return f"{prefixo_arquivo_ksb1(bu_nome, mes, ano, agrup_label)} - {ciclo}.XLSX"


def encontrar_arquivo_ksb1(pasta: Path, bu_nome: str, mes: int, ano: int, agrup_label: str, ciclo: str) -> Path:
    """Acha o arquivo bruto da KSB1 (Gestoriais/Sem Agrupamento) do mes/ano,
    do Ciclo pedido (ver nome_arquivo_ksb1). Para meses extraidos antes da
    mudanca de 2026-08-21 (sem Ciclo no nome do arquivo), cai para o arquivo
    mais recente com o prefixo antigo - mas nunca escolhe um arquivo que
    pertenca claramente a OUTRO Ciclo, para nao repetir o bug que motivou
    essa mudanca."""
    prefixo_base = prefixo_arquivo_ksb1(bu_nome, mes, ano, agrup_label)
    prefixo_ciclo = f"{prefixo_base} - {ciclo}"

    candidatos_ciclo = sorted(pasta.glob(f"{prefixo_ciclo}*.XLSX"), key=lambda p: p.stat().st_mtime)
    if candidatos_ciclo:
        return candidatos_ciclo[-1]

    prefixos_outros_ciclos = tuple(f"{prefixo_base} - {c}" for c in CICLOS if c != ciclo)
    candidatos_antigos = [
        p for p in pasta.glob(f"{prefixo_base}*.XLSX")
        if not p.stem.startswith(prefixos_outros_ciclos)
    ]
    if not candidatos_antigos:
        raise FileNotFoundError(
            f"Não encontrei nenhum arquivo do Ciclo '{ciclo}' (nem versão antiga sem Ciclo "
            f"no nome) começando com '{prefixo_base}' em {pasta}"
        )
    candidatos_antigos.sort(key=lambda p: p.stat().st_mtime)
    return candidatos_antigos[-1]


def localizar_extracao_ksb1(pasta_mes: Path, bu_nome: str, mes: int, ano: int, agrup_label: str, ciclo: str) -> Path:
    """Acha o arquivo bruto da extracao (Passo 1) do mes/ano/Ciclo/agrupamento.
    Desde 2026-08-24, a extracao passou a salvar dentro de uma subpasta do
    Ciclo (<MM>_<Mes3>_<Ciclo>/, mesmo padrao ja usado pelos Passos 3/4 - ver
    resolver_pasta_ciclo) em vez de solta direto na pasta do mes. Meses ja
    extraidos antes dessa mudanca (arquivos soltos na pasta do mes) continuam
    funcionando sem reorganizar nada: procura primeiro na subpasta do Ciclo;
    se nao achar nada la, cai para a pasta do mes (formato antigo) - decisao
    explicita da usuaria de nao mover os arquivos ja existentes."""
    pasta_ciclo = resolver_pasta_ciclo(pasta_mes, mes, ciclo)
    try:
        return encontrar_arquivo_ksb1(pasta_ciclo, bu_nome, mes, ano, agrup_label, ciclo)
    except FileNotFoundError:
        try:
            return encontrar_arquivo_ksb1(pasta_mes, bu_nome, mes, ano, agrup_label, ciclo)
        except FileNotFoundError as erro_pasta_mes:
            raise FileNotFoundError(
                f"Não encontrei a extração de '{agrup_label}' (Ciclo {ciclo}) nem em "
                f"{pasta_ciclo} nem em {pasta_mes} (formato antigo, sem subpasta)."
            ) from erro_pasta_mes


def encontrar_arquivo_mais_recente(pasta: Path, nome_base: str) -> Path | None:
    """Acha a versão mais recente de um arquivo gerado via nome_com_versao
    (nome_base.xlsx, nome_base_v2.xlsx, nome_base_v3.xlsx, ...) — usada
    pelos passos que LEEM um arquivo gerado por um passo anterior do mesmo
    fluxo (ex: Finalização lê o BASE_KSB1 que "Atualizar Pivot KSB1" acabou
    de gerar, ou o passo do mês seguinte lê o Actual do mês anterior).
    Achado real em 2026-08-22: essas leituras buscavam só o nome exato
    (sem "_v2" etc.), então se o passo anterior fosse rodado de novo (ex:
    pra corrigir algo), o resultado corrigido ficava "invisível" pros
    passos seguintes, que continuavam lendo a versão antiga em silêncio.
    Devolve None se nenhuma versão existir — quem chama decide o erro."""
    base = Path(nome_base)
    stem, ext = base.stem, base.suffix
    candidatos = list(pasta.glob(f"{stem}*{ext}")) if pasta.exists() else []
    if not candidatos:
        return None
    return max(candidatos, key=lambda p: p.stat().st_mtime)


def abrir_excel_isolado(log=print, pid_callback=None):
    """Abre uma instancia isolada e invisivel do Excel (DispatchEx - nao
    interfere com o Excel que a usuaria tiver aberto) e captura o PID do
    processo EXCEL.EXE correspondente, via o Hwnd da propria instancia
    (existe mesmo com Visible=False). Se pid_callback for passado, e'
    chamado com esse PID assim que capturado - usado pelo watchdog de
    travamento da GUI (rodar_em_thread em atualizar_ksb1_gui.py) pra saber
    qual processo encerrar se essa instancia travar de verdade."""
    log("Abrindo Excel (instância isolada, oculta)...")
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    if pid_callback is not None:
        try:
            pid = win32process.GetWindowThreadProcessId(excel.Hwnd)[1]
        except Exception:
            pid = None
        pid_callback(pid)
    return excel


def fechar_excel_se_aberto(caminho_arquivo: Path, log=print) -> bool:
    """Algumas configuracoes de exportacao nativa do SAP (Lista > Exportar >
    Planilha eletronica) abrem o arquivo gerado automaticamente no Excel logo
    depois de salvar - isso trava o arquivo (WinError 32) se o codigo tentar
    mover/renomear ele em seguida.

    GetObject(Class='Excel.Application') so' enxerga UMA instancia (ambigua
    se a usuaria tiver o Excel dela mesma aberto ao mesmo tempo) - em vez
    disso, cada pasta de trabalho aberta se registra na Running Object Table
    (COM) com o proprio caminho completo como nome. Por isso a busca aqui e'
    direto por esse nome (via pythoncom), o que acha a planilha certa mesmo
    com varias instancias do Excel rodando, sem precisar adivinhar qual
    Application pegar - e sem mexer em nenhuma outra planilha aberta.

    Tudo aqui e' "melhor esforco": se o Excel estiver ocupado (ex: mostrando
    um dialogo modal proprio, tipo "nao encontrei o arquivo") as chamadas de
    COM podem falhar - nesse caso so' desiste e deixa a retentativa de mover
    (no chamador) continuar tentando; nao e' garantido que isso feche o
    Excel de verdade, so' aumenta a chance. Retorna True se fechou algo."""
    import pythoncom

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
            log(f"O SAP abriu '{Path(caminho_arquivo).name}' automaticamente no Excel - fechando pra liberar o arquivo...")
            wb.Close(SaveChanges=False)
            return True
        except Exception:
            continue
    return False


def limpar_excel_orfao(log=print):
    """Mesmo depois de fechar_excel_se_aberto liberar o arquivo, o Excel as
    vezes ainda mostra um aviso nativo ("Sorry, we couldn't find ...") -
    porque o proprio codigo moveu o arquivo de onde o Excel tinha aberto
    ele - e fica pra tras uma janela vazia (sem nenhuma pasta de trabalho
    aberta). Fecha os dois:
    1. Clica "OK" em qualquer dialogo nativo do Windows titulado
       "Microsoft Excel" (o titulo padrao desse aviso especifico).
    2. Fecha (Quit) qualquer instancia do Excel que nao tenha NENHUMA pasta
       de trabalho aberta - nunca mexe numa instancia com algo aberto, entao
       nunca fecha um Excel que a usuaria esteja usando de verdade."""

    def _clicar_ok_se_for_o_aviso(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        if win32gui.GetClassName(hwnd) != "#32770":  # classe padrao de dialogo do Windows
            return True
        if win32gui.GetWindowText(hwnd).strip() != "Microsoft Excel":
            return True

        def _achar_botao_ok(h, _):
            texto = win32gui.GetWindowText(h)
            if win32gui.GetClassName(h) == "Button" and texto.strip().upper() == "OK":
                win32gui.PostMessage(h, win32con.BM_CLICK, 0, 0)
            return True

        try:
            win32gui.EnumChildWindows(hwnd, _achar_botao_ok, None)
            log("Fechei um aviso do Excel que sobrou depois da extração.")
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(_clicar_ok_se_for_o_aviso, None)
    except Exception:
        pass

    try:
        excel = win32com.client.GetObject(Class="Excel.Application")
        if excel.Workbooks.Count == 0:
            excel.Quit()
            log("Fechei uma janela do Excel que ficou vazia depois da extração.")
    except Exception:
        pass


def nome_com_versao(pasta: Path, nome_base: str) -> str:
    # Nunca sobrescrever um arquivo ja existente na pasta: se ja existe um
    # arquivo com esse nome, salva como "_v2", se "_v2" tambem ja existir,
    # "_v3", e assim por diante.
    base = Path(nome_base)
    stem, ext = base.stem, base.suffix
    candidato = pasta / nome_base
    versao = 2
    while candidato.exists():
        candidato = pasta / f"{stem}_v{versao}{ext}"
        versao += 1
    return candidato.name
