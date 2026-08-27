#!/usr/bin/env python3
"""
Janela grafica para extrair a KSB1 da Fitted Units (Gestoriais + Sem Agrupamento),
salvar os arquivos na area de rede e conferir o agrupamento gestorial.

Rodado sempre localmente (nao precisa ser copiado para a rede — so o atalho/
.bat na rede que aponta pra ca). Depende de check_agrupamentos_ksb1.py, na
mesma pasta, pro botao "Gerar Check de Agrupamentos".

Pre-requisitos na maquina de quem for rodar:
- Python 3 com pywin32 e openpyxl instalados (pip install -r requirements.txt)
- SAP GUI Scripting habilitado (Alt+F12 > Opcoes > Acessibilidade e Scripting > Scripting)
- SAP GUI aberto e logada (nao precisa estar na KSB1, o script abre a transacao sozinho)
"""
import ctypes
import math
import queue
import shutil
import sys
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk

import psutil
import pythoncom
from PIL import Image, ImageDraw, ImageTk

# Sem isso, o Windows nao sabe que o Tkinter lida com DPI sozinho e "estica"
# a janela como bitmap pra bater com o zoom da tela (125%/150% etc.) — e' o
# que deixa o texto borrado. Precisa rodar antes de qualquer janela do Tk
# ser criada. Tentativa em cascata (API mais nova -> mais antiga) porque
# SetProcessDpiAwareness so existe a partir do Windows 8.1 (shcore.dll).
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  # PROCESS_SYSTEM_DPI_AWARE
except (AttributeError, OSError):
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except (AttributeError, OSError):
        pass

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "_shared"))
from ksb1_core import (  # noqa: E402
    BU,
    MESES_NOMES,
    MESES_PASTA,
    REDE_BASE,
    abrir_ksb1,
    connect_session,
    fechar_excel_se_aberto,
    limpar_excel_orfao,
    nome_arquivo_ksb1,
    nome_com_versao,
    resolver_pasta_ciclo,
    voltar_para_selecao,
)

# Logo da Pirelli embutido (base64) para o script continuar autossuficiente
# (nao depender de arquivo de imagem separado ao copiar para a rede).
LOGO_PIRELLI_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAJsAAAAvCAYAAAD0OrjvAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAU4SURBVHhe7Zvri1VlFIcnm5qpbLLLt7KLlpk5Y4ZI6nwIDE0s8MOEgiRFN0UKTJQUu2AYo0KW1HRRpMuoYVaa1hBdUCu7Rx+ke/0rb8+PMweOb2vfztTrmWF9eEBnr7X23u969n735ey28HdbcJwUuGxOMlw2Jxkum5MMl81JhsvmJMNlc5LhsjnJcNmcZFSX7c2zQrjvvBD6upyxzLILQ3ikI4Q99PynyIEmqS7bdmS7iQ25/DJnLDMRrrs4hDmIt7YzhC8NFyrisjn5XAHTkW4js9nPhg8VSCPbVZeGsJSNffQcJxVr4EGYx7hbPamChFs0PoRDhg8VSCPb5EtCGDg3hB/Id9LxDbzaHsKNE+y+VGEWPd9BPcuJkqSR7Vpkex3ZrHrO/8vRcYjCWcnqSxXU836jfgXSyHYN0+gzyDZEvpOGD+EtWMVUqssYqy9VmEvPX6Ge5URJ0simOf8Gzm4zuNB00tENmlWsnlSlj55/bvhQgTSyOaOb3q4QdtH3Pw0fKuCyOdlM4qy45PwQ9tHz3wwXKpLumm1DRwiD5DqjgwPwGZyCPwwPmiCNbHr08Ro3CDoNO6OHv8ByoEnSyOaPPhw4s7I9BYv4+/xOm6Fx/87ZVJBThsXtIWykzteQdfTeRcxtRm5ZFsL+qKamo/2MnxUvFnCpsSrKyeIkLGd8rDribsbo4yjH4ivYkFNHbyE+inKa5MzKtpKLT02xVo54hx2Nc+6nzmRu6a34qkyfEMJBBtq6JrmZZXpkY+WV4WquU3dFNX+Hl1mfFS+uJGfx+NNzsjhGnbmMn1VH9F4UwvtGXsxxuIcxtWqIhfT6cJTTJGNPNj3AVKPr6P950sykKSfZp/gMVySbxFDtLLTPu6OaLpu9IJNWl63/7BD2ErdvmGdhNgOvn8xY8aKfxv1KXON6imR7gJjnctgJJ6CxpstmL8ik1WU7gmy6k2rM2Q5Tc3LupXGnopwi2V4gRtNvHvHZ0mWzF2TS6rK9i2z63ZXOVEL/fprGXJ+znhVNyLaenLdz0LXg91FNl81ekEmry/YYOc/TiJ3DPAkzcsTR3zcT80u0Hr9BqOGy5chWFb2w/pR9iqdel62Gy/YfyKbGzqTOAE2z3vsVyaZfVNySQy/5g1FNl81ekEmryzaLZfNo9DSYSPOsGDGNuD0Zooki2bYR810B8dTsstkLMml12Qa5QdDPoR+HvDtQPQsbIDZ+5FGnSLatxHxbgIRrlLmMbAsYW6tWnfpndUWyzUG2NxryYn4EbZvLNkwzstUffehV1Hy2Uw204sRsziJDxvWaKJLtDrZtZQEPt5/+rK1INq2vh/VatcRq6m1j/1SrSLap1Flq1KjzBHWOUcdlG2YksilWU52mSyuuzoqO2pFe9Q1CGSYheuNUViRbEbrp6LugVqtItiJu7QrhIHVaWrYdyKadnEITy9LDKX0vR2VcS19cd7PMyhGHjJzVCNjDQFnx4oMG2fSF0Z2cvXSUW7FCU+2L7FN87dbLIE/JyStDN/lHG2pKtt2sy4otg65Dl3fWah2njl6UW3FluJ1xeY86+qn3Q4ypFSOWIPcRYur7MAKqy6YN3ERD13A0lGUd0nzC4MS19O5Qy6wcccLI0UcXeTnxe05dt6xnMK3YOluIiS/mNc3o20srvizryG/8klwHgT5GsWLLsJZ6ehWmWjqQNrOvVlwZtjCGmuJV5yXOklaM2ErcF8PrHCHVZXOcJnHZnGS4bE4yXDYnGS6bkwyXzUmGy+Ykw2VzkuGyOclw2ZxEtIV/AAlejwcdSLwvAAAAAElFTkSuQmCC"
)


def _gerar_frames_pneu(diametro=18, n_frames=12):
    """Gera os frames (ImageTk.PhotoImage) de um icone de pneu com calota
    (banda preta com marcas de sulco, aro prateado com raios e miolo escuro
    com detalhe amarelo claro - cores do cockpit) girando, usado como
    indicador de "processando" na barra abaixo do cabecalho. Desenhado em
    runtime via PIL em vez de arquivo/base64 separado, pra manter o script
    autossuficiente - mesma filosofia do logo Pirelli embutido acima."""
    escala = 8  # desenha bem maior e reduz depois (raios da calota ficam limpos)
    d_grande = diametro * escala
    img = Image.new("RGBA", (d_grande, d_grande), (0, 0, 0, 0))
    desenho = ImageDraw.Draw(img)
    raio = d_grande / 2
    centro = (raio, raio)

    # Banda do pneu (preta, com sulcos)
    desenho.ellipse(
        [d_grande * 0.04, d_grande * 0.04, d_grande * 0.96, d_grande * 0.96],
        fill=(20, 20, 20, 255),
    )
    for i in range(12):
        ang = math.radians(i * 30)
        x1 = centro[0] + raio * 0.87 * math.cos(ang)
        y1 = centro[1] + raio * 0.87 * math.sin(ang)
        x2 = centro[0] + raio * 0.70 * math.cos(ang)
        y2 = centro[1] + raio * 0.70 * math.sin(ang)
        desenho.line([x1, y1, x2, y2], fill=(60, 60, 60, 255), width=max(1, escala // 3))

    # Calota (aro prateado com raios, miolo escuro com detalhe amarelo claro)
    raio_calota = raio * 0.62
    desenho.ellipse(
        [centro[0] - raio_calota, centro[1] - raio_calota, centro[0] + raio_calota, centro[1] + raio_calota],
        fill=(196, 199, 204, 255), outline=(120, 122, 126, 255), width=max(1, escala // 4),
    )
    n_raios = 6
    largura_raio = math.radians(10)
    for i in range(n_raios):
        ang = math.radians(i * (360 / n_raios))
        pontos = []
        for delta in (-largura_raio, largura_raio):
            a = ang + delta
            pontos.append((centro[0] + raio_calota * 0.94 * math.cos(a), centro[1] + raio_calota * 0.94 * math.sin(a)))
        pontos.insert(1, (
            centro[0] + raio_calota * 0.94 * math.cos(ang), centro[1] + raio_calota * 0.94 * math.sin(ang)
        ))
        desenho.polygon([centro, pontos[0], pontos[1], pontos[2]], fill=(146, 149, 155, 255))

    raio_miolo = raio_calota * 0.34
    desenho.ellipse(
        [centro[0] - raio_miolo, centro[1] - raio_miolo, centro[0] + raio_miolo, centro[1] + raio_miolo],
        fill=(30, 30, 30, 255),
    )
    raio_logo = raio_miolo * 0.4
    desenho.ellipse(
        [centro[0] - raio_logo, centro[1] - raio_logo, centro[0] + raio_logo, centro[1] + raio_logo],
        fill=(255, 233, 168, 255),
    )

    frames = []
    for i in range(n_frames):
        rotacionado = img.rotate(-i * (360 / n_frames), resample=Image.BICUBIC)
        reduzido = rotacionado.resize((diametro, diametro), Image.LANCZOS)
        frames.append(ImageTk.PhotoImage(reduzido))
    return frames


def extrair_um(session, mes, ano, ciclo, koagr, agrup_label, log):
    import calendar

    ultimo_dia = calendar.monthrange(ano, mes)[1]
    data_de = f"01.{mes:02d}.{ano}"
    data_ate = f"{ultimo_dia:02d}.{mes:02d}.{ano}"

    wnd = session.FindById("wnd[0]")
    wnd.FindById("usr/ctxtP_KOKRS").Text = "0580"
    if wnd.FindById("usr/ctxtKSTGR", False) is not None:
        wnd.FindById("usr/ctxtKSTGR").Text = BU["kstgr"]
    wnd.FindById("usr/ctxtKOAGR").Text = koagr
    wnd.FindById("usr/ctxtR_BUDAT-LOW").Text = data_de
    wnd.FindById("usr/ctxtR_BUDAT-HIGH").Text = data_ate
    wnd.FindById("usr/ctxtP_DISVAR").Text = BU["disvar"]

    log(f"Executando KSB1 ({agrup_label}, Ciclo {ciclo})...")
    session.FindById("wnd[0]").SendVKey(8)

    # Desde 2026-08-24, a extracao salva dentro de uma subpasta do Ciclo
    # (<MM>_<Mes3>_<Ciclo>/, mesmo padrao ja usado pelos Passos 3/4 - ver
    # resolver_pasta_ciclo), em vez de solta direto na pasta do mes. Meses
    # ja extraidos antes disso continuam soltos na pasta do mes (nao foram
    # reorganizados) - o Passo 2/3 sabem procurar nos dois lugares.
    pasta_mes = REDE_BASE / str(ano) / "00.Extração Base KSB1" / MESES_PASTA[mes]
    pasta_rede = resolver_pasta_ciclo(pasta_mes, mes, ciclo)
    pasta_rede.mkdir(parents=True, exist_ok=True)
    nome_arquivo = nome_com_versao(
        pasta_rede, nome_arquivo_ksb1(BU["nome"], mes, ano, agrup_label, ciclo)
    )

    # Desde 2026-08-25: o SAP sempre exporta pra essa pasta "Temporario" fixa
    # (nunca muda de mes/Ciclo/ano dentro do mesmo ano), em vez de exportar
    # direto na subpasta do Ciclo (que muda todo mes). O popup "Seguranca
    # SAPGUI" so aparece quando o SAP escreve numa pasta ainda nao autorizada
    # em saprules.xml - como agora ele so escreve nessa pasta fixa, so' pede
    # autorizacao 1x (a primeira vez, por ano) em vez de 1x por mes/Ciclo.
    # A realocacao pra pasta final (pasta_rede) e' feita com shutil.move, uma
    # copia de arquivo comum que nao passa pelo SAP GUI Scripting - por isso
    # nao e' vigiada pela seguranca do SAP e nao gera popup nenhum.
    pasta_staging = REDE_BASE / str(ano) / "00.Extração Base KSB1" / "Temporario"
    pasta_staging.mkdir(parents=True, exist_ok=True)
    arquivo_staging = pasta_staging / nome_arquivo
    if arquivo_staging.exists():
        # Sobra de uma tentativa anterior que falhou antes de mover - remove
        # pra nao correr risco do SAP perguntar "sobrescrever?" (dialogo que
        # o script nao trata) na hora de exportar de novo. Se ainda estiver
        # travada (Excel demorando pra soltar), tenta de novo por ate 30s em
        # vez de derrubar a extracao inteira com um erro.
        for _ in range(30):
            try:
                arquivo_staging.unlink()
                break
            except OSError:
                fechar_excel_se_aberto(arquivo_staging, log)
                time.sleep(1)
        else:
            raise ErroComTitulo(
                "Arquivo temporário travado",
                f"'{arquivo_staging.name}' ainda está aberto em algum programa (provavelmente Excel) "
                f"e não consegui apagar a sobra da tentativa anterior em {pasta_staging}.\n\n"
                "Feche o arquivo e tente de novo.",
            )

    session.FindById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select()
    wnd1 = session.FindById("wnd[1]")
    wnd1.FindById("usr/ctxtDY_PATH").Text = str(pasta_staging)
    wnd1.FindById("usr/ctxtDY_FILENAME").Text = nome_arquivo
    wnd1.FindById("tbar[0]/btn[0]").Press()  # Gerar

    for _ in range(20):
        if arquivo_staging.exists():
            break
        time.sleep(0.5)

    arquivo_final = pasta_rede / nome_arquivo
    if arquivo_staging.exists():
        # O SAP as vezes abre o arquivo recem-exportado automaticamente no
        # Excel (comportamento padrao dessa exportacao em alguns ambientes),
        # o que trava o arquivo (WinError 32) na hora de mover. Tenta mover
        # com retentativa; se continuar travado, fecha so essa aba especifica
        # (sem mexer em outra planilha aberta) e tenta de novo.
        movido = False
        for tentativa in range(30):
            try:
                shutil.move(str(arquivo_staging), str(arquivo_final))
                movido = True
                break
            except OSError:
                fechar_excel_se_aberto(arquivo_staging, log)
                time.sleep(1)
        if movido:
            log(f"{agrup_label}: salvo em {arquivo_final}")
            # O Excel que o SAP abriu automaticamente as vezes ainda mostra um
            # aviso ("nao encontrei o arquivo") so' porque o arquivo que ele
            # tinha aberto acabou de ser movido daqui pra pasta final - efeito
            # colateral inofensivo. Fecha o aviso e a janela vazia que sobra,
            # tentando algumas vezes (o aviso pode demorar um pouco a aparecer).
            for _ in range(5):
                time.sleep(1)
                limpar_excel_orfao(log)
        else:
            log(
                f"AVISO: {agrup_label} ficou travado na pasta temporária "
                f"({arquivo_staging}) e não pôde ser movido pra pasta final. "
                f"Feche o arquivo se estiver aberto e mova manualmente pra {arquivo_final}."
            )
    else:
        log(f"AVISO: não encontrei o arquivo de {agrup_label} na pasta temporária. Confira manualmente.")

    voltar_para_selecao(session, log)


class ErroComTitulo(Exception):
    """Erro com titulo proprio pro messagebox (a mensagem de erro do SAP
    sozinha, sem contexto, nao ajuda a usuaria a saber o que fazer). Usado
    pra 'rodar' poder levantar excecao em vez de chamar messagebox direto -
    precisa ser assim pra rodar numa thread separada (ver rodar_em_thread em
    main() - so' a thread principal do Tk pode mostrar messagebox)."""

    def __init__(self, titulo, mensagem):
        super().__init__(mensagem)
        self.titulo = titulo
        self.mensagem = mensagem


def rodar(mes, ano, ciclo, log):
    try:
        session = connect_session()
    except Exception as e:
        raise ErroComTitulo(
            "Erro de conexão",
            f"Não consegui conectar ao SAP GUI.\n\nDetalhe: {e}\n\n"
            "Verifique se o SAP GUI está aberto e logada, e se o Scripting está "
            "habilitado (Alt+F12 > Opções > Acessibilidade e Scripting > Scripting).",
        ) from e

    try:
        abrir_ksb1(session, log)
    except RuntimeError as e:
        raise ErroComTitulo(
            "Não consegui abrir a KSB1",
            f"{e}\n\nConfirme se seu usuário tem acesso à transação KSB1 e tente de novo.",
        ) from e

    log(f"Extraindo KSB1 - {MESES_NOMES[mes]}/{ano} (Ciclo {ciclo})...")
    try:
        extrair_um(session, mes, ano, ciclo, "gestoriais", "Gestoriais", log)
        extrair_um(session, mes, ano, ciclo, "", "Sem Agrupamento", log)
    except Exception as e:
        raise ErroComTitulo("Erro durante a extração", str(e)) from e

    log("\nConcluído!")


AMARELO_CLARO = "#FFE9A8"
CINZA_TEXTO = "#555555"

# Watchdog de travamento (ver rodar_em_thread): se uma operacao ficar rodando
# mais tempo que isso sem terminar, avisa que pode estar travada. 12 min foi
# escolhido pra dar folga a operacoes grandes (colagem linha a linha em meses
# com muitas linhas), sem deixar a usuaria esperando longe demais sem
# feedback - confirmado com ela em 2026-08-24.
TIMEOUT_AVISO_SEGUNDOS = 12 * 60

# Paleta "cockpit": cabecalho escuro com logo Pirelli (trim vermelho/amarelo);
# corpo abaixo do trim em fundo branco/letras pretas, a pedido da usuaria.
BG_ROOT = "#0b0c0e"
BG_PAINEL = "#ffffff"
BG_CARD = "#ffffff"
BG_CAMPO = "#f0f0f2"
BORDA = "#d5d6d9"
TEXTO_CLARO = "#111111"
TEXTO_SECUNDARIO = "#5a5c60"
LOG_BG = "#ffffff"
LOG_FG = "#111111"

PASSOS = [
    {
        "aba": "①  Extração",
        "titulo": "Passo 1 · Extrair KSB1",
        "descricao": (
            "Baixa a KSB1 direto do SAP (Gestoriais + Sem Agrupamento) pro mês/ano/Ciclo "
            "escolhidos e salva os dois arquivos na área de rede, já identificados com o "
            "Ciclo no nome. Pré-requisito: SAP GUI aberto e logado na tela inicial (o "
            "script abre a transação sozinho)."
        ),
        "botoes": ["Extrair KSB1 (Gestoriais + Sem Agrupamento)"],
    },
    {
        "aba": "②  Check de Agrupamentos",
        "titulo": "Passo 2 · Check de Agrupamentos",
        "descricao": (
            "Confere se toda conta contábil do Sem Agrupamento está vinculada a um "
            "agrupamento gestorial, comparando com o arquivo Gestoriais do mesmo mês/Ciclo. "
            "Usa os arquivos já extraídos no Passo 1 — não acessa o SAP."
        ),
        "botoes": ["Gerar Check de Agrupamentos"],
    },
    {
        "aba": "③  Provisões",
        "titulo": "Passo 3 · Provisões",
        "descricao": (
            "Só pro Ciclo Flash: cria a Base Intermediária Flash do mês (a partir do Actual "
            "do mês anterior) e preenche as linhas coloridas com as provisões/reclassificações "
            "do 'Fast Provisão' mais recente da pasta de rede. 'Lançar Provisões' cria o "
            "arquivo pela primeira vez; 'Atualizar Provisões' relê o Fast Provisão (ex: depois "
            "de uma correção) e atualiza um arquivo já criado. O Fast Provisão precisa estar "
            "fechado e salvo antes de rodar qualquer um dos dois."
        ),
        "botoes": ["Lançar Provisões", "Atualizar Provisões"],
    },
    {
        "aba": "④  Base Intermediária",
        "titulo": "Passo 4 · Atualizar KSB1 Pivot",
        "descricao": (
            "Atualiza o KSB1 acumulado do ano (BASE_KSB1 + Pivot Tables nativas) com "
            "as linhas do mês e prepara os valores pra colar na Base Intermediária. "
            "Usa o Ciclo selecionado (Actual/Flash) pra escolher a extração certa do "
            "Passo 1 e também no nome do arquivo final. Depois de atualizar o Pivot, use "
            "'Finalização da Base Intermediária' pra colar os valores na Intermediária "
            "(no Flash, rode o Passo 3 — Lançar Provisões — antes)."
        ),
        "botoes": ["Atualizar Pivot KSB1", "Finalização da Base Intermediária"],
    },
    {
        "aba": "⑤  Rateio de Custos",
        "titulo": "Passo 5 · Rateio de Custos",
        "descricao": (
            "Ainda em validação com a usuária. 'Abertura de Custos por Unidade' lê a "
            "Base Intermediária do mês/Ciclo (Passo 4 já rodado) e gera o arquivo de "
            "Rateio de Custos, com o check por unidade e a aba 'Comentários'. "
            "'Atualizar Rateio' abre um quadro editável com o % vigente pra cada "
            "unidade — se você não mudar num Ciclo, o rateio anterior continua valendo."
        ),
        "botoes": ["Abertura de Custos por Unidade", "Atualizar Rateio"],
    },
    {
        "aba": "⑥  Mensalização",
        "titulo": "Passo 6 · Mensalização",
        "descricao": (
            "Gera o arquivo de Mensalização - copia a base certa (Forecast do mês, ou "
            "o Flash já fechado quando o Ciclo for Actual), aplica os ajustes de cenário "
            "quando necessário, e cola os valores do Passo 5 (Rateio de Custos) na "
            "coluna do mês sendo fechado. 'Atualizar Faturamento' ainda não foi "
            "automatizado (Net Sales continua manual por enquanto)."
        ),
        "botoes": ["Atualizar Faturamento", "Atualizar Custo"],
    },
]


def _configurar_estilo(root):
    style = ttk.Style()
    style.theme_use("clam")

    style.configure("TFrame", background=BG_PAINEL)
    style.configure("Card.TFrame", background=BG_CARD)
    style.configure("TLabel", background=BG_PAINEL, foreground=TEXTO_CLARO, font=("Segoe UI", 10))
    style.configure("Card.TLabel", background=BG_CARD, foreground=TEXTO_CLARO, font=("Segoe UI", 10))
    style.configure(
        "Titulo.TLabel", background=BG_CARD, foreground=TEXTO_CLARO,
        font=("Segoe UI", 14, "bold"),
    )
    style.configure(
        "Descricao.TLabel", background=BG_CARD, foreground=TEXTO_SECUNDARIO,
        font=("Segoe UI", 10), wraplength=620,
    )

    style.configure(
        "TCombobox",
        fieldbackground=BG_CAMPO, background=BG_CAMPO, foreground=TEXTO_CLARO,
        arrowcolor=TEXTO_CLARO, bordercolor=BORDA, lightcolor=BG_CAMPO, darkcolor=BG_CAMPO,
    )
    style.map(
        "TCombobox",
        fieldbackground=[("readonly", BG_CAMPO)],
        foreground=[("readonly", TEXTO_CLARO)],
    )
    style.configure(
        "TEntry",
        fieldbackground=BG_CAMPO, foreground=TEXTO_CLARO, insertcolor=TEXTO_CLARO,
        bordercolor=BORDA, lightcolor=BG_CAMPO, darkcolor=BG_CAMPO,
    )
    style.configure("Pirelli.TButton", font=("Segoe UI", 11, "bold"), foreground="black", borderwidth=0)
    style.map(
        "Pirelli.TButton",
        background=[("!disabled", AMARELO_CLARO), ("disabled", "#ecdfb0")],
        foreground=[("!disabled", "black"), ("disabled", "#8a8a8a")],
    )

    # Notebook (abas) — tema "clam" permite recolorir tab a tab, o padrao do
    # Windows (vista/xpnative) ignora essas cores.
    style.configure("TNotebook", background=BG_PAINEL, borderwidth=0, tabmargins=(8, 8, 8, 0))
    style.configure(
        "TNotebook.Tab", background=BG_CARD, foreground=TEXTO_SECUNDARIO,
        font=("Segoe UI", 10, "bold"), padding=(18, 10), borderwidth=0,
    )
    style.map(
        "TNotebook.Tab",
        background=[("selected", AMARELO_CLARO), ("!selected", BG_CARD)],
        foreground=[("selected", "black"), ("!selected", TEXTO_SECUNDARIO)],
        expand=[("selected", (0, 0, 0, 0))],
    )

    # Combobox usa listas suspensas nativas do Tk (nao ttk) — precisam ser
    # coloridas separadamente, senao ficam brancas mesmo com o tema escuro.
    root.option_add("*TCombobox*Listbox.background", BG_CAMPO)
    root.option_add("*TCombobox*Listbox.foreground", TEXTO_CLARO)
    root.option_add("*TCombobox*Listbox.selectBackground", AMARELO_CLARO)
    root.option_add("*TCombobox*Listbox.selectForeground", "black")


def main():
    root = tk.Tk()
    root.title("Fitted Units · Cockpit Fechamento")
    root.geometry("1317x800")
    root.minsize(1000, 650)
    root.configure(bg=BG_ROOT)
    # Abrir sempre maximizada (pedido explicito da usuaria, 2026-08-26) -
    # "zoomed" e' o estado nativo do Windows (maximiza mantendo bordas/
    # barra de titulo, diferente de "-fullscreen" que tira tudo isso).
    root.state("zoomed")

    logo_img = tk.PhotoImage(data=LOGO_PIRELLI_B64)
    root.iconphoto(True, logo_img)

    _configurar_estilo(root)

    # --- Cabecalho ---------------------------------------------------
    # Altura NAO e' fixa em pixels (sem pack_propagate(False)) — com DPI
    # awareness a fonte do titulo renderiza no tamanho real, e uma altura
    # travada em pixels espreme o subtitulo contra a linha vermelha.
    header = tk.Frame(root, bg=BG_ROOT)
    header.pack(fill=tk.X, side=tk.TOP)
    tk.Label(header, image=logo_img, bg=BG_ROOT).pack(side=tk.LEFT, padx=(24, 12), pady=16)
    titulo_box = tk.Frame(header, bg=BG_ROOT)
    titulo_box.pack(side=tk.LEFT, pady=16)
    tk.Label(
        titulo_box, text="COCKPIT FECHAMENTO FITTED", bg=BG_ROOT, fg="#e9e9eb",
        font=("Segoe UI", 15, "bold"),
    ).pack(anchor="w")
    tk.Label(
        titulo_box, text="Fitted Units · Despesas", bg=BG_ROOT, fg="#9a9da2",
        font=("Consolas", 9, "bold"),
    ).pack(anchor="w", pady=(4, 0))

    # Status/spinner fica no cabecalho (canto direito, area escura fixa) em
    # vez de no corpo - o corpo tem conteudo que varia de altura (descricao
    # de cada aba) e o console de log e' expansivel, entao um indicador
    # colado la embaixo corre risco de ser espremido pra fora da janela
    # visivel. O cabecalho nunca encolhe, garante que sempre aparece.
    status_var = tk.StringVar(value="")
    tk.Label(
        header, textvariable=status_var, bg=BG_ROOT, fg=AMARELO_CLARO,
        font=("Consolas", 10, "bold"),
    ).pack(side=tk.RIGHT, padx=(0, 24), pady=16)

    tk.Frame(root, bg=AMARELO_CLARO, height=3).pack(fill=tk.X, side=tk.TOP)

    # Barra de "progresso" (indeterminada, sem %) - sempre visivel logo abaixo
    # do trim, mesmo lugar/altura o tempo todo (nunca pack/pack_forget, pra
    # nunca correr risco de ficar escondida). Em vez do bloco padrao do
    # ttk.Progressbar, desenha o pneuzinho Pirelli girando, deslizando de um
    # lado a outro - pedido explicito da usuaria.
    ALTURA_BARRA_PROGRESSO = 32
    canvas_progresso = tk.Canvas(root, height=ALTURA_BARRA_PROGRESSO, bg=BG_CAMPO, highlightthickness=0)
    canvas_progresso.pack(fill=tk.X, side=tk.TOP)

    _frames_pneu = _gerar_frames_pneu(diametro=ALTURA_BARRA_PROGRESSO - 2)
    _pneu = {"ativo": False, "x": 4.0, "direcao": 1, "indice_frame": 0}

    def _animar_pneu():
        canvas_progresso.delete("pneu")
        if _pneu["ativo"]:
            largura = canvas_progresso.winfo_width() or 400
            tam = _frames_pneu[0].width()
            limite = max(4, largura - tam - 4)
            _pneu["x"] += _pneu["direcao"] * 5
            if _pneu["x"] >= limite:
                _pneu["x"] = limite
                _pneu["direcao"] = -1
            elif _pneu["x"] <= 4:
                _pneu["x"] = 4
                _pneu["direcao"] = 1
            _pneu["indice_frame"] = (_pneu["indice_frame"] + 1) % len(_frames_pneu)
            canvas_progresso.create_image(
                _pneu["x"], ALTURA_BARRA_PROGRESSO // 2,
                anchor="w", image=_frames_pneu[_pneu["indice_frame"]], tags="pneu",
            )
        root.after(40, _animar_pneu)

    root.after(40, _animar_pneu)

    def iniciar_progresso():
        _pneu["ativo"] = True

    def parar_progresso():
        _pneu["ativo"] = False
        canvas_progresso.delete("pneu")

    # --- Área rolável (corpo inteiro) ---------------------------------
    # Pedido explícito da usuária, 2026-08-26: janela não cabia inteira em
    # telas/zoom menores (ficou pior depois do banner de aviso de Janeiro).
    # Canvas + Scrollbar em volta do "corpo" (painel + abas + log) - o
    # conteúdo de cada aba continua do mesmo jeito, só o CONTAINER de fora
    # passa a rolar se não couber tudo na altura visível da janela.
    scroll_area = tk.Frame(root, bg=BG_PAINEL)
    scroll_area.pack(fill=tk.BOTH, expand=True)

    corpo_canvas = tk.Canvas(scroll_area, bg=BG_PAINEL, highlightthickness=0)
    corpo_scrollbar = ttk.Scrollbar(scroll_area, orient="vertical", command=corpo_canvas.yview)
    corpo_canvas.configure(yscrollcommand=corpo_scrollbar.set)
    corpo_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    corpo_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    corpo = ttk.Frame(corpo_canvas, padding=(24, 18, 24, 18), style="TFrame")
    _corpo_janela = corpo_canvas.create_window((0, 0), window=corpo, anchor="nw")

    def _atualizar_scrollregion(event=None):
        corpo_canvas.configure(scrollregion=corpo_canvas.bbox("all"))

    corpo.bind("<Configure>", _atualizar_scrollregion)

    def _ajustar_largura_corpo(event):
        corpo_canvas.itemconfig(_corpo_janela, width=event.width)

    corpo_canvas.bind("<Configure>", _ajustar_largura_corpo)

    def _rolar_com_mouse(event):
        corpo_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _ligar_scroll_mouse(event):
        corpo_canvas.bind_all("<MouseWheel>", _rolar_com_mouse)

    def _desligar_scroll_mouse(event):
        corpo_canvas.unbind_all("<MouseWheel>")

    corpo_canvas.bind("<Enter>", _ligar_scroll_mouse)
    corpo_canvas.bind("<Leave>", _desligar_scroll_mouse)

    # --- Painel de instrumentos (Mes / Ano / Ciclo, compartilhado) ---
    hoje = datetime.now()
    if hoje.month == 1:
        mes_padrao, ano_padrao = 12, hoje.year - 1
    else:
        mes_padrao, ano_padrao = hoje.month - 1, hoje.year

    painel = ttk.Frame(corpo, style="Card.TFrame", padding=16)
    painel.pack(fill=tk.X, side=tk.TOP)

    ttk.Label(painel, text="ANO", style="Card.TLabel", font=("Consolas", 8, "bold")).grid(row=0, column=0, sticky="w")
    ano_var = tk.StringVar(value=str(ano_padrao))
    anos_disponiveis = [str(a) for a in range(hoje.year - 2, hoje.year + 2)]
    ttk.Combobox(
        painel, textvariable=ano_var, values=anos_disponiveis, width=8
    ).grid(row=1, column=0, sticky="w", padx=(0, 24), pady=(2, 0))

    ttk.Label(painel, text="MÊS", style="Card.TLabel", font=("Consolas", 8, "bold")).grid(row=0, column=1, sticky="w")
    mes_var = tk.StringVar(value=MESES_NOMES[mes_padrao])
    ttk.Combobox(
        painel, textvariable=mes_var, values=list(MESES_NOMES.values()), state="readonly", width=12
    ).grid(row=1, column=1, sticky="w", padx=(0, 24), pady=(2, 0))

    ttk.Label(painel, text="CICLO", style="Card.TLabel", font=("Consolas", 8, "bold")).grid(row=0, column=2, sticky="w")
    ciclo_var = tk.StringVar(value="Actual")
    ttk.Combobox(
        painel, textvariable=ciclo_var, values=["Actual", "Flash"], state="readonly", width=10
    ).grid(row=1, column=2, sticky="w", pady=(2, 0))

    def ler_mes_ano():
        nome_para_numero = {v: k for k, v in MESES_NOMES.items()}
        mes = nome_para_numero[mes_var.get()]
        try:
            ano = int(ano_var.get())
        except ValueError:
            messagebox.showerror("Ano inválido", "Digite um ano válido, ex: 2026.")
            return None
        return mes, ano

    # --- Abas, uma por passo do processo (ordem fixa) -----------------
    notebook = ttk.Notebook(corpo)
    notebook.pack(fill=tk.BOTH, expand=True, pady=(16, 0))

    botoes = {}

    def _rateio_precisa_confirmacao_janeiro():
        """True se hoje for Janeiro E ninguém tiver salvo uma entrada de
        rateio específica pra esse ano (vigente_desde == 'AAAA-01') ainda -
        pedido explícito da usuária (2026-08-25, implementado 2026-08-26):
        o rateio geralmente muda em Janeiro, e ela quer ser avisada na tela
        do cockpit se esquecer de confirmar/atualizar. Não bloqueia nada -
        se ninguém confirmar, `carregar_rateio_vigente` simplesmente segue
        usando a entrada anterior (comportamento já garantido, sem risco de
        travar o fechamento por causa do aviso)."""
        hoje = datetime.now()
        if hoje.month != 1:
            return False
        try:
            from gerar_rateio_custos import RATEIO_CONFIG_PATH
            import json as _json

            dados = _json.loads(RATEIO_CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            return False
        alvo = f"{hoje.year:04d}-01"
        return not any(e.get("vigente_desde") == alvo for e in dados.get("entradas", []))

    aviso_rateio_janeiro = _rateio_precisa_confirmacao_janeiro()

    def fazer_aba(passo, indice):
        aba = ttk.Frame(notebook, style="Card.TFrame", padding=24)
        notebook.add(aba, text=passo["aba"])

        ttk.Label(aba, text=passo["titulo"], style="Titulo.TLabel").pack(anchor="w")
        ttk.Label(aba, text=passo["descricao"], style="Descricao.TLabel", justify="left").pack(
            anchor="w", pady=(8, 20)
        )

        if indice == 4 and aviso_rateio_janeiro:
            tk.Label(
                aba,
                text=(
                    "⚠  É Janeiro — ninguém confirmou o % de rateio da Gerência pra este "
                    "ano ainda. Se não mudou, clique 'Atualizar Rateio' e salve com a mesma "
                    "vigência pra confirmar; se mudou, atualize os percentuais."
                ),
                bg="#FBEAEA", fg="#9C0006", font=("Segoe UI", 10, "bold"),
                justify="left", wraplength=620, padx=12, pady=8,
            ).pack(anchor="w", fill=tk.X, pady=(0, 16))

        widgets = []
        rotulos = passo["botoes"]
        for i, rotulo in enumerate(rotulos):
            btn = ttk.Button(aba, text=rotulo, style="Pirelli.TButton", cursor="hand2")
            btn.pack(fill=tk.X, ipady=8, pady=(0, 8) if i < len(rotulos) - 1 else 0)
            widgets.append(btn)
        botoes[indice] = widgets
        return aba

    for i, passo in enumerate(PASSOS):
        fazer_aba(passo, i)

    if aviso_rateio_janeiro:
        # Alem do banner na aba (que so' aparece se ela clicar la'), avisa na
        # hora que a janela abre - garante que ela ve mesmo sem entrar na
        # aba de Rateio de Custos.
        root.after(
            600,
            lambda: messagebox.showwarning(
                "Rateio da Gerência — confirmar pra este ano",
                "É Janeiro e ninguém confirmou (ou atualizou) o % de rateio da Gerência "
                "pra este ano ainda.\n\nVá na aba '⑤ Rateio de Custos' → 'Atualizar Rateio' "
                "pra confirmar (mesmo que os percentuais não mudem).",
            ),
        )

    # --- Console de log (compartilhado, sempre visivel embaixo) ------
    ttk.Label(corpo, text="LOG", font=("Consolas", 8, "bold"), style="TLabel").pack(
        anchor="w", pady=(16, 4)
    )
    log_widget = tk.Text(
        corpo, height=9, wrap="word", relief="solid", borderwidth=1,
        bg=LOG_BG, fg=LOG_FG, insertbackground=LOG_FG,
        highlightbackground=BORDA, highlightcolor=BORDA, highlightthickness=1,
        font=("Consolas", 9),
    )
    log_widget.pack(fill=tk.BOTH, expand=True)

    # Icone girando (spinner) junto do texto de status - roda em loop
    # independente (root.after) e so' mexe no texto quando "ativo" (setado
    # por rodar_em_thread), pra nao gastar ciclo a toa quando esta ocioso.
    FRAMES_SPINNER = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
    _spinner = {"indice": 0, "ativo": False, "descricao": ""}

    def _animar_spinner():
        if _spinner["ativo"]:
            frame = FRAMES_SPINNER[_spinner["indice"] % len(FRAMES_SPINNER)]
            _spinner["indice"] += 1
            status_var.set(f"{frame}  Processando: {_spinner['descricao']}...")
        root.after(120, _animar_spinner)

    root.after(120, _animar_spinner)

    # --- Log thread-safe -----------------------------------------------
    # As operacoes rodam numa thread separada (ver rodar_em_thread) pra
    # janela nao travar - so' a thread principal do Tk pode mexer em widget,
    # entao 'log' so' enfileira e quem escreve de verdade e' _drenar_fila,
    # chamada em loop via root.after (sempre na thread principal).
    fila_log = queue.Queue()

    def log(msg):
        fila_log.put(msg)

    def _drenar_fila():
        try:
            while True:
                msg = fila_log.get_nowait()
                log_widget.insert(tk.END, msg + "\n")
                log_widget.see(tk.END)
        except queue.Empty:
            pass
        root.after(80, _drenar_fila)

    root.after(80, _drenar_fila)

    # Botoes que ficam desabilitados por padrao (funcionalidade ainda nao
    # automatizada) e NUNCA devem ser reativados pelo "liberar janela" no
    # fim de uma operacao - "Atualizar Faturamento" (Passo 6), pedido
    # explicito da usuaria, 2026-08-26 (Net Sales continua manual).
    botoes_sempre_desabilitados = set()

    def _todos_botoes(estado):
        for lista in botoes.values():
            for btn in lista:
                if estado == "normal" and btn in botoes_sempre_desabilitados:
                    btn.config(state="disabled")
                else:
                    btn.config(state=estado)

    def _liberar_janela():
        parar_progresso()
        _spinner["ativo"] = False
        status_var.set("")
        root.config(cursor="")
        _todos_botoes("normal")

    def _perguntar_sim_nao(titulo, mensagem):
        """Dialogo Sim/Nao em portugues. O messagebox.askyesno padrao do Tk
        usa botoes fixos em ingles ('Yes'/'No'), mesmo com o resto do texto
        em portugues - quebra a regra do projeto de manter tudo em
        portugues (REGRAS_RAPIDAS #11, pedido explicito da usuaria). Modal
        (grab_set + wait_window), devolve True (Sim) ou False (Nao/fechar)."""
        dialogo = tk.Toplevel(root)
        dialogo.title(titulo)
        dialogo.configure(bg=BG_CARD)
        dialogo.resizable(False, False)
        dialogo.transient(root)
        dialogo.grab_set()

        resposta = {"valor": False}

        corpo = tk.Frame(dialogo, bg=BG_CARD, padx=24, pady=20)
        corpo.pack(fill=tk.BOTH, expand=True)
        tk.Label(
            corpo, text=f"⚠  {mensagem}", bg=BG_CARD, fg=TEXTO_CLARO,
            font=("Segoe UI", 10), justify="left", wraplength=420,
        ).pack(anchor="w")

        botoes_frame = tk.Frame(corpo, bg=BG_CARD)
        botoes_frame.pack(fill=tk.X, pady=(20, 0))

        def responder(valor):
            resposta["valor"] = valor
            dialogo.destroy()

        ttk.Button(
            botoes_frame, text="Não", cursor="hand2", command=lambda: responder(False)
        ).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(
            botoes_frame, text="Sim", style="Pirelli.TButton", cursor="hand2",
            command=lambda: responder(True),
        ).pack(side=tk.RIGHT)

        dialogo.protocol("WM_DELETE_WINDOW", lambda: responder(False))
        dialogo.update_idletasks()
        x = root.winfo_rootx() + (root.winfo_width() - dialogo.winfo_width()) // 2
        y = root.winfo_rooty() + (root.winfo_height() - dialogo.winfo_height()) // 2
        dialogo.geometry(f"+{max(0, x)}+{max(0, y)}")

        dialogo.wait_window()
        return resposta["valor"]

    def _avisar_travamento(descricao, decorrido_s, caixa_resultado, estado, permite_forcar_excel):
        """Mostra o aviso de possivel travamento (watchdog). Se a operacao
        usa Excel isolado (DispatchEx) E ja capturamos o PID dessa instancia
        (via pid_callback - ver abrir_excel_isolado em ksb1_core.py), oferece
        forcar o encerramento so' desse processo especifico. Nunca oferece
        encerrar o SAP GUI automaticamente: mataria TODAS as sessoes abertas
        dele, nao so' a desta automacao - decisao confirmada com a usuaria em
        2026-08-24."""
        minutos = int(decorrido_s // 60)
        pid = caixa_resultado.get("excel_pid") if permite_forcar_excel else None

        if pid:
            forcar = _perguntar_sim_nao(
                "Pode estar travado",
                f"'{descricao}' está rodando há mais de {minutos} minuto(s) sem terminar.\n\n"
                "Pode ser normal (bases grandes demoram) ou um travamento real do Excel.\n\n"
                "SIM = forçar o encerramento da instância isolada do Excel usada por esta "
                "operação (processo próprio, não afeta outros Excel que você tenha aberto) "
                "e cancelar a operação.\n"
                "NÃO = continuar aguardando.",
            )
            if not forcar:
                return
            log(f"\nForçando o encerramento do Excel desta operação (PID {pid})...")
            try:
                proc = psutil.Process(pid)
                if proc.name().upper() != "EXCEL.EXE":
                    log(
                        f"AVISO: o processo {pid} não é mais o EXCEL.EXE esperado "
                        "(já deve ter terminado sozinho) — nada foi encerrado."
                    )
                else:
                    proc.terminate()
                    log(f"Processo Excel (PID {pid}) encerrado à força.")
            except Exception as e:
                log(f"Não consegui encerrar o processo Excel (PID {pid}): {e}")
            estado["abandonado"] = True
            _liberar_janela()
            messagebox.showinfo(
                "Operação cancelada",
                f"'{descricao}' foi cancelada. Confira o log e, se precisar, rode a operação de novo.",
            )
            return

        if permite_forcar_excel:
            motivo = (
                "esta operação ainda não abriu o Excel (pode estar lendo um arquivo grande ou "
                "aguardando o SAP) — ainda não há um processo específico pra encerrar."
            )
        else:
            motivo = (
                "esta etapa usa o SAP GUI, não o Excel — encerrar o processo do SAP fecharia "
                "TODAS as suas sessões abertas, não só esta automação, então não faço isso "
                "automaticamente."
            )
        messagebox.showwarning(
            "Pode estar travado",
            f"'{descricao}' está rodando há mais de {minutos} minuto(s) sem terminar.\n\n"
            f"Pode ser normal ou um travamento real — {motivo}\n\n"
            "Se tiver certeza que travou, você pode encerrar manualmente pelo Gerenciador de "
            "Tarefas e tentar de novo. A janela continua aberta normalmente enquanto isso.",
        )

    def rodar_em_thread(descricao, func, ao_concluir, permite_forcar_excel=True):
        """Roda func(log, pid_callback) numa thread separada (com
        CoInitialize/CoUninitialize pro COM do SAP/Excel funcionar isolado
        por thread), mantendo a janela responsiva. func deve devolver o
        resultado (ou levantar excecao); pid_callback(pid) e' como func avisa
        o PID da instancia isolada do Excel que abriu (via abrir_excel_isolado
        em ksb1_core.py), se abrir uma - usado pelo watchdog abaixo pra saber
        o que encerrar se travar. ao_concluir(resultado, erro) roda de volta
        na thread principal, via root.after - nunca mexe em widget Tk fora da
        thread principal (trava ou corrompe a interface).

        Watchdog: se a operacao passar de TIMEOUT_AVISO_SEGUNDOS sem
        terminar, avisa (repetindo a cada novo intervalo enquanto continuar
        presa). Se a usuaria forcar o encerramento do Excel, a janela e'
        liberada na hora (estado['abandonado']=True) - a thread de fundo
        (daemon) pode continuar existindo ate' a chamada COM travada
        finalmente falhar (com o processo morto), mas isso acontece em
        segundo plano, sem prender mais a interface."""
        log_widget.delete("1.0", tk.END)
        log_widget.insert(tk.END, f"⏳ Processando: {descricao}...\n")
        log_widget.see(tk.END)
        _todos_botoes("disabled")
        root.config(cursor="watch")
        _spinner["ativo"] = True
        _spinner["descricao"] = descricao
        iniciar_progresso()
        # Forca redesenhar AGORA (cursor, botoes desabilitados, barra e texto
        # do log) antes de iniciar a thread - sem isso, se a operacao for
        # rapida (ex: Excel ja "aquecido" de uma rodada anterior), a janela
        # podia pular direto pro "Concluido" sem o usuario ver o estado
        # "processando" nem por um instante.
        root.update_idletasks()

        caixa_resultado = {}
        estado = {"abandonado": False, "inicio": time.monotonic(), "aviso_intervalo": 0}

        def registrar_pid(pid):
            caixa_resultado["excel_pid"] = pid

        def alvo():
            pythoncom.CoInitialize()
            try:
                caixa_resultado["valor"] = func(log, registrar_pid)
            except Exception as e:
                caixa_resultado["erro"] = e
            finally:
                pythoncom.CoUninitialize()

        thread = threading.Thread(target=alvo, daemon=True)
        thread.start()

        def checar():
            if thread.is_alive():
                decorrido = time.monotonic() - estado["inicio"]
                intervalo_atual = int(decorrido // TIMEOUT_AVISO_SEGUNDOS)
                if intervalo_atual > estado["aviso_intervalo"]:
                    estado["aviso_intervalo"] = intervalo_atual
                    _avisar_travamento(descricao, decorrido, caixa_resultado, estado, permite_forcar_excel)
                if estado["abandonado"]:
                    return  # janela ja liberada - so' para de monitorar, thread continua em segundo plano
                root.after(150, checar)
                return
            if estado["abandonado"]:
                # Ja tinha liberado a janela quando a usuaria forcou o
                # encerramento - so' registra no log que a thread terminou,
                # sem reabrir popup de conclusao/erro (ja mostrado antes).
                erro = caixa_resultado.get("erro")
                log(f"(Operação cancelada anteriormente terminou agora. {'Erro: ' + str(erro) if erro else 'Terminou sem erro.'})")
                return
            _liberar_janela()
            ao_concluir(caixa_resultado.get("valor"), caixa_resultado.get("erro"))

        root.after(150, checar)

    def ao_clicar_extrair():
        mes_ano = ler_mes_ano()
        if mes_ano is None:
            return
        mes, ano = mes_ano
        ciclo = ciclo_var.get()

        def func(log, pid_callback):
            rodar(mes, ano, ciclo, log)

        def ao_concluir(resultado, erro):
            if erro is not None:
                if isinstance(erro, ErroComTitulo):
                    messagebox.showerror(erro.titulo, erro.mensagem)
                else:
                    messagebox.showerror("Erro durante a extração", str(erro))
                return
            messagebox.showinfo("Concluído", "Extração da KSB1 finalizada (Gestoriais + Sem Agrupamento).")

        # Passo 1 usa so' o SAP GUI, nunca abre Excel - watchdog nao oferece
        # forcar encerramento (mataria todas as sessoes do SAP, nao so' esta).
        rodar_em_thread("Extraindo KSB1 do SAP", func, ao_concluir, permite_forcar_excel=False)

    def ao_clicar_check():
        from check_agrupamentos_ksb1 import gerar_check

        mes_ano = ler_mes_ano()
        if mes_ano is None:
            return
        mes, ano = mes_ano
        ciclo = ciclo_var.get()

        def func(log, pid_callback):
            return gerar_check(mes, ano, ciclo, log=log)

        def ao_concluir(resultado, erro):
            if erro is not None:
                messagebox.showerror("Erro ao gerar o check", str(erro))
                return
            messagebox.showinfo("Concluído", f"Check de agrupamentos gerado:\n{resultado}")

        # Passo 2 so' le arquivos ja extraidos via openpyxl, nao abre Excel/COM.
        rodar_em_thread("Gerando Check de Agrupamentos", func, ao_concluir, permite_forcar_excel=False)

    def ao_clicar_lancar_provisoes():
        from gerar_base_intermediaria import lancar_provisoes

        mes_ano = ler_mes_ano()
        if mes_ano is None:
            return
        mes, ano = mes_ano

        # Provisões e' sempre Flash, independente do Ciclo selecionado no
        # painel compartilhado (esse passo nem existe pro Actual).
        pasta_saida = resolver_pasta_ciclo(REDE_BASE / str(ano) / MESES_PASTA[mes], mes, "Flash")

        def func(log, pid_callback):
            return lancar_provisoes(mes, ano, pasta_saida, log=log, pid_callback=pid_callback)

        def ao_concluir(resultado, erro):
            if erro is not None:
                messagebox.showerror("Erro ao lançar as provisões", str(erro))
                return
            messagebox.showinfo("Concluído", f"Provisões lançadas:\n{resultado}")

        rodar_em_thread("Lançando Provisões", func, ao_concluir)

    def ao_clicar_atualizar_provisoes():
        from gerar_base_intermediaria import atualizar_provisoes

        mes_ano = ler_mes_ano()
        if mes_ano is None:
            return
        mes, ano = mes_ano

        pasta_saida = resolver_pasta_ciclo(REDE_BASE / str(ano) / MESES_PASTA[mes], mes, "Flash")

        def func(log, pid_callback):
            return atualizar_provisoes(mes, ano, pasta_saida, log=log, pid_callback=pid_callback)

        def ao_concluir(resultado, erro):
            if erro is not None:
                messagebox.showerror("Erro ao atualizar as provisões", str(erro))
                return
            messagebox.showinfo("Concluído", f"Provisões atualizadas:\n{resultado}")

        rodar_em_thread("Atualizando Provisões", func, ao_concluir)

    def ao_clicar_pivot():
        from gerar_ksb1_mensal import gerar_ksb1_mensal

        mes_ano = ler_mes_ano()
        if mes_ano is None:
            return
        mes, ano = mes_ano
        ciclo = ciclo_var.get()

        # Pasta de rede oficial do ciclo - tolera excecoes de nome (ex: pastas
        # de marco/abril usam o mes por extenso em vez da abreviacao padrao).
        pasta_saida = resolver_pasta_ciclo(REDE_BASE / str(ano) / MESES_PASTA[mes], mes, ciclo)

        def func(log, pid_callback):
            return gerar_ksb1_mensal(mes, ano, ciclo, pasta_saida, log=log, pid_callback=pid_callback)

        def ao_concluir(resultado, erro):
            if erro is not None:
                messagebox.showerror("Erro ao atualizar o Pivot KSB1", str(erro))
                return
            messagebox.showinfo("Concluído", f"Pivot KSB1 atualizado:\n{resultado}")

        rodar_em_thread("Atualizando Pivot KSB1", func, ao_concluir)

    def ao_clicar_finalizar_intermediaria():
        from gerar_base_intermediaria import atualizar_base_intermediaria

        mes_ano = ler_mes_ano()
        if mes_ano is None:
            return
        mes, ano = mes_ano
        ciclo = ciclo_var.get()

        pasta_saida = resolver_pasta_ciclo(REDE_BASE / str(ano) / MESES_PASTA[mes], mes, ciclo)

        def func(log, pid_callback):
            return atualizar_base_intermediaria(mes, ano, ciclo, pasta_saida, log=log, pid_callback=pid_callback)

        def ao_concluir(resultado, erro):
            if erro is not None:
                messagebox.showerror("Erro ao finalizar a Base Intermediária", str(erro))
                return
            caminho, caminho_historico, aviso_comparacao = resultado
            msg = f"Base Intermediária finalizada:\n{caminho}"
            if caminho_historico:
                msg += f"\n\nHistórico de unidades encerradas (enviar pra contabilidade):\n{caminho_historico}"
            messagebox.showinfo("Concluído", msg)
            if aviso_comparacao:
                messagebox.showwarning("Quadro de comparação", aviso_comparacao)

        rodar_em_thread("Finalizando a Base Intermediária", func, ao_concluir)

    def ao_clicar_rateio_custos():
        from gerar_rateio_custos import gerar_arquivo_rateio_custos

        mes_ano = ler_mes_ano()
        if mes_ano is None:
            return
        mes, ano = mes_ano
        ciclo = ciclo_var.get()

        pasta_saida = resolver_pasta_ciclo(REDE_BASE / str(ano) / MESES_PASTA[mes], mes, ciclo)

        def func(log, pid_callback):
            return gerar_arquivo_rateio_custos(mes, ano, ciclo, pasta_saida, log=log)

        def ao_concluir(resultado, erro):
            if erro is not None:
                messagebox.showerror("Erro ao gerar o Rateio de Custos", str(erro))
                return
            messagebox.showinfo("Concluído", f"Rateio de Custos gerado:\n{resultado}")

        # So le a Base Intermediaria via openpyxl (sem Excel/COM) - watchdog
        # nao precisa oferecer forcar Excel.
        rodar_em_thread("Gerando Abertura de Custos por Unidade", func, ao_concluir, permite_forcar_excel=False)

    def ao_clicar_atualizar_rateio():
        """Abre um dialogo pra editar o % de rateio da Gerencia por unidade
        (SJP/IBI/GOI/RES), pre-preenchido com o rateio vigente hoje. Salva
        como uma entrada nova (ou atualiza uma existente com a mesma
        'vigente_desde') em ontology/rateio_gerencia.json - pedido explicito
        da usuaria, 2026-08-26. Se ela nao mexer em nada num Ciclo, o rateio
        anterior continua valendo sozinho (carregar_rateio_vigente ja pega
        sempre a entrada mais recente <= mes/ano pedido - nao precisa de
        logica nova pra isso)."""
        import json as _json

        from gerar_rateio_custos import ORDEM_UNIDADES_ATIVAS, RATEIO_CONFIG_PATH, carregar_rateio_vigente

        hoje = datetime.now()
        try:
            percentuais_atuais, vigente_desde_atual = carregar_rateio_vigente(hoje.month, hoje.year)
        except RuntimeError as e:
            messagebox.showerror("Rateio não configurado", str(e))
            return

        dialogo = tk.Toplevel(root)
        dialogo.title("Atualizar Rateio da Gerência")
        dialogo.configure(bg=BG_CARD)
        dialogo.resizable(False, False)
        dialogo.transient(root)
        dialogo.grab_set()

        corpo_dlg = tk.Frame(dialogo, bg=BG_CARD, padx=24, pady=20)
        corpo_dlg.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            corpo_dlg, text=f"Rateio vigente hoje (desde {vigente_desde_atual}):",
            bg=BG_CARD, fg=TEXTO_CLARO, font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))

        entradas = {}
        for i, sigla in enumerate(ORDEM_UNIDADES_ATIVAS):
            tk.Label(corpo_dlg, text=sigla, bg=BG_CARD, fg=TEXTO_CLARO, font=("Segoe UI", 10)).grid(
                row=1 + i, column=0, sticky="w", pady=4
            )
            var = tk.StringVar(value=f"{percentuais_atuais.get(sigla, 0) * 100:.1f}".replace(".", ","))
            ttk.Entry(corpo_dlg, textvariable=var, width=10).grid(
                row=1 + i, column=1, sticky="w", padx=(12, 4), pady=4
            )
            tk.Label(corpo_dlg, text="%", bg=BG_CARD, fg=TEXTO_CLARO).grid(row=1 + i, column=2, sticky="w")
            entradas[sigla] = var

        linha_vig = 1 + len(ORDEM_UNIDADES_ATIVAS)
        tk.Label(
            corpo_dlg, text="Vigente a partir de (AAAA-MM):", bg=BG_CARD, fg=TEXTO_CLARO,
            font=("Segoe UI", 10),
        ).grid(row=linha_vig, column=0, columnspan=3, sticky="w", pady=(16, 4))
        vigencia_var = tk.StringVar(value=f"{hoje.year:04d}-{hoje.month:02d}")
        ttk.Entry(corpo_dlg, textvariable=vigencia_var, width=10).grid(
            row=linha_vig + 1, column=0, sticky="w"
        )

        aviso_var = tk.StringVar(value="")
        tk.Label(
            corpo_dlg, textvariable=aviso_var, bg=BG_CARD, fg="#9C0006",
            font=("Segoe UI", 9), wraplength=320, justify="left",
        ).grid(row=linha_vig + 2, column=0, columnspan=3, sticky="w", pady=(8, 0))

        botoes_frame = tk.Frame(corpo_dlg, bg=BG_CARD)
        botoes_frame.grid(row=linha_vig + 3, column=0, columnspan=3, sticky="e", pady=(16, 0))

        def salvar():
            vigencia = vigencia_var.get().strip()
            if len(vigencia) != 7 or vigencia[4] != "-":
                aviso_var.set("Formato inválido. Use AAAA-MM, ex: 2026-08.")
                return

            try:
                novos_percentuais = {
                    sigla: round(float(var.get().replace(",", ".")) / 100, 4)
                    for sigla, var in entradas.items()
                }
            except ValueError:
                aviso_var.set("Digite só números nos campos de %, ex: 21 ou 21,5.")
                return

            soma = sum(novos_percentuais.values())
            if abs(soma - 1.0) > 0.001:
                prosseguir = _perguntar_sim_nao(
                    "Soma diferente de 100%",
                    f"Os percentuais somam {soma * 100:.1f}%, não 100%.\n\nQuer salvar assim mesmo?",
                )
                if not prosseguir:
                    return

            dados = _json.loads(RATEIO_CONFIG_PATH.read_text(encoding="utf-8"))
            existente = next((e for e in dados["entradas"] if e["vigente_desde"] == vigencia), None)
            carimbo = f"Atualizado pelo cockpit em {datetime.now():%Y-%m-%d %H:%M}."
            if existente is not None:
                existente["percentuais"] = novos_percentuais
                existente["observacao"] = carimbo
            else:
                dados["entradas"].append(
                    {"vigente_desde": vigencia, "percentuais": novos_percentuais, "observacao": carimbo}
                )
            RATEIO_CONFIG_PATH.write_text(_json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
            dialogo.destroy()
            resumo = "\n".join(f"{s}: {p * 100:.1f}%" for s, p in novos_percentuais.items())
            messagebox.showinfo("Rateio atualizado", f"Rateio salvo, vigente a partir de {vigencia}:\n{resumo}")

        ttk.Button(botoes_frame, text="Cancelar", cursor="hand2", command=dialogo.destroy).pack(
            side=tk.RIGHT, padx=(8, 0)
        )
        ttk.Button(
            botoes_frame, text="Salvar", style="Pirelli.TButton", cursor="hand2", command=salvar
        ).pack(side=tk.RIGHT)

        dialogo.update_idletasks()
        x = root.winfo_rootx() + (root.winfo_width() - dialogo.winfo_width()) // 2
        y = root.winfo_rooty() + (root.winfo_height() - dialogo.winfo_height()) // 2
        dialogo.geometry(f"+{max(0, x)}+{max(0, y)}")

    def ao_clicar_atualizar_custo_mensalizacao():
        from gerar_mensalizacao import MESES_INGLES, REDE_BASE_MENSALIZACAO, gerar_arquivo_mensalizacao

        mes_ano = ler_mes_ano()
        if mes_ano is None:
            return
        mes, ano = mes_ano
        ciclo = ciclo_var.get()

        pasta_saida = REDE_BASE_MENSALIZACAO / ciclo / str(ano) / f"{mes:02d} - {MESES_INGLES[mes]}"

        def func(log, pid_callback):
            return gerar_arquivo_mensalizacao(mes, ano, pasta_saida, ciclo, log=log, pid_callback=pid_callback)

        def ao_concluir(resultado, erro):
            if erro is not None:
                messagebox.showerror("Erro ao atualizar o Custo (Mensalização)", str(erro))
                return
            caminho, checks = resultado
            messagebox.showinfo("Concluído", f"Mensalização (Custo) atualizada:\n{caminho}")
            # Pedido explicito da usuaria, 2026-08-26: ate' o Faturamento ser
            # automatizado, avisar pra atualizar essa parte na mao assim que
            # o Custo terminar.
            messagebox.showwarning(
                "Atualize o Faturamento manualmente",
                "O Custo do Passo 6 foi atualizado. A parte de Faturamento (Net Sales) "
                "ainda não foi automatizada — não esqueça de atualizar essa parte "
                "manualmente no arquivo antes de considerar a Mensalização completa.",
            )

        rodar_em_thread("Atualizando Custo (Mensalização)", func, ao_concluir)

    botoes[0][0].config(command=ao_clicar_extrair)
    botoes[1][0].config(command=ao_clicar_check)
    botoes[2][0].config(command=ao_clicar_lancar_provisoes)
    botoes[2][1].config(command=ao_clicar_atualizar_provisoes)
    botoes[3][0].config(command=ao_clicar_pivot)
    botoes[3][1].config(command=ao_clicar_finalizar_intermediaria)
    botoes[4][0].config(command=ao_clicar_rateio_custos)
    botoes[4][1].config(command=ao_clicar_atualizar_rateio)
    botoes[5][1].config(command=ao_clicar_atualizar_custo_mensalizacao)

    # "Atualizar Faturamento" (Passo 6) ainda nao foi automatizado - fica
    # desabilitado, e o helper _todos_botoes ja sabe manter ele assim mesmo
    # depois de outras operacoes terminarem (ver botoes_sempre_desabilitados).
    botoes_sempre_desabilitados.add(botoes[5][0])
    botoes[5][0].config(state="disabled")

    root.mainloop()


if __name__ == "__main__":
    main()
