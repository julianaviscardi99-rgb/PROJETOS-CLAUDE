#!/usr/bin/env python3
"""
Cockpit do Circuito Panamericano (CP) - controladoria / extracao SAP.

Mesmo formato do cockpit da Fitted Units (painel Ano/Mes/Ciclo, abas por passo, barra de
progresso, log), mas aplicativo proprio do CP: nao importa nada da Fitted. Passo 1 = extracao
da KSB1 (area 2281, todas as despesas). Os demais passos ficam em branco ate a Juliana
definir cada um. Cabecalho com a foto das pistas (pedido dela).

Pre-requisitos: Python 3 com pywin32, psutil e Pillow (pip install -r requirements.txt);
SAP GUI aberto e logado, com Scripting habilitado.
"""
import ctypes
import math
import os
import queue
import sys
import threading
import time
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, ttk

import psutil
import pythoncom
from PIL import Image, ImageDraw, ImageEnhance, ImageTk

PASTA = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, PASTA)
import ksb1_cp  # noqa: E402  (Passo 1: extracao KSB1)
from ksb1_cp import MESES_NOMES, ErroComTitulo  # noqa: E402

FOTO_CABECALHO = os.path.join(PASTA, "assets", "pista_circuito.jpg")
ICONE = os.path.join(PASTA, "assets", "pirelli_tire.ico")
ALTURA_CABECALHO = 170
FOCO_VERTICAL = 0.52  # faixa da foto que aparece no cabecalho (0 = topo, 1 = base)

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

AMARELO_CLARO = "#FFE9A8"
CINZA_TEXTO = "#555555"
GRAFITE = "#3a3b40"

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


class _Tooltip:
    """Caixa de texto que aparece ao passar o mouse sobre um widget (ex:
    botao) e some ao tirar o mouse - usado pra explicar o que cada botao faz
    sem deixar o texto sempre visivel ocupando espaco da tela (a usuaria
    achou o texto sempre visivel "um horror" em 2026-09-01, depois de pedir
    os passos lado a lado - preferiu botoes logo no topo da aba, com o
    detalhe so' aparecendo sob demanda)."""

    def __init__(self, widget, texto, wraplength=420):
        self.widget = widget
        self.texto = texto
        self.wraplength = wraplength
        self.janela = None
        widget.bind("<Enter>", self._mostrar)
        widget.bind("<Leave>", self._esconder)
        widget.bind("<Button-1>", self._esconder)

    def _mostrar(self, event=None):
        if self.janela or not self.texto:
            return
        self.janela = tk.Toplevel(self.widget)
        self.janela.wm_overrideredirect(True)
        self.janela.wm_geometry(f"+{event.x_root + 16}+{event.y_root + 12}")
        tk.Label(
            self.janela, text=self.texto, justify="left", wraplength=self.wraplength,
            bg="#FFFBE6", fg="#111111", font=("Segoe UI", 9),
            relief="solid", borderwidth=1, padx=10, pady=8,
        ).pack()

    def _esconder(self, event=None):
        if self.janela:
            self.janela.destroy()
            self.janela = None


def _foto_cabecalho(largura):
    """Foto cortada na faixa FOCO_VERTICAL pra preencher a largura, com degrade
    escuro a esquerda pra o titulo ficar legivel."""
    img = Image.open(FOTO_CABECALHO).convert("RGB")
    escala = largura / img.width
    img = img.resize((largura, max(ALTURA_CABECALHO, int(img.height * escala))), Image.LANCZOS)
    topo = int((img.height - ALTURA_CABECALHO) * FOCO_VERTICAL)
    img = img.crop((0, topo, largura, topo + ALTURA_CABECALHO))
    img = ImageEnhance.Brightness(img).enhance(0.9)
    degrade = Image.new("L", (largura, ALTURA_CABECALHO), 0)
    px = degrade.load()
    corte = int(largura * 0.55)
    for x in range(largura):
        alfa = int(190 * max(0.0, 1 - x / corte)) if x < corte else 0
        for y in range(ALTURA_CABECALHO):
            px[x, y] = alfa
    return Image.composite(Image.new("RGB", img.size, (11, 12, 14)), img, degrade)


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
    # Mesmo artefato do tema "clam" corrigido no Notebook.Tab acima: o layout
    # padrao inclui um sub-elemento de foco (contorno/traco mais claro) que
    # so' aparece no botao com foco de teclado (o primeiro botao carregado,
    # por padrao) - a usuaria viu como uma "linha" estranha no botao de
    # Extrair KSB1 (2026-08-28). Layout customizado remove esse sub-elemento.
    style.layout(
        "Pirelli.TButton",
        [
            (
                "Button.border",
                {
                    "sticky": "nswe",
                    "border": "1",
                    "children": [
                        (
                            "Button.padding",
                            {"sticky": "nswe", "children": [("Button.label", {"sticky": "nswe"})]},
                        )
                    ],
                },
            )
        ],
    )


    # Botao cinza clarinho pra acoes secundarias/alternativas (ex: "Atualizar
    # Provisoes", que so' se usa depois de uma correcao - "Lancar Provisoes"
    # e' a acao principal do par) - pedido da usuaria, 2026-09-01. Reusa o
    # mesmo layout do Pirelli.TButton (sem o artefato de foco do tema clam).
    style.configure("Secundario.TButton", font=("Segoe UI", 11, "bold"), foreground="black", borderwidth=0)
    style.map(
        "Secundario.TButton",
        background=[("!disabled", "#E7E7EA"), ("disabled", "#f1f1f2")],
        foreground=[("!disabled", "black"), ("disabled", "#8a8a8a")],
    )
    style.layout("Secundario.TButton", style.layout("Pirelli.TButton"))

    # Combobox usa listas suspensas nativas do Tk (nao ttk) — precisam ser
    # coloridas separadamente, senao ficam brancas mesmo com o tema escuro.
    root.option_add("*TCombobox*Listbox.background", BG_CAMPO)
    root.option_add("*TCombobox*Listbox.foreground", TEXTO_CLARO)
    root.option_add("*TCombobox*Listbox.selectBackground", AMARELO_CLARO)
    root.option_add("*TCombobox*Listbox.selectForeground", "black")



PASSOS = [
    {
        "aba": "①  Extração",
        "titulo": "Passo 1 · Extrair KSB1",
        "botoes": ["Extrair KSB1 (empresa 2281)"],
        "tooltips": [
            "Baixa a KSB1 direto do SAP com todas as despesas da empresa 2281 (sem grupo de "
            "centros de custo e sem agrupamento), pro mês/ano escolhidos, e salva na pasta do "
            "ciclo (02. Flash / 03. Actual) como 'KSB1 - Circuito Panamericano MM.AAAA - Ciclo' (se já existir, salva "
            "como _v2, _v3...). Pré-requisito: SAP GUI aberto e logado."
        ],
    },
    {"aba": "②  (em branco)", "titulo": "Passo 2 · a definir", "botoes": []},
    {"aba": "③  (em branco)", "titulo": "Passo 3 · a definir", "botoes": []},
    {"aba": "④  (em branco)", "titulo": "Passo 4 · a definir", "botoes": []},
    {"aba": "⑤  (em branco)", "titulo": "Passo 5 · a definir", "botoes": []},
]

def main():
    root = tk.Tk()
    root.title("Circuito Panamericano · Cockpit Fechamento")
    root.geometry("1317x800")
    root.minsize(1000, 650)
    root.configure(bg=BG_ROOT)
    root.state("zoomed")
    try:
        root.iconbitmap(ICONE)
    except tk.TclError:
        pass

    _configurar_estilo(root)

    # --- Cabecalho com a foto das pistas (pedido da Juliana) -----------------
    cab = tk.Canvas(root, height=ALTURA_CABECALHO, bg=BG_ROOT, highlightthickness=0)
    cab.pack(fill=tk.X, side=tk.TOP)
    estado_cab = {"tk": None, "largura": 0}
    status_var = tk.StringVar(value="")

    def desenhar_cabecalho(_evento=None):
        largura = cab.winfo_width()
        if largura < 50 or largura == estado_cab["largura"]:
            return
        estado_cab["largura"] = largura
        cab.delete("all")
        try:
            estado_cab["tk"] = ImageTk.PhotoImage(_foto_cabecalho(largura))
            cab.create_image(0, 0, image=estado_cab["tk"], anchor="nw")
        except OSError:
            pass  # sem foto: fica o fundo escuro
        cab.create_text(32, 62, text="COCKPIT FECHAMENTO CIRCUITO", anchor="w",
                        fill="#ffffff", font=("Segoe UI", 22, "bold"))
        cab.create_text(34, 102, text="Circuito Panamericano · Controladoria", anchor="w",
                        fill=AMARELO_CLARO, font=("Consolas", 11, "bold"))
        cab.create_window(largura - 24, ALTURA_CABECALHO - 26, anchor="se", window=tk.Label(
            cab, textvariable=status_var, bg=BG_ROOT, fg=AMARELO_CLARO, font=("Consolas", 10, "bold")))

    cab.bind("<Configure>", desenhar_cabecalho)

    tk.Frame(root, bg=AMARELO_CLARO, height=8).pack(fill=tk.X, side=tk.TOP)

    # Barra de "progresso" (indeterminada, sem %) - sempre visivel logo abaixo
    # do trim, mesmo lugar/altura o tempo todo (nunca pack/pack_forget, pra
    # nunca correr risco de ficar escondida). Em vez do bloco padrao do
    # ttk.Progressbar, desenha o pneuzinho Pirelli girando, deslizando de um
    # lado a outro - pedido explicito da usuaria. Faixa e pneu aumentados
    # (2026-08-28, pedido da usuaria) pra ficar mais visivel/parecer melhor.
    ALTURA_BARRA_PROGRESSO = 56
    canvas_progresso = tk.Canvas(root, height=ALTURA_BARRA_PROGRESSO, bg=BG_CAMPO, highlightthickness=0)
    canvas_progresso.pack(fill=tk.X, side=tk.TOP)

    _frames_pneu = _gerar_frames_pneu(diametro=ALTURA_BARRA_PROGRESSO - 2)
    _pneu = {"ativo": False, "x": 4.0, "direcao": 1, "indice_frame": 0}

    # Mesmo pneu (com calota) do indicador de "processando" acima, girando -
    # usado como icone antes do numero da aba, mas so' na aba SELECIONADA/
    # amarela (pedido da usuaria, 2026-08-28: aparece so' quando selecionado,
    # e fica girando continuamente enquanto essa aba estiver selecionada).
    _frames_pneu_aba = _gerar_frames_pneu(diametro=40, n_frames=12)
    _pneu_aba = {"indice_frame": 0, "aba_atual": 0}

    def _animar_pneu_aba():
        _pneu_aba["indice_frame"] = (_pneu_aba["indice_frame"] + 1) % len(_frames_pneu_aba)
        item = labels_aba.get(_pneu_aba["aba_atual"])
        if item is not None:
            item[1].config(image=_frames_pneu_aba[_pneu_aba["indice_frame"]])
        root.after(80, _animar_pneu_aba)

    root.after(80, _animar_pneu_aba)

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
    # Barra de abas propria (Frame + Label por aba) em vez de ttk.Notebook -
    # o tema "clam" (unico que permite recolorir aba a aba - ver
    # _configurar_estilo) desenha um corte diagonal no canto interno de cada
    # aba (efeito visual do proprio tema, nao um bug de layout), o que a
    # usuaria viu como "desalinhado" no Passo 1 mesmo depois de remover o
    # sub-elemento de foco (2026-08-28). Widgets Tk puros dao controle total
    # de cor/formato/alinhamento, sem esse artefato.
    abas_container = tk.Frame(
        corpo, bg=BG_PAINEL, highlightbackground=BORDA, highlightthickness=1, bd=0,
    )
    abas_container.pack(fill=tk.BOTH, expand=True, pady=(16, 0))

    barra_abas = tk.Frame(abas_container, bg=BG_PAINEL)
    barra_abas.pack(fill=tk.X, side=tk.TOP)

    paginas_container = tk.Frame(abas_container, bg=BG_PAINEL)
    paginas_container.pack(fill=tk.BOTH, expand=True)
    paginas_container.grid_rowconfigure(0, weight=1)
    paginas_container.grid_columnconfigure(0, weight=1)

    botoes = {}
    paginas = {}
    labels_aba = {}


    aviso_rateio_janeiro = False  # (so existe na Fitted)

    def selecionar_aba(indice):
        _pneu_aba["aba_atual"] = indice
        for i, (frame, icone_lbl, texto_lbl) in labels_aba.items():
            cor_bg = AMARELO_CLARO if i == indice else GRAFITE
            cor_fg = "black" if i == indice else "#e9e9eb"
            imagem = _frames_pneu_aba[_pneu_aba["indice_frame"]] if i == indice else ""
            icone_lbl.config(image=imagem, bg=cor_bg)
            frame.config(bg=cor_bg)
            texto_lbl.config(bg=cor_bg, fg=cor_fg)
        paginas[indice].tkraise()

    def fazer_aba(passo, indice):
        aba = ttk.Frame(paginas_container, style="Card.TFrame", padding=24)
        aba.grid(row=0, column=0, sticky="nsew")
        paginas[indice] = aba

        # Icone e texto em widgets separados (em vez de um so' Label com
        # compound="left") - o Tk classic reaproveita o "padx" do Label como
        # espaco TANTO na borda quanto entre imagem/texto, o que deixava o
        # pneu longe do numero mesmo com padx baixo (2026-08-28, achado da
        # usuaria). Com widgets separados da' pra controlar cada gap sozinho.
        tab_frame = tk.Frame(barra_abas, bg=GRAFITE, cursor="hand2")
        tab_frame.pack(side=tk.LEFT, padx=(0, 2))

        icone_lbl = tk.Label(tab_frame, bg=GRAFITE, borderwidth=0, cursor="hand2")
        icone_lbl.pack(side=tk.LEFT, padx=(14, 0), pady=10)

        texto_lbl = tk.Label(
            tab_frame, text=passo["aba"], bg=GRAFITE, fg="#e9e9eb",
            font=("Segoe UI", 10, "bold"), cursor="hand2",
        )
        texto_lbl.pack(side=tk.LEFT, padx=(4, 14), pady=10)

        for widget in (tab_frame, icone_lbl, texto_lbl):
            widget.bind("<Button-1>", lambda e, i=indice: selecionar_aba(i))
        labels_aba[indice] = (tab_frame, icone_lbl, texto_lbl)

        ttk.Label(aba, text=passo["titulo"], style="Titulo.TLabel").pack(anchor="w", pady=(0, 14))

        if indice == 3 and aviso_rateio_janeiro:
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
        tooltips = passo.get("tooltips", [])
        estilos = passo.get("estilos", [])
        for i, rotulo in enumerate(rotulos):
            estilo = estilos[i] if i < len(estilos) else "Pirelli.TButton"
            btn = ttk.Button(aba, text=rotulo, style=estilo, cursor="hand2")
            btn.pack(fill=tk.X, ipady=8, pady=(0, 8) if i < len(rotulos) - 1 else 0)
            widgets.append(btn)
            if i < len(tooltips) and tooltips[i]:
                _Tooltip(btn, tooltips[i])
        botoes[indice] = widgets
        return aba

    for i, passo in enumerate(PASSOS):
        fazer_aba(passo, i)
    selecionar_aba(0)

    # --- Console de log (compartilhado, sempre visivel embaixo) ------
    ttk.Label(corpo, text="LOG", font=("Consolas", 8, "bold"), style="TLabel").pack(
        anchor="w", pady=(16, 4)
    )
    log_widget = tk.Text(
        corpo, height=9, wrap="word", relief="flat", borderwidth=0,
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
    # fim de uma operacao - "Atualizar Faturamento" (Passo 5), pedido
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
            ksb1_cp.rodar(mes, ano, ciclo, log)

        def ao_concluir(resultado, erro):
            if erro is not None:
                if isinstance(erro, ErroComTitulo):
                    messagebox.showerror(erro.titulo, erro.mensagem)
                else:
                    messagebox.showerror("Erro durante a extração", str(erro))
                return
            messagebox.showinfo("Concluído", "Extração da KSB1 finalizada (empresa 2281).")

        # Passo 1 usa so' o SAP GUI, nunca abre Excel - watchdog nao oferece
        # forcar encerramento (mataria todas as sessoes do SAP, nao so' esta).
        rodar_em_thread("Extraindo KSB1 do SAP", func, ao_concluir, permite_forcar_excel=False)

    botoes[0][0].config(command=ao_clicar_extrair)

    root.mainloop()


if __name__ == "__main__":
    main()
