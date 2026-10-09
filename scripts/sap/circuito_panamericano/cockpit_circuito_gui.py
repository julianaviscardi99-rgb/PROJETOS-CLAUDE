"""Cockpit do Circuito Panamericano (CP) - controladoria / extracao SAP.

Aplicativo proprio do CP, separado do cockpit da Fitted Units. Por enquanto e'
so a casca (cabecalho com foto das pistas, area de log); as etapas entram aqui
conforme forem definidas.
"""
import os
import tkinter as tk
from tkinter import ttk

from PIL import Image, ImageEnhance, ImageTk

PASTA = os.path.dirname(os.path.abspath(__file__))
FOTO_CABECALHO = os.path.join(PASTA, "assets", "pista_circuito.jpg")
ICONE = os.path.join(PASTA, "assets", "pirelli_tire.ico")

AMARELO = "#FFE9A8"
BG_ROOT = "#0b0c0e"  # so aparece se a foto nao carregar
BG_PAINEL = "#ffffff"
BG_CAMPO = "#f0f0f2"
BORDA = "#d5d6d9"
TEXTO = "#111111"
TEXTO_SECUNDARIO = "#5a5c60"
ALTURA_CABECALHO = 170
FOCO_VERTICAL = 0.52  # faixa da foto que aparece no cabecalho (0 = topo, 1 = base)


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


def main():
    root = tk.Tk()
    root.title("Circuito Panamericano")
    root.geometry("1100x720")
    root.minsize(800, 520)
    root.configure(bg=BG_PAINEL)
    try:
        root.iconbitmap(ICONE)
    except tk.TclError:
        pass
    root.state("zoomed")

    cab = tk.Canvas(root, height=ALTURA_CABECALHO, bg=BG_ROOT, highlightthickness=0)
    cab.pack(fill=tk.X, side=tk.TOP)
    estado = {"tk": None, "largura": 0}

    def desenhar(_evento=None):
        largura = cab.winfo_width()
        if largura < 50 or largura == estado["largura"]:
            return
        estado["largura"] = largura
        cab.delete("all")
        try:
            estado["tk"] = ImageTk.PhotoImage(_foto_cabecalho(largura))
            cab.create_image(0, 0, image=estado["tk"], anchor="nw")
        except OSError:
            pass  # sem foto: fica o fundo escuro
        cab.create_text(32, 66, text="CIRCUITO PANAMERICANO", anchor="w",
                        fill="#ffffff", font=("Segoe UI", 24, "bold"))
        cab.create_text(34, 106, text="Controladoria  ·  Extração SAP", anchor="w",
                        fill=AMARELO, font=("Consolas", 11, "bold"))

    cab.bind("<Configure>", desenhar)
    tk.Frame(root, bg=AMARELO, height=8).pack(fill=tk.X, side=tk.TOP)

    corpo = tk.Frame(root, bg=BG_PAINEL)
    corpo.pack(fill=tk.BOTH, expand=True, padx=24, pady=20)
    ttk.Label(corpo, text="Nenhuma etapa configurada ainda.", background=BG_PAINEL,
              foreground=TEXTO, font=("Segoe UI", 12, "bold")).pack(anchor="w")
    ttk.Label(corpo, text="As etapas do CP (extração SAP, classificação de custos, "
                          "mensalização) serão adicionadas aqui.",
              background=BG_PAINEL, foreground=TEXTO_SECUNDARIO,
              font=("Segoe UI", 10)).pack(anchor="w", pady=(4, 16))

    ttk.Label(corpo, text="LOG", background=BG_PAINEL, foreground=TEXTO,
              font=("Consolas", 8, "bold")).pack(anchor="w")
    log = tk.Text(corpo, bg="#ffffff", fg=TEXTO, relief="flat", height=10,
                  highlightbackground=BORDA, highlightcolor=BORDA, highlightthickness=1,
                  font=("Consolas", 10))
    log.pack(fill=tk.BOTH, expand=True)
    log.insert("end", "Cockpit do Circuito Panamericano iniciado.\n")
    log.config(state="disabled")

    root.mainloop()


if __name__ == "__main__":
    main()
