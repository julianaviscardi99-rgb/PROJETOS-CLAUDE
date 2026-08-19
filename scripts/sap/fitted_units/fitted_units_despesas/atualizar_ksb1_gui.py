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
import sys
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk

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
    nome_com_versao,
    voltar_para_selecao,
)

# Logo da Pirelli embutido (base64) para o script continuar autossuficiente
# (nao depender de arquivo de imagem separado ao copiar para a rede).
LOGO_PIRELLI_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAJsAAAAvCAYAAAD0OrjvAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAU4SURBVHhe7Zvri1VlFIcnm5qpbLLLt7KLlpk5Y4ZI6nwIDE0s8MOEgiRFN0UKTJQUu2AYo0KW1HRRpMuoYVaa1hBdUCu7Rx+ke/0rb8+PMweOb2vfztTrmWF9eEBnr7X23u969n735ey28HdbcJwUuGxOMlw2Jxkum5MMl81JhsvmJMNlc5LhsjnJcNmcZFSX7c2zQrjvvBD6upyxzLILQ3ikI4Q99PynyIEmqS7bdmS7iQ25/DJnLDMRrrs4hDmIt7YzhC8NFyrisjn5XAHTkW4js9nPhg8VSCPbVZeGsJSNffQcJxVr4EGYx7hbPamChFs0PoRDhg8VSCPb5EtCGDg3hB/Id9LxDbzaHsKNE+y+VGEWPd9BPcuJkqSR7Vpkex3ZrHrO/8vRcYjCWcnqSxXU836jfgXSyHYN0+gzyDZEvpOGD+EtWMVUqssYqy9VmEvPX6Ge5URJ0simOf8Gzm4zuNB00tENmlWsnlSlj55/bvhQgTSyOaOb3q4QdtH3Pw0fKuCyOdlM4qy45PwQ9tHz3wwXKpLumm1DRwiD5DqjgwPwGZyCPwwPmiCNbHr08Ro3CDoNO6OHv8ByoEnSyOaPPhw4s7I9BYv4+/xOm6Fx/87ZVJBThsXtIWykzteQdfTeRcxtRm5ZFsL+qKamo/2MnxUvFnCpsSrKyeIkLGd8rDribsbo4yjH4ivYkFNHbyE+inKa5MzKtpKLT02xVo54hx2Nc+6nzmRu6a34qkyfEMJBBtq6JrmZZXpkY+WV4WquU3dFNX+Hl1mfFS+uJGfx+NNzsjhGnbmMn1VH9F4UwvtGXsxxuIcxtWqIhfT6cJTTJGNPNj3AVKPr6P950sykKSfZp/gMVySbxFDtLLTPu6OaLpu9IJNWl63/7BD2ErdvmGdhNgOvn8xY8aKfxv1KXON6imR7gJjnctgJJ6CxpstmL8ik1WU7gmy6k2rM2Q5Tc3LupXGnopwi2V4gRtNvHvHZ0mWzF2TS6rK9i2z63ZXOVEL/fprGXJ+znhVNyLaenLdz0LXg91FNl81ekEmry/YYOc/TiJ3DPAkzcsTR3zcT80u0Hr9BqOGy5chWFb2w/pR9iqdel62Gy/YfyKbGzqTOAE2z3vsVyaZfVNySQy/5g1FNl81ekEmryzaLZfNo9DSYSPOsGDGNuD0Zooki2bYR810B8dTsstkLMml12Qa5QdDPoR+HvDtQPQsbIDZ+5FGnSLatxHxbgIRrlLmMbAsYW6tWnfpndUWyzUG2NxryYn4EbZvLNkwzstUffehV1Hy2Uw204sRsziJDxvWaKJLtDrZtZQEPt5/+rK1INq2vh/VatcRq6m1j/1SrSLap1Flq1KjzBHWOUcdlG2YksilWU52mSyuuzoqO2pFe9Q1CGSYheuNUViRbEbrp6LugVqtItiJu7QrhIHVaWrYdyKadnEITy9LDKX0vR2VcS19cd7PMyhGHjJzVCNjDQFnx4oMG2fSF0Z2cvXSUW7FCU+2L7FN87dbLIE/JyStDN/lHG2pKtt2sy4otg65Dl3fWah2njl6UW3FluJ1xeY86+qn3Q4ypFSOWIPcRYur7MAKqy6YN3ERD13A0lGUd0nzC4MS19O5Qy6wcccLI0UcXeTnxe05dt6xnMK3YOluIiS/mNc3o20srvizryG/8klwHgT5GsWLLsJZ6ehWmWjqQNrOvVlwZtjCGmuJV5yXOklaM2ErcF8PrHCHVZXOcJnHZnGS4bE4yXDYnGS6bkwyXzUmGy+Ykw2VzkuGyOclw2ZxEtIV/AAlejwcdSLwvAAAAAElFTkSuQmCC"
)


def extrair_um(session, mes, ano, koagr, agrup_label, log):
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

    log(f"Executando KSB1 ({agrup_label})...")
    session.FindById("wnd[0]").SendVKey(8)

    pasta_rede = REDE_BASE / str(ano) / "00.Extração Base KSB1" / MESES_PASTA[mes]
    pasta_rede.mkdir(parents=True, exist_ok=True)
    nome_arquivo = nome_com_versao(
        pasta_rede, f"KSB1 - {BU['nome']} {mes:02d}.{ano} - {agrup_label}.XLSX"
    )

    session.FindById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select()
    wnd1 = session.FindById("wnd[1]")
    wnd1.FindById("usr/ctxtDY_PATH").Text = str(pasta_rede)
    wnd1.FindById("usr/ctxtDY_FILENAME").Text = nome_arquivo
    wnd1.FindById("tbar[0]/btn[0]").Press()  # Gerar

    # As notificacoes de seguranca de scripting ja ficam desativadas nas
    # opcoes do SAP GUI (Acessibilidade & scripting > Scripting), entao o
    # popup "Seguranca SAPGUI" nao aparece de fato. Em vez de travar a
    # extracao esperando confirmacao manual, so aguarda alguns instantes
    # para o SAP terminar de escrever o arquivo na rede.
    arquivo_final = pasta_rede / nome_arquivo
    for _ in range(20):
        if arquivo_final.exists():
            break
        time.sleep(0.5)

    if arquivo_final.exists():
        log(f"{agrup_label}: salvo em {arquivo_final}")
    else:
        log(f"AVISO: não encontrei o arquivo de {agrup_label} na pasta esperada. Confira manualmente.")

    voltar_para_selecao(session, log)


def rodar(mes, ano, log_widget):
    def log(msg):
        log_widget.insert(tk.END, msg + "\n")
        log_widget.see(tk.END)
        log_widget.update()

    log_widget.delete("1.0", tk.END)

    try:
        session = connect_session()
    except Exception as e:
        messagebox.showerror(
            "Erro de conexão",
            f"Não consegui conectar ao SAP GUI.\n\nDetalhe: {e}\n\n"
            "Verifique se o SAP GUI está aberto e logada, e se o Scripting está "
            "habilitado (Alt+F12 > Opções > Acessibilidade e Scripting > Scripting).",
        )
        return

    try:
        abrir_ksb1(session, log)
    except RuntimeError as e:
        messagebox.showerror(
            "Não consegui abrir a KSB1",
            f"{e}\n\nConfirme se seu usuário tem acesso à transação KSB1 e tente de novo.",
        )
        return

    log(f"Extraindo KSB1 - {MESES_NOMES[mes]}/{ano}...")
    try:
        extrair_um(session, mes, ano, "gestoriais", "Gestoriais", log)
        extrair_um(session, mes, ano, "", "Sem Agrupamento", log)
    except Exception as e:
        messagebox.showerror("Erro durante a extração", str(e))
        return

    log("\nConcluído!")
    messagebox.showinfo("Concluído", "Extração da KSB1 finalizada (Gestoriais + Sem Agrupamento).")


AMARELO_PIRELLI = "#FFD400"
VERMELHO_PIRELLI = "#DA291C"
CINZA_TEXTO = "#555555"

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
BG_RODAPE = "#ffffff"
TREAD_DARK = "#1f2126"
TREAD_LINE = "#c7c9cc"

PASSOS = [
    {
        "aba": "①  Extração",
        "titulo": "Passo 1 · Extrair KSB1",
        "descricao": (
            "Baixa a KSB1 direto do SAP (Gestoriais + Sem Agrupamento) pro mês/ano "
            "escolhido e salva os dois arquivos na área de rede. Pré-requisito: SAP "
            "GUI aberto e logado na tela inicial (o script abre a transação sozinho)."
        ),
        "botao": "Extrair KSB1 (Gestoriais + Sem Agrupamento)",
    },
    {
        "aba": "②  Check de Agrupamentos",
        "titulo": "Passo 2 · Check de Agrupamentos",
        "descricao": (
            "Confere se toda conta contábil do Sem Agrupamento está vinculada a um "
            "agrupamento gestorial, comparando com o arquivo Gestoriais do mesmo mês. "
            "Usa os arquivos já extraídos no Passo 1 — não acessa o SAP."
        ),
        "botao": "Gerar Check de Agrupamentos",
    },
    {
        "aba": "③  Base Intermediária",
        "titulo": "Passo 3 · Atualizar KSB1 Pivot",
        "descricao": (
            "Atualiza o KSB1 acumulado do ano (BASE_KSB1 + Pivot Tables nativas) com "
            "as linhas do mês e prepara os valores pra colar na Base Intermediária. "
            "Usa o Ciclo selecionado (Actual/Flash) só no nome do arquivo final."
        ),
        "botao": "Atualizar KSB1 Pivot",
    },
]


def desenhar_rastro_pneu(canvas, largura, altura):
    """Desenha uma faixa com padrao de sulco de pneu (blocos repetidos),
    puramente vetorial via Canvas — sem depender de imagem externa,
    mesmo espirito do logo embutido em base64."""
    y_centro = altura // 2
    canvas.create_line(0, y_centro, largura, y_centro, fill=TREAD_LINE, width=1)
    bloco_w, vao = 12, 7
    x = 4
    while x < largura - bloco_w:
        canvas.create_rectangle(
            x, 3, x + bloco_w, altura - 3,
            fill=TREAD_DARK, outline="",
        )
        x += bloco_w + vao


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
        background=[("!disabled", VERMELHO_PIRELLI), ("disabled", "#4a2320")],
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
        background=[("selected", VERMELHO_PIRELLI), ("!selected", BG_CARD)],
        foreground=[("selected", "black"), ("!selected", TEXTO_SECUNDARIO)],
        expand=[("selected", (0, 0, 0, 0))],
    )

    # Combobox usa listas suspensas nativas do Tk (nao ttk) — precisam ser
    # coloridas separadamente, senao ficam brancas mesmo com o tema escuro.
    root.option_add("*TCombobox*Listbox.background", BG_CAMPO)
    root.option_add("*TCombobox*Listbox.foreground", TEXTO_CLARO)
    root.option_add("*TCombobox*Listbox.selectBackground", VERMELHO_PIRELLI)
    root.option_add("*TCombobox*Listbox.selectForeground", "black")


def main():
    root = tk.Tk()
    root.title("Fitted Units · Cockpit Fechamento")
    root.geometry("860x640")
    root.minsize(780, 560)
    root.configure(bg=BG_ROOT)

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

    tk.Frame(root, bg=VERMELHO_PIRELLI, height=3).pack(fill=tk.X, side=tk.TOP)

    corpo = ttk.Frame(root, padding=(24, 18, 24, 18), style="TFrame")
    corpo.pack(fill=tk.BOTH, expand=True)

    # --- Painel de instrumentos (Mes / Ano / Ciclo, compartilhado) ---
    hoje = datetime.now()
    if hoje.month == 1:
        mes_padrao, ano_padrao = 12, hoje.year - 1
    else:
        mes_padrao, ano_padrao = hoje.month - 1, hoje.year

    painel = ttk.Frame(corpo, style="Card.TFrame", padding=16)
    painel.pack(fill=tk.X, side=tk.TOP)

    ttk.Label(painel, text="MÊS", style="Card.TLabel", font=("Consolas", 8, "bold")).grid(row=0, column=0, sticky="w")
    mes_var = tk.StringVar(value=MESES_NOMES[mes_padrao])
    ttk.Combobox(
        painel, textvariable=mes_var, values=list(MESES_NOMES.values()), state="readonly", width=12
    ).grid(row=1, column=0, sticky="w", padx=(0, 24), pady=(2, 0))

    ttk.Label(painel, text="ANO", style="Card.TLabel", font=("Consolas", 8, "bold")).grid(row=0, column=1, sticky="w")
    ano_var = tk.StringVar(value=str(ano_padrao))
    ttk.Entry(painel, textvariable=ano_var, width=8).grid(row=1, column=1, sticky="w", padx=(0, 24), pady=(2, 0))

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

    def fazer_aba(passo, indice):
        aba = ttk.Frame(notebook, style="Card.TFrame", padding=24)
        notebook.add(aba, text=passo["aba"])

        ttk.Label(aba, text=passo["titulo"], style="Titulo.TLabel").pack(anchor="w")
        ttk.Label(aba, text=passo["descricao"], style="Descricao.TLabel", justify="left").pack(
            anchor="w", pady=(8, 20)
        )

        btn = ttk.Button(aba, text=passo["botao"], style="Pirelli.TButton")
        btn.pack(fill=tk.X, ipady=8)
        botoes[indice] = btn
        return aba

    for i, passo in enumerate(PASSOS):
        fazer_aba(passo, i)

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

    def log(msg):
        log_widget.insert(tk.END, msg + "\n")
        log_widget.see(tk.END)
        log_widget.update()

    def _todos_botoes(estado):
        for btn in botoes.values():
            btn.config(state=estado)

    def ao_clicar_extrair():
        mes_ano = ler_mes_ano()
        if mes_ano is None:
            return
        mes, ano = mes_ano
        log_widget.delete("1.0", tk.END)
        _todos_botoes("disabled")
        try:
            rodar(mes, ano, log_widget)
        finally:
            _todos_botoes("normal")

    def ao_clicar_check():
        from check_agrupamentos_ksb1 import gerar_check

        mes_ano = ler_mes_ano()
        if mes_ano is None:
            return
        mes, ano = mes_ano

        log_widget.delete("1.0", tk.END)
        _todos_botoes("disabled")
        try:
            caminho = gerar_check(mes, ano, log=log)
            messagebox.showinfo("Concluído", f"Check de agrupamentos gerado:\n{caminho}")
        except Exception as e:
            messagebox.showerror("Erro ao gerar o check", str(e))
        finally:
            _todos_botoes("normal")

    def ao_clicar_pivot():
        from gerar_base_intermediaria import atualizar_base_intermediaria

        mes_ano = ler_mes_ano()
        if mes_ano is None:
            return
        mes, ano = mes_ano
        ciclo = ciclo_var.get()

        log_widget.delete("1.0", tk.END)
        _todos_botoes("disabled")
        try:
            caminho = atualizar_base_intermediaria(mes, ano, ciclo, log=log)
            messagebox.showinfo(
                "Concluído",
                f"Base Intermediária gerada:\n{caminho}\n\nConfira a aba 'Pendências' antes de colar no arquivo de trabalho.",
            )
        except Exception as e:
            messagebox.showerror("Erro ao atualizar a Base Intermediária", str(e))
        finally:
            _todos_botoes("normal")

    botoes[0].config(command=ao_clicar_extrair)
    botoes[1].config(command=ao_clicar_check)
    botoes[2].config(command=ao_clicar_pivot)

    rastro_canvas = tk.Canvas(root, height=18, bg=BG_RODAPE, highlightthickness=0)
    rastro_canvas.pack(fill=tk.X, side=tk.BOTTOM)
    root.update_idletasks()
    desenhar_rastro_pneu(rastro_canvas, root.winfo_width(), 18)

    root.mainloop()


if __name__ == "__main__":
    main()
