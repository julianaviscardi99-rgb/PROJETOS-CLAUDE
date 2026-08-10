#!/usr/bin/env python3
"""
Janela grafica para extrair a KSB1 da Fitted Units (Gestoriais + Sem Agrupamento)
e salvar os arquivos na area de rede. Feito para qualquer pessoa com acesso a
rede usar, sem precisar de terminal nem conhecer o projeto.

Script autossuficiente: nao depende de outros arquivos do projeto (nao le/
escreve ontologia). Pode ser copiado sozinho para a area de rede.

Pre-requisitos na maquina de quem for rodar:
- Python 3 com pywin32 instalado (pip install pywin32)
- SAP GUI Scripting habilitado (Alt+F12 > Opcoes > Acessibilidade e Scripting > Scripting)
- SAP GUI aberto, logada, com a transacao KSB1 aberta na tela de selecao
"""
import shutil
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk

import win32com.client

BU = {"nome": "Fitted Units", "kstgr": "0495", "disvar": "/DESPFITTED"}

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
    nome_arquivo = f"KSB1 - {BU['nome']} {mes:02d}.{ano} - {agrup_label}.XLSX"

    session.FindById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select()
    wnd1 = session.FindById("wnd[1]")
    wnd1.FindById("usr/ctxtDY_PATH").Text = str(pasta_rede)
    wnd1.FindById("usr/ctxtDY_FILENAME").Text = nome_arquivo
    wnd1.FindById("tbar[0]/btn[0]").Press()  # Gerar

    messagebox.showinfo(
        "Confirme no SAP",
        f"Se aparecer o popup 'Segurança SAPGUI' pedindo autorização, clique em 'Permitir' "
        f"na tela do SAP.\n\nDepois clique OK aqui para continuar.",
    )

    arquivo_final = pasta_rede / nome_arquivo
    if arquivo_final.exists():
        log(f"{agrup_label}: salvo em {arquivo_final}")
    else:
        log(f"AVISO: não encontrei o arquivo de {agrup_label} na pasta esperada. Confira manualmente.")

    # Volta para a tela de selecao para a proxima extracao (ou deixa limpo no final)
    session.FindById("wnd[0]").SendVKey(3)


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

    if session.Info.Transaction != "KSB1":
        log("Abrindo a transação KSB1...")
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nKSB1"
        session.FindById("wnd[0]").SendVKey(0)  # Enter

    if session.Info.Transaction != "KSB1":
        messagebox.showerror(
            "Não consegui abrir a KSB1",
            f"Tentei abrir a transação KSB1, mas a tela atual é '{session.Info.Transaction}'.\n\n"
            "Confirme se seu usuário tem acesso à transação KSB1 e tente de novo.",
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


def main():
    root = tk.Tk()
    root.title("Atualizar KSB1 - Fitted Units")
    root.geometry("420x320")

    frame = ttk.Frame(root, padding=16)
    frame.pack(fill=tk.BOTH, expand=True)

    ttk.Label(frame, text="Mês:").grid(row=0, column=0, sticky="w", pady=4)
    mes_var = tk.StringVar(value=MESES_NOMES[datetime.now().month])
    mes_combo = ttk.Combobox(frame, textvariable=mes_var, values=list(MESES_NOMES.values()), state="readonly")
    mes_combo.grid(row=0, column=1, sticky="ew", pady=4)

    ttk.Label(frame, text="Ano:").grid(row=1, column=0, sticky="w", pady=4)
    ano_var = tk.StringVar(value=str(datetime.now().year))
    ano_entry = ttk.Entry(frame, textvariable=ano_var)
    ano_entry.grid(row=1, column=1, sticky="ew", pady=4)

    frame.columnconfigure(1, weight=1)

    log_widget = tk.Text(frame, height=10, wrap="word")
    log_widget.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(12, 0))
    frame.rowconfigure(3, weight=1)

    def ao_clicar_extrair():
        nome_para_numero = {v: k for k, v in MESES_NOMES.items()}
        mes = nome_para_numero[mes_var.get()]
        try:
            ano = int(ano_var.get())
        except ValueError:
            messagebox.showerror("Ano inválido", "Digite um ano válido, ex: 2026.")
            return
        extrair_btn.config(state="disabled")
        try:
            rodar(mes, ano, log_widget)
        finally:
            extrair_btn.config(state="normal")

    extrair_btn = ttk.Button(frame, text="Extrair KSB1 (Gestoriais + Sem Agrupamento)", command=ao_clicar_extrair)
    extrair_btn.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(12, 0))

    ttk.Label(
        frame,
        text="Antes de clicar: deixe o SAP GUI aberto, logada, com a KSB1 na tela de seleção.",
        wraplength=380,
        foreground="#555",
    ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))

    root.mainloop()


if __name__ == "__main__":
    main()
