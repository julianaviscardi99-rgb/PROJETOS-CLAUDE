#!/usr/bin/env python3
"""
Preenche a tela de selecao da KSB1 (despesas) e executa, para uma BU e periodo
informados interativamente. Nao clica em nada que altere dados no SAP.

Pre-requisito: SAP GUI aberto, logada, com a transacao KSB1 aberta na tela de
selecao (1a tela) antes de rodar este script.
"""
import calendar
import json
import shutil
import sys
from pathlib import Path

import win32com.client

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
ONTOLOGY_DIR = PROJECT_ROOT / "ontology"

# Area de rede da Pirelli onde a Juliana quer uma copia dos exports brutos da
# KSB1 (apenas Fitted Units por enquanto). Nao e um caminho pessoal do
# usuario, e um recurso corporativo compartilhado - por isso fica fixo aqui.
REDE_FITTED_BASE = Path(
    r"\\FSS024-01BR.group.pirelli.com\GFU_DAC\Custos Fitted Units\Resultados Fitted"
)
MESES_PASTA = {
    1: "01 - Jan", 2: "02 - Feb", 3: "03 - Mar", 4: "04 - Apr",
    5: "05 - May", 6: "06 - Jun", 7: "07 - Jul", 8: "08 - Aug",
    9: "09 - Sep", 10: "10 - Oct", 11: "11 - Nov", 12: "12 - Dec",
}

# Valores conhecidos por BU (grupo de centro de custo e variante de exibicao).
# None = ainda nao mapeado; o script pergunta e oferece salvar para a proxima vez.
BUS = {
    "1": {
        "nome": "Fitted Units",
        "arquivo_ontologia": "fitted_units.json",
        "kstgr": "0495",
        "disvar": "/DESPFITTED",
    },
    "2": {
        "nome": "Circuito Panamericano",
        "arquivo_ontologia": "circuito_panamericano.json",
        "kstgr": None,
        "disvar": None,
    },
}


def connect_session():
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(0)
    return connection.Children(0)


def perguntar_bu():
    print("Qual BU?")
    for chave, bu in BUS.items():
        print(f"  {chave}) {bu['nome']}")
    escolha = input("Escolha (1/2): ").strip()
    if escolha not in BUS:
        print("Opcao invalida.")
        sys.exit(1)
    return BUS[escolha]


def perguntar_periodo():
    mes = int(input("Mes (1-12): ").strip())
    ano = int(input("Ano (ex: 2026): ").strip())
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    data_de = f"{1:02d}.{mes:02d}.{ano}"
    data_ate = f"{ultimo_dia:02d}.{mes:02d}.{ano}"
    return mes, ano, data_de, data_ate


def perguntar_agrupamento():
    print("Agrupamento:")
    print("  1) Com gestoriais")
    print("  2) Sem agrupamento")
    escolha = input("Escolha (1/2): ").strip()
    return "gestoriais" if escolha == "1" else ""


def garantir_parametros_bu(bu):
    if bu["kstgr"] and bu["disvar"]:
        return bu

    print(f"\nAinda nao tenho o 'Grupo de centro de custo' e a 'Variante de exibicao' de {bu['nome']} mapeados.")
    kstgr = input("Grupo de centro de custo (ex: 0495): ").strip()
    disvar = input("Variante de exibicao (ex: /DESPFITTED): ").strip()
    bu["kstgr"] = kstgr
    bu["disvar"] = disvar

    salvar = input("Salvar esses valores na ontologia para a proxima vez? (s/n): ").strip().lower()
    if salvar == "s":
        caminho = ONTOLOGY_DIR / bu["arquivo_ontologia"]
        dados = json.loads(caminho.read_text(encoding="utf-8"))
        dados.setdefault("sistemas", {}).setdefault("transacoes_sap", [])
        entrada_ksb1 = None
        for t in dados["sistemas"]["transacoes_sap"]:
            if t.get("transacao") == "KSB1":
                entrada_ksb1 = t
                break
        if entrada_ksb1 is None:
            entrada_ksb1 = {"transacao": "KSB1", "uso": "Extração de despesas."}
            dados["sistemas"]["transacoes_sap"].append(entrada_ksb1)
        entrada_ksb1["parametros_observados"] = {
            "grupo_de_centros_de_custo": kstgr,
            "variante_de_exibicao": disvar,
        }
        caminho.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"Salvo em {caminho.relative_to(PROJECT_ROOT)}")

    return bu


def preencher_e_executar(session, bu, data_de, data_ate, koagr):
    wnd = session.FindById("wnd[0]")

    if wnd.FindById("usr/ctxtKSTGR", False) is not None:
        wnd.FindById("usr/ctxtKSTGR").Text = bu["kstgr"]

    wnd.FindById("usr/ctxtKOAGR").Text = koagr
    wnd.FindById("usr/ctxtR_BUDAT-LOW").Text = data_de
    wnd.FindById("usr/ctxtR_BUDAT-HIGH").Text = data_ate
    wnd.FindById("usr/ctxtP_DISVAR").Text = bu["disvar"]

    print(f"\nFiltros preenchidos: KSTGR={bu['kstgr']} | KOAGR='{koagr}' | Periodo={data_de} a {data_ate} | Variante={bu['disvar']}")
    confirmar = input("Executar a KSB1 agora (F8)? (s/n): ").strip().lower()
    if confirmar != "s":
        print("Cancelado pela usuaria. Filtros ficam preenchidos na tela para revisao manual.")
        return False

    session.FindById("wnd[0]").SendVKey(8)
    print("Executado. Verifique o resultado na tela do SAP.")
    return True


def exportar_para_excel(session, bu, mes, ano, koagr):
    confirmar = input("\nExportar o resultado para Excel em data/raw/ agora? (s/n): ").strip().lower()
    if confirmar != "s":
        print("Exportacao pulada. O resultado continua na tela do SAP se quiser exportar manualmente.")
        return

    pasta_saida = PROJECT_ROOT / "data" / "raw"
    pasta_saida.mkdir(parents=True, exist_ok=True)

    bu_slug = bu["nome"].lower().replace(" ", "_")
    agrup_slug = "gestoriais" if koagr else "sem_agrupamento"
    nome_arquivo = f"KSB1_{bu_slug}_{ano}{mes:02d}_{agrup_slug}.XLSX"

    session.FindById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select()
    wnd1 = session.FindById("wnd[1]")
    wnd1.FindById("usr/ctxtDY_PATH").Text = str(pasta_saida)
    wnd1.FindById("usr/ctxtDY_FILENAME").Text = nome_arquivo
    wnd1.FindById("tbar[0]/btn[0]").Press()  # botao "Gerar"

    arquivo_local = pasta_saida / nome_arquivo
    print(f"\nExportacao solicitada: {arquivo_local}")
    print("Se aparecer um popup 'Seguranca SAPGUI' pedindo autorizacao para criar o arquivo,")
    print("clique em 'Permitir' na tela do SAP para concluir a exportacao.")
    input("Depois de confirmar no SAP (e o arquivo aparecer salvo), aperte Enter aqui... ")

    copiar_para_rede(bu, mes, ano, arquivo_local)


def copiar_para_rede(bu, mes, ano, arquivo_local):
    if bu["nome"] != "Fitted Units":
        print(f"Copia para a rede ainda nao configurada para {bu['nome']}. Pulando.")
        return

    if not arquivo_local.exists():
        print(f"AVISO: {arquivo_local} nao foi encontrado, nao copiei para a rede.")
        print("Confirme se a exportacao no SAP realmente terminou e rode a copia manualmente se precisar.")
        return

    pasta_rede = REDE_FITTED_BASE / str(ano) / "00.Extração Base KSB1" / MESES_PASTA[mes]
    try:
        pasta_rede.mkdir(parents=True, exist_ok=True)
        destino = pasta_rede / arquivo_local.name
        shutil.copy2(arquivo_local, destino)
        print(f"Copiado para a rede: {destino}")
    except OSError as e:
        print(f"AVISO: nao consegui copiar para a rede ({e}). O arquivo local em data/raw/ esta ok.")


def main():
    try:
        session = connect_session()
    except Exception as e:
        print(f"ERRO ao conectar: {e}")
        sys.exit(1)

    if session.Info.Transaction != "KSB1":
        print(f"A transacao atual e '{session.Info.Transaction}', nao KSB1.")
        print("Abra a KSB1 (tela de selecao) no SAP e rode o script de novo.")
        sys.exit(1)

    bu = perguntar_bu()
    bu = garantir_parametros_bu(bu)
    mes, ano, data_de, data_ate = perguntar_periodo()
    koagr = perguntar_agrupamento()

    executou = preencher_e_executar(session, bu, data_de, data_ate, koagr)
    if executou:
        exportar_para_excel(session, bu, mes, ano, koagr)


if __name__ == "__main__":
    main()
