# -*- coding: utf-8 -*-
"""mp27_draft.py -- draft de despesas do Master Plan 2027 da TLM (0981).

OBJETIVO: primeira proposta de valores do MP27 sem esperar RH e area tecnica.
OPERACAO: base 2026 ano cheio (Actual jan-ago + Fcst R09 set-dez, intermediaria por CC x conta)
          x indice por conta gestorial (inflacao MP2027 / contrato / zero), fora do perimetro
          Cabreuva e Barueri MP, e escreve o caderno + a carga no layout da YKCCP003_LT (J01/2027).
ALINHAMENTO: e o metodo do MP26 (R09 x inflacao) com o indice aberto por conta, decisao do Caetano
          em 28/09/2026. Nada e escrito no SAP; a carga e so o Excel pronto para revisao.

Decisoes (28/09/2026): indice por conta (A) - base ano cheio (A) - versao J01 ano 2027 -
Cabreuva nao existe em 2027 - faturamento MP Barueri nao existe em 2027.
"""
import os
import sys
from datetime import datetime

import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(BASE_DIR, "scripts", "antidumping"))
from caderno import escrever_caderno                                   # noqa: E402

REDE = r"\\SRV054-01BR.group.pirelli.com\DAC_TLM"
ANO_MP = datetime.now().year + 1
ANO_BASE = ANO_MP - 1
AA_MP, AA_BASE = ANO_MP % 100, ANO_BASE % 100
PASTA_MP = os.path.join(REDE, "MP%d" % AA_MP, "1\u00aa Vers\u00e3o MP%d vs R09" % AA_MP)
PASTA_SAIDA = os.path.join(PASTA_MP, "TLM_Automation")
BASES = os.path.join(REDE, str(ANO_BASE), "Bases")
CENTROS = os.path.join(REDE, str(ANO_BASE), "Centros TLM - %d.xlsx" % ANO_BASE)
CARGA_MP_ANTERIOR = os.path.join(REDE, "MP%d" % AA_BASE, "1\u00aa Vers\u00e3o MP%d vs R09" % AA_BASE,
                                 "Carga Custo MP%d.xlsx" % AA_BASE)

MESES = ["January", "February", "March", "April", "May", "June",
         "July", "August", "September", "October", "November", "December"]
MESES_CARGA = ["JAN", "FEV", "MAR", "ABR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
ULTIMO_MES_ACT = 8                      # Actual ate agosto, R09 cobre set-dez

# Macro MP2027 - Brasil (MP27\Macroeconomics 1st draft 10_09.pdf + Budget 2027 ForEx 1st draft 10_09.pdf)
MACRO = {"inflacao_media": 0.044, "inflacao_fim_periodo": 0.043, "pib_pirelli": 0.015,
         "brl_usd_2026_r9": 5.16, "brl_usd_mp2027": 5.25, "brl_eur_2026_r9": 6.02, "brl_eur_mp2027": 6.14}
FX_USD = MACRO["brl_usd_mp2027"] / MACRO["brl_usd_2026_r9"] - 1

# Perimetro 2027 (Caetano, 28/09/2026)
WH_FORA = {"CBLog": "sai - Cabre\u00fava n\u00e3o existe em 2027",
           "CB_AG": "sai - Cabre\u00fava n\u00e3o existe em 2027",
           "BAMP": "sai - faturamento MP Barueri n\u00e3o existe em 2027"}

# Indice por conta gestorial. Grupo -> (indice, regra). Conta sem grupo cai em INFLACAO e e marcada REVISAR.
GRUPOS = {
    "INFLACAO": (MACRO["inflacao_media"], "infla\u00e7\u00e3o m\u00e9dia Brasil MP2027 (S&P/Pirelli)"),
    "RH":       (0.0, "n\u00e3o indexar - substitu\u00eddo pelo Labour Cost / HC / Layoff do RH"),
    "CONTRATO": (0.0, "IFRS16 / deprecia\u00e7\u00e3o - vem do contrato e do imobilizado, n\u00e3o de \u00edndice"),
    "FX":       (FX_USD, "BRL/USD 5,16 -> 5,25 (ForEx MP2027)"),
    "REVISAR":  (0.0, "conta sem regra clara - decidir antes da 1\u00aa rodada oficial"),
}
CONTAS_RH = ["4200000", "4200100", "4200110", "4200120", "4200130", "4200140", "4200150", "4200160",
             "4200170", "4200180", "4200400", "4201000", "4201100", "4201110", "4201120",
             "4201130", "4201140", "4201150", "4201160", "4201170", "4201180", "4201410",
             "4245000", "4245100", "4247100", "4247101_HR", "4254100", "4254110"]
CONTAS_INFLACAO = ["4202000", "4202100", "4203050", "4203055", "4205250", "4205350", "4211000", "4222000",
                   "4222100", "4230000", "4232000", "4233000", "4234000", "4236020", "4236050", "4236100",
                   "4237000", "4237100", "4238000", "4241000", "4243000", "4243100", "4244000", "4246000",
                   "4247200", "4248000", "4254000", "4257000", "4258000", "4258010", "4260000", "4262100",
                   "4263000"]
CONTA_GRUPO = {c: "RH" for c in CONTAS_RH}
CONTA_GRUPO.update({c: "INFLACAO" for c in CONTAS_INFLACAO})
CONTA_GRUPO.update({
    "4257001": "CONTRATO", "4255002": "CONTRATO", "4255000": "CONTRATO",   # IFRS16 / depreciacao
    "4266000": "REVISAR",                                                  # Ajuste de Custo
    "4200200": "REVISAR",                                                  # Transitoria folha
})


def ler_intermediaria(cenario):
    partes = []
    for fv in ("Fix", "Var"):
        arq = os.path.join(BASES, "Base_interm_%s_%s.XLSX" % (cenario, fv))
        d = pd.read_excel(arq, dtype={"Centro de Custo": str, "Conta Gestorial": str, "Conta Fiscal": str})
        d["VAR."] = fv[0]
        partes.append(d)
    d = pd.concat(partes, ignore_index=True)
    d = d.rename(columns={"Descri\u00e7\u00e3o": "Conta Gestorial Desc", "Descri\u00e7\u00e3o.1": "Classe Desc",
                          "Descri\u00e7\u00e3o.2": "Mini-F\u00e1brica Desc",
                          "Tp.Custo (Pr\u00f3prio /  Aprop.)": "ORIG."})
    d[MESES] = d[MESES].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    return d[["Centro de Custo", "Mini-F\u00e1brica", "Mini-F\u00e1brica Desc", "Conta Gestorial",
              "Conta Gestorial Desc", "Conta Fiscal", "Classe Desc", "VAR.", "ORIG."] + MESES]


def ler_centros():
    c = pd.read_excel(CENTROS, sheet_name=0, header=3, dtype=str).iloc[:, 1:8]
    c.columns = ["Centro de Custo", "CC Desc", "Local", "BU", "WH", "F/V", "Respons\u00e1vel"]
    c = c.dropna(subset=["Centro de Custo"]).drop_duplicates("Centro de Custo")
    return c[["Centro de Custo", "CC Desc", "WH", "Respons\u00e1vel"]]


CHAVE = ["Centro de Custo", "Mini-F\u00e1brica", "Mini-F\u00e1brica Desc", "Conta Gestorial", "Conta Gestorial Desc",
         "Conta Fiscal", "Classe Desc", "VAR.", "ORIG."]


def base_ano_cheio():
    """Actual jan-ago + F09 set-dez, na mesma linha CC x conta fiscal x VAR x ORIG."""
    act = ler_intermediaria("Actual_%02d.%d" % (ULTIMO_MES_ACT, ANO_BASE))
    f09 = ler_intermediaria("Fcst_R09_09.%d" % ANO_BASE)
    act = act.groupby(CHAVE, as_index=False)[MESES[:ULTIMO_MES_ACT]].sum()
    f09 = f09.groupby(CHAVE, as_index=False)[MESES[ULTIMO_MES_ACT:]].sum()
    base = act.merge(f09, on=CHAVE, how="outer").fillna(0.0)
    base = base[base[MESES].abs().sum(axis=1) > 0].copy()
    base["Total %d" % ANO_BASE] = base[MESES].sum(axis=1)
    return base


def montar():
    tot_base = "Total %d" % ANO_BASE
    tot_mp = "Total MP%d" % AA_MP
    print("[1] base %d ano cheio (Actual jan-%02d + R09)" % (ANO_BASE, ULTIMO_MES_ACT))
    base = base_ano_cheio().merge(ler_centros(), on="Centro de Custo", how="left")
    base["WH"] = base["WH"].fillna("? sem WH no Centros TLM")
    base["Per\u00edmetro %d" % ANO_MP] = base["WH"].map(WH_FORA).fillna("fica")
    per = "Per\u00edmetro %d" % ANO_MP

    print("[2] indice por conta")
    base["Grupo"] = base["Conta Gestorial"].map(CONTA_GRUPO)
    base["Revisar"] = base["Grupo"].isna()
    base["Grupo"] = base["Grupo"].fillna("INFLACAO")
    base["\u00cdndice"] = base["Grupo"].map(lambda g: GRUPOS[g][0])

    print("[3] draft MP%d" % AA_MP)
    draft = base[base[per] == "fica"].copy()
    for m in MESES:
        draft[m] = draft[m] * (1 + draft["\u00cdndice"])
    draft[tot_mp] = draft[MESES].sum(axis=1)
    draft["Observa\u00e7\u00e3o"] = draft["Grupo"].map(lambda g: GRUPOS[g][1])
    draft.loc[draft["Revisar"], "Observa\u00e7\u00e3o"] = "sem regra expl\u00edcita - caiu em infla\u00e7\u00e3o; REVISAR"

    print("[4] carga no layout YKCCP003_LT (vers\u00e3o J01 / ano %d)" % ANO_MP)
    carregadas = set(pd.read_excel(CARGA_MP_ANTERIOR, dtype=str)["CLASSE DE CUSTO"].dropna())
    carga = pd.DataFrame({
        "EMPRESA": "0981", "ANO": str(ANO_MP), "VERS\u00c3O": "J01",
        "CENTRO DE CUSTO": draft["Centro de Custo"].values, "CLASSE DE CUSTO": draft["Conta Fiscal"].values,
        "DESCRI\u00c7\u00c3O DA CLASSE": draft["Classe Desc"].values,
        "VAR.": draft["VAR."].values, "ORIG.": draft["ORIG."].values})
    for mc, m in zip(MESES_CARGA, MESES):
        carga[mc] = draft[m].round(2).values
    carga["TOTAL"] = carga[MESES_CARGA].sum(axis=1)
    carga["Carregada no MP%d" % AA_BASE] = carga["CLASSE DE CUSTO"].isin(carregadas).map(
        {True: "sim", False: "N\u00c3O - conferir"})
    carga["Grupo \u00edndice"] = draft["Grupo"].values
    carga = carga[carga["TOTAL"].abs() > 0.005].reset_index(drop=True)
    # A intermediaria e o RESULTADO do SAP: traz classes de simulacao de folha (_HR, PF/PI, N2..., N61...)
    # que nascem do rateio e nunca entram na YKCCP003_LT. So sobe o que ja subiu no MP anterior;
    # o resto fica em aba propria para o Caetano decidir.
    fora_layout = carga[carga["Carregada no MP%d" % AA_BASE] != "sim"].reset_index(drop=True)
    carga = carga[carga["Carregada no MP%d" % AA_BASE] == "sim"].reset_index(drop=True)

    print("[5] resumos")
    r_conta = (base.groupby(["Conta Gestorial", "Conta Gestorial Desc", "Grupo", "\u00cdndice"], as_index=False)[tot_base].sum()
               .merge(draft.groupby("Conta Gestorial", as_index=False)[tot_mp].sum(), on="Conta Gestorial", how="left")
               .fillna({tot_mp: 0.0}))
    r_conta["Delta"] = r_conta[tot_mp] - r_conta[tot_base]
    r_conta["Regra"] = r_conta["Grupo"].map(lambda g: GRUPOS[g][1])
    r_conta = r_conta.sort_values(tot_base, key=abs, ascending=False)

    r_wh = (base.groupby(["WH", per], as_index=False)[tot_base].sum()
            .merge(draft.groupby("WH", as_index=False)[tot_mp].sum(), on="WH", how="left").fillna({tot_mp: 0.0}))
    r_wh["Delta"] = r_wh[tot_mp] - r_wh[tot_base]
    r_wh = r_wh.sort_values(tot_base, key=abs, ascending=False)

    premissas = pd.DataFrame([
        ("Base", "Ano %d cheio = Actual jan-%02d + Fcst R09 set-dez (intermedi\u00e1ria Fix+Var, CC x conta fiscal)"
                 % (ANO_BASE, ULTIMO_MES_ACT)),
        ("Vers\u00e3o SAP", "J01 / ano %d (carga YKCCP003_LT) - N\u00c3O carregada, s\u00f3 o Excel" % ANO_MP),
        ("Infla\u00e7\u00e3o m\u00e9dia Brasil MP2027", "%.1f%% (S&P 4,5%%; 2026 = 4,6%%)" % (MACRO["inflacao_media"] * 100)),
        ("Infla\u00e7\u00e3o fim de per\u00edodo 2027", "%.1f%%" % (MACRO["inflacao_fim_periodo"] * 100)),
        ("PIB Brasil 2027 (Pirelli)", "%.1f%%" % (MACRO["pib_pirelli"] * 100)),
        ("BRL/USD", "2026 R9 %.2f -> MP2027 %.2f (%+.1f%%)" % (MACRO["brl_usd_2026_r9"], MACRO["brl_usd_mp2027"], FX_USD * 100)),
        ("BRL/EUR", "2026 R9 %.2f -> MP2027 %.2f" % (MACRO["brl_eur_2026_r9"], MACRO["brl_eur_mp2027"])),
        ("Per\u00edmetro", "Cabre\u00fava (CBLog + CB_AG) n\u00e3o existe em 2027 - fora inteira, sem migra\u00e7\u00e3o de volume"),
        ("Per\u00edmetro", "Faturamento MP Barueri (BAMP) n\u00e3o existe em 2027 - fora"),
        ("M\u00e3o de obra", "classes 4\u2026 mantidas no valor %d (\u00edndice 0) - RH substitui com Labour Cost / HC / Layoff" % ANO_BASE),
        ("Fora deste draft", "Volume (\u00e1rea t\u00e9cnica) \u00b7 Faturamento e tarifa \u00b7 HC \u00b7 premissas de contrato Campneus/Hevea"),
        ("Fonte macro", r"MP27\Macroeconomics 1st draft 10_09.pdf + Budget 2027 ForEx 1st draft 10_09.pdf"),
    ], columns=["Premissa", "Valor"])

    indices = pd.DataFrame([(g, v[0], v[1]) for g, v in GRUPOS.items()], columns=["Grupo", "\u00cdndice", "Regra"])

    os.makedirs(PASTA_SAIDA, exist_ok=True)
    destino = os.path.join(PASTA_SAIDA, "Draft MP%d Despesas - %s.xlsx" % (AA_MP, datetime.now().strftime("%Y%m%d")))
    sub = "gerado em %s por scripts/tlm/mp27_draft.py - base %d ano cheio x \u00edndice por conta" % (
        datetime.now().strftime("%d/%m/%Y %H:%M"), ANO_BASE)
    cols = ["WH", "Centro de Custo", "CC Desc", "Mini-F\u00e1brica Desc", "Conta Gestorial", "Conta Gestorial Desc",
            "Conta Fiscal", "Classe Desc", "VAR.", "ORIG.", per, "Grupo", "\u00cdndice"]
    escrever_caderno(destino, [
        ("Premissas", "Draft MP%d TLM - premissas" % AA_MP, sub, premissas),
        ("Indice_por_grupo", "\u00cdndice por grupo de conta", "grupo de cada conta no Resumo_Conta", indices),
        ("Resumo_Conta", "%d x MP%d por conta gestorial" % (ANO_BASE, AA_MP),
         "R$ - %s inclui Cabre\u00fava/BAMP; MP%d j\u00e1 sem" % (tot_base, AA_MP), r_conta),
        ("Resumo_WH", "%d x MP%d por WH" % (ANO_BASE, AA_MP), "R$", r_wh),
        ("Base_%d" % ANO_BASE, "Base %d ano cheio" % ANO_BASE, "Actual jan-%02d + R09 set-dez" % ULTIMO_MES_ACT,
         base[cols + MESES + [tot_base]]),
        ("Draft_MP%d" % AA_MP, "Draft MP%d por CC x conta" % AA_MP, "base x (1 + \u00edndice), per\u00edmetro %d" % ANO_MP,
         draft[cols + MESES + [tot_mp, "Observa\u00e7\u00e3o"]]),
        ("Carga_J01", "Carga MP%d - layout YKCCP003_LT" % AA_MP,
         "vers\u00e3o J01 - ano %d - s\u00f3 classes que subiram no MP%d - N\u00c3O carregada no SAP" % (ANO_MP, AA_BASE), carga),
        ("Fora_do_layout", "Classes que n\u00e3o sobem na YKCCP003_LT",
         "resultado de rateio/simula\u00e7\u00e3o de folha no SAP (_HR, PF/PI, N2..., N61...) - decidir o que fazer", fora_layout),
    ])
    print("  fora do layout (%4d linhas): %12.0f" % (len(fora_layout), fora_layout["TOTAL"].sum()))
    print("\nSaida:", destino)
    print("Total %d ano cheio     : %14.0f" % (ANO_BASE, base[tot_base].sum()))
    print("  fora do perimetro %d : %14.0f" % (ANO_MP, base.loc[base[per] != "fica", tot_base].sum()))
    print("Draft MP%d              : %14.0f" % (AA_MP, draft[tot_mp].sum()))
    print("  carga (%4d linhas)    : %14.0f" % (len(carga), carga["TOTAL"].sum()))
    print("Contas sem regra (REVISAR):", sorted(base.loc[base["Revisar"], "Conta Gestorial"].unique()))
    return destino


if __name__ == "__main__":
    montar()
