# -*- coding: utf-8 -*-
"""mp27_base_areas.py -- despesas do MP27 da TLM (0981) a partir da base enviada pelas areas.

OBJETIVO: o Resultado TLM - MP27 vs R09 passa a refletir a base de custos que as areas montaram item
          a item (e-mail "MP27 | Base custos TLM", Rafaela, 09/10/2026), no lugar do R09 x indice.
OPERACAO: le a versao "MP27 Inf" da aba Base (ja com inflacao 4,4 %), poe WH pelo CC (Centros TLM),
          VAR. pelo R09/Actual 2026 do mesmo CC x classe, e troca TODAS as despesas nao-RH do draft
          por ela. Folha/beneficios (grupo RH) continuam do draft por indice ate o RH mandar o MP27.
          Grava "Draft MP27 Despesas - <data> - base areas.xlsx" -- o mp27_resultado.py pega o ultimo.
ALINHAMENTO: a base das areas e o metodo do MP26 (R09 ajustado item a item pelo responsavel); o draft
          por indice era so a ponte ate ela chegar. Nada e escrito no SAP.

Decisoes do Caetano (14/10/2026): VAR. = a do R09 do mesmo CC x classe (a base das areas vem toda F,
e a tarifa variavel sai do custo variavel da WH); linhas IFRS16 repetidas ficam como vieram ate a
Rafaela responder.
"""
import glob
import os
import sys
from datetime import datetime

import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(BASE_DIR, "scripts", "antidumping"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from caderno import escrever_caderno                                   # noqa: E402
import mp27_draft as dr                                                # noqa: E402

PASTA_BASES = os.path.join(dr.PASTA_MP, "Arquivos Recebidos")
VERSAO_AREAS = "MP%d Inf" % dr.AA_MP
SUFIXO = " - base unica"
PREFIXO_IANA = "MP27 RH TLM 0981 - Iana Brito"
ARQ_LC = "Labour Cost TLM - MP%d.xlsx" % dr.AA_MP
FONTE, ENVIADO = "Fonte", "Enviado por"
TOT_MP = "Total MP%d" % dr.AA_MP
PER = "Perímetro %d" % dr.ANO_MP
COLS = ["WH", "Centro de Custo", "CC Desc", "Mini-Fábrica Desc", "Conta Gestorial", "Conta Gestorial Desc",
        "Conta Fiscal", "Classe Desc", "VAR.", "ORIG.", PER, "Grupo", "Índice"]


def _ultimo(padrao, excluir=None):
    arqs = sorted(a for a in glob.glob(padrao) if not (excluir and excluir in os.path.basename(a))
                  and not os.path.basename(a).startswith("~$"))
    if not arqs:
        raise FileNotFoundError(padrao)
    return arqs[-1]


def ler_base_areas(arq):
    d = pd.read_excel(arq, sheet_name="Base", header=3)
    d.columns = [str(c).strip() for c in d.columns]
    d["Linha Excel"] = d.index + 5
    d = d[d["Emp"].notna() & (d["Versão"] == VERSAO_AREAS)].copy()
    d[dr.MESES] = d[dr.MESES].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    d["Centro de Custo"] = d["CC"].astype(int).astype(str)
    d["Conta Gestorial"] = d["Conta Gest"].astype(int).astype(str)
    d["Conta Fiscal"] = d["Conta Fisc"].astype(str).str.strip()
    d[TOT_MP] = d[dr.MESES].sum(axis=1)
    return d.rename(columns={"WH": "Site"})


def var_do_r09(base):
    """VAR. dominante (maior valor absoluto no ano 2026) por CC x classe e, na falta, por CC x conta gestorial."""
    b = base.assign(abs_=base["Total %d" % dr.ANO_BASE].abs())
    por_classe = (b.groupby(["Centro de Custo", "Conta Fiscal", "VAR."])["abs_"].sum().reset_index()
                  .sort_values("abs_").drop_duplicates(["Centro de Custo", "Conta Fiscal"], keep="last")
                  .set_index(["Centro de Custo", "Conta Fiscal"])["VAR."])
    por_conta = (b.groupby(["Centro de Custo", "Conta Gestorial", "VAR."])["abs_"].sum().reset_index()
                 .sort_values("abs_").drop_duplicates(["Centro de Custo", "Conta Gestorial"], keep="last")
                 .set_index(["Centro de Custo", "Conta Gestorial"])["VAR."])
    mini = (b.groupby(["Centro de Custo", "Mini-Fábrica Desc"])["abs_"].sum().reset_index()
            .sort_values("abs_").drop_duplicates("Centro de Custo", keep="last")
            .set_index("Centro de Custo")["Mini-Fábrica Desc"])
    return por_classe, por_conta, mini


def comparativo_r09(base, areas_mp, minimo_ponte=300000.0):
    """R09 ano cheio (sem RH) x MP27 da base das areas: ponte, por WH, por conta e WH x conta.
    O R09 e pos-rateio (A + P no CC que recebe) e a base das areas esta no CC final - batem por WH."""
    tot_r09 = "R09 %d FY" % dr.ANO_BASE
    sem_inf = "MP%d sem inflação" % dr.AA_MP
    b = base.merge(dr.ler_centros(), on="Centro de Custo", how="left")
    b = b[b["Conta Gestorial"].map(dr.CONTA_GRUPO).fillna("INFLACAO") != "RH"].copy()
    b["WH"] = b["WH"].fillna("? sem WH no Centros TLM")
    b[tot_r09] = b["Total %d" % dr.ANO_BASE]
    fica = ~b["WH"].isin(list(dr.WH_FORA))
    a = areas_mp.copy()
    a[sem_inf] = a[TOT_MP] / (1 + dr.MACRO["inflacao_media"])

    def juntar(chaves):
        x = pd.concat([b[fica].groupby(chaves)[tot_r09].sum(), a.groupby(chaves)[sem_inf].sum(),
                       a.groupby(chaves)[TOT_MP].sum()], axis=1).fillna(0.0)
        x["Variação de base"] = x[sem_inf] - x[tot_r09]
        x["Inflação"] = x[TOT_MP] - x[sem_inf]
        x["Delta MP x R09"] = x[TOT_MP] - x[tot_r09]
        x["Delta %"] = (x["Delta MP x R09"] / x[tot_r09].where(x[tot_r09].abs() > 0.5)).round(4)
        return x.reset_index()

    desc = pd.concat([b.drop_duplicates("Conta Gestorial").set_index("Conta Gestorial")["Conta Gestorial Desc"],
                      a.drop_duplicates("Conta Gestorial").set_index("Conta Gestorial")["Conta Gestorial Desc"]])
    desc = desc[~desc.index.duplicated()]
    por_conta = juntar(["Conta Gestorial"])
    por_conta.insert(1, "Descrição", por_conta["Conta Gestorial"].map(desc))
    por_conta = por_conta.sort_values("Delta MP x R09", key=abs, ascending=False)
    por_wh = juntar(["WH"]).sort_values(tot_r09, key=abs, ascending=False)
    wh_conta = juntar(["WH", "Conta Gestorial"])
    wh_conta.insert(2, "Descrição", wh_conta["Conta Gestorial"].map(desc))
    wh_conta = wh_conta[wh_conta[[tot_r09, TOT_MP]].abs().sum(axis=1) > 0.5].sort_values(["WH", "Delta MP x R09"])

    ponte = [("R09 %d ano cheio - despesas sem RH (Actual jan-%02d + R09 set-dez)" % (dr.ANO_BASE, dr.ULTIMO_MES_ACT),
              b[tot_r09].sum())]
    for wh, motivo in dr.WH_FORA.items():
        v = b.loc[b["WH"] == wh, tot_r09].sum()
        if abs(v) > 0.5:
            ponte.append(("%s (%s)" % (motivo, wh), -v))
    ponte.append(("Subtotal - R09 no perímetro %d" % dr.ANO_MP, b.loc[fica, tot_r09].sum()))
    grandes = por_conta[por_conta["Variação de base"].abs() >= minimo_ponte].sort_values("Variação de base")
    for _, r in grandes.iterrows():
        ponte.append(("variação de base - %s %s" % (r["Conta Gestorial"], r["Descrição"]), r["Variação de base"]))
    ponte.append(("variação de base - demais contas (|delta| < %.0f mil)" % (minimo_ponte / 1000),
                  por_conta["Variação de base"].sum() - grandes["Variação de base"].sum()))
    ponte.append(("inflação %.1f%% (versão '%s')" % (dr.MACRO["inflacao_media"] * 100, VERSAO_AREAS),
                  por_conta["Inflação"].sum()))
    ponte.append(("Total - MP%d base das áreas" % dr.AA_MP, a[TOT_MP].sum()))
    ponte.append(("Delta MP%d x R09 (com Cabreúva/BAMP no R09)" % dr.AA_MP, a[TOT_MP].sum() - b[tot_r09].sum()))
    ponte = pd.DataFrame(ponte, columns=["Passo", "R$"])
    ponte["R$ mil"] = (ponte["R$"] / 1000).round(0)
    return ponte, por_wh, por_conta, wh_conta


def wh_e_var(d, base, var_propria=None):
    """WH pelo CC (Centros TLM) e VAR.: a da propria fonte quando confiavel (Labour Cost), senao a do R09 do
    mesmo CC x classe, senao CC x conta gestorial, senao F/V do CC."""
    centros = dr.ler_centros()
    centros_fv = pd.read_excel(dr.CENTROS, sheet_name=0, header=3, dtype=str).iloc[:, [1, 6]]
    centros_fv.columns = ["Centro de Custo", "F/V CC"]
    d = d.merge(centros, on="Centro de Custo", how="left").merge(
        centros_fv.dropna().drop_duplicates("Centro de Custo"), on="Centro de Custo", how="left")
    d["WH"] = d["WH"].fillna("? sem WH no Centros TLM")
    por_classe, por_conta, mini = var_do_r09(base)
    v0 = pd.Series(list(var_propria) if var_propria is not None else [None] * len(d), index=d.index, dtype=object)
    v1 = pd.Series([por_classe.get(k) for k in zip(d["Centro de Custo"], d["Conta Fiscal"])], index=d.index)
    v2 = pd.Series([por_conta.get(k) for k in zip(d["Centro de Custo"], d["Conta Gestorial"])], index=d.index)
    v3 = d["F/V CC"].str.strip().str[:1].where(d["F/V CC"].str.strip().str[:1].isin(["F", "V"]))
    d["VAR."] = v0.fillna(v1).fillna(v2).fillna(v3).fillna("F")
    d["Origem VAR."] = (v0.notna().map({True: "da própria fonte", False: None})
                        .fillna(v1.notna().map({True: "R09 CC x classe", False: None}))
                        .fillna(v2.notna().map({True: "R09 CC x conta gest.", False: None}))
                        .fillna(v3.notna().map({True: "Centros TLM (F/V do CC)", False: None}))
                        .fillna("sem referência - F"))
    d["Mini-Fábrica Desc"] = d["Centro de Custo"].map(mini).fillna("")
    d[PER] = d["WH"].map(dr.WH_FORA).fillna("fica")
    return d


def mapa_contas(base, areas):
    """conta fiscal -> conta gestorial e descricoes, pelo R09 e pela base das areas."""
    m = pd.concat([base[["Conta Fiscal", "Conta Gestorial", "Conta Gestorial Desc", "Classe Desc"]],
                   areas[["Conta Fiscal", "Conta Gestorial", "Descrição", "Descrição2"]].set_axis(
                       ["Conta Fiscal", "Conta Gestorial", "Conta Gestorial Desc", "Classe Desc"], axis=1)])
    m["Conta Fiscal"] = m["Conta Fiscal"].astype(str).str.strip()
    return m.drop_duplicates("Conta Fiscal").set_index("Conta Fiscal")


def _contas(d, mapa):
    """Conta gestorial e descricoes pelo plano (R09 + base das areas); a classe fica exatamente como veio.
    Classe que nao existe no plano da 0981 fica sem gestorial - nada e inventado (Caetano 15/10/2026)."""
    d["Conta Fiscal"] = d["Conta Fiscal"].astype(str).str.strip()
    gest = d["Conta Fiscal"].map(mapa["Conta Gestorial"])
    d["Conta Gestorial"] = gest.fillna(d["Conta Fiscal"].str.extract(r"^(\d{7})")[0]).fillna("? sem conta gestorial")
    d["Conta Gestorial Desc"] = d["Conta Fiscal"].map(mapa["Conta Gestorial Desc"]).fillna(d["Classe Desc"])
    vazia = d["Classe Desc"].isna() | (d["Classe Desc"].astype(str).str.strip().isin(["", "nan"]))
    d["Classe Desc"] = d["Classe Desc"].where(~vazia, d["Conta Fiscal"].map(mapa["Classe Desc"]))
    return d


def ler_iana(arq, mapa):
    """Custo fixo do RH (Iana Brito), ja filtrado na 0981 por mp27_rh_iana_981.py."""
    i = pd.read_excel(arq, sheet_name="Base_0981", header=3, dtype={"CC": str, "CLASSE DE CUSTO": str})
    meses_iana = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
    d = pd.DataFrame({"Centro de Custo": i["CC"].str.strip(), "Conta Fiscal": i["CLASSE DE CUSTO"],
                      "Classe Desc": i.iloc[:, i.columns.get_loc("CLASSE DE CUSTO") + 1]})
    for m, mi in zip(dr.MESES, meses_iana):
        d[m] = pd.to_numeric(i[mi], errors="coerce").fillna(0.0)
    d[TOT_MP] = d[dr.MESES].sum(axis=1)
    d["ORIG."] = i["ORIG."]
    d["VAR. Iana"] = i["VAR."].astype(str).str.strip().str.upper().where(i["VAR."].notna())
    d["Grupo"] = "RH_FIXO"
    d["Observação"] = ("RH custo fixo - " + i["OBSERVAÇÃO"].fillna("").astype(str) + " "
                       + i["JUSTIFICATIVA"].fillna("").astype(str)).str.strip()
    d[FONTE] = "Custo fixo RH 0981 - %s" % os.path.basename(arq)
    d[ENVIADO] = "Iana Brito (RH corporativo)"
    return _contas(d, mapa)


def ler_labour_cost(arq, mapa):
    """Mao de obra do Pyplan (RH) - e-mail 'Versão Disponível Labour Cost - TLM'."""
    lc = pd.read_excel(arq, dtype={"Cost Center": str, "SAP Account": str})
    lc = lc[lc["SAP Code"].astype(str).str.lstrip("0") == "981"].reset_index(drop=True)
    meses_lc = ["%d.%02d" % (dr.ANO_MP, m) for m in range(1, 13)]
    d = pd.DataFrame({"Centro de Custo": lc["Cost Center"].str.strip(), "Conta Fiscal": lc["SAP Account"],
                      "Classe Desc": lc["SAP Account Description"].astype(str).str.strip()})
    for m, ml in zip(dr.MESES, meses_lc):
        d[m] = pd.to_numeric(lc[ml], errors="coerce").fillna(0.0)
    d[TOT_MP] = d[dr.MESES].sum(axis=1)
    d["ORIG."] = "P"
    d["Grupo"] = "MDO"
    d["Observação"] = "MDO - " + lc["Category"].astype(str) + " / " + lc["Function"].astype(str)
    d[FONTE] = "Labour Cost MP%d (Pyplan) - %s" % (dr.AA_MP, os.path.basename(arq))
    d[ENVIADO] = "RH via Pyplan (No Reply - Pyplan)"
    var = lc["Cost Center Variability"].astype(str).str.strip().str[:1].str.upper().map({"F": "F", "V": "V"})
    return _contas(d, mapa), list(var)


def montar():
    arq_areas = _ultimo(os.path.join(PASTA_BASES, "*MP'%d Custos TLM*.xlsx" % dr.AA_MP))
    arq_draft = _ultimo(os.path.join(dr.PASTA_SAIDA, "Draft MP%d Despesas - ????????.xlsx" % dr.AA_MP))
    arq_iana = _ultimo(os.path.join(PASTA_BASES, PREFIXO_IANA + "*.xlsx"))
    arq_lc = os.path.join(PASTA_BASES, ARQ_LC)
    print("[1] base das areas  <-", os.path.basename(arq_areas))
    print("    custo fixo RH    <-", os.path.basename(arq_iana))
    print("    mao de obra      <-", os.path.basename(arq_lc))
    print("    draft por indice <-", os.path.basename(arq_draft), "(so comparacao)")
    areas = ler_base_areas(arq_areas)
    draft = pd.read_excel(arq_draft, sheet_name="Draft_MP%d" % dr.AA_MP, header=3,
                          dtype={"Centro de Custo": str, "Conta Gestorial": str, "Conta Fiscal": str})
    draft = draft[draft["Conta Gestorial"].notna()].copy()

    print("[2] WH pelo CC, VAR. pelo R09 2026 (MDO: VAR. do Labour Cost)")
    base = dr.base_ano_cheio()
    mapa = mapa_contas(base, areas)
    areas = wh_e_var(areas, base)
    sem_wh = sorted(areas.loc[areas["WH"] == "? sem WH no Centros TLM", "Centro de Custo"].unique())
    areas["ORIG."] = "A"
    areas[FONTE] = "Base custos das áreas - %s (versão %s)" % (os.path.basename(arq_areas), VERSAO_AREAS)
    areas[ENVIADO] = "Rafaela Moreira (e-mail MP27 | Base custos TLM, 09/10)"
    areas["Grupo"] = "AREAS"
    areas["Índice"] = dr.MACRO["inflacao_media"]
    areas["Conta Gestorial Desc"] = areas["Descrição"]
    areas["Classe Desc"] = areas["Descrição2"].astype(str).str.strip()
    areas["Observação"] = ("base das áreas (" + areas["Área"].astype(str) + ") - "
                                + areas["Comentários"].fillna("").astype(str))
    fora = areas[areas[PER] != "fica"]
    areas_mp = areas[(areas[PER] == "fica") & (areas[TOT_MP].abs() > 0.005)]

    ponte, cmp_wh, cmp_conta, cmp_wh_conta = comparativo_r09(base, areas_mp)

    print("[3] base unica = areas + custo fixo RH (Iana) + mao de obra (Labour Cost), somadas")
    iana = ler_iana(arq_iana, mapa)
    iana = wh_e_var(iana, base, var_propria=iana["VAR. Iana"])
    lc, var_lc = ler_labour_cost(arq_lc, mapa)
    lc = wh_e_var(lc, base, var_propria=var_lc)
    for x in (iana, lc):
        x["Índice"] = 0.0
    rh_fora = pd.concat([iana, lc])
    rh_fora = rh_fora[rh_fora[PER] != "fica"]
    iana_mp = iana[(iana[PER] == "fica") & (iana[TOT_MP].abs() > 0.005)]
    lc_mp = lc[(lc[PER] == "fica") & (lc[TOT_MP].abs() > 0.005)]
    rh = draft[draft["Grupo"] == "RH"]
    sai = draft[draft["Grupo"] != "RH"]
    cols_novo = [FONTE, ENVIADO] + COLS + dr.MESES + [TOT_MP, "Observação"]
    novo = pd.concat([areas_mp[cols_novo], iana_mp[cols_novo], lc_mp[cols_novo]], ignore_index=True)
    kc = ["Centro de Custo", "Conta Fiscal"]
    sobre = novo.groupby(kc + [ENVIADO])[TOT_MP].sum().unstack(ENVIADO).fillna(0.0)
    sobre = sobre[(sobre != 0).sum(axis=1) > 1].reset_index()
    por_fonte = novo.groupby([ENVIADO, FONTE, "Grupo"], as_index=False)[TOT_MP].sum()
    por_fonte = pd.concat([por_fonte, pd.DataFrame([{ENVIADO: "TOTAL", TOT_MP: por_fonte[TOT_MP].sum()}])])

    print("[4] resumos e carga")
    r_conta = pd.concat([sai.groupby(["Conta Gestorial"])[TOT_MP].sum().rename("Draft por índice"),
                         areas_mp.groupby(["Conta Gestorial"])[TOT_MP].sum().rename("Base das áreas")],
                        axis=1).fillna(0.0)
    descr = pd.concat([sai.set_index("Conta Gestorial")["Conta Gestorial Desc"],
                       areas_mp.set_index("Conta Gestorial")["Conta Gestorial Desc"]])
    r_conta.insert(0, "Descrição", descr[~descr.index.duplicated()].reindex(r_conta.index))
    r_conta["Delta"] = r_conta["Base das áreas"] - r_conta["Draft por índice"]
    r_conta = r_conta.reset_index().sort_values("Delta", key=abs, ascending=False)
    r_wh = pd.concat([sai.groupby("WH")[TOT_MP].sum().rename("Draft por índice"),
                      areas_mp.groupby("WH")[TOT_MP].sum().rename("Base das áreas"),
                      novo.groupby("WH")[TOT_MP].sum().rename("MP27 novo (com RH)")], axis=1).fillna(0.0)
    r_wh["Delta despesas"] = r_wh["Base das áreas"] - r_wh["Draft por índice"]
    r_wh = r_wh.reset_index().rename(columns={"index": "WH"}).sort_values("MP27 novo (com RH)", key=abs, ascending=False)
    r_var = areas_mp.groupby(["Origem VAR.", "VAR."], as_index=False)[TOT_MP].sum()

    carga = pd.DataFrame({
        "EMPRESA": "0981", "ANO": str(dr.ANO_MP), "VERSÃO": "J01",
        "CENTRO DE CUSTO": novo["Centro de Custo"].values, "CLASSE DE CUSTO": novo["Conta Fiscal"].values,
        "DESCRIÇÃO DA CLASSE": novo["Classe Desc"].values,
        "VAR.": novo["VAR."].values, "ORIG.": novo["ORIG."].values})
    for mc, m in zip(dr.MESES_CARGA, dr.MESES):
        carga[mc] = novo[m].round(2).values
    carga["TOTAL"] = carga[dr.MESES_CARGA].sum(axis=1)
    carga[ENVIADO] = novo[ENVIADO].values

    chave_dup = ["Conta Fisc", "CC", "Comentários"] + dr.MESES
    a_bruto = ler_base_areas(arq_areas)
    mp_sem_inf = pd.read_excel(arq_areas, sheet_name="Base", header=3)
    mp_sem_inf.columns = [str(c).strip() for c in mp_sem_inf.columns]
    mp_sem_inf["Linha Excel"] = mp_sem_inf.index + 5
    mp_sem_inf = mp_sem_inf[mp_sem_inf["Versão"] == "MP%d" % dr.AA_MP]
    dup = mp_sem_inf[mp_sem_inf.duplicated(chave_dup, keep=False)
                     & mp_sem_inf["Conta Fisc"].astype(str).str.endswith("RT")]
    dup = dup[["Linha Excel", "WH", "Conta Fisc", "Descrição2", "CC", "Comentários", "Total Ano"]]

    premissas = pd.DataFrame([
        ("Fonte das despesas", "%s - versão '%s' (já com inflação %.1f%%) - %d linhas"
         % (os.path.basename(arq_areas), VERSAO_AREAS, dr.MACRO["inflacao_media"] * 100, len(a_bruto))),
        ("Custo fixo RH", "%s - Iana Brito, só 0981 (%d linhas, R$ %.0f)"
         % (os.path.basename(arq_iana), len(iana_mp), iana_mp[TOT_MP].sum())),
        ("Mão de obra", "%s - Labour Cost do Pyplan (%d linhas, R$ %.0f); VAR. = Cost Center Variability"
         % (os.path.basename(arq_lc), len(lc_mp), lc_mp[TOT_MP].sum())),
        ("Sobreposição", "Caetano 15/10/2026: as três fontes SOMAM, mesmo na mesma CC x classe (aba Sobreposicao)"),
        ("Regra", "Caetano 15/10/2026: nada entra que não tenha vindo das bases recebidas - classe, valor, VAR. e ORIG. como "
         "vieram (só a base das áreas tem VAR. do R09, decisão de 14/10, porque veio toda F)"),
        ("Classe fora do plano 0981", "%d linhas, R$ %.0f, sem conta gestorial (ex.: Aprendizes da Iana em N14215M001/"
         "N14219M000) - corrigir com quem mandou; o Resultado não as enxerga"
         % ((novo["Conta Gestorial"] == "? sem conta gestorial").sum(),
            novo.loc[novo["Conta Gestorial"] == "? sem conta gestorial", TOT_MP].sum())),
        ("Draft por índice", "%s - não entra mais; o grupo RH dele (R$ %.0f) foi trocado por Labour Cost + Iana"
         % (os.path.basename(arq_draft), rh[TOT_MP].sum())),
        ("Pendente das áreas", "alocadas (e-mail de 09/10/2026)"),
        ("WH", "pelo CC no %s" % os.path.basename(dr.CENTROS)),
        ("VAR.", "a do R09/Actual %d no mesmo CC x classe; senão CC x conta gestorial; senão F/V do CC (a base vem toda F)" % dr.ANO_BASE),
        ("Perímetro", "Cabreúva e BAMP fora (áreas: %d linhas, R$ %.0f · RH: %d linhas, R$ %.0f)"
         % (len(fora), fora[TOT_MP].sum(), len(rh_fora), rh_fora[TOT_MP].sum())),
        ("Em aberto", "linhas IFRS16 repetidas (aba Duplicidades) - mantidas como vieram, pergunta feita à Rafaela"),
        ("Versão SAP", "J01 / ano %d - NÃO carregada, só o Excel" % dr.ANO_MP),
    ], columns=["Premissa", "Valor"])

    destino = os.path.join(dr.PASTA_SAIDA, "Draft MP%d Despesas - %s%s.xlsx"
                           % (dr.AA_MP, datetime.now().strftime("%Y%m%d"), SUFIXO))
    sub = "gerado em %s por scripts/tlm/mp27_base_areas.py" % datetime.now().strftime("%d/%m/%Y %H:%M")
    escrever_caderno(destino, [
        ("Premissas", "MP%d TLM - base única de custos" % dr.AA_MP, sub, premissas),
        ("Por_Fonte", "De onde vem cada custo e quem mandou", "R$", por_fonte),
        ("Sobreposicao", "CC x classe com valor em mais de uma fonte (somadas)", "R$ - decisão do Caetano 15/10/2026", sobre),
        ("Ponte_R09", "Ponte R09 %d -> MP%d - despesas sem RH" % (dr.ANO_BASE, dr.AA_MP),
         "R09 = intermediária (Actual + F09, A + P) - MP = base das áreas; RH e alocadas fora dos dois lados", ponte),
        ("R09_x_MP_WH", "R09 x MP%d por WH" % dr.AA_MP, "R$ - R09 no perímetro %d" % dr.ANO_MP, cmp_wh),
        ("R09_x_MP_Conta", "R09 x MP%d por conta gestorial" % dr.AA_MP, "R$ - R09 no perímetro %d" % dr.ANO_MP, cmp_conta),
        ("R09_x_MP_WH_Conta", "R09 x MP%d por WH x conta" % dr.AA_MP, "R$ - R09 no perímetro %d" % dr.ANO_MP, cmp_wh_conta),
        ("Resumo_WH", "Draft por índice x base das áreas por WH", "R$ - despesas sem RH; última coluna com RH", r_wh),
        ("Resumo_Conta", "Draft por índice x base das áreas por conta gestorial", "R$ - sem RH", r_conta),
        ("Origem_VAR", "De onde veio a VAR. de cada linha", "R$", r_var),
        ("Draft_MP%d" % dr.AA_MP, "Base única MP%d por CC x conta" % dr.AA_MP,
         "áreas (AREAS) + custo fixo RH da Iana (RH_FIXO) + Labour Cost do Pyplan (MDO) - colunas Fonte / Enviado por", novo),
        ("Base_Areas", "Base das áreas com WH e VAR.", os.path.basename(arq_areas),
         areas[["Linha Excel", "Área", "Comentários", "Origem VAR."] + COLS + dr.MESES + [TOT_MP]]),
        ("Carga_J01", "Carga MP%d - layout YKCCP003_LT" % dr.AA_MP,
         "versão J01 - as três fontes - NÃO carregada no SAP", carga),
        ("Duplicidades", "Linhas IFRS16 repetidas na base das áreas", "versão MP%d sem inflação - aguardando a Rafaela" % dr.AA_MP, dup),
    ])
    arq_carga = os.path.join(dr.PASTA_SAIDA, "Carga Custo MP%d - %s.xlsx" % (dr.AA_MP, datetime.now().strftime("%Y%m%d")))
    escrever_caderno(arq_carga, [
        ("Carga_J01", "Carga Custo MP%d TLM - layout YKCCP003_LT" % dr.AA_MP,
         "versão J01 / %d - áreas + custo fixo RH + MDO - NÃO carregada no SAP" % dr.ANO_MP, carga.drop(columns=ENVIADO)),
        ("Por_Fonte", "Carga por fonte", "R$", carga.groupby(ENVIADO, as_index=False)["TOTAL"].sum())])
    print("\nSaida:", destino)
    print("Carga:", arq_carga, "- %d linhas, R$ %.0f" % (len(carga), carga["TOTAL"].sum()))
    print(por_fonte.to_string(index=False))
    print("CC sem WH:", sem_wh)
    print("Despesas draft por indice : %14.0f" % sai[TOT_MP].sum())
    print("Despesas base das areas   : %14.0f  (%d linhas)" % (areas_mp[TOT_MP].sum(), len(areas_mp)))
    print("RH do draft (saiu)        : %14.0f" % rh[TOT_MP].sum())
    print("Custo fixo RH (Iana)      : %14.0f" % iana_mp[TOT_MP].sum())
    print("Mao de obra (Labour Cost) : %14.0f" % lc_mp[TOT_MP].sum())
    print("MP%d novo total           : %14.0f" % (dr.AA_MP, novo[TOT_MP].sum()))
    print(r_var.to_string(index=False))
    print(ponte[["Passo", "R$ mil"]].to_string(index=False))
    print(cmp_wh.drop(columns="Delta %").assign(**{c: lambda x, c=c: (x[c] / 1000).round(0) for c in cmp_wh.columns[1:-1]}).to_string(index=False))
    return destino


if __name__ == "__main__":
    import mp27_rh_iana_981 as recebidos
    print("[0] última versão recebida (Outlook)")
    recebidos.main()
    recebidos.baixar_labour_cost()
    montar()
