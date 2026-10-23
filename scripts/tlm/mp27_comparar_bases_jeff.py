# -*- coding: utf-8 -*-
"""mp27_comparar_bases_jeff.py -- o que variou entre duas versoes da base de custos do Jefferson (MP27 Inf).

USO: python scripts/tlm/mp27_comparar_bases_jeff.py <arq_antigo> <arq_novo> [<arq_referencia_sem_RH>]
     (nomes dentro de Arquivos Recebidos[/_Versoes_anteriores] ou caminhos completos)
SAIDA: data/processed/tlm/<data>/Comparativo base Jeff.xlsx (nunca sobrescreve: _v2, _v3...)
"""
import glob
import os
import sys
from datetime import datetime

import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import mp27_draft as dr                                                # noqa: E402

PASTA = os.path.join(dr.PASTA_MP, "Arquivos Recebidos")
CHAVE = ["WH", "CC", "Conta Gest", "Descrição"]


def achar(nome):
    if os.path.isfile(nome):
        return nome
    for sub in ("", "_Versoes_anteriores"):
        r = glob.glob(os.path.join(PASTA, sub, nome + "*")) or glob.glob(os.path.join(PASTA, sub, nome))
        if r:
            return sorted(r)[0]
    raise FileNotFoundError(nome)


def ler(arq):
    d = pd.read_excel(arq, sheet_name="Base", header=3)
    d.columns = [str(c).strip() for c in d.columns]
    d = d[d["Emp"].notna() & (d["Versão"] == "MP27 Inf")].copy()
    d["CC"] = d["CC"].astype(int)
    d["Conta Gest"] = d["Conta Gest"].astype(int)
    d[dr.MESES] = d[dr.MESES].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    d["Total Ano"] = d[dr.MESES].sum(axis=1)
    return d


def agrupar(d):
    return d.groupby(CHAVE, dropna=False)[dr.MESES + ["Total Ano"]].sum()


def comparar(a, b):
    ga, gb = agrupar(a), agrupar(b)
    j = ga.join(gb, how="outer", lsuffix="_ant", rsuffix="_novo").fillna(0.0)
    j["Anterior"], j["Novo"] = j["Total Ano_ant"], j["Total Ano_novo"]
    j["Variação"] = j["Novo"] - j["Anterior"]
    j["Var %"] = j["Variação"] / j["Anterior"].where(j["Anterior"] != 0)
    j["Tipo"] = "Alterada"
    j.loc[j["Anterior"] == 0, "Tipo"] = "Nova"
    j.loc[j["Novo"] == 0, "Tipo"] = "Removida"
    return j[j["Variação"].abs() > 1].reset_index().sort_values("Variação")


def main():
    ant, nov = achar(sys.argv[1]), achar(sys.argv[2])
    a, b = ler(ant), ler(nov)
    sem_rh = ler(achar(sys.argv[3])) if len(sys.argv) > 3 else None
    resumo = [("Anterior: " + os.path.basename(ant), a["Total Ano"].sum()),
              ("(-) Área RH (a base vira 'Sem RH')", -(a[a["Área"] == "RH"]["Total Ano"].sum() - b[b["Área"] == "RH"]["Total Ano"].sum())),
              ("= Anterior sem RH", a["Total Ano"].sum() - a[a["Área"] == "RH"]["Total Ano"].sum() + b[b["Área"] == "RH"]["Total Ano"].sum())]
    c_total = comparar(a, b)
    c_semrh = comparar(a[a["Área"] != "RH"], b[b["Área"] != "RH"])
    resumo.append(("Variações das demais áreas", c_semrh["Variação"].sum()))
    resumo.append(("Novo: " + os.path.basename(nov), b["Total Ano"].sum()))
    out = {"Resumo": pd.DataFrame(resumo, columns=["Passo", "R$"]),
           "Variação total com RH": c_total,
           "Variação sem RH": c_semrh}
    if sem_rh is not None:
        out["Novo vs Sem RH 16-10"] = comparar(sem_rh, b)
    pasta = os.path.join(BASE_DIR, "data", "processed", "tlm", datetime.now().strftime("%Y%m%d"))
    os.makedirs(pasta, exist_ok=True)
    dest, n = os.path.join(pasta, "Comparativo base Jeff.xlsx"), 1
    while os.path.exists(dest):
        n += 1
        dest = os.path.join(pasta, "Comparativo base Jeff_v%d.xlsx" % n)
    with pd.ExcelWriter(dest) as w:
        for nome, df in out.items():
            df.to_excel(w, sheet_name=nome[:31], index=False)
    print(dest)
    print(out["Resumo"].to_string(index=False))


if __name__ == "__main__":
    main()
