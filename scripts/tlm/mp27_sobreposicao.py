# -*- coding: utf-8 -*-
"""Possiveis duplicidades entre as bases recebidas do MP27 da TLM (0981).

OBJETIVO: perguntar a quem mandou se o mesmo custo veio em duas bases (areas x RH custo fixo x Labour Cost).
OPERACAO: le a aba Draft_MP27 da ultima base unica (mp27_base_areas.py) e lista cada CC x classe com valor em mais de
          uma fonte - resumo (uma linha por CC x classe, valor de cada fonte) e detalhe (cada linha como veio, com o
          comentario de quem mandou). Grava 'Sobreposicao fontes MP27 - <data>.xlsx' em TLM_Automation.
ALINHAMENTO: hoje as tres fontes SOMAM (decisao do Caetano 15/10/2026); o arquivo e o anexo do e-mail que confirma
          com as areas e o RH se a soma esta certa. Nada e alterado na base.
Uso: python scripts/tlm/mp27_sobreposicao.py
"""
import glob
import os
import sys
from datetime import datetime

import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE_DIR, "scripts", "antidumping"))

import mp27_draft as dr                                                # noqa: E402
from caderno import escrever_caderno                                   # noqa: E402

TOT = "Total MP%d" % dr.AA_MP
ENVIADO = "Enviado por"
CURTO = {"Rafaela": "Base das áreas (Rafaela)", "Iana": "Custo fixo RH (Iana)", "Pyplan": "Labour Cost MDO (Pyplan)"}


def _curto(enviado):
    return next((v for k, v in CURTO.items() if k in enviado), enviado)


def main():
    arq = sorted(glob.glob(os.path.join(dr.PASTA_SAIDA, "Draft MP%d Despesas - * - base unica.xlsx" % dr.AA_MP)))[-1]
    d = pd.read_excel(arq, sheet_name="Draft_MP%d" % dr.AA_MP, header=3,
                      dtype={"Centro de Custo": str, "Conta Fiscal": str, "Conta Gestorial": str})
    d = d[d["Conta Fiscal"].notna()].copy()
    d["Base"] = d[ENVIADO].map(_curto)
    k = ["Centro de Custo", "Conta Fiscal"]
    n = d.groupby(k)["Base"].nunique()
    chaves = n[n > 1].index
    det = d.set_index(k).loc[chaves].reset_index()

    res = det.pivot_table(index=k, columns="Base", values=TOT, aggfunc="sum", fill_value=0.0)
    res = res.reindex(columns=[c for c in CURTO.values() if c in res.columns])
    res["Total somado no MP27"] = res.sum(axis=1)
    info = det.drop_duplicates(k).set_index(k)[["WH", "CC Desc", "Classe Desc"]]
    res = info.join(res).reset_index().sort_values("Total somado no MP27", key=abs, ascending=False)
    res.insert(0, "Item", range(1, len(res) + 1))
    item = dict(zip(zip(res["Centro de Custo"], res["Conta Fiscal"]), res["Item"]))

    det["Item"] = [item[x] for x in zip(det["Centro de Custo"], det["Conta Fiscal"])]
    det = det.sort_values(["Item", "Base"])
    det = det[["Item", "Base", "WH", "Centro de Custo", "CC Desc", "Conta Fiscal", "Classe Desc", "Observação"]
              + dr.MESES + [TOT]]

    tot = res[[c for c in res.columns if c in CURTO.values()] + ["Total somado no MP27"]].sum()
    resumo_fontes = pd.DataFrame({"Base": tot.index, "R$ nas CC x classe repetidas": tot.values})

    destino = os.path.join(dr.PASTA_SAIDA, "Sobreposicao fontes MP%d - %s.xlsx" % (dr.AA_MP, datetime.now().strftime("%Y%m%d")))
    sub = "CC x classe com valor em mais de uma base recebida - hoje somadas no MP%d; fonte: %s" % (dr.AA_MP, os.path.basename(arq))
    escrever_caderno(destino, [
        ("Resumo", "MP%d TLM - possíveis duplicidades entre bases recebidas" % dr.AA_MP, sub, res),
        ("Detalhe", "Linhas como vieram de cada base", "R$ - a coluna Observação traz o comentário de quem mandou", det),
        ("Por_base", "Total por base nas CC x classe repetidas", "R$", resumo_fontes)])
    print("Saida:", destino)
    print("CC x classe repetidas: %d | linhas: %d" % (len(res), len(det)))
    print(resumo_fontes.to_string(index=False))
    print(res.drop(columns=["CC Desc", "Classe Desc"]).round(0).to_string(index=False))
    return destino


if __name__ == "__main__":
    main()
