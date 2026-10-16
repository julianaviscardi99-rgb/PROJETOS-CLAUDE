# -*- coding: utf-8 -*-
"""Custo fixo do RH no MP27 da TLM (0981) — base da Iana Brito (RH corporativo, fio "Custo fixo - MP27").

A base da Iana é do RH inteiro (todas as empresas) e vem em versões pelo mesmo fio de e-mail. Pedido do Caetano
(15/10/2026): pegar sempre o ÚLTIMO anexo recebido, filtrar só a empresa 981, gravar e conferir formato e totais.
  1. Outlook: e-mail mais recente da Iana com .xlsx no fio "Custo fixo - MP27" → original salvo em
     Arquivos Recebidos\\_Original\\<AAAAMMDD_HHMM>_Iana Brito_<anexo> (não sobrescreve se já existe).
  2. Filtro COD. EMPRESA = 981 → `MP27 RH TLM 0981 - Iana Brito <dd-mm HHhMM>.xlsx` (padrão azul), abas
     Base_0981 · Conferencia (texto em campo de valor, TOTAL da Iana x soma dos meses, total por unidade,
     Cabreúva fora do perímetro, comparação com a versão anterior da 981).
  3. Arquivo anterior da 981 na raiz vai para Arquivos Recebidos\\_Versoes_anteriores.
Uso: python scripts/tlm/mp27_rh_iana_981.py
"""
import glob
import os
import shutil
import sys
from datetime import datetime

import openpyxl
import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE_DIR, "scripts", "antidumping"))

import mp27_draft as dr                                                # noqa: E402
from caderno import escrever_caderno                                   # noqa: E402

PASTA = os.path.join(dr.PASTA_MP, "Arquivos Recebidos")
PASTA_ORIGINAL = os.path.join(PASTA, "_Original")
PASTA_ANTERIORES = os.path.join(PASTA, "_Versoes_anteriores")
REMETENTE = "Brito Iana"
ASSUNTO = "custo fixo - mp27"
EMPRESA = 981
MESES = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
PREFIXO = "MP27 RH TLM 0981 - Iana Brito"
CC_CABREUVA = range(7500, 7576)     # perímetro MP27: Cabreúva sai (ontology/tlm_logic.json › master_plan.perimetro_mp27)


def baixar_ultimo():
    import win32com.client
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    def pastas(f):
        yield f
        for s in f.Folders:
            yield from pastas(s)

    melhor = None
    for f in pastas(ns.GetDefaultFolder(6)):
        try:
            itens = f.Items.Restrict("[ReceivedTime] >= '10/01/%d 00:00'" % datetime.now().year)
        except Exception:
            continue
        for m in itens:
            try:
                if REMETENTE not in (m.SenderName or "") or ASSUNTO not in (m.Subject or "").lower():
                    continue
                anexos = [m.Attachments.Item(i + 1) for i in range(m.Attachments.Count)
                          if m.Attachments.Item(i + 1).FileName.lower().endswith(".xlsx")]
                quando = datetime(*m.ReceivedTime.timetuple()[:6])
                if anexos and (melhor is None or quando > melhor[0]):
                    melhor = (quando, anexos[0])
            except Exception:
                pass
    if melhor is None:
        raise SystemExit("Nenhum e-mail da Iana com .xlsx no fio '%s'." % ASSUNTO)
    quando, anexo = melhor
    os.makedirs(PASTA_ORIGINAL, exist_ok=True)
    destino = os.path.join(PASTA_ORIGINAL, "%s_Iana Brito_%s" % (quando.strftime("%Y%m%d_%H%M"), anexo.FileName))
    if not os.path.exists(destino):
        anexo.SaveAsFile(destino)
    print("  último anexo da Iana: %s (%s)" % (anexo.FileName, quando.strftime("%d/%m %H:%M")))
    return quando, destino


def baixar_labour_cost():
    """Último 'Labour Cost TLM - MP27.xlsx' do Pyplan no Outlook; troca o da pasta se mudou (o antigo vai p/ _Versoes_anteriores)."""
    import filecmp
    import win32com.client
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    alvo = "Labour Cost TLM - MP%d.xlsx" % dr.AA_MP

    def pastas(f):
        yield f
        for s in f.Folders:
            yield from pastas(s)

    melhor = None
    for f in pastas(ns.GetDefaultFolder(6)):
        try:
            itens = f.Items.Restrict("[ReceivedTime] >= '10/01/%d 00:00'" % datetime.now().year)
        except Exception:
            continue
        for m in itens:
            try:
                for i in range(m.Attachments.Count):
                    a = m.Attachments.Item(i + 1)
                    quando = datetime(*m.ReceivedTime.timetuple()[:6])
                    if a.FileName == alvo and (melhor is None or quando > melhor[0]):
                        melhor = (quando, a)
            except Exception:
                pass
    if melhor is None:
        print("  Labour Cost: nenhum e-mail com %s - fica o da pasta" % alvo)
        return
    quando, anexo = melhor
    os.makedirs(PASTA_ORIGINAL, exist_ok=True)
    original = os.path.join(PASTA_ORIGINAL, "%s_Pyplan_%s" % (quando.strftime("%Y%m%d_%H%M"), alvo))
    if not os.path.exists(original):
        anexo.SaveAsFile(original)
    atual = os.path.join(PASTA, alvo)
    if os.path.exists(atual) and not filecmp.cmp(atual, original, shallow=False):
        os.makedirs(PASTA_ANTERIORES, exist_ok=True)
        shutil.move(atual, os.path.join(PASTA_ANTERIORES, "%s_%s" % (datetime.now().strftime("%Y%m%d_%H%M"), alvo)))
    if not os.path.exists(atual):
        shutil.copy2(original, atual)
    print("  último Labour Cost do Pyplan: %s (%s)" % (alvo, quando.strftime("%d/%m %H:%M")))


def ler_981(arq):
    wb_f = openpyxl.load_workbook(arq)
    wb_v = openpyxl.load_workbook(arq, data_only=True)
    ws_f, ws_v = wb_f.worksheets[0], wb_v.worksheets[0]
    cab = [c.value for c in ws_f[1]]
    i_emp, i_tot = cab.index("COD. EMPRESA"), cab.index("TOTAL")
    i_mes = [cab.index(m) for m in MESES]
    linhas, problemas = [], []
    for rf, rv in zip(ws_f.iter_rows(min_row=2), ws_v.iter_rows(min_row=2)):
        if str(rf[i_emp].value).strip().lstrip("0") != str(EMPRESA):
            continue
        reg = {c: rf[j].value for j, c in enumerate(cab) if c not in MESES + ["TOTAL", "CHAVE"]}
        reg["Linha Iana"] = rf[0].row
        for m, j in zip(MESES, i_mes):
            v = rf[j].value
            if v is not None and not isinstance(v, (int, float)):
                problemas.append(("texto em campo de valor", rf[j].coordinate, repr(v)))
                v = pd.to_numeric(str(v).replace(".", "").replace(",", "."), errors="coerce")
            reg[m] = float(v or 0)
        reg["TOTAL"] = sum(reg[m] for m in MESES)
        tot_iana = rv[i_tot].value
        if isinstance(tot_iana, (int, float)) and abs(tot_iana - reg["TOTAL"]) > 0.01:
            problemas.append(("TOTAL da Iana diferente da soma dos meses", rf[i_tot].coordinate,
                              "%.2f x %.2f" % (tot_iana, reg["TOTAL"])))
        reg["TOTAL Iana"] = tot_iana
        linhas.append(reg)
    df = pd.DataFrame(linhas)
    for c in ["COD. EMPRESA", "ANO", "CC"]:
        df[c] = df[c].astype(str)
    return df, problemas


def main():
    quando, arq = baixar_ultimo()
    df, problemas = ler_981(arq)
    ordem = [c for c in df.columns if c not in MESES + ["TOTAL", "TOTAL Iana"]]
    ordem = ordem[:ordem.index("ORIG.") + 1] + MESES + ["TOTAL", "TOTAL Iana"] + ordem[ordem.index("ORIG.") + 1:]
    df = df[ordem]

    anterior = sorted(glob.glob(os.path.join(PASTA, PREFIXO + "*.xlsx")))
    nome = "%s %s.xlsx" % (PREFIXO, quando.strftime("%d-%m %Hh%M"))
    destino = os.path.join(PASTA, nome)

    conf = [("Arquivo da Iana (original)", os.path.basename(arq)),
            ("Recebido em", quando.strftime("%d/%m/%Y %H:%M")),
            ("Linhas da 981", len(df)),
            ("Total 981 (soma dos meses)", round(df["TOTAL"].sum(), 2)),
            ("Total 981 (coluna TOTAL da Iana)", round(pd.to_numeric(df["TOTAL Iana"], errors="coerce").sum(), 2)),
            ("Células de valor em texto", sum(1 for p in problemas if p[0].startswith("texto"))),
            ("Linhas com TOTAL diferente da soma dos meses", sum(1 for p in problemas if p[0].startswith("TOTAL")))]
    cab = df[df["CC"].astype(int).isin(CC_CABREUVA)]
    conf.append(("Cabreúva (CC 7500-7575) — fora do perímetro MP27", "%d linhas · %.2f" % (len(cab), cab["TOTAL"].sum())))
    sem_desc = df[df.iloc[:, df.columns.get_loc("CLASSE DE CUSTO") + 1].isna()]
    conf.append(("Linhas sem descrição da classe", len(sem_desc)))
    comparar = [a for a in anterior if os.path.basename(a) != nome] + glob.glob(os.path.join(PASTA, "MP27_env*.xlsx"))
    guardados = [a for a in glob.glob(os.path.join(PASTA_ANTERIORES, "*.xlsx"))
                 if os.path.basename(a) not in map(os.path.basename, comparar)]
    for arq_ant in comparar + guardados:
        ant = pd.read_excel(arq_ant, header=3 if os.path.basename(arq_ant).startswith(PREFIXO) else 0)
        ant = ant[ant["COD. EMPRESA"].astype(str).str.lstrip("0") == str(EMPRESA)]
        tot_ant = ant[MESES].apply(pd.to_numeric, errors="coerce").fillna(0).values.sum()
        conf.append(("Versão anterior: %s" % os.path.basename(arq_ant),
                     "%d linhas · %.2f · Δ %.2f" % (len(ant), tot_ant, df["TOTAL"].sum() - tot_ant)))
    df_conf = pd.DataFrame(conf, columns=["Conferência", "Resultado"])
    df_conf["Resultado"] = df_conf["Resultado"].astype(str)
    por_un = df.groupby("UNIDADE", as_index=False)[MESES + ["TOTAL"]].sum()
    por_un = pd.concat([por_un, pd.DataFrame([{"UNIDADE": "TOTAL", **por_un[MESES + ["TOTAL"]].sum().to_dict()}])])
    df_prob = pd.DataFrame(problemas or [("nenhum", "", "")], columns=["Problema", "Célula", "Valor"])

    sub = "Base da Iana Brito (RH) recebida em %s — só COD. EMPRESA 981; valores em R$" % quando.strftime("%d/%m/%Y %H:%M")
    escrever_caderno(destino, [
        ("Base_0981", "MP27 — Custo fixo RH — TLM (0981)", sub, df),
        ("Conferencia", "Conferência da base da Iana — 0981", sub, df_conf),
        ("Total_por_unidade", "Total por unidade — 0981", sub, por_un),
        ("Problemas", "Células com problema na base da Iana (0981)", sub, df_prob)])
    os.makedirs(PASTA_ANTERIORES, exist_ok=True)
    for a in [a for a in anterior if os.path.basename(a) != nome] + glob.glob(os.path.join(PASTA, "MP27_env*.xlsx")):
        shutil.move(a, os.path.join(PASTA_ANTERIORES, os.path.basename(a)))
    print("  gravado: %s" % destino)
    for c, r in conf:
        print("   - %s: %s" % (c, r))


if __name__ == "__main__":
    main()
