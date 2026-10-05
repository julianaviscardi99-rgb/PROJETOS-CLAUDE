"""Lista a caixa de entrada do Outlook (somente leitura) e destaca o que parece urgente.

Uso:
    python scripts/ler_caixa_entrada.py [--dias 3] [--max 200]

Não envia, move, apaga nem marca nada como lido. Salva um CSV (sem corpo do e-mail)
em data/processed/.
"""
import argparse
import csv
from datetime import datetime, timedelta
from pathlib import Path

import win32com.client

RAIZ = Path(__file__).resolve().parent.parent
SAIDA = RAIZ / "data" / "processed"

PALAVRAS_URGENTES = [
    "urgente", "urgência", "urgencia", "asap", "prazo", "hoje", "imediato",
    "pendente", "atraso", "atrasado", "vencimento", "vence", "cobrança", "cobranca",
    "fechamento", "deadline", "escalation", "importante", "ação necessária",
    "acao necessaria", "action required", "reminder", "lembrete",
]
OL_IMPORTANCIA_ALTA = 2
OL_PASTA_CAIXA_ENTRADA = 6


def pontuar(item):
    motivos = []
    if getattr(item, "Importance", 1) == OL_IMPORTANCIA_ALTA:
        motivos.append("importância alta")
    if getattr(item, "FlagStatus", 0) == 2:
        motivos.append("sinalizado")
    if getattr(item, "UnRead", False):
        motivos.append("não lido")
    texto = f"{item.Subject or ''} {(item.Body or '')[:500]}".lower()
    achadas = [p for p in PALAVRAS_URGENTES if p in texto]
    if achadas:
        motivos.append("palavras: " + ", ".join(achadas[:3]))
    peso = (
        3 * ("importância alta" in motivos)
        + 2 * ("sinalizado" in motivos)
        + 1 * ("não lido" in motivos)
        + 2 * bool(achadas)
    )
    return peso, motivos


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dias", type=int, default=3)
    ap.add_argument("--max", type=int, default=200)
    ap.add_argument("--nao-lidos", action="store_true", help="só não lidos (use --dias 0 para sem limite de data)")
    args = ap.parse_args()

    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    itens = ns.GetDefaultFolder(OL_PASTA_CAIXA_ENTRADA).Items
    itens.Sort("[ReceivedTime]", True)
    desde = (datetime.now() - timedelta(days=args.dias)).strftime("%m/%d/%Y %H:%M")
    if args.dias > 0:
        itens = itens.Restrict(f"[ReceivedTime] >= '{desde}'")
    if args.nao_lidos:
        itens = itens.Restrict("[UnRead] = True")

    linhas = []
    for i, item in enumerate(itens):
        if i >= args.max:
            break
        try:
            if item.Class != 43:  # só e-mails (MailItem)
                continue
            peso, motivos = pontuar(item)
            linhas.append({
                "recebido": item.ReceivedTime.strftime("%Y-%m-%d %H:%M"),
                "remetente": item.SenderName,
                "assunto": item.Subject,
                "lido": "não" if item.UnRead else "sim",
                "peso": peso,
                "motivos": "; ".join(motivos),
            })
        except Exception:
            continue

    linhas.sort(key=lambda r: (-r["peso"], r["recebido"]), reverse=False)
    SAIDA.mkdir(parents=True, exist_ok=True)
    arq = SAIDA / f"caixa_entrada_{datetime.now():%Y%m%d_%H%M}.csv"
    with open(arq, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(linhas[0]) if linhas else ["vazio"])
        w.writeheader()
        w.writerows(linhas)

    print(f"{len(linhas)} e-mails dos últimos {args.dias} dia(s). CSV: {arq}\n")
    print("=== POSSIVELMENTE URGENTES (peso >= 3) ===")
    for r in [r for r in linhas if r["peso"] >= 3]:
        print(f"[{r['peso']}] {r['recebido']} | {r['remetente']} | {r['assunto']} | {r['motivos']}")
    print("\n=== DEMAIS NÃO LIDOS ===")
    for r in [r for r in linhas if r["peso"] < 3 and r["lido"] == "não"]:
        print(f"{r['recebido']} | {r['remetente']} | {r['assunto']}")


if __name__ == "__main__":
    main()
