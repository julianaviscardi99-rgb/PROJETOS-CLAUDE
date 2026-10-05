"""Backup das conversas do Claude Code (transcritos .jsonl) + versão legível em .md.

Roda a cada 3 min via Agendador de Tarefas. Só copia/converte o que mudou.
Saída: data/processed/conversas/ (fora do Git, ver .gitignore).
Uso: python scripts/backup_conversas.py
"""
import json
import shutil
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
# Claude Code guarda os transcritos em ~/.claude/projects/<caminho-do-projeto-com-hifens>/
SLUG = str(RAIZ).replace(":", "-").replace("\\", "-").replace("/", "-").replace(" ", "-")
ORIGEM = Path.home() / ".claude" / "projects" / SLUG
DESTINO = RAIZ / "data" / "processed" / "conversas"


def texto_da_mensagem(msg):
    c = msg.get("content")
    if isinstance(c, str):
        return c
    partes = []
    for p in c or []:
        if p.get("type") == "text":
            partes.append(p["text"])
        elif p.get("type") == "tool_use":
            partes.append(f"`[ferramenta: {p.get('name')}]`")
    return "\n".join(partes)


def para_markdown(jsonl, md):
    linhas = []
    for bruta in jsonl.read_text(encoding="utf-8", errors="replace").splitlines():
        try:
            o = json.loads(bruta)
        except ValueError:
            continue
        if o.get("type") not in ("user", "assistant"):
            continue
        t = texto_da_mensagem(o.get("message", {})).strip()
        if t:
            quem = "USUÁRIA" if o["type"] == "user" else "CLAUDE"
            linhas.append(f"### [{o.get('timestamp', '')}] {quem}\n\n{t}\n")
    md.write_text("\n".join(linhas), encoding="utf-8")


def main():
    if not ORIGEM.exists():
        return
    DESTINO.mkdir(parents=True, exist_ok=True)
    for f in ORIGEM.glob("*.jsonl"):
        copia = DESTINO / f.name
        if copia.exists() and copia.stat().st_mtime >= f.stat().st_mtime \
                and copia.stat().st_size == f.stat().st_size:
            continue
        shutil.copy2(f, copia)
        para_markdown(f, copia.with_suffix(".md"))


if __name__ == "__main__":
    main()
