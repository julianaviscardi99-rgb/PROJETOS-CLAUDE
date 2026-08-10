#!/usr/bin/env python3
"""Arquiva um snapshot do BRIEFING.md em memory/long_term/ e sincroniza o projeto com o GitHub (commit + push)."""
import subprocess
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
BRIEFING = PROJECT_ROOT / "memory" / "BRIEFING.md"
LONG_TERM = PROJECT_ROOT / "memory" / "long_term"


def archive_briefing():
    if not BRIEFING.exists():
        return None
    LONG_TERM.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    destino = LONG_TERM / f"{timestamp}_briefing_snapshot.md"
    destino.write_text(BRIEFING.read_text(encoding="utf-8"), encoding="utf-8")
    return destino


def run_git(*args):
    return subprocess.run(
        ["git", *args], cwd=PROJECT_ROOT, capture_output=True, text=True
    )


def main():
    destino = archive_briefing()

    run_git("add", "-A")
    diff = run_git("diff", "--cached", "--quiet")
    if diff.returncode == 0:
        print("Nada novo para commitar.")
        return

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    commit = run_git("commit", "-m", f"Auto session transition - {timestamp}")
    if commit.returncode != 0:
        print(f"Falha no commit: {commit.stderr.strip()}")
        return

    push = run_git("push")
    if push.returncode != 0:
        print(f"Commit feito, mas push falhou: {push.stderr.strip()}")
        return

    msg = "Nenhum BRIEFING.md encontrado para arquivar" if destino is None else f"Snapshot arquivado em {destino.name}"
    print(f"{msg}. Commit e push concluidos.")


if __name__ == "__main__":
    main()
