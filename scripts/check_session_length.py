#!/usr/bin/env python3
"""Conta acoes (tool calls) da sessao atual do Claude Code e dispara o backup/archive ao atingir o limite."""
import json
import subprocess
import sys
from pathlib import Path

LIMITE = 45

PROJECT_ROOT = Path(__file__).resolve().parent.parent
STATE_DIR = PROJECT_ROOT / ".claude" / "session_state"
TRANSITION_SCRIPT = PROJECT_ROOT / "scripts" / "session_transition.py"


def main():
    try:
        payload = json.load(sys.stdin)
    except (json.JSONDecodeError, ValueError):
        payload = {}

    session_id = payload.get("session_id") or "default"

    STATE_DIR.mkdir(parents=True, exist_ok=True)
    state_file = STATE_DIR / f"{session_id}.count"

    count = 0
    if state_file.exists():
        try:
            count = int(state_file.read_text().strip() or "0")
        except ValueError:
            count = 0

    count += 1

    if count >= LIMITE:
        state_file.write_text("0")
        resultado = subprocess.run(
            [sys.executable, str(TRANSITION_SCRIPT)],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
        )
        resumo = resultado.stdout.strip() or resultado.stderr.strip() or "sem saida"
        output = {
            "systemMessage": (
                f"Sessao longa detectada ({LIMITE} acoes). Backup automatico executado: {resumo}"
            ),
            "hookSpecificOutput": {
                "hookEventName": "PostToolUse",
                "additionalContext": (
                    "ALERTA DE SESSAO LONGA: esta sessao atingiu o limite de "
                    f"{LIMITE} acoes. O backup/archive automatico ja rodou (session_transition.py). "
                    "Regra do CLAUDE.md: atualize memory/BRIEFING.md com o progresso desta sessao "
                    "agora, e informe a usuaria que ela pode encerrar esta janela e abrir uma nova "
                    "para continuar (o contexto ja esta salvo)."
                ),
            },
        }
        print(json.dumps(output))
    else:
        state_file.write_text(str(count))


if __name__ == "__main__":
    main()
