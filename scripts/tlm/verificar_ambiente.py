# -*- coding: utf-8 -*-
"""Confere se a máquina tem o necessário para rodar o MP27 TLM. Uso: python verificar_ambiente.py"""
import importlib
import os
import sys

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
REDE = r"\\SRV054-01BR.group.pirelli.com\DAC_TLM\MP27\1ª Versão MP27 vs R09"

ok = True
print(f"Python {sys.version.split()[0]}")
for mod in ("pandas", "openpyxl", "win32com.client"):
    try:
        importlib.import_module(mod)
        print(f"  [ok] {mod}")
    except ImportError:
        ok = False
        print(f"  [FALTA] {mod} — rode: pip install -r requirements.txt")

for sub in ("Arquivos Recebidos", "TLM_Automation", "Volume"):
    p = os.path.join(REDE, sub)
    existe = os.path.isdir(p)
    ok &= existe
    print(f"  [{'ok' if existe else 'SEM ACESSO'}] {p}")

sys.path.insert(0, os.path.join(BASE_DIR, "scripts", "tlm"))
sys.path.insert(0, os.path.join(BASE_DIR, "scripts", "antidumping"))
try:
    import caderno  # noqa: F401
    import mp27_draft  # noqa: F401
    print("  [ok] scripts do projeto importam")
except Exception as e:
    ok = False
    print(f"  [ERRO] scripts do projeto: {e}")

try:
    import win32com.client
    xl = win32com.client.DispatchEx("Excel.Application")
    print(f"  [ok] Excel {xl.Version}")
    xl.Quit()
except Exception as e:
    ok = False
    print(f"  [ERRO] Excel via win32com: {e}")

print("\nTudo certo." if ok else "\nCorrija os itens acima antes de rodar os scripts.")
