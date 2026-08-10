# SAP GUI Scripting habilitado e testado com sucesso

Confirmado que a conexão via SAP GUI Scripting funciona de ponta a ponta na máquina da Juliana:

- **Cliente (SAP GUI local):** "Ativar scripting" já estava marcado em Opções → Acessibilidade & Scripting → Scripting.
- **Servidor (TI/Basis da Pirelli):** também permite scripting — não foi preciso abrir chamado.
- **Teste:** `scripts/sap/test_conexao_sap.py`, usando `pywin32` (`win32com.client.GetObject("SAPGUI")`), conectou com sucesso na sessão logada (Sistema G20/Produção, mandante 210).

**Como conectar em scripts futuros:**
```python
import win32com.client
sap_gui_auto = win32com.client.GetObject("SAPGUI")
application = sap_gui_auto.GetScriptingEngine
session = application.Children(0).Children(0)  # primeira conexao, primeira sessao
```

**Pré-requisito para qualquer script SAP:** a Juliana precisa estar com o SAP GUI aberto e logada no sistema antes de rodar o script — o script não faz login sozinho, ele se conecta a uma sessão já aberta.

**Cuidado:** o script tem acesso de leitura E escrita à sessão real (produção). Qualquer automação que vá além de ler/exportar dados (ex: preencher campos, salvar, executar transações que alteram dados) precisa seguir a regra do `CLAUDE.md`/`REGRAS_RAPIDAS.md`: confirmar com a usuária antes.
