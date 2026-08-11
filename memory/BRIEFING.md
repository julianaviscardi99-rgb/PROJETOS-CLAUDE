# BRIEFING — Documento Vivo da Sessão
> Atualizado por Claude em tempo real. Lido no início de cada sessão.
> Manter apenas as últimas 2 sessões inline — sessões mais antigas vão para long_term/.

---
## Resumo do dia 2026-08-11 — Fitted Units Despesas (popup SAP + estrutura de pastas + regra da Base Intermediária)
- **Popup "Definir área contab.custos" do SAP:** usuária relatou que, ao sair/voltar a entrar no SAP, um popup pede a área contábil (0580) antes de liberar qualquer transação. Primeira tentativa de fechar automaticamente adivinhou o campo errado e causou erro em cascata no SAP. Criado `scripts/sap/diagnosticar_popup.py` (+ atalho `DIAGNOSTICAR POPUP.bat`) pra inspecionar o popup ao vivo; descobriu-se que o campo fica dentro de um subscreen (`usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE`), não direto em `usr`. Corrigido com busca recursiva por `GuiCTextField` em `atualizar_ksb1_gui.py` e `extrair_ksb1.py`. **Testado pela usuária e funcionando.** Detalhe completo em `memory/errors/2026-08-11_popup_area_contabil_ao_reentrar_sap.md`.
- **Regra de negócio confirmada — passo 3 do processo (montar a base intermediária):**
  - Depois de extrair a KSB1 (Gestoriais + Sem Agrupamento) e rodar o Check: se o total do Gestoriais bate com o Sem Agrupamento, usa-se o Gestoriais; se não bater (porque existem contas fora do agrupamento gestorial), usa-se o Sem Agrupamento, mas ainda excluindo as contas do Check 1 (fixas + qualquer conta iniciada em "B", que são bens de investimento e nunca entram como despesa).
  - Isso acontece 2x/mês (Flash e Actual). Forecast existe mas está **fora do escopo por enquanto**.
  - Cada fechamento parte do arquivo do Actual do mês anterior (que já acumula o efetivo histórico) — ex: fechando julho, parte-se do arquivo "June Actual" e inserem-se as linhas de julho nele, sem apagar os meses anteriores.
  - Estrutura descoberta nos arquivos de referência (`.../2026/<mês>/<ciclo>/`): `KSB1 <Mês> <Ciclo> <Ano>.xlsx` → aba `BASE_KSB1` (tabela acumulada com todos os meses, colunas derivadas como Gestorial/MF/DG-MO são fórmulas que a usuária arrasta) → aba `Pivot_Inter.` alimenta `Base Intermediária Fitted <Mês> <Ciclo> <Ano>.xlsx` (matriz Conta Gestorial × mês, com Total Ano).
  - **Ainda não implementado em script** — só a regra de negócio foi confirmada e documentada. Automatizar esse passo é o próximo trabalho pesado.
- **Reorganização de pastas na rede** (`\\FSS024-01BR.group.pirelli.com\GFU_DAC\Custos Fitted Units\Resultados Fitted\2026\`): antes cada ciclo ficava solto na raiz (`01_Jan_Actual`, `01_Jan_Flash`...); agora cada mês tem sua pasta própria (`01 - Jan` a `12 - Dec`, mesmo padrão já usado em `00.Extração Base KSB1`) com Actual/Flash/Forecast dentro. Jan-Jul: pastas movidas sem alterar nenhum arquivo (uma pasta de julho — Forecast — ficou presa por um arquivo aberto no Excel, resolvido depois que a usuária fechou o arquivo). Ago-Dez: subpastas Actual/Flash/Forecast já criadas vazias, prontas pro uso.
- Pendente registrar formalmente em `ontology/fitted_units.json` a nova estrutura de pastas e a regra da base intermediária (só ficou no BRIEFING por enquanto — fazer isso na próxima sessão antes de começar a programar o passo 3).
- **Ajuste final da reorganização de pastas (mesma sessão, depois do alerta de sessão longa):** as subpastas de Ago-Dez tinham ficado só como `Actual`/`Flash`/`Forecast` (sem prefixo). Usuária pediu para seguir o mesmo padrão dos meses já usados (`MM_Mon_Ciclo`, ex: `07_Jul_Actual`). Renomeadas para `08_Aug_Actual/Flash/Forecast`, `09_Sep_...`, `10_Oct_...`, `11_Nov_...`, `12_Dec_...` — dentro das pastas `08 - Aug` a `12 - Dec`. Estrutura de pastas de 2026 agora está 100% padronizada e completa.

---
## Resumo do dia 2026-08-10 — Fitted Units Despesas (KSB1)
- **Instrução da usuária para as próximas sessões:** focar só em Fitted Units por enquanto. Circuito Panamericano fica pausado até ela pedir de novo — não avançar nele sem ela pedir.
- **O que foi entregue e validado hoje** (sub-projeto batizado por ela de "Fitted Units Despesas" — ver `memory/PROJECT_MAP.md`):
  1. `scripts/sap/atualizar_ksb1_gui.py` — GUI que extrai a KSB1 da Fitted Units (Gestoriais + Sem Agrupamento) direto do SAP via GUI Scripting e salva na área de rede. Nunca sobrescreve arquivo existente (versiona `_v2`, `_v3`...). Mês/ano padrão: ano atual dinâmico, mês = mês anterior ao atual. Navegação entre as duas extrações usa o botão "Voltar" da toolbar do SAP (mais robusto que F3). Sem popups manuais de confirmação (as notificações de segurança de scripting já estão desativadas no SAP GUI da usuária).
  2. `scripts/sap/check_agrupamentos_ksb1.py` — conferência automática: Check 1 lista as contas contábeis sempre ignoradas (6 contas fixas, mesmo com valor 0, mais qualquer conta iniciada em "B" encontrada no mês); Check 2 lista contas do Sem Agrupamento sem vínculo no agrupamento gestorial (pra mandar pra controladoria central); Resumo compara o total das duas bases. Tudo numa aba só, arquivo `Check de agrupamentos - MM.YYYY.xlsx` salvo na pasta do mês.
  3. As duas funções ficam na mesma janela/app, acessada por um único atalho na rede: `ATUALIZAR KSB1.lnk` (ícone de pneu Pirelli), com os botões "Extrair KSB1" e "Gerar Check de Agrupamentos".
  4. Regras de negócio documentadas em `ontology/fitted_units.json` (→ `classificacao_despesas.check_de_agrupamentos`); decisões e motivos em `memory/DECISOES.md`; novas regras gerais em `memory/REGRAS_RAPIDAS.md` (#11 respostas em português, #12 nunca sobrescrever arquivo — versionar).
  5. `requirements.txt` atualizado com `openpyxl` (necessário pro check).
- Histórico detalhado passo a passo de hoje (todos os bugs e correções, em ordem cronológica): `memory/long_term/2026-08-10_*_briefing_snapshot.md`.
- Tudo commitado e sincronizado com o GitHub (branch `main`, sem pendência de commit/push).

---
## Próximos passos (retomar — só Fitted Units)
- **Prioridade:** registrar em `ontology/fitted_units.json` a regra de negócio da base intermediária e a nova estrutura de pastas (ver resumo de 2026-08-11 acima) — ainda não formalizado lá, só no BRIEFING.
- Depois disso, começar a automatizar o passo 3 do processo (`ontology/fitted_units.json` → `processo_recorrente`): montar/atualizar a `BASE_KSB1` acumulada com as linhas do mês fechado (Gestoriais quando bate, Sem Agrupamento filtrado quando não bate), a partir do arquivo do Actual do mês anterior copiado.
- Depois: passos 4-5 do processo — rateio dos custos da Gerência (GER) pras demais unidades, depois carregar no arquivo de P&L.
- Usuária ainda precisa rodar "Gerar Check de Agrupamentos" com um mês real (Gestoriais + Sem Agrupamento do mesmo mês) pela primeira vez desde o último ajuste (aba única + Check 1 sempre listando as 6 contas fixas) e confirmar se bate com o esperado — se isso já rolou informalmente durante as extrações desta sessão, confirmar com a usuária antes de assumir que ainda está pendente.
- `memory/PROJECT_MAP.md`: Original Equipment ainda não detalhado (fora do escopo por ora, já que o foco agora é só Fitted Units).
- Circuito Panamericano: não mexer até a usuária pedir explicitamente.
- Forecast: fora do escopo por enquanto (confirmado pela usuária em 2026-08-11), mas as pastas de rede já foram deixadas prontas pra quando for retomado.

---
## Contexto permanente do projeto
- Esta pasta (`C:\Users\silveju001\Projetos Claude`) está estruturada seguindo o "Guia de Onboarding — Como Trabalhar com o Claude de Forma Profissional" (maio 2026, baseado no projeto Cockpit Ind — Pirelli Planning & Control).
- Objetivo real deste projeto: automatizar controladoria (Fitted Units e Circuito Panamericano) hoje feita em Excel — resultado, despesas, faturamento, EBIT, P&L mensal. Detalhes completos em `CLAUDE.md`. **Foco atual (a partir de 2026-08-10): só Fitted Units.**
- Repositório Git já configurado com backup remoto no GitHub; backup automático diário às 18h (Agendador de Tarefas do Windows) além do backup por sessão longa (a cada 45 ações).
- GUI compartilhável da KSB1 (`scripts/sap/atualizar_ksb1_gui.py`, atalho único `ATUALIZAR KSB1.lnk` na rede) e o Check de Agrupamentos (`scripts/sap/check_agrupamentos_ksb1.py`) já em uso — ver `memory/DECISOES.md` para o histórico completo de decisões sobre eles.
