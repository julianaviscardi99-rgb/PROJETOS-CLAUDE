# 2026-10-16 — TLM: inventário do que o antecessor fez com o Claude (handoff)

Juliana assumiu a controladoria da **TLM** (empresa SAP **0981**; unidades Barueri, Campinas, Feira de Santana, São Bernardo; Cabreúva sai em 2027).
Antecessor: **Jose Caetano Farillo** (aparece como autor dos arquivos e nas premissas "decisão do Caetano").

## Onde está
- Rede: `\SRV054-01BR.group.pirelli.com\DAC_TLM` (pastas por ano, MP21..MP27, SP).
- **O código NÃO está na rede.** Os arquivos citam `scripts/tlm/mp27_base_areas.py` e `scripts/tlm/mp27_faturamento_base.py` → estão no PC/repositório do Caetano. Pedir a pasta de scripts + memória/CLAUDE.md dele.

## Frente 1 — Faturamento mensal (ago–set/2026) — `2026\TLM_Automation\`
- `Faturamento_TLM_<Mes>.xlsx` (+ `_Envio`): Painel (mês, premissas de impostos), Impostos por Localidade, Mapa de Abas, Base Flash, Drivers, Faturamento do Mês (líquido → bruto), Texto das NFs, **Conferência** antes de mandar emitir, uma aba por bloco (BA-Car, BA-Moto, BAMP, Campneus, Hevea, Fitted, BAQA...).
- Feito para Julho (conferido contra 17 NFs), Agosto e Setembro (30/09). **Outubro ainda não.**
- Alertas abertos que ele deixou: aba FSCheck com título errado; linhas deslocadas no "Resumo Fat. Logística"; fator ISS da Fitted oscilando 2%/5% (definir com fiscal); CPFab fora do padrão; aba "Faturamento Industrial" do fiscal com #REF.
- A seção "variações" da Conferência de Setembro ainda traz texto de julho x junho (texto não atualizado).

## Frente 2 — MP27 (1ª versão MP27 vs R09) — `MP27\1ª Versão MP27 vs R09\TLM_Automation\` e `Volume\`
- Despesas: `Draft MP27 Despesas - 20261015 - base unica.xlsx` = 3 fontes somadas (áreas/Rafaela Moreira 34,2 mi; RH fixo/Iana Brito 2,9 mi; Labour Cost Pyplan 54,4 mi; total 91,6 mi). Ponte R09→MP27, comparativos por WH/conta, `Carga_J01` no layout YKCCP003. **Versão J01 NÃO carregada no SAP.**
- `Base Total MP27.xlsx` (aba Intermediaria) alimenta `Resultado TLM - MP27 vs R09.xlsm` → `Resumo Resultado ... .xlsx`.
- Faturamento MP27: `Volume\Faturamento MP27 - base R09 - 20261015.xlsx` (tarifa a custo, método MP26; toneladas são SIMULAÇÃO até a área técnica mandar o Volume Movimentado).
- Pendências abertas: 19 linhas (R$ 695.940) com classe fora do plano 0981; linhas IFRS16 repetidas (pergunta à Rafaela); sobreposição de fontes (R$ 6,8 mi na mesma CC x classe — ele decidiu somar); custo variável MP27 de Campinas e Feira = 0 no arquivo de faturamento; toneladas MP27 vazias.

## Fechamento mensal 2026 (Flash/Actual/R10)
- Arquivos `Resultado TLM - ... .xlsm`, Base KSB1, Bases intermediárias: **sem marca de script** → aparentemente feitos à mão no Excel. A automação foi só faturamento + MP27.

## Diagnóstico MP27 aprofundado (16/10, a pedido dela: "começar o MP")
- Fluxo de referência (MP26): Bases (áreas, Labour Cost, Fixo RH, Ecolabel) → Carga Custo → carga SAP (YKCCP003/004) → Intermediária Fix/Var → Base Total → Resultado .xlsm → Resumo; mais Volume Movimentado (técnica), HC/Hours, Layoff, Premissas Campneus/Hevea/MP. MP27: pastas `Bases 2027`, `Intermediaria` e `Mão de Obra` vazias (o Caetano pulou o SAP: a Base Total vem direto do Excel).
- O Resultado puxa cada linha por SUMIFS em `Intermediaria[Desc.Resultado]` = rótulo da coluna A das abas de WH (71 rótulos).
- **Base Total MP27 (15/10 23:02) = R$ 91,6 mi, mas R$ 45,5 mi NÃO casam com nenhum rótulo:** Pyplan R$ 44,8 mi (Desc.Resultado com nome cru do Pyplan: "SALARIOS", "ENC S/ SALRIOS HORISTA", "ASSISTENCIA MEDICA HORISTA"...; na base 2026 é "Salários", "Enc.Salários Hor.", "Assistência Médica-Horista") + Iana R$ 0,7 mi (sem conta gestorial). Áreas (Rafaela) 100% ok.
- `Resultado TLM - MP27 vs R09.xlsm` da pasta (02/10) é cópia do Flash Set: 25 links, nenhum para Base Total MP27 nem Faturamento MP27. A única versão ligada ao MP27 é a de 08/10 em `_Versoes_anteriores` (base antiga por índice + faturamento 08/10). O Resumo de 15/10 aponta para o Resultado de 02/10 → números do Resumo NÃO confiáveis (MO var −1 mil; vendas 53 mi vs R09 93 mi).
- Faturamento MP27 é tarifa a custo → sem a MO, a tarifa também fica subestimada.

## De-para Pyplan (16/10) — PROPOSTA, aguardando OK dela
- Script `scripts/sap/tlm/mp2027/gerar_depara_pyplan.py` (só leitura) → `data/processed/tlm/mp2027/De-para Pyplan MP27 - 20261016.xlsx` (abas Resumo / De-para / Fica fora).
- Chave = Conta Gestorial; rótulo = Desc.Resultado da base 2026. 12 contas trocam a descrição; 4247100 (previdência) variável→fixo (2026 é 100% F).
- Simulação (SUMIFS case-insensitive, como o Excel): Pyplan fora hoje R$ 44.854.852 → 0 com de-para. Sobra só Iana R$ 695.940 (sem conta gestorial).
- Próximo passo após o OK: gerar Base Total MP27 corrigida (arquivo novo) e religar o Resultado MP27.

## Pacote de passagem do Caetano (16/10 tarde)
- `\SRV054-01BR.group.pirelli.com\DAC_TLM\MP27\Projeto_Claude_MP27_TLM\` (+ zip). CLAUDE.md, LEIA-ME, BRIEFING (dele), tlm_logic.json, 8 scripts `scripts/tlm/`, caderno.py/copc_estilo.py (padrão azul).
- **Causa raiz do bug da MO confirmada no código:** `mp27_resultado.py › montar_base` mapeia Desc.Resultado pela aba `Parametros` da `Base Total_Act_Flash_Fct.xlsx` (conta FISCAL → desc). Pyplan usa conta fiscal `4201000P` (não existe no mapa) → `fillna(Conta Gestorial Desc)` = nome cru do Pyplan. Bug nasceu na base única de 15/10; antes disso o RH vinha do draft por índice, com nomes certos.
- A última execução (15/10 ~23:02) gravou a Base Total e parou no passo 2: o Resultado na rede é a cópia crua do Flash Set. O BRIEFING dele não registra PBIT pós-base única (último: 14/10, PBIT 2.501 k com RH antigo de 46,0 M).
- `montar_resultado` faz `shutil.copy2(FLASH, RESULTADO)` → sobrescreve o Resultado anterior sem guardar versão (o de 14/10 se perdeu).
- Conflito de regra: Caetano = "VAR. como veio"; minha proposta de previdência 4247100 V→F (R$ 12 mil) contraria isso → decisão dela.
