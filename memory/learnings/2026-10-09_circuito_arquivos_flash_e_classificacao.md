# Circuito Panamericano — arquivos do Flash Set/26 e teste de classificação (2026-10-09)

## Onde ficam os arquivos (rede)
`\\FSS024-01BR.group.pirelli.com\CONTROLLING\Reporting\Reporting ACT_FCST_MP Cons. e Ind\Circuito Panamericano\2026\09 - Sep\02. Flash`
(irmãs: `01. Forecast`, `03. Actual`). Share diferente do GFU_DAC da Fitted Units.

## Arquivos na pasta Flash (Set/26)
- `KSB1 - Jan-Dez_.xlsx` — **onde a Juliana classifica manualmente.** 1 aba por mês (Jan–Dez) + `Parâmetros` (descrição/texto do pedido → classificação, ~100 linhas) + `Check` (realizado x orçamento por classificação). Coluna final `Classificação` só preenchida em Ago e Set (Jan–Jul têm layout antigo, sem essa coluna).
- `Base KSB1 - 09.2026.XLSX` — extração crua da KSB1 do mês (162 linhas).
- `MENS CP 2026 - Flash September.xlsx` — mensalização (abas LOCAL, INTEGRADA, Mensalização, MDO, Proposta).
- `09_P&L Circuito Panamericano_Flash_September_26.xlsx` (e versão `_`) — P&L (abas Resumo Resultado Mês/Ano).
- `Base de Custo Flash September.xlsx`, `Fast Provisão_final_..._CP.xlsx`, `Reclassificação Motorsport - September.xlsx` — apoio.

## Estrutura da KSB1 classificada
Linhas "RH" (rateio de pessoal, par centro 4000 / HR_DUMMY que se anulam) = 90 das 163 linhas de Set. Linhas sem fornecedor/classe de custo = lançamentos DAC / reclassificações MS (Motorsport).

## Teste (treino = Ago classificado, prova = Set classificado)
Chave **Classe de custo + Fornecedor**: cobre 145/163 (89%) das linhas, acerta 141/145 (97%). Só classe de custo: cobre 96% mas acerta 81% (ambígua). Só texto do pedido: cobre 17%, acerta 100%.
Falhas: linhas novas (sem histórico, 15 linhas, líquido R$ 36,5k, algumas já cobertas pela aba Parâmetros por texto do pedido) e fornecedores que prestam mais de um serviço (ex.: 4211333742 = Limpeza ou Jardinagem; 4211330102 = café insumo ou locação de máquina; 4211323022 = locação carro, 3 subtipos). Para esses o texto do pedido desempata.
Teste feito só com 2 meses; ainda não usado Parâmetros nem Jan–Jul.

## Trabalho feito sobre cópia
Arquivos copiados para o scratchpad da sessão; originais da rede NÃO foram alterados. Sem pandas no ambiente (requirements.txt só tem openpyxl) — usar openpyxl.

## Atualização (mesma sessão): script criado
`scripts/sap/circuito_panamericano/classificar_ksb1.py` — sugere a classificação da KSB1 do mês. Histórico = Jan–Jul (coluna 'Nº documento', pintada de verde) + Ago em diante (coluna 'Classificação') + aba Parâmetros. Gera xlsx novo em `data/processed/circuito_panamericano/` (nunca altera o original; copia para pasta temporária antes de ler). `--avaliar` faz o teste contra o que já foi classificado.
Resultado do teste (confiança alta): Mar 95%, Mai 98%, Jul 94%, Ago 94%, Set 99% de acerto; cobertura alta 61–90% das linhas. Obs.: Mar–Ago usam meses posteriores como histórico também (otimista); Set é o caso realista.
Correções que subiram o acerto: (1) linhas sem classe de custo e sem fornecedor (DAC / reclassificação Motorsport) vão sempre para revisão manual; (2) histórico exato tem prioridade sobre a aba Parâmetros (desatualizada, ex.: Optimus hoje é 'I.T - Software Optimus', não 'Predial').
Fornecedores multi-serviço seguem como pontos de atenção: 4211319723 e 4211328098 (irrigação x predial/elétrica), 4211333742 (limpeza x jardinagem).

## Correções da Juliana (2026-10-09)
- Optimus é software de **IT** (não predial). Regra fixa em `scripts/sap/circuito_panamericano/regras_manuais.json` (vale acima do histórico e dos Parâmetros). A aba Parâmetros da rede continua com "Predial" — não foi alterada.
- Ela NÃO quer que o arquivo/pasta da área na rede seja modificado por enquanto: só leitura, saída em data/processed/.
