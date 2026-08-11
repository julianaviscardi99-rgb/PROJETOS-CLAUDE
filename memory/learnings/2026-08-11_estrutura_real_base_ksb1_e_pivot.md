# Estrutura real do BASE_KSB1 e das Pivot Tables (inspecionado ao vivo em 2026-08-11)

Investigação feita direto no arquivo real `KSB1 June Actual 2026.xlsx` (rede, só leitura) pra sustentar a automação do passo 3 (voltar a alimentar via BASE_KSB1 + Pivot nativo — ver `memory/DECISOES.md` → reversão de 2026-08-11).

## Aba `BASE_KSB1`
- **Não é uma Tabela do Excel (ListObject)** — é um range simples. Sem defined names no workbook.
- Dimensão real em junho/2026: `A1:AM44976` (39 colunas, ~45 mil linhas).
- **Colunas 1-18 (A-R)** = dados brutos colados direto do export da KSB1. Confirmado 1:1 com o extrato bruto (`Bases SAP/KSB1_Actual_June_gestoriais.XLSX`): mesmo nome, mesma ordem de coluna (Data de lançamento, Data do documento, Centro custo, Classe de custo, Denom.classe custo, Fornecedor, Nome 1, Texto do pedido, Denominação, Documento de compras, Material, Moeda da transação, Nome do usuário, Texto breve material, Texto de cabeçalho de documento, Quantidade, Valor/MR, Soc.parc.negócios). Colar linha nova é copiar direto, sem remapear coluna.
- **Colunas 19-35 (S-AI)** = fórmulas por linha (VLOOKUP/IF), mesmo padrão em toda a base, só o número da linha muda (ex: `=VLOOKUP(D2,[1]Contas!$A:$J,10,0)` na linha 2, `=VLOOKUP(D44976,[1]Contas!$A:$J,10,0)` na última linha). "Arrastar fórmula" = pegar o texto da fórmula da última linha existente e trocar o número da linha pelo da linha nova, coluna por coluna.
  - Colunas: S=Mês, T=Gestorial, U=Descrição da Gestorial, V=Centro de Montagem, W=Variabilidade, X=M/H, Y=Variabilidade, Z=MF, AA=Chave Conta Razão&CM&Mês&CC, AB=Descrição, AC=DG/MO, AD=Detalhes, AE=Variabilidade CC, AF=Variabilidade Conta Razão, AG=Check Variabilidade, AH=Centro de Montagem(2), AI=Conta Gestão.
- **Colunas 36-39 (AJ-AM)**: AJ vazia; AK, AL, AM só têm fórmula na **linha 1** (célula de resumo isolada — `=SUMIF(S:S,AM1,Q:Q)`, `=SUM(Q:Q)`, `=MONTH(TODAY())-1`), não são padrão por linha. Não replicar essas pra linhas novas.
- **2 links externos** no workbook: `[1] Contas` = `Base_Contas_Contábeis_Fitted_22.xlsx` (base real usada nos VLOOKUP de Contas/Centros) e `[2] Indice` = `RHFitted <Mês> Actual <Ano>_.xlsx` — **é o link já confirmado como lixo/desatualizado** (ver `memory/DECISOES.md` 2026-08-11), mas mesmo assim é referenciado nas fórmulas de Variabilidade/M-H — não mexer, só replicar a fórmula como está.

## Pivot Tables nativas (`Pivot_Inter.` e `Pivot_Detalhes`)
- **Range de origem já fixo em `A1:AH1048576` (Pivot_Inter.) e `A1:AD1048576` (Pivot_Detalhes)** — cobre até a última linha possível do Excel. **Não é preciso reajustar o range da pivot ao adicionar linhas novas no BASE_KSB1**, só dar refresh (`PivotTable.RefreshTable`/`PivotCache.Refresh` via COM) depois de colar as linhas novas.
- Layout da `Pivot_Inter.`: linhas A3:H (campos de linha) = Gestorial, Descrição da Gestorial, Classe de custo (= Conta Fiscal), Denom.classe custo, Centro custo, MF, Centro de Montagem(2), Variabilidade. A partir da coluna I, uma coluna por número de mês (1, 2, 3...), com "Sum of Valor/MR". Dados começam na linha 5.
- **Classe de custo + Centro custo já são as mesmas duas chaves usadas na aba `Intermediária`** — dá pra ler o valor do mês direto do Pivot já refreshado, agrupando por essas duas colunas (pode haver split por Variabilidade/MF pra uma mesma combinação, então ao ler é preciso somar todas as linhas do pivot que batem na chave, não assumir uma linha só).

## Arquivo de teste disponível pra validação sem tocar em nada real
- `KSB1 May Actual 2026.xlsx` (mês anterior fechado) existe na rede — dá pra simular o fechamento de junho (maio + extrato bruto de junho de `Bases SAP/`) e comparar o resultado com o `KSB1 June Actual 2026.xlsx` real, que já tem o Pivot correto pronto (mesmo tipo de validação cega feita antes pro método antigo).
