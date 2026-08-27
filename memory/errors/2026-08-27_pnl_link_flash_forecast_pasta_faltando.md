# Erro encontrado: links de Flash/Forecast do P&L Actual apontam para uma pasta que não existe (falta o nível "MM - Mês")

**Data:** 2026-08-27
**Onde:** arquivos `<MM>_P&L Fitted Units_Actual_<Mês>-26.xlsx` (Fitted Units), aba "Resumo Resultado Mês" (link de Flash, coluna E) e aba "Resumo Resultado Ano" (link de Forecast, coluna AF).

## O que estava errado

Confirmado em **Maio, Junho e Julho/2026** (todos os meses fechados checados): o link externo de Flash aponta para
```
...\Resultados Fitted\2026\<MM>_<Mês3>_Flash\<MM>_P&L Fitted Units_Flash_<Mês>-26...xlsx
```
— sem o nível de pasta do mês (`<MM> - <Mês3>`). O caminho real onde o arquivo está é:
```
...\Resultados Fitted\2026\<MM> - <Mês3>\<MM>_<Mês3>_Flash\<MM>_P&L Fitted Units_Flash_<Mês>-26...xlsx
```
Confirmado com `Test-Path` (PowerShell nativo, não só `ls` do Git Bash) que o caminho curto **não existe** — os três meses testados (Mai/Jun/Jul) devolveram `False`. Mesmo problema no link de Forecast (coluna AF, "Resumo Resultado Ano").

## Por que o valor mostrado parece certo (armadilha)

A célula mostra um número que hoje bate com o valor real do Flash de Julho (451.071) — mas é **valor em cache**, congelado de antes do link quebrar (provavelmente de quando a estrutura de pastas ainda não tinha o nível "MM - Mês"). Não há erro visível (`#REF!`) porque o Excel só mostra erro se tentar atualizar o link e falhar — com `UpdateLinks=0` (ou se ninguém forçar "Editar Links > Atualizar valores"), ele mantém o último valor calculado silenciosamente, mesmo que a fonte tenha mudado depois. Se o Flash de Julho for revisado agora, esse número no Actual **não vai acompanhar**, sem nenhum aviso.

## Achado relacionado (não confirmado como erro): share EO_CONSUMER vs EO_FITTED

O link de Mensalização (que aponta pro `MENS FITTED ACTUAL <MÊS>.xls`, saída do Passo 6) usa `\\FSS024-01BR.group.pirelli.com\EO_CONSUMER\BU FITTED\Forecast\...` em Maio/Junho, mas `\\FSS024-01BR.group.pirelli.com\EO_FITTED\BU FITTED\Forecast\...` em Julho. Os dois compartilhamentos existem de fato na rede. **A usuária confirmou (2026-08-27) que `EO_FITTED` é o correto/atual** — não investigado a fundo se `EO_CONSUMER` está desatualizado ou é um saldo residual de outra reorganização.

## Causa confirmada pela usuária

Não é um erro aleatório: a usuária confirmou (2026-08-27) que a pasta de rede foi reorganizada em algum momento (nível "MM - Mês" adicionado por cima da pasta de Ciclo) e os links dentro dos arquivos de P&L, que trocam manualmente todo mês, nunca foram re-apontados pra incluir esse nível novo — o mesmo padrão de "correção nunca propagada" do bug do link de PY (que também nasceu de uma mudança pontual não refletida no arquivo seguinte).

## Decisão (confirmada com a usuária, 2026-08-27)

- **Não corrigir retroativamente** Maio/Junho/Julho (mesma política já usada pro bug do link de PY, ver `2026-08-27_pnl_link_py_apontava_2024.md`).
- A automação do Passo 7 (`gerar_pnl.py`, ainda a implementar) deve sempre montar o caminho de destino do `ChangeLink` **incluindo o nível "MM - Mês"** (usar `resolver_pasta_ciclo`/mesmo padrão do resto do projeto, nunca replicar o caminho "curto" que os arquivos atuais têm) e sempre usar `EO_FITTED` (nunca `EO_CONSUMER`) pro link de Mensalização.

## Mecanismo de verificação usado

`Workbook.LinkSources(1)` pra listar os links de cada arquivo; comparação do caminho contra `Test-Path` (PowerShell, não só `ls` do Git Bash — o Git Bash pode mascarar diferenças de resolução de caminho UNC). Tudo em modo leitura (`ReadOnly=True`, `UpdateLinks=0`, `Close(SaveChanges=False)`) — nenhum arquivo real foi alterado.

Ver também `memory/DECISOES.md` (entrada 2026-08-27, continuação) e `memory/errors/2026-08-27_pnl_link_py_apontava_2024.md` (bug irmão, mesmo arquivo, link de PY).
