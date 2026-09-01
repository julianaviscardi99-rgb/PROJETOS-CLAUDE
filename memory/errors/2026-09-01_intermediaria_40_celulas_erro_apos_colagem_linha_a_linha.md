# Erro: "40 célula(s) gravada(s) como erro (#N/A) mesmo colando linha por linha" ao clicar em "③ Finalização da Base Intermediária"

**Sintoma:** popup "Erro ao finalizar a Base Intermediária" com "N célula(s) gravada(s) como erro (#N/A) mesmo colando linha por linha — abortando sem salvar." — reapareceu em 2026-09-01 mesmo depois das correções 5a (retry na leitura do Pivot_Inter, `ler_pivot_inter`) e 5b (limpar Y:AJ das linhas amarelas sobrando, `limpar_provisoes`/`preencher_provisoes_flash`), que resolveram outras 2 causas do MESMO texto de erro no mesmo dia — não é repetição delas.

**Causa raiz:** local diferente das anteriores. Essa checagem fica em `atualizar_base_intermediaria` (perto da linha 928, depois do loop de colagem linha a linha em `ws.Range(...).Value = [linha]`, linha ~924). O comentário do código já documentava o bug conhecido: escrever o array inteiro de uma vez via COM (pywin32) corrompe células aleatoriamente em erro (#N/A) — colar linha por linha (1 linha por chamada COM) reduz drasticamente a frequência (testado: 166 erros colando tudo de uma vez, 25 em blocos de 50, 0 num teste de 601 linhas colando linha a linha), **mas não elimina o risco por completo** — é uma corrupção probabilística de marshalling do COM, não determinística. O código antigo só conferia DEPOIS de colar tudo e abortava direto na primeira falha, sem tentar se autocorrigir — exigia que a usuária clicasse "Rode de novo" manualmente (reiniciando o processo inteiro do zero).

**Fix:** nova função `_corrigir_celulas_com_erro` (perto de `_celulas_com_erro`) — depois da colagem linha a linha, reconfere a área colada; se sobrar célula com erro, reescreve SÓ as células ruins (não a linha inteira, célula a célula, ainda mais granular) e reconfere de novo, até 5 tentativas (2s de espera entre elas) antes de desistir de vez. Só se a corrupção persistir depois de 5 rodadas de reparo é que aborta — e nesse caso a mensagem agora lista as linhas afetadas, não só a contagem.

---

## ATUALIZAÇÃO (mesmo dia): a causa raiz REAL era outra — não era bug de COM

O retry acima **não resolveu**: o erro voltou idêntico (40 células). A hipótese certa foi da usuária ("será que não tem a ver com células N/A da Base Intermediária"). Investigando os arquivos direto com `openpyxl(data_only=True)` — sem abrir Excel, técnica rápida e segura, repetir sempre que aparecer suspeita de #N/A:

- As 40 células estavam na **coluna G ("Centro de Montagem(2)") do Pivot_Inter.** do KSB1 — coluna de **RÓTULO**, não de valor.
- Todas de **RESENDE (MF 0483, CC 8333/8348/8349)** — R$ 78.289,32, primeiro mês da unidade com custo (Agosto/2026).
- **Causa:** fórmula da coluna AH do BASE_KSB1 = `=VLOOKUP(Z2,[1]Centros!$K$2:$L$9,2,0)`, com range **travado até a linha 9**; na `Base_Contas_Contábeis_Fitted_22.xlsx` (aba Centros, K:L) Resende/0483 está na **linha K10** — fora do range. A fórmula se propaga mês a mês via AutoFill (S:AI) do `gerar_ksb1_mensal.py`.
- **Por que despistou:** `_celulas_com_erro` só checava colunas de VALOR (I em diante), então o erro passava batido na leitura do Pivot_Inter e só estourava depois, na conferência pós-colagem (que varre tudo, inclusive rótulos), com a mensagem de "bug de marshalling conhecido".

**Fix definitivo:** `normalizar_formula_centro_montagem` em `gerar_ksb1_mensal.py` reescreve a coluna AH com o range ampliado (`$K$2:$L$100`) a cada geração; e `_celulas_com_erro` agora checa também as colunas de rótulo A-H, com a mensagem apontando a coluna afetada e o suspeito nº 1 (unidade/MF nova fora do range do de-para).

**Lição:** quando a MESMA contagem de células se repete rodada após rodada, NÃO é corrupção aleatória de COM — corrupção de marshalling é probabilística e varia. Contagem estável = dado real. Checar a origem com openpyxl antes de culpar o COM.

**Como reconhecer de novo:** se voltar a acontecer mesmo com esse retry (ou seja, sobrar erro depois das 5 tentativas de reparo célula a célula), aí sim é sinal de algo novo — não é mais o bug de marshalling (que já tenta se curar sozinho), pode ser um problema real nos dados de origem (ver `ler_pivot_inter`, que já filtra erro na leitura do Pivot_Inter antes mesmo de colar) ou uma instabilidade mais séria do COM/Excel na máquina.

**Nota:** existem HOJE (2026-09-01) duas checagens de "célula com erro #N/A" diferentes no mesmo arquivo — não confundir:
1. `ler_pivot_inter` (leitura do Pivot_Inter, origem) — retry por espera (link externo "assentando").
2. `_corrigir_celulas_com_erro` (escrita na Intermediária, destino, este arquivo) — retry por reescrita célula a célula (bug de marshalling do COM).
