# Erro: "The message filter indicated that the application is busy" (-2147417846) ao clicar em "Finalização da Base Intermediária"

**Sintoma:** popup "Erro ao finalizar a Base Intermediária" com `(-2147417846, 'The message filter indicated that the application is busy.', None, None)` — acontecia toda vez (não intermitente), no Passo 3 → botão "Finalização da Base Intermediária".

**Causa raiz:** `com_retry` (ksb1_core.py) já existia pra esse exato erro COM (RPC_E_SERVERCALL_RETRYLATER — Excel "assentando" logo após um `RefreshAll`/`CalculateFullRebuild` pesado), mas duas chamadas em `gerar_base_intermediaria.py` ficaram sem essa proteção:
- `atualizar_comparacao_flash` (linha ~235) — roda logo depois do `RefreshAll` da aba Pivot (linha 925), exatamente o cenário que o `com_retry` foi feito pra cobrir.
- `ler_forecast_despesas_mao_de_obra` (linha ~356) — chamada por `atualizar_comparacao_forecast`, mesmo padrão, ciclo Flash.

Por rodarem sempre logo após o RefreshAll pesado, o erro acontecia de forma consistente (não aleatória), diferente do resto do arquivo onde o `com_retry` já mascarava o problema.

**Fix:** envolvidas em `com_retry(excel.CalculateFullRebuild, log=log)`, igual ao padrão já usado em todo o resto do arquivo.

Aproveitado pra corrigir o mesmo padrão desprotegido (`excel.CalculateFullRebuild()` / `wb.Save()` / `wb.Close()` sem `com_retry`) em `lancar_provisoes` e `atualizar_provisoes` (botões "Lançar Provisões" / "Atualizar Provisões" do Passo 3) — ainda não tinham dado erro, mas mesma causa raiz, corrigido preventivamente.

**Como reconhecer de novo:** se aparecer esse mesmo código de erro `-2147417846` em qualquer botão do fluxo KSB1/Base Intermediária, procurar por chamadas COM (`excel.*`, `wb.*`, `ws.*`) SEM `com_retry` em volta — especialmente logo após `RefreshAll`/`CalculateFullRebuild`/`Save` pesados.
