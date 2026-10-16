# MP27 TLM — Master Plan 2027 da TLM (empresa 0981)

> Projeto herdado do Caetano (Planning & Control) em 16/10/2026. A partir de agora quem controla o custo
> da TLM é a Juliana. Esta pasta é autocontida: scripts, regras de negócio, histórico de decisões e pendências.

## Ao iniciar qualquer sessão — ler nesta ordem
1. `memory/BRIEFING.md` — o que foi feito, decisões, números, pendências e a **ordem para refazer**.
   As seções estão em ordem de data; a última é o estado atual.
2. `ontology/tlm_logic.json` → chave `master_plan` — as regras validadas (perímetro, tarifa, base única...).
   Usar Grep pela sub-chave, não ler o JSON inteiro.
3. `memory/errors/` e `memory/learnings/` — armadilhas já pagas. Nunca repetir um erro registrado.
4. `memory/contexto/` — briefings irmãos (Resultado mensal e Faturamento da TLM), só se o assunto tocar neles.

**O briefing diz ONDE olhar, não O QUE É VERDADE.** Antes de dizer que uma pendência está aberta,
conferir a fonte (Outlook, pasta `Arquivos Recebidos`, arquivos da `TLM_Automation`).

## Onde moram os dados (rede DAC_TLM)
- Raiz: `\\SRV054-01BR.group.pirelli.com\DAC_TLM\MP27\1ª Versão MP27 vs R09\`
- `Arquivos Recebidos\` — tudo que as áreas mandam (Rafaela: despesas · Iana: RH · Labour Cost). **Nunca editar.**
- `TLM_Automation\` — só o que os scripts geram (draft, carga J01, Base Total, Resultado, Resumo).
- `Volume\` — planilha de faturamento MP27 (tarifa, toneladas).
- Base R09 / Actual 2026: `\DAC_TLM\2026\Bases\` e `\DAC_TLM\2026\1. FCST\R09\`. Gabarito: `\DAC_TLM\MP26\`.

## Scripts (`scripts/tlm/`) — rodam com `python scripts/tlm/<script>.py`, sem argumentos
| Script | O que faz |
|---|---|
| `mp27_draft.py` | draft original R09 ano cheio × índice por conta (substituído pela base das áreas, mantido como referência) |
| `mp27_rh_iana_981.py` | pega no Outlook o último anexo da Iana (fio "Custo fixo - MP27") e filtra a empresa 981 |
| `mp27_base_areas.py` | **base única**: áreas (Rafaela) + custo fixo RH (Iana) + MDO (Labour Cost) → draft + carga J01 |
| `mp27_sobreposicao.py` | lista CC × classe que aparece em mais de uma base (possível duplicidade) |
| `mp27_resultado.py` | Base Total MP27 + `Resultado TLM - MP27 vs R09.xlsm` (usa Excel via win32com) |
| `mp27_faturamento_base.py` / `mp27_tarifa.py` | planilha de faturamento e tarifa a custo (regra MP26) |
| `mp27_resumo.py` | `Resumo Resultado TLM - MP27 vs R09.xlsx` |

`scripts/antidumping/caderno.py` e `scripts/copc/copc_estilo.py` são só o gravador de Excel no padrão azul — não mexer.

**Ordem para refazer depois de trocar qualquer custo:**
`mp27_base_areas.py` → `mp27_resultado.py` → `mp27_faturamento_base.py` → `mp27_resultado.py` (2ª volta) → `mp27_resumo.py`.
Sem a 2ª volta a tarifa fica com o custo velho (já deu PBIT errado por isso).

## Regras do projeto
- **Nada entra que não tenha vindo das bases recebidas** — classe, valor, VAR. e ORIG. como vieram.
  Única exceção: VAR. (fixo/variável) da base das áreas vem do R09/Actual 2026 do mesmo CC × classe.
- **Perímetro 2027:** Cabreúva (mini-fábrica 108, WH CBLog/CB_AG) não existe; faturamento MP Barueri (BAMP) fora.
- Mão de obra nunca é indexada — vem do RH.
- **Nunca sobrescrever arquivo original da rede.** Saída nova com data no nome; versão anterior vai para
  `TLM_Automation\_Versoes_anteriores\`. Arquivo recebido nunca é editado.
- Todo Excel gravado sai formatado (padrão azul via `caderno.py`) — nada de planilha crua.
- Excel por win32com: abrir as fontes dos links antes de recalcular (SUMIF em pasta fechada dá #VALUE!).
  **Nunca `taskkill /IM EXCEL.exe`** (derruba as planilhas abertas da usuária) — se precisar, matar por PID.
- Nada é escrito no SAP por este projeto. A carga J01 é gerada em arquivo; quem sobe é a usuária, manualmente.
- E-mail: só rascunhar; enviar só com confirmação explícita.

## Jeito de trabalhar
- Português BR sempre. Respostas curtas e diretas.
- Testar (rodar o script, conferir totais) antes de entregar.
- Ao aprender regra nova, decisão ou armadilha: atualizar na hora `memory/BRIEFING.md` (seção com data),
  `ontology/tlm_logic.json › master_plan` e criar `memory/learnings/AAAA-MM-DD_HH-MM-SS_<nome>.md`.
  Ao errar: `memory/errors/AAAA-MM-DD_HH-MM-SS_<nome>.md` (Data / Tarefa / Erro / Causa / Solução / Evitar).
- Pessoas: Rafaela e Jefferson (TLM, despesas das áreas) · Iana Brito (RH corporativo) · Caetano (dono anterior).
