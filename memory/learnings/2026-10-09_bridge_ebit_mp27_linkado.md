# 2026-10-09 — Bridge de EBIT MP'27 linkado à v3 teste daniel

- Revisão da `MENS FITTED MP27_v3_teste daniel.xlsx`: erros da v2 daniel corrigidos por ela (IBI linha 77 CNH usava E54; MO da RES sem o reajuste de 7%). Volumes agora vêm de `BU FITTED\PO's\Volume\MP'27.xlsx`.
- Respostas dela: o preço da RES de 7,3 é o preço de lista (a unidade identifica os ganhos); o frete da IBI subiu pela nova regra da ANTT; o frete da GO antes estava superestimado; o ROS esperado foi ajustado para 13,8%; o DU baixo da GO em outubro é a parada de férias coletivas da Jeep; a MO é estimativa até a informação oficial do RH; o rateio fixo só será ajustado depois de quadrar o faturamento.
- Criado `MP 2027\Bridge EBIT MP27 - linkado v3 teste daniel_v2.xlsx` (a v1 ficou sem a premissa da Jaguar). O script está só no scratchpad (`build.py`).
- **Técnica que funcionou:** criar o vínculo externo real no openpyxl com `ExternalLink`/`ExternalBook` + `Relationship(Target="<nome do arquivo>")` relativo (mesma pasta), fórmulas no formato `[1]SJP!$Q$8` e cache dos valores em `ExternalSheetData`. Testado via Excel COM com uma cópia da fonte ao lado: links resolvem e 0 erros.
- Lógica da IBI: Volume = T8 − JLR!K24; Mix = JLR!L27; Preço = JLR!M27; Jaguar = −margem direta; Demais var = T18 − (fat JLR − MD) − ANTT. ANTT = vol. Fiat × (−1,9763 − (−1,57073125)) = −1.073,8.
- Jaguar vs R9 informado por ela = **−550** (margem direta calculada −522,5; a diferença de −27,5 cai na linha de rateio fixo até ela inserir o rateio de Ibirité).
- Achado: o R9 da aba TOTAL (S43 = 13.001,4) é 188,2 maior que a soma das unidades — está em `TOTAL!L47` (agosto, digitado). O MP'26 da RES tem só o EBIT (337,5), sem detalhe na coluna V.
