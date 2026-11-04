# HoraLu — Time André: passo a passo de implantação

## 1. Trocar o código
1. Na planilha do André: **Extensões > Apps Script**.
2. Faça uma cópia de segurança do código atual (cole num bloco de notas).
3. Apague tudo e cole o conteúdo de `HoraLu_Andre.gs`. Salve.
4. Recarregue a planilha. Vai aparecer o menu **HoraLu — Time André**.

## 2. Conferir a aba AGENTES
- Coluna G: e-mail de cada agente. Coluna H: escreva `OK` depois de conferir.
- Coluna C (Operação): quem tiver `Aprendiz` recebe o e-mail de aprendiz (TOP 3, visual colorido). Os demais recebem o de agente (TOP 5).

## 3. Testar (os envios para o time já vêm desligados)
1. Menu **Processar relatórios agora**. Na primeira vez o Google pede autorização (Gmail e Sheets): aceite.
   - Atenção: o **consolidado do gestor continua ligado** e vai para o André com a Luana em cópia. Para o teste não chegar à Luana, mude `ENVIAR_PAINEL_GESTOR` para `false` até terminar.
2. Menu **Enviar teste (só para o gestor)**. O André recebe o consolidado e um e-mail de cada agente, com o assunto "[TESTE VISÃO AGENTE] Como chegará para: <nome>".
3. Confira as metas Nuvens/Estrelas, o ranking e os nomes. Repita durante 1 ou 2 dias.

## 4. Ligar para o time
1. No código, mude `ENVIAR_EMAILS_INDIVIDUAIS: false` para `true` (e `ENVIAR_PAINEL_GESTOR` de volta para `true`, se tiver desligado). Salve.
2. Menu **Instalar execução automática (15 min)**. Esse comando também remove o gatilho antigo de 1 hora.

## O que mudou em relação ao script antigo
- **E-mails do HoraLu**: Meta Nuvens = meta ajustada; Meta Estrelas = meta ajustada × 1,2. Ranking por % da meta ajustada, TOP 5 (agentes) e TOP 3 (aprendizes), com "faltam X atendimentos para passar Fulano". O consolidado mostra o ranking completo com o status do André (Meta Batida / Atenção / Crítico).
- **Falta/Folga**: quem não aparece nos CSVs da Genesys fica de fora do ranking e do dashboard, aparece cinza no consolidado e não recebe e-mail. O processamento só é cancelado se faltar mais da metade do time.
- **Gatilho a cada 15 min**: com relatórios às 14h, 15h30, 17h30 e 19h, o e-mail sai até 15 min depois de o relatório chegar. Não duplica: cada par de relatórios é processado uma vez só.
- **Gmail**: igual ao script antigo (qualquer CSV recente com o cabeçalho da Genesys).
- **Aba LOG** (criada sozinha): registra cada processamento e qualquer falha de envio. Uma falha não impede os outros envios.
- Se um par de relatórios der erro, o script não tenta de novo a cada 15 min. Depois de corrigir, use **Liberar reprocessamento**.

## Ajustes rápidos (bloco CFG, no topo do código)
| Campo | Para quê |
|---|---|
| `EMAIL_GESTOR`, `EMAIL_COORDENADORA` | Quem recebe o consolidado (coordenação em cópia) |
| `EMAIL_TESTE` | Quem recebe tudo no modo teste |
| `NOME_TIME` | Assinatura "TIME ANDRÉ, o melhor time de todos!" |
| `FATOR_ESTRELAS`, `TOP_AGENTES`, `TOP_APRENDIZES` | Regras do HoraLu |
| `LIMITE_AUSENTES` | Fração do time ausente que faz cancelar (0,5 = metade) |

## CSAT e RPC (atualização de 30/10)
- O script lê também o relatório **Avaliações.csv** da Genesys (o mesmo e-mail, nos mesmos horários) e conta, por agente, as etiquetas `[CSAT SAC] - PROMOTOR/NEUTRO/DETRATOR` e `[CSAT SAC] - RPC SIM/NÃO` da coluna "Resultados com êxito". Só conta ligações do dia analisado; em transferência, a avaliação vai para o último agente da lista.
- CSAT = promotores ÷ respostas (meta 92%). RPC = sim ÷ respostas (meta 82%).
- **Nota HoraLu** (ordena o ranking): 70% produção + 15% CSAT + 15% RPC, cada um em relação à meta. Com menos de 3 avaliações, o indicador conta como na meta.
- Resultado gravado nas abas **RAW_AVALIACOES** e **CONSOLIDADO_AVALIACOES**, nas colunas M a S da BASE_CONSOLIDADA (inclui detratores e neutros) e na coluna E (Nota HoraLu) da aba RANKING.
- Se o relatório de avaliações não chegar, os e-mails saem do mesmo jeito (CSAT/RPC aparecem como "sem avaliações") e a aba LOG registra um aviso.
- Ajustes no CFG: `META_CSAT`, `META_RPC`, `MIN_AVALIACOES`, `PESOS_NOTA`.
