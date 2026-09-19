# Retrofit Flow — Automação de Relatórios de Ativos

Aplicação desktop (Python + interface gráfica própria) que elimina o preenchimento manual, célula a célula, de planilhas de controle de ativos (MBP) durante projetos de retrofit de hardware em larga escala.

## O problema

Em um projeto de retrofit de computadores (200–300 unidades por ciclo), cada equipamento trocado precisava ser lançado manualmente em uma planilha padrão: número de patrimônio, service tag, modelo e valor contábil — um processo repetitivo, sujeito a erro de digitação e que consumia uma analista em tempo integral por meses.

## A solução

1. Um técnico em campo preenche um checklist digital padronizado (app de terceiros) durante o retrofit.
2. O relatório exportado (Excel) é carregado na interface do Retrofit Flow.
3. O software identifica automaticamente as colunas de cada tipo de equipamento (notebook, CPU, monitor) via pandas/regex, limpa dados inconsistentes ("ilegível", "N/A") e faz uma validação (double-check) antes de gravar.
4. Os dados são injetados diretamente no template oficial via `openpyxl`, respeitando células mescladas e formatação contábil — uma planilha final por unidade, gerada em lote.

## Impacto medido

| Métrica | Processo manual | Retrofit Flow |
|---|---|---|
| Tempo por ciclo (200–300 unidades) | ~4,5 meses (~800h) | < 15 minutos |
| Tempo médio por unidade | ~172 minutos | ~3,6 segundos |
| Risco de erro de digitação | Alto | Eliminado (validação automática) |
| Escalabilidade | Limitada pela capacidade humana | Processa 300 ou 1000 unidades no mesmo tempo |

## Stack

- Python
- `pandas` + `regex` para extração e limpeza de dados
- `openpyxl` para geração das planilhas finais
- `customtkinter` para a interface gráfica desktop

> Versão sanitizada para portfólio: nome de arquivo-modelo e valores contábeis foram substituídos por placeholders genéricos.
