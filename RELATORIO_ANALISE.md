# Relatório de análise - Análise.xlsx

Fonte: `Análise.xlsx`, aba `Análise`

Período analisado: 01/01/2026 a 30/01/2026

## Resumo executivo

- Foram analisadas 26.437 ocorrências.
- A principal categoria é `PERTURBAÇÃO AO SOSSEGO ALHEIO`, com 18.725 registros, equivalente a 70,8% da base.
- `MARIA DA PENHA` aparece em segundo lugar, com 3.942 registros.
- 71,1% das ocorrências aconteceram no período de noite ou madrugada.
- A base possui 787 cidades mapeadas.

## Principais categorias de crime

| Crime | Registros |
| --- | ---: |
| PERTURBAÇÃO AO SOSSEGO ALHEIO | 18.725 |
| MARIA DA PENHA | 3.942 |
| FURTO | 1.266 |
| CVP | 901 |
| DISPARO DE ARMA | 552 |
| LESÃO CORPORAL | 534 |

## Turnos

| Turno | Registros |
| --- | ---: |
| Noite | 12.277 |
| Madrugada | 6.516 |
| Tarde | 4.793 |
| Manhã | 2.851 |

## Batalhões com mais registros

| Batalhão | Registros |
| --- | ---: |
| 14º BPM | 2.453 |
| 19º BPM | 1.946 |
| 6º BPM | 1.634 |
| 9º BPM | 1.367 |
| 11º BPM | 1.337 |
| 3º BPM | 1.307 |
| 12º BPM | 1.304 |
| 18º BPM | 1.266 |

## Cidades com maior taxa por 10 mil habitantes

| Cidade | Registros | Taxa por 10 mil hab. |
| --- | ---: | ---: |
| Maracanaú | 1.282 | 51,0 |
| Paracuru | 212 | 50,6 |
| Pacatuba | 430 | 49,7 |
| Aquiraz | 409 | 47,9 |
| Camocim | 289 | 44,3 |
| Russas | 312 | 41,7 |
| Pindoretama | 102 | 40,9 |
| Beberibe | 217 | 38,8 |

## Observações

- A análise foi incorporada ao dashboard Flask do projeto.
- O backend não depende de `pandas` ou `openpyxl`; ele lê o `.xlsx` diretamente usando a biblioteca padrão do Python.
- A rota `/api/dashboard` retorna os agregados em JSON para uso no front-end.
