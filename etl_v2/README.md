# ETL v2 — IberPhenol

Relê o Excel bruto (`Base dados_CP.xlsx`) para o modelo v2, em que o pico anotado é a unidade
central e cada valor guarda a aba e a célula de origem. Não altera o Excel nem o banco antigo
(`webapp/phenoldb.sqlite`).

```bash
python etl_v2/etl_v2.py
```

Saída em `etl_v2/out/`:

| Arquivo | Conteúdo |
| --- | --- |
| `iberphenol_v2.sqlite` | Banco v2 (esquema em `schema.sql`) |
| `csv/*.csv` | Uma tabela por arquivo, UTF-8 |
| `validation_report.md` | Contagens, leitura dos valores, comparação com o pipeline antigo, problemas |
| `csv/issue.csv` | Cada problema com aba e célula, para a curadoria |

## O que o v2 lê que o pipeline antigo perdia

- A linha 3 de cada aba como dado, não só como metadado.
- Cada coluna de concentração como uma condição: preparo (infusão, decocção…), sub-amostra
  ((a) judia, BB2 Montesinho…) e modo de quantificação (padrão DAD, MS, área de pico).
- As 13 abas "Planilha" (tabelas de picos com tempo de retenção), ainda sem espécie: o relatório
  sugere a aba de origem por sobreposição de nomes.
- Abas com o nome do composto na coluna de concentração (Cereja, Amendoa, Mirtilos), abas de
  totais por classe (Wine, Vitis vinifera_grape_2) e linhas de continuação (mesmo composto, outra fonte).
- Valor, desvio, unidade, base (extrato, peso seco, peso fresco, volume), preparo e flags
  (tr, nd, nq) separados; o texto original fica em `raw_text`.
- Todas as citações de cada aba, com detecção de textos que não são citação.

## Limites desta versão (0.1.0)

- IDs de composto (`IPH-C00001…`) são provisórios: mudam se a ordem das abas mudar. Um registro
  fixo de IDs entra antes da v1.0.
- Compostos são agrupados só por grafia (caixa, espaços, hífens); sinônimos e estruturas vêm
  na etapa de identidade química.
- Táxons são uma leitura inicial do nome da aba; a resolução em WFO/NCBI é a próxima etapa.
- Estruturas são as candidatas do pipeline antigo (busca por nome no PubChem), marcadas AUT-NR e
  conferidas só pela massa nominal.
