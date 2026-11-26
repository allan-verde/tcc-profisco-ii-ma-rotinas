# Rotinas de extração e conferência — Apêndice C

Monografia: *Modernização fiscal e tributária: o prometido, o registrado e o publicizado no Projeto de
Modernização da Gestão Fiscal do Estado do Maranhão (PROFISCO II-MA), 2019–2025*, de Allan Verde Rodrigues
(Ciências Contábeis, UFMA, 2026).

Esta pasta reúne as rotinas em Python descritas no Apêndice C e os arquivos que elas geraram. Os documentos
de entrada são públicos. Nenhuma rotina tira dado do texto da monografia; só o bloco P4 de
`checar_afirmacoes.py` lê o `.docx`, para conferir os quadros contra a extração (ver abaixo). Os 16 informes do BID não estão
aqui: baixe-os do repositório do Banco (operação BR-L1500). Os insumos da seção 6.3 estão em
`4 — arrecadação do ICMS (6.3)\insumos`.

## O que há em cada pasta

| Pasta | Rotina | O que faz | O que grava |
|---|---|---|---|
| 1 — informes do BID (PMR) | `extrair_pmr.py` | Lê os 16 informes semestrais de monitoramento (PMR) da operação BR-L1500 e grava, para cada produto e cada informe, a meta planejada (P), a meta ajustada (P(a)) e o registrado como alcançado (A), além das tabelas de alterações. Cada matriz é lida por dois caminhos (linhas da grade e posição das palavras), comparados célula a célula; uma amostra fixa de três produtos por informe é conferida contra o texto bruto da página. | `tabela_mestra_pmr.csv`, `alteracoes_pmr.csv`, `informes_pmr.csv` (incluídos) |
| 1 — informes do BID (PMR) | `checar_afirmacoes.py` | Recalcula, a partir dos três CSVs, as afirmações numéricas do capítulo 5 e indica o informe e a página que sustentam cada uma. | `resultado_auditoria.md` |
| 2 — Matriz A | `matriz_a_tres_camadas.py` | Classifica os 24 produtos pelas três metas de comparação (PME, planejada e ajustada), pelas regras da seção 4.4. As metas do PME vêm de `metas_pme.csv` (transcrição do Cuadro 1 do PME, conferida contra uma segunda transcrição em `meta_pme.py`). Resultado na camada do PME: 20 · 3 · 0 · 1; com `--estrito`, 17 · 3 · 0 · 4. | saída na tela |
| 3 — Diário Oficial | `doe_busca_profisco.py`, `doe_busca_termos_por_produto.py` | Pesquisam o termo "PROFISCO" no Diário Oficial do Estado do Maranhão, de 2019 a 2025, e, para os 12 produtos sem menção nesses resultados, um termo próprio de cada um, retendo os trechos que citam o programa ou o BID. | JSON com os resultados |
| 4 — arrecadação do ICMS (6.3) | `baixar_serie_icms_siconfi.py`, `baixar_insumos_6_3.py`, `calcular_6_3.py`, `grafico_6_3.py` | Baixam a Declaração de Contas Anuais (SICONFI), o IPCA, as projeções de população e o PIB (IBGE); corrigem os valores para reais de 2024; calculam as taxas da seção 6.3 e geram os Gráficos 1 e 2. O Gráfico 2 lê `tabela_mestra_pmr.csv`, da pasta 1; os percentuais de desembolso sob os anos foram lidos na primeira página dos informes e estão fixados em `grafico_6_3.py` (`MARCOS`). Uma conferência das contas de 2019 a 2021 lê o texto dos Relatórios da Administração Fazendária de 2020 e 2021. | os CSVs de insumo e o texto dos dois relatórios (incluídos em `insumos\`), `resultado_6_3_ate2025.csv` (incluído) e os gráficos (`grafico_6_3_icms.png` e `grafico_6_3_produtos.png`) |

## Como rodar

- Python 3.11 ou mais novo. Bibliotecas: `PyMuPDF` (fitz), `python-docx`, `pandas` e `matplotlib`.
- **PMR:** baixe os 16 informes do repositório do BID (operação BR-L1500) para uma pasta e rode
  `python extrair_pmr.py --fontes "<pasta dos PDFs>" --saida "<pasta de saída>"`. Depois,
  `python checar_afirmacoes.py`, na pasta dos CSVs.
- **Diário Oficial:** `python doe_busca_profisco.py PROFISCO saida.json`; a busca por produto grava em
  `doe\doe_termos_por_produto.json` (crie a subpasta `doe` antes).
- **Gráficos 1 e 2:** `python grafico_6_3.py`, na pasta 4. Roda sem ajuste nesta pasta: lê `resultado_6_3_ate2025.csv`
  ao lado e a `tabela_mestra_pmr.csv` da pasta 1.
- **Caminhos fixos:** as rotinas foram escritas na pasta de trabalho do autor. Antes de rodar, ajuste as
  constantes de caminho no início de cada arquivo (`PASTA_FONTES` em `extrair_pmr.py`; `PMR`
  em `matriz_a_tres_camadas.py`; `FONTES` nas rotinas da 6.3 e `SAIDA` em `baixar_serie_icms_siconfi.py`, que devem apontar para a pasta `insumos`).
- **Parte que lê a monografia:** só o bloco P4 de `checar_afirmacoes.py`, que compara, célula a célula, os
  quadros dos 24 produtos (o da seção 5.1 e o do Apêndice A) com a extração, nas colunas P, P(a) e A. É
  conferência do texto, não fonte de dado. Sem o `.docx`, esse bloco não roda; o resto independe dele.

## Diferença em relação à pasta de trabalho

Duas: em `checar_afirmacoes.py`, a lista das metas do PME passou a vir de `meta_pme.py` (cópia sem
alteração), porque na pasta de trabalho ela mora num script de montagem do documento; e, em
`matriz_a_tres_camadas.py`, o caminho do informe final (`PMR`) ficou para o leitor preencher.

Pasta preparada em 8 de novembro de 2026. Em 9 de novembro de 2026, `grafico_6_3.py` foi trocado pela versão do
Gráfico 1 com dois painéis (a arrecadação e os produtos).

**Versão 1.0, de 26 de novembro de 2026:** a Matriz A deixou de ler o `.docx` (as metas do PME vêm de
`metas_pme.csv`), a regra da seção 4.4 passou a ser a da opção B (20 · 3 · 0 · 1) e o Gráfico 1 foi dividido nos
Gráficos 1 e 2.
