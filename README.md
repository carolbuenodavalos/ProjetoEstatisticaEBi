# PM3_Tratamento_Dados - Tratamento de Dados

Projeto de tratamento, limpeza, transformação, normalização, discretização, feature engineering e documentação de uma base real do Sistema de Informação sobre Mortalidade (SIM/DATASUS).

## Tema

Óbitos de crianças de 0 a 4 anos em Foz do Iguaçu (PR), com dados de ano, causa básica, idade, sexo, raça/cor e local de ocorrência.

## Fonte dos dados

- Fonte: Sistema de Informação sobre Mortalidade (SIM), Ministério da Saúde/DATASUS.
- Link: https://dadosabertos.saude.gov.br/dataset/sim
- Data de acesso: 16/05/2026.
- Arquivo bruto preservado: `dados_brutos/dadosSIM2005-.csv`.

## Estrutura DataOps

```text
PM3_Tratamento_Dados/
├── dados_brutos/
│   ├── dadosSIM2005-.csv
│   └── base_original.csv
├── dados_tratados/
│   ├── dataset_final_tratado.csv
│   └── dataset_final_tratado_sem_outliers.csv
├── documentacao/
│   ├── catalogo_dados.xlsx
│   ├── catalogo_dados.md
│   ├── relatorio_final.pdf
│   ├── problemas_qualidade.csv
│   ├── valores_faltantes_antes_depois.csv
│   ├── estatisticas_descritivas.csv
│   ├── frequencias_categoricas.csv
│   ├── agregacao_obitos_por_ano.csv
│   ├── agregacao_obitos_por_ano_faixa.csv
│   ├── agregacao_causa_local.csv
│   ├── cruzamento_causa_faixa_etaria.csv
│   ├── cruzamento_causa_externa_faixa_etaria.csv
│   ├── cruzamento_causa_local_resumo.csv
│   └── graficos/
├── notebooks/
│   └── tratamento_dados_pm3.ipynb
├── projMens3.py
├── executar_projeto.bat
├── requirements.txt
└── README.md
```

## Como reproduzir

No PowerShell, a partir da pasta `PM3_Tratamento_Dados`:

Opção mais simples no Windows:

```powershell
.\executar_projeto.bat
```

Esse arquivo instala as dependências do `requirements.txt` e executa o projeto.

Opção manual:

```powershell
python projMens3.py
```

Se precisar instalar dependências em outro ambiente:

```powershell
python -m pip install -r requirements.txt
python projMens3.py
```

Ao executar `python projMens3.py`, o professor recria automaticamente:

- o dataset final em `dados_tratados/dataset_final_tratado.csv`;
- o relatório em `documentacao/relatorio_final.pdf`;
- o catálogo em `documentacao/catalogo_dados.xlsx`;
- as tabelas de evidência em `documentacao/*.csv`;
- os gráficos em `documentacao/graficos/*.png`.

Os gráficos também aparecem dentro do `relatorio_final.pdf`. Caso o professor queira ver cada imagem separadamente, basta abrir a pasta:

```text
documentacao/graficos/
```

O notebook `notebooks/tratamento_dados_pm3.ipynb` também executa o pipeline e possui uma célula final que exibe os gráficos gerados.

## Entregáveis principais

- Base original: `dados_brutos/dadosSIM2005-.csv`
- Dataset final tratado: `dados_tratados/dataset_final_tratado.csv`
- Dataset alternativo sem outliers estatísticos: `dados_tratados/dataset_final_tratado_sem_outliers.csv`
- Notebook: `notebooks/tratamento_dados_pm3.ipynb`
- Relatório final: `documentacao/relatorio_final.pdf`
- Catálogo de dados: `documentacao/catalogo_dados.xlsx`
- Evidências de análise: tabelas e gráficos na pasta `documentacao/`
- Cruzamentos analíticos: causa por faixa etária, causa externa por faixa etária e causa por local de ocorrência.

## Conferência dos requisitos

- Base original preservada em `dados_brutos/`.
- Dataset final tratado em `dados_tratados/`.
- Notebook e script Python para reprodução do tratamento.
- Arquivo `executar_projeto.bat` para instalar dependências e executar o pipeline.
- Relatório final em PDF com introdução, tema, fonte, objetivo, descrição da base, relação com BI/Big Data/Data Mining, diagnóstico de qualidade, EDA, seleção, limpeza, faltantes, outliers, transformações, agregações, normalização, discretização, feature engineering, dataset final, catálogo, DataOps, conclusão, limitações e próximos passos.
- Catálogo de dados no formato: Coluna, Descrição, Tipo, Exemplo, Origem, Tratamento aplicado e Uso.
