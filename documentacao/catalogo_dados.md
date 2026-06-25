# Catálogo de Dados - PM3

## Fonte
- Base: Sistema de Informação sobre Mortalidade (SIM) - Ministério da Saúde/DATASUS.
- Link: https://dadosabertos.saude.gov.br/dataset/sim
- Data de acesso: 16/05/2026
- Descrição: Sistema de Informação sobre Mortalidade (SIM), Ministério da Saúde/DATASUS, com registros de óbitos e variáveis socioeconômicas, local de residência, local de ocorrência e causa básica do óbito.

## Métricas de processamento
- linhas_base_original: 1110
- colunas_base_original: 9
- linhas_apos_tratamento: 1052
- colunas_apos_tratamento: 21
- duplicatas_removidas: 58
- faltantes_antes: 1026
- faltantes_depois: 0
- escolaridade_faltante_pct: 91.16
- outliers_idade_iqr: 196
- limite_iqr_idade_inferior: 0.0
- limite_iqr_idade_superior: 0.61
- taxa_retencao_pct: 94.77
- fonte: https://dadosabertos.saude.gov.br/dataset/sim
- data_acesso: 16/05/2026

## Variáveis do dataset final

| Coluna | Descrição | Tipo | Exemplo | Origem | Tratamento | Uso |
|---|---|---|---|---|---|---|
| ano | Ano do óbito registrado no SIM. | Inteiro | 2021 | Base original | Conversão para inteiro. | Análise temporal |
| sigla_uf | Unidade da Federação da residência. | Texto | PR | Base original | Padronização em maiúsculas. | Dimensão geográfica |
| id_municipio_residencia | Código IBGE do município de residência. | Inteiro | 4108304 | Base original | Conversão para inteiro. | Dimensão geográfica |
| causa_basica | Código CID-10 da causa básica do óbito. | Texto | P369 | Base original | Remoção de espaços e padronização em maiúsculas. | Dimensão clínica |
| grupo_causa | Agrupamento analítico da causa básica. | Texto | Afecções perinatais | Criada | Derivação pelo primeiro caractere do CID-10. | Segmentação de causas |
| classificacao_evitabilidade | Classificação simplificada para leitura de evitabilidade. | Texto | Potencialmente evitável por atenção à saúde | Criada | Regra analítica baseada no grupo de causa. | Priorização de análise |
| idade | Idade no óbito em anos. | Numérico | 0.25 | Base original | Conversão para número e validação de 0 a 4 anos. | Medida |
| idade_meses | Idade convertida para meses. | Numérico | 3.00 | Criada | Idade em anos multiplicada por 12. | Medida interpretável |
| faixa_etaria | Faixa etária discretizada. | Categórico | Pós-neonatal | Criada | Discretização por regras de saúde infantil. | Segmentação |
| faixa_etaria_bi | Faixa etária consolidada para painéis de BI. | Categórico | Menor de 1 ano | Criada | Agrupamento em menor de 1 ano e 1 a 4 anos. | Filtro de dashboard |
| sexo | Código do sexo informado no registro. | Inteiro | 1 | Base original | Conversão para inteiro. | Dimensão categórica |
| sexo_desc | Descrição do sexo. | Categórico | Masculino | Criada | Mapeamento dos códigos de sexo. | Dimensão categórica |
| raca_cor | Código de raça/cor. | Inteiro | 4 | Base original | Valores ausentes preenchidos com código 9. | Dimensão sociodemográfica |
| raca_cor_desc | Descrição de raça/cor. | Categórico | Parda | Criada | Mapeamento dos códigos de raça/cor. | Dimensão sociodemográfica |
| local_ocorrencia | Código do local de ocorrência do óbito. | Inteiro | 1 | Base original | Conversão para inteiro. | Dimensão operacional |
| local_ocorrencia_desc | Descrição do local de ocorrência. | Categórico | Hospital | Criada | Mapeamento dos códigos de local. | Dimensão operacional |
| causa_externa | Indica se a causa pertence ao grupo de causas externas. | Booleano | False | Criada | Grupo de causa igual a causas externas. | Filtro analítico |
| obito_hospitalar | Indica se o óbito ocorreu em hospital. | Booleano | True | Criada | Local de ocorrência igual a hospital. | Indicador |
| idade_outlier_iqr | Indica outlier estatístico da idade pelo método IQR. | Booleano | False | Criada | Cálculo por intervalo interquartil. | Controle de extremos |
| idade_normalizada_minmax | Idade normalizada entre 0 e 1. | Numérico | 0.0625 | Criada | Min-Max Scaling manual. | Data Mining |
| idade_padronizada_zscore | Idade padronizada em z-score. | Numérico | -0.42 | Criada | Padronização pela média e desvio padrão. | Data Mining |