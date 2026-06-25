# ProjetoEstatisticaEBi

Dashboard interativo de mortalidade infantil (0 a 4 anos) para Foz do Iguaçu, com dados tratados do SIM/DATASUS e uma versão em Streamlit para visualização.

## Conteúdo do projeto

- Dados brutos preservados em [dados_brutos](dados_brutos)
- Dataset tratado em [dados_tratados/dataset_final_tratado.csv](dados_tratados/dataset_final_tratado.csv)
- Dashboard em [app.py](app.py)
- Documentação e relatórios em [documentacao](documentacao)
- Notebook com o pipeline em [notebooks/tratamento_dados_pm3.ipynb](notebooks/tratamento_dados_pm3.ipynb)

## Como executar localmente

```powershell
python -m pip install -r requirements.txt
streamlit run app.py
```

## Deploy público

Este projeto pode ser publicado no Streamlit Cloud apontando o app para [app.py](app.py).
