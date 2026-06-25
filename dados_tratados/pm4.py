from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

# Configuração da página
st.set_page_config(page_title="Dashboard SIM - PM4", layout="wide")
st.title("📊 Dashboard de Mortalidade Infantil (0 a 4 anos) - Foz do Iguaçu")
st.markdown("**Fonte dos dados:** Sistema de Informação sobre Mortalidade (SIM/DATASUS) | **Autor:** Seu Nome")

# Carregando a base tratada do PM3
@st.cache_data
def load_data():
    base_dir = Path(__file__).resolve().parent
    candidates = [
        base_dir / "dataset_final_tratado.csv",
        base_dir.parent / "dados_tratados" / "dataset_final_tratado.csv",
        base_dir.parent / "dataset_final_tratado.csv",
    ]

    for path in candidates:
        if path.exists():
            return pd.read_csv(path, sep=';', encoding='utf-8-sig')

    raise FileNotFoundError("Arquivo dataset_final_tratado.csv não encontrado. Verifique se o arquivo existe na pasta dados_tratados.")

df = load_data()

# ----------------- FILTRO INTERATIVO -----------------
st.sidebar.header("Filtros do Dashboard")
anos_selecionados = st.sidebar.slider("Selecione o Período", int(df['ano'].min()), int(df['ano'].max()), (2015, 2021))
faixa_selecionada = st.sidebar.multiselect("Faixa Etária", df['faixa_etaria_bi'].unique(), default=df['faixa_etaria_bi'].unique())

# Aplicando filtro
df_filtrado = df[(df['ano'] >= anos_selecionados[0]) & 
                 (df['ano'] <= anos_selecionados[1]) & 
                 (df['faixa_etaria_bi'].isin(faixa_selecionada))]

# ----------------- INDICADORES (KPIs) -----------------
col1, col2, col3 = st.columns(3)
col1.metric("Total de Óbitos (Período)", len(df_filtrado))
col2.metric("Média de Idade (Anos)", round(df_filtrado['idade'].mean(), 2))

if len(df_filtrado) > 0:
    pct_hosp = (len(df_filtrado[df_filtrado['local_ocorrencia_desc'] == 'Hospital']) / len(df_filtrado)) * 100
else:
    pct_hosp = 0
col3.metric("% Óbitos Hospitalares", f"{pct_hosp:.1f}%")

st.divider()

# ----------------- GRÁFICOS -----------------
c1, c2 = st.columns(2)

# Gráfico 1: Linha
evolucao = df_filtrado.groupby('ano').size().reset_index(name='obitos')
fig1 = px.line(evolucao, x='ano', y='obitos', markers=True, title="1. Evolução Histórica de Óbitos")
c1.plotly_chart(fig1, use_container_width=True)

# Gráfico 2: Barras
ranking_causas = df_filtrado['grupo_causa'].value_counts().reset_index()
ranking_causas.columns = ['Causa', 'Total']
fig2 = px.bar(ranking_causas, x='Total', y='Causa', orientation='h', title="2. Principais Grupos de Causas", color='Total', color_continuous_scale='Reds')
fig2.update_layout(yaxis={'categoryorder':'total ascending'})
c2.plotly_chart(fig2, use_container_width=True)

c3, c4 = st.columns(2)

# Gráfico 3: Colunas (Substituindo o de Pizza)
faixa_contagem = df_filtrado['faixa_etaria_bi'].value_counts().reset_index()
faixa_contagem.columns = ['Faixa Etária', 'Total de Óbitos']

fig3 = px.bar(faixa_contagem, 
              x='Faixa Etária', 
              y='Total de Óbitos', 
              title="3. Volume de Óbitos por Faixa Etária", 
              text_auto=True, 
              color='Faixa Etária', 
              color_discrete_sequence=['#1f77b4', '#ff7f0e'])

# Escondendo a legenda lateral para deixar o gráfico mais limpo
fig3.update_layout(showlegend=False) 
c3.plotly_chart(fig3, use_container_width=True)

# Gráfico 4: Barras Empilhadas 100%
local_faixa = df_filtrado.groupby(['faixa_etaria_bi', 'local_ocorrencia_desc']).size().reset_index(name='contagem')

if not local_faixa.empty:
    local_faixa_pivot = local_faixa.pivot(index='faixa_etaria_bi', columns='local_ocorrencia_desc', values='contagem').fillna(0)
    total_por_faixa = local_faixa_pivot.sum(axis=1)
    local_faixa_percentual = local_faixa_pivot.div(total_por_faixa, axis=0) * 100
    local_faixa_percentual = local_faixa_percentual.reset_index()
    fig4 = px.bar(
        local_faixa_percentual,
        x='faixa_etaria_bi',
        y=local_faixa_percentual.columns[1:],
        title='4. Local de Ocorrência por Faixa Etária',
        barmode='stack'
    )
    fig4.update_layout(yaxis_title='Percentual (%)', xaxis_title='Faixa Etária')
else:
    fig4 = px.bar(title='4. Local de Ocorrência por Faixa Etária')

c4.plotly_chart(fig4, use_container_width=True)