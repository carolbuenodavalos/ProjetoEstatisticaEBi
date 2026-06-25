from pathlib import Path
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Dashboard SIM - PM4", layout="wide")
st.title("📊 Dashboard de Mortalidade Infantil (0 a 4 anos) - Foz do Iguaçu")
st.markdown("**Fonte dos dados:** Sistema de Informação sobre Mortalidade (SIM/DATASUS) | **Autor:** Seu Nome")

@st.cache_data
def load_data():
    base_dir = Path(__file__).resolve().parent
    csv_path = base_dir / "dados_tratados" / "dataset_final_tratado.csv"
    return pd.read_csv(csv_path, sep=';', encoding='utf-8-sig')


df = load_data()

st.sidebar.header("Filtros do Dashboard")
anos_selecionados = st.sidebar.slider("Selecione o Período", int(df['ano'].min()), int(df['ano'].max()), (2015, 2021))
faixa_selecionada = st.sidebar.multiselect("Faixa Etária", df['faixa_etaria_bi'].unique(), default=df['faixa_etaria_bi'].unique())

df_filtrado = df[(df['ano'] >= anos_selecionados[0]) & (df['ano'] <= anos_selecionados[1]) & (df['faixa_etaria_bi'].isin(faixa_selecionada))]

col1, col2, col3 = st.columns(3)
col1.metric("Total de Óbitos (Período)", len(df_filtrado))
col2.metric("Média de Idade (Anos)", round(df_filtrado['idade'].mean(), 2))

if len(df_filtrado) > 0:
    pct_hosp = (len(df_filtrado[df_filtrado['local_ocorrencia_desc'] == 'Hospital']) / len(df_filtrado)) * 100
else:
    pct_hosp = 0
col3.metric("% Óbitos Hospitalares", f"{pct_hosp:.1f}%")

st.divider()

c1, c2 = st.columns(2)

evolucao = df_filtrado.groupby('ano').size().reset_index(name='obitos')
fig1 = px.line(evolucao, x='ano', y='obitos', markers=True, title='1. Evolução Histórica de Óbitos')
c1.plotly_chart(fig1, use_container_width=True)

ranking_causas = df_filtrado['grupo_causa'].value_counts().reset_index()
ranking_causas.columns = ['Causa', 'Total']
fig2 = px.bar(ranking_causas, x='Total', y='Causa', orientation='h', title='2. Principais Grupos de Causas', color='Total', color_continuous_scale='Reds')
fig2.update_layout(yaxis={'categoryorder':'total ascending'})
c2.plotly_chart(fig2, use_container_width=True)

c3, c4 = st.columns(2)
fig3 = px.pie(df_filtrado, names='faixa_etaria_bi', title='3. Proporção por Faixa Etária', hole=0.4, color_discrete_sequence=['#1f77b4', '#ff7f0e'])
c3.plotly_chart(fig3, use_container_width=True)

local_faixa = df_filtrado.groupby(['faixa_etaria_bi', 'local_ocorrencia_desc']).size().reset_index(name='contagem')
if not local_faixa.empty:
    local_faixa_pivot = local_faixa.pivot(index='faixa_etaria_bi', columns='local_ocorrencia_desc', values='contagem').fillna(0)
    total_por_faixa = local_faixa_pivot.sum(axis=1)
    local_faixa_percentual = local_faixa_pivot.div(total_por_faixa, axis=0) * 100
    local_faixa_percentual = local_faixa_percentual.reset_index()
    fig4 = px.bar(local_faixa_percentual, x='faixa_etaria_bi', y=local_faixa_percentual.columns[1:], title='4. Local de Ocorrência por Faixa Etária', barmode='stack')
    fig4.update_layout(yaxis_title='Percentual (%)', xaxis_title='Faixa Etária')
else:
    fig4 = px.bar(title='4. Local de Ocorrência por Faixa Etária')
c4.plotly_chart(fig4, use_container_width=True)
