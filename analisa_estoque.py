import pandas as pd
import numpy as np
import plotly.express as px
import streamlit as st
import re

# Configuração de Página Streamlit
st.set_page_config(page_title="Análise de Estoque e Excessos", layout="wide")

# CSS Customizado para impedir que as métricas cortem
st.markdown("""
    <style>
    [data-testid="stMetricValue"] {
        font-size: 1.5rem !important;
        font-weight: bold !important;
        white-space: nowrap !important;
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.90rem !important;
    }
    </style>
""", unsafe_allow_html=True)

def extrair_disponivel(val):
    """
    Extrai a quantidade correta de barras ou unidades de textos como '2022ml (337 br)'
    """
    if pd.isna(val):
        return 0.0
    val_str = str(val).strip()
    
    # Busca por números dentro do parênteses (ex: '337 br')
    match_br = re.search(r'\((.*?)\)', val_str)
    if match_br:
        conteudo = match_br.group(1)
        num_match = re.search(r'[\d\.\,]+', conteudo)
        if num_match:
            num_s = num_match.group(0).replace('.', '').replace(',', '.')
            try:
                return float(num_s)
            except:
                pass
                
    # Caso não tenha parênteses, limpa sufixos de m², ml, etc.
    cleaned = re.sub(r'[^\d\,\.]', '', val_str)
    if not cleaned:
        return 0.0
    if ',' in cleaned:
        cleaned = cleaned.replace('.', '').replace(',', '.')
    try:
        return float(cleaned)
    except:
        return 0.0

def converter_moeda(val):
    """
    Limpa caracteres de moeda R$ e converte com precisão para float
    """
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace('R$', '').replace(' ', '').strip()
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except:
        return 0.0

@st.cache_data
def carregar_dados():
    df_raw = pd.read_excel('EstoqueReal.xls', sheet_name=0)
    df_header = df_raw.iloc[6]
    df = df_raw.iloc[8:].copy()
    df.columns = df_header
    df = df[df['Cód.'].notna()]
    df = df[~df['Cód.'].isin(['Cód.', 'TOTAL', 'Totais'])]
    
    # Tratamento das colunas numéricas
    df['Qtd_Disponivel'] = df['Disponível'].apply(extrair_disponivel)
    df['Valor_Custo_Total'] = df['Preço Total de Custo'].apply(converter_moeda)
    df['Valor_Venda_Total'] = df['Preço Total de Venda'].apply(converter_moeda)
    
    # Cálculo seguro do custo unitário por barra/unidade
    df['Custo_Unitario'] = np.where(
        df['Qtd_Disponivel'] > 0,
        df['Valor_Custo_Total'] / df['Qtd_Disponivel'],
        0.0
    )
    
    return df

df = carregar_dados()

st.title("📦 Painel de Análise de Estoque e Excessos")

# Top Cards de Métricas sem cortes
total_itens = len(df)
total_est_investido = df['Valor_Custo_Total'].sum()
total_pecas = df['Qtd_Disponivel'].sum()

col1, col2, col3 = st.columns(3)
col1.metric("Total de Itens Cadastrados", f"{total_itens:,}".replace(',', '.'))
col2.metric("Est. Investimento (Custo)", f"R$ {total_est_investido:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
col3.metric("Quantidade Total em Estoque", f"{total_pecas:,.0f}".replace(',', '.'))

st.markdown("---")

# Gráfico Top 10 Capital Imobilizado sem Notação Científica
top10 = df.sort_values(by='Valor_Custo_Total', ascending=False).head(10).copy()

fig = px.bar(
    top10,
    x='Valor_Custo_Total',
    y='Produto',
    orientation='h',
    text='Valor_Custo_Total',
    title="Top 10 Produtos com Maior Capital Imobilizado em Estoque (R$)"
)

# Força formatação monetária direta nos rótulos e eixos do Plotly
fig.update_traces(
    texttemplate='R$ %{x:,.2f}', 
    textposition='outside'
)
fig.update_layout(
    xaxis_title="Valor Imobilizado (R$)",
    yaxis_title="Produto",
    yaxis=dict(autorange="reversed"),
    xaxis=dict(showgrid=True, tickprefix="R$ "),
    height=500
)

st.plotly_chart(fig, use_container_width=True)