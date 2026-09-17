import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import re

# Configuração da Página
st.set_page_config(
    page_title="Painel Integrado de Estoque",
    page_icon="📦",
    layout="wide"
)

# Estilização CSS
st.markdown("""
    <style>
    .main { padding-top: 1rem; }
    .stMetric { background-color: #f8f9fa; padding: 12px; border-radius: 8px; border: 1px solid #e9ecef; }
    </style>
""", unsafe_allow_html=True)

# Topo: Logo + Título + Nome dos Elaboradores
col_logo, col_titulo = st.columns([1, 4])

with col_logo:
    try:
        st.image('logo.png', width=160)
    except:
        st.write("📦 **[LOGO]**")

with col_titulo:
    st.title("Painel Integrado de Gestão de Estoque")
    st.caption("Elaborado por: **EXPEDIÇÃO & LOGÍSTICA** | Sistema WGlass")

st.markdown("---")

@st.cache_data
def carregar_dados():
    # Carregar planilha de estoque
    df_raw = pd.read_excel('EstoqueReal.xls', sheet_name=0)
    df_header = df_raw.iloc[6]
    df_est = df_raw.iloc[8:].copy()
    df_est.columns = df_header
    df_est = df_est[df_est['Cód.'].notna()]
    df_est = df_est[~df_est['Cód.'].isin(['Cód.', 'TOTAL', 'Totais'])].copy()

    def extrair_num(val):
        if pd.isna(val): return 0.0
        val_str = str(val).strip()
        match_br = re.search(r'\((.*?)\)', val_str)
        if match_br:
            c = match_br.group(1)
            nm = re.search(r'[\d\.\,]+', c)
            if nm: return float(nm.group(0).replace('.', '').replace(',', '.'))
        cleaned = re.sub(r'[^\d\,\.]', '', val_str)
        if not cleaned: return 0.0
        if ',' in cleaned: cleaned = cleaned.replace('.', '').replace(',', '.')
        try: return float(cleaned)
        except: return 0.0

    def converter_moeda(val):
        if pd.isna(val): return 0.0
        if isinstance(val, (int, float)): return float(val)
        s = str(val).replace('R$', '').replace(' ', '').strip()
        if ',' in s: s = s.replace('.', '').replace(',', '.')
        try: return float(s)
        except: return 0.0

    df_est['Qtd_Disponivel'] = df_est['Disponível'].apply(extrair_num)
    df_est['Preco_Custo_Unit'] = df_est['Preço Custo'].apply(converter_moeda)
    df_est['Valor_Custo_Total'] = df_est['Preço Total de Custo'].apply(converter_moeda)

    # Classificação de Status
    def classificar_status(row):
        qtd = row['Qtd_Disponivel']
        if qtd <= 0:
            return '🚨 HORA DE COMPRAR (ZERADO)'
        elif qtd < 20:
            return '⚠️ ATENÇÃO (ESTOQUE BAIXO)'
        elif qtd > 1000:
            return '🔴 ESTOQUE SATURADO / EXCESSO'
        else:
            return '✅ ESTOQUE CERTO / IDEAL'

    df_est['Status_Estoque'] = df_est.apply(classificar_status, axis=1)
    return df_est

df = carregar_dados()

# Cards do Topo
m1, m2, m3, m4 = st.columns(4)
m1.metric("Total de Itens", f"{len(df):,}".replace(',', '.'))
m2.metric("Investimento Total", f"R$ {df['Valor_Custo_Total'].sum():,.2f}".replace('.', 'X').replace(',', '.').replace('X', ','))
m3.metric("Unidades em Estoque", f"{df['Qtd_Disponivel'].sum():,.0f}".replace(',', '.'))
m4.metric("Itens Saturados/Excesso", f"{len(df[df['Status_Estoque'].str.contains('SATURADO')]):,}".replace(',', '.'))

st.markdown("---")

# Abas de Navegação por Categoria
aba1, aba2, aba3, aba4, aba5 = st.tabs([
    "📊 Visão Geral", 
    "🔴 Estoque Saturado / Parado", 
    "🚨 Hora de Comprar", 
    "✅ Estoque Certo / Ideal",
    "🔍 Consulta por Produto"
])

with aba1:
    col_graf1, col_graf2 = st.columns(2)
    with col_graf1:
        st.subheader("Divisão dos Status do Estoque")
        df_status = df['Status_Estoque'].value_counts().reset_index()
        df_status.columns = ['Status', 'Quantidade']
        fig_pie = px.pie(df_status, names='Status', values='Quantidade', hole=0.4,
                         color_discrete_sequence=['#ff4b4b', '#ffa800', '#00c853', '#29b6f6'])
        st.plotly_chart(fig_pie, use_container_width=True)

    with col_graf2:
        st.subheader("Top 10 Capital Imobilizado (R$)")
        top10 = df.sort_values(by='Valor_Custo_Total', ascending=False).head(10)
        fig_bar = px.bar(top10, x='Valor_Custo_Total', y='Produto', orientation='h',
                         labels={'Valor_Custo_Total': 'Custo Total (R$)', 'Produto': 'Produto'},
                         color='Valor_Custo_Total', color_continuous_scale='Blues')
        fig_bar.update_layout(yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig_bar, use_container_width=True)

with aba2:
    st.subheader("🔴 Estoque Saturado e Capital Imobilizado")
    df_saturado = df[df['Status_Estoque'].str.contains('SATURADO')].sort_values(by='Valor_Custo_Total', ascending=False)
    st.dataframe(df_saturado[['Cód.', 'Produto', 'Qtd_Disponivel', 'Preco_Custo_Unit', 'Valor_Custo_Total']], use_container_width=True)

with aba3:
    st.subheader("🚨 Itens para Reposição Urgente (Hora de Comprar)")
    df_comprar = df[df['Status_Estoque'].str.contains('HORA DE COMPRAR|ATENÇÃO')].sort_values(by='Qtd_Disponivel', ascending=True)
    st.dataframe(df_comprar[['Cód.', 'Produto', 'Qtd_Disponivel', 'Preco_Custo_Unit', 'Valor_Custo_Total']], use_container_width=True)

with aba4:
    st.subheader("✅ Itens com Estoque Balanceado / Certo")
    df_certo = df[df['Status_Estoque'].str.contains('CERTO')].sort_values(by='Qtd_Disponivel', ascending=False)
    st.dataframe(df_certo[['Cód.', 'Produto', 'Qtd_Disponivel', 'Preco_Custo_Unit', 'Valor_Custo_Total']], use_container_width=True)

with aba5:
    st.subheader("🔍 Buscar Qualquer Produto no Sistema")
    busca = st.text_input("Digite o nome ou código do item:")
    if busca:
        df_filt = df[df['Produto'].astype(str).str.contains(busca, case=False) | df['Cód.'].astype(str).str.contains(busca, case=False)]
        st.dataframe(df_filt[['Cód.', 'Produto', 'Qtd_Disponivel', 'Status_Estoque', 'Valor_Custo_Total']], use_container_width=True)
    else:
        st.dataframe(df[['Cód.', 'Produto', 'Qtd_Disponivel', 'Status_Estoque', 'Valor_Custo_Total']], use_container_width=True)