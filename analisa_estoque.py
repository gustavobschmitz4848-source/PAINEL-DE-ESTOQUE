import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import re

st.set_page_config(
    page_title="Painel de Gestão e Giro de Estoque",
    page_icon="📦",
    layout="wide"
)

st.markdown("""
    <style>
    .main { padding-top: 1rem; }
    .stMetric { background-color: #f8f9fa; padding: 12px; border-radius: 8px; border: 1px solid #e9ecef; }
    </style>
""", unsafe_allow_html=True)

# Topo
col_logo, col_titulo = st.columns([1, 4])
with col_logo:
    try:
        st.image('logo.png', width=160)
    except:
        st.write("📦 **[LOGO]**")

with col_titulo:
    st.title("Painel de Gestão e Giro de Estoque - Temper Plus")
    st.caption("Elaborado por: **EXPEDIÇÃO & LOGÍSTICA** | Sistema WGlass")

st.markdown("---")

@st.cache_data
def carregar_dados():
    # 1. Carregar EstoqueReal
    df_raw = pd.read_excel('EstoqueReal.xls', header=None)
    idx_header = 6
    for idx in range(min(15, len(df_raw))):
        linha_texto = " ".join([str(v) for v in df_raw.iloc[idx].values if pd.notna(v)])
        if 'Cód' in linha_texto:
            idx_header = idx
            break

    df_est = df_raw.iloc[idx_header + 1:].copy()
    df_est.columns = [str(c).strip() for c in df_raw.iloc[idx_header].values]
    df_est = df_est.loc[:, ~df_est.columns.str.contains('nan|None', case=False, na=False)]

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

    def buscar_coluna(df, padrao):
        for col in df.columns:
            if padrao.lower() in str(col).lower():
                return col
        return None

    col_cod = buscar_coluna(df_est, 'Cód')
    col_disp = buscar_coluna(df_est, 'Disponível') or buscar_coluna(df_est, 'Disp')
    col_custo_unit = buscar_coluna(df_est, 'Preço Custo') or buscar_coluna(df_est, 'Custo')
    col_custo_tot = buscar_coluna(df_est, 'Preço Total de Custo') or buscar_coluna(df_est, 'Total')
    col_prod = buscar_coluna(df_est, 'Produto') or buscar_coluna(df_est, 'Descrição')

    df_est = df_est[df_est[col_cod].notna()].copy()
    df_est = df_est[~df_est[col_cod].astype(str).str.upper().isin(['CÓD.', 'CÓD', 'TOTAL', 'TOTAIS'])].copy()

    df_est['Cód.'] = df_est[col_cod].astype(str).str.strip()
    df_est['Qtd_Disponivel'] = df_est[col_disp].apply(extrair_num) if col_disp else 0.0
    df_est['Preco_Custo_Unit'] = df_est[col_custo_unit].apply(converter_moeda) if col_custo_unit else 0.0
    df_est['Valor_Custo_Total'] = df_est[col_custo_tot].apply(converter_moeda) if col_custo_tot else (df_est['Qtd_Disponivel'] * df_est['Preco_Custo_Unit'])
    df_est['Produto'] = df_est[col_prod] if col_prod else 'Sem Nome'

    # 2. Carregar Vendas (vendasProd.xls)
    try:
        df_v = pd.read_excel('vendasProd.xls')
        c_v_cod = buscar_coluna(df_v, 'Cód') or df_v.columns[0]
        c_v_qtd = buscar_coluna(df_v, 'Qtd') or buscar_coluna(df_v, 'Quantidade') or df_v.columns[1]
        
        df_v['Cód.'] = df_v[c_v_cod].astype(str).str.strip()
        df_v['Qtd_Vendida_Ano'] = df_v[c_v_qtd].apply(extrair_num)
        
        df_v_agrupado = df_v.groupby('Cód.')['Qtd_Vendida_Ano'].sum().reset_index()
        df = pd.merge(df_est, df_v_agrupado, on='Cód.', how='left')
        df['Qtd_Vendida_Ano'] = df['Qtd_Vendida_Ano'].fillna(0.0)
    except:
        df = df_est.copy()
        df['Qtd_Vendida_Ano'] = 0.0

    # 3. Cálculos de Giro
    df['Giro_Diario'] = df['Qtd_Vendida_Ano'] / 365.0
    df['Giro_Semanal'] = df['Qtd_Vendida_Ano'] / 52.0
    df['Giro_Mensal'] = df['Qtd_Vendida_Ano'] / 12.0

    # Dias de Cobertura de Estoque
    df['Dias_Cobertura'] = np.where(df['Giro_Diario'] > 0, df['Qtd_Disponivel'] / df['Giro_Diario'], 9999)

    # Classificação Analítica
    def classificar_movimento(row):
        if row['Qtd_Vendida_Ano'] <= 0 and row['Qtd_Disponivel'] > 0:
            return '🧊 Sem Movimento (1 Ano+)'
        elif row['Dias_Cobertura'] > 180 and row['Qtd_Disponivel'] > 0:
            return '🐢 Movimento Muito Baixo'
        elif row['Qtd_Disponivel'] <= 0:
            return '🚨 Hora de Comprar (Zerado)'
        else:
            return '⚡ Giro Normal / Ideal'

    df['Status_Giro'] = df.apply(classificar_movimento, axis=1)

    return df

df = carregar_dados()

# Cards
m1, m2, m3, m4 = st.columns(4)
m1.metric("Total de Produtos", f"{len(df):,}".replace(',', '.'))
m2.metric("Investimento Total", f"R$ {df['Valor_Custo_Total'].sum():,.2f}".replace('.', 'X').replace(',', '.').replace('X', ','))
m3.metric("Sem Movimento (1 Ano)", f"{len(df[df['Status_Giro'].str.contains('Sem Movimento')]):,}".replace(',', '.'))
m4.metric("Movimento Baixo (>180 dias)", f"{len(df[df['Status_Giro'].str.contains('Muito Baixo')]):,}".replace(',', '.'))

st.markdown("---")

# Abas Principais
aba1, aba2, aba3, aba4 = st.tabs([
    "📊 Visão Geral & Giro", 
    "🧊 Sem Movimento / Baixo Giro", 
    "📈 Análise Média de Giro (Dia/Sem/Mês)", 
    "🔍 Relatório Completo"
])

with aba1:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Análise de Movimentação do Estoque")
        fig_pie = px.pie(df, names='Status_Giro', hole=0.4,
                         color_discrete_sequence=['#ef5350', '#ffb74d', '#26a69a', '#ab47bc'])
        st.plotly_chart(fig_pie, use_container_width=True)

    with col2:
        st.subheader("Top 10 Capital Imobilizado (R$)")
        top10 = df.sort_values(by='Valor_Custo_Total', ascending=False).head(10)
        fig_bar = px.bar(top10, x='Valor_Custo_Total', y='Produto', orientation='h',
                         color='Valor_Custo_Total', color_continuous_scale='Blues',
                         labels={'Valor_Custo_Total': 'Valor em Estoque (R$)'})
        fig_bar.update_layout(yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig_bar, use_container_width=True)

with aba2:
    st.subheader("🧊 Produtos Parados Sem Movimento ou Com Baixo Giro")
    filtro_status = st.multiselect(
        "Filtrar Categoria:",
        options=df['Status_Giro'].unique(),
        default=['🧊 Sem Movimento (1 Ano+)', '🐢 Movimento Muito Baixo']
    )
    df_parado = df[df['Status_Giro'].isin(filtro_status)].sort_values(by='Valor_Custo_Total', ascending=False)
    
    st.write(f"**Total de Capital Imobilizado nos itens selecionados:** R$ {df_parado['Valor_Custo_Total'].sum():,.2f}".replace('.', 'X').replace(',', '.').replace('X', ','))
    
    st.dataframe(
        df_parado[['Cód.', 'Produto', 'Qtd_Disponivel', 'Qtd_Vendida_Ano', 'Dias_Cobertura', 'Valor_Custo_Total', 'Status_Giro']],
        use_container_width=True
    )

with aba3:
    st.subheader("📈 Média de Consumo/Giro por Item")
    st.write("Estimativa baseada no histórico de vendas do `vendasProd.xls`:")
    
    st.dataframe(
        df[['Cód.', 'Produto', 'Qtd_Disponivel', 'Giro_Diario', 'Giro_Semanal', 'Giro_Mensal', 'Dias_Cobertura']]
        .sort_values(by='Giro_Mensal', ascending=False),
        column_config={
            "Giro_Diario": st.column_config.NumberColumn("Média Diária", format="%.2f"),
            "Giro_Semanal": st.column_config.NumberColumn("Média Semanal", format="%.2f"),
            "Giro_Mensal": st.column_config.NumberColumn("Média Mensal", format="%.2f"),
            "Dias_Cobertura": st.column_config.NumberColumn("Cobertura (Dias)", format="%.0f d"),
        },
        use_container_width=True
    )

with aba4:
    st.subheader("🔍 Filtro Geral e Exportação")
    busca = st.text_input("Buscar Produto por Nome ou Código:")
    df_exp = df.copy()
    if busca:
        df_exp = df_exp[df_exp['Produto'].astype(str).str.contains(busca, case=False) | df_exp['Cód.'].astype(str).str.contains(busca, case=False)]
    
    st.dataframe(df_exp[['Cód.', 'Produto', 'Qtd_Disponivel', 'Qtd_Vendida_Ano', 'Giro_Mensal', 'Valor_Custo_Total', 'Status_Giro']], use_container_width=True)
