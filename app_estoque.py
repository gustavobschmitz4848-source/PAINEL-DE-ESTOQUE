import io
import os
import pandas as pd
import plotly.express as px
import streamlit as st

# Fixa o diretório de trabalho na pasta do projeto
PASTA_TRABALHO = r"C:\Users\Pc\OneDrive\Documentos\ESTOQUE"
os.chdir(PASTA_TRABALHO)

st.set_page_config(
    page_title="Painel e Gestão de Estoque - Temper Plus",
    layout="wide",
    page_icon="📊",
)

st.markdown(
    """
    <style>
        section[data-testid="stSidebar"] {
            background-color: #f8f9fa;
            border-right: 1px solid #dee2e6;
        }
        section[data-testid="stSidebar"] label, 
        section[data-testid="stSidebar"] span,
        section[data-testid="stSidebar"] p {
            color: #1a252f !important;
            font-weight: 700 !important;
        }
        .header-box {
            background: linear-gradient(135deg, #0f2027 0%, #203a43 50%, #2c5364 100%);
            padding: 15px 25px;
            border-radius: 10px;
            color: white;
            box-shadow: 0 4px 10px rgba(0,0,0,0.15);
            margin-bottom: 20px;
        }
        .badge-autor {
            background-color: rgba(255, 255, 255, 0.2);
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 0.95rem;
            font-weight: 600;
        }
    </style>
""",
    unsafe_allow_html=True,
)


@st.cache_data(ttl=5)
def carregar_estoque():
  path_relatorio = os.path.join(
      PASTA_TRABALHO, "Relatorio_Estoque_Avancado_WGlass.xlsx"
  )
  if os.path.exists(path_relatorio):
    return pd.read_excel(path_relatorio)
  return None


@st.cache_data(ttl=5)
def carregar_compras():
  path_compras = os.path.join(
      PASTA_TRABALHO, "Historico_Compras_Processado.xlsx"
  )
  if os.path.exists(path_compras):
    return pd.read_excel(path_compras)
  return None


df = carregar_estoque()
df_compras = carregar_compras()

# Cabeçalho
col_logo, col_titulo = st.columns([1, 5])
with col_logo:
  for logo_file in ["logo.png", "logo.jpg", "logo.png.jpg"]:
    if os.path.exists(os.path.join(PASTA_TRABALHO, logo_file)):
      st.image(os.path.join(PASTA_TRABALHO, logo_file), width=130)
      break

with col_titulo:
  st.markdown(
      """
    <div class="header-box">
        <h2 style="margin:0; padding:0; font-size: 2rem;">PAINEL E GESTÃO DE ESTOQUE</h2>
        <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 8px;">
            <span>Setor de Logística & Expedição • Temper Plus</span>
            <span class="badge-autor">👨‍💻 Elaborado por: ANDRÉ LUIZ & GUSTAVO SCHMITZ</span>
        </div>
    </div>
""",
      unsafe_allow_html=True,
  )

if df is None or "Custo_Estimado_Compra" not in df.columns:
  st.error(
      "A planilha atualizada não foi encontrada na pasta. Execute o comando"
      " 'python analisa_estoque.py' no Prompt de Comando."
  )
else:
  aba_logistica, aba_compras, aba_individual = st.tabs([
      "📊 Operação de Estoque & Alertas",
      "💰 Histórico de Compras Finalizadas",
      "🔍 Análise Detalhada por Item",
  ])

  with aba_logistica:
    comprar_df = df[df["Status"] == "🚨 COMPRAR"]
    saturado_df = df[df["Status"] == "⚠️ SATURADO"]

    custo_total_compra = comprar_df["Custo_Estimado_Compra"].sum()
    capital_saturado = saturado_df["Valor_Imobilizado_Estoque"].sum()

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("🚨 Necessitam Compra", f"{len(comprar_df)} itens")
    k2.metric(
        "💰 Est. Investimento Compra", f"R$ {custo_total_compra:,.2f}"
    )
    k3.metric("⚠️ Estoque Saturado", f"{len(saturado_df)} itens")
    k4.metric("📉 Capital Imobilizado", f"R$ {capital_saturado:,.2f}")
    k5.metric("📦 Total de Itens", f"{len(df)} itens")

    st.markdown("---")

    g1, g2 = st.columns(2)
    with g1:
      fig_status = px.pie(
          df,
          names="Status",
          title="Distribuição do Status Operacional",
          hole=0.45,
          color_discrete_sequence=px.colors.qualitative.Set1,
      )
      st.plotly_chart(fig_status, use_container_width=True)

    with g2:
      saturados_top = saturado_df.sort_values(
          by="Valor_Imobilizado_Estoque", ascending=False
      ).head(10)
      fig_sat = px.bar(
          saturados_top,
          x="Valor_Imobilizado_Estoque",
          y="Produto",
          orientation="h",
          title="Top 10 - Maior Capital Imobilizado em Excessos (R$)",
          color_discrete_sequence=["#ff7f0e"],
      )
      fig_sat.update_layout(yaxis={"categoryorder": "total ascending"})
      st.plotly_chart(fig_sat, use_container_width=True)

    st.markdown("---")

    # Filtros da Sidebar
    st.sidebar.header("🔍 Filtros de Busca")

    status_sel = st.sidebar.multiselect(
        "Filtrar por Status:",
        options=df["Status"].unique(),
        default=[s for s in df["Status"].unique() if s != "⚪ Sem Movimento"],
    )

    curva_sel = st.sidebar.multiselect(
        "Curva ABC (Prioridade):",
        options=df["Curva_ABC"].unique(),
        default=df["Curva_ABC"].unique(),
    )

    busca_txt = st.sidebar.text_input("Buscar por Nome ou Código:")

    df_f = df[(df["Status"].isin(status_sel)) & (df["Curva_ABC"].isin(curva_sel))]
    if busca_txt:
      df_f = df_f[
          df_f["Codigo"].astype(str).str.contains(busca_txt, case=False)
          | df_f["Produto"].astype(str).str.contains(busca_txt, case=False)
      ]

    col_t1, col_t2 = st.columns([4, 1])
    with col_t1:
      st.subheader(
          f"📋 Tabela de Controle de Estoque ({len(df_f)} itens exibidos)"
      )
    with col_t2:
      compras_export = df_f[df_f["Status"] == "🚨 COMPRAR"][
          ["Codigo", "Produto", "Grupo/ Subgrupo", "Qtd_Sugerida_Compra"]
      ]
      buffer = io.BytesIO()
      # USO DO OPENPYXL PARA EVITAR O ERRO
      with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        compras_export.to_excel(
            writer, index=False, sheet_name="Necessidade_Compra"
        )

      st.download_button(
          label="📥 Exportar Lista de Compras",
          data=buffer.getvalue(),
          file_name="Lista_Necessidade_Compras_TemperPlus.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      )

    st.dataframe(
        df_f.style.format({
            "Faturamento_Anual_R$": "R$ {:,.2f}",
            "Consumo_Diario": "{:.2f}",
            "Estoque_Atual": "{:.0f}",
            "Estoque_Minimo": "{:.0f}",
            "Ponto_Pedido": "{:.0f}",
            "Estoque_Maximo": "{:.0f}",
            "Qtd_Sugerida_Compra": "{:.0f}",
            "Autonomia_Dias": "{:.1f}",
            "Custo_Estimado_Compra": "R$ {:,.2f}",
            "Valor_Imobilizado_Estoque": "R$ {:,.2f}",
        }),
        use_container_width=True,
        hide_index=True,
    )

  with aba_compras:
    st.subheader("📊 Histórico de Compras Finalizadas")
    if df_compras is not None:
      c_tot1, c_tot2, c_tot3 = st.columns(3)
      c_tot1.metric(
          "Total Acumulado Comprado",
          f"R$ {df_compras['Total_Valor'].sum():,.2f}",
      )
      c_tot2.metric("Total de Pedidos", f"{len(df_compras)} pedidos")
      c_tot3.metric(
          "Ticket Médio / Pedido",
          f"R$ {df_compras['Total_Valor'].mean():,.2f}",
      )

      st.markdown("---")

      mc1, mc2 = st.columns(2)
      with mc1:
        top_forn = (
            df_compras.groupby("Fornecedor")["Total_Valor"]
            .sum()
            .reset_index()
            .sort_values(by="Total_Valor", ascending=False)
            .head(10)
        )
        fig_forn = px.bar(
            top_forn,
            x="Total_Valor",
            y="Fornecedor",
            orientation="h",
            title="Top 10 Fornecedores por Volume Comprado (R$)",
            color_discrete_sequence=["#2ca02c"],
        )
        fig_forn.update_layout(yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig_forn, use_container_width=True)

      with mc2:
        top_func = (
            df_compras.groupby("Funcionario")["Total_Valor"]
            .sum()
            .reset_index()
            .sort_values(by="Total_Valor", ascending=False)
        )
        fig_func = px.pie(
            top_func,
            names="Funcionario",
            values="Total_Valor",
            title="Volume Comprado por Responsável",
            hole=0.4,
        )
        st.plotly_chart(fig_func, use_container_width=True)

      st.dataframe(df_compras, use_container_width=True, hide_index=True)
    else:
      st.info("Nenhum histórico de compras processado.")

  with aba_individual:
    st.subheader("🔍 Consulta Individual de Produto")
    prod_sel = st.selectbox(
        "Selecione o Produto para Analisar:", options=df["Produto"].unique()
    )
    prod_data = df[df["Produto"] == prod_sel].iloc[0]

    p1, p2, p3, p4 = st.columns(4)
    p1.metric("Código", str(prod_data["Codigo"]))
    p2.metric("Status Operacional", str(prod_data["Status"]))
    p3.metric("Estoque Atual", f"{prod_data['Estoque_Atual']:.0f} un")
    p4.metric("Autonomia Estimada", f"{prod_data['Autonomia_Dias']:.1f} dias")

    st.markdown("---")
    st.markdown(f"**Grupo:** {prod_data['Grupo/ Subgrupo']}")
    st.markdown(f"**Curva ABC:** {prod_data['Curva_ABC']}")
    st.markdown(f"**Consumo Diário:** {prod_data['Consumo_Diario']:.2f} un/dia")
    st.markdown(f"**Ponto de Pedido:** {prod_data['Ponto_Pedido']:.0f} un")
    st.markdown(
        f"**Sugestão de Compra:** {prod_data['Qtd_Sugerida_Compra']:.0f} un"
    )

  st.markdown("---")
  st.markdown(
      "<div style='text-align: center; color: #6c757d; font-size: 0.9rem;'>Temper"
      " Plus • Gestão Logística Integrada | Elaborado por André Luiz & Gustavo"
      " Schmitz</div>",
      unsafe_allow_html=True,
  )