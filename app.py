"""
Painel Integrado de Estoque - Temper Plus
Elaborado por: André Luiz & Gustavo Schmitz
Setor: Expedição & Logística
"""

import io
import re
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# CONFIGURAÇÕES
# ============================================================
st.set_page_config(
    page_title="Painel de Estoque - Temper Plus",
    page_icon="📦",
    layout="wide",
)

ARQ_ESTOQUE = "EstoqueReal.xls"
ARQ_VENDAS = "vendasProd.xls"
LOGO = "logo.png"

COMPRIMENTO_BARRA_M = 6.0

# ---------- CONFIGURAÇÃO POR GRUPO ----------
# ⚠️ VIDRO REMOVIDO conforme solicitado
GRUPOS_CONFIG = {
    "PERFIL DE ALUMINIO": {"lead_time": 52, "dias_seguranca": 10, "dias_cobertura": 30, "label": "Perfis de Alumínio",  "cor": "#1976d2"},
    "GUARNIÇ":            {"lead_time": 20, "dias_seguranca": 5,  "dias_cobertura": 20, "label": "Guarnições",          "cor": "#7b1fa2"},
    "FERRAGEM":           {"lead_time": 20, "dias_seguranca": 5,  "dias_cobertura": 20, "label": "Ferragens",           "cor": "#c2185b"},
    "SELANTE":            {"lead_time": 20, "dias_seguranca": 5,  "dias_cobertura": 20, "label": "Selantes",            "cor": "#00838f"},
    "KITS TERCEIROS":     {"lead_time": 20, "dias_seguranca": 5,  "dias_cobertura": 20, "label": "Kits de Terceiros",   "cor": "#ef6c00"},
    "KITS PRODU":         {"lead_time": 2,  "dias_seguranca": 1,  "dias_cobertura": 10, "label": "Kits de Produção",    "cor": "#2e7d32"},
    "KIT P/ MONTAGEM":    {"lead_time": 2,  "dias_seguranca": 1,  "dias_cobertura": 10, "label": "Kits p/ Montagem",    "cor": "#558b2f"},
}
PADRAO = {"lead_time": 20, "dias_seguranca": 5, "dias_cobertura": 20, "label": "Outros", "cor": "#616161"}

STATUS_CORES = {
    "🔴 URGENTE":       "#e53935",
    "🚨 COMPRAR":       "#fb8c00",
    "⚠️ SATURADO":      "#fdd835",
    "✅ OK":            "#43a047",
    "⚪ Parado":        "#9e9e9e",
    "⚪ Sem Movimento": "#bdbdbd",
}


def config_do_grupo(grupo: str) -> dict:
    g = str(grupo).upper()
    for chave, cfg in GRUPOS_CONFIG.items():
        if chave.upper() in g:
            return cfg
    return PADRAO


# ============================================================
# CSS
# ============================================================
st.markdown("""
<style>
  .header-box {
    background: linear-gradient(135deg, #0f2027 0%, #203a43 50%, #2c5364 100%);
    padding: 16px 24px; border-radius: 12px; color: white;
    box-shadow: 0 4px 12px rgba(0,0,0,0.15); margin-bottom: 20px;
  }
  .header-box h2 { margin: 0; font-size: 1.9rem; }
  .badge {
    background-color: rgba(255,255,255,0.18); padding: 4px 12px;
    border-radius: 20px; font-size: 0.9rem; font-weight: 600;
  }
  .kpi {
    background: #f8f9fa; padding: 14px 16px; border-radius: 10px;
    border-left: 5px solid #2c5364; box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    height: 100%;
  }
  .kpi h4 { margin: 0; font-size: 0.75rem; color: #666; font-weight: 600;
             text-transform: uppercase; letter-spacing: 0.6px; }
  .kpi p  { margin: 6px 0 0 0; font-size: 1.45rem; font-weight: 700; color: #1a252f; }
  .kpi small { color: #888; font-size: 0.75rem; }
  .kpi.danger { border-left-color: #e53935; }
  .kpi.warn   { border-left-color: #fb8c00; }
  .kpi.ok     { border-left-color: #43a047; }
  .kpi.info   { border-left-color: #1e88e5; }

  .status-card {
    padding: 16px 20px; border-radius: 10px; margin-bottom: 12px;
    border-left: 6px solid; background: #f8f9fa;
  }
  .status-card h4 { margin: 0 0 6px 0; font-size: 1rem; }
  .status-card p  { margin: 0; color: #555; font-size: 0.9rem; }
  .s-urgente  { border-left-color: #e53935; background: #ffebee; }
  .s-comprar  { border-left-color: #fb8c00; background: #fff3e0; }
  .s-saturado { border-left-color: #fdd835; background: #fffde7; }
  .s-ok       { border-left-color: #43a047; background: #e8f5e9; }
  .s-parado   { border-left-color: #9e9e9e; background: #f5f5f5; }
  .s-semmov   { border-left-color: #bdbdbd; background: #fafafa; }

  .formula-box {
    background: #f8f9fa; border-left: 4px solid #1e88e5;
    padding: 12px 16px; border-radius: 6px; margin: 8px 0;
    font-family: 'Courier New', monospace; font-size: 0.9rem;
  }
</style>
""", unsafe_allow_html=True)


# ============================================================
# PARSERS
# ============================================================
def numero_br(texto):
    if texto is None:
        return 0.0
    if isinstance(texto, float) and np.isnan(texto):
        return 0.0
    s = str(texto).strip()
    if not s or s.lower() in ("nan", "none"):
        return 0.0
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def limpar_moeda(v):
    if pd.isna(v):
        return 0.0
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v)
    return numero_br(str(v).replace("R$", "").strip())


def extrair_qtd_unidade(valor, grupo):
    if pd.isna(valor):
        return 0.0, "un"
    s = str(valor).strip()
    g = str(grupo).upper()

    m = re.search(r"\(([\d.,]+)\s*br\)", s, re.IGNORECASE)
    if m:
        return numero_br(m.group(1)), "br"

    if "PERFIL DE ALUMINIO" in g and "ml" in s.lower():
        metros = numero_br(re.sub(r"ml", "", s, flags=re.IGNORECASE))
        return round(metros / COMPRIMENTO_BARRA_M, 2), "br"

    if "ml" in s.lower():
        return numero_br(re.sub(r"ml", "", s, flags=re.IGNORECASE)), "m"

    if "m²" in s or "m2" in s.lower():
        return numero_br(re.sub(r"m²|m2", "", s, flags=re.IGNORECASE)), "m2"

    return numero_br(s), "un"


def extrair_periodo_dias(df_raw):
    texto = " ".join(str(v) for v in df_raw.head(10).astype(str).values.flatten())
    datas = re.findall(r"(\d{2}/\d{2}/\d{4})", texto)
    if len(datas) >= 2:
        try:
            d1 = datetime.strptime(datas[0], "%d/%m/%Y")
            d2 = datetime.strptime(datas[1], "%d/%m/%Y")
            if (d2 - d1).days > 0:
                return (d2 - d1).days
        except ValueError:
            pass
    return 365


# ============================================================
# CARREGAMENTO
# ============================================================
@st.cache_data(show_spinner="📥 Lendo estoque...")
def carregar_estoque(caminho):
    df = pd.read_excel(caminho, sheet_name=0, header=7)
    df = df[df["Cód."].notna()].copy()
    df = df[~df["Cód."].astype(str).isin(["Cód.", "TOTAL", "Totais"])].copy()

    df["Codigo"] = df["Cód."].astype(str).str.strip()
    df["Produto"] = df["Produto"].astype(str).str.strip()
    df["Grupo"] = df["Grupo/ Subgrupo"].astype(str).str.strip()

    # 🚫 Remove -M
    df = df[~df["Codigo"].str.upper().str.endswith("-M")].copy()

    disp = df.apply(lambda r: extrair_qtd_unidade(r.get("Disponível", 0), r["Grupo"]), axis=1)
    total = df.apply(lambda r: extrair_qtd_unidade(r.get("Qtd. em Estoque", 0), r["Grupo"])[0], axis=1)

    df["Qtd_Disponivel"] = [v[0] for v in disp]
    df["Unidade"] = [v[1] for v in disp]
    df["Qtd_Total"] = total

    df["Custo_Total"] = df["Preço Total de Custo"].apply(limpar_moeda)
    df["Venda_Total"] = df["Preço Total de Venda"].apply(limpar_moeda)
    df["Custo_Unit"] = np.where(df["Qtd_Total"] > 0, df["Custo_Total"] / df["Qtd_Total"], 0.0)

    return df[[
        "Codigo", "Produto", "Grupo", "Unidade",
        "Qtd_Disponivel", "Qtd_Total",
        "Custo_Total", "Venda_Total", "Custo_Unit",
    ]]


@st.cache_data(show_spinner="📥 Lendo vendas...")
def carregar_vendas(caminho):
    df_raw = pd.read_excel(caminho, sheet_name=0, header=None)
    dias = extrair_periodo_dias(df_raw)

    df = pd.read_excel(caminho, sheet_name=0, header=8)
    df = df[df["Cod."].notna()].copy()
    df = df[~df["Cod."].astype(str).isin(["Cod.", "TOTAL", "Totais"])].copy()

    df["Codigo"] = df["Cod."].astype(str).str.strip()
    df = df[~df["Codigo"].str.upper().str.endswith("-M")].copy()

    df["Qtde"] = df["Qtde"].apply(numero_br)
    df["Metros"] = df["Total M²"].apply(
        lambda v: numero_br(re.sub(r"ml", "", str(v), flags=re.IGNORECASE))
    )
    df["Vendido"] = df["Total Vendido"].apply(limpar_moeda)

    vendas = df.groupby("Codigo", as_index=False)[["Qtde", "Metros", "Vendido"]].sum()
    return vendas, dias


# ============================================================
# CÁLCULOS
# ============================================================
def gerar_folego(row):
    aut = row["Autonomia"]
    lt = row["Lead_Time"]
    giro = row["Giro_Diario"]
    status = row["Status"]

    if giro == 0:
        if row["Qtd_Disponivel"] == 0:
            return "Sem estoque e sem vendas no período"
        return f"Parado — {row['Qtd_Disponivel']:.0f} un. sem giro"

    if pd.isna(aut):
        return "—"

    if status == "🔴 URGENTE":
        if aut < lt:
            return f"⏰ Acaba em {aut:.0f} d — fornecedor leva {lt} d (⚠️ risco de ruptura)"
        return f"Acaba em {aut:.0f} d — abaixo do mínimo"
    if status == "🚨 COMPRAR":
        return f"Acaba em {aut:.0f} d — fornecedor leva {lt} d"
    if status == "⚠️ SATURADO":
        if aut > 365:
            return f"Parado há mais de 1 ano ({aut:.0f} d de estoque)"
        return f"Excesso — {aut:.0f} d de estoque parado"
    folga = aut - lt
    if folga > 30:
        return f"Confortável — {aut:.0f} d de autonomia (folga de {folga:.0f} d)"
    return f"OK — {aut:.0f} d de autonomia (folga de {folga:.0f} d)"


def montar_painel(df_est, df_vend, dias):
    df = df_est.merge(df_vend, on="Codigo", how="left")
    for c in ("Qtde", "Metros", "Vendido"):
        df[c] = df[c].fillna(0)

    def vendas_na_unidade(r):
        if r["Unidade"] == "br":
            return round(r["Metros"] / COMPRIMENTO_BARRA_M, 2)
        if r["Unidade"] == "m":
            return r["Metros"]
        return r["Qtde"]

    df["Vendas_Periodo"] = df.apply(vendas_na_unidade, axis=1)
    df["Giro_Diario"] = (df["Vendas_Periodo"] / dias).round(3)

    cfgs = df["Grupo"].apply(config_do_grupo)
    df["Lead_Time"] = [c["lead_time"] for c in cfgs]
    df["Dias_Seguranca"] = [c["dias_seguranca"] for c in cfgs]
    df["Dias_Cobertura"] = [c["dias_cobertura"] for c in cfgs]
    df["Grupo_Label"] = [c["label"] for c in cfgs]
    df["Grupo_Cor"] = [c["cor"] for c in cfgs]

    df["Estoque_Seguranca"] = (df["Giro_Diario"] * df["Dias_Seguranca"]).round(2)
    df["Estoque_Minimo"] = df["Estoque_Seguranca"]
    df["Ponto_Pedido"] = (
        df["Giro_Diario"] * df["Lead_Time"] + df["Estoque_Seguranca"]
    ).round(2)
    df["Estoque_Maximo"] = (
        df["Ponto_Pedido"] + df["Giro_Diario"] * df["Dias_Cobertura"]
    ).round(2)

    df["Qtd_Comprar"] = np.where(
        df["Qtd_Disponivel"] <= df["Ponto_Pedido"],
        (df["Estoque_Maximo"] - df["Qtd_Disponivel"]).clip(lower=0),
        0,
    ).round(2)

    df["Autonomia"] = np.where(
        df["Giro_Diario"] > 0,
        (df["Qtd_Disponivel"] / df["Giro_Diario"]).round(1),
        np.nan,
    )

    df["Custo_Comprar"] = (df["Qtd_Comprar"] * df["Custo_Unit"]).round(2)
    df["Valor_Imobilizado"] = (df["Qtd_Disponivel"] * df["Custo_Unit"]).round(2)
    df["Faturamento"] = (df["Vendido"] / dias * 365).round(2)

    hoje = datetime.today().date()
    df["Chegada_Se_Comprar"] = [
        (hoje + timedelta(days=int(lt))).strftime("%d/%m/%Y") for lt in df["Lead_Time"]
    ]

    def status(r):
        if r["Giro_Diario"] == 0:
            return "⚪ Sem Movimento" if r["Qtd_Disponivel"] == 0 else "⚪ Parado"
        if r["Qtd_Disponivel"] <= r["Estoque_Minimo"]:
            return "🔴 URGENTE"
        if r["Qtd_Disponivel"] <= r["Ponto_Pedido"]:
            return "🚨 COMPRAR"
        if r["Estoque_Maximo"] > 0 and r["Qtd_Disponivel"] > r["Estoque_Maximo"]:
            return "⚠️ SATURADO"
        return "✅ OK"

    df["Status"] = df.apply(status, axis=1)
    df["Fôlego"] = df.apply(gerar_folego, axis=1)

    # Curva ABC
    df = df.sort_values("Vendido", ascending=False).reset_index(drop=True)
    tot = df["Vendido"].sum()
    df["_pct"] = df["Vendido"].cumsum() / tot if tot > 0 else 1.0

    def curva(p, vendido):
        if vendido <= 0:
            return "D (Sem Giro)"
        if p <= 0.80:
            return "A (Alta)"
        if p <= 0.95:
            return "B (Média)"
        return "C (Baixa)"

    df["Curva_ABC"] = [curva(p, v) for p, v in zip(df["_pct"], df["Vendido"])]
    return df.drop(columns=["_pct"])


# ============================================================
# INTERFACE
# ============================================================
col_logo, col_titulo = st.columns([1, 6])
with col_logo:
    try:
        st.image(LOGO, width=140)
    except Exception:
        st.markdown("### 📦")

with col_titulo:
    st.markdown("""
    <div class="header-box">
      <h2>PAINEL INTEGRADO DE GESTÃO DE ESTOQUE</h2>
      <div style="display:flex; justify-content:space-between;
                  align-items:center; margin-top:8px;">
        <span>Temper Plus • Setor de Logística &amp; Expedição</span>
        <span class="badge">👨‍💻 Elaborado por: André Luiz &amp; Gustavo Schmitz</span>
      </div>
    </div>
    """, unsafe_allow_html=True)

try:
    df_est = carregar_estoque(ARQ_ESTOQUE)
    df_vend, dias = carregar_vendas(ARQ_VENDAS)
except FileNotFoundError as e:
    st.error(f"❌ Arquivo não encontrado: {e.filename}")
    st.stop()

df = montar_painel(df_est, df_vend, dias)

st.caption(
    f"📅 Período de análise: {dias} dias • 🗂️ {len(df)} itens • "
    f"🚫 Códigos '-M' ocultados • "
    f"🔄 Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}"
)

# ---------- FILTROS ----------
with st.sidebar:
    st.header("🔎 Filtros")

    grupos_disp = sorted(df["Grupo_Label"].unique())
    grupos_sel = st.multiselect("Grupo:", grupos_disp, default=grupos_disp)

    status_disp = ["🔴 URGENTE", "🚨 COMPRAR", "⚠️ SATURADO", "✅ OK",
                   "⚪ Parado", "⚪ Sem Movimento"]
    status_sel = st.multiselect("Status:", status_disp, default=status_disp)

    curva_disp = ["A (Alta)", "B (Média)", "C (Baixa)", "D (Sem Giro)"]
    curva_sel = st.multiselect("Curva ABC:", curva_disp, default=curva_disp)

    busca = st.text_input("🔍 Nome ou código:")

    st.markdown("---")
    st.caption("💡 Filtros afetam as abas Análise por Grupo e Consulta.")

dff = df[df["Grupo_Label"].isin(grupos_sel) & df["Status"].isin(status_sel) & df["Curva_ABC"].isin(curva_sel)]
if busca:
    dff = dff[dff["Produto"].str.contains(busca, case=False, na=False)
              | dff["Codigo"].str.contains(busca, case=False, na=False)]

# ---------- KPIs ----------
urgente = df[df["Status"] == "🔴 URGENTE"]
comprar = df[df["Status"] == "🚨 COMPRAR"]
saturado = df[df["Status"] == "⚠️ SATURADO"]
ok = df[df["Status"] == "✅ OK"]

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.markdown(f'<div class="kpi info"><h4>📦 Total de Itens</h4><p>{len(df)}</p></div>', unsafe_allow_html=True)
k2.markdown(f'<div class="kpi danger"><h4>🔴 Urgente</h4><p>{len(urgente)}</p><small>abaixo do mínimo</small></div>', unsafe_allow_html=True)
k3.markdown(f'<div class="kpi danger"><h4>🚨 Comprar</h4><p>{len(comprar)}</p><small>até o ponto de pedido</small></div>', unsafe_allow_html=True)
k4.markdown(f'<div class="kpi danger"><h4>💸 Investimento</h4><p>R$ {comprar["Custo_Comprar"].sum():,.0f}</p><small>para reposição</small></div>', unsafe_allow_html=True)
k5.markdown(f'<div class="kpi warn"><h4>⚠️ Saturados</h4><p>{len(saturado)}</p><small>acima do máximo</small></div>', unsafe_allow_html=True)
k6.markdown(f'<div class="kpi warn"><h4>📉 Capital Parado</h4><p>R$ {saturado["Valor_Imobilizado"].sum():,.0f}</p></div>', unsafe_allow_html=True)

st.markdown("---")

# ---------- ABAS ----------
t1, t2, t3, t4, t5, t6, t7 = st.tabs([
    "📊 Visão Geral",
    "🔴 Compras",
    "⚠️ Saturado",
    "📈 Análise por Grupo",
    "🔍 Consulta",
    "📖 Legenda & Lógica",
    "⚙️ Configuração",
])

# ============ ABA 1 ============
with t1:
    c1, c2 = st.columns(2)
    with c1:
        fig = px.pie(df, names="Status", hole=0.45,
                     title="Distribuição do Status Operacional",
                     color="Status", color_discrete_map=STATUS_CORES)
        fig.update_traces(textinfo="label+percent")
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        top = df[df["Valor_Imobilizado"] > 0].sort_values(
            "Valor_Imobilizado", ascending=False).head(10).copy()
        top["Label"] = top["Codigo"].str.slice(0, 20) + "  •  " + top["Produto"].str.slice(0, 32)
        fig = px.bar(top, x="Valor_Imobilizado", y="Label", orientation="h",
                     title="Top 10 — Maior Capital Imobilizado (R$)",
                     color="Grupo_Label",
                     color_discrete_map={c["label"]: c["cor"] for c in GRUPOS_CONFIG.values()} | {"Outros": "#616161"})
        fig.update_layout(yaxis={"categoryorder": "total ascending"}, showlegend=False)
        fig.update_traces(hovertemplate="<b>%{y}</b><br>R$ %{x:,.2f}<extra></extra>")
        st.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        top_giro = df[df["Giro_Diario"] > 0].sort_values(
            "Giro_Diario", ascending=False).head(10).copy()
        top_giro["Label"] = top_giro["Codigo"].str.slice(0, 20) + "  •  " + top_giro["Produto"].str.slice(0, 32)
        fig = px.bar(top_giro, x="Giro_Diario", y="Label", orientation="h",
                     title="Top 10 — Maior Giro Diário (un./dia)",
                     color="Grupo_Label",
                     color_discrete_map={c["label"]: c["cor"] for c in GRUPOS_CONFIG.values()} | {"Outros": "#616161"})
        fig.update_layout(yaxis={"categoryorder": "total ascending"}, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    with c4:
        abc = df["Curva_ABC"].value_counts().reset_index()
        abc.columns = ["Curva", "Itens"]
        fig = px.bar(abc, x="Curva", y="Itens",
                     title="Distribuição Curva ABC",
                     color="Curva",
                     color_discrete_sequence=px.colors.qualitative.Set2, text="Itens")
        st.plotly_chart(fig, use_container_width=True)

    # ---- NOVOS GRÁFICOS ----
    st.markdown("### 🗺️ Capital Imobilizado por Grupo e Status")
    df_tree = df[df["Valor_Imobilizado"] > 0]
    if len(df_tree):
        fig = px.treemap(
            df_tree, path=["Grupo_Label", "Status"], values="Valor_Imobilizado",
            color="Status", color_discrete_map=STATUS_CORES,
        )
        fig.update_traces(
            textinfo="label+value+percent parent",
            hovertemplate="<b>%{label}</b><br>R$ %{value:,.2f}<br>%{percentParent:.1%} do grupo<extra></extra>",
        )
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### 📉 Pareto da Curva ABC")
    top_par = df[df["Vendido"] > 0].head(30).copy()
    if len(top_par):
        top_par["Acum_%"] = (top_par["Vendido"].cumsum() / top_par["Vendido"].sum() * 100).round(1)
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=top_par["Codigo"], y=top_par["Vendido"],
            name="Faturamento", marker_color="#1e88e5",
            hovertemplate="<b>%{x}</b><br>R$ %{y:,.0f}<extra></extra>",
        ))
        fig.add_trace(go.Scatter(
            x=top_par["Codigo"], y=top_par["Acum_%"],
            name="Acumulado %", yaxis="y2",
            line=dict(color="#e53935", width=3), mode="lines+markers",
        ))
        fig.update_layout(
            title="Pareto — Top 30 SKUs por Faturamento",
            yaxis=dict(title="Faturamento (R$)"),
            yaxis2=dict(title="Acumulado %", overlaying="y", side="right",
                        range=[0, 105], showgrid=False),
            xaxis=dict(tickangle=-45),
            legend=dict(orientation="h", y=1.1),
        )
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### 🎯 Mapa de Risco — Giro × Autonomia")
    st.caption("Itens no canto inferior esquerdo (baixa autonomia com giro alto) são críticos.")
    df_sc = df[(df["Giro_Diario"] > 0) & (df["Autonomia"].notna())].copy()
    if len(df_sc):
        df_sc["Tam"] = df_sc["Valor_Imobilizado"].clip(lower=1)
        fig = px.scatter(
            df_sc, x="Giro_Diario", y="Autonomia",
            color="Status", color_discrete_map=STATUS_CORES,
            size="Tam", hover_data=["Codigo", "Produto", "Lead_Time"],
            title="Giro Diário × Autonomia (tamanho = capital em estoque)",
        )
        # Linha de referência do lead time por grupo — usa a moda
        lt_medio = df_sc["Lead_Time"].median()
        fig.add_hline(y=lt_medio, line_dash="dash", line_color="red",
                      annotation_text=f"Lead time médio ({lt_medio:.0f} d)",
                      annotation_position="right")
        fig.update_layout(xaxis_title="Giro Diário (un./dia)",
                          yaxis_title="Autonomia (dias)")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### ⚖️ Autonomia × Lead Time — Top 15 críticos")
    st.caption("Compara quantos dias o estoque dura com quantos dias o fornecedor leva. "
               "Barras vermelhas = risco de ruptura.")
    df_crit = df[(df["Giro_Diario"] > 0) & df["Autonomia"].notna()].copy()
    df_crit["Risco"] = (df_crit["Lead_Time"] - df_crit["Autonomia"]).round(1)
    df_crit = df_crit.sort_values("Risco", ascending=False).head(15)
    if len(df_crit):
        df_crit["Label"] = df_crit["Codigo"].str.slice(0, 18)
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=df_crit["Label"], y=df_crit["Autonomia"],
            name="Autonomia atual (d)", marker_color="#43a047",
        ))
        fig.add_trace(go.Bar(
            x=df_crit["Label"], y=df_crit["Lead_Time"],
            name="Lead time fornecedor (d)", marker_color="#e53935",
        ))
        fig.update_layout(
            barmode="group",
            title="Top 15 — Autonomia × Lead Time",
            xaxis_title="Código do produto", yaxis_title="Dias",
            legend=dict(orientation="h", y=1.1),
            xaxis_tickangle=-45,
        )
        st.plotly_chart(fig, use_container_width=True)

# ============ ABA 2 — COMPRAS ============
with t2:
    st.subheader(f"🔴 Itens que precisam de reposição — {len(urgente) + len(comprar)} itens")

    cols_mostrar = ["Codigo", "Produto", "Unidade", "Qtd_Disponivel",
                    "Estoque_Minimo", "Ponto_Pedido", "Estoque_Maximo",
                    "Qtd_Comprar", "Custo_Comprar", "Fôlego",
                    "Lead_Time", "Chegada_Se_Comprar"]
    renomear = {
        "Codigo": "Código", "Produto": "Produto", "Unidade": "Un.",
        "Qtd_Disponivel": "Disponível", "Estoque_Minimo": "Mínimo",
        "Ponto_Pedido": "Pto. Pedido", "Estoque_Maximo": "Máximo",
        "Qtd_Comprar": "Comprar", "Custo_Comprar": "Custo R$",
        "Fôlego": "Fôlego", "Lead_Time": "Lead (d)",
        "Chegada_Se_Comprar": "Chega em",
    }

    if len(urgente):
        st.markdown("#### 🔴 Urgente — Já abaixo do Estoque Mínimo")
        st.dataframe(urgente[cols_mostrar].rename(columns=renomear),
                     use_container_width=True, hide_index=True)

    if len(comprar):
        st.markdown("#### 🚨 Comprar — Atingiu o Ponto de Pedido")
        st.dataframe(comprar[cols_mostrar].rename(columns=renomear),
                     use_container_width=True, hide_index=True)

    export = pd.concat([urgente, comprar])
    if len(export):
        cols_exp = ["Codigo", "Produto", "Grupo_Label", "Unidade",
                    "Qtd_Disponivel", "Ponto_Pedido", "Qtd_Comprar",
                    "Custo_Comprar", "Lead_Time", "Chegada_Se_Comprar"]
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            export[cols_exp].to_excel(w, index=False, sheet_name="Comprar")
        st.download_button(
            "📥 Baixar Lista de Compras (Excel)",
            buf.getvalue(),
            file_name=f"Lista_Compras_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.success("🎉 Nenhum item precisa de reposição no momento!")

# ============ ABA 3 — SATURADO ============
with t3:
    st.subheader(f"⚠️ Estoque Saturado / Parado — {len(saturado)} itens")
    if len(saturado):
        st.dataframe(
            saturado[["Codigo", "Produto", "Grupo_Label", "Unidade",
                      "Qtd_Disponivel", "Estoque_Maximo",
                      "Fôlego", "Valor_Imobilizado", "Curva_ABC"]]
                .sort_values("Valor_Imobilizado", ascending=False)
                .rename(columns={
                    "Codigo": "Código", "Produto": "Produto",
                    "Grupo_Label": "Grupo", "Unidade": "Un.",
                    "Qtd_Disponivel": "Disponível", "Estoque_Maximo": "Máximo",
                    "Fôlego": "Fôlego",
                    "Valor_Imobilizado": "Capital R$",
                    "Curva_ABC": "Curva",
                }),
            use_container_width=True, hide_index=True,
        )
        st.metric("💸 Capital total imobilizado em excesso",
                  f"R$ {saturado['Valor_Imobilizado'].sum():,.2f}")
    else:
        st.success("🎉 Nenhum item saturado!")

# ============ ABA 4 — ANÁLISE POR GRUPO ============
with t4:
    st.subheader("📈 Análise Consolidada por Grupo")

    resumo = dff.groupby("Grupo_Label").agg(
        Itens=("Codigo", "count"),
        Giro_Total=("Giro_Diario", "sum"),
        Capital_Estoque=("Valor_Imobilizado", "sum"),
        Custo_Comprar=("Custo_Comprar", "sum"),
        Faturamento_Anual=("Faturamento", "sum"),
    ).reset_index().rename(columns={
        "Grupo_Label": "Grupo", "Giro_Total": "Giro Diário Total",
        "Capital_Estoque": "Capital em Estoque (R$)",
        "Custo_Comprar": "A Comprar (R$)",
        "Faturamento_Anual": "Faturamento Anual (R$)",
    })
    st.dataframe(resumo, use_container_width=True, hide_index=True)

    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(resumo, x="Grupo", y="Capital em Estoque (R$)",
                     title="Capital Imobilizado por Grupo",
                     color="Grupo", text_auto=".2s",
                     color_discrete_sequence=px.colors.qualitative.Set2)
        fig.update_layout(showlegend=False, xaxis_tickangle=-30)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.bar(resumo, x="Grupo", y="A Comprar (R$)",
                     title="Necessidade de Compra por Grupo",
                     color="Grupo", text_auto=".2s",
                     color_discrete_sequence=px.colors.qualitative.Pastel)
        fig.update_layout(showlegend=False, xaxis_tickangle=-30)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Status por Grupo")
    cross = pd.crosstab(dff["Grupo_Label"], dff["Status"])
    st.dataframe(cross, use_container_width=True)

# ============ ABA 5 — CONSULTA ============
with t5:
    st.subheader("🔍 Consulta Detalhada")
    st.caption(f"{len(dff)} itens (ajuste os filtros na barra lateral)")
    if len(dff):
        st.dataframe(
            dff[["Codigo", "Produto", "Grupo_Label", "Unidade",
                 "Qtd_Disponivel", "Giro_Diario", "Estoque_Minimo",
                 "Ponto_Pedido", "Estoque_Maximo", "Qtd_Comprar",
                 "Fôlego", "Status", "Curva_ABC", "Valor_Imobilizado"]]
                .rename(columns={
                    "Codigo": "Código", "Produto": "Produto",
                    "Grupo_Label": "Grupo", "Unidade": "Un.",
                    "Qtd_Disponivel": "Disponível", "Giro_Diario": "Giro/dia",
                    "Estoque_Minimo": "Mínimo", "Ponto_Pedido": "Pto. Pedido",
                    "Estoque_Maximo": "Máximo", "Qtd_Comprar": "Comprar",
                    "Fôlego": "Fôlego", "Status": "Status",
                    "Curva_ABC": "Curva", "Valor_Imobilizado": "Capital R$",
                }),
            use_container_width=True, hide_index=True, height=600,
        )

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            dff.to_excel(w, index=False, sheet_name="Consulta")
        st.download_button("📥 Baixar consulta filtrada",
                           buf.getvalue(),
                           file_name=f"Consulta_Estoque_{datetime.now().strftime('%Y%m%d')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("Nenhum item corresponde aos filtros.")

# ============ ABA 6 — LEGENDA & LÓGICA ============
with t6:
    st.subheader("📖 Legenda & Lógica do Painel")
    st.caption("Explicação em linguagem de negócio de cada indicador, status e regra usada.")

    st.markdown("### 🚦 O que significa cada Status")

    st.markdown("""
    <div class="status-card s-urgente">
      <h4>🔴 URGENTE — Abaixo do Estoque Mínimo</h4>
      <p>A quantidade disponível é <b>menor ou igual ao Estoque Mínimo</b>. 
      Isso significa que se um pedido grande entrar hoje, você pode não conseguir atender. 
      É o estado mais crítico — deve ser resolvido imediatamente.</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="status-card s-comprar">
      <h4>🚨 COMPRAR — Atingiu o Ponto de Pedido</h4>
      <p>A quantidade está acima do Mínimo, mas <b>já alcançou o Ponto de Pedido</b>. 
      É o gatilho ideal para disparar a compra: se você comprar agora, o material chega 
      antes do estoque acabar (considerando o Lead Time do fornecedor).</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="status-card s-saturado">
      <h4>⚠️ SATURADO — Acima do Estoque Máximo</h4>
      <p>Você tem <b>mais do que o necessário</b> para cobrir o consumo previsto. 
      Isso é dinheiro parado em prateleira. Itens saturados raramente precisam de compra 
      e devem ser monitorados para não virarem obsoletos.</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="status-card s-ok">
      <h4>✅ OK — Estoque Balanceado</h4>
      <p>Quantidade está entre o Ponto de Pedido e o Estoque Máximo. 
      É a <b>zona ideal</b>: nem risco de ruptura, nem capital parado em excesso.</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="status-card s-parado">
      <h4>⚪ PARADO — Tem estoque, mas não vendeu no período</h4>
      <p>Existe quantidade em estoque, mas <b>nenhuma venda foi registrada</b> no período analisado. 
      Não gera alerta de compra (não há consumo a repor), mas merece atenção: 
      é capital imobilizado que pode estar obsoleto.</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="status-card s-semmov">
      <h4>⚪ SEM MOVIMENTO — Zerado e sem vendas</h4>
      <p>Estoque zerado e nenhuma venda no período. O item está fora da operação — 
      normalmente é item descontinuado, sazonal ou cadastro inativo.</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 📊 Como funciona a Curva ABC")

    st.markdown("""
    A Curva ABC classifica os SKUs por **importância no faturamento** durante o período analisado.
    A ideia é aplicar a **Lei de Pareto** (80/20) ao estoque: a minoria dos itens gera a maioria do resultado.
    """)

    st.markdown("""
    <div class="formula-box">
    <b>Regra de classificação:</b><br>
    Ordena todos os SKUs por faturamento decrescente. Calcula o % acumulado.<br><br>
    🅰️ <b>Curva A (Alta)</b> — SKUs que, juntos, somam os <b>primeiros 80%</b> do faturamento.<br>
    &nbsp;&nbsp;&nbsp;&nbsp;Poucos itens, mas são <b>o coração do negócio</b>. Nunca podem faltar.<br><br>
    🅱️ <b>Curva B (Média)</b> — Itens que vão dos <b>80% até 95%</b> do faturamento.<br>
    &nbsp;&nbsp;&nbsp;&nbsp;Importância intermediária. Compra programada.<br><br>
    🅲 <b>Curva C (Baixa)</b> — Itens que vão dos <b>95% até 100%</b>.<br>
    &nbsp;&nbsp;&nbsp;&nbsp;Vendem pouco, mas ainda têm giro. Compra pontual.<br><br>
    🅳 <b>Curva D (Sem Giro)</b> — Itens que <b>não venderam nada</b> no período.<br>
    &nbsp;&nbsp;&nbsp;&nbsp;Ou estão descontinuados, ou são estoque morto.
    </div>
    """, unsafe_allow_html=True)

    st.info(
        "💡 **Como usar na prática:** a política de estoque muda por curva. "
        "Itens Curva A merecem estoque de segurança maior e monitoramento semanal. "
        "Itens Curva C podem trabalhar com estoque enxuto (compra sob demanda). "
        "Itens Curva D devem ser revisados para venda, descarte ou descontinuação."
    )

    st.markdown("---")
    st.markdown("### 📐 Fórmulas usadas nos cálculos")

    st.markdown("""
    <div class="formula-box">
    <b>Giro Diário</b> = Vendas do período ÷ Dias do período<br>
    <i>Quanto o item vende por dia, em média.</i>
    </div>

    <div class="formula-box">
    <b>Estoque de Segurança</b> = Giro Diário × Dias de Segurança (por grupo)<br>
    <i>Proteção contra atrasos de fornecedor e picos inesperados de venda.</i>
    </div>

    <div class="formula-box">
    <b>Estoque Mínimo</b> = Estoque de Segurança<br>
    <i>Quantidade mínima que nunca deveria faltar.</i>
    </div>

    <div class="formula-box">
    <b>Ponto de Pedido</b> = (Giro Diário × Lead Time) + Estoque de Segurança<br>
    <i>Quando o estoque chega aqui, é hora de comprar — o material chega antes de acabar.</i>
    </div>

    <div class="formula-box">
    <b>Estoque Máximo</b> = Ponto de Pedido + (Giro Diário × Dias de Cobertura)<br>
    <i>Limite superior. Acima disso, é estoque em excesso.</i>
    </div>

    <div class="formula-box">
    <b>Quantidade a Comprar</b> = Estoque Máximo − Estoque Atual<br>
    <i>Só é calculada quando o estoque atual está ≤ Ponto de Pedido.</i>
    </div>

    <div class="formula-box">
    <b>Autonomia</b> = Estoque Disponível ÷ Giro Diário<br>
    <i>Quantos dias o estoque atual dura no ritmo de venda.</i>
    </div>

    <div class="formula-box">
    <b>Fôlego</b> = Comparação entre Autonomia e Lead Time do fornecedor<br>
    <i>Traduz em texto se o estoque vai acabar antes ou depois do fornecedor entregar.</i>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 🕒 Lead Time por Grupo de Produto")

    conf = pd.DataFrame([
        {"Grupo": c["label"], "Lead Time (dias)": c["lead_time"],
         "Estoque Segurança (dias)": c["dias_seguranca"],
         "Dias Cobertura (após PP)": c["dias_cobertura"]}
        for c in GRUPOS_CONFIG.values()
    ])
    st.dataframe(conf, use_container_width=True, hide_index=True)

    st.info(
        "**Lead Time** é o tempo que o fornecedor leva para entregar o produto desde o pedido. "
        "**Dias de Segurança** é a margem extra de proteção. "
        "**Dias de Cobertura** é quanto de giro adicional entra no Estoque Máximo depois do Ponto de Pedido."
    )

# ============ ABA 7 — CONFIGURAÇÃO ============
with t7:
    st.subheader("⚙️ Parâmetros de Cálculo por Grupo")
    st.caption("Estes são os valores usados para calcular Mínimo, Ponto de Pedido e Máximo.")

    conf = pd.DataFrame([
        {"Grupo": c["label"], "Lead Time (dias)": c["lead_time"],
         "Estoque Segurança (dias)": c["dias_seguranca"],
         "Dias Cobertura (após PP)": c["dias_cobertura"]}
        for c in GRUPOS_CONFIG.values()
    ])
    st.dataframe(conf, use_container_width=True, hide_index=True)

    st.markdown("""
    **Fórmulas aplicadas:**
    - **Giro Diário** = vendas do período ÷ dias do período
    - **Estoque de Segurança** = Giro Diário × Dias de Segurança
    - **Estoque Mínimo** = Estoque de Segurança
    - **Ponto de Pedido** = (Giro Diário × Lead Time) + Estoque de Segurança
    - **Estoque Máximo** = Ponto de Pedido + (Giro Diário × Dias de Cobertura)
    - **Fôlego** = comparação entre Autonomia e Lead Time, em texto
    - **Qtd a Comprar** = Estoque Máximo − Qtd Disponível (só quando ≤ Ponto de Pedido)
    """)

# ---------- RODAPÉ ----------
st.markdown("---")
st.markdown(
    "<div style='text-align:center; color:#6c757d; font-size:0.9rem;'>"
    "Temper Plus • Gestão Logística Integrada | "
    "Elaborado por <b>André Luiz</b> &amp; <b>Gustavo Schmitz</b></div>",
    unsafe_allow_html=True,
)
