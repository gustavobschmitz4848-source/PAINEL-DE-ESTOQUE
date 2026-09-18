"""
Painel Integrado de Estoque - Temper Plus
Elaborado por: André Luiz & Gustavo Schmitz
Setor: Expedição & Logística
"""

import io
import re
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
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
LEAD_TIME_PADRAO_DIAS = 30
LEAD_TIME_POR_GRUPO = {"PERFIL DE ALUMINIO": 45, "GUARNIÇ": 20}
FATOR_SEGURANCA = 1.3
DIAS_COBERTURA_MAXIMO = 30

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
  }
  .kpi h4 { margin: 0; font-size: 0.8rem; color: #666; font-weight: 500;
             text-transform: uppercase; letter-spacing: 0.5px; }
  .kpi p  { margin: 4px 0 0 0; font-size: 1.5rem; font-weight: 700; color: #1a252f; }
  .kpi.danger  { border-left-color: #e53935; }
  .kpi.warn    { border-left-color: #fb8c00; }
  .kpi.ok      { border-left-color: #43a047; }
</style>
""", unsafe_allow_html=True)


# ============================================================
# PARSERS (convertem texto em número com segurança)
# ============================================================
def numero_br(texto):
    """Converte '1.234,56' ou '1234.56' em float."""
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
    """Converte 'R$ 24.896,00' → 24896.00"""
    if pd.isna(v):
        return 0.0
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v)
    return numero_br(str(v).replace("R$", "").strip())


def extrair_qtd_unidade(valor, grupo):
    """
    Devolve (quantidade, unidade).
    ─ Perfis de alumínio → (barras, 'br')  ← foco principal
    ─ Guarnições/borrachas/escovas → (metros, 'm')
    ─ Vidros → (m², 'm2')
    ─ Restante → (unidades, 'un')
    """
    if pd.isna(valor):
        return 0.0, "un"
    s = str(valor).strip()
    g = str(grupo).upper()

    # "1626ml (271 br)" → 271 br  ← caminho preferido
    m = re.search(r"\(([\d.,]+)\s*br\)", s, re.IGNORECASE)
    if m:
        return numero_br(m.group(1)), "br"

    # Perfil de alumínio sem anotação, mas com "ml" → divide por 6
    if "PERFIL DE ALUMINIO" in g and "ml" in s.lower():
        metros = numero_br(re.sub(r"ml", "", s, flags=re.IGNORECASE))
        return round(metros / COMPRIMENTO_BARRA_M, 2), "br"

    # Guarnição / borracha / escova → metro linear direto
    if "ml" in s.lower():
        return numero_br(re.sub(r"ml", "", s, flags=re.IGNORECASE)), "m"

    # Vidro → m²
    if "m²" in s or "m2" in s.lower():
        return numero_br(re.sub(r"m²|m2", "", s, flags=re.IGNORECASE)), "m2"

    # Peças / unidades
    return numero_br(s), "un"


def extrair_periodo_dias(df_raw):
    """Lê 'Data Início Sit.: dd/mm/aaaa ... Data Fim Sit.: dd/mm/aaaa'."""
    texto = " ".join(str(v) for v in df_raw.head(10).astype(str).values.flatten())
    datas = re.findall(r"(\d{2}/\d{2}/\d{4})", texto)
    if len(datas) >= 2:
        try:
            d1 = datetime.strptime(datas[0], "%d/%m/%Y")
            d2 = datetime.strptime(datas[1], "%d/%m/%Y")
            dias = (d2 - d1).days
            if dias > 0:
                return dias
        except ValueError:
            pass
    return 365


def lead_time_do_grupo(grupo):
    g = str(grupo).upper()
    for chave, dias in LEAD_TIME_POR_GRUPO.items():
        if chave.upper() in g:
            return dias
    return LEAD_TIME_PADRAO_DIAS


# ============================================================
# CARREGAMENTO DOS ARQUIVOS
# ============================================================
@st.cache_data(show_spinner="📥 Lendo estoque...")
def carregar_estoque(caminho):
    df = pd.read_excel(caminho, sheet_name=0, header=7)
    df = df[df["Cód."].notna()].copy()
    df = df[~df["Cód."].astype(str).isin(["Cód.", "TOTAL", "Totais"])].copy()

    df["Codigo"] = df["Cód."].astype(str).str.strip()
    df["Produto"] = df["Produto"].astype(str).str.strip()
    df["Grupo"] = df["Grupo/ Subgrupo"].astype(str).str.strip()

    disp = df.apply(
        lambda r: extrair_qtd_unidade(r.get("Disponível", 0), r["Grupo"]), axis=1
    )
    total = df.apply(
        lambda r: extrair_qtd_unidade(r.get("Qtd. em Estoque", 0), r["Grupo"])[0],
        axis=1,
    )

    df["Qtd_Disponivel"] = [v[0] for v in disp]
    df["Unidade"] = [v[1] for v in disp]
    df["Qtd_Total"] = total

    df["Custo_Total"] = df["Preço Total de Custo"].apply(limpar_moeda)
    df["Venda_Total"] = df["Preço Total de Venda"].apply(limpar_moeda)
    df["Custo_Unit"] = np.where(
        df["Qtd_Total"] > 0, df["Custo_Total"] / df["Qtd_Total"], 0.0
    )

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
    df["Qtde"] = df["Qtde"].apply(numero_br)
    df["Metros"] = df["Total M²"].apply(
        lambda v: numero_br(re.sub(r"ml", "", str(v), flags=re.IGNORECASE))
    )
    df["Vendido"] = df["Total Vendido"].apply(limpar_moeda)

    vendas = df.groupby("Codigo", as_index=False)[["Qtde", "Metros", "Vendido"]].sum()
    return vendas, dias


# ============================================================
# CÁLCULOS DO PAINEL
# ============================================================
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
    df["Consumo_Diario"] = df["Vendas_Periodo"] / dias
    df["Lead_Time"] = df["Grupo"].apply(lead_time_do_grupo)

    demanda = df["Consumo_Diario"] * df["Lead_Time"]
    df["Est_Minimo"] = (demanda * (FATOR_SEGURANCA - 1)).round(2)
    df["Ponto_Pedido"] = (demanda + df["Est_Minimo"]).round(2)
    df["Est_Maximo"] = (
        df["Ponto_Pedido"] + df["Consumo_Diario"] * DIAS_COBERTURA_MAXIMO
    ).round(2)

    df["Qtd_Comprar"] = np.where(
        df["Qtd_Disponivel"] <= df["Ponto_Pedido"],
        (df["Est_Maximo"] - df["Qtd_Disponivel"]).clip(lower=0),
        0,
    ).round(2)

    df["Autonomia"] = np.where(
        df["Consumo_Diario"] > 0,
        df["Qtd_Disponivel"] / df["Consumo_Diario"],
        np.nan,
    ).round(1)

    df["Custo_Comprar"] = (df["Qtd_Comprar"] * df["Custo_Unit"]).round(2)
    df["Valor_Imobilizado"] = (df["Qtd_Disponivel"] * df["Custo_Unit"]).round(2)
    df["Faturamento"] = (df["Vendido"] / dias * 365).round(2)

    def status(r):
        if r["Consumo_Diario"] == 0:
            return "⚪ Sem Movimento" if r["Qtd_Disponivel"] == 0 else "⚪ Parado"
        if r["Qtd_Disponivel"] <= r["Ponto_Pedido"]:
            return "🚨 COMPRAR"
        if r["Est_Maximo"] > 0 and r["Qtd_Disponivel"] > r["Est_Maximo"]:
            return "⚠️ SATURADO"
        return "✅ OK"

    df["Status"] = df.apply(status, axis=1)

    # Curva ABC por faturamento
    df = df.sort_values("Vendido", ascending=False).reset_index(drop=True)
    tot = df["Vendido"].sum()
    df["_pct"] = df["Vendido"].cumsum() / tot if tot > 0 else 1.0

    def curva(p, vendido):
        if vendido <= 0:
            return "D (Sem Giro)"
        if p <= 0.80:
            return "A (Alta Prioridade)"
        if p <= 0.95:
            return "B (Média Prioridade)"
        return "C (Baixa Prioridade)"

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

# ---------- Carrega os arquivos do repositório ----------
try:
    df_est = carregar_estoque(ARQ_ESTOQUE)
    df_vend, dias = carregar_vendas(ARQ_VENDAS)
except FileNotFoundError as e:
    st.error(
        f"❌ Arquivo não encontrado: {e.filename}. "
        "Verifique se `EstoqueReal.xls` e `vendasProd.xls` estão na raiz do repositório."
    )
    st.stop()

df = montar_painel(df_est, df_vend, dias)

st.caption(f"📅 Período de análise: {dias} dias • 🗂️ {len(df)} itens carregados")

# ---------- KPIs ----------
comprar = df[df["Status"] == "🚨 COMPRAR"]
saturado = df[df["Status"] == "⚠️ SATURADO"]
ok = df[df["Status"] == "✅ OK"]

k1, k2, k3, k4, k5 = st.columns(5)
k1.markdown(f'<div class="kpi"><h4>Total de Itens</h4><p>{len(df)}</p></div>', unsafe_allow_html=True)
k2.markdown(f'<div class="kpi danger"><h4>🚨 A Comprar</h4><p>{len(comprar)}</p></div>', unsafe_allow_html=True)
k3.markdown(f'<div class="kpi danger"><h4>💸 Investimento Compra</h4><p>R$ {comprar["Custo_Comprar"].sum():,.2f}</p></div>', unsafe_allow_html=True)
k4.markdown(f'<div class="kpi warn"><h4>⚠️ Saturados</h4><p>{len(saturado)}</p></div>', unsafe_allow_html=True)
k5.markdown(f'<div class="kpi warn"><h4>📉 Capital Imobilizado</h4><p>R$ {saturado["Valor_Imobilizado"].sum():,.2f}</p></div>', unsafe_allow_html=True)

st.markdown("---")

# ---------- Abas ----------
t1, t2, t3, t4, t5 = st.tabs([
    "📊 Visão Geral",
    "🚨 Comprar",
    "⚠️ Saturado",
    "✅ OK",
    "🔍 Consulta por Produto",
])

with t1:
    c1, c2 = st.columns(2)
    with c1:
        fig = px.pie(df, names="Status", hole=0.45,
                     title="Distribuição do Status Operacional",
                     color_discrete_sequence=px.colors.qualitative.Set2)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        top = df.sort_values("Valor_Imobilizado", ascending=False).head(10)
        fig = px.bar(top, x="Valor_Imobilizado", y="Produto", orientation="h",
                     title="Top 10 Capital Imobilizado (R$)",
                     color_discrete_sequence=["#ff7f0e"])
        fig.update_layout(yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("📋 Ranking por Curva ABC")
    abc = df["Curva_ABC"].value_counts().reset_index()
    abc.columns = ["Curva", "Itens"]
    st.dataframe(abc, use_container_width=True, hide_index=True)

with t2:
    st.subheader(f"🚨 Necessidade de Compra — {len(comprar)} itens")
    if len(comprar):
        cols = ["Codigo", "Produto", "Unidade", "Qtd_Disponivel",
                "Ponto_Pedido", "Qtd_Comprar", "Custo_Comprar",
                "Autonomia", "Curva_ABC"]
        st.dataframe(comprar[cols], use_container_width=True, hide_index=True)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            comprar[cols].to_excel(w, index=False, sheet_name="Comprar")
        st.download_button(
            "📥 Baixar Lista de Compras (Excel)",
            buf.getvalue(),
            file_name="Lista_Compras_TemperPlus.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.success("🎉 Nenhum item precisa de reposição no momento!")

with t3:
    st.subheader(f"⚠️ Estoque Saturado / Parado — {len(saturado)} itens")
    if len(saturado):
        cols = ["Codigo", "Produto", "Unidade", "Qtd_Disponivel",
                "Est_Maximo", "Valor_Imobilizado", "Autonomia", "Curva_ABC"]
        st.dataframe(saturado[cols].sort_values("Valor_Imobilizado", ascending=False),
                     use_container_width=True, hide_index=True)
    else:
        st.success("🎉 Nenhum item saturado!")

with t4:
    st.subheader(f"✅ Estoque Balanceado — {len(ok)} itens")
    cols = ["Codigo", "Produto", "Unidade", "Qtd_Disponivel",
            "Ponto_Pedido", "Est_Maximo", "Autonomia", "Curva_ABC"]
    st.dataframe(ok[cols], use_container_width=True, hide_index=True)

with t5:
    st.subheader("🔍 Consulta Individual")
    busca = st.text_input("Digite o código ou nome do produto:")
    base = df if not busca else df[
        df["Produto"].str.contains(busca, case=False, na=False)
        | df["Codigo"].str.contains(busca, case=False, na=False)
    ]
    st.dataframe(
        base[["Codigo", "Produto", "Unidade", "Qtd_Disponivel",
              "Ponto_Pedido", "Est_Maximo", "Status", "Autonomia",
              "Valor_Imobilizado", "Curva_ABC"]],
        use_container_width=True, hide_index=True,
    )

# ---------- Rodapé ----------
st.markdown("---")
st.markdown(
    "<div style='text-align:center; color:#6c757d; font-size:0.9rem;'>"
    "Temper Plus • Gestão Logística Integrada | "
    "Elaborado por <b>André Luiz</b> &amp; <b>Gustavo Schmitz</b></div>",
    unsafe_allow_html=True,
)
