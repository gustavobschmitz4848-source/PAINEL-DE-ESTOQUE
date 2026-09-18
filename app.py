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

# ---------- PALETA TEMPER PLUS ----------
TP_AZUL_ESCURO = "#0F2557"
TP_AZUL_CLARO  = "#21A0B5"
TP_BRANCO      = "#FFFFFF"
TP_PRETO       = "#0D1117"
TP_CINZA_FUNDO = "#C9D3E0"   # Fundo mais escuro para destacar gráficos
TP_CINZA_BORDA = "#9FB0C4"

# Escala para gráficos contínuos
ESCALA_AZUL = [[0.0, "#7FB8C4"], [0.5, "#21A0B5"], [1.0, "#0F2557"]]

# ---------- GRUPOS ----------
GRUPOS_CONFIG = {
    "PERFIL DE ALUMINIO": {"lead_time": 52, "dias_seguranca": 10, "dias_cobertura": 30, "label": "Perfis de Alumínio",  "cor": TP_AZUL_ESCURO},
    "GUARNIÇ":            {"lead_time": 20, "dias_seguranca": 5,  "dias_cobertura": 20, "label": "Guarnições",          "cor": TP_AZUL_CLARO},
    "FERRAGEM":           {"lead_time": 20, "dias_seguranca": 5,  "dias_cobertura": 20, "label": "Ferragens",           "cor": "#2E5C8A"},
    "SELANTE":            {"lead_time": 20, "dias_seguranca": 5,  "dias_cobertura": 20, "label": "Selantes",            "cor": "#5BAABF"},
    "KITS TERCEIROS":     {"lead_time": 20, "dias_seguranca": 5,  "dias_cobertura": 20, "label": "Kits de Terceiros",   "cor": "#1B4F72"},
    "KITS PRODU":         {"lead_time": 2,  "dias_seguranca": 1,  "dias_cobertura": 10, "label": "Kits de Produção",    "cor": "#3D8B99"},
    "KIT P/ MONTAGEM":    {"lead_time": 2,  "dias_seguranca": 1,  "dias_cobertura": 10, "label": "Kits p/ Montagem",    "cor": "#7FB8C4"},
}
PADRAO = {"lead_time": 20, "dias_seguranca": 5, "dias_cobertura": 20, "label": "Outros", "cor": "#6C7A89"}

STATUS_CORES = {
    "🔴 URGENTE":        "#D32F2F",
    "🚨 COMPRAR":        "#F57C00",
    "⚠️ SATURADO":      "#F9A825",
    "✅ OK":            "#2E7D32",
    "⚪ Parado":        "#757575",
    "⚪ Sem Movimento": "#BDBDBD",
}

CURVA_CORES = {
    "A (Alta)":  TP_AZUL_ESCURO,
    "B (Média)": TP_AZUL_CLARO,
    "C (Baixa)": "#7FB8C4",
    "D (Sem Giro)": "#BDBDBD",
}


def config_do_grupo(grupo: str) -> dict:
    g = str(grupo).upper()
    for chave, cfg in GRUPOS_CONFIG.items():
        if chave.upper() in g:
            return cfg
    return PADRAO


def fmt_brl(v):
    """Formata número como R$ 1.234,56 (pt-BR)."""
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"


def fmt_num(v):
    """Formata número inteiro com ponto de milhar (pt-BR)."""
    try:
        return f"{int(v):,}".replace(",", ".")
    except Exception:
        return "0"


def fmt_pct(v):
    """Formata percentual em pt-BR."""
    try:
        return f"{float(v):.1f}%".replace(".", ",")
    except Exception:
        return "0,0%"


# ============================================================
# CSS ESTILIZADO
# ============================================================
st.markdown(f"""
<style>
  .stApp {{ background-color: {TP_CINZA_FUNDO}; }}

  .header-box {{
    background: linear-gradient(135deg, {TP_AZUL_ESCURO} 0%, #1B4A7A 55%, {TP_AZUL_CLARO} 100%);
    padding: 18px 26px; border-radius: 12px; color: {TP_BRANCO};
    box-shadow: 0 6px 18px rgba(15,37,87,0.25); margin-bottom: 16px;
    border-bottom: 4px solid {TP_AZUL_CLARO};
  }}
  .header-box h2 {{ margin: 0; font-size: 1.85rem; letter-spacing: 0.5px; }}
  .header-box span {{ font-size: 0.95rem; opacity: 0.92; }}
  .badge {{
    background-color: rgba(255,255,255,0.18);
    padding: 5px 14px; border-radius: 20px;
    font-size: 0.88rem; font-weight: 600;
    border: 1px solid rgba(255,255,255,0.25);
  }}

  .kpi {{
    background: {TP_BRANCO};
    padding: 14px 16px; border-radius: 10px;
    border-left: 5px solid {TP_AZUL_ESCURO};
    box-shadow: 0 2px 6px rgba(15,37,87,0.08);
    height: 100%;
  }}
  .kpi h4 {{
    margin: 0; font-size: 0.72rem; color: {TP_AZUL_ESCURO};
    font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px;
  }}
  .kpi p {{
    margin: 8px 0 0 0; font-size: 1.35rem;
    font-weight: 700; color: {TP_PRETO};
  }}
  .kpi small {{ color: #6C7A89; font-size: 0.72rem; }}
  .kpi.danger {{ border-left-color: #D32F2F; }}
  .kpi.warn   {{ border-left-color: #F57C00; }}
  .kpi.ok     {{ border-left-color: #2E7D32; }}
  .kpi.info   {{ border-left-color: {TP_AZUL_CLARO}; }}
  .kpi.gold   {{ border-left-color: #C9A227; }}

  .alert-box {{
    background: #FFF3E0; border-left: 6px solid #D32F2F;
    padding: 14px 20px; border-radius: 10px; margin-bottom: 16px;
    box-shadow: 0 2px 8px rgba(211,47,47,0.12);
  }}
  .alert-box h4 {{ margin: 0 0 8px 0; color: #B71C1C; font-size: 1rem; }}
  .alert-box p  {{ margin: 4px 0; color: #333; font-size: 0.9rem; }}
  .alert-box code {{
    background: #FFE0B2; padding: 2px 6px; border-radius: 4px;
    color: #B71C1C; font-weight: 600;
  }}

  .stTabs [data-baseweb="tab-list"] {{ gap: 4px; }}
  .stTabs [data-baseweb="tab"] {{
    background-color: {TP_BRANCO};
    border-radius: 8px 8px 0 0;
    padding: 10px 18px;
    color: {TP_AZUL_ESCURO};
    font-weight: 600;
    border: 1px solid {TP_CINZA_BORDA};
    border-bottom: none;
  }}
  .stTabs [aria-selected="true"] {{
    background-color: {TP_AZUL_ESCURO} !important;
    color: {TP_BRANCO} !important;
    border-color: {TP_AZUL_ESCURO} !important;
  }}

  .stButton button, .stDownloadButton button {{
    background-color: {TP_AZUL_CLARO};
    color: {TP_BRANCO};
    border: none; border-radius: 8px;
    font-weight: 600;
  }}
  .stButton button:hover, .stDownloadButton button:hover {{
    background-color: {TP_AZUL_ESCURO};
  }}

  section[data-testid="stSidebar"] {{
    background-color: #E4EAF2;
    border-right: 3px solid {TP_AZUL_CLARO};
  }}

  div[data-testid="stPlotlyChart"] {{
    background-color: {TP_BRANCO};
    border: 1px solid {TP_CINZA_BORDA};
    border-radius: 12px;
    padding: 10px;
    box-shadow: 0 3px 10px rgba(15,37,87,0.15);
  }}
</style>
""", unsafe_allow_html=True)


# ============================================================
# PARSERS & UTILITÁRIOS
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

    df = df[~df["Codigo"].str.upper().str.endswith("-M")].copy()

    disp = df.apply(lambda r: extrair_qtd_unidade(r.get("Disponível", 0), r["Grupo"]), axis=1)
    total = df.apply(lambda r: extrair_qtd_unidade(r.get("Qtd. em Estoque", 0), r["Grupo"])[0], axis=1)

    df["Qtd_Disponivel"] = [v[0] for v in disp]
    df["Unidade"] = [v[1] for v in disp]
    df["Qtd_Total"] = total

    df["Custo_Total"] = df["Preço Total de Custo"].apply(limpar_moeda)
    df["Venda_Total"] = df["Preço Total de Venda"].apply(limpar_moeda)
    df["Custo_Unit"] = np.where(df["Qtd_Total"] > 0, df["Custo_Total"] / df["Qtd_Total"], 0.0)
    df["Preco_Venda_Unit"] = np.where(df["Qtd_Total"] > 0, df["Venda_Total"] / df["Qtd_Total"], 0.0)

    return df[["Codigo", "Produto", "Grupo", "Unidade",
               "Qtd_Disponivel", "Qtd_Total",
               "Custo_Total", "Venda_Total", "Custo_Unit", "Preco_Venda_Unit"]]


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

    if "Lucro" in df.columns:
        df["Lucro"] = df["Lucro"].apply(limpar_moeda)
    else:
        df["Lucro"] = 0.0

    df["_pedido"] = 1

    vendas = df.groupby("Codigo", as_index=False).agg({
        "Qtde": "sum",
        "Metros": "sum",
        "Vendido": "sum",
        "Lucro": "sum",
        "_pedido": "sum",
    }).rename(columns={"_pedido": "N_Pedidos"})

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
    for c in ("Qtde", "Metros", "Vendido", "Lucro", "N_Pedidos"):
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

    df["Ruptura_Dias"] = np.where(
        df["Autonomia"].notna(),
        (df["Lead_Time"] - df["Autonomia"]).round(1),
        np.nan,
    )

    df["Custo_Comprar"] = (df["Qtd_Comprar"] * df["Custo_Unit"]).round(2)
    df["Valor_Imobilizado"] = (df["Qtd_Disponivel"] * df["Custo_Unit"]).round(2)

    df["Fat_Periodo"] = df["Vendido"].round(2)
    df["Fat_Mensal"]  = (df["Vendido"] / dias * 30).round(2)
    df["Fat_Anual"]   = (df["Vendido"] / dias * 365).round(2)

    df["Lucro_Periodo"] = df["Lucro"].round(2)
    df["Margem_%"] = np.where(
        df["Vendido"] > 0,
        (df["Lucro"] / df["Vendido"] * 100).round(1),
        np.nan,
    )

    df["Ticket_Medio"] = np.where(
        df["N_Pedidos"] > 0,
        (df["Vendido"] / df["N_Pedidos"]).round(2),
        0,
    )

    tot_vendido = df["Vendido"].sum()
    df["%_Fat_Total"] = np.where(
        tot_vendido > 0,
        (df["Vendido"] / tot_vendido * 100).round(2),
        0,
    )

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
# EXPORTAÇÃO EXCEL
# ============================================================
def exportar_excel_completo(df):
    buf = io.BytesIO()

    cols_compra = ["Codigo", "Produto", "Grupo_Label", "Unidade",
                   "Qtd_Disponivel", "Estoque_Minimo", "Ponto_Pedido",
                   "Qtd_Comprar", "Custo_Comprar", "Autonomia",
                   "Lead_Time", "Chegada_Se_Comprar", "Curva_ABC"]

    df_compra = df[df["Status"].isin(["🔴 URGENTE", "🚨 COMPRAR"])][cols_compra]
    df_sat    = df[df["Status"] == "⚠️ SATURADO"][cols_compra]
    df_todos  = df[cols_compra]

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_compra.to_excel(writer, index=False, sheet_name="Compras")
        df_sat.to_excel(writer, index=False, sheet_name="Saturado")
        df_todos.to_excel(writer, index=False, sheet_name="Todos")

        from openpyxl.styles import Alignment, Font, PatternFill
        fill = PatternFill("solid", fgColor="0F2557")
        font = Font(bold=True, color="FFFFFF")
        align = Alignment(horizontal="center", vertical="center")

        for ws in writer.book.worksheets:
            for cell in ws[1]:
                cell.fill = fill
                cell.font = font
                cell.alignment = align
            for col in ws.columns:
                max_len = max((len(str(c.value)) if c.value else 0) for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 42)

    return buf.getvalue()


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
    st.markdown(f"""
    <div class="header-box">
      <h2>PAINEL INTEGRADO DE GESTÃO DE ESTOQUE</h2>
      <div style="display:flex; justify-content:space-between; align-items:center; margin-top:8px;">
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
    f"📅 Período: {dias} dias • 🗂️ {len(df)} itens • "
    f"🚫 Códigos '-M' ocultados • "
    f"🔄 Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}"
)

# ============================================================
# BANNER DE ALERTAS CRÍTICOS
# ============================================================
risco = df[
    (df["Status"].isin(["🔴 URGENTE", "🚨 COMPRAR"]))
    & (df["Ruptura_Dias"] > 0)
    & (df["Giro_Diario"] > 0)
].sort_values("Ruptura_Dias", ascending=False).head(5)

if len(risco):
    itens_html = "".join(
        f"<p>• <code>{r['Codigo']}</code> <b>{str(r['Produto'])[:45]}</b> — "
        f"acaba em <b>{r['Autonomia']:.0f} dias</b>, fornecedor leva {int(r['Lead_Time'])} dias "
        f"→ <b>vai faltar {r['Ruptura_Dias']:.0f} dias</b></p>"
        for _, r in risco.iterrows()
    )
    st.markdown(f"""
    <div class="alert-box">
      <h4>🚨 {len(risco)} itens vão acabar ANTES do fornecedor entregar</h4>
      {itens_html}
    </div>
    """, unsafe_allow_html=True)

# ============================================================
# FILTROS LATERAIS
# ============================================================
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
    excel_data = exportar_excel_completo(df)
    st.download_button(
        label="📥 Baixar Relatório Completo (Excel)",
        data=excel_data,
        file_name=f"Relatorio_Estoque_TemperPlus_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

dff = df[df["Grupo_Label"].isin(grupos_sel) & df["Status"].isin(status_sel) & df["Curva_ABC"].isin(curva_sel)]
if busca:
    dff = dff[dff["Produto"].str.contains(busca, case=False, na=False)
              | dff["Codigo"].str.contains(busca, case=False, na=False)]

# ============================================================
# KPIS SUPERIORES
# ============================================================
urgente = df[df["Status"] == "🔴 URGENTE"]
comprar = df[df["Status"] == "🚨 COMPRAR"]
saturado = df[df["Status"] == "⚠️ SATURADO"]
ok = df[df["Status"] == "✅ OK"]

k1, k2, k3, k4, k5 = st.columns(5)
k1.markdown(f'<div class="kpi info"><h4>📦 Total de Itens</h4><p>{fmt_num(len(df))}</p></div>', unsafe_allow_html=True)
k2.markdown(f'<div class="kpi danger"><h4>🔴 Urgente</h4><p>{fmt_num(len(urgente))}</p><small>abaixo do mínimo</small></div>', unsafe_allow_html=True)
k3.markdown(f'<div class="kpi danger"><h4>🚨 Comprar</h4><p>{fmt_num(len(comprar))}</p><small>até o ponto de pedido</small></div>', unsafe_allow_html=True)
k4.markdown(f'<div class="kpi warn"><h4>⚠️ Saturados</h4><p>{fmt_num(len(saturado))}</p><small>acima do máximo</small></div>', unsafe_allow_html=True)
k5.markdown(f'<div class="kpi ok"><h4>✅ Estoque OK</h4><p>{fmt_num(len(ok))}</p><small>zona ideal</small></div>', unsafe_allow_html=True)

v1, v2, v3, v4, v5 = st.columns(5)
v1.markdown(f'<div class="kpi info"><h4>💰 Faturamento Período</h4><p>{fmt_brl(df["Fat_Periodo"].sum())}</p></div>', unsafe_allow_html=True)
v2.markdown(f'<div class="kpi info"><h4>📅 Faturamento Mensal</h4><p>{fmt_brl(df["Fat_Mensal"].sum())}</p></div>', unsafe_allow_html=True)
v3.markdown(f'<div class="kpi gold"><h4>💎 Lucro do Período</h4><p>{fmt_brl(df["Lucro_Periodo"].sum())}</p><small>margem: {fmt_pct((df["Lucro_Periodo"].sum() / df["Fat_Periodo"].sum() * 100) if df["Fat_Periodo"].sum() > 0 else 0)}</small></div>', unsafe_allow_html=True)
v4.markdown(f'<div class="kpi danger"><h4>💸 Investimento Compra</h4><p>{fmt_brl(comprar["Custo_Comprar"].sum())}</p><small>para repor</small></div>', unsafe_allow_html=True)
v5.markdown(f'<div class="kpi warn"><h4>📉 Capital Parado</h4><p>{fmt_brl(saturado["Valor_Imobilizado"].sum())}</p></div>', unsafe_allow_html=True)

st.markdown("---")

# ============================================================
# ABAS DA APLICAÇÃO
# ============================================================
t1, t2, t3, t4, t5, t6, t7 = st.tabs([
    "📊 Visão Geral",
    "🔴 Compras",
    "⚠️ Saturado",
    "📈 Análise por Grupo",
    "💎 Lucratividade",
    "🔍 Consulta",
    "📖 Legenda & Lógica",
])

# ============ ABA 1: VISÃO GERAL ============
with t1:
    c1, c2 = st.columns(2)
    with c1:
        fig = px.pie(df, names="Status", hole=0.45,
                     title="Distribuição do Status Operacional",
                     color="Status", color_discrete_map=STATUS_CORES)
        fig.update_traces(textinfo="label+percent")
        fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        top = df[df["Valor_Imobilizado"] > 0].sort_values(
            "Valor_Imobilizado", ascending=False).head(10).copy()
        top["Label"] = top["Codigo"].str.slice(0, 20) + "  •  " + top["Produto"].str.slice(0, 32)
        fig = px.bar(top, x="Valor_Imobilizado", y="Label", orientation="h",
                     title="Top 10 — Maior Capital Imobilizado (R$)",
                     color="Grupo_Label",
                     color_discrete_map={**{c["label"]: c["cor"] for c in GRUPOS_CONFIG.values()}, "Outros": "#6C7A89"})
        fig.update_layout(yaxis={"categoryorder": "total ascending"}, showlegend=False,
                          paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        fig.update_traces(hovertemplate="<b>%{y}</b><br>%{x:,.2f}<extra></extra>")
        st.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        top_giro = df[df["Giro_Diario"] > 0].sort_values(
            "Giro_Diario", ascending=False).head(10).copy()
        top_giro["Label"] = top_giro["Codigo"].str.slice(0, 20) + "  •  " + top_giro["Produto"].str.slice(0, 32)
        fig = px.bar(top_giro, x="Giro_Diario", y="Label", orientation="h",
                     title="Top 10 — Maior Giro Diário (un./dia)",
                     color="Grupo_Label",
                     color_discrete_map={**{c["label"]: c["cor"] for c in GRUPOS_CONFIG.values()}, "Outros": "#6C7A89"})
        fig.update_layout(yaxis={"categoryorder": "total ascending"}, showlegend=False,
                          paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)

    with c4:
        abc = df["Curva_ABC"].value_counts().reset_index()
        abc.columns = ["Curva", "Itens"]
        fig = px.bar(abc, x="Curva", y="Itens",
                     title="Distribuição Curva ABC",
                     color="Curva", text="Itens",
                     color_discrete_map=CURVA_CORES)
        fig.update_layout(showlegend=False, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)

# ============ ABA 2: COMPRAS ============
with t2:
    st.subheader(f"🔴 Itens que precisam de reposição — {len(urgente) + len(comprar)} itens")

    col_cfg = {
        "Código": st.column_config.TextColumn(width="small"),
        "Produto": st.column_config.TextColumn(width="large"),
        "Un.": st.column_config.TextColumn(width="small"),
        "Disponível": st.column_config.NumberColumn(format="%.0f"),
        "Mínimo": st.column_config.NumberColumn(format="%.0f"),
        "Pto. Pedido": st.column_config.NumberColumn(format="%.0f"),
        "Máximo": st.column_config.NumberColumn(format="%.0f"),
        "Comprar": st.column_config.NumberColumn(format="%.0f"),
        "Custo R$": st.column_config.NumberColumn(format="R$ %.2f"),
        "Lead (d)": st.column_config.NumberColumn(format="%d d"),
        "Chega em": st.column_config.TextColumn(),
        "Fôlego": st.column_config.TextColumn(width="large"),
    }

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
        st.dataframe(
            urgente[cols_mostrar].rename(columns=renomear),
            column_config=col_cfg,
            use_container_width=True, hide_index=True,
        )

    if len(comprar):
        st.markdown("#### 🚨 Comprar — Atingiu o Ponto de Pedido")
        st.dataframe(
            comprar[cols_mostrar].rename(columns=renomear),
            column_config=col_cfg,
            use_container_width=True, hide_index=True,
        )

# ============ ABA 3: SATURADO ============
with t3:
    st.subheader(f"⚠️ Itens com Estoque Acima do Máximo — {len(saturado)} itens")
    if len(saturado):
        st.dataframe(
            saturado[cols_mostrar].rename(columns=renomear),
            column_config=col_cfg,
            use_container_width=True, hide_index=True,
        )
    else:
        st.success("Nenhum item saturado no momento.")

# ============ ABA 4: ANÁLISE POR GRUPO ============
with t4:
    st.subheader("📈 Análise Consolidada por Grupo de Produto")
    grp = df.groupby("Grupo_Label").agg({
        "Codigo": "count",
        "Valor_Imobilizado": "sum",
        "Fat_Periodo": "sum",
        "Lucro_Periodo": "sum",
    }).reset_index().rename(columns={"Codigo": "Itens"})
    
    st.dataframe(grp, use_container_width=True, hide_index=True)

# ============ ABA 5: LUCRATIVIDADE ============
with t5:
    st.subheader("💎 Ranking de Lucratividade por SKU")
    top_lucro = df.sort_values("Lucro_Periodo", ascending=False).head(20)
    st.dataframe(
        top_lucro[["Codigo", "Produto", "Grupo_Label", "Fat_Periodo", "Lucro_Periodo", "Margem_%"]],
        use_container_width=True, hide_index=True
    )

# ============ ABA 6: CONSULTA ============
with t6:
    st.subheader("🔍 Consulta Geral Filtrada")
    st.dataframe(dff, use_container_width=True, hide_index=True)

# ============ ABA 7: LEGENDA & LÓGICA ============
with t7:
    st.markdown("""
    ### 📖 Regras de Negócio e Fórmulas Aplicadas
    
    * **Giro Diário:** `Vendas no Período / Dias do Período`
    * **Estoque de Segurança:** `Giro Diário * Dias de Segurança`
    * **Ponto de Pedido:** `(Giro Diário * Lead Time) + Estoque de Segurança`
    * **Estoque Máximo:** `Ponto de Pedido + (Giro Diário * Dias de Cobertura)`
    * **Autonomia:** `Estoque Disponível / Giro Diário`
    """)
