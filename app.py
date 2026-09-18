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
TP_CINZA_FUNDO = "#C9D3E0"   # fundo do site (mais escuro para destacar os gráficos)
TP_CINZA_BORDA = "#9FB0C4"

# Escala para gráficos contínuos (as barras claras continuam visíveis)
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
    "🔴 URGENTE":       "#D32F2F",
    "🚨 COMPRAR":       "#F57C00",
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
# CSS
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

  .status-card {{
    padding: 16px 20px; border-radius: 10px; margin-bottom: 12px;
    border-left: 6px solid; background: {TP_BRANCO};
    box-shadow: 0 2px 6px rgba(15,37,87,0.06);
  }}
  .status-card h4 {{ margin: 0 0 6px 0; font-size: 1rem; color: {TP_AZUL_ESCURO}; }}
  .status-card p  {{ margin: 0; color: #3A4A5C; font-size: 0.9rem; line-height: 1.5; }}
  .s-urgente  {{ border-left-color: #D32F2F; background: #FFEBEE; }}
  .s-comprar  {{ border-left-color: #F57C00; background: #FFF3E0; }}
  .s-saturado {{ border-left-color: #F9A825; background: #FFFDE7; }}
  .s-ok       {{ border-left-color: #2E7D32; background: #E8F5E9; }}
  .s-parado   {{ border-left-color: #757575; background: #F5F5F5; }}
  .s-semmov   {{ border-left-color: #BDBDBD; background: #FAFAFA; }}

  .formula-box {{
    background: {TP_BRANCO}; border-left: 4px solid {TP_AZUL_CLARO};
    padding: 12px 16px; border-radius: 6px; margin: 8px 0;
    font-family: 'Courier New', monospace; font-size: 0.9rem;
    color: {TP_AZUL_ESCURO};
    box-shadow: 0 1px 4px rgba(15,37,87,0.05);
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

  /* ---------- Sidebar ---------- */
  section[data-testid="stSidebar"] {{
    background-color: #E4EAF2;
    border-right: 3px solid {TP_AZUL_CLARO};
  }}
  section[data-testid="stSidebar"] h2,
  section[data-testid="stSidebar"] h3 {{ color: {TP_AZUL_ESCURO}; }}

  /* Tags dos filtros: azul da marca em vez de vermelho */
  span[data-baseweb="tag"] {{
    background-color: {TP_AZUL_ESCURO} !important;
    border-radius: 6px;
  }}
  span[data-baseweb="tag"] span,
  span[data-baseweb="tag"] svg {{
    color: {TP_BRANCO} !important;
    fill: {TP_BRANCO} !important;
  }}

  /* Campos de filtro com contorno visível */
  section[data-testid="stSidebar"] div[data-baseweb="select"] > div,
  section[data-testid="stSidebar"] input {{
    background-color: {TP_BRANCO};
    border: 1px solid {TP_CINZA_BORDA};
  }}

  /* ---------- Gráficos em cards brancos ---------- */
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

    # Lucro
    if "Lucro" in df.columns:
        df["Lucro"] = df["Lucro"].apply(limpar_moeda)
    else:
        df["Lucro"] = 0.0

    # Nº de pedidos (linhas de venda)
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

    # Ruptura potencial: quantos dias vai faltar (negativo = ok, positivo = vai faltar)
    df["Ruptura_Dias"] = np.where(
        df["Autonomia"].notna(),
        (df["Lead_Time"] - df["Autonomia"]).round(1),
        np.nan,
    )

    df["Custo_Comprar"] = (df["Qtd_Comprar"] * df["Custo_Unit"]).round(2)
    df["Valor_Imobilizado"] = (df["Qtd_Disponivel"] * df["Custo_Unit"]).round(2)

    # Faturamento em 3 perspectivas
    df["Fat_Periodo"] = df["Vendido"].round(2)
    df["Fat_Mensal"]  = (df["Vendido"] / dias * 30).round(2)
    df["Fat_Anual"]   = (df["Vendido"] / dias * 365).round(2)

    # Lucro e margem
    df["Lucro_Periodo"] = df["Lucro"].round(2)
    df["Margem_%"] = np.where(
        df["Vendido"] > 0,
        (df["Lucro"] / df["Vendido"] * 100).round(1),
        np.nan,
    )

    # Ticket médio por pedido (do item)
    df["Ticket_Medio"] = np.where(
        df["N_Pedidos"] > 0,
        (df["Vendido"] / df["N_Pedidos"]).round(2),
        0,
    )

    # % do faturamento total
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
# EXPORTAÇÃO EXCEL (com abas e formatação)
# ============================================================
def exportar_excel_completo(df):
    """Gera Excel com 3 abas: Compras, Saturado, Todos."""
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

        # Formatar cabeçalhos (fundo azul, texto branco, negrito)
        from openpyxl.styles import Font, PatternFill, Alignment
        fill = PatternFill("solid", fgColor="0F2557")
        font = Font(bold=True, color="FFFFFF")
        align = Alignment(horizontal="center", vertical="center")

        for ws in writer.book.worksheets:
            for cell in ws[1]:
                cell.fill = fill
                cell.font = font
                cell.alignment = align
            # Largura automática simples
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
# FILTROS
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
    st.caption("💡 Filtros afetam a aba Consulta.")

dff = df[df["Grupo_Label"].isin(grupos_sel) & df["Status"].isin(status_sel) & df["Curva_ABC"].isin(curva_sel)]
if busca:
    dff = dff[dff["Produto"].str.contains(busca, case=False, na=False)
              | dff["Codigo"].str.contains(busca, case=False, na=False)]

# ============================================================
# KPIs
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

# Linha 2 de KPIs — valores R$
v1, v2, v3, v4, v5 = st.columns(5)
v1.markdown(f'<div class="kpi info"><h4>💰 Faturamento Período</h4><p>{fmt_brl(df["Fat_Periodo"].sum())}</p></div>', unsafe_allow_html=True)
v2.markdown(f'<div class="kpi info"><h4>📅 Faturamento Mensal</h4><p>{fmt_brl(df["Fat_Mensal"].sum())}</p></div>', unsafe_allow_html=True)
v3.markdown(f'<div class="kpi gold"><h4>💎 Lucro do Período</h4><p>{fmt_brl(df["Lucro_Periodo"].sum())}</p><small>margem: {fmt_pct((df["Lucro_Periodo"].sum() / df["Fat_Periodo"].sum() * 100) if df["Fat_Periodo"].sum() > 0 else 0)}</small></div>', unsafe_allow_html=True)
v4.markdown(f'<div class="kpi danger"><h4>💸 Investimento Compra</h4><p>{fmt_brl(comprar["Custo_Comprar"].sum())}</p><small>para repor</small></div>', unsafe_allow_html=True)
v5.markdown(f'<div class="kpi warn"><h4>📉 Capital Parado</h4><p>{fmt_brl(saturado["Valor_Imobilizado"].sum())}</p></div>', unsafe_allow_html=True)

st.markdown("---")

# ============================================================
# ABAS
# ============================================================
t1, t2, t3, t4, t5, t6, t7, t8 = st.tabs([
    "📊 Visão Geral",
    "🔴 Compras",
    "⚠️ Saturado",
    "📈 Análise por Grupo",
    "💎 Lucratividade",
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

    st.markdown("### 🗺️ Capital Imobilizado por Grupo e Status")
    df_tree = df[df["Valor_Imobilizado"] > 0]
    if len(df_tree):
        fig = px.treemap(
            df_tree, path=["Grupo_Label", "Status"], values="Valor_Imobilizado",
            color="Status", color_discrete_map=STATUS_CORES,
        )
        fig.update_traces(
            textinfo="label+value+percent parent",
            hovertemplate="<b>%{label}</b><br>%{value:,.2f}<br>%{percentParent:.1%} do grupo<extra></extra>",
        )
        fig.update_layout(paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### 📉 Pareto — Top 30 SKUs por Faturamento")
    top_par = df[df["Vendido"] > 0].head(30).copy()
    if len(top_par):
        top_par["Acum_%"] = (top_par["Vendido"].cumsum() / top_par["Vendido"].sum() * 100).round(1)
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=top_par["Codigo"], y=top_par["Vendido"],
            name="Faturamento", marker_color=TP_AZUL_ESCURO,
            hovertemplate="<b>%{x}</b><br>R$ %{y:,.2f}<extra></extra>",
        ))
        fig.add_trace(go.Scatter(
            x=top_par["Codigo"], y=top_par["Acum_%"],
            name="Acumulado %", yaxis="y2",
            line=dict(color=TP_AZUL_CLARO, width=3), mode="lines+markers",
        ))
        fig.update_layout(
            yaxis=dict(title="Faturamento (R$)"),
            yaxis2=dict(title="Acumulado %", overlaying="y", side="right",
                        range=[0, 105], showgrid=False),
            xaxis=dict(tickangle=-45),
            legend=dict(orientation="h", y=1.1),
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig, use_container_width=True)

# ============ ABA 2 — COMPRAS ============
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
        "Ruptura (d)": st.column_config.NumberColumn(
            format="%d d",
            help="Dias que o estoque vai faltar. > 0 = vai faltar antes do fornecedor entregar.",
        ),
        "Fôlego": st.column_config.TextColumn(width="large"),
    }

    cols_mostrar = ["Codigo", "Produto", "Unidade", "Qtd_Disponivel",
                    "Estoque_Minimo", "Ponto_Pedido", "Estoque_Maximo",
                    "Qtd_Comprar", "Custo_Comprar", "Fôlego",
                    "Lead_Time", "Chegada_Se_Comprar", "Ruptura_Dias"]
    renomear = {
        "Codigo": "Código", "Produto": "Produto", "Unidade": "Un.",
        "Qtd_Disponivel": "Disponível", "Estoque_Minimo": "Mínimo",
        "Ponto_Pedido": "Pto. Pedido", "Estoque_Maximo": "Máximo",
        "Qtd_Comprar": "Comprar", "Custo_Comprar": "Custo R$",
        "Fôlego": "Fôlego", "Lead_Time": "Lead (d)",
        "Chegada_Se_Comprar": "Chega em", "Ruptura_Dias": "Ruptura (d)",
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

    export = pd.concat([urgente, comprar])
    if len(export):
        st.download_button(
            "📥 Baixar Excel completo (Compras + Saturado + Todos)",
            exportar_excel_completo(df),
            file_name=f"Estoque_TemperPlus_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.success("🎉 Nenhum item precisa de reposição no momento!")

# ============ ABA 3 — SATURADO ============
with t3:
    st.subheader(f"⚠️ Estoque Saturado / Parado — {len(saturado)} itens")
    if len(saturado):
        col_sat = {
            "Código": st.column_config.TextColumn(width="small"),
            "Produto": st.column_config.TextColumn(width="large"),
            "Grupo": st.column_config.TextColumn(),
            "Un.": st.column_config.TextColumn(width="small"),
            "Disponível": st.column_config.NumberColumn(format="%.0f"),
            "Máximo": st.column_config.NumberColumn(format="%.0f"),
            "Auton. (d)": st.column_config.NumberColumn(format="%d d"),
            "Capital R$": st.column_config.NumberColumn(format="R$ %.2f"),
            "Curva": st.column_config.TextColumn(width="small"),
            "Fôlego": st.column_config.TextColumn(width="large"),
        }
        df_sat_show = saturado[["Codigo", "Produto", "Grupo_Label", "Unidade",
                                 "Qtd_Disponivel", "Estoque_Maximo", "Autonomia",
                                 "Valor_Imobilizado", "Curva_ABC", "Fôlego"]].sort_values(
            "Valor_Imobilizado", ascending=False).rename(columns={
                "Codigo": "Código", "Produto": "Produto",
                "Grupo_Label": "Grupo", "Unidade": "Un.",
                "Qtd_Disponivel": "Disponível", "Estoque_Maximo": "Máximo",
                "Autonomia": "Auton. (d)", "Valor_Imobilizado": "Capital R$",
                "Curva_ABC": "Curva", "Fôlego": "Fôlego",
            })
        st.dataframe(df_sat_show, column_config=col_sat,
                     use_container_width=True, hide_index=True)
        st.metric("💸 Capital total imobilizado em excesso",
                  fmt_brl(saturado["Valor_Imobilizado"].sum()))
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
        Faturamento=("Fat_Periodo", "sum"),
        Lucro=("Lucro_Periodo", "sum"),
    ).reset_index()

    resumo["Margem_%"] = np.where(
        resumo["Faturamento"] > 0,
        (resumo["Lucro"] / resumo["Faturamento"] * 100).round(1),
        0,
    )

    resumo_show = resumo.rename(columns={
        "Grupo_Label": "Grupo", "Giro_Total": "Giro/dia",
        "Capital_Estoque": "Capital R$",
        "Custo_Comprar": "A Comprar R$",
        "Faturamento": "Faturamento R$",
        "Lucro": "Lucro R$",
        "Margem_%": "Margem %",
    })

    st.dataframe(
        resumo_show,
        column_config={
            "Giro/dia": st.column_config.NumberColumn(format="%.1f"),
            "Capital R$": st.column_config.NumberColumn(format="R$ %.2f"),
            "A Comprar R$": st.column_config.NumberColumn(format="R$ %.2f"),
            "Faturamento R$": st.column_config.NumberColumn(format="R$ %.2f"),
            "Lucro R$": st.column_config.NumberColumn(format="R$ %.2f"),
            "Margem %": st.column_config.NumberColumn(format="%.1f%%"),
        },
        use_container_width=True, hide_index=True,
    )

    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(resumo, x="Grupo_Label", y="Capital_Estoque",
                     title="Capital Imobilizado por Grupo",
                     color="Grupo_Label", text_auto=".2s",
                     color_discrete_map={**{c["label"]: c["cor"] for c in GRUPOS_CONFIG.values()}, "Outros": "#6C7A89"})
        fig.update_layout(showlegend=False, xaxis_tickangle=-30,
                          paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.bar(resumo, x="Grupo_Label", y="Faturamento",
                     title="Faturamento por Grupo (período)",
                     color="Grupo_Label", text_auto=".2s",
                     color_discrete_map={**{c["label"]: c["cor"] for c in GRUPOS_CONFIG.values()}, "Outros": "#6C7A89"})
        fig.update_layout(showlegend=False, xaxis_tickangle=-30,
                          paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Status por Grupo")
    cross = pd.crosstab(dff["Grupo_Label"], dff["Status"])
    st.dataframe(cross, use_container_width=True)

# ============ ABA 5 — LUCRATIVIDADE ============
with t5:
    st.subheader("💎 Ranking por Lucratividade")
    st.caption("Itens que mais contribuem para o lucro no período.")

    top_lucro = df[df["Lucro_Periodo"] > 0].sort_values(
        "Lucro_Periodo", ascending=False).head(20).copy()

    if len(top_lucro):
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### Top 20 — Maior Lucro (R$)")
            fig = px.bar(top_lucro, x="Lucro_Periodo", y="Codigo", orientation="h",
                         color="Margem_%", color_continuous_scale=ESCALA_AZUL,
                         labels={"Lucro_Periodo": "Lucro (R$)", "Codigo": "Código"})
            fig.update_layout(yaxis={"categoryorder": "total ascending"},
                              paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            fig.update_traces(hovertemplate="<b>%{y}</b><br>Lucro: R$ %{x:,.2f}<extra></extra>")
            st.plotly_chart(fig, use_container_width=True)

        with c2:
            st.markdown("#### Top 20 — Maior Margem (%)")
            top_margem = df[(df["Vendido"] > 1000) & df["Margem_%"].notna()].sort_values(
                "Margem_%", ascending=False).head(20).copy()
            fig = px.bar(top_margem, x="Margem_%", y="Codigo", orientation="h",
                         color="Margem_%", color_continuous_scale=ESCALA_AZUL,
                         labels={"Margem_%": "Margem (%)", "Codigo": "Código"})
            fig.update_layout(yaxis={"categoryorder": "total ascending"},
                              paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            fig.update_traces(hovertemplate="<b>%{y}</b><br>Margem: %{x:.1f}%<extra></extra>")
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### Tabela Detalhada — Top 30 por Lucro")
        cols_lucro = ["Codigo", "Produto", "Unidade", "Vendas_Periodo",
                      "Fat_Periodo", "Lucro_Periodo", "Margem_%",
                      "Ticket_Medio", "%_Fat_Total", "Curva_ABC"]
        st.dataframe(
            top_lucro.head(30)[cols_lucro].rename(columns={
                "Codigo": "Código", "Produto": "Produto", "Unidade": "Un.",
                "Vendas_Periodo": "Vendas", "Fat_Periodo": "Faturamento R$",
                "Lucro_Periodo": "Lucro R$", "Margem_%": "Margem %",
                "Ticket_Medio": "Ticket Médio R$", "%_Fat_Total": "% do Total",
                "Curva_ABC": "Curva",
            }),
            column_config={
                "Vendas": st.column_config.NumberColumn(format="%.0f"),
                "Faturamento R$": st.column_config.NumberColumn(format="R$ %.2f"),
                "Lucro R$": st.column_config.NumberColumn(format="R$ %.2f"),
                "Margem %": st.column_config.NumberColumn(format="%.1f%%"),
                "Ticket Médio R$": st.column_config.NumberColumn(format="R$ %.2f"),
                "% do Total": st.column_config.NumberColumn(format="%.2f%%"),
            },
            use_container_width=True, hide_index=True,
        )
    else:
        st.info("Nenhum item com lucro registrado no período.")

# ============ ABA 6 — CONSULTA ============
with t6:
    st.subheader("🔍 Consulta Detalhada")
    st.caption(f"{len(dff)} itens (ajuste os filtros na barra lateral)")
    if len(dff):
        cols = ["Codigo", "Produto", "Grupo_Label", "Unidade",
                "Qtd_Disponivel", "Giro_Diario", "Estoque_Minimo",
                "Ponto_Pedido", "Estoque_Maximo", "Qtd_Comprar",
                "Autonomia", "Ruptura_Dias", "Fôlego", "Status",
                "Curva_ABC", "Fat_Periodo", "Lucro_Periodo",
                "Margem_%", "Valor_Imobilizado"]

        st.dataframe(
            dff[cols].rename(columns={
                "Codigo": "Código", "Produto": "Produto",
                "Grupo_Label": "Grupo", "Unidade": "Un.",
                "Qtd_Disponivel": "Disponível", "Giro_Diario": "Giro/dia",
                "Estoque_Minimo": "Mínimo", "Ponto_Pedido": "Pto. Pedido",
                "Estoque_Maximo": "Máximo", "Qtd_Comprar": "Comprar",
                "Autonomia": "Auton. (d)", "Ruptura_Dias": "Ruptura (d)",
                "Fôlego": "Fôlego", "Status": "Status",
                "Curva_ABC": "Curva", "Fat_Periodo": "Faturamento R$",
                "Lucro_Periodo": "Lucro R$", "Margem_%": "Margem %",
                "Valor_Imobilizado": "Capital R$",
            }),
            column_config={
                "Disponível": st.column_config.NumberColumn(format="%.0f"),
                "Giro/dia": st.column_config.NumberColumn(format="%.2f"),
                "Mínimo": st.column_config.NumberColumn(format="%.0f"),
                "Pto. Pedido": st.column_config.NumberColumn(format="%.0f"),
                "Máximo": st.column_config.NumberColumn(format="%.0f"),
                "Comprar": st.column_config.NumberColumn(format="%.0f"),
                "Auton. (d)": st.column_config.NumberColumn(format="%d d"),
                "Ruptura (d)": st.column_config.NumberColumn(format="%d d"),
                "Faturamento R$": st.column_config.NumberColumn(format="R$ %.2f"),
                "Lucro R$": st.column_config.NumberColumn(format="R$ %.2f"),
                "Margem %": st.column_config.NumberColumn(format="%.1f%%"),
                "Capital R$": st.column_config.NumberColumn(format="R$ %.2f"),
            },
            use_container_width=True, hide_index=True, height=600,
        )

        st.download_button("📥 Baixar consulta filtrada",
                           exportar_excel_completo(dff),
                           file_name=f"Consulta_Estoque_{datetime.now().strftime('%Y%m%d')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("Nenhum item corresponde aos filtros.")

# ============ ABA 7 — LEGENDA & LÓGICA ============
with t7:
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
    <div class="status-card s-comprar">
      <h4>🚨 COMPRAR — Atingiu o Ponto de Pedido</h4>
      <p>A quantidade está acima do Mínimo, mas <b>já alcançou o Ponto de Pedido</b>. 
      É o gatilho ideal para disparar a compra: se você comprar agora, o material chega 
      antes do estoque acabar (considerando o Lead Time do fornecedor).</p>
    </div>
    <div class="status-card s-saturado">
      <h4>⚠️ SATURADO — Acima do Estoque Máximo</h4>
      <p>Você tem <b>mais do que o necessário</b> para cobrir o consumo previsto. 
      Isso é dinheiro parado em prateleira. Itens saturados raramente precisam de compra 
      e devem ser monitorados para não virarem obsoletos.</p>
    </div>
    <div class="status-card s-ok">
      <h4>✅ OK — Estoque Balanceado</h4>
      <p>Quantidade está entre o Ponto de Pedido e o Estoque Máximo. 
      É a <b>zona ideal</b>: nem risco de ruptura, nem capital parado em excesso.</p>
    </div>
    <div class="status-card s-parado">
      <h4>⚪ PARADO — Tem estoque, mas não vendeu no período</h4>
      <p>Existe quantidade em estoque, mas <b>nenhuma venda foi registrada</b> no período. 
      Não gera alerta de compra, mas é capital imobilizado que pode estar obsoleto.</p>
    </div>
    <div class="status-card s-semmov">
      <h4>⚪ SEM MOVIMENTO — Zerado e sem vendas</h4>
      <p>Estoque zerado e nenhuma venda no período. O item está fora da operação — 
      normalmente é item descontinuado, sazonal ou cadastro inativo.</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 🚨 Ruptura Potencial (coluna 'Ruptura (d)')")
    st.markdown("""
    Mostra **quantos dias o estoque vai faltar** antes do fornecedor entregar. Fórmula:

    <div class="formula-box">
    <b>Ruptura Potencial</b> = Lead Time − Autonomia<br><br>
    <b>Se resultado &gt; 0</b> → você vai ficar sem estoque antes do fornecedor entregar (⚠️ RISCO).<br>
    <b>Se resultado ≤ 0</b> → o material chega antes do estoque acabar (✅ OK).
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 💰 Indicadores de Faturamento e Lucro")
    st.markdown("""
    <div class="formula-box">
    <b>Faturamento Período</b> = vendas registradas no período analisado.<br>
    <b>Faturamento Mensal</b> = Faturamento Período × (30 ÷ dias do período).<br>
    <b>Faturamento Anual</b> = Faturamento Período × (365 ÷ dias do período).<br>
    <b>Lucro</b> = soma dos lucros por venda registrados no arquivo.<br>
    <b>Margem %</b> = (Lucro ÷ Faturamento) × 100.<br>
    <b>Ticket Médio</b> = Faturamento ÷ Nº de pedidos do item.<br>
    <b>% do Total</b> = participação do item no faturamento total.
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 📊 Como funciona a Curva ABC")
    st.markdown("""
    A Curva ABC classifica os SKUs por **importância no faturamento**. 
    A ideia é aplicar a **Lei de Pareto (80/20)**: a minoria dos itens gera a maioria do resultado.
    """)
    st.markdown("""
    <div class="formula-box">
    Ordena todos os SKUs por faturamento decrescente. Calcula o % acumulado.<br><br>
    🅰️ <b>Curva A (Alta)</b> — SKUs que somam os <b>primeiros 80%</b> do faturamento.<br>
    🅱️ <b>Curva B (Média)</b> — Itens que vão dos <b>80% até 95%</b>.<br>
    🅲 <b>Curva C (Baixa)</b> — Itens que vão dos <b>95% até 100%</b>.<br>
    🅳 <b>Curva D (Sem Giro)</b> — Itens que <b>não venderam nada</b> no período.
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 📐 Fórmulas de Estoque")
    st.markdown("""
    <div class="formula-box">
    <b>Giro Diário</b> = Vendas do período ÷ Dias do período<br>
    <b>Estoque de Segurança</b> = Giro Diário × Dias de Segurança (por grupo)<br>
    <b>Estoque Mínimo</b> = Estoque de Segurança<br>
    <b>Ponto de Pedido</b> = (Giro Diário × Lead Time) + Estoque de Segurança<br>
    <b>Estoque Máximo</b> = Ponto de Pedido + (Giro Diário × Dias de Cobertura)<br>
    <b>Quantidade a Comprar</b> = Estoque Máximo − Estoque Atual (só se ≤ Ponto de Pedido)<br>
    <b>Autonomia</b> = Estoque Disponível ÷ Giro Diário<br>
    <b>Fôlego</b> = Comparação entre Autonomia e Lead Time, em texto
    </div>
    """, unsafe_allow_html=True)

# ============ ABA 8 — CONFIGURAÇÃO ============
with t8:
    st.subheader("⚙️ Parâmetros de Cálculo por Grupo")
    st.caption("Estes são os valores usados para calcular Mínimo, Ponto de Pedido e Máximo.")

    conf = pd.DataFrame([
        {"Grupo": c["label"], "Lead Time (dias)": c["lead_time"],
         "Estoque Segurança (dias)": c["dias_seguranca"],
         "Dias Cobertura (após PP)": c["dias_cobertura"]}
        for c in GRUPOS_CONFIG.values()
    ])
    st.dataframe(conf, use_container_width=True, hide_index=True)

# ---------- RODAPÉ ----------
st.markdown("---")
st.markdown(
    "<div style='text-align:center; color:#3A4A5C; font-size:0.9rem;'>"
    "Temper Plus • Gestão Logística Integrada | "
    "Elaborado por <b>André Luiz</b> &amp; <b>Gustavo Schmitz</b></div>",
    unsafe_allow_html=True,
)
