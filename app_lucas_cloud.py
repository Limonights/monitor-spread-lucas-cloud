# -*- coding: utf-8 -*-
from pathlib import Path
from typing import Iterable

import pandas as pd
import plotly.express as px
import streamlit as st


# =============================================================================
# CONFIGURAÇÕES CLOUD
# =============================================================================

APP_DIR = Path(__file__).resolve().parent
DATA_FILE = APP_DIR / "data" / "monitor_spread_lucas_cloud.xlsx"

EXPECTED_SHEETS = {
    "dashboard": pd.DataFrame(),
    "alertas": pd.DataFrame(),
    "ranking_aberturas": pd.DataFrame(),
    "ranking_fechamentos": pd.DataFrame(),
    "historico": pd.DataFrame(),
    "papeis_monitorados": pd.DataFrame(),
    "log_consultas": pd.DataFrame(),
}

USUARIO_NOME = "Lucas"
USUARIO_AREA = "Crédito Privado"

# Colunas solicitadas para não aparecerem nas tabelas do site.
# A normalização permite ocultar mesmo se vier com maiúsculas, acentos ou espaços.
COLUNAS_OCULTAR_SITE = [
    "grupo_economico",
    "subsetor",
    "rating",
    "agencia",
    "data_rating",
    "fonte_rating",
    "intervalo_indicativo_minimo",
    "intervalo_indicativo_maximo",
    "pct_vne",
    "referencia_ntnb",
    "spread_incentivados_sem_gross_up",
    "vna",
]


# =============================================================================
# CONFIG STREAMLIT
# =============================================================================

st.set_page_config(
    page_title="Monitoramento de Spreads",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =============================================================================
# CSS - VISUAL DO DASHBOARD LOCAL, ADAPTADO PARA CLOUD
# =============================================================================

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

:root {
    --navy: #061b3a;
    --blue: #073b82;
    --green: #027a48;
    --red: #d92d20;
    --orange: #f79009;
    --text: #101828;
    --muted: #667085;
    --line: #e6eaf0;
    --bg: #f7f9fc;
    --card: #ffffff;
}

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: var(--bg); }
.block-container { padding-top: 0.65rem; padding-left: 1.2rem; padding-right: 1.2rem; padding-bottom: 1.8rem; max-width: 100%; }
[data-testid="stSidebar"] { background: linear-gradient(180deg, #061b3a 0%, #062449 60%, #02162f 100%); border-right: 1px solid rgba(255,255,255,0.08); min-width: 250px; }
[data-testid="stSidebar"] > div:first-child { padding-top: 1rem; }
[data-testid="stSidebar"] * { color: #f9fafb !important; }
.sidebar-logo { display: flex; align-items: center; gap: 11px; padding: 10px 8px 18px 8px; margin-bottom: 6px; }
.sidebar-icon { width: 34px; height: 34px; border-radius: 10px; background: rgba(255,255,255,0.10); display: flex; align-items: center; justify-content: center; font-size: 18px; }
.sidebar-title { font-size: 15px; font-weight: 800; line-height: 1.18; color: #ffffff !important; }
.sidebar-subtitle { font-size: 11px; color: #b7c6df !important; margin-top: 3px; }
.sidebar-footer { position: fixed; bottom: 18px; left: 20px; width: 210px; font-size: 11px; color: #dbeafe !important; }
div[role="radiogroup"] label { background: transparent !important; border-radius: 10px; padding: 8px 10px !important; margin-bottom: 2px; }
div[role="radiogroup"] label:hover { background: rgba(255,255,255,0.08) !important; }
.topbar { min-height: 50px; background: rgba(255,255,255,0.96); border: 1px solid var(--line); border-radius: 14px; padding: 8px 12px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 8px 24px rgba(16, 24, 40, 0.04); margin-bottom: 10px; }
.topbar-left { display: flex; align-items: center; gap: 10px; color: var(--text); font-size: 12px; font-weight: 700; }
.topbar-pill { background: #f3f6fb; border: 1px solid var(--line); border-radius: 10px; padding: 7px 10px; color: #344054; font-size: 12px; font-weight: 700; white-space: nowrap; }
.topbar-user { display: flex; align-items: center; gap: 10px; }
.avatar { width: 35px; height: 35px; border-radius: 999px; background: #eef4ff; border: 1px solid #d6e4ff; display: flex; align-items: center; justify-content: center; color: #0b2f69; font-weight: 800; }
.user-name { font-size: 12px; font-weight: 800; color: #101828; white-space: nowrap; }
.user-area { font-size: 11px; color: #667085; white-space: nowrap; }
.kpi-card { background: var(--card); border: 1px solid var(--line); border-radius: 16px; padding: 13px 12px; box-shadow: 0 8px 24px rgba(16, 24, 40, 0.045); min-height: 96px; display: flex; align-items: center; gap: 10px; }
.kpi-icon { width: 38px; height: 38px; border-radius: 999px; display: flex; align-items: center; justify-content: center; font-size: 17px; flex-shrink: 0; }
.kpi-blue { background: #eef4ff; color: #175cd3; }
.kpi-red { background: #fff1f0; color: #d92d20; }
.kpi-green { background: #ecfdf3; color: #067647; }
.kpi-orange { background: #fff7ed; color: #f79009; }
.kpi-content { width: 100%; overflow: hidden; }
.kpi-label { font-size: 10.5px; color: #475467; font-weight: 800; margin-bottom: 5px; line-height: 1.1; white-space: normal; }
.kpi-value { font-size: 22px; color: #101828; font-weight: 800; line-height: 1.05; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.kpi-delta-pos { color: var(--green); font-size: 10.5px; font-weight: 800; margin-top: 7px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.kpi-delta-neg { color: var(--red); font-size: 10.5px; font-weight: 800; margin-top: 7px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.panel-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 6px; padding-top: 2px; }
.panel-title { font-size: 14px; font-weight: 800; color: #101828; }
.panel-link { font-size: 11px; font-weight: 800; color: #175cd3; }
.summary-text { background: #f8fbff; border: 1px solid #d9e7ff; border-left: 4px solid #175cd3; border-radius: 14px; padding: 11px 13px; color: #344054; font-size: 12px; line-height: 1.45; margin-top: 14px; margin-bottom: 8px; }
.status-pill { display: inline-block; padding: 4px 8px; border-radius: 999px; font-size: 10px; font-weight: 800; white-space: nowrap; }
.status-red { color: #b42318; background: #fee4e2; border: 1px solid #fecdca; }
.stDataFrame { border-radius: 14px; overflow: hidden; border: 1px solid var(--line); }
[data-testid="stVerticalBlock"] { gap: 0.62rem; }
button[kind="secondary"] { border-radius: 10px !important; border: 1px solid var(--line) !important; background: #ffffff !important; color: #0f2f63 !important; font-weight: 700 !important; }
input { border-radius: 12px !important; }
[data-baseweb="select"] > div { border-radius: 12px !important; }
[data-testid="stMultiSelect"] [data-baseweb="tag"] { background-color: #063b82 !important; color: white !important; border-radius: 6px !important; max-width: 130px !important; }
.spacer-after-kpi { height: 22px; }
.footer-note { color: #667085; font-size: 11px; margin-top: 10px; }
header[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"], [data-testid="stStatusWidget"] { display: none !important; }
#MainMenu { visibility: hidden !important; }
footer { visibility: hidden !important; }
.stDeployButton { display: none !important; }
[data-testid="stAppViewContainer"] { padding-top: 0rem !important; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# =============================================================================
# FUNÇÕES
# =============================================================================

def normalize_text(value: object) -> str:
    return (
        str(value).strip().lower()
        .replace("á", "a").replace("à", "a").replace("ã", "a").replace("â", "a")
        .replace("é", "e").replace("ê", "e")
        .replace("í", "i")
        .replace("ó", "o").replace("ô", "o").replace("õ", "o")
        .replace("ú", "u").replace("ç", "c")
        .replace("%", "pct").replace("/", "_").replace("-", "_").replace(".", "")
        .replace("(", "").replace(")", "").replace("+", "mais").replace(" ", "_")
    )


def first_existing_column(df: pd.DataFrame, candidates: Iterable[str]) -> str | None:
    if df is None or df.empty:
        return None
    normalized_columns = {normalize_text(col): col for col in df.columns}
    for candidate in candidates:
        found = normalized_columns.get(normalize_text(candidate))
        if found is not None:
            return found
    for col_norm, col_original in normalized_columns.items():
        for candidate in candidates:
            candidate_norm = normalize_text(candidate)
            if candidate_norm in col_norm or col_norm in candidate_norm:
                return col_original
    return None


def numeric_series(df: pd.DataFrame, column: str | None) -> pd.Series:
    if column is None or column not in df.columns:
        return pd.Series(dtype="float64")
    return pd.to_numeric(df[column], errors="coerce")


def date_series(df: pd.DataFrame, column: str | None) -> pd.Series:
    if column is None or column not in df.columns:
        return pd.Series(dtype="datetime64[ns]")
    return pd.to_datetime(df[column], errors="coerce", dayfirst=True)


def fmt_int(value) -> str:
    try:
        if pd.isna(value):
            return "-"
        return f"{int(float(value)):,}".replace(",", ".")
    except Exception:
        return str(value)


def limpar_texto(valor, limite=None):
    texto = "" if pd.isna(valor) else str(valor).strip()
    if texto.lower() in ["none", "nan", "nat", "<na>"]:
        texto = ""
    if limite and len(texto) > limite:
        return texto[:limite - 3] + "..."
    return texto


def safe_value_dashboard(df_dashboard: pd.DataFrame, indicador: str, padrao="-"):
    if df_dashboard.empty:
        return padrao
    ind_col = first_existing_column(df_dashboard, ["indicador"])
    val_col = first_existing_column(df_dashboard, ["valor"])
    if not ind_col or not val_col:
        return padrao
    filtro = df_dashboard[ind_col].astype(str) == indicador
    if not filtro.any():
        return padrao
    return df_dashboard.loc[filtro, val_col].iloc[0]


def card_kpi(label, value, delta, icon, cor="blue"):
    delta = "" if delta is None else str(delta)
    delta_class = "kpi-delta-neg" if delta.strip().startswith("-") else "kpi-delta-pos"
    return f"""
    <div class="kpi-card">
        <div class="kpi-icon kpi-{cor}">{icon}</div>
        <div class="kpi-content">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value" title="{value}">{value}</div>
            <div class="{delta_class}" title="{delta}">{delta}</div>
        </div>
    </div>
    """


def plotly_base(fig, height=280, showlegend=True):
    fig.update_layout(
        template="plotly_white",
        height=height,
        paper_bgcolor="rgba(255,255,255,0)",
        plot_bgcolor="rgba(255,255,255,0)",
        margin=dict(l=8, r=8, t=12, b=8),
        font=dict(family="Inter, Arial", size=10, color="#344054"),
        hovermode="closest",
        showlegend=showlegend,
        legend_title_text="",
        title=None,
    )
    fig.update_xaxes(title_text="", showgrid=False, zeroline=False, linecolor="#e6eaf0", tickfont=dict(color="#667085", size=9))
    fig.update_yaxes(title_text="", gridcolor="#eef2f6", zeroline=False, linecolor="#e6eaf0", tickfont=dict(color="#667085", size=9))
    return fig


def obter_lista(df, coluna):
    if df is None or df.empty or not coluna or coluna not in df.columns:
        return []
    valores = df[coluna].dropna().astype(str).str.strip().unique().tolist()
    return sorted([x for x in valores if x and x.lower() not in ["nan", "none", "nat", "<na>"]])


def garantir_colunas(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    date_col = first_existing_column(out, ["data_ref", "data", "data_referencia", "dt_referencia", "referencia"])
    if date_col:
        out[date_col] = pd.to_datetime(out[date_col], errors="coerce")
    for col in ["taxa_indicativa", "taxa_d_1", "pu", "duration", "variacao_dia_bps", "variacao_5d_bps", "variacao_21d_bps", "z_score_20d", "percentil_historico", "impacto_preco_estimado_pct", "spread_medio", "taxa_media", "variacao_media_bps"]:
        real = first_existing_column(out, [col])
        if real:
            out[real] = pd.to_numeric(out[real], errors="coerce")
    cod_col = first_existing_column(out, ["codigo_ativo", "ativo", "codigo", "ticker", "papel", "debenture"])
    if cod_col:
        out[cod_col] = out[cod_col].astype(str).str.strip().str.upper()
    return out


@st.cache_data(show_spinner="Carregando Excel consolidado...")
def load_workbook(path: str) -> dict[str, pd.DataFrame]:
    workbook_path = Path(path)
    sheets = pd.read_excel(workbook_path, sheet_name=None, engine="openpyxl")
    result = {k: v.copy() for k, v in EXPECTED_SHEETS.items()}
    for sheet_name, df in sheets.items():
        key = normalize_text(sheet_name)
        if key in result:
            cleaned = df.copy()
            cleaned.columns = [str(col).strip() for col in cleaned.columns]
            result[key] = garantir_colunas(cleaned.dropna(how="all"))
    return result


def make_excel_download(path: Path) -> bytes:
    with open(path, "rb") as f:
        return f.read()


def make_csv(df: pd.DataFrame) -> bytes:
    if df.empty:
        return "".encode("utf-8-sig")
    return df.to_csv(index=False, sep=";").encode("utf-8-sig")


def aplicar_filtros(df, ativo, status, indexador, setor, rating, texto_busca):
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    col_ativo = first_existing_column(out, ["codigo_ativo", "ativo", "codigo", "ticker", "papel", "debenture"])
    col_status = first_existing_column(out, ["status_alerta", "status"])
    col_indexador = first_existing_column(out, ["indexador", "remuneracao", "tipo_remuneracao"])
    col_setor = first_existing_column(out, ["setor", "segmento", "subsetor"])
    col_rating = first_existing_column(out, ["rating"])
    col_emissor = first_existing_column(out, ["emissor", "emissor_anbima", "nome_emissor"])
    if ativo != "Todos" and col_ativo:
        out = out[out[col_ativo].astype(str).str.strip().str.upper() == str(ativo).strip().upper()]
    if status and col_status:
        out = out[out[col_status].astype(str).isin(status)]
    if indexador and col_indexador:
        out = out[out[col_indexador].astype(str).isin(indexador)]
    if setor and col_setor:
        out = out[out[col_setor].astype(str).isin(setor)]
    if rating and col_rating:
        out = out[out[col_rating].astype(str).isin(rating)]
    if texto_busca:
        texto = texto_busca.upper().strip()
        cols_busca = [c for c in [col_ativo, col_emissor, col_setor, col_rating, col_indexador, first_existing_column(out, ["grupo_comparavel", "grupo_economico"])] if c and c in out.columns]
        if cols_busca:
            mask = pd.Series(False, index=out.index)
            for col in cols_busca:
                mask = mask | out[col].astype(str).str.upper().str.contains(texto, na=False)
            out = out[mask]
    return out


def preparar_tabela_alertas_card(df_alertas):
    if df_alertas.empty:
        return pd.DataFrame()
    cols = {
        "Ativo": first_existing_column(df_alertas, ["codigo_ativo", "ativo", "codigo", "ticker"]),
        "Emissor": first_existing_column(df_alertas, ["emissor", "emissor_anbima", "nome_emissor"]),
        "Var. bps": first_existing_column(df_alertas, ["variacao_dia_bps", "variacao", "delta_spread"]),
        "Status": first_existing_column(df_alertas, ["status_alerta", "status"]),
    }
    data = {final: df_alertas[real] for final, real in cols.items() if real}
    if not data:
        return pd.DataFrame()
    out = pd.DataFrame(data).head(8)
    if "Emissor" in out.columns:
        out["Emissor"] = out["Emissor"].apply(lambda x: limpar_texto(x, 26))
    return out


def preparar_tabela_principal(df_alertas):
    if df_alertas.empty:
        return pd.DataFrame()
    mapping = {
        "Ativo": ["codigo_ativo", "ativo", "codigo", "ticker"],
        "Emissor": ["emissor", "emissor_anbima", "nome_emissor"],
        "Indexador": ["indexador", "remuneracao", "tipo_remuneracao"],
        "Setor": ["setor"],
        "Rating": ["rating"],
        "Duration": ["duration"],
        "Taxa D-1": ["taxa_d_1"],
        "Taxa Atual": ["taxa_indicativa", "taxa_atual"],
        "Variação (bps)": ["variacao_dia_bps", "variacao", "delta_spread"],
        "Percentil": ["percentil_historico", "percentil"],
        "Z-score": ["z_score_20d", "z_score"],
        "Status": ["status_alerta", "status"],
    }
    data = {}
    for final, candidates in mapping.items():
        real = first_existing_column(df_alertas, candidates)
        if real:
            data[final] = df_alertas[real]
    if not data:
        return pd.DataFrame()
    out = pd.DataFrame(data)
    if "Emissor" in out.columns:
        out["Emissor"] = out["Emissor"].apply(lambda x: limpar_texto(x, 38))
    return out


def leitura_do_dia(df_dashboard, df_alertas):
    maior_abertura = safe_value_dashboard(df_dashboard, "maior_abertura", "-")
    maior_fechamento = safe_value_dashboard(df_dashboard, "maior_fechamento", "-")
    alertas_fortes = safe_value_dashboard(df_dashboard, "alertas_fortes", "0")
    aberturas = safe_value_dashboard(df_dashboard, "aberturas_identificadas", "0")
    fechamentos = safe_value_dashboard(df_dashboard, "fechamentos_identificados", "0")
    texto = (
        f"Mercado com {aberturas} ativos em abertura contra {fechamentos} em fechamento. "
        f"Foram identificados {alertas_fortes} alertas fortes. "
        f"Destaque de abertura: {maior_abertura}. Destaque de fechamento: {maior_fechamento}."
    )
    status_col = first_existing_column(df_alertas, ["status_alerta", "status"])
    if not df_alertas.empty and status_col:
        qtd_ok = (df_alertas[status_col].astype(str) == "OK").sum()
        texto += f" Ativos sem alerta relevante: {qtd_ok}."
    return texto



def ocultar_colunas_site(df: pd.DataFrame) -> pd.DataFrame:
    """Remove colunas que não devem aparecer nas tabelas do site."""
    if df is None or df.empty:
        return pd.DataFrame()

    out = df.copy()
    colunas_ocultas_normalizadas = {normalize_text(col) for col in COLUNAS_OCULTAR_SITE}

    colunas_para_remover = [
        col for col in out.columns
        if normalize_text(col) in colunas_ocultas_normalizadas
    ]

    if colunas_para_remover:
        out = out.drop(columns=colunas_para_remover, errors="ignore")

    return out



def render_sidebar():
    with st.sidebar:
        st.markdown(
            """
            <div class="sidebar-logo">
                <div class="sidebar-icon">▥</div>
                <div>
                    <div class="sidebar-title">Monitoramento de<br>Spreads de<br>Debêntures</div>
                    <div class="sidebar-subtitle">Crédito Privado</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        pagina = st.radio(
            "Navegação",
            ["Dashboard", "Papéis Monitorados", "Alertas", "Histórico", "Setores", "Ratings", "Configurações"],
            label_visibility="collapsed",
        )
        st.markdown(
            """
            <div class="sidebar-footer">
                <div style="display:flex;align-items:center;gap:8px;">
                    <span style="width:9px;height:9px;border-radius:99px;background:#32d583;display:inline-block;"></span>
                    <span>Dados atualizados</span>
                </div>
                <div style="margin-top:6px;color:#b7c6df !important;">Fonte: ANBIMA</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    return pagina


def render_topbar(data_ref, arquivo_disponivel=True):
    top1, top2 = st.columns([5.8, 1.1])
    with top1:
        st.markdown(
            f"""
            <div class="topbar">
                <div class="topbar-left">
                    <span class="topbar-pill">🔒 Acesso Interno</span>
                    <span class="topbar-pill">Data base&nbsp;&nbsp; {data_ref}</span>
                </div>
                <div class="topbar-user">
                    <div class="avatar">L</div>
                    <div>
                        <div class="user-name">{USUARIO_NOME}</div>
                        <div class="user-area">{USUARIO_AREA}</div>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with top2:
        if arquivo_disponivel:
            st.download_button(
                "⬇ Exportar",
                data=make_excel_download(DATA_FILE),
                file_name=DATA_FILE.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )


def render_table(df: pd.DataFrame, height=620):
    df_view = ocultar_colunas_site(df)

    if df_view.empty:
        st.warning("Sem dados para exibir.")
        return

    st.dataframe(df_view, use_container_width=True, hide_index=True, height=height)


def main():
    pagina = render_sidebar()
    if not DATA_FILE.exists():
        render_topbar("-", arquivo_disponivel=False)
        st.error("Arquivo Excel não encontrado.")
        st.write("Coloque o arquivo consolidado neste caminho:")
        st.code("data/monitor_spread_lucas_cloud.xlsx")
        st.stop()
    try:
        frames = load_workbook(str(DATA_FILE))
    except Exception as exc:
        render_topbar("-", arquivo_disponivel=False)
        st.error("Não foi possível ler o Excel consolidado.")
        st.exception(exc)
        st.stop()

    df_dashboard = frames["dashboard"]
    df_alertas = frames["alertas"]
    df_aberturas = frames["ranking_aberturas"]
    df_fechamentos = frames["ranking_fechamentos"]
    df_historico = frames["historico"]
    df_papeis = frames["papeis_monitorados"]
    df_logs = frames["log_consultas"]

    data_ref = safe_value_dashboard(df_dashboard, "data_referencia", "-")
    if data_ref == "-":
        date_col_hist = first_existing_column(df_historico, ["data_ref", "data", "data_referencia", "dt_referencia"])
        last_date = date_series(df_historico, date_col_hist).max()
        data_ref = "-" if pd.isna(last_date) else last_date.strftime("%Y-%m-%d")
    render_topbar(data_ref, arquivo_disponivel=True)

    col_ativo_hist = first_existing_column(df_historico, ["codigo_ativo", "ativo", "codigo", "ticker"])
    col_status = first_existing_column(df_alertas, ["status_alerta", "status"])
    col_indexador = first_existing_column(df_alertas, ["indexador", "remuneracao", "tipo_remuneracao"])
    col_setor = first_existing_column(df_alertas, ["setor"])
    col_rating = first_existing_column(df_alertas, ["rating"])

    ativos_disponiveis = obter_lista(df_historico, col_ativo_hist)
    status_disponiveis = obter_lista(df_alertas, col_status)
    indexadores = obter_lista(df_alertas, col_indexador)
    setores = obter_lista(df_alertas, col_setor)
    ratings = obter_lista(df_alertas, col_rating)

    ftop1, ftop2 = st.columns([1.2, 4.8])
    with ftop1:
        ativo_selecionado = st.selectbox("Ativo", options=["Todos"] + ativos_disponiveis, index=0, label_visibility="collapsed")
    with ftop2:
        texto_busca = st.text_input("Buscar", placeholder="Buscar ativo, emissor, setor, rating...", label_visibility="collapsed")

    f1, f2, f3, f4 = st.columns(4)
    with f1:
        status_selecionado = st.multiselect("Status", options=status_disponiveis, default=[], placeholder="Todos os status")
    with f2:
        indexador_selecionado = st.multiselect("Indexador", options=indexadores, default=[], placeholder="Todos os indexadores")
    with f3:
        setor_selecionado = st.multiselect("Setor", options=setores, default=[], placeholder="Todos os setores")
    with f4:
        rating_selecionado = st.multiselect("Rating", options=ratings, default=[], placeholder="Todos os ratings")

    df_alertas_filtrado = aplicar_filtros(df_alertas, ativo_selecionado, status_selecionado, indexador_selecionado, setor_selecionado, rating_selecionado, texto_busca)
    df_historico_filtrado = aplicar_filtros(df_historico, ativo_selecionado, [], indexador_selecionado, setor_selecionado, rating_selecionado, texto_busca)

    papeis_monitorados = safe_value_dashboard(df_dashboard, "papeis_monitorados", df_papeis.shape[0] if not df_papeis.empty else df_historico[col_ativo_hist].nunique() if col_ativo_hist else "-")
    alertas_fortes = safe_value_dashboard(df_dashboard, "alertas_fortes", "-")
    aberturas = safe_value_dashboard(df_dashboard, "aberturas_identificadas", "-")
    fechamentos = safe_value_dashboard(df_dashboard, "fechamentos_identificados", "-")
    registros = safe_value_dashboard(df_dashboard, "registros_historicos", len(df_historico))
    maior_abertura = safe_value_dashboard(df_dashboard, "maior_abertura", "-")
    maior_fechamento = safe_value_dashboard(df_dashboard, "maior_fechamento", "-")

    if pagina == "Dashboard":
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        with k1:
            st.markdown(card_kpi("Papéis", fmt_int(papeis_monitorados), "monitorados", "▤", "blue"), unsafe_allow_html=True)
        with k2:
            st.markdown(card_kpi("Alertas Fortes", fmt_int(alertas_fortes), "crítico", "⚠", "red"), unsafe_allow_html=True)
        with k3:
            st.markdown(card_kpi("Aberturas", fmt_int(aberturas), "spread maior", "↗", "green"), unsafe_allow_html=True)
        with k4:
            st.markdown(card_kpi("Fechamentos", fmt_int(fechamentos), "spread menor", "↘", "blue"), unsafe_allow_html=True)
        with k5:
            valor = str(maior_abertura).split("|")[0].strip() if maior_abertura != "-" else "-"
            delta = str(maior_abertura).split("|")[1].strip() if "|" in str(maior_abertura) else ""
            st.markdown(card_kpi("Máx Abertura", valor, delta, "↗", "orange"), unsafe_allow_html=True)
        with k6:
            valor = str(maior_fechamento).split("|")[0].strip() if maior_fechamento != "-" else "-"
            delta = str(maior_fechamento).split("|")[1].strip() if "|" in str(maior_fechamento) else ""
            st.markdown(card_kpi("Máx Fechamento", valor, delta, "↘", "blue"), unsafe_allow_html=True)

        st.markdown('<div class="spacer-after-kpi"></div>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns([2.25, 1.05, 1.38])

        with c1:
            st.markdown('<div class="panel-header"><div class="panel-title">Histórico de Spread / Taxa Indicativa</div><div class="panel-link">Selecionável</div></div>', unsafe_allow_html=True)
            df_hist_base = df_historico_filtrado.copy()
            cod_col = first_existing_column(df_hist_base, ["codigo_ativo", "ativo", "codigo", "ticker"])
            date_col = first_existing_column(df_hist_base, ["data_ref", "data", "data_referencia", "dt_referencia"])
            taxa_col = first_existing_column(df_hist_base, ["taxa_indicativa", "spread", "spread_atual", "valor"])
            if df_hist_base.empty or not cod_col or not date_col or not taxa_col:
                st.warning("Sem histórico para exibir.")
            else:
                opcoes_papeis_grafico = obter_lista(df_hist_base, cod_col)
                if ativo_selecionado != "Todos" and ativo_selecionado in opcoes_papeis_grafico:
                    papeis_default = [ativo_selecionado]
                else:
                    var_col = first_existing_column(df_alertas_filtrado, ["variacao_dia_bps", "variacao", "delta_spread"])
                    cod_alerta = first_existing_column(df_alertas_filtrado, ["codigo_ativo", "ativo", "codigo", "ticker"])
                    if not df_alertas_filtrado.empty and var_col and cod_alerta:
                        ativos_top = df_alertas_filtrado.copy().assign(abs_var=lambda x: pd.to_numeric(x[var_col], errors="coerce").abs()).sort_values("abs_var", ascending=False).head(5)[cod_alerta].astype(str).str.strip().unique().tolist()
                    else:
                        ativos_top = opcoes_papeis_grafico[:5]
                    papeis_default = [x for x in ativos_top if x in opcoes_papeis_grafico] or opcoes_papeis_grafico[:5]
                papeis_grafico = st.multiselect("Selecionar papéis no gráfico", options=opcoes_papeis_grafico, default=papeis_default, placeholder="Escolha um ou mais papéis", key="papeis_grafico_historico")
                if not papeis_grafico:
                    st.info("Selecione pelo menos um papel para exibir o histórico.")
                else:
                    df_hist_plot = df_hist_base[df_hist_base[cod_col].astype(str).isin(papeis_grafico)].copy()
                    df_hist_plot[date_col] = pd.to_datetime(df_hist_plot[date_col], errors="coerce")
                    df_hist_plot[taxa_col] = pd.to_numeric(df_hist_plot[taxa_col], errors="coerce")
                    df_hist_plot = df_hist_plot.dropna(subset=[date_col, taxa_col]).sort_values(date_col)
                    if df_hist_plot.empty:
                        st.warning("Sem histórico para os papéis selecionados.")
                    else:
                        hover_cols = [c for c in [first_existing_column(df_hist_plot, ["emissor", "emissor_anbima"]), first_existing_column(df_hist_plot, ["pu"]), first_existing_column(df_hist_plot, ["duration"]), first_existing_column(df_hist_plot, ["z_score_20d", "z_score"])] if c]
                        fig_hist = px.line(df_hist_plot, x=date_col, y=taxa_col, color=cod_col, markers=False, labels={date_col: "", taxa_col: "", cod_col: ""}, hover_data=hover_cols)
                        fig_hist.update_traces(line=dict(width=2.5))
                        fig_hist = plotly_base(fig_hist, height=315, showlegend=True)
                        st.plotly_chart(fig_hist, use_container_width=True)

        with c2:
            st.markdown('<div class="panel-header"><div class="panel-title">Top Aberturas</div></div>', unsafe_allow_html=True)
            var_col = first_existing_column(df_alertas_filtrado, ["variacao_dia_bps", "variacao", "delta_spread"])
            cod_col_alerta = first_existing_column(df_alertas_filtrado, ["codigo_ativo", "ativo", "codigo", "ticker"])
            if not df_alertas_filtrado.empty and var_col and cod_col_alerta:
                df_top_ab = df_alertas_filtrado.copy()
                df_top_ab[var_col] = pd.to_numeric(df_top_ab[var_col], errors="coerce")
                df_top_ab = df_top_ab.dropna(subset=[var_col]).sort_values(var_col, ascending=False).head(6)
                if df_top_ab.empty:
                    st.warning("Sem aberturas.")
                else:
                    fig_ab = px.bar(df_top_ab, x=var_col, y=cod_col_alerta, orientation="h", text=var_col, labels={var_col: "", cod_col_alerta: ""})
                    fig_ab.update_traces(marker_color="#063b82", texttemplate="+%{text:.0f}", textposition="outside")
                    fig_ab.update_layout(yaxis=dict(autorange="reversed"))
                    fig_ab = plotly_base(fig_ab, height=315, showlegend=False)
                    st.plotly_chart(fig_ab, use_container_width=True)
            else:
                st.warning("Sem aberturas.")

        with c3:
            qtd_alertas = len(df_alertas_filtrado) if not df_alertas_filtrado.empty else 0
            st.markdown(f'<div class="panel-header"><div class="panel-title">Alertas do Dia <span class="status-pill status-red">{qtd_alertas}</span></div><div class="panel-link">Ver todos</div></div>', unsafe_allow_html=True)
            df_card_alertas = preparar_tabela_alertas_card(df_alertas_filtrado)
            if df_card_alertas.empty:
                st.warning("Nenhum alerta.")
            else:
                config = {"Var. bps": st.column_config.NumberColumn(format="%.2f")} if "Var. bps" in df_card_alertas.columns else {}
                st.dataframe(df_card_alertas, use_container_width=True, hide_index=True, height=279, column_config=config)

        st.markdown(f'<div class="summary-text"><b>Leitura automática do dia:</b> {leitura_do_dia(df_dashboard, df_alertas_filtrado)}</div>', unsafe_allow_html=True)
        df_tabela = ocultar_colunas_site(preparar_tabela_principal(df_alertas_filtrado))
        if not df_tabela.empty:
            column_config = {}
            for col in ["Variação (bps)", "Taxa D-1", "Taxa Atual", "Duration", "Z-score"]:
                if col in df_tabela.columns:
                    column_config[col] = st.column_config.NumberColumn(format="%.2f")
            if "Percentil" in df_tabela.columns:
                column_config["Percentil"] = st.column_config.NumberColumn(format="%.1f%%")
            st.dataframe(df_tabela, use_container_width=True, hide_index=True, height=280, column_config=column_config)
        else:
            st.warning("Sem dados para tabela principal.")

        b1, b2, b3 = st.columns([1.1, 1.1, 1.7])
        with b1:
            st.markdown('<div class="panel-header"><div class="panel-title">Análise por Setor</div></div>', unsafe_allow_html=True)
            setor_col = first_existing_column(df_alertas_filtrado, ["setor"])
            taxa_col = first_existing_column(df_alertas_filtrado, ["taxa_indicativa", "spread", "spread_atual"])
            if not df_alertas_filtrado.empty and setor_col and taxa_col and df_alertas_filtrado[setor_col].notna().any():
                df_setor = df_alertas_filtrado.copy()
                df_setor[taxa_col] = pd.to_numeric(df_setor[taxa_col], errors="coerce")
                df_setor = df_setor.groupby(setor_col, as_index=False).agg(spread_medio=(taxa_col, "mean")).sort_values("spread_medio", ascending=False).head(7)
                fig_setor = px.bar(df_setor, x="spread_medio", y=setor_col, orientation="h", text="spread_medio", labels={"spread_medio": "", setor_col: ""})
                fig_setor.update_traces(marker_color="#063b82", texttemplate="%{text:.2f}", textposition="outside")
                fig_setor.update_layout(yaxis=dict(autorange="reversed"))
                fig_setor = plotly_base(fig_setor, height=248, showlegend=False)
                st.plotly_chart(fig_setor, use_container_width=True)
            else:
                st.info("Preencher setor na base para habilitar esta visão.")
        with b2:
            st.markdown('<div class="panel-header"><div class="panel-title">Distribuição por Status</div></div>', unsafe_allow_html=True)
            status_col = first_existing_column(df_alertas_filtrado, ["status_alerta", "status"])
            cod_col = first_existing_column(df_alertas_filtrado, ["codigo_ativo", "ativo", "codigo", "ticker"])
            if not df_alertas_filtrado.empty and status_col:
                count_col = cod_col or status_col
                df_status = df_alertas_filtrado.groupby(status_col, as_index=False).agg(qtd=(count_col, "count")).sort_values("qtd", ascending=False)
                fig_status = px.pie(df_status, names=status_col, values="qtd", hole=0.58, color=status_col, color_discrete_map={"ALERTA FORTE - ABERTURA": "#d92d20", "ALERTA FORTE - FECHAMENTO": "#027a48", "ALERTA MODERADO - ABERTURA": "#f79009", "ALERTA MODERADO - FECHAMENTO": "#12b76a", "ABERTURA LEVE": "#fdb022", "FECHAMENTO LEVE": "#32d583", "OK": "#98a2b3", "Sem D-1": "#d0d5dd"})
                fig_status.update_traces(textposition="inside", textinfo="percent")
                fig_status = plotly_base(fig_status, height=248, showlegend=True)
                st.plotly_chart(fig_status, use_container_width=True)
            else:
                st.warning("Sem status.")
        with b3:
            st.markdown('<div class="panel-header"><div class="panel-title">Comparação vs Pares</div><div class="panel-link">Duration x Taxa</div></div>', unsafe_allow_html=True)
            duration_col = first_existing_column(df_alertas_filtrado, ["duration"])
            taxa_col = first_existing_column(df_alertas_filtrado, ["taxa_indicativa", "spread", "spread_atual"])
            cod_col = first_existing_column(df_alertas_filtrado, ["codigo_ativo", "ativo", "codigo", "ticker"])
            status_col = first_existing_column(df_alertas_filtrado, ["status_alerta", "status"])
            if not df_alertas_filtrado.empty and duration_col and taxa_col:
                df_scatter = df_alertas_filtrado.copy()
                df_scatter[duration_col] = pd.to_numeric(df_scatter[duration_col], errors="coerce")
                df_scatter[taxa_col] = pd.to_numeric(df_scatter[taxa_col], errors="coerce")
                df_scatter = df_scatter.dropna(subset=[duration_col, taxa_col])
                if df_scatter.empty:
                    st.info("Sem dados válidos de duration/taxa.")
                else:
                    fig_scatter = px.scatter(df_scatter, x=duration_col, y=taxa_col, color=status_col if status_col else None, hover_name=cod_col if cod_col else None, labels={duration_col: "", taxa_col: "", status_col or "": ""})
                    fig_scatter.update_traces(marker=dict(size=8, opacity=0.82))
                    fig_scatter = plotly_base(fig_scatter, height=248, showlegend=True)
                    st.plotly_chart(fig_scatter, use_container_width=True)
            else:
                st.warning("Sem dados para comparação.")
        st.markdown('<div class="footer-note">Fonte: ANBIMA | Elaboração: Crédito Privado | Dados referentes ao fechamento do dia anterior.</div>', unsafe_allow_html=True)

    elif pagina == "Papéis Monitorados":
        st.markdown('<div class="panel-header"><div class="panel-title">Papéis Monitorados</div></div>', unsafe_allow_html=True)
        render_table(df_papeis)
    elif pagina == "Alertas":
        st.markdown('<div class="panel-header"><div class="panel-title">Alertas</div></div>', unsafe_allow_html=True)
        render_table(df_alertas_filtrado)
        st.download_button("Baixar alertas em CSV", data=make_csv(ocultar_colunas_site(df_alertas_filtrado)), file_name="alertas_monitor_spread_lucas.csv", mime="text/csv")
    elif pagina == "Histórico":
        st.markdown('<div class="panel-header"><div class="panel-title">Histórico Completo</div></div>', unsafe_allow_html=True)
        render_table(df_historico_filtrado)
        st.download_button("Baixar histórico em CSV", data=make_csv(ocultar_colunas_site(df_historico_filtrado)), file_name="historico_monitor_spread_lucas.csv", mime="text/csv")
    elif pagina == "Setores":
        st.markdown('<div class="panel-header"><div class="panel-title">Setores</div></div>', unsafe_allow_html=True)
        setor_col = first_existing_column(df_alertas, ["setor"])
        taxa_col = first_existing_column(df_alertas, ["taxa_indicativa", "spread", "spread_atual"])
        var_col = first_existing_column(df_alertas, ["variacao_dia_bps", "variacao", "delta_spread"])
        cod_col = first_existing_column(df_alertas, ["codigo_ativo", "ativo", "codigo", "ticker"])
        if not df_alertas.empty and setor_col and df_alertas[setor_col].notna().any():
            df_setores = df_alertas.copy()
            if taxa_col: df_setores[taxa_col] = pd.to_numeric(df_setores[taxa_col], errors="coerce")
            if var_col: df_setores[var_col] = pd.to_numeric(df_setores[var_col], errors="coerce")
            agg_dict = {"papeis": (cod_col or setor_col, "count")}
            if taxa_col: agg_dict["taxa_media"] = (taxa_col, "mean")
            if var_col: agg_dict["variacao_media_bps"] = (var_col, "mean")
            df_setores = df_setores.groupby(setor_col, as_index=False).agg(**agg_dict).sort_values(list(agg_dict.keys())[1] if len(agg_dict) > 1 else "papeis", ascending=False)
            render_table(df_setores)
        else:
            st.info("Setor ainda não está preenchido na base de papéis monitorados.")
    elif pagina == "Ratings":
        st.markdown('<div class="panel-header"><div class="panel-title">Ratings</div></div>', unsafe_allow_html=True)
        rating_col = first_existing_column(df_alertas, ["rating"])
        taxa_col = first_existing_column(df_alertas, ["taxa_indicativa", "spread", "spread_atual"])
        var_col = first_existing_column(df_alertas, ["variacao_dia_bps", "variacao", "delta_spread"])
        cod_col = first_existing_column(df_alertas, ["codigo_ativo", "ativo", "codigo", "ticker"])
        if not df_alertas.empty and rating_col and df_alertas[rating_col].notna().any():
            df_ratings = df_alertas.copy()
            if taxa_col: df_ratings[taxa_col] = pd.to_numeric(df_ratings[taxa_col], errors="coerce")
            if var_col: df_ratings[var_col] = pd.to_numeric(df_ratings[var_col], errors="coerce")
            agg_dict = {"papeis": (cod_col or rating_col, "count")}
            if taxa_col: agg_dict["taxa_media"] = (taxa_col, "mean")
            if var_col: agg_dict["variacao_media_bps"] = (var_col, "mean")
            df_ratings = df_ratings.groupby(rating_col, as_index=False).agg(**agg_dict).sort_values(list(agg_dict.keys())[1] if len(agg_dict) > 1 else "papeis", ascending=False)
            render_table(df_ratings)
        else:
            st.info("Rating ainda não está preenchido na base de papéis monitorados.")
    else:
        st.markdown('<div class="panel-header"><div class="panel-title">Configurações</div></div>', unsafe_allow_html=True)
        st.write("Arquivo cloud carregado:")
        st.code(str(DATA_FILE))
        st.write("Abas esperadas:")
        st.code(", ".join(EXPECTED_SHEETS.keys()))
        st.write("Registros históricos:")
        st.code(str(registros))
        st.success("Arquivo encontrado e carregado com sucesso.")
        if not df_logs.empty:
            st.markdown('<div class="panel-header"><div class="panel-title">Log de consultas</div></div>', unsafe_allow_html=True)
            render_table(df_logs, height=360)


if __name__ == "__main__":
    main()
