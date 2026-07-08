from pathlib import Path
from typing import Iterable

import pandas as pd
import plotly.express as px
import streamlit as st


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


st.set_page_config(
    page_title="Monitor de Spreads - Debentures",
    page_icon="📊",
    layout="wide",
)


def normalize_text(value: object) -> str:
    return (
        str(value)
        .strip()
        .lower()
        .replace("á", "a")
        .replace("à", "a")
        .replace("ã", "a")
        .replace("â", "a")
        .replace("é", "e")
        .replace("ê", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ô", "o")
        .replace("õ", "o")
        .replace("ú", "u")
        .replace("ç", "c")
        .replace("%", "pct")
        .replace("/", "_")
        .replace("-", "_")
        .replace(" ", "_")
    )


def first_existing_column(df: pd.DataFrame, candidates: Iterable[str]) -> str | None:
    if df.empty:
        return None

    normalized_columns = {normalize_text(col): col for col in df.columns}
    for candidate in candidates:
        found = normalized_columns.get(normalize_text(candidate))
        if found is not None:
            return found
    return None


def numeric_series(df: pd.DataFrame, column: str | None) -> pd.Series:
    if column is None or column not in df.columns:
        return pd.Series(dtype="float64")
    return pd.to_numeric(df[column], errors="coerce")


def date_series(df: pd.DataFrame, column: str | None) -> pd.Series:
    if column is None or column not in df.columns:
        return pd.Series(dtype="datetime64[ns]")
    return pd.to_datetime(df[column], errors="coerce", dayfirst=True)


def format_number(value: float | int | None, suffix: str = "") -> str:
    if value is None or pd.isna(value):
        return "-"
    return f"{value:,.2f}{suffix}".replace(",", "X").replace(".", ",").replace("X", ".")


@st.cache_data(show_spinner="Carregando Excel consolidado...")
def load_workbook(path: str) -> dict[str, pd.DataFrame]:
    workbook_path = Path(path)
    sheets = pd.read_excel(workbook_path, sheet_name=None, engine="openpyxl")
    result = EXPECTED_SHEETS.copy()

    for sheet_name, df in sheets.items():
        key = normalize_text(sheet_name)
        if key in result:
            cleaned = df.copy()
            cleaned.columns = [str(col).strip() for col in cleaned.columns]
            result[key] = cleaned.dropna(how="all")

    return result


def filter_by_sidebar(
    frames: dict[str, pd.DataFrame],
    selected_assets: list[str],
    selected_sectors: list[str],
) -> dict[str, pd.DataFrame]:
    filtered = {}
    asset_candidates = ["ativo", "codigo", "ticker", "papel", "debenture"]
    sector_candidates = ["setor", "segmento", "classe", "emissor"]

    for name, df in frames.items():
        current = df.copy()
        asset_col = first_existing_column(current, asset_candidates)
        sector_col = first_existing_column(current, sector_candidates)

        if selected_assets and asset_col:
            current = current[current[asset_col].astype(str).isin(selected_assets)]
        if selected_sectors and sector_col:
            current = current[current[sector_col].astype(str).isin(selected_sectors)]

        filtered[name] = current

    return filtered


def collect_filter_values(frames: dict[str, pd.DataFrame], candidates: list[str]) -> list[str]:
    values: set[str] = set()
    for df in frames.values():
        column = first_existing_column(df, candidates)
        if column:
            values.update(df[column].dropna().astype(str).str.strip())
    return sorted(value for value in values if value)


def render_table(title: str, df: pd.DataFrame, empty_message: str) -> None:
    st.subheader(title)
    if df.empty:
        st.info(empty_message)
        return
    st.dataframe(df, use_container_width=True, hide_index=True)


def render_ranking_chart(title: str, df: pd.DataFrame, color: str) -> None:
    if df.empty:
        return

    asset_col = first_existing_column(df, ["ativo", "codigo", "ticker", "papel", "debenture"])
    value_col = first_existing_column(
        df,
        [
            "variacao_spread",
            "delta_spread",
            "abertura",
            "fechamento",
            "spread_atual",
            "spread",
            "valor",
        ],
    )

    if not asset_col or not value_col:
        return

    chart_df = df.copy()
    chart_df[value_col] = numeric_series(chart_df, value_col)
    chart_df = chart_df.dropna(subset=[value_col]).head(10)
    if chart_df.empty:
        return

    fig = px.bar(
        chart_df,
        x=value_col,
        y=asset_col,
        orientation="h",
        title=title,
        text=value_col,
        color_discrete_sequence=[color],
    )
    fig.update_layout(yaxis={"categoryorder": "total ascending"}, height=420)
    fig.update_traces(texttemplate="%{text:.2f}", textposition="outside")
    st.plotly_chart(fig, use_container_width=True)


def render_history(historico: pd.DataFrame, assets: list[str]) -> None:
    st.subheader("Histórico por ativo")
    if historico.empty:
        st.info("A aba historico esta vazia ou nao foi encontrada no Excel.")
        return

    asset_col = first_existing_column(historico, ["ativo", "codigo", "ticker", "papel", "debenture"])
    date_col = first_existing_column(historico, ["data", "dt_referencia", "data_referencia", "referencia"])
    spread_col = first_existing_column(historico, ["spread", "spread_atual", "taxa_spread", "valor"])

    if not asset_col or not date_col or not spread_col:
        st.warning("Nao encontrei colunas suficientes para montar o grafico historico.")
        st.dataframe(historico, use_container_width=True, hide_index=True)
        return

    available_assets = sorted(historico[asset_col].dropna().astype(str).unique())
    default_assets = assets if assets else available_assets[:5]
    selected = st.multiselect(
        "Ativos no historico",
        options=available_assets,
        default=[asset for asset in default_assets if asset in available_assets],
    )

    chart_df = historico.copy()
    chart_df[date_col] = date_series(chart_df, date_col)
    chart_df[spread_col] = numeric_series(chart_df, spread_col)

    if selected:
        chart_df = chart_df[chart_df[asset_col].astype(str).isin(selected)]

    chart_df = chart_df.dropna(subset=[date_col, spread_col]).sort_values(date_col)
    if chart_df.empty:
        st.info("Nao ha dados suficientes para o grafico com os filtros atuais.")
    else:
        fig = px.line(
            chart_df,
            x=date_col,
            y=spread_col,
            color=asset_col,
            markers=True,
            title="Evolucao do spread",
        )
        fig.update_layout(height=470)
        st.plotly_chart(fig, use_container_width=True)

    st.dataframe(historico, use_container_width=True, hide_index=True)


def render_kpis(frames: dict[str, pd.DataFrame]) -> None:
    dashboard = frames["dashboard"]
    alertas = frames["alertas"]
    papeis = frames["papeis_monitorados"]
    historico = frames["historico"]

    spread_col = first_existing_column(dashboard, ["spread", "spread_atual", "taxa_spread", "valor"])
    variation_col = first_existing_column(dashboard, ["variacao_spread", "delta_spread", "mudanca_spread"])
    date_col = first_existing_column(historico, ["data", "dt_referencia", "data_referencia", "referencia"])

    total_assets = len(papeis) if not papeis.empty else len(dashboard)
    active_alerts = len(alertas)
    average_spread = numeric_series(dashboard, spread_col).mean()
    average_change = numeric_series(dashboard, variation_col).mean()
    last_date = date_series(historico, date_col).max()

    kpi_cols = st.columns(5)
    kpi_cols[0].metric("Papeis monitorados", f"{total_assets:,}".replace(",", "."))
    kpi_cols[1].metric("Alertas", f"{active_alerts:,}".replace(",", "."))
    kpi_cols[2].metric("Spread medio", format_number(average_spread))
    kpi_cols[3].metric("Variacao media", format_number(average_change))
    kpi_cols[4].metric("Ultima data", "-" if pd.isna(last_date) else last_date.strftime("%d/%m/%Y"))


def make_csv(frames: dict[str, pd.DataFrame]) -> bytes:
    base = frames["dashboard"]
    if base.empty:
        base = pd.concat(
            [df.assign(aba=name) for name, df in frames.items() if not df.empty],
            ignore_index=True,
            sort=False,
        )
    return base.to_csv(index=False, sep=";").encode("utf-8-sig")


def main() -> None:
    st.title("Monitor de Spreads - Debentures")
    st.caption("Versao simplificada para Streamlit Cloud, baseada apenas no Excel consolidado.")

    if not DATA_FILE.exists():
        st.error("Arquivo Excel nao encontrado.")
        st.write("Coloque o arquivo consolidado neste caminho:")
        st.code("data/monitor_spread_lucas_cloud.xlsx")
        st.stop()

    try:
        frames = load_workbook(str(DATA_FILE))
    except Exception as exc:
        st.error("Nao foi possivel ler o Excel consolidado.")
        st.exception(exc)
        st.stop()

    with st.sidebar:
        st.header("Filtros")
        assets = collect_filter_values(frames, ["ativo", "codigo", "ticker", "papel", "debenture"])
        sectors = collect_filter_values(frames, ["setor", "segmento", "classe", "emissor"])

        selected_assets = st.multiselect("Ativos", options=assets)
        selected_sectors = st.multiselect("Setores / emissores", options=sectors)

        st.divider()
        st.download_button(
            "Baixar CSV",
            data=make_csv(filter_by_sidebar(frames, selected_assets, selected_sectors)),
            file_name="monitor_spread_lucas.csv",
            mime="text/csv",
            use_container_width=True,
        )

    filtered = filter_by_sidebar(frames, selected_assets, selected_sectors)

    render_kpis(filtered)

    tab_resumo, tab_rankings, tab_alertas, tab_historico, tab_base = st.tabs(
        ["Resumo", "Rankings", "Alertas", "Historico", "Base completa"]
    )

    with tab_resumo:
        render_table("Resumo / Dashboard", filtered["dashboard"], "A aba dashboard esta vazia.")

    with tab_rankings:
        left, right = st.columns(2)
        with left:
            render_ranking_chart("Top 10 aberturas", filtered["ranking_aberturas"], "#C0392B")
            render_table(
                "Top 10 aberturas",
                filtered["ranking_aberturas"].head(10),
                "A aba ranking_aberturas esta vazia.",
            )
        with right:
            render_ranking_chart("Top 10 fechamentos", filtered["ranking_fechamentos"], "#1F7A4D")
            render_table(
                "Top 10 fechamentos",
                filtered["ranking_fechamentos"].head(10),
                "A aba ranking_fechamentos esta vazia.",
            )

    with tab_alertas:
        render_table("Alertas", filtered["alertas"], "Nao ha alertas para os filtros atuais.")

    with tab_historico:
        render_history(filtered["historico"], selected_assets)

    with tab_base:
        selected_sheet = st.selectbox(
            "Aba",
            options=list(filtered.keys()),
            format_func=lambda item: item.replace("_", " ").title(),
        )
        render_table(
            selected_sheet.replace("_", " ").title(),
            filtered[selected_sheet],
            "Nao ha dados para exibir nesta aba.",
        )


if __name__ == "__main__":
    main()
