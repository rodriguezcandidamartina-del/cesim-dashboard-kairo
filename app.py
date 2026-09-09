from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from parser import procesar_archivo
from kpis_config import MI_EMPRESA, REGIONES, TECNOLOGIAS

st.set_page_config(
    page_title="CESIM Dashboard — KAIRO",
    page_icon=str(Path(__file__).resolve().parent / "assets" / "kairo_logo.png"),
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
:root {
    --kairo-blue: #2563eb;
    --kairo-blue-2: #0ea5e9;
    --kairo-navy: #0f172a;
    --kairo-muted: #64748b;
    --kairo-border: #dbe3ee;
    --kairo-soft: #f7faff;
}

html, body, [class*="css"] {
    font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
}

[data-testid="stAppViewContainer"] {
    background:
      radial-gradient(circle at 88% 0%, rgba(37,99,235,.10), transparent 28%),
      linear-gradient(180deg, #ffffff 0%, #f8fbff 100%);
}

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #f8fbff 0%, #eef5ff 100%);
    border-right: 1px solid var(--kairo-border);
}

[data-testid="stSidebar"] > div:first-child {
    padding-top: 1.2rem;
}

.block-container {
    padding-top: 1.1rem;
    padding-bottom: 2rem;
    max-width: 1600px;
}

.kairo-brand-wrap {
    display: flex;
    align-items: center;
    gap: 18px;
    margin-bottom: 8px;
}

.kairo-brand-title {
    font-size: 2.05rem;
    font-weight: 850;
    line-height: 1;
    color: var(--kairo-navy);
    letter-spacing: -.02em;
}

.kairo-brand-sub {
    font-size: .95rem;
    color: var(--kairo-muted);
    margin-top: 8px;
}

.kairo-strategy {
    border-left: 4px solid var(--kairo-blue);
    background: linear-gradient(90deg, rgba(37,99,235,.08), rgba(14,165,233,.04));
    border-radius: 12px;
    padding: 13px 16px;
    font-size: .95rem;
    color: #1e3a8a;
    margin: 8px 0 18px 0;
}

.kpi-card {
    border: 1px solid var(--kairo-border);
    border-radius: 16px;
    padding: 15px 16px;
    min-height: 112px;
    background: rgba(255,255,255,.88);
    box-shadow: 0 8px 24px rgba(15,23,42,.045);
}

.kpi-label {
    font-size: .80rem;
    color: var(--kairo-muted);
    margin-bottom: 7px;
}

.kpi-value {
    font-size: 1.55rem;
    font-weight: 800;
    line-height: 1.1;
    color: var(--kairo-navy);
    white-space: nowrap;
}

.kpi-note {
    font-size: .76rem;
    color: var(--kairo-muted);
    margin-top: 8px;
    min-height: 18px;
}

.region-card {
    border: 1px solid var(--kairo-border);
    border-radius: 15px;
    padding: 14px 16px;
    background: rgba(255,255,255,.90);
    box-shadow: 0 6px 18px rgba(15,23,42,.035);
    min-height: 100px;
}

.region-name {
    font-size: .84rem;
    color: var(--kairo-muted);
}

.region-share {
    font-size: 1.65rem;
    font-weight: 800;
    color: var(--kairo-navy);
    margin-top: 3px;
}

.region-rank {
    font-size: .77rem;
    color: var(--kairo-muted);
    margin-top: 5px;
}

.section-card {
    border: 1px solid var(--kairo-border);
    border-radius: 18px;
    padding: 18px;
    background: rgba(255,255,255,.88);
    box-shadow: 0 8px 22px rgba(15,23,42,.04);
    margin-bottom: 18px;
}

.sidebar-brand {
    text-align:center;
    padding: 4px 0 12px 0;
}

.sidebar-title {
    font-size: 1.2rem;
    font-weight: 800;
    color: var(--kairo-navy);
    margin-top: 4px;
}

.sidebar-subtitle {
    color: var(--kairo-muted);
    font-size: .84rem;
}

div[data-testid="stTabs"] button {
    font-size: .90rem;
    font-weight: 650;
}

div[data-testid="stTabs"] [data-baseweb="tab-list"] {
    gap: 10px;
}

div[data-testid="stTabs"] button[aria-selected="true"] {
    color: var(--kairo-blue);
}

[data-testid="stDataFrame"] {
    border: 1px solid var(--kairo-border);
    border-radius: 12px;
    overflow: hidden;
}

h1, h2, h3 {
    color: var(--kairo-navy);
}
</style>
""", unsafe_allow_html=True)

# ---------- Carga automática de rondas ----------
LOGO_PATH = Path(__file__).resolve().parent / "assets" / "kairo_logo.png"

with st.sidebar:
    st.image(str(LOGO_PATH), width="stretch")
    st.markdown(
        """
        <div class="sidebar-brand">
            <div class="sidebar-title">KAIRO · CESIM</div>
            <div class="sidebar-subtitle">Dashboard de análisis por ronda</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

CARPETA_RESULTADOS = Path(__file__).resolve().parent / "resultados"
CARPETA_RESULTADOS.mkdir(exist_ok=True)

archivos_locales = sorted(CARPETA_RESULTADOS.glob("*.xls")) + sorted(CARPETA_RESULTADOS.glob("*.xlsx"))

if not archivos_locales:
    st.markdown('<div class="kairo-title">CESIM Dashboard</div>', unsafe_allow_html=True)
    st.markdown('<div class="kairo-sub">KAIRO · Control de gestión y competencia</div>', unsafe_allow_html=True)
    st.error(
        "No encontré archivos de resultados en la carpeta `resultados`. "
        "Agregá ahí archivos como `results-r01.xlsx`, `results-r02.xlsx`, etc."
    )
    st.stop()

@st.cache_data(show_spinner=False)
def parsear_archivo_local(path_str, modified_time):
    path = Path(path_str)
    return procesar_archivo(path.name, path.read_bytes())

partes = [
    parsear_archivo_local(str(path), path.stat().st_mtime_ns)
    for path in archivos_locales
]

def juntar(clave):
    dfs = [p[clave] for p in partes if not p[clave].empty]
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

mercado = juntar("mercado")
shares = juntar("market_share")
finanzas = juntar("finanzas")
unmet = juntar("demanda_insatisfecha")

rondas = sorted(set(p["ronda"] for p in partes))
ronda_sel = st.sidebar.selectbox("Ronda", rondas, index=len(rondas) - 1)

equipos = sorted(set(
    list(mercado["equipo"].dropna().unique()) if not mercado.empty else []
) | set(
    list(shares["equipo"].dropna().unique()) if not shares.empty else []
) | set(
    list(finanzas["equipo"].dropna().unique()) if not finanzas.empty else []
))

st.sidebar.markdown("---")
st.sidebar.caption(f"{len(archivos_locales)} archivo(s) precargado(s) · {len(rondas)} ronda(s)")
st.sidebar.success("✓ Resultados cargados automáticamente")
with st.sidebar.expander("Archivos disponibles"):
    for path in archivos_locales:
        st.write(f"• {path.name}")


hcol1, hcol2 = st.columns([1.25, 3.75], vertical_alignment="center")
with hcol1:
    st.image(str(LOGO_PATH), width="stretch")
with hcol2:
    st.markdown(
        """
        <div class="kairo-brand-title">CESIM Dashboard</div>
        <div class="kairo-brand-sub">KAIRO · Control de gestión y competencia</div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        f"""
        <div class="kairo-strategy">
            <b>Ronda {ronda_sel}</b> · comparación con {max(len(equipos)-1, 0)} competidores ·
            Estrategia, análisis y resultados en una sola vista.
        </div>
        """,
        unsafe_allow_html=True,
    )

# ---------- Helpers ----------
def base_ronda(df):
    return df[df.ronda == ronda_sel].copy() if not df.empty else pd.DataFrame()

merc_r = base_ronda(mercado)
shares_r = base_ronda(shares)
fin_r = base_ronda(finanzas)
unmet_r = base_ronda(unmet)

def valor_fin(kpi, equipo=MI_EMPRESA, ronda=None):
    if finanzas.empty:
        return None
    rr = ronda_sel if ronda is None else ronda
    x = finanzas[
        (finanzas.ronda == rr) &
        (finanzas.equipo == equipo) &
        (finanzas.kpi == kpi)
    ]
    return None if x.empty else float(x.iloc[0].valor)

def valor_share(region, tecnologia="Total", equipo=MI_EMPRESA, ronda=None):
    if shares.empty:
        return None
    rr = ronda_sel if ronda is None else ronda
    x = shares[
        (shares.ronda == rr) &
        (shares.region == region) &
        (shares.tecnologia == tecnologia) &
        (shares.equipo == equipo)
    ]
    return None if x.empty else float(x.iloc[0].market_share_pct)

def ranking_share(region, tecnologia="Total", equipo=MI_EMPRESA):
    if shares_r.empty:
        return None, None
    x = shares_r[
        (shares_r.region == region) &
        (shares_r.tecnologia == tecnologia)
    ].sort_values("market_share_pct", ascending=False).reset_index(drop=True)
    if x.empty or equipo not in x.equipo.values:
        return None, None
    puesto = int(x.index[x.equipo == equipo][0]) + 1
    return puesto, len(x)

def fmt_money_k(v):
    # Los estados financieros CESIM vienen en miles USD.
    if v is None or pd.isna(v):
        return "—"
    av = abs(v)
    if av >= 1_000_000:
        return f"US$ {v/1_000_000:.2f} B"
    if av >= 1_000:
        return f"US$ {v/1_000:.1f} M"
    return f"US$ {v:,.0f} mil".replace(",", ".")

def fmt_pct(v):
    return "—" if v is None or pd.isna(v) else f"{v:.2f}%"

def fmt_units(v):
    if v is None or pd.isna(v):
        return "—"
    if abs(v) >= 1000:
        return f"{v/1000:.2f} M"
    return f"{v:,.0f} mil".replace(",", ".")

def delta_vs_anterior(kpi):
    idx = rondas.index(ronda_sel)
    if idx == 0:
        return ""
    actual = valor_fin(kpi, ronda=ronda_sel)
    prev = valor_fin(kpi, ronda=rondas[idx - 1])
    if actual is None or prev in (None, 0):
        return ""
    return f"{((actual / prev) - 1) * 100:+.1f}% vs R{rondas[idx - 1]}"

def render_kpi(col, label, value, note=""):
    col.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

def ranking_bar(df, metric, titulo, etiqueta, suffix="", prefix=""):
    x = df[["equipo", metric]].dropna().copy()
    if x.empty:
        return None
    x = x.sort_values(metric, ascending=True)
    x["destacado"] = x["equipo"].apply(lambda e: "KAIRO" if e == MI_EMPRESA else "Competencia")
    x["texto"] = x[metric].apply(
        lambda v: f"{prefix}{v:,.2f}{suffix}".replace(",", ".")
    )
    fig = px.bar(
        x,
        x=metric,
        y="equipo",
        orientation="h",
        color="destacado",
        text="texto",
        title=titulo,
        labels={metric: etiqueta, "equipo": "", "destacado": ""},
    )
    fig.update_traces(textposition="outside")
    fig.update_layout(
        showlegend=True,
        legend_orientation="h",
        legend_y=1.08,
        height=390,
        margin=dict(l=10, r=35, t=65, b=20),
    )
    return fig

# ---------- KPIs superiores ----------
global_ms = valor_share("Global", "Total")
global_rank = ranking_share("Global", "Total")

demanda_insat_kairo = None
if not unmet_r.empty:
    aux = unmet_r[unmet_r.equipo == MI_EMPRESA]
    if not aux.empty:
        demanda_insat_kairo = float(aux.demanda_insatisfecha.sum())

cols = st.columns(6)
render_kpi(cols[0], "Ingresos", fmt_money_k(valor_fin("Ingresos por ventas")), delta_vs_anterior("Ingresos por ventas"))
render_kpi(cols[1], "EBITDA", fmt_money_k(valor_fin("EBITDA")), delta_vs_anterior("EBITDA"))
render_kpi(cols[2], "Beneficio", fmt_money_k(valor_fin("Beneficio de la ronda")), delta_vs_anterior("Beneficio de la ronda"))

ms_note = f"#{global_rank[0]} de {global_rank[1]}" if global_rank[0] is not None else ""
render_kpi(cols[3], "Market share global", fmt_pct(global_ms), ms_note)
render_kpi(cols[4], "ROE", fmt_pct(valor_fin("ROE, %")), delta_vs_anterior("ROE, %"))
render_kpi(cols[5], "Demanda insatisfecha", fmt_units(demanda_insat_kairo), "")

tabs = st.tabs([
    "🏠 Resumen",
    "💰 Comercial",
    "📦 Demanda",
    "📱 Producto",
    "📣 Marketing",
    "🏭 Producción",
    "💵 Finanzas",
    "⚔️ Competencia",
    "📈 Evolución",
])

# ---------- Resumen ----------
with tabs[0]:
    st.subheader("Cuota de mercado por región — KAIRO vs competidores")
    st.caption(
        "Comparación directa de la cuota total de mercado de las 7 empresas. "
        "EE.UU., Asia y Europa se muestran por separado."
    )

    total_shares = shares_r[
        (shares_r.tecnologia == "Total") &
        (shares_r.region.isin(["EE.UU.", "Asia", "Europa"]))
    ].copy()

    if not total_shares.empty:
        # Tabla comparativa: empresa x región
        pivot = total_shares.pivot(
            index="equipo",
            columns="region",
            values="market_share_pct"
        )
        orden_cols = [c for c in ["EE.UU.", "Asia", "Europa"] if c in pivot.columns]
        pivot = pivot[orden_cols]

        # Ordenar por promedio regional para una lectura estable.
        pivot["Promedio 3 regiones"] = pivot.mean(axis=1)
        pivot = pivot.sort_values("Promedio 3 regiones", ascending=False)
        pivot_tabla = pivot.drop(columns=["Promedio 3 regiones"])

        st.markdown("### Comparación completa")
        st.dataframe(
            pivot_tabla.style.format("{:.2f}%"),
            width="stretch",
        )

        # Tres gráficos visibles simultáneamente.
        st.markdown("### Ranking de cuota de mercado")
        chart_cols = st.columns(3)

        for col, region in zip(chart_cols, ["EE.UU.", "Asia", "Europa"]):
            sub = total_shares[total_shares.region == region].copy()
            fig = ranking_bar(
                sub,
                "market_share_pct",
                f"{region}",
                "Market share (%)",
                suffix="%",
            )
            if fig:
                fig.update_layout(
                    height=430,
                    showlegend=False,
                    margin=dict(l=5, r=30, t=55, b=20),
                )
                col.plotly_chart(fig, width="stretch")

        # Tarjetas KAIRO: cuota y posición en cada región.
        st.markdown("### Posición de KAIRO")
        region_cols = st.columns(3)

        for col, region in zip(region_cols, ["EE.UU.", "Asia", "Europa"]):
            v = valor_share(region, "Total")
            puesto, total = ranking_share(region, "Total")
            rank_txt = f"#{puesto} de {total}" if puesto is not None else "Sin dato"
            col.markdown(
                f"""
                <div class="region-card">
                    <div class="region-name">{region}</div>
                    <div class="region-share">{fmt_pct(v)}</div>
                    <div class="region-rank">{rank_txt}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        # Desglose por tecnología dentro de cada región.
        st.markdown("### Market share por tecnología")
        st.caption(
            "Elegí una región para ver cuánto tiene cada empresa en Tec 1, Tec 2, Tec 3 y Tec 4."
        )

        region_tec = st.selectbox(
            "Región",
            ["EE.UU.", "Asia", "Europa"],
            key="resumen_share_region_tec",
        )

        tech_shares = shares_r[
            (shares_r.region == region_tec) &
            (shares_r.tecnologia.isin(TECNOLOGIAS))
        ].copy()

        if not tech_shares.empty:
            tech_pivot = tech_shares.pivot(
                index="equipo",
                columns="tecnologia",
                values="market_share_pct"
            )
            tech_cols = [t for t in TECNOLOGIAS if t in tech_pivot.columns]
            tech_pivot = tech_pivot[tech_cols]
            st.dataframe(
                tech_pivot.style.format("{:.2f}%"),
                width="stretch",
            )

            tec_cols = st.columns(4)
            for col, tech in zip(tec_cols, TECNOLOGIAS):
                sub = tech_shares[tech_shares.tecnologia == tech].copy()
                if sub.empty:
                    col.info(f"{tech}: sin datos")
                    continue
                fig = ranking_bar(
                    sub,
                    "market_share_pct",
                    tech,
                    "Market share (%)",
                    suffix="%",
                )
                if fig:
                    fig.update_layout(
                        height=390,
                        showlegend=False,
                        margin=dict(l=5, r=25, t=50, b=15),
                    )
                    col.plotly_chart(fig, width="stretch")
    else:
        st.info("No encontré las cuotas de mercado regionales en este archivo.")

# ---------- Comercial ----------
with tabs[1]:
    st.subheader("Precio, ventas y posición competitiva")

    f1, f2 = st.columns(2)
    reg = f1.selectbox("Región", REGIONES, key="com_reg")
    tec = f2.selectbox("Tecnología", TECNOLOGIAS, key="com_tec")

    sub = merc_r[
        (merc_r.region == reg) &
        (merc_r.tecnologia == tec)
    ].copy()

    share_sub = shares_r[
        (shares_r.region == reg) &
        (shares_r.tecnologia == tec)
    ][["equipo", "market_share_pct"]]

    if sub.empty:
        st.info("No hay oferta para esa combinación en esta ronda.")
    else:
        sub = sub.merge(share_sub, on="equipo", how="left")

        kairo = sub[sub.equipo == MI_EMPRESA]
        metric_cols = st.columns(5)

        if not kairo.empty:
            kr = kairo.iloc[0]
            moneda = str(kr.get("moneda", ""))
            render_kpi(metric_cols[0], "Precio KAIRO", f"{moneda} {kr.precio:,.0f}".replace(",", "."))
            render_kpi(metric_cols[1], "Ventas KAIRO", fmt_units(kr.ventas))
            render_kpi(metric_cols[2], "Market share", fmt_pct(kr.market_share_pct))
            render_kpi(metric_cols[3], "Características", f"{kr.caracteristicas:.0f}" if pd.notna(kr.caracteristicas) else "—")
            render_kpi(metric_cols[4], "Marketing", str(kr.marketing) if pd.notna(kr.marketing) else "—")

        st.markdown("### Rankings")
        rank_tabs = st.tabs(["💲 Precio", "🛒 Ventas", "📊 Market Share"])

        with rank_tabs[0]:
            moneda = str(sub["moneda"].dropna().iloc[0]) if not sub["moneda"].dropna().empty else ""
            fig = ranking_bar(
                sub,
                "precio",
                f"Ranking de precios · {reg} · {tec}",
                f"Precio ({moneda})",
                prefix=f"{moneda} ",
            )
            if fig:
                st.plotly_chart(fig, width="stretch")
            st.caption("Mayor barra = precio más alto. KAIRO aparece separado visualmente de la competencia.")

        with rank_tabs[1]:
            fig = ranking_bar(
                sub,
                "ventas",
                f"Ranking de ventas · {reg} · {tec}",
                "Miles de unidades",
            )
            if fig:
                st.plotly_chart(fig, width="stretch")

        with rank_tabs[2]:
            fig = ranking_bar(
                sub,
                "market_share_pct",
                f"Ranking de market share · {reg} · {tec}",
                "Market share (%)",
                suffix="%",
            )
            if fig:
                st.plotly_chart(fig, width="stretch")

        st.markdown("### Comparación integral")
        fig = px.scatter(
            sub,
            x="precio",
            y="ventas",
            size="caracteristicas",
            color="equipo",
            text="equipo",
            hover_data=["market_share_pct", "marketing", "demanda"],
            title=f"Precio vs ventas · tamaño de burbuja = características · {reg} · {tec}",
            labels={
                "precio": f"Precio ({sub.moneda.iloc[0]})",
                "ventas": "Ventas (miles)",
                "equipo": "Empresa",
            },
        )
        fig.update_traces(textposition="top center")
        fig.update_layout(height=500)
        st.plotly_chart(fig, width="stretch")

        mostrar = sub[
            ["equipo", "precio", "moneda", "ventas", "demanda",
             "market_share_pct", "caracteristicas", "marketing"]
        ].copy()

        mostrar = mostrar.rename(columns={
            "equipo": "Empresa",
            "precio": "Precio",
            "moneda": "Moneda",
            "ventas": "Ventas (miles)",
            "demanda": "Demanda (miles)",
            "market_share_pct": "Market share %",
            "caracteristicas": "Características",
            "marketing": "Marketing",
        })

        mostrar = mostrar.sort_values("Market share %", ascending=False)

        st.dataframe(
            mostrar.style.format({
                "Precio": "{:,.0f}",
                "Ventas (miles)": "{:,.3f}",
                "Demanda (miles)": "{:,.3f}",
                "Market share %": "{:.2f}%",
                "Características": "{:.0f}",
            }),
            width="stretch",
            hide_index=True,
        )

# ---------- Demanda ----------
with tabs[2]:
    st.subheader("Demanda, ventas y demanda insatisfecha")

    d1, d2 = st.columns(2)
    reg = d1.selectbox("Región", ["Todas"] + REGIONES, key="dem_reg")
    tec = d2.selectbox("Tecnología", ["Todas"] + TECNOLOGIAS, key="dem_tec")

    sub = merc_r.copy()
    if reg != "Todas":
        sub = sub[sub.region == reg]
    if tec != "Todas":
        sub = sub[sub.tecnologia == tec]

    if not unmet_r.empty:
        sub = sub.merge(
            unmet_r,
            on=["ronda", "region", "tecnologia", "equipo"],
            how="left",
        )
    else:
        sub["demanda_insatisfecha"] = sub["demanda_insatisfecha_calculada"]

    if "demanda_insatisfecha_calculada" in sub:
        sub["demanda_insatisfecha"] = sub["demanda_insatisfecha"].fillna(
            sub["demanda_insatisfecha_calculada"]
        )

    tabla = sub[
        ["equipo", "region", "tecnologia", "ventas", "demanda",
         "demanda_insatisfecha", "cobertura_demanda_pct"]
    ].copy()

    st.dataframe(
        tabla.sort_values("demanda_insatisfecha", ascending=False),
        width="stretch",
        hide_index=True,
    )

    if not sub.empty:
        agg = sub.groupby("equipo", as_index=False)[["ventas", "demanda_insatisfecha"]].sum()
        long = agg.melt("equipo", var_name="concepto", value_name="miles_unidades")
        fig = px.bar(
            long,
            x="equipo",
            y="miles_unidades",
            color="concepto",
            barmode="stack",
            title="Ventas + demanda insatisfecha",
        )
        st.plotly_chart(fig, width="stretch")

# ---------- Producto ----------
with tabs[3]:
    st.subheader("Características y posicionamiento de producto")

    p1, p2 = st.columns(2)
    reg = p1.selectbox("Región", REGIONES, key="prod_reg")
    tec = p2.selectbox("Tecnología", TECNOLOGIAS, key="prod_tec")

    sub = merc_r[(merc_r.region == reg) & (merc_r.tecnologia == tec)].copy()

    if not sub.empty:
        fig = ranking_bar(
            sub,
            "caracteristicas",
            f"Ranking de características · {reg} · {tec}",
            "Cantidad de características",
        )
        if fig:
            st.plotly_chart(fig, width="stretch")

        fig = px.scatter(
            sub,
            x="caracteristicas",
            y="precio",
            size="ventas",
            text="equipo",
            hover_data=["marketing", "demanda"],
            title=f"Precio vs características · {reg} · {tec}",
        )
        fig.update_traces(textposition="top center")
        st.plotly_chart(fig, width="stretch")
    else:
        st.info("Esa tecnología todavía no tiene oferta en esta ronda.")

# ---------- Marketing ----------
with tabs[4]:
    st.subheader("Estrategia de marketing")
    sub = merc_r.dropna(subset=["marketing"]).copy()

    if not sub.empty:
        st.dataframe(
            sub[["equipo", "region", "tecnologia", "marketing", "precio", "caracteristicas", "ventas"]]
            .sort_values(["region", "tecnologia", "ventas"], ascending=[True, True, False]),
            width="stretch",
            hide_index=True,
        )

        conteo = sub.groupby(["equipo", "marketing"], as_index=False).size()
        fig = px.bar(
            conteo,
            x="equipo",
            y="size",
            color="marketing",
            title="Enfoques de marketing utilizados por empresa",
        )
        st.plotly_chart(fig, width="stretch")
    else:
        st.info("No encontré enfoques de marketing.")

# ---------- Producción ----------
with tabs[5]:
    st.subheader("Producción y restricciones")
    st.caption("Se muestra la demanda insatisfecha reportada por CESIM como señal principal de restricción.")

    if not unmet_r.empty:
        st.dataframe(
            unmet_r.sort_values("demanda_insatisfecha", ascending=False),
            width="stretch",
            hide_index=True,
        )

        fig = px.bar(
            unmet_r,
            x="equipo",
            y="demanda_insatisfecha",
            color="region",
            facet_col="tecnologia",
            title="Demanda insatisfecha por tecnología y región",
        )
        st.plotly_chart(fig, width="stretch")
    else:
        st.info("No encontré demanda insatisfecha logística.")

# ---------- Finanzas ----------
with tabs[6]:
    st.subheader("Finanzas — comparación global")

    if not fin_r.empty:
        kpi = st.selectbox("Indicador", sorted(fin_r.kpi.unique()), key="fin_kpi")
        sub = fin_r[fin_r.kpi == kpi].copy()

        fig = ranking_bar(
            sub.rename(columns={"valor": "metrica"}),
            "metrica",
            f"{kpi} · Ronda {ronda_sel}",
            kpi,
        )
        if fig:
            st.plotly_chart(fig, width="stretch")

        st.dataframe(
            sub.sort_values("valor", ascending=False),
            width="stretch",
            hide_index=True,
        )
    else:
        st.info("No encontré indicadores financieros.")

# ---------- Competencia ----------
with tabs[7]:
    st.subheader("KAIRO vs competencia")

    c1, c2 = st.columns(2)
    reg = c1.selectbox("Región", REGIONES, key="comp_reg")
    tec = c2.selectbox("Tecnología", TECNOLOGIAS, key="comp_tec")

    sub = merc_r[(merc_r.region == reg) & (merc_r.tecnologia == tec)].copy()
    share_sub = shares_r[
        (shares_r.region == reg) &
        (shares_r.tecnologia == tec)
    ][["equipo", "market_share_pct"]]

    if not sub.empty:
        sub = sub.merge(share_sub, on="equipo", how="left")

        fig = px.scatter(
            sub,
            x="precio",
            y="ventas",
            size="caracteristicas",
            color="market_share_pct",
            text="equipo",
            hover_data=["marketing", "demanda", "market_share_pct"],
            title=f"Mapa competitivo · {reg} · {tec}",
        )
        fig.update_traces(textposition="top center")
        st.plotly_chart(fig, width="stretch")

        comp_cols = st.columns(2)
        with comp_cols[0]:
            fig = ranking_bar(
                sub,
                "market_share_pct",
                "Ranking de market share",
                "Market share (%)",
                suffix="%",
            )
            if fig:
                st.plotly_chart(fig, width="stretch")

        with comp_cols[1]:
            fig = ranking_bar(
                sub,
                "ventas",
                "Ranking de ventas",
                "Miles de unidades",
            )
            if fig:
                st.plotly_chart(fig, width="stretch")
    else:
        st.info("No hay oferta para esa combinación.")

# ---------- Evolución ----------
with tabs[8]:
    st.subheader("Evolución entre rondas")

    if len(rondas) < 2:
        st.info("Cuando cargues dos o más rondas, acá vas a ver la evolución.")
    else:
        e1, e2 = st.columns(2)
        kpi = e1.selectbox("KPI financiero", sorted(finanzas.kpi.unique()), key="evol_fin")
        empresas_fin = sorted(finanzas.equipo.unique())
        default_idx = empresas_fin.index(MI_EMPRESA) if MI_EMPRESA in empresas_fin else 0
        eq = e2.selectbox("Empresa", empresas_fin, index=default_idx)

        sub = finanzas[
            (finanzas.kpi == kpi) &
            (finanzas.equipo == eq)
        ].sort_values("ronda")

        fig = px.line(
            sub,
            x="ronda",
            y="valor",
            markers=True,
            title=f"{kpi} · {eq}",
        )
        fig.update_xaxes(dtick=1)
        st.plotly_chart(fig, width="stretch")

        st.markdown("### Market share de KAIRO por región")
        sub_ms = shares[
            (shares.equipo == MI_EMPRESA) &
            (shares.tecnologia == "Total") &
            (shares.region.isin(["EE.UU.", "Asia", "Europa", "Global"]))
        ].copy()

        if not sub_ms.empty:
            fig = px.line(
                sub_ms,
                x="ronda",
                y="market_share_pct",
                color="region",
                markers=True,
                title="Evolución del market share de KAIRO",
                labels={"market_share_pct": "Market share (%)", "ronda": "Ronda"},
            )
            fig.update_xaxes(dtick=1)
            st.plotly_chart(fig, width="stretch")

with st.expander("Datos extraídos"):
    st.write("Mercado")
    st.dataframe(mercado, width="stretch", hide_index=True)
    st.write("Market share")
    st.dataframe(shares, width="stretch", hide_index=True)
    st.write("Finanzas")
    st.dataframe(finanzas, width="stretch", hide_index=True)
