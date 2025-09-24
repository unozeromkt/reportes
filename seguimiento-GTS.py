# app.py
# Dashboard – Seguimiento GTS 3C Safety (GF COL)
# Funciona con: "Seguimiento GTS 3C Safety GF COL.xlsx" en la misma carpeta
# Requisitos: pip install streamlit pandas plotly openpyxl

import os
import unicodedata
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px


# =========================
# Utilidades de limpieza
# =========================
def norm_text(x: str) -> str:
    """Normaliza texto (minúsculas, sin acentos, trim)."""
    s = str(x)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("utf8", "ignore")
    return s.strip().lower()


def as_percent(series: pd.Series) -> pd.Series:
    """Convierte serie a proporción 0..1 si detecta 0..100."""
    s = pd.to_numeric(series, errors="coerce")
    if s.dropna().mean() > 1.5:  # si parece estar en %
        s = s / 100.0
    return s


def weighted_mean(values: pd.Series, weights: pd.Series) -> float:
    """Promedio ponderado con manejo de NaN."""
    v = pd.to_numeric(values, errors="coerce")
    w = pd.to_numeric(weights, errors="coerce").fillna(0)
    mask = v.notna() & w.notna()
    if mask.sum() == 0:
        return np.nan
    wsum = w[mask].sum()
    return (v[mask] * w[mask]).sum() / wsum if wsum > 0 else v[mask].mean()


# =========================
# Carga – SEGUIMIENTO GTS (detalle por área)
# =========================
@st.cache_data(show_spinner=False)
def load_seguimiento_gts(xlsx_path: str) -> pd.DataFrame:
    """
    Lee la hoja 'SEGUIMIENTO GTS' (estructura típica de tu archivo).
    Columnas por índice (0-based) usadas:
      7: Category (se propaga con ffill)
      9: Area
      10: Requeridas
      11: Faltantes
      12: Cerradas
      13: Mandatorio
      14: Avance
    Devuelve columnas estándar: Category, Area, Requeridas, Faltantes, Cerradas, Mandatorio, Avance
    """
    df = pd.read_excel(xlsx_path, sheet_name="SEGUIMIENTO GTS", header=None)

    # Propagar categoría y levantar columnas de interés
    df["Category"] = df[7].ffill()
    out = df[[9, 10, 11, 12, 13, 14, "Category"]].copy()
    out = out.rename(columns={
        9: "Area",
        10: "Requeridas",
        11: "Faltantes",
        12: "Cerradas",
        13: "Mandatorio",
        14: "Avance"
    })

    # Coerciones numéricas
    for c in ["Requeridas", "Faltantes", "Cerradas", "Mandatorio", "Avance"]:
        out[c] = pd.to_numeric(out[c], errors="coerce")
    out["Avance"] = as_percent(out["Avance"])

    # Limpieza de filas inválidas / resumen
    out = out.dropna(subset=["Area", "Avance"])
    out["Area"] = out["Area"].astype(str).str.strip()
    out["Category"] = out["Category"].astype(str).str.strip()
    out = out[~out["Area"].str.contains("total", case=False, na=False)]

    return out


# =========================
# Carga – Plan de trabajo (robusto a columnas duplicadas)
# =========================
@st.cache_data(show_spinner=False)
def load_plan_trabajo(xlsx_path: str) -> pd.DataFrame:
    """
    Lee 'Plan de trabajo' aunque existan columnas duplicadas/variantes.
    Devuelve columnas estándar:
      GTS, Totales, Completadas, Faltantes, Avance, Pdte, Mes, CierreMes, FechaEntrega
    """
    raw = pd.read_excel(xlsx_path, sheet_name="Plan de trabajo", header=None)

    # Detectar fila de encabezados buscando palabras clave
    header_idx = None
    for i in range(min(40, len(raw))):
        row = [norm_text(x) for x in raw.iloc[i].tolist()]
        if ("gts" in row or "bloque" in row) and \
           "totales" in row and \
           ("completadas" in row or "completa" in row) and \
           ("faltantes" in row or "faltante" in row):
            header_idx = i
            break
    if header_idx is None:
        header_idx = 0  # fallback

    df = pd.read_excel(xlsx_path, sheet_name="Plan de trabajo", header=header_idx)
    df.columns = [norm_text(c) for c in df.columns]

    # Helpers para fusionar columnas duplicadas o variantes
    def pick_text(patterns):
        cols = [c for c in df.columns if any(p in c for p in patterns)]
        if not cols:
            return pd.Series([pd.NA] * len(df))
        block = df[cols].astype(str)
        s = block.bfill(axis=1).iloc[:, 0]
        s = s.where(s.str.strip().ne(""), pd.NA)
        return s

    def pick_num(patterns):
        cols = [c for c in df.columns if any(p in c for p in patterns)]
        if not cols:
            return pd.Series([np.nan] * len(df))
        block = df[cols].copy()
        for c in block.columns:
            block[c] = pd.to_numeric(block[c], errors="coerce")
        s = block.bfill(axis=1).iloc[:, 0]
        return s

    out = pd.DataFrame({
        "GTS":          pick_text(["gts", "bloque", "modulo"]),
        "Totales":      pick_num(["total"]),
        "Completadas":  pick_num(["completa"]),
        "Faltantes":    pick_num(["faltant", "pend"]),
        "Avance":       pick_num(["avance", "%"]),
        "Pdte":         pick_num(["pdte", "pend"]),
        "Mes":          pick_text(["mes"]),
        "CierreMes":    pick_num(["cierre", "mes"]),
        "FechaEntrega": pick_text(["fecha", "entrega"])
    })

    out = out.dropna(subset=["GTS"], how="all")
    out["GTS"] = out["GTS"].astype(str).str.strip()
    out["Avance"] = as_percent(out["Avance"])
    return out


# =========================
# Carga – Comparativo (Hoja2) opcional
# =========================
@st.cache_data(show_spinner=False)
def load_comp_hoja2(xlsx_path: str) -> pd.DataFrame:
    """
    Lee 'Hoja2' y arma un comparativo: Disciplina | LE | REAL | DIF.
    Soporta encabezados en distintas filas y columnas duplicadas/variantes.
    Si no existe o no se puede interpretar, devuelve DF vacío.
    """
    import numpy as np

    # 1) Intentar leer la hoja
    try:
        raw = pd.read_excel(xlsx_path, sheet_name="Hoja2", header=None)
    except Exception:
        return pd.DataFrame()

    # 2) Detectar fila de encabezados buscando 'LE' y 'REAL'
    header_idx = None
    for i in range(min(40, len(raw))):
        row = [norm_text(x) for x in raw.iloc[i].tolist()]
        if "le" in row and any("real" in r for r in row):
            header_idx = i
            break
    if header_idx is None:
        header_idx = 0  # fallback

    try:
        df = pd.read_excel(xlsx_path, sheet_name="Hoja2", header=header_idx)
    except Exception:
        return pd.DataFrame()

    # 3) Normalizar nombres
    df.columns = [norm_text(c) for c in df.columns]

    # --- helpers con manejo de duplicados ---
    def coalesce_text(patterns):
        cols = [c for c in df.columns if any(p in c for p in patterns)]
        if not cols:
            return pd.Series([pd.NA] * len(df))
        block = df.loc[:, cols].astype(str)
        # primer no-nulo por fila
        s = block.bfill(axis=1).iloc[:, 0]
        s = s.where(s.str.strip().ne(""), pd.NA)
        return s

    def coalesce_num(patterns):
        cols = [c for c in df.columns if any((p == c) or (p in c) for p in patterns)]
        if not cols:
            return pd.Series([np.nan] * len(df))
        block = df.loc[:, cols].copy()

        # Evitar problemas de nombres duplicados: hacemos únicas las columnas por índice
        block.columns = [f"{col}__{i}" for i, col in enumerate(block.columns)]

        # Convertir TODO el bloque a numérico columna por columna (por índice)
        for i in range(block.shape[1]):
            block.iloc[:, i] = pd.to_numeric(block.iloc[:, i], errors="coerce")

        # Tomar el primer valor no-nulo por fila
        s = block.bfill(axis=1).iloc[:, 0]
        return s

    # 4) Construir salida estándar
    disciplina = coalesce_text(["disciplina", "area", "bloque", "modulo", "categoria"])
    le   = coalesce_num(["le"])     # exacto 'le' o variantes
    real = coalesce_num(["real"])   # contiene 'real'
    dif  = coalesce_num(["dif"])    # contiene 'dif'

    out = pd.DataFrame({
        "Disciplina": disciplina,
        "LE": le,
        "REAL": real,
        "DIF": dif
    }).dropna(subset=["Disciplina"], how="all")

    return out



# =========================
# UI – Streamlit
# =========================
st.set_page_config(page_title="Dashboard Seguimiento GTS", layout="wide")
st.title("📊 Dashboard – Seguimiento GTS 3C Safety (GF COL)")
st.caption("Visualización de métricas de seguridad, cumplimiento y plan de trabajo.")

# Selector de archivo
st.sidebar.header("📁 Fuente de datos")
uploaded = st.sidebar.file_uploader("Sube el Excel (.xlsx)", type=["xlsx"])
default_path = "Seguimiento GTS 3C Safety GF COL.xlsx"
xlsx_path = uploaded if uploaded else (default_path if os.path.exists(default_path) else None)
if not xlsx_path:
    st.error("Sube el archivo o coloca 'Seguimiento GTS 3C Safety GF COL.xlsx' en esta carpeta.")
    st.stop()

# Carga de datos
try:
    areas = load_seguimiento_gts(xlsx_path)
except Exception as e:
    st.error(f"No se pudo leer 'SEGUIMIENTO GTS'. Detalle: {e}")
    st.stop()

plan = load_plan_trabajo(xlsx_path)
comp = load_comp_hoja2(xlsx_path)

# =========================
# Filtros
# =========================
st.sidebar.header("🎛️ Filtros")
cats_all = sorted(areas["Category"].dropna().unique().tolist())
areas_all = sorted(areas["Area"].dropna().unique().tolist())

cats_sel = st.sidebar.multiselect("Categorías (bloques)", cats_all, default=cats_all)
areas_sel = st.sidebar.multiselect("Áreas", areas_all, default=areas_all[: min(20, len(areas_all))])

areas_fil = areas.copy()
if cats_sel:
    areas_fil = areas_fil[areas_fil["Category"].isin(cats_sel)]
if areas_sel:
    areas_fil = areas_fil[areas_fil["Area"].isin(areas_sel)]

# =========================
# KPIs
# =========================
col1, col2, col3, col4 = st.columns(4)

with col1:
    cg = weighted_mean(areas_fil["Avance"], areas_fil["Requeridas"])
    if np.isnan(cg):
        cg = areas_fil["Avance"].mean()
    st.metric("Cumplimiento Global", f"{(cg * 100):.1f}%")

with col2:
    tot_req = pd.to_numeric(areas_fil["Requeridas"], errors="coerce").sum()
    tot_cer = pd.to_numeric(areas_fil["Cerradas"], errors="coerce").sum()
    st.metric("Ítems Cerrados / Requeridos", f"{int(tot_cer)} / {int(tot_req)}")

with col3:
    falt = pd.to_numeric(areas_fil["Faltantes"], errors="coerce").sum()
    st.metric("Faltantes (backlog)", int(falt))

with col4:
    mand_abiertos = int(((areas_fil["Mandatorio"] > 0) & (areas_fil["Faltantes"] > 0)).sum())
    st.metric("Mandatorios abiertos", mand_abiertos)

st.divider()

# =========================
# Tabs de visualización
# =========================
tab1, tab2, tab3, tab4 = st.tabs(["Cumplimientos", "Plan de trabajo", "Comparativo (Hoja2)", "Insights"])

with tab1:
    left, right = st.columns((1.3, 1))

    with left:
        st.subheader("Cumplimiento por Área")
        df_plot = areas_fil.sort_values("Avance", ascending=False).copy()
        if df_plot.empty:
            st.info("No hay datos para los filtros seleccionados.")
        else:
            fig = px.bar(
                df_plot,
                x="Area",
                y="Avance",
                text=(df_plot["Avance"] * 100).round(1).astype(str) + "%",
                labels={"Area": "Área", "Avance": "Cumplimiento"}
            )
            fig.update_traces(textposition="outside")
            fig.update_yaxes(range=[0, 1], tickformat=".0%")
            st.plotly_chart(fig, use_container_width=True)

    with right:
        st.subheader("Cumplimiento por Categoría (promedio)")
        cat = areas_fil.groupby("Category", as_index=False)["Avance"].mean().sort_values("Avance", ascending=True)
        if cat.empty:
            st.info("Sin datos.")
        else:
            fig2 = px.bar(
                cat,
                x="Avance",
                y="Category",
                orientation="h",
                text=(cat["Avance"] * 100).round(1).astype(str) + "%",
                labels={"Category": "Categoría", "Avance": "Cumplimiento"}
            )
            fig2.update_traces(textposition="outside")
            fig2.update_xaxes(range=[0, 1], tickformat=".0%")
            st.plotly_chart(fig2, use_container_width=True)

with tab2:
    st.subheader("Plan de trabajo – Totales / Completadas / Faltantes")
    if plan.empty or not {"Totales", "Completadas", "Faltantes"}.issubset(set(plan.columns)):
        st.info("No se pudieron detectar correctamente las columnas del 'Plan de trabajo'.")
    else:
        plan_plot = plan.copy()
        plan_plot["Completadas"] = pd.to_numeric(plan_plot["Completadas"], errors="coerce")
        plan_plot["Faltantes"] = pd.to_numeric(plan_plot["Faltantes"], errors="coerce")
        plan_plot["Totales"] = pd.to_numeric(plan_plot["Totales"], errors="coerce")

        fig3 = px.bar(
            plan_plot,
            x="GTS",
            y=["Completadas", "Faltantes"],
            barmode="stack",
            labels={"value": "Cantidad", "GTS": "Bloque", "variable": "Estado"}
        )
        st.plotly_chart(fig3, use_container_width=True)

        # Donut global
        tot_c = plan_plot["Completadas"].sum()
        tot_f = plan_plot["Faltantes"].sum()
        donut_df = pd.DataFrame({"Estado": ["Completadas", "Faltantes"], "Cantidad": [tot_c, tot_f]})
        fig4 = px.pie(donut_df, values="Cantidad", names="Estado", hole=0.4)
        st.plotly_chart(fig4, use_container_width=True)

with tab3:
    st.subheader("Comparativo LE vs REAL (si aplica)")
    if comp.empty:
        st.info("No se encontraron columnas comparables en 'Hoja2'.")
    else:
        fig5 = px.scatter(
            comp,
            x="LE",
            y="REAL",
            hover_name="Disciplina",
            trendline="ols",
            labels={"LE": "LE (Plan)", "REAL": "Real"}
        )
        st.plotly_chart(fig5, use_container_width=True)

with tab4:
    st.subheader("📌 Interpretaciones automáticas")
    if areas_fil.empty:
        st.info("Ajusta los filtros para ver insights.")
    else:
        worst = areas_fil.sort_values(
            ["Mandatorio", "Faltantes", "Avance"], ascending=[False, False, True]
        ).head(5)
        best = areas_fil.sort_values("Avance", ascending=False).head(3)

        st.markdown("**Top 5 áreas a priorizar (mandatorios/faltantes altos y menor % de avance):**")
        st.dataframe(worst[["Category", "Area", "Mandatorio", "Faltantes", "Requeridas", "Cerradas", "Avance"]])

        st.markdown("**Top 3 áreas con mejor desempeño:**")
        st.dataframe(best[["Category", "Area", "Requeridas", "Cerradas", "Faltantes", "Avance"]])

        # Comentario ejecutivo
        cg_txt = (cg * 100) if not np.isnan(cg) else float("nan")
        if not np.isnan(cg_txt):
            falt_total = int(pd.to_numeric(areas_fil["Faltantes"], errors="coerce").sum())
            if cg_txt >= 90:
                st.success(
                    f"Cumplimiento global alto ({cg_txt:.1f}%). Mantén el ritmo y cierra mandatorios abiertos ({int(((areas_fil['Mandatorio']>0)&(areas_fil['Faltantes']>0)).sum())})."
                )
            elif cg_txt >= 85:
                st.warning(
                    f"Cumplimiento global medio-alto ({cg_txt:.1f}%). Prioriza backlog crítico ({falt_total} ítems) y cierra mandatorios."
                )
            else:
                st.error(
                    f"Cumplimiento global bajo ({cg_txt:.1f}%). Concentrar recursos en las 5 áreas críticas listadas arriba."
                )
