# app.py
# Dashboard GTS 3C Brewing – Limpio, robusto y con filtros
# Requisitos: pip install streamlit pandas plotly openpyxl

import os
import unicodedata
import pandas as pd
import streamlit as st
import plotly.express as px


# =========================
# Utilidades de limpieza
# =========================
def norm_text(x: str) -> str:
    """Normaliza texto: quita acentos, baja a minúsculas y recorta."""
    s = str(x)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("utf8", "ignore")
    return s.strip().lower()


def as_percent(series: pd.Series) -> pd.Series:
    """Convierte una serie a porcentaje en rango 0..1 detectando si está en 0..100."""
    s = pd.to_numeric(series, errors="coerce")
    if s.dropna().mean() > 1.5:  # si parece 0..100, lo pasamos a 0..1
        s = s / 100.0
    return s


# =========================
# Carga de datos robusta
# =========================
def load_dashboard_df(xlsx_path: str) -> pd.DataFrame:
    """Lee la hoja 'Dashboard' y devuelve columnas estandarizadas: Área, Cumplimiento (0..1)."""
    df_raw = pd.read_excel(xlsx_path, sheet_name="Dashboard", header=0)

    # Detectar columna de áreas (texto) y columna de cumplimiento (numérica)
    area_col, value_col = None, None
    best_text_count, best_num_count = -1, -1

    for col in df_raw.columns:
        s = df_raw[col]
        # Contar "texto no-vacío"
        text_count = s.astype(str).replace("nan", pd.NA).dropna().shape[0]
        # Contar numéricos válidos
        num_count = pd.to_numeric(s, errors="coerce").dropna().shape[0]

        # Elegimos como Área la columna con MÁS texto y MENOS números
        if text_count > best_text_count and num_count < text_count * 0.4:
            best_text_count, area_col = text_count, col

        # Elegimos como Cumplimiento la columna con MÁS números
        if num_count > best_num_count:
            best_num_count, value_col = num_count, col

    df = pd.DataFrame({
        "Área": df_raw[area_col],
        "Cumplimiento": as_percent(df_raw[value_col])
    })

    # Limpiar filas inválidas
    df = df.dropna(subset=["Área", "Cumplimiento"])
    # Quitar filas-resumen típicas
    df = df[~df["Área"].astype(str).str.contains("total", case=False, na=False)]
    return df


ESTADOS_CANON = ["retrasada", "en proceso", "no iniciada", "completa", "cancelada"]
ESTADOS_SET = set(ESTADOS_CANON + ["abierto", "abierta", "cerrado", "cerrada"])


def load_action_log(xlsx_path: str) -> pd.DataFrame:
    """Lee 'ACTION LOG' y detecta columnas de Responsable y Estado de forma flexible."""
    df = pd.read_excel(xlsx_path, sheet_name="ACTION LOG", header=0)

    # Detectar columna de estado: la que más coincide con ESTADOS_SET
    estado_col, max_hits = None, -1
    for col in df.columns:
        hits = df[col].astype(str).map(norm_text).isin(ESTADOS_SET).sum()
        if hits > max_hits:
            max_hits, estado_col = hits, col

    # Detectar columna de responsable: la que tenga más valores únicos (excluyendo estado)
    resp_col, best_unique = None, -1
    for col in df.columns:
        if col == estado_col:
            continue
        uniq = df[col].astype(str).nunique(dropna=True)
        if uniq > best_unique:
            best_unique, resp_col = uniq, col

    # Normalizamos nombres
    df = df.rename(columns={resp_col: "Responsable", estado_col: "Estado"})
    # Limpiar
    df["Estado"] = df["Estado"].astype(str)
    df["Estado_norm"] = df["Estado"].map(norm_text)

    return df[["Responsable", "Estado", "Estado_norm"]].dropna(how="all")


def load_plan_accion(xlsx_path: str) -> pd.DataFrame:
    """
    Lee 'Plan de Acción (2)' y extrae conteos por estado.
    Busca celdas con los nombres de estado y toma el número en la/s columna/s vecinas.
    """
    df = pd.read_excel(xlsx_path, sheet_name="Plan de Acción (2)", header=0)
    counts = {e: 0 for e in ESTADOS_CANON}

    n_rows, n_cols = df.shape
    for r in range(n_rows):
        for c in range(n_cols):
            label = norm_text(df.iat[r, c])
            if label in ESTADOS_SET:
                # Buscar un número a la derecha (hasta 3 columnas)
                num = None
                for k in range(1, 4):
                    if c + k < n_cols:
                        num_try = pd.to_numeric(df.iat[r, c + k], errors="coerce")
                        if pd.notnull(num_try):
                            num = int(num_try)
                            break
                # Mapear abiertos/cerrados a categorías si fuera el caso
                if label in ["abierto", "abierta"]:
                    # No entra en el pie, pero lo usamos como KPI de abiertas si aparece
                    counts.setdefault("abierto", 0)
                    counts["abierto"] = counts["abierto"] + (num or 0)
                elif label in ["cerrado", "cerrada"]:
                    counts.setdefault("cerrado", 0)
                    counts["cerrado"] = counts["cerrado"] + (num or 0)
                else:
                    counts[label] = num or counts.get(label, 0)

    plan_df = pd.DataFrame({
        "Estado": [e.title() for e in ESTADOS_CANON],
        "Cantidad": [counts.get(e, 0) for e in ESTADOS_CANON]
    })
    return plan_df


# =========================
# UI – Streamlit
# =========================
st.set_page_config(page_title="Dashboard GTS 3C Brewing", layout="wide")

st.title("📊 Dashboard de Seguridad - GTS 3C Brewing")
st.markdown("Visualización de métricas de seguridad, cumplimiento y gestión de acciones.")

# Uploader opcional para facilitar pruebas / cambios de archivo
st.sidebar.header("⚙️ Fuente de datos")
uploaded = st.sidebar.file_uploader("Sube el Excel (.xlsx)", type=["xlsx"])
default_path = "GTS_3C_Brewing.xlsx"
xlsx_path = default_path

if uploaded:
    xlsx_path = uploaded
else:
    if not os.path.exists(default_path):
        st.error("No se encontró el archivo '01. GTS 3C Brewing.xlsx' en esta carpeta. Sube el Excel en la barra lateral.")
        st.stop()

# Cargar datasets
try:
    dashboard_df = load_dashboard_df(xlsx_path)
except Exception as e:
    st.error(f"No se pudo leer la hoja 'Dashboard'. Detalle: {e}")
    st.stop()

try:
    action_log = load_action_log(xlsx_path)
except Exception:
    # Si falla, seguimos sin esa sección
    action_log = pd.DataFrame(columns=["Responsable", "Estado", "Estado_norm"])

try:
    plan_df = load_plan_accion(xlsx_path)
except Exception:
    plan_df = pd.DataFrame({"Estado": ESTADOS_CANON, "Cantidad": [0]*len(ESTADOS_CANON)})

# =========================
# Filtros
# =========================
st.sidebar.header("🎛️ Filtros")
areas_all = sorted(dashboard_df["Área"].astype(str).unique())
areas_sel = st.sidebar.multiselect("Áreas", areas_all, default=areas_all)

if not action_log.empty:
    responsables_all = sorted(action_log["Responsable"].dropna().astype(str).unique())
    # Para no saturar, por defecto seleccionamos hasta 15
    default_resp = responsables_all if len(responsables_all) <= 15 else responsables_all[:15]
    resp_sel = st.sidebar.multiselect("Responsables", responsables_all, default=default_resp)
else:
    resp_sel = []

# Aplicar filtro a dashboard (áreas)
dash_filtrado = dashboard_df[dashboard_df["Área"].astype(str).isin(areas_sel)] if areas_sel else dashboard_df.copy()

# =========================
# KPIs
# =========================
col1, col2, col3 = st.columns(3)

with col1:
    if not dash_filtrado.empty:
        cumplimiento_global = dash_filtrado["Cumplimiento"].mean() * 100
        st.metric("Cumplimiento Global", f"{cumplimiento_global:.1f}%")
    else:
        st.metric("Cumplimiento Global", "—")

with col2:
    abiertas_count = 0
    if not action_log.empty:
        abiertas_count = action_log["Estado_norm"].isin(["abierto", "abierta", "en proceso", "retrasada", "no iniciada"]).sum()
    st.metric("Acciones Abiertas", int(abiertas_count))

with col3:
    completas_count = int(plan_df.loc[plan_df["Estado"].str.lower() == "completa", "Cantidad"].sum()) if not plan_df.empty else 0
    st.metric("Acciones Completas", completas_count)

# =========================
# Gráfico: Cumplimiento por Área
# =========================
st.subheader("Cumplimiento por Área")
if dash_filtrado.empty:
    st.info("No hay datos para las áreas seleccionadas.")
else:
    fig_area = px.bar(
        dash_filtrado.sort_values("Cumplimiento", ascending=False),
        x="Área",
        y="Cumplimiento",
        text=(dash_filtrado.sort_values("Cumplimiento", ascending=False)["Cumplimiento"] * 100).round(1).astype(str) + "%",
        labels={"Área": "Área", "Cumplimiento": "Cumplimiento"},
    )
    fig_area.update_traces(textposition="outside")
    fig_area.update_yaxes(range=[0, 1], tickformat=".0%")
    st.plotly_chart(fig_area, use_container_width=True)

# =========================
# Gráfico: Estado de Acciones (Plan)
# =========================
st.subheader("Estado de Acciones (Plan de Acción)")
if plan_df["Cantidad"].sum() == 0:
    st.info("No se encontraron cantidades por estado en 'Plan de Acción (2)'.")
else:
    fig_plan = px.pie(plan_df, values="Cantidad", names="Estado", hole=0.35)
    st.plotly_chart(fig_plan, use_container_width=True)

# =========================
# Gráfico: Acciones por Responsable (solo si hay ACTION LOG)
# =========================
st.subheader("Acciones por Responsable")
if action_log.empty:
    st.info("No se pudo interpretar 'ACTION LOG'.")
else:
    al = action_log.copy()
    if resp_sel:
        al = al[al["Responsable"].astype(str).isin(resp_sel)]
    if al.empty:
        st.info("No hay acciones para los responsables seleccionados.")
    else:
        fig_resp = px.histogram(
            al,
            x="Responsable",
            color="Estado",
            barmode="group",
            title="Acciones por responsable y estado"
        )
        st.plotly_chart(fig_resp, use_container_width=True)

# =========================
# Interpretación automática
# =========================
st.subheader("📌 Interpretación rápida")
if not dash_filtrado.empty:
    best = dash_filtrado.sort_values("Cumplimiento", ascending=False).head(1)
    worst = dash_filtrado.sort_values("Cumplimiento", ascending=True).head(1)
    best_area, best_val = best["Área"].iloc[0], best["Cumplimiento"].iloc[0] * 100
    worst_area, worst_val = worst["Área"].iloc[0], worst["Cumplimiento"].iloc[0] * 100

    if cumplimiento_global >= 90:
        st.success(f"Cumplimiento global alto ({cumplimiento_global:.1f}%). Mejor área: **{best_area}** ({best_val:.1f}%). Requiere atención: **{worst_area}** ({worst_val:.1f}%).")
    elif cumplimiento_global >= 75:
        st.warning(f"Cumplimiento global medio ({cumplimiento_global:.1f}%). Prioriza mejoras en **{worst_area}** ({worst_val:.1f}%).")
    else:
        st.error(f"Cumplimiento global bajo ({cumplimiento_global:.1f}%). Activa plan urgente en **{worst_area}** ({worst_val:.1f}%).")
else:
    st.write("Selecciona alguna área en los filtros para ver insights.")
