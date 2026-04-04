import streamlit as st
import pandas as pd
import io
import shutil
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(
    page_title="Transpositor Excel",
    page_icon="⇄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Estilos ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}

/* Fondo general */
.stApp {
    background-color: #0f0f13;
    color: #e8e6f0;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background-color: #16151e;
    border-right: 1px solid #2a2838;
}
section[data-testid="stSidebar"] .stMarkdown p,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] .stRadio label {
    color: #b8b4cc !important;
    font-size: 0.82rem;
}

/* Título sidebar */
.sidebar-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.7rem;
    font-weight: 500;
    letter-spacing: 0.15em;
    color: #6c5ecf !important;
    text-transform: uppercase;
    margin-bottom: 0.5rem;
    padding-bottom: 0.5rem;
    border-bottom: 1px solid #2a2838;
}

/* Encabezado principal */
.main-header {
    display: flex;
    align-items: baseline;
    gap: 1rem;
    margin-bottom: 0.25rem;
}
.main-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.6rem;
    font-weight: 500;
    color: #e8e6f0;
    letter-spacing: -0.02em;
}
.main-tag {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.7rem;
    color: #6c5ecf;
    letter-spacing: 0.1em;
    background: #1e1b2e;
    padding: 2px 8px;
    border-radius: 2px;
    border: 1px solid #3a3560;
}
.main-sub {
    font-size: 0.88rem;
    color: #6b6880;
    margin-bottom: 2rem;
}

/* Cards / paneles */
.panel-card {
    background: #16151e;
    border: 1px solid #2a2838;
    border-radius: 6px;
    padding: 1.25rem 1.5rem;
    margin-bottom: 1rem;
}
.panel-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.65rem;
    letter-spacing: 0.12em;
    color: #6c5ecf;
    text-transform: uppercase;
    margin-bottom: 0.75rem;
}

/* Inputs */
.stTextInput input, .stNumberInput input, .stSelectbox select {
    background-color: #1e1b2e !important;
    border: 1px solid #2a2838 !important;
    border-radius: 4px !important;
    color: #e8e6f0 !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.85rem !important;
}
.stTextInput input:focus, .stNumberInput input:focus {
    border-color: #6c5ecf !important;
    box-shadow: 0 0 0 1px #6c5ecf33 !important;
}

/* Multiselect */
.stMultiSelect [data-baseweb="tag"] {
    background-color: #2a2060 !important;
    border: 1px solid #6c5ecf !important;
    border-radius: 2px !important;
}
.stMultiSelect [data-baseweb="tag"] span {
    color: #c4bfee !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.78rem !important;
}

/* Botones */
.stButton > button {
    background: #6c5ecf !important;
    color: #fff !important;
    border: none !important;
    border-radius: 4px !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.82rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.05em !important;
    padding: 0.5rem 1.5rem !important;
    transition: background 0.15s !important;
}
.stButton > button:hover {
    background: #5a4db8 !important;
}

/* Download button */
.stDownloadButton > button {
    background: #1e3a2f !important;
    color: #4ade80 !important;
    border: 1px solid #2d5a42 !important;
    border-radius: 4px !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.82rem !important;
    letter-spacing: 0.05em !important;
}
.stDownloadButton > button:hover {
    background: #2a4d3c !important;
}

/* Radio */
.stRadio [data-baseweb="radio"] span {
    color: #b8b4cc !important;
    font-size: 0.85rem !important;
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    background: transparent !important;
    border-bottom: 1px solid #2a2838 !important;
    gap: 0 !important;
}
.stTabs [data-baseweb="tab"] {
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.78rem !important;
    letter-spacing: 0.08em !important;
    color: #6b6880 !important;
    background: transparent !important;
    border-bottom: 2px solid transparent !important;
    padding: 0.5rem 1.25rem !important;
}
.stTabs [aria-selected="true"] {
    color: #c4bfee !important;
    border-bottom-color: #6c5ecf !important;
}

/* DataFrame */
.stDataFrame {
    border: 1px solid #2a2838 !important;
    border-radius: 4px !important;
}

/* Alerts */
.stSuccess {
    background: #0d2b1f !important;
    border: 1px solid #1a4a32 !important;
    color: #4ade80 !important;
    border-radius: 4px !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.82rem !important;
}
.stError {
    background: #2b0d0d !important;
    border: 1px solid #4a1a1a !important;
    border-radius: 4px !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.82rem !important;
}
.stInfo {
    background: #0d1a2b !important;
    border: 1px solid #1a324a !important;
    border-radius: 4px !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.82rem !important;
}
.stWarning {
    background: #2b1f0d !important;
    border: 1px solid #4a3a1a !important;
    border-radius: 4px !important;
}

/* Expander */
.streamlit-expanderHeader {
    background: #16151e !important;
    border: 1px solid #2a2838 !important;
    border-radius: 4px !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.78rem !important;
    color: #b8b4cc !important;
}

/* Métricas */
[data-testid="metric-container"] {
    background: #16151e;
    border: 1px solid #2a2838;
    border-radius: 6px;
    padding: 0.75rem 1rem;
}
[data-testid="metric-container"] label {
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.65rem !important;
    letter-spacing: 0.1em !important;
    color: #6b6880 !important;
    text-transform: uppercase !important;
}
[data-testid="metric-container"] [data-testid="stMetricValue"] {
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 1.4rem !important;
    color: #c4bfee !important;
}

/* Checkbox */
.stCheckbox label span {
    color: #b8b4cc !important;
    font-size: 0.85rem !important;
}

/* Selectbox label */
.stSelectbox label, .stMultiSelect label, .stTextInput label,
.stNumberInput label, .stRadio label div {
    color: #8884a0 !important;
    font-size: 0.78rem !important;
    font-family: 'IBM Plex Mono', monospace !important;
    letter-spacing: 0.05em !important;
}

/* File uploader */
[data-testid="stFileUploader"] {
    background: #16151e !important;
    border: 1px dashed #3a3560 !important;
    border-radius: 6px !important;
}
[data-testid="stFileUploader"] label {
    color: #8884a0 !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.8rem !important;
}

/* Divider */
hr {
    border-color: #2a2838 !important;
}

/* Scrollbar */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #0f0f13; }
::-webkit-scrollbar-thumb { background: #2a2838; border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: #3a3560; }
</style>
""", unsafe_allow_html=True)


# ── Funciones de procesamiento ────────────────────────────────────────────────

def apply_excel_styling(wb, sheet_name, do_headers, do_autofit, do_freeze):
    ws = wb[sheet_name]
    if do_headers:
        hfill = PatternFill("solid", fgColor="534AB7")
        hfont = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
        for cell in ws[1]:
            cell.fill = hfill
            cell.font = hfont
            cell.alignment = Alignment(horizontal="center", vertical="center")
    if do_autofit:
        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=8) + 4
            ws.column_dimensions[col[0].column_letter].width = min(max_len, 45)
    if do_freeze:
        ws.freeze_panes = "A2"


def process_dataframe(df, mode, params):
    if mode == "Ancho → Largo (melt)":
        id_cols   = params.get("id_cols", [])
        val_cols  = params.get("val_cols", [])
        var_name  = params.get("var_name", "Variable")
        val_name  = params.get("val_name", "Valor")
        if not val_cols:
            raise ValueError("Debes seleccionar al menos una columna a trasponer.")
        return df.melt(id_vars=id_cols, value_vars=val_cols,
                       var_name=var_name, value_name=val_name)

    elif mode == "Largo → Ancho (pivot)":
        idx  = params.get("pivot_index")
        cols = params.get("pivot_cols")
        vals = params.get("pivot_vals")
        agg  = params.get("agg_func", "sum")
        fill = params.get("fill_val", None)
        if not idx or not cols or not vals:
            raise ValueError("Para pivot debes especificar índice, columna de variables y columna de valores.")
        result = df.pivot_table(index=idx, columns=cols, values=vals, aggfunc=agg)
        if fill is not None and fill != "":
            try:
                result = result.fillna(float(fill))
            except ValueError:
                result = result.fillna(fill)
        result = result.reset_index()
        result.columns.name = None
        return result

    elif mode == "Transponer todo (.T)":
        sel_cols = params.get("val_cols", [])
        subset = df[sel_cols] if sel_cols else df
        result = subset.T.reset_index()
        result.columns = ["Columna"] + [f"Fila_{i}" for i in range(len(result.columns) - 1)]
        return result

    raise ValueError("Modo no reconocido.")


def to_excel_bytes(df, sheet_name, index, do_headers, do_autofit, do_freeze):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=index)
    buf.seek(0)
    wb = load_workbook(buf)
    apply_excel_styling(wb, sheet_name, do_headers, do_autofit, do_freeze)
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()


# ── UI ────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="main-header">
  <span class="main-title">⇄ Transpositor Excel</span>
  <span class="main-tag">v1.0</span>
</div>
<p class="main-sub">Transpone, pivotea y reestructura hojas Excel sin código — descarga el resultado al instante.</p>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="sidebar-title">Archivo de entrada</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Sube tu archivo Excel", type=["xlsx", "xls"],
                                label_visibility="collapsed")

    st.markdown("---")
    st.markdown('<div class="sidebar-title">Parámetros de lectura</div>', unsafe_allow_html=True)
    sheet_input = st.text_input("Nombre de hoja", value="Sheet1")
    header_row  = st.number_input("Fila de encabezados", min_value=1, value=1, step=1)
    skip_rows   = st.number_input("Saltar filas al inicio", min_value=0, value=0, step=1)
    nrows_val   = st.number_input("Leer máx. N filas (0 = todo)", min_value=0, value=0, step=1)

    st.markdown("---")
    st.markdown('<div class="sidebar-title">Archivo de salida</div>', unsafe_allow_html=True)
    out_filename = st.text_input("Nombre archivo resultado", value="resultado.xlsx")
    out_sheet    = st.text_input("Nombre hoja resultado", value="Resultado")
    include_idx  = st.checkbox("Incluir índice", value=False)

    st.markdown("---")
    st.markdown('<div class="sidebar-title">Formato de salida</div>', unsafe_allow_html=True)
    do_headers = st.checkbox("Encabezados con color", value=True)
    do_autofit = st.checkbox("Ajustar anchos automáticamente", value=True)
    do_freeze  = st.checkbox("Congelar primera fila", value=False)


# ── Contenido principal ───────────────────────────────────────────────────────
if uploaded is None:
    st.info("⬆  Sube un archivo Excel en el panel izquierdo para comenzar.")
    st.stop()

# Leer datos
try:
    nrows_arg = int(nrows_val) if nrows_val > 0 else None
    df_raw = pd.read_excel(
        uploaded,
        sheet_name=sheet_input,
        header=int(header_row) - 1,
        skiprows=int(skip_rows) if skip_rows > 0 else None,
        nrows=nrows_arg
    )
except Exception as e:
    st.error(f"Error al leer el archivo: {e}")
    st.stop()

all_columns = list(df_raw.columns)

# ── Métricas rápidas
c1, c2, c3, c4 = st.columns(4)
c1.metric("Filas", f"{len(df_raw):,}")
c2.metric("Columnas", f"{len(df_raw.columns):,}")
c3.metric("Hoja leída", sheet_input)
c4.metric("Celdas totales", f"{len(df_raw) * len(df_raw.columns):,}")

st.markdown("---")

# ── Tabs principales
tab1, tab2, tab3 = st.tabs(["01  DATOS ORIGINALES", "02  CONFIGURACIÓN", "03  RESULTADO"])

# ── Tab 1: Vista previa
with tab1:
    st.markdown('<div class="panel-label">Vista previa — primeras 50 filas</div>', unsafe_allow_html=True)

    # Filtro rápido
    with st.expander("Filtro rápido de filas"):
        fc1, fc2, fc3 = st.columns([2, 1, 2])
        with fc1:
            filter_col = st.selectbox("Columna", ["(ninguno)"] + all_columns, key="fcol")
        with fc2:
            filter_op = st.selectbox("Condición", ["==", "!=", ">", ">=", "<", "<="], key="fop")
        with fc3:
            filter_val = st.text_input("Valor", key="fval")

    df_preview = df_raw.copy()
    if filter_col != "(ninguno)" and filter_val.strip():
        try:
            fv = float(filter_val) if filter_val.replace(".", "").replace("-", "").isdigit() else filter_val
            df_preview = df_preview[df_preview[filter_col].apply(
                lambda x: eval(f"x {filter_op} fv", {"x": x, "fv": fv})
            )]
            st.caption(f"Mostrando {len(df_preview):,} filas tras el filtro.")
        except Exception as e:
            st.warning(f"Filtro no aplicado: {e}")

    st.dataframe(df_preview.head(50), use_container_width=True, height=320)

# ── Tab 2: Configuración
with tab2:
    st.markdown('<div class="panel-label">Modo de transposición</div>', unsafe_allow_html=True)
    mode = st.radio(
        "modo",
        ["Ancho → Largo (melt)", "Largo → Ancho (pivot)", "Transponer todo (.T)"],
        label_visibility="collapsed"
    )

    st.markdown("---")
    params = {}

    if mode == "Ancho → Largo (melt)":
        st.markdown('<div class="panel-label">Columnas identificadoras (no se trasponen)</div>', unsafe_allow_html=True)
        params["id_cols"] = st.multiselect(
            "id_cols", all_columns, label_visibility="collapsed",
            placeholder="Selecciona columnas ID (ej: ID, Nombre, Region...)"
        )
        remaining = [c for c in all_columns if c not in params["id_cols"]]
        st.markdown('<div class="panel-label">Columnas a trasponer</div>', unsafe_allow_html=True)
        params["val_cols"] = st.multiselect(
            "val_cols", remaining, default=remaining, label_visibility="collapsed",
            placeholder="Selecciona columnas a convertir en filas..."
        )
        mc1, mc2 = st.columns(2)
        with mc1:
            params["var_name"] = st.text_input("Nombre columna 'variable'", value="Variable")
        with mc2:
            params["val_name"] = st.text_input("Nombre columna 'valor'", value="Valor")

    elif mode == "Largo → Ancho (pivot)":
        pc1, pc2, pc3 = st.columns(3)
        with pc1:
            params["pivot_index"] = st.selectbox("Columna índice (ID)", all_columns)
        with pc2:
            params["pivot_cols"] = st.selectbox("Columna de variables", all_columns)
        with pc3:
            params["pivot_vals"] = st.selectbox("Columna de valores", all_columns)
        popt1, popt2 = st.columns(2)
        with popt1:
            params["agg_func"] = st.selectbox("Función de agregación",
                                               ["sum", "mean", "count", "max", "min", "first"])
        with popt2:
            params["fill_val"] = st.text_input("Rellenar vacíos con", value="0")

    elif mode == "Transponer todo (.T)":
        st.markdown('<div class="panel-label">Columnas a incluir (vacío = todas)</div>', unsafe_allow_html=True)
        params["val_cols"] = st.multiselect(
            "cols_T", all_columns, label_visibility="collapsed",
            placeholder="Vacío = usa todas las columnas"
        )

    st.markdown("---")
    st.markdown('<div class="panel-label">Ordenamiento del resultado</div>', unsafe_allow_html=True)
    so1, so2 = st.columns([3, 1])
    with so1:
        sort_col = st.text_input("Ordenar por columna (nombre exacto, dejar vacío = no ordenar)", value="")
    with so2:
        sort_asc = st.selectbox("Orden", ["Ascendente", "Descendente"])

    params["sort_col"] = sort_col.strip() if sort_col.strip() else None
    params["sort_asc"] = sort_asc == "Ascendente"

    # Filtro aplicado al procesamiento
    params["filter_col"] = st.session_state.get("fcol", "(ninguno)")
    params["filter_op"]  = st.session_state.get("fop", "==")
    params["filter_val"] = st.session_state.get("fval", "")

    st.markdown("---")
    run_btn = st.button("⇄  Aplicar transposición", use_container_width=True)

# ── Tab 3: Resultado
with tab3:
    if "df_result" not in st.session_state:
        st.info("Configura los parámetros en la pestaña 02 y haz clic en 'Aplicar transposición'.")
    else:
        df_res = st.session_state["df_result"]
        st.markdown('<div class="panel-label">Vista previa del resultado</div>', unsafe_allow_html=True)

        rm1, rm2, rm3 = st.columns(3)
        rm1.metric("Filas resultado", f"{len(df_res):,}")
        rm2.metric("Columnas resultado", f"{len(df_res.columns):,}")
        rm3.metric("Modo aplicado", st.session_state.get("applied_mode", "—"))

        st.dataframe(df_res.head(100), use_container_width=True, height=380)

        excel_bytes = to_excel_bytes(
            df_res, out_sheet, include_idx, do_headers, do_autofit, do_freeze
        )
        fname = out_filename if out_filename.endswith(".xlsx") else out_filename + ".xlsx"
        st.download_button(
            label="↓  Descargar resultado Excel",
            data=excel_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# ── Ejecutar transposición (fuera de los tabs para que actualice sesión)
if run_btn:
    try:
        df_work = df_raw.copy()

        # Aplicar filtro si existe
        fc = params.get("filter_col", "(ninguno)")
        fv_str = params.get("filter_val", "").strip()
        if fc != "(ninguno)" and fv_str:
            fop = params.get("filter_op", "==")
            try:
                fv2 = float(fv_str) if fv_str.replace(".", "").replace("-", "").isdigit() else fv_str
                df_work = df_work[df_work[fc].apply(
                    lambda x: eval(f"x {fop} fv2", {"x": x, "fv2": fv2})
                )]
            except Exception:
                pass

        df_result = process_dataframe(df_work, mode, params)

        # Ordenar
        sc = params.get("sort_col")
        if sc and sc in df_result.columns:
            df_result = df_result.sort_values(sc, ascending=params.get("sort_asc", True))

        st.session_state["df_result"]    = df_result
        st.session_state["applied_mode"] = mode.split("(")[0].strip()
        st.success(f"Transposición completada — {len(df_result):,} filas × {len(df_result.columns):,} columnas")
        st.rerun()

    except Exception as e:
        st.error(f"Error al procesar: {e}")
