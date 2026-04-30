"""
Comparador de Excel — Archivo A vs Archivo B
Compara estructura, orden de columnas y datos entre dos archivos Excel.
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import difflib

# ──────────────────────────────────────────────
# Configuración de página
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="Comparador de Excel — A vs B",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Comparador de Excel — Archivo A vs Archivo B")
st.markdown(
    "Carga los dos archivos para analizar **estructura**, **orden de columnas** y **datos**."
)

# ──────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes, label: str) -> dict[str, pd.DataFrame]:
    """Carga todas las hojas de un Excel y devuelve {nombre_hoja: DataFrame}."""
    try:
        xls = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
        return {sheet: xls.parse(sheet) for sheet in xls.sheet_names}
    except Exception as exc:
        st.error(f"Error al leer '{label}': {exc}")
        return {}


def compare_columns(df_legacy: pd.DataFrame, df_migrated: pd.DataFrame) -> dict:
    """Compara columnas entre dos DataFrames."""
    cols_legacy = list(df_legacy.columns)
    cols_migrated = list(df_migrated.columns)

    only_in_legacy = [c for c in cols_legacy if c not in cols_migrated]
    only_in_migrated = [c for c in cols_migrated if c not in cols_legacy]
    common = [c for c in cols_legacy if c in cols_migrated]

    # Orden de las columnas comunes
    order_legacy = [cols_legacy.index(c) for c in common]
    order_migrated = [cols_migrated.index(c) for c in common]
    order_diff = [
        c for c, ol, om in zip(common, order_legacy, order_migrated) if ol != om
    ]

    return {
        "cols_legacy": cols_legacy,
        "cols_migrated": cols_migrated,
        "only_in_legacy": only_in_legacy,
        "only_in_migrated": only_in_migrated,
        "common": common,
        "order_changed": order_diff,
        "order_ok": len(order_diff) == 0,
    }


def compare_dtypes(df_legacy: pd.DataFrame, df_migrated: pd.DataFrame, common_cols: list) -> pd.DataFrame:
    """Compara tipos de dato para columnas comunes."""
    rows = []
    for col in common_cols:
        dl = str(df_legacy[col].dtype)
        dm = str(df_migrated[col].dtype)
        rows.append({"Columna": col, "Tipo Archivo A": dl, "Tipo Archivo B": dm, "¿Igual?": dl == dm})
    return pd.DataFrame(rows)


def compare_row_counts(df_legacy: pd.DataFrame, df_migrated: pd.DataFrame) -> dict:
    return {
        "filas_legacy": len(df_legacy),
        "filas_migrated": len(df_migrated),
        "diferencia": len(df_migrated) - len(df_legacy),
    }


def compare_data_sample(
    df_legacy: pd.DataFrame,
    df_migrated: pd.DataFrame,
    common_cols: list,
    key_col: str | None = None,
    max_diffs: int = 50,
) -> pd.DataFrame:
    """
    Compara fila a fila las columnas comunes.
    Si se indica key_col se hace un merge por esa clave, sino se compara por posición.
    """
    diffs = []
    df_l = df_legacy[common_cols].copy()
    df_m = df_migrated[common_cols].copy()

    if key_col and key_col in common_cols:
        merged = df_l.merge(df_m, on=key_col, suffixes=("_legacy", "_migrated"), how="outer")
        value_cols = [c for c in common_cols if c != key_col]
        for _, row in merged.iterrows():
            for col in value_cols:
                v_l = row.get(f"{col}_legacy")
                v_m = row.get(f"{col}_migrated")
                if pd.isna(v_l) and pd.isna(v_m):
                    continue
                if v_l != v_m:
                    diffs.append({
                        "Clave": row[key_col],
                        "Columna": col,
                        "Valor Archivo A": v_l,
                        "Valor Archivo B": v_m,
                    })
                    if len(diffs) >= max_diffs:
                        break
            if len(diffs) >= max_diffs:
                break
    else:
        # Comparación posicional
        min_rows = min(len(df_l), len(df_m))
        for i in range(min_rows):
            for col in common_cols:
                v_l = df_l.iloc[i][col]
                v_m = df_m.iloc[i][col]
                eq = (v_l == v_m) if not (pd.isna(v_l) and pd.isna(v_m)) else True
                if not eq:
                    diffs.append({
                        "Fila": i + 2,  # +2 para reflejar fila Excel (1-indexed + header)
                        "Columna": col,
                        "Valor Archivo A": v_l,
                        "Valor Archivo B": v_m,
                    })
                    if len(diffs) >= max_diffs:
                        break
            if len(diffs) >= max_diffs:
                break

    return pd.DataFrame(diffs) if diffs else pd.DataFrame(columns=["Sin diferencias"])


def render_badge(ok: bool, ok_text="✅ OK", ko_text="⚠️ Diferencias"):
    return ok_text if ok else ko_text


def build_report_txt(
    file_a_name: str,
    file_b_name: str,
    sheet: str,
    rc: dict,
    col_info: dict,
    dtype_df: pd.DataFrame,
    diffs_df: pd.DataFrame,
    nulls_df: pd.DataFrame,
    issues: list[str],
) -> str:
    """Genera el reporte completo en texto plano."""
    SEP = "=" * 72
    SEP2 = "-" * 72
    lines = []

    from datetime import datetime
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    lines += [
        SEP,
        "  REPORTE DE COMPARACIÓN DE EXCEL — ARCHIVO A vs ARCHIVO B",
        f"  Generado: {ts}",
        SEP,
        f"  Archivo A : {file_a_name}",
        f"  Archivo B : {file_b_name}",
        f"  Hoja analizada : {sheet}",
        SEP,
        "",
    ]

    # 1. Estadísticas generales
    lines += [
        "1. ESTADÍSTICAS GENERALES",
        SEP2,
        f"  Filas Archivo A : {rc['filas_legacy']}",
        f"  Filas Archivo B : {rc['filas_migrated']}",
        f"  Diferencia      : {rc['diferencia']:+d}",
        "",
    ]

    # 2. Estructura de columnas
    lines += [
        "2. ESTRUCTURA DE COLUMNAS",
        SEP2,
        f"  Columnas Archivo A : {len(col_info['cols_legacy'])}",
        f"  Columnas Archivo B : {len(col_info['cols_migrated'])}",
        f"  Columnas comunes   : {len(col_info['common'])}",
    ]
    if col_info["only_in_legacy"]:
        lines.append(f"  Solo en Archivo A : {', '.join(str(c) for c in col_info['only_in_legacy'])}")
    if col_info["only_in_migrated"]:
        lines.append(f"  Solo en Archivo B : {', '.join(str(c) for c in col_info['only_in_migrated'])}")
    if not col_info["only_in_legacy"] and not col_info["only_in_migrated"]:
        lines.append("  [OK] Ambos archivos tienen exactamente las mismas columnas.")
    lines.append("")

    # 3. Orden de columnas
    lines += ["3. ORDEN DE COLUMNAS", SEP2]
    if col_info["order_ok"]:
        lines.append("  [OK] El orden de columnas comunes es idéntico.")
    else:
        lines.append(f"  [DIFF] Columnas con posición distinta ({len(col_info['order_changed'])}):")
        for c in col_info["order_changed"]:
            pos_l = col_info["cols_legacy"].index(c) + 1
            pos_m = col_info["cols_migrated"].index(c) + 1
            lines.append(f"    - {c}  (Archivo A: col {pos_l}  |  Archivo B: col {pos_m})")
    lines.append("")

    # Orden completo
    lines.append("  Tabla de orden completo:")
    max_len = max(len(col_info["cols_legacy"]), len(col_info["cols_migrated"]))
    lines.append(f"  {'Pos':>4}  {'Columna Archivo A':<35}  {'Columna Archivo B':<35}  {'¿Igual?'}")
    lines.append(f"  {'-'*4}  {'-'*35}  {'-'*35}  {'-'*7}")
    for i in range(max_len):
        cl = col_info["cols_legacy"][i] if i < len(col_info["cols_legacy"]) else "—"
        cm = col_info["cols_migrated"][i] if i < len(col_info["cols_migrated"]) else "—"
        eq = "OK" if cl == cm else "DIFF"
        lines.append(f"  {i+1:>4}  {str(cl):<35}  {str(cm):<35}  {eq}")
    lines.append("")

    # 4. Tipos de dato
    lines += ["4. TIPOS DE DATO (columnas comunes)", SEP2]
    mismatches = dtype_df[dtype_df["¿Igual?"] == False]  # noqa: E712
    if mismatches.empty:
        lines.append("  [OK] Todos los tipos de dato coinciden.")
    else:
        lines.append(f"  [DIFF] {len(mismatches)} columna(s) con tipo diferente:")
        lines.append(f"  {'Columna':<35}  {'Tipo Archivo A':<20}  {'Tipo Archivo B'}")
        lines.append(f"  {'-'*35}  {'-'*20}  {'-'*20}")
        for _, row in mismatches.iterrows():
            lines.append(f"  {str(row['Columna']):<35}  {str(row['Tipo Archivo A']):<20}  {str(row['Tipo Archivo B'])}")
    lines.append("")
    lines.append("  Detalle completo de tipos:")
    lines.append(f"  {'Columna':<35}  {'Tipo Archivo A':<20}  {'Tipo Archivo B':<20}  Estado")
    lines.append(f"  {'-'*35}  {'-'*20}  {'-'*20}  {'-'*6}")
    for _, row in dtype_df.iterrows():
        estado = "OK" if row["¿Igual?"] else "DIFF"
        lines.append(f"  {str(row['Columna']):<35}  {str(row['Tipo Archivo A']):<20}  {str(row['Tipo Archivo B']):<20}  {estado}")
    lines.append("")

    # 5. Diferencias de datos
    lines += ["5. DIFERENCIAS DE DATOS", SEP2]
    if "Sin diferencias" in diffs_df.columns:
        lines.append("  [OK] No se encontraron diferencias en los datos comparados.")
    else:
        lines.append(f"  [DIFF] {len(diffs_df)} diferencia(s) encontradas:")
        col_names = list(diffs_df.columns)
        header = "  " + "  ".join(f"{str(c):<25}" for c in col_names)
        lines.append(header)
        lines.append("  " + "-" * (26 * len(col_names)))
        for _, row in diffs_df.iterrows():
            lines.append("  " + "  ".join(f"{str(row[c]):<25}" for c in col_names))
    lines.append("")

    # 6. Valores nulos
    lines += ["6. VALORES NULOS POR COLUMNA", SEP2]
    lines.append(f"  {'Columna':<35}  {'Nulos Arch. A':>13}  {'Nulos Arch. B':>13}  {'Diferencia':>10}")
    lines.append(f"  {'-'*35}  {'-'*13}  {'-'*13}  {'-'*10}")
    for col_name, row in nulls_df.iterrows():
        nl = int(row.get("Nulos Archivo A", 0))
        nm = int(row.get("Nulos Archivo B", 0))
        diff_n = int(row.get("Diferencia", 0))
        marker = "  <--" if diff_n != 0 else ""
        lines.append(f"  {str(col_name):<35}  {nl:>12}  {nm:>13}  {diff_n:>+10}{marker}")
    lines.append("")

    # Resumen ejecutivo
    lines += [SEP, "7. RESUMEN EJECUTIVO", SEP]
    if issues:
        for issue in issues:
            # quitar emojis para texto plano
            clean = issue.replace("❌ ", "[ERROR] ").replace("⚠️ ", "[WARN]  ")
            lines.append(f"  {clean}")
    else:
        lines.append("  [OK] Los archivos son estructuralmente idénticos y sin diferencias de datos detectadas.")
    lines += ["", SEP, "  FIN DEL REPORTE", SEP]

    return "\n".join(lines)


# ──────────────────────────────────────────────
# Sidebar — carga de archivos
# ──────────────────────────────────────────────
st.sidebar.header("📁 Archivos de entrada")
file_legacy = st.sidebar.file_uploader("📂 Archivo A", type=["xlsx", "xls"], key="legacy")
file_migrated = st.sidebar.file_uploader("📂 Archivo B", type=["xlsx", "xls"], key="migrated")

st.sidebar.markdown("---")
st.sidebar.subheader("⚙️ Opciones")
max_diff_rows = st.sidebar.slider("Máx. diferencias a mostrar", 10, 500, 50)
show_full_data = st.sidebar.checkbox("Mostrar preview de datos completos", value=False)

both_loaded = file_legacy is not None and file_migrated is not None

if both_loaded:
    st.sidebar.markdown("---")
    procesar = st.sidebar.button("🔍 Procesar comparación", type="primary", use_container_width=True)
else:
    procesar = False
    if file_legacy is None and file_migrated is None:
        pass  # mensaje al pie del sidebar
    elif file_legacy is None:
        st.sidebar.info("Falta cargar el **Archivo A**.")
    else:
        st.sidebar.info("Falta cargar el **Archivo B**.")

# ──────────────────────────────────────────────
# Análisis principal
# ──────────────────────────────────────────────
if both_loaded and procesar:
    with st.spinner("Leyendo archivos..."):
        sheets_legacy = load_excel(file_legacy.getvalue(), "Archivo A")
        sheets_migrated = load_excel(file_migrated.getvalue(), "Archivo B")

    if not sheets_legacy or not sheets_migrated:
        st.stop()

    # Selector de hoja
    common_sheets = [s for s in sheets_legacy if s in sheets_migrated]
    only_legacy_sheets = [s for s in sheets_legacy if s not in sheets_migrated]
    only_migrated_sheets = [s for s in sheets_migrated if s not in sheets_legacy]

    st.subheader("📑 Hojas del Excel")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Hojas en Archivo A", len(sheets_legacy))
    with col2:
        st.metric("Hojas en Archivo B", len(sheets_migrated))
    with col3:
        st.metric("Hojas en común", len(common_sheets))

    if only_legacy_sheets:
        st.warning(f"Solo en Archivo A: **{', '.join(only_legacy_sheets)}**")
    if only_migrated_sheets:
        st.warning(f"Solo en Archivo B: **{', '.join(only_migrated_sheets)}**")

    if not common_sheets:
        st.error("No hay hojas con el mismo nombre. No se puede comparar.")
        st.stop()

    selected_sheet = st.selectbox("Selecciona hoja a analizar", common_sheets)

    df_l = sheets_legacy[selected_sheet]
    df_m = sheets_migrated[selected_sheet]

    st.markdown("---")

    # ── 1. Estadísticas generales ──────────────
    st.subheader("1️⃣ Estadísticas generales")
    rc = compare_row_counts(df_l, df_m)
    c1, c2, c3 = st.columns(3)
    c1.metric("Filas Archivo A", rc["filas_legacy"])
    c2.metric("Filas Archivo B", rc["filas_migrated"])
    delta_color = "off" if rc["diferencia"] == 0 else "inverse"
    c3.metric("Diferencia filas", rc["diferencia"], delta_color=delta_color)

    # ── 2. Estructura de columnas ──────────────
    st.subheader("2️⃣ Estructura de columnas")
    col_info = compare_columns(df_l, df_m)

    c1, c2, c3 = st.columns(3)
    c1.metric("Columnas Archivo A", len(col_info["cols_legacy"]))
    c2.metric("Columnas Archivo B", len(col_info["cols_migrated"]))
    c3.metric("Columnas en común", len(col_info["common"]))

    if col_info["only_in_legacy"]:
        st.error(f"❌ Solo en Archivo A ({len(col_info['only_in_legacy'])}): `{'`, `'.join(col_info['only_in_legacy'])}`")
    if col_info["only_in_migrated"]:
        st.error(f"❌ Solo en Archivo B ({len(col_info['only_in_migrated'])}): `{'`, `'.join(col_info['only_in_migrated'])}`")
    if not col_info["only_in_legacy"] and not col_info["only_in_migrated"]:
        st.success("✅ Ambos archivos tienen exactamente las mismas columnas.")

    # ── 3. Orden de columnas ───────────────────
    st.subheader("3️⃣ Orden de columnas")

    if col_info["order_ok"]:
        st.success("✅ El orden de columnas comunes es idéntico.")
    else:
        st.warning(
            f"⚠️ {len(col_info['order_changed'])} columna(s) tienen posición distinta: "
            f"`{'`, `'.join(col_info['order_changed'])}`"
        )

    with st.expander("Ver orden completo lado a lado"):
        max_len = max(len(col_info["cols_legacy"]), len(col_info["cols_migrated"]))
        order_df = pd.DataFrame({
            "Pos.": range(1, max_len + 1),
            "Columna Archivo A": col_info["cols_legacy"] + ["—"] * (max_len - len(col_info["cols_legacy"])),
            "Columna Archivo B": col_info["cols_migrated"] + ["—"] * (max_len - len(col_info["cols_migrated"])),
        })
        order_df["¿Igual?"] = order_df.apply(
            lambda r: "✅" if r["Columna Archivo A"] == r["Columna Archivo B"] else "⚠️", axis=1
        )
        st.dataframe(order_df, use_container_width=True, hide_index=True)

    # ── 4. Tipos de dato ──────────────────────
    st.subheader("4️⃣ Tipos de dato (columnas comunes)")
    dtype_df = compare_dtypes(df_l, df_m, col_info["common"])
    mismatches = dtype_df[dtype_df["¿Igual?"] == False]  # noqa: E712

    if mismatches.empty:
        st.success("✅ Todos los tipos de dato coinciden.")
    else:
        st.warning(f"⚠️ {len(mismatches)} columna(s) con tipo diferente.")

    with st.expander("Ver detalle de tipos"):
        def highlight_mismatch(row):
            color = "" if row["¿Igual?"] else "background-color: #fff3cd"
            return [color] * len(row)
        st.dataframe(dtype_df.style.apply(highlight_mismatch, axis=1), use_container_width=True, hide_index=True)

    # ── 5. Comparación de datos ────────────────
    st.subheader("5️⃣ Comparación de datos")

    key_col = st.selectbox(
        "Columna clave para el join (opcional — si no aplica, elige 'Sin clave / posicional')",
        ["Sin clave / posicional"] + col_info["common"],
    )
    key_col_val = None if key_col == "Sin clave / posicional" else key_col

    diffs_df = compare_data_sample(df_l, df_m, col_info["common"], key_col=key_col_val, max_diffs=max_diff_rows)

    if "Sin diferencias" in diffs_df.columns:
        st.success("✅ No se encontraron diferencias en los datos comparados.")
    else:
        st.warning(f"⚠️ Se encontraron **{len(diffs_df)}** diferencia(s) (máx. mostradas: {max_diff_rows}).")
        st.dataframe(diffs_df, use_container_width=True, hide_index=True)

    # ── 6. Valores nulos ──────────────────────
    st.subheader("6️⃣ Valores nulos por columna")
    with st.expander("Ver reporte de nulos"):
        nulls_l = df_l[col_info["common"]].isnull().sum().rename("Nulos Archivo A")
        nulls_m = df_m[col_info["common"]].isnull().sum().rename("Nulos Archivo B")
        nulls_df = pd.concat([nulls_l, nulls_m], axis=1)
        nulls_df["Diferencia"] = nulls_df["Nulos Archivo B"] - nulls_df["Nulos Archivo A"]
        st.dataframe(nulls_df.style.highlight_max(axis=1, color="#fff3cd"), use_container_width=True)

    # ── 7. Preview datos ──────────────────────
    if show_full_data:
        st.subheader("7️⃣ Preview de datos")
        c1, c2 = st.columns(2)
        with c1:
            st.caption("Archivo A")
            st.dataframe(df_l.head(100), use_container_width=True)
        with c2:
            st.caption("Archivo B")
            st.dataframe(df_m.head(100), use_container_width=True)

    # ── Resumen ejecutivo ─────────────────────
    st.markdown("---")
    st.subheader("📋 Resumen ejecutivo")

    issues = []
    if col_info["only_in_legacy"]:
        issues.append(f"❌ Columnas solo en Archivo A: {col_info['only_in_legacy']}")
    if col_info["only_in_migrated"]:
        issues.append(f"❌ Columnas solo en Archivo B: {col_info['only_in_migrated']}")
    if not col_info["order_ok"]:
        issues.append(f"⚠️ Columnas con orden diferente: {col_info['order_changed']}")
    if not mismatches.empty:
        issues.append(f"⚠️ Tipos de dato distintos en {len(mismatches)} columna(s)")
    if rc["diferencia"] != 0:
        issues.append(f"⚠️ Diferencia de {abs(rc['diferencia'])} fila(s) entre archivos")
    if "Sin diferencias" not in diffs_df.columns:
        issues.append(f"⚠️ {len(diffs_df)} diferencia(s) de datos encontradas")

    if issues:
        for issue in issues:
            st.markdown(f"- {issue}")
    else:
        st.success("✅ Los archivos son estructuralmente idénticos y sin diferencias de datos detectadas.")

    # ── Descarga TXT ──────────────────────────
    st.markdown("---")
    report_txt = build_report_txt(
        file_a_name=file_legacy.name,
        file_b_name=file_migrated.name,
        sheet=selected_sheet,
        rc=rc,
        col_info=col_info,
        dtype_df=dtype_df,
        diffs_df=diffs_df,
        nulls_df=nulls_df,
        issues=issues,
    )
    from datetime import datetime
    fname = f"reporte_comparacion_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    st.download_button(
        label="⬇️ Descargar reporte TXT",
        data=report_txt.encode("utf-8"),
        file_name=fname,
        mime="text/plain",
        use_container_width=True,
    )

elif both_loaded and not procesar:
    st.success("✅ Archivos cargados correctamente. Presiona **🔍 Procesar comparación** en el panel izquierdo.")
    st.markdown(f"- **Archivo A:** `{file_legacy.name}`")
    st.markdown(f"- **Archivo B:** `{file_migrated.name}`")
else:
    st.info("👈 Carga los dos archivos Excel en el panel izquierdo para comenzar el análisis.")
    st.markdown("""
    ### ¿Qué analiza esta herramienta?

    | # | Análisis | Descripción |
    |---|----------|-------------|
    | 1 | **Estadísticas generales** | Cantidad de filas en cada archivo |
    | 2 | **Estructura de columnas** | Columnas presentes / faltantes |
    | 3 | **Orden de columnas** | Si el orden es el mismo en ambos archivos |
    | 4 | **Tipos de dato** | Si los tipos coinciden por columna |
    | 5 | **Datos fila a fila** | Diferencias de valores (por clave o posición) |
    | 6 | **Valores nulos** | Cantidad de nulos por columna en cada archivo |
    """)
