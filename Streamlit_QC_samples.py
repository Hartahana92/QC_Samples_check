import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="LabQC Normalization & Z‑Score", layout="wide")

st.title("🧪 LabQC Normalization & Z‑Score App")
st.write(
    "Загрузите Excel-файл с колонками **Group**, **Sample** и метаболитами (начиная с 3‑й колонки).\n"
    "Группы: `Sample` — образцы, `labQC` — лабораторные QC."
)

# -------------------------------
# Helpers
# -------------------------------
@st.cache_data(show_spinner=False)
def load_excel(file, sheet_name=None):
    """Load Excel as DataFrame. If sheet_name is None, return dict of sheets."""
    if sheet_name is None:
        # Return all sheets as dict
        xl = pd.read_excel(file, sheet_name=None)
        return xl
    else:
        return pd.read_excel(file, sheet_name=sheet_name)


def compute_norm_and_z(df: pd.DataFrame):
    required_cols = {"Group", "Sample"}
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Отсутствуют обязательные колонки: {', '.join(missing)}")

    # metabolites = all columns after the 2nd (index 2:)
    if len(df.columns) < 3:
        raise ValueError("В файле должно быть минимум 3 колонки: Group, Sample и хотя бы один метаболит.")

    metabolites = df.columns[2:]

    # Split groups
    samples = df[df["Group"] == "Sample"].copy()
    labqcs = df[df["Group"] == "labQC"].copy()

    if samples.empty:
        raise ValueError("Не найдены строки с Group == 'Sample'.")
    if labqcs.empty:
        raise ValueError("Не найдены строки с Group == 'labQC'.")

    # Compute LabQC medians per metabolite
    labqc_ref = labqcs[metabolites].median(axis=0)

    # Normalize sample values by LabQC medians
    # Align division (avoid SettingWithCopy on loop, vectorize instead)
    norm = samples.copy()
    norm_df = pd.DataFrame({"Name": norm["Sample"].values})

    # Avoid division by zero: replace 0 medians with NaN to prevent inf
    safe_ref = labqc_ref.replace(0, np.nan)
    norm_values = norm[metabolites].div(safe_ref, axis=1)
    norm_df = pd.concat([norm_df, norm_values.reset_index(drop=True)], axis=1)

    # Summ and Z
    norm_df["Summ"] = norm_df.iloc[:, 1:].sum(axis=1, numeric_only=True)
    summ = norm_df["Summ"]
    norm_df["Z"] = (summ - summ.mean()) / summ.std(ddof=0)

    # Build LabQC reference table for display
    labqc_table = pd.DataFrame({"Metabolite": metabolites, "LabQC_median": labqc_ref.values})

    return norm_df, labqc_table, metabolites


def to_excel_bytes(df_dict: dict):
    """Pack multiple DataFrames into one XLSX in-memory."""
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        for name, d in df_dict.items():
            d.to_excel(writer, index=False, sheet_name=name[:31] or "Sheet1")
    bio.seek(0)
    return bio


# -------------------------------
# Sidebar — load data
# -------------------------------
st.sidebar.header("1) Загрузка данных")
file = st.sidebar.file_uploader("Загрузите Excel (.xlsx)", type=["xlsx", "xlsm", "xls"])

if file is not None:
    # Read sheet names
    try:
        all_sheets = load_excel(file, sheet_name=None)
        sheet_names = list(all_sheets.keys())
        sheet = st.sidebar.selectbox("Лист Excel", sheet_names, index=0)
        data_raw = all_sheets[sheet]
    except Exception as e:
        st.error(f"Ошибка при чтении файла: {e}")
        st.stop()
else:
    st.info("Загрузите файл слева, чтобы продолжить.")
    st.stop()

# -------------------------------
# Preview
# -------------------------------
st.header("Предпросмотр исходных данных")
st.dataframe(data_raw.head(50), use_container_width=True)

with st.expander("Проверка колонок / структуры"):
    st.write({"columns": list(data_raw.columns), "rows": len(data_raw)})

# -------------------------------
# Compute
# -------------------------------
try:
    df_norm, labqc_table, metabolites = compute_norm_and_z(data_raw)
except Exception as e:
    st.error(f"Ошибка вычислений: {e}")
    st.stop()

st.header("Результаты нормализации и Z‑оценки")
st.markdown("**Таблица LabQC медиан по метаболитам**")
st.dataframe(labqc_table, use_container_width=True)

st.markdown("**Нормализованные образцы** (\"Sample\")")
st.dataframe(df_norm, use_container_width=True)

# -------------------------------
# Simple analytics / visuals
# -------------------------------
st.subheader("Аналитика Z‑оценки")

# Top / bottom samples by Z
k = st.slider("Сколько топ/анти‑топ образцов показать?", 3, min(20, len(df_norm)), 5)

topk = df_norm.nlargest(k, "Z")["Name"].tolist()
botk = df_norm.nsmallest(k, "Z")["Name"].tolist()

col1, col2 = st.columns(2)
with col1:
    st.write("**Top Z**")
    st.table(df_norm.nlargest(k, "Z")[["Name", "Summ", "Z"]].reset_index(drop=True))
with col2:
    st.write("**Bottom Z**")
    st.table(df_norm.nsmallest(k, "Z")[["Name", "Summ", "Z"]].reset_index(drop=True))

# Histogram of Z
st.write("**Гистограмма Z**")
st.bar_chart(df_norm.set_index("Name")["Z"], use_container_width=True)

# -------------------------------
# Downloads
# -------------------------------
st.header("Экспорт")

excel_bytes = to_excel_bytes({
    "LabQC_medians": labqc_table,
    "Normalized": df_norm,
})

st.download_button(
    label="⬇️ Скачать результаты (XLSX)",
    data=excel_bytes,
    file_name="labqc_norm_z_results.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption(
    "Логика: для каждого метаболита берётся медиана по строкам `labQC`, затем значения `Sample` делятся на эту медиану;\n"
    "далее суммируются нормализованные метаболиты по строке и считается Z‑оценка по сумме."
)
