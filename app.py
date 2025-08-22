import streamlit as st
import pandas as pd
import re
from pathlib import Path

st.set_page_config(page_title="Fiyat Karşılaştırması Dashboard", layout="wide")

DATA_DIR = Path("data")

# ---------------- Helpers ----------------
def list_excel_files(data_dir: Path) -> list[Path]:
    if not data_dir.exists():
        return []
    files = list(data_dir.glob("*.xlsx"))
    files.sort(key=lambda p: ("fiyat" not in p.name.lower(), -p.stat().st_mtime, p.name.lower()))
    return files

def to_numeric_safe(series: pd.Series) -> pd.Series:
    if series.dtype == "object":
        s = series.astype(str).str.replace("%", "", regex=False).str.replace(" ", "", regex=False)
        s = s.str.replace(",", ".", regex=False)
        s = s.str.replace(r"\.(?=\d{3}(\D|$))", "", regex=True)
        return pd.to_numeric(s, errors="coerce")
    return pd.to_numeric(series, errors="coerce")

def parse_percent_series_mixed(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        ser = pd.to_numeric(s, errors="coerce")
        if ser.notna().sum() and (ser.dropna() > 1).mean() > 0.5:
            ser = ser / 100.0
        return ser
    txt = s.astype(str).str.strip().str.replace("%", "", regex=False).str.replace(" ", "", regex=False)
    txt = txt.str.replace(",", ".", regex=False)
    txt = txt.str.replace(r"\.(?=\d{3}(\D|$))", "", regex=True)
    def keep_first_num(x: str) -> str:
        m = re.findall(r"-?\d+(?:\.\d+)?", x)
        return m[0] if m else ""
    cleaned = txt.apply(keep_first_num)
    ser = pd.to_numeric(cleaned, errors="coerce")
    if ser.notna().sum() and (ser.dropna() > 1).mean() > 0.5:
        ser = ser / 100.0
    return ser

@st.cache_data(show_spinner=False)
def load_data(path: Path) -> pd.DataFrame:
    df = pd.read_excel(
        path,
        sheet_name=0,      # her zaman ilk sheet
        usecols="D:Q",
        skiprows=3,        # 4. satırdan başlat
        header=None,
        engine="openpyxl",
    )
    df.columns = [
        "Marka",           # D
        "Model",           # E
        "Paket",           # F
        "_G",
        "Stoktaki en uygun otomobil fiyatı",  # H
        "Fiyat konumu",    # I
        "İndirim oranı",   # J
        "_K", "_L", "_M", "_N",
        "İndirimli fiyat",                   # O
        "İndirimli fiyat konumu",            # P
        "Spec adjusted fiyat konumu",        # Q
    ]
    df["Marka"] = df["Marka"].replace(r"^\s*$", pd.NA, regex=True)
    df["__group_id__"] = df["Marka"].isna().cumsum()
    df["İndirim oranı"] = parse_percent_series_mixed(df["İndirim oranı"])
    return df

def fmt_numeric(df: pd.DataFrame) -> pd.DataFrame:
    for c in ["Stoktaki en uygun otomobil fiyatı", "İndirimli fiyat"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in ["Fiyat konumu", "İndirimli fiyat konumu", "Spec adjusted fiyat konumu"]:
        if c in df.columns:
            conv = pd.to_numeric(df[c], errors="coerce")
            if conv.isna().all():
                conv = to_numeric_safe(df[c])
            df[c] = conv
    return df

# ---------------- UI ----------------
top_l, top_r = st.columns([4, 1])
with top_l:
    files = list_excel_files(DATA_DIR)
    if not files:
        st.error("`data/` klasöründe .xlsx bulunamadı. Lütfen Excel yükleyin.")
        st.stop()
    labels = [f.name for f in files]
    file_name = st.selectbox("Excel dosyası", options=labels, index=0)
    EXCEL_PATH = next(p for p in files if p.name == file_name)

with top_r:
    if st.button("Yenile"):
        st.cache_data.clear()
        if hasattr(st, "rerun"): st.rerun()

# --------------- Yükle & BMW karşılaştırma ---------------
df_raw = load_data(EXCEL_PATH)

st.markdown("## BMW Rakip Karşılaştırma")

df_bmw = df_raw[(df_raw["Marka"].astype(str).str.strip().str.upper() == "BMW")]
df_bmw = df_bmw[df_bmw["Model"].notna() & df_bmw["Paket"].notna()]

if df_bmw.empty:
    st.warning("Excel içinde BMW satırı bulunamadı.")
    st.stop()

c1, c2, _ = st.columns([2, 2, 1])
with c1:
    model_list = sorted(df_bmw["Model"].astype(str).unique().tolist())
    selected_model = st.selectbox("BMW Model", options=model_list, index=0)
with c2:
    pkg_list = sorted(df_bmw.loc[df_bmw["Model"].astype(str) == selected_model, "Paket"].astype(str).unique())
    selected_pkg = st.selectbox("Paket", options=pkg_list, index=0)

df_sel = df_bmw[(df_bmw["Model"].astype(str) == selected_model) & (df_bmw["Paket"].astype(str) == selected_pkg)]
if df_sel.empty:
    st.info("Seçime uygun satır bulunamadı.")
    st.stop()

group_id = int(df_sel["__group_id__"].iloc[0])
df_group = df_raw[(df_raw["__group_id__"] == group_id) & (df_raw["Marka"].notna())].copy()

display_cols = [
    "Marka",
    "Model",
    "Paket",
    "Stoktaki en uygun otomobil fiyatı",
    "Fiyat konumu",
    "İndirim oranı",
    "İndirimli fiyat",
    "İndirimli fiyat konumu",
    "Spec adjusted fiyat konumu",
]

df_group_fmt = fmt_numeric(df_group[display_cols].copy())

def highlight_selected(row):
    if (str(row["Marka"]).strip().upper() == "BMW") and \
       (str(row["Model"]) == selected_model) and \
       (str(row["Paket"]) == selected_pkg):
        return ["font-weight: bold;"] * len(row)
    return [""] * len(row)

styled = df_group_fmt.style.apply(highlight_selected, axis=1).format(
    {
        "Stoktaki en uygun otomobil fiyatı": "{:,.0f}",
        "İndirimli fiyat": "{:,.0f}",
        "Fiyat konumu": "{:.1f}",
        "İndirimli fiyat konumu": "{:.1f}",
        "Spec adjusted fiyat konumu": "{:.1f}",
        "İndirim oranı": "{:.1%}",
    }
)
st.dataframe(styled, use_container_width=True, hide_index=True)
