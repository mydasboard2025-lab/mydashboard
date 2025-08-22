import streamlit as st
import pandas as pd
import re
from pathlib import Path

st.set_page_config(page_title="Fiyat Karşılaştırması Dashboard", layout="wide")

DATA_DIR = Path("data")

# ---------- Helpers ----------
def find_latest_excel(data_dir: Path) -> Path | None:
    if not data_dir.exists():
        return None
    files = list(data_dir.glob("*.xlsx"))
    if not files:
        return None
    # 'fiyat' geçenleri öne al, sonra mtime DESC, sonra ada göre
    files.sort(key=lambda p: ("fiyat" not in p.name.lower(), -p.stat().st_mtime, p.name.lower()))
    return files[0]

def to_numeric_locale_aware(s: pd.Series) -> pd.Series:
    """Virgül ondalık, nokta binlik, para/metin temizliği ile güvenli numeric çeviri."""
    t = s.astype(str).str.strip()
    t = t.replace({
        "": pd.NA, "-": pd.NA, "—": pd.NA, "–": pd.NA,
        "N/A": pd.NA, "n/a": pd.NA, "na": pd.NA, "NaN": pd.NA,
        "#N/A": pd.NA, "#NA": pd.NA, "#VALUE!": pd.NA,
    })
    t = t.str.replace(r"[^\d,.\-]", "", regex=True)         # para/etiket sil
    t = t.str.replace(r"\.(?=\d{3}(\D|$))", "", regex=True) # binlik noktası sil
    t = t.str.replace(",", ".", regex=False)                # ondalık virgülü noktaya
    return pd.to_numeric(t, errors="coerce")

def parse_percent_series_mixed(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        ser = pd.to_numeric(s, errors="coerce")
        if ser.notna().sum() and (ser.dropna() > 1).mean() > 0.5:
            ser = ser / 100.0
        return ser
    ser = to_numeric_locale_aware(s)
    if ser.notna().sum() and (ser.dropna() > 1).mean() > 0.5:
        ser = ser / 100.0
    return ser

@st.cache_data(show_spinner=False)
def load_data_auto(path: Path) -> pd.DataFrame:
    # Her zaman ilk sheet + 4. satırdan (skiprows=3) D:Q aralığı
    df = pd.read_excel(
        path,
        sheet_name=0,
        usecols="D:Q",
        skiprows=3,
        header=None,
        engine="openpyxl",
    )
    df.columns = [
        "Marka",           # D
        "Model",           # E
        "Paket",           # F
        "_G",              # G (kullanılmıyor)
        "Stoktaki en uygun otomobil fiyatı",  # H
        "Fiyat konumu",    # I
        "İndirim oranı",   # J
        "_K", "_L", "_M", "_N",              # K..N
        "İndirimli fiyat",                   # O
        "İndirimli fiyat konumu",            # P
        "Spec adjusted fiyat konumu",        # Q
    ]

    # --- 1) Önce grup sınırlarını oluştur (D sütunu boş satır = ayraç) ---
    df["Marka"] = df["Marka"].replace(r"^\s*$", pd.NA, regex=True)
    df["__group_id__"] = df["Marka"].isna().cumsum()   # <-- AYRAÇLAR KORUNDU

    # --- 2) Şimdi H filtresini uygula; grup_id korunur ---
    h_col = "Stoktaki en uygun otomobil fiyatı"
    h_num = to_numeric_locale_aware(df[h_col])

    # Excel #N/A genelde NaN olur; ayrıca numerik NaN'lar da düşsün
    is_na = df[h_col].isna() | h_num.isna()
    # 0 olanlar düşsün
    is_zero = h_num.fillna(0).eq(0)

    df = df[~(is_na | is_zero)].copy()

    # Ayraç (Marka boş) satırları zaten göstermiyoruz
    # Ama grup kimliği korunmuş durumda

    # Yüzde sütunu normalize (0-1)
    df["İndirim oranı"] = parse_percent_series_mixed(df["İndirim oranı"])

    return df

def fmt_numeric(df: pd.DataFrame) -> pd.DataFrame:
    for c in ["Stoktaki en uygun otomobil fiyatı", "İndirimli fiyat"]:
        if c in df.columns:
            df[c] = to_numeric_locale_aware(df[c])
    for c in ["Fiyat konumu", "İndirimli fiyat konumu", "Spec adjusted fiyat konumu"]:
        if c in df.columns:
            conv = pd.to_numeric(df[c], errors="coerce")
            if conv.isna().all():
                conv = to_numeric_locale_aware(df[c])
            df[c] = conv
    return df

# ---------- Flow ----------
EXCEL_PATH = find_latest_excel(DATA_DIR)
if EXCEL_PATH is None or not EXCEL_PATH.exists():
    st.error("`data/` klasöründe .xlsx bulunamadı. Lütfen Excel dosyanı `data/` içine yükle.")
    st.stop()

st.markdown("## BMW Rakip Karşılaştırma")

df_raw = load_data_auto(EXCEL_PATH)

# Filtre kaynakları: sadece BMW
df_bmw = df_raw[(df_raw["Marka"].astype(str).str.strip().str.upper() == "BMW")]
df_bmw = df_bmw[df_bmw["Model"].notna() & df_bmw["Paket"].notna()]

if df_bmw.empty:
    st.warning("Excel içinde (H=0/#N/A filtreleri sonrası) BMW satırı bulunamadı.")
    st.stop()

c1, c2, _ = st.columns([2, 2, 1])
with c1:
    model_list = sorted(df_bmw["Model"].astype(str).unique().tolist())
    selected_model = st.selectbox("BMW Model", options=model_list, index=0)
with c2:
    pkg_list = sorted(df_bmw.loc[df_bmw["Model"].astype(str) == selected_model, "Paket"].astype(str).unique())
    selected_pkg = st.selectbox("Paket", options=pkg_list, index=0)

# Seçilen satır → grup → rakipler
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
