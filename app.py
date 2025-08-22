import streamlit as st
import pandas as pd
import re
from pathlib import Path

st.set_page_config(page_title="Fiyat Karşılaştırması Dashboard", layout="wide")

DATA_DIR = Path("data")
ASSETS_DIR = Path("assets")
IMAGE_PATH = ASSETS_DIR / "Fiyat Konumu tablo.png"

# ---------- Yardımcılar ----------
def list_excel_files(data_dir: Path) -> list[Path]:
    """data/ içindeki tüm .xlsx dosyalarını mtime DESC + isim DESC sıralayıp döndür."""
    if not data_dir.exists():
        return []
    files = list(data_dir.glob("*.xlsx"))
    # "Fiyat" geçenleri üstte tut (öncelik), sonra mtime, sonra isim
    files.sort(key=lambda p: ( "fiyat" not in p.name.lower(), -p.stat().st_mtime, p.name.lower() ))
    return files

@st.cache_data(show_spinner=False)
def load_data(path: Path) -> pd.DataFrame:
    df = pd.read_excel(
        path,
        sheet_name=0,
        usecols="D:Q",
        skiprows=3,   # veri 4. satırdan başlıyor
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
        "İndirim oranı",   # J (Excel'de Percentage => 0.10 gibi float)
        "_K", "_L", "_M", "_N",
        "İndirimli fiyat",                   # O
        "İndirimli fiyat konumu",            # P
        "Spec adjusted fiyat konumu",        # Q
    ]
    # Boş görünen marka hücrelerini gerçek NaN yap (gruplama sağlıklı olsun)
    df["Marka"] = df["Marka"].replace(r"^\s*$", pd.NA, regex=True)
    # Grup (rakip seti): Marka boşsa yeni grup
    df["__group_id__"] = df["Marka"].isna().cumsum()
    return df

def to_numeric_safe(series):
    if series.dtype == "object":
        s = series.astype(str).str.replace(",", ".", regex=False)
        s = s.str.replace(r"\.(?=\d{3}(\D|$))", "", regex=True)  # binlik noktalarını temizle
        return pd.to_numeric(s, errors="coerce")
    return pd.to_numeric(series, errors="coerce")

def fmt_numeric(df):
    # Fiyat kolonları
    for c in ["Stoktaki en uygun otomobil fiyatı", "İndirimli fiyat"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    # Konum kolonları (tek ondalık)
    for c in ["Fiyat konumu", "İndirimli fiyat konumu", "Spec adjusted fiyat konumu"]:
        if c in df.columns:
            conv = pd.to_numeric(df[c], errors="coerce")
            if conv.isna().all():
                conv = to_numeric_safe(df[c])
            df[c] = conv
    return df

# ---------- Üst bar: dosya seçimi + yenile ----------
left, mid, right = st.columns([3, 2, 1])
with left:
    files = list_excel_files(DATA_DIR)
    if not files:
        st.error("`data/` klasöründe .xlsx bulunamadı. Lütfen bir Excel yükleyin.")
        st.stop()
    # Kullanıcıya listeden seçtir (varsayılan ilk eleman: öncelikli/en yeni)
    file_labels = [f.name for f in files]
    default_idx = 0
    selected_label = st.selectbox("Excel dosyası", options=file_labels, index=default_idx)
    EXCEL_PATH = next(p for p in files if p.name == selected_label)

with right:
    if st.button("Yenile"):
        st.cache_data.clear()
        st.experimental_rerun()

st.caption(f"Kullanılan dosya: `{EXCEL_PATH.name}`")

# ---------- Veri yükle ----------
df_raw = load_data(EXCEL_PATH)

# ---------- Üst görsel ----------
if IMAGE_PATH.exists():
    st.image(str(IMAGE_PATH), caption="Fiyat Konumu (kurumsal format)")

st.markdown("### Kaynak Excel (doğrudan tablo görünümü)")
st.dataframe(
    df_raw[df_raw["Marka"].notna()].drop(columns=[c for c in df_raw.columns if c.startswith("_") or c.startswith("__")]),
    use_container_width=True,
    hide_index=True,
)

st.markdown("---")
st.markdown("## BMW Rakip Karşılaştırma")

# Sadece BMW satırları (filtre kaynakları)
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
    "İndirim oranı",                # Excel'de 0.10 gibi gelir
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

st.markdown("### Seçilen model ve rakipleri")
styled = df_group_fmt.style.apply(highlight_selected, axis=1).format(
    {
        "Stoktaki en uygun otomobil fiyatı": "{:,.0f}",
        "İndirimli fiyat": "{:,.0f}",
        "Fiyat konumu": "{:.1f}",
        "İndirimli fiyat konumu": "{:.1f}",
        "Spec adjusted fiyat konumu": "{:.1f}",
        "İndirim oranı": "{:.1%}",  # 0.10 -> %10.0
    }
)
st.dataframe(styled, use_container_width=True, hide_index=True)
