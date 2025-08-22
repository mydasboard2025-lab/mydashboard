import streamlit as st
import pandas as pd
from pathlib import Path

st.set_page_config(page_title="Fiyat Karşılaştırması Dashboard", layout="wide")

# ---- Dosya Yolları (repo içi) ----
EXCEL_PATH = Path("data") / "Fiyat Karşılaştırması - 08.08.2025.xlsx"
IMAGE_PATH = Path("assets") / "Fiyat Konumu tablo.png"

# ---- Yardımcı: Veri Yükleme ----
@st.cache_data(show_spinner=False)
def load_data(path: Path) -> pd.DataFrame:
    df = pd.read_excel(
        path,
        sheet_name=0,
        usecols="D:Q",
        skiprows=3,     # veri 4. satırdan başlıyor
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
    # Blok/grup: D (Marka) boş satırları blok ayırıcısı
    df["__group_id__"] = df["Marka"].isna().cumsum()
    return df

# ---- Yükle ----
if not EXCEL_PATH.exists():
    st.error(f"Excel bulunamadı: {EXCEL_PATH.resolve()}")
    st.stop()

df_raw = load_data(EXCEL_PATH)

# Üstte kurumsal “Fiyat Konumu” görseli (direkt)
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

# Sadece BMW satırları
df_bmw = df_raw[(df_raw["Marka"].astype(str).str.strip().str.upper() == "BMW")].copy()
df_bmw = df_bmw[df_bmw["Model"].notna() & df_bmw["Paket"].notna()]

if df_bmw.empty:
    st.warning("Excel içinde BMW satırı bulunamadı.")
    st.stop()

# Filtreler
col1, col2, _ = st.columns([2, 2, 1])

with col1:
    model_list = sorted(df_bmw["Model"].astype(str).unique().tolist())
    selected_model = st.selectbox("BMW Model", options=model_list, index=0, key="model_select")

with col2:
    pkg_list = sorted(
        df_bmw.loc[df_bmw["Model"].astype(str) == selected_model, "Paket"]
        .astype(str).unique().tolist()
    )
    selected_pkg = st.selectbox("Paket", options=pkg_list, index=0, key="pkg_select")

# Seçilen BMW satırı
df_selected_bmw = df_bmw[
    (df_bmw["Model"].astype(str) == selected_model) &
    (df_bmw["Paket"].astype(str) == selected_pkg)
]

if df_selected_bmw.empty:
    st.info("Seçime uygun satır bulunamadı.")
    st.stop()

# Grup kimliği ve rakipleri getir
group_id = int(df_selected_bmw["__group_id__"].iloc[0])
df_group = df_raw[(df_raw["__group_id__"] == group_id) & (df_raw["Marka"].notna())].copy()

# Gösterilecek kolonlar (İndirim oranı eklendi)
display_cols = [
    "Marka",
    "Model",
    "Paket",
    "Stoktaki en uygun otomobil fiyatı",
    "Fiyat konumu",
    "İndirim oranı",                # <--- eklendi
    "İndirimli fiyat",
    "İndirimli fiyat konumu",
    "Spec adjusted fiyat konumu",
]

# Vurgulama
def highlight_selected(row):
    if (str(row["Marka"]).strip().upper() == "BMW") and \
       (str(row["Model"]) == selected_model) and \
       (str(row["Paket"]) == selected_pkg):
        return ["font-weight: bold;"] * len(row)
    return [""] * len(row)

# Numerik format: fiyatlar ve konumlar
def to_numeric_safe(s):
    # Nokta/virgül farklarını güvenli çevirmek için
    if s.dtype == "object":
        return pd.to_numeric(s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce")
    return pd.to_numeric(s, errors="coerce")

def fmt_numeric(df):
    # Fiyat kolonlarını sayı yap
    price_cols = ["Stoktaki en uygun otomobil fiyatı", "İndirimli fiyat"]
    for c in price_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # Konum kolonlarını tek ondalık için sayı yap
    pos_cols = ["Fiyat konumu", "İndirimli fiyat konumu", "Spec adjusted fiyat konumu"]
    for c in pos_cols:
        if c in df.columns:
            # Önce doğrudan, olmazsa locale varyasyonları için to_numeric_safe
            converted = pd.to_numeric(df[c], errors="coerce")
            if converted.isna().all():
                converted = to_numeric_safe(df[c].astype(str))
            df[c] = converted
    return df

df_group_fmt = fmt_numeric(df_group[display_cols].copy())

st.markdown("### Seçilen model ve rakipleri")
styled = df_group_fmt.style.apply(highlight_selected, axis=1).format(
    {
        "Stoktaki en uygun otomobil fiyatı": "{:,.0f}",
        "İndirimli fiyat": "{:,.0f}",
        "Fiyat konumu": "{:.1f}",
        "İndirimli fiyat konumu": "{:.1f}",
        "Spec adjusted fiyat konumu": "{:.1f}",
        # "İndirim oranı": "{:.1f}%"  # İstersen bunu aktifleştiririz (değerler 0-100 ise güzel durur)
    }
)
st.dataframe(styled, use_container_width=True, hide_index=True)

with st.expander("Açıklama / Notlar"):
    st.markdown(
        "- Veri blokları D sütunundaki boş satırlar ile ayrılmıştır.\n"
        "- Filtreler yalnızca D sütununda **BMW** olan satırlardan türetilmiştir (Model=E, Paket=F).\n"
        "- Fiyatlar binlik ayraçla, *konum* sütunları tek ondalık basamakla gösterilir.\n"
        "- **İndirim oranı** hücreleri içeriği neyse aynen gösterilir (%, nokta/virgül vb.)."
    )
