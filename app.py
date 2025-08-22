import streamlit as st
import pandas as pd
import re
from pathlib import Path

st.set_page_config(page_title="Fiyat Karşılaştırması Dashboard", layout="wide")

# -------------------------------
# Dosya bulma: data/ içindeki en yeni .xlsx
# (Öncelik: adında 'Fiyat' geçenler; yoksa en yeni .xlsx)
# -------------------------------
DATA_DIR = Path("data")
ASSETS_DIR = Path("assets")

def find_latest_excel(data_dir: Path) -> Path | None:
    if not data_dir.exists():
        return None
    files = list(data_dir.glob("*.xlsx"))
    if not files:
        return None
    # Adında "Fiyat" geçenleri öne al
    fiyat_files = [f for f in files if "fiyat" in f.name.lower()]
    candidates = fiyat_files if fiyat_files else files
    # En son değişene göre sırala
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0]

EXCEL_PATH = find_latest_excel(DATA_DIR)
IMAGE_PATH = ASSETS_DIR / "Fiyat Konumu tablo.png"

# -------------------------------
# Yardımcı: yüzde serisini 0-1'e çevir (J sütunu)
# -------------------------------
def parse_percent_series(s: pd.Series) -> pd.Series:
    # Direkt sayısal ise akıllı yorumla
    if pd.api.types.is_numeric_dtype(s):
        ser = pd.to_numeric(s, errors="coerce")
        if ser.notna().sum() > 0:
            # Çoğunluk 1'den büyükse muhtemelen % değer; /100 yap
            if (ser.dropna() > 1).mean() > 0.5:
                ser = ser / 100.0
        return ser

    # String/obj ise: %, boşluk ve görünmez karakterleri temizle
    txt = (
        s.astype(str)
         .str.strip()
         .str.replace("%", "", regex=False)
         .str.replace("\u200f", "", regex=False)  # RTL
         .str.replace("\u200e", "", regex=False)  # LTR
    )
    # Ondalık için virgülü noktaya çevir
    txt = txt.str.replace(",", ".", regex=False)

    # Binlik noktalarını temizle: 1.234.567,89 gibi kalıplarda
    # Yalnızca binlik olan nokta: ardından 3 rakam geliyorsa sil
    txt = txt.str.replace(r"\.(?=\d{3}(\D|$))", "", regex=True)

    # Sadece sayı ve tek bir ondalık nokta bırak
    # (Örn. "1.234.5" -> "1234.5")
    def keep_numeric(x: str) -> str:
        m = re.findall(r"-?\d+(?:\.\d+)?", x)
        return m[0] if m else ""

    cleaned = txt.apply(keep_numeric)
    ser = pd.to_numeric(cleaned, errors="coerce")

    if ser.notna().sum() == 0:
        return ser

    # Çoğunluk 1'den büyükse /100
    if (ser.dropna() > 1).mean() > 0.5:
        ser = ser / 100.0
    return ser

# -------------------------------
# Veri yükleme
# -------------------------------
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
        "_G",              # G (kullanılmıyor)
        "Stoktaki en uygun otomobil fiyatı",  # H
        "Fiyat konumu",    # I
        "İndirim oranı",   # J  (yüzde)
        "_K", "_L", "_M", "_N",              # K..N
        "İndirimli fiyat",                   # O
        "İndirimli fiyat konumu",            # P
        "Spec adjusted fiyat konumu",        # Q
    ]
    # Blok/grup: D sütunundaki boş satır her yeni grubu başlatır
    df["__group_id__"] = df["Marka"].isna().cumsum()

    # J sütunu: yüzdeleri 0-1'e çevir
    df["İndirim oranı"] = parse_percent_series(df["İndirim oranı"])

    return df

# -------------------------------
# Başla
# -------------------------------
if EXCEL_PATH is None or not EXCEL_PATH.exists():
    st.error("`data/` klasöründe .xlsx bulunamadı. Lütfen Excel dosyanı `data/` içine yükle.")
    st.stop()

df_raw = load_data(EXCEL_PATH)

# Görseli göster
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

# Filtre listeleri: yalnızca BMW satırları
df_bmw = df_raw[(df_raw["Marka"].astype(str).str.strip().str.upper() == "BMW")].copy()
df_bmw = df_bmw[df_bmw["Model"].notna() & df_bmw["Paket"].notna()]

if df_bmw.empty:
    st.warning("Excel içinde BMW satırı bulunamadı.")
    st.stop()

col1, col2, _ = st.columns([2, 2, 1])
with col1:
    model_list = sorted(df_bmw["Model"].astype(str).unique().tolist())
    selected_model = st.selectbox("BMW Model", options=model_list, index=0)
with col2:
    pkg_list = sorted(
        df_bmw.loc[df_bmw["Model"].astype(str) == selected_model, "Paket"].astype(str).unique().tolist()
    )
    selected_pkg = st.selectbox("Paket", options=pkg_list, index=0)

# Seçilen satır
df_selected_bmw = df_bmw[
    (df_bmw["Model"].astype(str) == selected_model) &
    (df_bmw["Paket"].astype(str) == selected_pkg)
]
if df_selected_bmw.empty:
    st.info("Seçime uygun satır bulunamadı.")
    st.stop()

# Aynı grup (rakipleriyle birlikte)
group_id = int(df_selected_bmw["__group_id__"].iloc[0])
df_group = df_raw[(df_raw["__group_id__"] == group_id) & (df_raw["Marka"].notna())].copy()

# Gösterilecek kolonlar
display_cols = [
    "Marka",
    "Model",
    "Paket",
    "Stoktaki en uygun otomobil fiyatı",
    "Fiyat konumu",
    "İndirim oranı",                # yüzde (0-1)
    "İndirimli fiyat",
    "İndirimli fiyat konumu",
    "Spec adjusted fiyat konumu",
]

# Vurgulama (seçilen BMW kalın)
def highlight_selected(row):
    if (str(row["Marka"]).strip().upper() == "BMW") and \
       (str(row["Model"]) == selected_model) and \
       (str(row["Paket"]) == selected_pkg):
        return ["font-weight: bold;"] * len(row)
    return [""] * len(row)

# Sayısal biçim: fiyatlar ve konumlar
def to_numeric_safe(series):
    if series.dtype == "object":
        # , -> . ve binlik noktalarını temizle
        s = series.astype(str).str.replace(",", ".", regex=False)
        s = s.str.replace(r"\.(?=\d{3}(\D|$))", "", regex=True)
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

df_group_fmt = fmt_numeric(df_group[display_cols].copy())

st.markdown("### Seçilen model ve rakipleri")
styled = df_group_fmt.style.apply(highlight_selected, axis=1).format(
    {
        "Stoktaki en uygun otomobil fiyatı": "{:,.0f}",
        "İndirimli fiyat": "{:,.0f}",
        "Fiyat konumu": "{:.1f}",
        "İndirimli fiyat konumu": "{:.1f}",
        "Spec adjusted fiyat konumu": "{:.1f}",
        "İndirim oranı": "{:.1%}",   # tek ondalık yüzde
    }
)
st.dataframe(styled, use_container_width=True, hide_index=True)

with st.expander("Açıklama / Notlar"):
    st.markdown(
        "- Veriler 4. satırdan itibaren okunur; D sütunundaki **boş satır** yeni gruptur.\n"
        "- Filtreler yalnızca D sütununda **BMW** olan satırlardan türetilir (Model=E, Paket=F).\n"
        "- Fiyatlar binlik ayraçlı; *konum* sütunları tek ondalık; **İndirim oranı** yüzde (tek ondalık) gösterilir.\n"
        f"- Okunan Excel: `{EXCEL_PATH.name}` (data/ içindeki en yeni .xlsx)."
    )
