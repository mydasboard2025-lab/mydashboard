import streamlit as st
import pandas as pd
import re
from pathlib import Path

st.set_page_config(page_title="Fiyat Karşılaştırması Dashboard", layout="wide")

DATA_DIR = Path("data")
ASSETS_DIR = Path("assets")
IMAGE_PATH = ASSETS_DIR / "Fiyat Konumu tablo.png"

# ---------------- Helpers ----------------
def list_excel_files(data_dir: Path) -> list[Path]:
    if not data_dir.exists():
        return []
    files = list(data_dir.glob("*.xlsx"))
    # "Fiyat" geçenleri öne al; sonra mtime DESC; sonra ada göre
    files.sort(key=lambda p: ("fiyat" not in p.name.lower(), -p.stat().st_mtime, p.name.lower()))
    return files

def detect_start_row(df_txt: pd.DataFrame) -> int:
    """
    D,E,F sütunlarında (index 0,1,2) aynı satırda en az birinin dolu olduğu
    ilk satırın 1-index Excel satır numarasına göre başlangıç satırını verir.
    read_excel(skiprows=<return-1>) için kullanırız.
    """
    # df_txt: usecols="D:Q", header=None, dtype=str ile okunmuş olmalı
    for idx in range(len(df_txt)):
        d = str(df_txt.iat[idx, 0]).strip() if idx < len(df_txt) else ""
        e = str(df_txt.iat[idx, 1]).strip() if idx < len(df_txt) else ""
        f = str(df_txt.iat[idx, 2]).strip() if idx < len(df_txt) else ""
        if (d != "" or e != "" or f != "") and (d.lower() != "nan" or e.lower() != "nan" or f.lower() != "nan"):
            # Excel'de bu satır veri başı. pandas.skiprows = idx
            return idx + 1  # sadece kullanıcıya göstermek için Excel satır numarası
    return 4  # varsayılan: 4. satır

def to_numeric_safe(series: pd.Series) -> pd.Series:
    if series.dtype == "object":
        s = series.astype(str).str.replace("%", "", regex=False)
        s = s.str.replace(" ", "", regex=False)
        s = s.str.replace(",", ".", regex=False)
        # binlik noktası temizle (1.234.567 gibi)
        s = s.str.replace(r"\.(?=\d{3}(\D|$))", "", regex=True)
        return pd.to_numeric(s, errors="coerce")
    return pd.to_numeric(series, errors="coerce")

def parse_percent_series_mixed(s: pd.Series) -> pd.Series:
    """
    Excel yüzde ise zaten 0-1 float gelir. Metin ise (10%, 12,5, %7.5) normalize eder.
    Sonuç 0-1 aralığı float.
    """
    if pd.api.types.is_numeric_dtype(s):
        ser = pd.to_numeric(s, errors="coerce")
        # Çoğunluk 1'den büyükse % sayı kabul et (12.5 -> 0.125)
        if ser.notna().sum() and (ser.dropna() > 1).mean() > 0.5:
            ser = ser / 100.0
        return ser

    txt = s.astype(str).str.strip().str.replace("%", "", regex=False).str.replace(" ", "", regex=False)
    txt = txt.str.replace(",", ".", regex=False)
    txt = txt.str.replace(r"\.(?=\d{3}(\D|$))", "", regex=True)  # binlik nokta
    def keep_first_num(x: str) -> str:
        m = re.findall(r"-?\d+(?:\.\d+)?", x)
        return m[0] if m else ""
    cleaned = txt.apply(keep_first_num)
    ser = pd.to_numeric(cleaned, errors="coerce")
    if ser.notna().sum() and (ser.dropna() > 1).mean() > 0.5:
        ser = ser / 100.0
    return ser

@st.cache_data(show_spinner=False)
def read_as_text(path: Path, sheet_name) -> pd.DataFrame:
    # Ham görüntü için metin olarak oku (D:Q)
    return pd.read_excel(path, sheet_name=sheet_name, usecols="D:Q", header=None, dtype=str, engine="openpyxl")

@st.cache_data(show_spinner=False)
def load_data(path: Path, sheet_name, skiprows: int) -> pd.DataFrame:
    df = pd.read_excel(
        path,
        sheet_name=sheet_name,
        usecols="D:Q",
        skiprows=skiprows,   # veri 4. satırdan başlıyorsa 3
        header=None,
        engine="openpyxl",
    )
    df.columns = [
        "Marka",           # D
        "Model",           # E
        "Paket",           # F
        "_G",              # G
        "Stoktaki en uygun otomobil fiyatı",  # H
        "Fiyat konumu",    # I
        "İndirim oranı",   # J (Excel'de Percentage ise 0-1 float)
        "_K", "_L", "_M", "_N",              # K..N
        "İndirimli fiyat",                   # O
        "İndirimli fiyat konumu",            # P
        "Spec adjusted fiyat konumu",        # Q
    ]
    # Boş görünen marka hücrelerini gerçek NaN yap (gruplama sağlıklı olsun)
    df["Marka"] = df["Marka"].replace(r"^\s*$", pd.NA, regex=True)

    # Grup (rakip seti): Marka boşsa yeni grup
    df["__group_id__"] = df["Marka"].isna().cumsum()

    # J sütunu: karma halde gelebilir; normalize et (0-1)
    df["İndirim oranı"] = parse_percent_series_mixed(df["İndirim oranı"])

    return df

def fmt_numeric(df: pd.DataFrame) -> pd.DataFrame:
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

def rerun():
    # Streamlit >=1.27: st.rerun, eski sürümde yoksa no-op
    if hasattr(st, "rerun"):
        st.rerun()

# ---------------- UI: Dosya & Sayfa & Başlangıç Satırı ----------------
top_l, top_m, top_r = st.columns([3, 2, 1])
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
        rerun()

# Sayfa (sheet) seçimi
try:
    sheet_names = pd.ExcelFile(EXCEL_PATH).sheet_names
except Exception as e:
    st.error(f"Excel açılamadı: {e}")
    st.stop()

sheet = st.selectbox("Sheet (sayfa)", options=sheet_names, index=0)

# Ham metin olarak oku ve otomatik başlangıç satırını tespit et
df_txt = read_as_text(EXCEL_PATH, sheet)
auto_start_excel_row = detect_start_row(df_txt)  # Excel satır numarası (1-based)
st.caption(f"Otomatik tespit edilen veri başlangıcı: Excel satırı {auto_start_excel_row}")

# Kullanıcı override
start_excel_row = st.number_input(
    "Veri başlangıç satırı (Excel satır numarası)", min_value=1, value=auto_start_excel_row, step=1
)
# pandas skiprows (0-based): başlangıç satırı n ise skiprows = n-1
skiprows = int(start_excel_row - 1)

# ---------------- Yükleme ----------------
df_raw = load_data(EXCEL_PATH, sheet, skiprows)

# Üst görsel
if IMAGE_PATH.exists():
    st.image(str(IMAGE_PATH), caption="Fiyat Konumu (kurumsal format)")

# Ham tablo (ayraç satırları gizleyip teknik kolonları at)
st.markdown("### Kaynak Excel (doğrudan tablo görünümü)")
st.dataframe(
    df_raw[df_raw["Marka"].notna()].drop(columns=[c for c in df_raw.columns if c.startswith("_") or c.startswith("__")]),
    use_container_width=True,
    hide_index=True,
)

st.markdown("---")
st.markdown("## BMW Rakip Karşılaştırma")

# Filtre kaynakları: sadece BMW
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
    "İndirim oranı",
    "İndirimli fiyat",
    "İndirimli fiyat konumu",
    "Spec adjusted fiyat konumu",
]

df_group_fmt = fmt_numeric(df_group[display_cols].copy())

def highlight_selected(row):
    if (str(row["Marka"]).strip().strupper() if hasattr(str, "strupper") else str(row["Marka"]).strip().upper()) == "BMW" and \
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

# Teşhis için opsiyonel panel
with st.expander("Teşhis Paneli (ham değerler)"):
    st.write("İlk 30 satır (metin olarak okunmuş):")
    st.dataframe(
        df_txt.head(30).rename(columns={0:"Marka",1:"Model",2:"Paket",4:"H_raw",5:"I_raw",6:"J_raw",11:"O_raw",12:"P_raw",13:"Q_raw"})[[0,1,2,4,5,6,11,12,13]],
        use_container_width=True, hide_index=True
    )
    st.caption(f"Kullanılan dosya: `{EXCEL_PATH.name}`, sheet: `{sheet}`, skiprows: {skiprows}")
