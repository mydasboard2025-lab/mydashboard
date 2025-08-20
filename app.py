# app.py
import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="DIO Günlük Bar Chart", page_icon="📊", layout="wide")
st.title("📊 DIO Model Dealer – Model Bazlı Günlük DIO")

EXCEL_PATH = "dashboard deneme.xlsx"      # repodaki Excel dosya adı
SHEET_NAME = "DIO Model Dealer"           # hedef sheet

@st.cache_data
def load_matrix(path: str, sheet: str) -> pd.DataFrame:
    """
    Yapı: E6→ tarihler, D9↓ model isimleri, kesişimde DIO.
    Header satırı: 6. satır → header=5 (0-index)
    Model kolonu: D sütunu → usecols 'D:ZZ' ile D'den sağa alıyoruz.
    """
    # D sütunundan itibaren tüm kolonları al (header=5 → 6. satır başlık kabul)
    df = pd.read_excel(
        path,
        sheet_name=sheet,
        header=5,         # 6. satır başlık
        usecols="D:ZZ",   # D'den sağa doğru
        engine="openpyxl"
    )

    # İlk kolon (D) model isimleri; isimlendirelim
    cols = df.columns.tolist()
    cols[0] = "Model"
    df.columns = cols

    # Model isimlerini temizle
    df["Model"] = df["Model"].astype(str).str.strip()

    # Model olmayan satırları at (boş / NaN / gereksiz başlıklar)
    df = df[df["Model"].notna() & (df["Model"] != "nan") & (df["Model"] != "")]

    # İstenen kalıp: "Yeni BMW" veya "BMW" ile başlayanlar
    df = df[df["Model"].str.startswith(("Yeni BMW", "BMW"), na=False)]

    # Tarih kolonlarını tespit et (başlıklardan tarih ayıkla)
    date_cols = []
    for c in df.columns[1:]:
        try:
            pd.to_datetime(c, dayfirst=True, errors="raise")
            date_cols.append(c)
        except Exception:
            # tarih olmayan başlıkları görmezden gel
            pass

    # Geniş matristen uzun formata çevir
    long_df = df.melt(
        id_vars=["Model"],
        value_vars=date_cols,
        var_name="Date",
        value_name="DIO"
    )

    # Tip dönüşümleri
    long_df["Date"] = pd.to_datetime(long_df["Date"], dayfirst=True, errors="coerce")
    # DIO'yu sayıya çevir (binlik/ondalık olasılıkları için basit temizleme)
    long_df["DIO"] = (
        long_df["DIO"]
        .astype(str)
        .str.replace("%", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(".", "", regex=False)   # binlik noktayı temizle
        .str.replace(",", ".", regex=False)  # ondalık virgülü noktaya çevir
    )
    long_df["DIO"] = pd.to_numeric(long_df["DIO"], errors="coerce")
    long_df = long_df.dropna(subset=["Date"])  # tarih olmayanları at

    return long_df

long_df = load_matrix(EXCEL_PATH, SHEET_NAME)

# --- Filtreler (sidebar) ---
st.sidebar.header("Filtreler")

# Model seçimi
all_models = sorted(long_df["Model"].dropna().unique().tolist())
selected_models = st.sidebar.multiselect(
    "Model seçin",
    options=all_models,
    default=all_models[:3] if len(all_models) > 3 else all_models
)

# Tarih aralığı
min_d, max_d = long_df["Date"].min(), long_df["Date"].max()
date_range = st.sidebar.date_input(
    "Tarih aralığı",
    value=(min_d.date() if pd.notna(min_d) else None,
           max_d.date() if pd.notna(max_d) else None)
)

filtered = long_df.copy()
if selected_models:
    filtered = filtered[filtered["Model"].isin(selected_models)]

if isinstance(date_range, (list, tuple)) and len(date_range) == 2 and all(date_range):
    start, end = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
    filtered = filtered[(filtered["Date"] >= start) & (filtered["Date"] <= end)]

# Günlük bazda birden fazla model seçildiyse: model kırılımı rengine gitsin
# (x=Date, y=DIO, color=Model). Tek model seçildiyse renksiz tek seri.
st.subheader("📈 Günlük DIO – Bar Chart")
if filtered.empty:
    st.warning("Seçilen filtrelerde veri bulunamadı.")
else:
    # Tarihe göre sırala
    filtered = filtered.sort_values("Date")

    if selected_models and len(selected_models) > 1:
        fig = px.bar(
            filtered,
            x="Date", y="DIO",
            color="Model",
            title="Günlük DIO (Model kırılımı)",
            labels={"Date": "Tarih", "DIO": "DIO Adedi"}
        )
    else:
        # Tek model ya da model seçilmediyse toplam göster
        # (model seçilmediyse tamamı tek seri olarak gösterilir)
        # Tek seri bar için toplamı almak istemiyorsan doğrudan filtered kullan.
        fig = px.bar(
            filtered,
            x="Date", y="DIO",
            title="Günlük DIO",
            labels={"Date": "Tarih", "DIO": "DIO Adedi"}
        )

    st.plotly_chart(fig, use_container_width=True)

# Tablo önizleme (performans için limit slider)
st.subheader("📋 Veri Önizleme")
n = st.slider("Kaç satır gösterilsin?", 10, max(10, len(filtered)), min(500, len(filtered)))
st.dataframe(filtered.head(n), use_container_width=True)
