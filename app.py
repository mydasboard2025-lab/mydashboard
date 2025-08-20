import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Excel Dashboard", page_icon="📊", layout="wide")
st.title("📊 Dashboard – Excel'den Görselleştirme")

# Excel dosya adı (aynı repoda olmalı!)
EXCEL_PATH = "dashboard deneme.xlsx"

@st.cache_data
def load_df(path):
    return pd.read_excel(path)  # openpyxl requirements.txt içinde olmalı

# Veriyi yükle
df = load_df(EXCEL_PATH)
st.success(f"{EXCEL_PATH} yüklendi ✔ • {df.shape[0]} satır, {df.shape[1]} sütun")

# Sayısal ve kategorik sütunları bul
num_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
all_cols = df.columns.tolist()

# Kullanıcıdan seçimler
dim = st.selectbox("Gruplama (Boyut)", all_cols, index=0)
metric = st.selectbox("Metrik (Sayısal)", num_cols or all_cols, index=0)
agg = st.selectbox("Toplama", ["sum", "mean", "count", "min", "max"], index=0)

# Grupla ve görselleştir
if agg == "count":
    g = df.groupby(dim).size().reset_index(name="value")
else:
    g = df.groupby(dim)[metric].agg(agg).reset_index(name="value")

st.plotly_chart(
    px.bar(g, x=dim, y="value", title=f"{dim} bazında {metric} ({agg})"),
    use_container_width=True,
)

# Ham verinin ilk 200 satırını göster
st.subheader("📋 Veri Önizleme")
st.dataframe(df.head(200))
