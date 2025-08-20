import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Excel Dashboard", page_icon="📊", layout="wide")
st.title("📊 Excel/CSV'den Web Dashboard")

EXCEL_PATH = "data.xlsx"   # repodaki dosya adı/yolu (ör. data/data.xlsx)

@st.cache_data
def load_df(path):
    return pd.read_excel(path)  # openpyxl sayesinde .xlsx okunur

df = load_df(EXCEL_PATH)
st.success(f"Yüklendi: {EXCEL_PATH} • {df.shape[0]} satır, {df.shape[1]} sütun")

num_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
dim = st.selectbox("Gruplama (boyut)", df.columns, index=0)
metric = st.selectbox("Metrik (sayısal)", num_cols or df.columns, index=0)
agg = st.selectbox("Toplama", ["sum","mean","count","min","max"], index=0)

if agg == "count":
    g = df.groupby(dim).size().reset_index(name="value")
else:
    g = df.groupby(dim)[metric].agg(agg).reset_index(name="value")

st.plotly_chart(px.bar(g, x=dim, y="value", title=f"{dim} bazında {metric} ({agg})"), use_container_width=True)
st.subheader("Veri (ilk 200 satır)")
st.dataframe(df.head(200))
