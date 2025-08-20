# app.py (teşhisli)
import os
import glob
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

st.set_page_config(page_title="DIO Günlük Bar", page_icon="📊", layout="wide")
st.title("📊 DIO Model Dealer – Model Bazlı Günlük DIO")

# ---- AYARLAR ----
EXCEL_PATH_DEFAULT = "dashboard deneme.xlsx"     # repodaki dosya adın
SHEET_NAME_DEFAULT = "DIO Model Dealer"          # hedef sheet

# ---- DOSYA KONTROL / YÜKLEME SEÇENEKLERİ ----
st.sidebar.header("Veri Kaynağı")
mode = st.sidebar.radio("Excel nasıl yüklensin?", ["Repo dosyası", "Dosya yükle (Upload)"], index=0)

excel_bytes = None
excel_path = None
if mode == "Repo dosyası":
    excel_path = st.sidebar.text_input("Repo içi yol/isim", EXCEL_PATH_DEFAULT)
    st.info(f"Çalışma klasörü: `{os.getcwd()}`")
    st.caption("Bu klasördeki dosyalar:")
    st.code("\n".join(glob.glob("*")))
    if not os.path.exists(excel_path):
        st.error(f"Excel bulunamadı: `{excel_path}`. Dosya adını/yolunu kontrol et (büyük-küçük harf, boşluk).")
        st.stop()
else:
    up = st.sidebar.file_uploader("Excel yükle (.xlsx)", type=["xlsx"])
    if up is None:
        st.warning("Devam etmek için bir Excel yükleyin.")
        st.stop()
    excel_bytes = up.read()

# ---- SHEET LİSTELE ----
try:
    if excel_bytes:
        xls = pd.ExcelFile(excel_bytes, engine="openpyxl")
    else:
        xls = pd.ExcelFile(excel_path, engine="openpyxl")
    st.success("Excel açıldı ✔")
    st.caption("Bulunan sheet'ler:")
    st.write(xls.sheet_names)
except Exception as e:
    st.exception(e)
    st.stop()

sheet_name = st.sidebar.selectbox("Sheet seç", options=xls.sheet_names,
                                  index=xls.sheet_names.index(SHEET_NAME_DEFAULT) if SHEET_NAME_DEFAULT in xls.sheet_names else 0)

# ---- PARAMETRELER (koordinatlar) ----
st.sidebar.header("Yerleşim Parametreleri")
ROW_DATES = st.sidebar.number_input("Tarih satırı (0-index, E6 → 5)", value=5, step=1)
COL_MODEL = st.sidebar.number_input("Model sütunu (0-index, D → 3)", value=3, step=1)
ROW_MODELS_START = st.sidebar.number_input("Model başlangıç satırı (0-index, D9 → 8)", value=8, step=1)

# ---- HAM OKUMA (header=None) ----
try:
    raw = pd.read_excel(xls, sheet_name=sheet_name, header=None, engine="openpyxl")
    st.success(f"Sheet okundu ✔  Şekil: {raw.shape}")
except Exception as e:
    st.exception(e)
    st.stop()

with st.expander("Ham veri ilk 12×20 önizleme"):
    st.dataframe(raw.iloc[:12, :20])

# ---- TARİHLERİ ÇEK (E6 →) ----
try:
    date_header_row = raw.iloc[int(ROW_DATES), int(COL_MODEL)+1:].astype(str).str.strip()
    date_parsed = pd.to_datetime(date_header_row, dayfirst=True, errors="coerce")
    if date_parsed.isna().mean() > 0.5:
        # dd.mm.yyyy dene
        date_parsed = pd.to_datetime(date_header_row, format="%d.%m.%Y", errors="coerce")
    valid_dates_mask = date_parsed.notna()
    date_cols_idx = np.where(valid_dates_mask.values)[0] + (int(COL_MODEL) + 1)
    dates = date_parsed[valid_dates_mask].reset_index(drop=True)
    st.info(f"Tarih kolon sayısı: {len(dates)} | İlk tarih: {dates.min()} | Son tarih: {dates.max()}")
    if len(dates) == 0:
        st.error("Tarih başlıkları bulunamadı. ROW_DATES / COL_MODEL parametrelerini kontrol edin.")
        st.stop()
except Exception as e:
    st.exception(e)
    st.stop()

# ---- MODELLERİ ÇEK (D9 ↓) ----
try:
    models_series = raw.iloc[int(ROW_MODELS_START):, int(COL_MODEL)].as_]()_
