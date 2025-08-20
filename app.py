# app.py
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

st.set_page_config(page_title="DIO Günlük Bar", page_icon="📊", layout="wide")
st.title("📊 DIO Model Dealer – Model Bazlı Günlük DIO")

EXCEL_PATH = "dashboard deneme.xlsx"
SHEET_NAME = "DIO Model Dealer"

# Konum sabitlemeleri (sorunda verdiğin düzene göre)
ROW_DATES = 5        # E6 sağa → başlık satırı (0-index: 5 => 6. satır)
COL_MODEL = 3        # D sütunu (0-index: 3)
ROW_MODELS_START = 8 # D9 ve aşağısı (0-index: 8 => 9. satır)

@st.cache_data
def load_long_df(path, sheet):
    # 1) Ham oku (header=None) -> ParserError riskini azaltır
    raw = pd.read_excel(path, sheet_name=sheet, header=None, engine="openpyxl")

    # 2) Tarih kolonlarını tespit et (E sütunundan sağa)
    date_cells = raw.iloc[ROW_DATES, COL_MODEL+1:]        # E6→
    date_parsed = pd.to_datetime(date_cells, dayfirst=True, errors="coerce")
    valid_mask = date_parsed.notna()
    date_cols_idx = np.where(valid_mask.values)[0] + (COL_MODEL + 1)  # gerçek kolon indeksleri
    dates = date_parsed[valid_mask].reset_index(drop=True)

    if len(date_cols_idx) == 0:
        raise ValueError("Tarih başlıkları (E6 ve sağa) okunamadı. Lütfen sheet yapısını kontrol edin.")

    # 3) Model sütununu (D9↓) al
    models = raw.iloc[ROW_MODELS_START:, COL_MODEL].astype(str).str.strip()
    # "Yeni BMW" veya "BMW" ile başlayanları filtrele
    model_mask = models.str.startswith(("Yeni BMW", "BMW"), na=False)
    models = models[model_mask].reset_index(drop=True)

    # 4) DIO değer matrisini al (modellerin yanındaki tarih kolonları)
    dio_mat = raw.iloc[ROW_MODELS_START:, date_cols_idx]           # satırlar: modeller
    dio_mat = dio_mat[model_mask.values]                           # aynı maskeyi uygula
    dio_mat.columns = dates                                        # kolon adları datetime

    # 5) Uzun formata dönüştür
    df = pd.concat([models.rename("Model"), dio_mat.reset_index(drop=True)], axis=1)
    long_df = df.melt(id_vars=["Model"], var_name="Date", value_name="DIO")

    # 6) Tip temizliği
    long_df["Date"] = pd.to_datetime(long_df["Date"], dayfirst=True, errors="coerce")
    long_df["DIO"] = (
        long_df["DIO"].astype(str)
        .str.replace("%", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(".", "", regex=False)   # binlik noktaları
        .str.replace(",", ".", regex=False)  # ondalık virgül
    )
    long_df["DIO"] = pd.to_numeric(long_df["DIO"], errors="coerce")

    # Boş tarihleri ve tamamen NaN DIO'ları ele
    long_df = long_df.dropna(subset=["Date"])
    return long_df

try:
    long_df = load_long_df(EXCEL_PATH, SHEET_NAME)
except Exception as e:
    st.error(f"Veri okunurken sorun oluştu: {e}")
