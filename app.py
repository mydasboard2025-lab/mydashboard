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
    models_series = raw.iloc[int(ROW_MODELS_START):, int(COL_MODEL)].astype(str).str.strip()
    model_mask = models_series.str.startswith(("Yeni BMW", "BMW"), na=False)
    models = models_series[model_mask].reset_index(drop=True)
    st.info(f"Model satır sayısı: {len(models)} (Yeni BMW/BMW ile başlayanlar)")
    if len(models) == 0:
        st.warning("Model bulunamadı. ROW_MODELS_START veya model metinlerini kontrol et.")
except Exception as e:
    st.exception(e)
    st.stop()

# ---- DIO MATRİSİ ----
try:
    dio_mat = raw.iloc[int(ROW_MODELS_START):, date_cols_idx]
    dio_mat = dio_mat[model_mask.values]
    dio_mat.columns = dates
    df_wide = pd.concat([models.rename("Model"), dio_mat.reset_index(drop=True)], axis=1)
    long_df = df_wide.melt(id_vars=["Model"], var_name="Date", value_name="DIO")
    long_df["Date"] = pd.to_datetime(long_df["Date"], errors="coerce")
    # DIO temizliği
    long_df["DIO"] = (
        long_df["DIO"].astype(str)
        .str.replace(r"[^0-9,.\-]", "", regex=True)
        .str.replace(",", ".", regex=False)
    )
    long_df["DIO"] = pd.to_numeric(long_df["DIO"], errors="coerce")
    long_df = long_df.dropna(subset=["Date"])
    st.success(f"Uzun form oluşturuldu ✔  Satır: {len(long_df)}  Modeller: {long_df['Model'].nunique()}")
except Exception as e:
    st.exception(e)
    st.stop()

with st.expander("Uzun form ilk 15 satır"):
    st.dataframe(long_df.head(15))

# ---- FİLTRELER ----
st.sidebar.header("Filtreler")
all_models = sorted(long_df["Model"].dropna().unique().tolist())
selected_models = st.sidebar.multiselect("Model seçin", all_models,
                                         default=all_models[:3] if len(all_models) > 3 else all_models)

min_d, max_d = long_df["Date"].min(), long_df["Date"].max()
date_range = st.sidebar.date_input("Tarih aralığı",
                                   value=(min_d.date() if pd.notna(min_d) else None,
                                          max_d.date() if pd.notna(max_d) else None))

f = long_df.copy()
if selected_models:
    f = f[f["Model"].isin(selected_models)]
if isinstance(date_range, (list, tuple)) and len(date_range) == 2 and all(date_range):
    start, end = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
    f = f[(f["Date"] >= start) & (f["Date"] <= end)]

st.caption(f"Filtre sonrası satır: {len(f)} | Model sayısı: {f['Model'].nunique()}")

# ---- GRAFİK ----
st.subheader("📈 Günlük DIO – Bar Chart")
if f.empty or f["DIO"].dropna().empty:
    st.warning("Grafik için uygun veri bulunamadı. (Sheet/koordinatlar ya da filtreler veriyi boşaltmış olabilir.)")
else:
    f = f.sort_values("Date")
    color_kw = {"color": "Model"} if selected_models and len(selected_models) > 1 else {}
    fig = px.bar(f, x="Date", y="DIO", **color_kw,
                 labels={"Date": "Tarih", "DIO": "DIO Adedi"},
                 title="Günlük DIO")
    st.plotly_chart(fig, use_container_width=True)

# ---- TABLO ----
st.subheader("📋 Veri Önizleme")
n = st.slider("Kaç satır gösterilsin?", 10, max(10, len(f)), min(500, len(f)))
st.dataframe(f.head(n), use_container_width=True)

