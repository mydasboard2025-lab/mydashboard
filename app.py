import streamlit as st
import pandas as pd
import numpy as np
import re
from pathlib import Path
from datetime import datetime
import zoneinfo
import pytz
import glob
import os

# ================== Genel Ayarlar ==================
st.set_page_config(page_title="Fiyat Karşılaştırması Dashboard", layout="wide")
DATA_DIR = Path("data")

# Dosya isimleri
PRICE_FILE_NAME = "Fiyat Karşılaştırması_v4.xlsx"      # Rakip karşılaştırma
PERF_FILE_NAME  = "Model aylık performans.xlsx"        # Retail/Handover/Presold/Free + DIO Model + Monthly Basis

# Ortak sabitler
IST_TZ = pytz.timezone("Europe/Istanbul")
MONTHS_EN = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

# ================== Yardımcı Fonksiyonlar (genel) ==================
def to_numeric_locale_aware(s: pd.Series) -> pd.Series:
    t = s.astype(str).str.strip()
    t = t.replace({
        "": pd.NA, "-": pd.NA, "—": pd.NA, "–": pd.NA,
        "N/A": pd.NA, "n/a": pd.NA, "na": pd.NA, "NaN": pd.NA,
        "#N/A": pd.NA, "#NA": pd.NA, "#VALUE!": pd.NA,
    })
    t = t.str.replace(r"[^\d,.\-]", "", regex=True)         # para vb temizle
    t = t.str.replace(r"\.(?=\d{3}(\D|$))", "", regex=True) # binlik noktası sil
    t = t.str.replace(",", ".", regex=False)                # ondalık virgülü noktaya çevir
    return pd.to_numeric(t, errors="coerce")

def parse_percent_series_mixed(s: pd.Series) -> pd.Series:
    # karışık % alanlarını 0-1'e çek
    if pd.api.types.is_numeric_dtype(s):
        ser = pd.to_numeric(s, errors="coerce")
        if ser.notna().sum() and (ser.dropna() > 1).mean() > 0.5:
            ser = ser / 100.0
        return ser
    ser = to_numeric_locale_aware(s)
    if ser.notna().sum() and (ser.dropna() > 1).mean() > 0.5:
        ser = ser / 100.0
    return ser

def fmt_numeric(df: pd.DataFrame) -> pd.DataFrame:
    for c in ["Stoktaki en uygun otomobil fiyatı", "İndirimli fiyat"]:
        if c in df.columns:
            df[c] = to_numeric_locale_aware(df[c])
    for c in ["Fiyat konumu", "İndirimli fiyat konumu", "Spec adjusted fiyat konumu"]:
        if c in df.columns:
            conv = pd.to_numeric(df[c], errors="coerce")
            if conv.isna().all():
                conv = to_numeric_locale_aware(df[c])
            df[c] = conv
    return df

def fmt_int(x):
    if x is None or pd.isna(x):
        return "—"
    return f"{float(x):,.0f}".replace(",", ".")

# ================== RAKİP KARŞILAŞTIRMA ==================
def find_price_excel(data_dir: Path) -> Path | None:
    if not data_dir.exists():
        return None
    exact = data_dir / PRICE_FILE_NAME
    if exact.exists():
        return exact

    files = list(data_dir.glob("*.xlsx"))
    if not files:
        return None

    # Önce adı 'fiyat' geçenler + en yeni dosya
    files.sort(key=lambda p: ("fiyat" not in p.name.lower(), -p.stat().st_mtime, p.name.lower()))
    return files[0]

@st.cache_data(show_spinner=False)
def load_price_compare(path: Path) -> pd.DataFrame:
    df = pd.read_excel(
        path,
        sheet_name=0,
        usecols="D:Q",
        skiprows=3,
        header=None,
        engine="openpyxl",
    )
    df.columns = [
        "Marka", "Model", "Paket", "_G",
        "Stoktaki en uygun otomobil fiyatı",  # H
        "Fiyat konumu",                       # I
        "İndirim oranı",                      # J
        "_K", "_L", "_M", "_N",
        "İndirimli fiyat",                    # O
        "İndirimli fiyat konumu",             # P
        "Spec adjusted fiyat konumu",         # Q
    ]

    # Grup kimliği: Marka boşsa grup ayırıcı
    df["Marka"] = df["Marka"].replace(r"^\s*$", pd.NA, regex=True)
    df["__group_id__"] = df["Marka"].isna().cumsum()

    # H kolonu 0/Na olanları at
    h_col = "Stoktaki en uygun otomobil fiyatı"
    h_num = to_numeric_locale_aware(df[h_col])
    is_na = df[h_col].isna() | h_num.isna()
    is_zero = h_num.fillna(0).eq(0)
    df = df[~(is_na | is_zero)].copy()

    df["İndirim oranı"] = parse_percent_series_mixed(df["İndirim oranı"])
    return df

def build_price_compare_ui(df_raw: pd.DataFrame, source_path: Path):
    st.markdown("## BMW Rakip Karşılaştırma")

    # BMW satırlarını çek
    df_bmw = df_raw[(df_raw["Marka"].astype(str).str.strip().str.upper() == "BMW")]
    df_bmw = df_bmw[df_bmw["Model"].notna() & df_bmw["Paket"].notna()]

    if df_bmw.empty:
        st.warning("Excel içinde (H=0/#N/A filtreleri sonrası) BMW satırı bulunamadı.")
        return

    c1, c2, _ = st.columns([2, 2, 1])
    with c1:
        model_list = sorted(df_bmw["Model"].astype(str).unique().tolist())
        selected_model = st.selectbox("BMW Model", options=model_list, index=0, key="bmw_model")
    with c2:
        pkg_list = sorted(
            df_bmw.loc[df_bmw["Model"].astype(str) == selected_model, "Paket"]
            .astype(str)
            .unique()
        )
        if len(pkg_list) == 0:
            st.info("Seçilen model için paket bulunamadı.")
            return
        selected_pkg = st.selectbox("Paket", options=pkg_list, index=0, key="bmw_pkg")

    df_sel = df_bmw[
        (df_bmw["Model"].astype(str) == selected_model)
        & (df_bmw["Paket"].astype(str) == selected_pkg)
    ]
    if df_sel.empty:
        st.info("Seçime uygun satır bulunamadı.")
        return

    # Aynı gruptaki rakipleri listele
    group_id = int(df_sel["__group_id__"].iloc[0])
    df_group = df_raw[(df_raw["__group_id__"] == group_id) & (df_raw["Marka"].notna())].copy()

    display_cols = [
        "Marka", "Model", "Paket",
        "Stoktaki en uygun otomobil fiyatı",
        "Fiyat konumu",
        "İndirim oranı",
        "İndirimli fiyat",
        "İndirimli fiyat konumu",
        "Spec adjusted fiyat konumu",
    ]
    df_group_fmt = fmt_numeric(df_group[display_cols].copy())

    def highlight_selected(row):
        if (
            (str(row["Marka"]).strip().upper() == "BMW")
            and (str(row["Model"]) == selected_model)
            and (str(row["Paket"]) == selected_pkg)
        ):
            return ["font-weight: bold;"] * len(row)
        return [""] * len(row)

    styled = df_group_fmt.style.apply(highlight_selected, axis=1).format(
        {
            "Stoktaki en uygun otomobil fiyatı": "{:,.0f}",
            "İndirimli fiyat": "{:,.0f}",
            "Fiyat konumu": "{:.1f}",
            "İndirimli fiyat konumu": "{:.1f}",
            "Spec adjusted fiyat konumu": "{:.1f}",
            "İndirim oranı": "{:.1%}",
        }
    )
    st.dataframe(styled, use_container_width=True, hide_index=True)
    st.caption(f"Kaynak: {source_path.name}")


# ================== MODEL AYLIK PERFORMANS / KPI / SİRKÜLER / GÜNLÜK DIO ==================

REQUIRED_SHEETS = {
    "Retail",
    "Handover Model",
    "Presold",
    "DIO Model",
    "Monthly Basis",
    "Series",
    "BMW Showroom Seri Bazlı",
    "Sirküler",
}

def find_performance_workbook(data_dir: Path) -> Path | None:
    if not data_dir.exists():
        return None

    exact = data_dir / PERF_FILE_NAME
    if exact.exists():
        return exact

    files = list(data_dir.glob("*.xlsx"))
    if not files:
        return None

    def _priority(p: Path):
        name = p.name.lower()
        prio_name = 0 if ("model" in name or "performans" in name or "performance" in name) else 1
        return (prio_name, -p.stat().st_mtime, name)

    files.sort(key=_priority)

    for f in files:
        try:
            xls = pd.ExcelFile(f, engine="openpyxl")
            sheets = set(xls.sheet_names)
            if REQUIRED_SHEETS.issubset(sheets):
                return f
        except Exception:
            continue

    # fallback
    return files[0] if files else None

def _read_sheet(path: Path, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name=sheet_name, header=None, engine="openpyxl")

def _header_row(df: pd.DataFrame, header_idx: int = 5) -> pd.Series:
    return df.iloc[header_idx].copy()

def _month_col_index_by_abbr(df: pd.DataFrame, month_abbr: str, start_col: int = 5, header_idx: int = 5) -> int | None:
    hdr = _header_row(df, header_idx=header_idx)
    target = month_abbr.strip().lower()
    for j in range(start_col, df.shape[1]):
        val = str(hdr.iloc[j]).strip().lower()
        if not val:
            continue
        if val == target or target in val:
            return j
    return None

def _row_indices_for_model(df: pd.DataFrame, model_name: str, model_col_idx: int = 3) -> list[int]:
    s = df.iloc[:, model_col_idx].astype(str).str.strip()
    mask = s.str.casefold() == model_name.strip().casefold()
    idx_list = mask[mask].index.tolist()
    return [int(i) for i in idx_list]

def _to_num(x):
    return to_numeric_locale_aware(pd.Series([x])).iloc[0]

@st.cache_data(show_spinner=False)
def load_model_lists(perf_path: Path) -> list[str]:
    models = set()
    for sh in ["Retail", "Handover Model", "Presold"]:
        df = _read_sheet(perf_path, sh)
        col_d = df.iloc[:, 3].astype(str).str.strip()  # D sütunu
        models |= set(col_d[col_d.ne("")])
    return sorted(models)

@st.cache_data(show_spinner=False)
def get_retail_handover_month_only(perf_path: Path, model_name: str, month_abbr: str) -> dict:
    out = {"retail_month": None, "handover_month": None}
    for sh, key_month in [
        ("Retail", "retail_month"),
        ("Handover Model", "handover_month"),
    ]:
        df = _read_sheet(perf_path, sh)
        rows = _row_indices_for_model(df, model_name, model_col_idx=3)
        if not rows:
            continue
        m_col = _month_col_index_by_abbr(df, month_abbr, start_col=5, header_idx=5)
        if m_col is None:
            continue
        vals = []
        for r in rows:
            try:
                vals.append(_to_num(df.iat[r, m_col]))
            except Exception:
                vals.append(pd.NA)
        ser = pd.to_numeric(pd.Series(vals), errors="coerce")
        total = float(ser.fillna(0).sum()) if ser.notna().any() else None
        out[key_month] = total
    return out

@st.cache_data(show_spinner=False)
def get_presold_free(perf_path: Path, model_name: str) -> dict:
    res = {"presold": None, "free": None}
    df = _read_sheet(perf_path, "Presold")
    rows = _row_indices_for_model(df, model_name, model_col_idx=3)
    if not rows:
        return res

    pres, free = [], []
    for r in rows:
        try:
            pres.append(_to_num(df.iat[r, 28]))  # AC
        except Exception:
            pres.append(pd.NA)
        try:
            free.append(_to_num(df.iat[r, 30]))  # AE
        except Exception:
            free.append(pd.NA)

    pres_s = pd.to_numeric(pd.Series(pres), errors="coerce")
    free_s = pd.to_numeric(pd.Series(free), errors="coerce")

    res["presold"] = float(pres_s.fillna(0).sum()) if pres_s.notna().any() else None
    res["free"]    = float(free_s.fillna(0).sum()) if free_s.notna().any() else None
    return res

# ---------- Series -> Showroom KPI (Walk-in / Test Sürüşü) ----------
def _find_series_for_model(perf_path: Path, model_name: str, sheet_name: str = "Series") -> str | None:
    try:
        df = pd.read_excel(perf_path, sheet_name=sheet_name, header=None, engine="openpyxl")
    except Exception:
        return None
    model_cf = str(model_name).strip().casefold()
    for i in range(df.shape[0]):
        row_vals = df.iloc[i, :].astype(str).str.strip().str.casefold()
        if (row_vals == model_cf).any():
            val = df.iat[i, 1]  # B sütunu
            return None if pd.isna(val) else str(val).strip()
    return None

def _get_showroom_kpis(perf_path: Path, series_name: str, sheet_name: str = "BMW Showroom Seri Bazlı"):
    try:
        df = pd.read_excel(perf_path, sheet_name=sheet_name, header=None, engine="openpyxl")
    except Exception:
        return {"walk_cur": None, "walk_prev": None, "test_cur": None, "test_prev": None}

    colB = df.iloc[:, 1].astype(str).str.strip()
    mask = colB.str.casefold() == str(series_name).strip().casefold()
    idx = mask[mask].index.tolist()
    if not idx:
        return {"walk_cur": None, "walk_prev": None, "test_cur": None, "test_prev": None}

    r = idx[0]
    def _num(x): return to_numeric_locale_aware(pd.Series([x])).iloc[0]

    vals = {
        "walk_cur": _num(df.iat[r, 10]) if 10 < df.shape[1] else pd.NA,  # K
        "walk_prev": _num(df.iat[r, 14]) if 14 < df.shape[1] else pd.NA, # O
        "test_cur": _num(df.iat[r, 11]) if 11 < df.shape[1] else pd.NA,  # L
        "test_prev": _num(df.iat[r, 16]) if 16 < df.shape[1] else pd.NA, # Q
    }
    for k,v in list(vals.items()):
        vals[k] = None if pd.isna(v) else float(v)
    return vals

def _delta_html(cur: float | None, prev: float | None) -> str:
    if (cur is not None) and (prev is not None) and (prev != 0):
        delta = (cur - prev) / prev
        cls = "delta-pos" if delta >= 0 else "delta-neg"
        sign = "+" if delta >= 0 else ""
        return (
            f'<div class="kv-delta {cls}">{sign}{delta:.0%}</div>'
            f'<div class="kv-sub">vs önceki ay ({fmt_int(prev)})</div>'
        )
    elif prev is not None:
        return f'<div class="kv-sub">vs önceki ay ({fmt_int(prev)})</div>'
    return ""

# ---------- DIO Model ----------
@st.cache_data(show_spinner=False)
def load_dio_sheet(perf_path: Path, sheet_name: str = "DIO Model") -> pd.DataFrame | None:
    try:
        df = pd.read_excel(perf_path, sheet_name=sheet_name, header=None, engine="openpyxl")
        return df
    except Exception:
        return None

def _find_model_rows_in_dio(df_dio: pd.DataFrame, model_name: str) -> list[int]:
    col = df_dio.iloc[:, 3].astype(str).str.strip()  # D sütunu
    mask = (col.str.casefold() == model_name.strip().casefold())
    return [int(i) for i in mask[mask].index.tolist()]

def _extract_day_headers_dates(df_dio: pd.DataFrame) -> tuple[list[pd.Timestamp], int]:
    headers = df_dio.iloc[5, 4:].tolist()  # E sütunundan sağa
    dates: list[pd.Timestamp] = []
    ncols = 0
    for h in headers:
        if pd.isna(h) or str(h).strip() == "":
            break
        dt = pd.to_datetime(h, dayfirst=True, errors="coerce")
        if pd.isna(dt):
            break
        dates.append(dt)
        ncols += 1
    return dates, ncols

def get_dio_timeseries_and_total(perf_path: Path, model_name: str):
    df_dio = load_dio_sheet(perf_path, "DIO Model")
    if df_dio is None:
        return None, None, "DIO Model sayfası bulunamadı."

    row_idxs = _find_model_rows_in_dio(df_dio, model_name)
    if not row_idxs:
        return None, None, f"'{model_name}' modeli DIO Model sayfasında bulunamadı."

    dates, ncols = _extract_day_headers_dates(df_dio)
    if ncols == 0:
        return None, None, "DIO Model sayfasında E6'dan başlayan tarih başlıkları okunamadı."

    daily_sum = np.zeros(ncols, dtype=float)
    for r in row_idxs:
        vals_raw = df_dio.iloc[r, 4:4+ncols].tolist()
        vals_num = to_numeric_locale_aware(pd.Series(vals_raw)).fillna(0).astype(float).values
        daily_sum += vals_num

    out = pd.DataFrame({"Tarih": dates, "Değer": daily_sum})
    out["TarihLabel"] = out["Tarih"].dt.strftime("%d.%m")
    out["TarihLabel"] = pd.Categorical(
        out["TarihLabel"],
        categories=out["TarihLabel"].tolist(),
        ordered=True
    )

    toplam_vals = []
    for r in row_idxs:
        total_cell = df_dio.iat[r, 4 + ncols] if (4 + ncols) < df_dio.shape[1] else None
        toplam_vals.append(to_numeric_locale_aware(pd.Series([total_cell])).iloc[0])
    toplam_series = pd.to_numeric(pd.Series(toplam_vals), errors="coerce")
    toplam = float(toplam_series.fillna(0).sum())

    return out, toplam, None

# ---------- Sirküler Sheet'ten Kampanyalar ----------
@st.cache_data(show_spinner=False)
def get_campaigns(perf_path: Path, model_name: str, sheet_name: str = "Sirküler"):
    """
    'Sirküler' sayfasında:
    - Model isimleri C sütunu (C3'ten itibaren)
    - Nakit destek = F sütunu
    - Takas destek = G sütunu
    - Kredi Kampanyası = I sütunu
    İlk eşleşen satırı döndür.
    """
    try:
        df = pd.read_excel(perf_path, sheet_name=sheet_name, header=None, engine="openpyxl")
    except Exception:
        return None

    # Excel C3 -> df.iloc[2, 2]
    sub_start_row = 2  # index 2 ~ satır 3
    model_col_idx = 2  # C
    nakit_idx = 5      # F
    takas_idx = 6      # G
    kredi_idx = 8      # I

    # ilgili satırı bul
    model_cf = model_name.strip().casefold()
    for r in range(sub_start_row, df.shape[0]):
        cell_model = str(df.iat[r, model_col_idx]).strip()
        if cell_model == "" or cell_model.lower() == "nan":
            continue
        if cell_model.casefold() == model_cf:
            nakit_val = df.iat[r, nakit_idx] if nakit_idx < df.shape[1] else None
            takas_val = df.iat[r, takas_idx] if takas_idx < df.shape[1] else None
            kredi_val = df.iat[r, kredi_idx] if kredi_idx < df.shape[1] else None
            return {
                "Nakit destek": "" if pd.isna(nakit_val) else str(nakit_val),
                "Takas destek": "" if pd.isna(takas_val) else str(takas_val),
                "Kredi Kampanyası": "" if pd.isna(kredi_val) else str(kredi_val),
            }

    return None

# ---------- UI: Model Aylık Performans Bölümü ----------
def build_monthly_performance_ui(perf_path: Path):
    st.markdown("## Model Aylık Performans (Retail / Handover / Presold / Free / Walk-in / Test Sürüşü)")

    if perf_path is None:
        st.warning("Aylık performans dosyası bulunamadı. Lütfen 'Model aylık performans.xlsx' dosyasını `data/` klasörüne koy.")
        return

    # ---- CSS ----
    st.markdown("""
    <style>
    .kv-row {
        display: grid;
        grid-template-columns: repeat(6, minmax(160px, 1fr));
        gap: 12px;
        align-items: stretch;
    }
    .kv-box {
        background:#eaf3ff;
        border:1px solid #d6e7ff;
        border-radius:18px;
        padding:16px 18px;
        text-align:center;
    }
    .kv-title { font-size:12px; color:#2a4a7a; letter-spacing:.35px; text-transform:uppercase; margin-bottom:6px; }
    .kv-value { font-size:26px; font-weight:800; color:#0c203a; }
    .kv-sub { font-size:12px; color:#5e738f; margin-top:4px; }
    .kv-delta { display:inline-block; font-weight:700; margin-top:6px; margin-right:6px; }
    .delta-pos { color:#0f9d58; }
    .delta-neg { color:#d93025; }
    @media (max-width: 1200px){
        .kv-row { grid-template-columns: repeat(3, minmax(160px, 1fr)); }
    }
    @media (max-width: 800px){
        .kv-row { grid-template-columns: repeat(2, minmax(140px, 1fr)); }
    }
    </style>
    """, unsafe_allow_html=True)

    # İçinde bulunduğumuz ay kısaltması (Istanbul TZ)
    ist_tz = zoneinfo.ZoneInfo("Europe/Istanbul")
    month_abbr = datetime.now(ist_tz).strftime("%b")

    # Model seçimi
    model_options = load_model_lists(perf_path)
    if not model_options:
        st.info("Model listesi boş görünüyor (D sütunu boş olabilir).")
        return

    selected_perf_model = st.selectbox(
        "Model seçiniz (Aylık performans)",
        options=model_options,
        index=0,
        key="perf_model_select"
    )

    # 1) Retail / Handover (ilgili ay kolonunu toplayarak)
    rh = get_retail_handover_month_only(perf_path, selected_perf_model, month_abbr)

    # 2) Presold / Free toplu
    pf = get_presold_free(perf_path, selected_perf_model)

    # 3) Walk-in / Test Sürüşü (Series -> BMW Showroom Seri Bazlı)
    series_name = _find_series_for_model(perf_path, selected_perf_model, sheet_name="Series")
    showroom = (
        _get_showroom_kpis(perf_path, series_name, sheet_name="BMW Showroom Seri Bazlı")
        if series_name
        else {"walk_cur": None, "walk_prev": None, "test_cur": None, "test_prev": None}
    )

    # KPI kutuları (6 kutu aynı sırada)
    html = f"""
    <div class="kv-row">
      <div class="kv-box">
        <div class="kv-title">Retail ({month_abbr})</div>
        <div class="kv-value">{fmt_int(rh.get("retail_month"))}</div>
      </div>

      <div class="kv-box">
        <div class="kv-title">Handover ({month_abbr})</div>
        <div class="kv-value">{fmt_int(rh.get("handover_month"))}</div>
      </div>

      <div class="kv-box">
        <div class="kv-title">Presold</div>
        <div class="kv-value">{fmt_int(pf.get("presold"))}</div>
      </div>

      <div class="kv-box">
        <div class="kv-title">Free</div>
        <div class="kv-value">{fmt_int(pf.get("free"))}</div>
      </div>

      <div class="kv-box">
        <div class="kv-title">Walk-in</div>
        <div class="kv-value">{fmt_int(showroom.get("walk_cur"))}</div>
        {_delta_html(showroom.get("walk_cur"), showroom.get("walk_prev"))}
      </div>

      <div class="kv-box">
        <div class="kv-title">Test Sürüşü</div>
        <div class="kv-value">{fmt_int(showroom.get("test_cur"))}</div>
        {_delta_html(showroom.get("test_cur"), showroom.get("test_prev"))}
      </div>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)
    st.caption(f"Kaynak: {perf_path.name}  •  Ay: {month_abbr}")

    # ---- Sirküler Kampanyaları ----
    st.markdown("#### Kampanyalar (Sirküler)")
    campaign_info = get_campaigns(perf_path, selected_perf_model, sheet_name="Sirküler")
    if campaign_info is None:
        st.info("Seçilen model için kampanya bilgisi bulunamadı.")
    else:
        camp_df = pd.DataFrame([campaign_info])
        st.dataframe(
            camp_df,
            hide_index=True,
            use_container_width=False
        )

    # ---- Günlük DIO Model Grafiği + Toplam ----
    st.markdown("### Günlük DIO Model")
    dio_df, dio_total, dio_err = get_dio_timeseries_and_total(perf_path, selected_perf_model)
    if dio_err:
        st.warning(dio_err)
    else:
        if dio_df is None or len(dio_df) == 0:
            st.info("Seçilen model için Günlük DIO Model verisi bulunamadı.")
        else:
            import altair as alt
            BAR_COLOR = "#2a4a7a"

            baslik = f"{selected_perf_model} • Günlük DIO Model • Toplam: {dio_total:,.0f}".replace(",", ".")
            base = alt.Chart(dio_df).encode(
                x=alt.X("TarihLabel:N", title="Gün", sort=list(dio_df["TarihLabel"].astype(str))),
                y=alt.Y("Değer:Q", title="Değer", scale=alt.Scale(nice=True, zero=True)),
                tooltip=[
                    alt.Tooltip("Tarih:T", title="Tarih", format="%d.%m.%Y"),
                    alt.Tooltip("Değer:Q", format=",.0f")
                ]
            )
            bars = base.mark_bar(color=BAR_COLOR).properties(height=260)
            labels = base.mark_text(dy=-5, fontSize=11, color=BAR_COLOR) \
                          .encode(text=alt.Text("Değer:Q", format=",.0f"))

            chart = (bars + labels).resolve_scale(y='shared').properties(title=baslik)
            st.altair_chart(chart, use_container_width=True)


# ================== ODMD SONUÇLARI (Monthly Basis) ==================

@st.cache_data(show_spinner=False)
def load_focus_segment_df_from_perf(perf_path: Path, sheet_name: str = "Monthly Basis"):
    """
    'Model aylık performans.xlsx' içerisindeki 'Monthly Basis' sayfasından
    D:S aralığını (D=Marka, E=Model, G..R=Jan..Dec, S=YTD) okur.
    """
    if perf_path is None or not perf_path.exists():
        raise ValueError("Model aylık performans dosyası bulunamadı.")

    xls = pd.ExcelFile(perf_path, engine="openpyxl")
    if sheet_name not in xls.sheet_names:
        raise ValueError(f"'{sheet_name}' sayfası dosyada yok. Mevcut sayfalar: {xls.sheet_names}")

    col_names = ["Marka", "Model"] + MONTHS_EN + ["YTD"]  # 15 kolon
    usecols_letters = ["D", "E"] + [chr(c) for c in range(ord("G"), ord("R")+1)] + ["S"]

    df = pd.read_excel(
        xls,
        sheet_name=sheet_name,
        header=None,
        skiprows=9,
        usecols=",".join(usecols_letters),
    )
    df.columns = col_names

    def to_num(x):
        if isinstance(x, str):
            x = x.strip().replace(".", "").replace(",", ".")
            x = re.sub(r"[^\d\.-]", "", x)
        return pd.to_numeric(x, errors="coerce")

    for c in MONTHS_EN + ["YTD"]:
        df[c] = df[c].apply(to_num)

    # Grup ID (D sütunundaki boş satırlar grup ayırıcısı)
    group_id = []
    g = -1
    for _, row in df.iterrows():
        marka = row["Marka"]
        if pd.isna(marka) or (isinstance(marka, str) and marka.strip() == ""):
            group_id.append(np.nan)
            g += 1
        else:
            if len(group_id) == 0 or pd.isna(group_id[-1]):
                g = g if g >= 0 else 0
            group_id.append(g)
    df["group_id"] = group_id

    data_df = df[~df["Marka"].isna()].copy()
    data_df["Marka"] = data_df["Marka"].astype(str).str.strip()
    data_df["Model"] = data_df["Model"].astype(str).str.strip()
    return data_df

def current_month_info():
    now = datetime.now(IST_TZ)
    cur_month_idx = now.month - 1               # içinde bulunduğumuz ay index
    prev_idx = (cur_month_idx - 1) % 12         # bir önceki ay
    last3 = [(prev_idx - 2) % 12, (prev_idx - 1) % 12, prev_idx]  # son 3 ay
    return cur_month_idx, prev_idx, last3, now

def compute_metrics(df: pd.DataFrame):
    cur_idx, prev_idx, last3, _ = current_month_info()
    prev_month_name = MONTHS_EN[prev_idx]
    last3_names = [MONTHS_EN[i] for i in last3]

    work = df.copy()
    work["Aylık Satış"] = work[prev_month_name]
    work["3 Aylık Satış"] = work[last3_names].mean(axis=1, skipna=True)

    # Ocak'ta 0'a bölmeyi engelle
    denom = max(cur_idx, 1)
    work["YTD Satış"] = work["YTD"] / denom

    out = work[["Marka", "Model", "Aylık Satış", "3 Aylık Satış", "YTD Satış", "YTD", "group_id"]].copy()
    return out, prev_month_name, last3_names, denom

def style_bmw_first(df: pd.DataFrame):
    df = df.copy()
    df["__bmw__"] = (df["Marka"].str.upper() == "BMW").astype(int)
    df = df.sort_values(by=["__bmw__", "Marka", "Model"], ascending=[False, True, True]).drop(columns="__bmw__")
    return df

def format_int(x):
    if pd.isna(x):
        return ""
    try:
        return f"{int(round(x)):,}".replace(",", ".")
    except:
        return str(x)

def build_odmd_ui(perf_path: Path):
    st.markdown("## ODMD Sonuçları")

    try:
        data_df = load_focus_segment_df_from_perf(perf_path, sheet_name="Monthly Basis")
    except ValueError as e:
        st.warning(str(e))
        return

    calc_df, prev_month_name, last3_names, ytd_denom = compute_metrics(data_df)

    # Boş YTD olanları at
    calc_df = calc_df[calc_df["YTD"].fillna(0) != 0].copy()

    # BMW modelleri -> filtre
    bmw_models = (
        calc_df.loc[calc_df["Marka"].str.upper() == "BMW", "Model"]
        .dropna()
        .drop_duplicates()
        .tolist()
    )
    if not bmw_models:
        st.info("BMW modeli bulunamadı. Lütfen 'Monthly Basis' sayfasındaki verileri kontrol edin.")
        return

    selected_bmw = st.selectbox(
        "BMW Model Filtresi",
        options=bmw_models,
        index=0,
        help="Bir BMW modeli seçtiğinde, o modelin rakip grubundaki satırlar listelenir.",
        key="odmd_bmw_filter"
    )

    target_groups = calc_df.loc[
        (calc_df["Marka"].str.upper() == "BMW") & (calc_df["Model"] == selected_bmw),
        "group_id"
    ].dropna().unique()

    if len(target_groups) == 0:
        st.info("Seçilen BMW modeline ait grup bulunamadı.")
        return

    gid = target_groups[0]
    group_view = calc_df[calc_df["group_id"] == gid].copy()
    view = group_view[["Marka", "Model", "Aylık Satış", "3 Aylık Satış", "YTD Satış"]].copy()
    view = style_bmw_first(view)

    cur_idx, prev_idx, last3, now = current_month_info()

    st.caption(
        f"Kaynak: {Path(perf_path).name} • Sayfa: 'Monthly Basis' • "
        f"İçinde bulunulan ay: **{MONTHS_EN[cur_idx]}** • "
        f"Aylık Satış = **{MONTHS_EN[prev_idx]}** • "
        f"3 Aylık = {', '.join(MONTHS_EN[i] for i in last3)} • "
        f"YTD Ortalama bölünen: **{ytd_denom}**"
    )

    styled = (
        view.style
        .apply(
            lambda s: [
                "font-weight: 700"
                if (s.name in view.index and view.loc[s.name, 'Marka'].upper()=='BMW')
                else ""
                for _ in s
            ],
            axis=1
        )
        .format({
            "Aylık Satış": format_int,
            "3 Aylık Satış": format_int,
            "YTD Satış": format_int
        })
    )

    st.dataframe(styled, use_container_width=True)

    csv_bytes = view.to_csv(index=False).encode("utf-8")
    st.download_button(
        "CSV indir (filtrelenmiş)",
        data=csv_bytes,
        file_name=(
            f"odmd_sonuclari_"
            f"{Path(perf_path).stem.replace(' ','_')}_"
            f"{selected_bmw.replace(' ','_')}.csv"
        ),
        mime="text/csv"
    )

# ================== UYGULAMA AKIŞI ==================
def main():
    # 1) Rakip Karşılaştırma
    price_excel = find_price_excel(DATA_DIR)
    if price_excel is None or not price_excel.exists():
        st.error(
            "`data/` klasöründe fiyat karşılaştırma için bir .xlsx bulunamadı. "
            f"Öncelik: `{PRICE_FILE_NAME}`."
        )
    else:
        df_raw = load_price_compare(price_excel)
        build_price_compare_ui(df_raw, price_excel)

    st.markdown("---")

    # 2) Model Aylık Performans + Kampanyalar (Sirküler) + Günlük DIO Model
    perf_excel = find_performance_workbook(DATA_DIR)
    build_monthly_performance_ui(perf_excel)

    st.markdown("---")

    # 3) ODMD Sonuçları (Monthly Basis)
    build_odmd_ui(perf_excel)

if __name__ == "__main__":
    main()
