import streamlit as st
import pandas as pd
import re
from pathlib import Path
from datetime import datetime
import zoneinfo

# ================== Genel Ayarlar ==================
st.set_page_config(page_title="Fiyat Karşılaştırması Dashboard", layout="wide")
DATA_DIR = Path("data")

# Dosya isimleri
PRICE_FILE_NAME = "Fiyat Karşılaştırması_v4.xlsx"      # Rakip karşılaştırma
PERF_FILE_NAME  = "Model aylık performans.xlsx"        # Retail/Handover/Presold/Free

# ================== Yardımcı Fonksiyonlar ==================
def to_numeric_locale_aware(s: pd.Series) -> pd.Series:
    t = s.astype(str).str.strip()
    t = t.replace({
        "": pd.NA, "-": pd.NA, "—": pd.NA, "–": pd.NA,
        "N/A": pd.NA, "n/a": pd.NA, "na": pd.NA, "NaN": pd.NA,
        "#N/A": pd.NA, "#NA": pd.NA, "#VALUE!": pd.NA,
    })
    t = t.str.replace(r"[^\d,.\-]", "", regex=True)         # para/etiket sil
    t = t.str.replace(r"\.(?=\d{3}(\D|$))", "", regex=True) # binlik noktası sil
    t = t.str.replace(",", ".", regex=False)                # ondalık virgülü noktaya
    return pd.to_numeric(t, errors="coerce")

def parse_percent_series_mixed(s: pd.Series) -> pd.Series:
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

# ================== Rakip Karşılaştırma (Önceki Kurgu) ==================
def find_price_excel(data_dir: Path) -> Path | None:
    if not data_dir.exists():
        return None
    exact = data_dir / PRICE_FILE_NAME
    if exact.exists():
        return exact
    files = list(data_dir.glob("*.xlsx"))
    if not files:
        return None
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
    df["Marka"] = df["Marka"].replace(r"^\s*$", pd.NA, regex=True)
    df["__group_id__"] = df["Marka"].isna().cumsum()

    h_col = "Stoktaki en uygun otomobil fiyatı"
    h_num = to_numeric_locale_aware(df[h_col])
    is_na = df[h_col].isna() | h_num.isna()
    is_zero = h_num.fillna(0).eq(0)
    df = df[~(is_na | is_zero)].copy()

    df["İndirim oranı"] = parse_percent_series_mixed(df["İndirim oranı"])
    return df

def build_price_compare_ui(df_raw: pd.DataFrame, source_path: Path):
    st.markdown("## BMW Rakip Karşılaştırma")

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
        pkg_list = sorted(df_bmw.loc[df_bmw["Model"].astype(str) == selected_model, "Paket"].astype(str).unique())
        if len(pkg_list) == 0:
            st.info("Seçilen model için paket bulunamadı.")
            return
        selected_pkg = st.selectbox("Paket", options=pkg_list, index=0, key="bmw_pkg")

    df_sel = df_bmw[(df_bmw["Model"].astype(str) == selected_model) & (df_bmw["Paket"].astype(str) == selected_pkg)]
    if df_sel.empty:
        st.info("Seçime uygun satır bulunamadı.")
        return

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
        if (str(row["Marka"]).strip().upper() == "BMW") and \
           (str(row["Model"]) == selected_model) and \
           (str(row["Paket"]) == selected_pkg):
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

# ================== Aylık Performans (Retail/Handover/Presold/Free) ==================
REQUIRED_SHEETS = {"Retail", "Handover Model", "Presold"}

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
    return None

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
        if val == target or target in val:  # 'Oct' veya 'October' eşleşsin
            return j
    return None

def _row_index_for_model(df: pd.DataFrame, model_name: str, model_col_idx: int = 3) -> int | None:
    s = df.iloc[:, model_col_idx].astype(str).str.strip()
    mask = s.str.casefold() == model_name.strip().casefold()
    idx = mask[mask].index
    if len(idx):
        return int(idx[0])
    return None

def _to_num(x):
    return to_numeric_locale_aware(pd.Series([x])).iloc[0]

@st.cache_data(show_spinner=False)
def load_model_lists(perf_path: Path) -> list[str]:
    models = set()
    for sh in ["Retail", "Handover Model", "Presold"]:
        df = _read_sheet(perf_path, sh)
        col_d = df.iloc[:, 3].astype(str).str.strip()
        models |= set(col_d[col_d.ne("")])
    return sorted(models)

@st.cache_data(show_spinner=False)
def get_retail_handover_month_only(perf_path: Path, model_name: str, month_abbr: str) -> dict:
    """
    Sadece içinde bulunduğumuz ay değerlerini döner:
      - Retail (month)
      - Handover (month)
    """
    out = {"retail_month": None, "handover_month": None}
    for sh, key_month in [
        ("Retail", "retail_month"),
        ("Handover Model", "handover_month"),
    ]:
        df = _read_sheet(perf_path, sh)
        r_idx = _row_index_for_model(df, model_name, model_col_idx=3)
        if r_idx is None:
            continue
        m_col = _month_col_index_by_abbr(df, month_abbr, start_col=5, header_idx=5)
        m_val = None
        if m_col is not None:
            try:
                m_val = _to_num(df.iat[r_idx, m_col])
            except Exception:
                m_val = None
        out[key_month] = m_val
    return out

@st.cache_data(show_spinner=False)
def get_presold_free(perf_path: Path, model_name: str) -> dict:
    res = {"presold": None, "free": None}
    df = _read_sheet(perf_path, "Presold")
    r_idx = _row_index_for_model(df, model_name, model_col_idx=3)
    if r_idx is None:
        return res
    try:
        res["presold"] = _to_num(df.iat[r_idx, 28])  # AC
    except Exception:
        res["presold"] = None
    try:
        res["free"] = _to_num(df.iat[r_idx, 30])     # AE
    except Exception:
        res["free"] = None
    return res

def build_monthly_performance_ui(perf_path: Path):
    with st.expander("📊 Model Aylık Performans (Retail / Handover / Presold / Free)", expanded=True):
        if perf_path is None:
            st.warning("Aylık performans dosyası bulunamadı. Lütfen 'Model aylık performans.xlsx' dosyasını `data/` klasörüne koy.")
            return

        # ---- CSS: Açık mavi kutular ----
        st.markdown("""
        <style>
        .kv-row {display:flex; gap:12px; flex-wrap:wrap;}
        .kv-box {
            flex:1 1 0;
            background:#eaf3ff;                 /* açık mavi */
            border:1px solid #d6e7ff;
            border-radius:12px;
            padding:14px 16px;
            min-width:180px;
            text-align:center;
        }
        .kv-title {
            font-size:12px;
            color:#2a4a7a;
            letter-spacing:.3px;
            text-transform:uppercase;
            margin-bottom:6px;
        }
        .kv-value {
            font-size:22px;
            font-weight:700;
        }
        </style>
        """, unsafe_allow_html=True)

        # Kullanıcı TZ: Europe/Istanbul, ay başlığı İngilizce kısaltma ('Oct' gibi)
        ist_tz = zoneinfo.ZoneInfo("Europe/Istanbul")
        month_abbr = datetime.now(ist_tz).strftime("%b")  # 'Oct', 'Nov', ...

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

        rh = get_retail_handover_month_only(perf_path, selected_perf_model, month_abbr)
        pf = get_presold_free(perf_path, selected_perf_model)

        def _fmt(v): 
            return "—" if v is None or pd.isna(v) else f"{float(v):,.0f}"

        # ---- 4 kutu tek satır ----
        st.markdown(
            f"""
            <div class="kv-row">
              <div class="kv-box">
                <div class="kv-title">Retail ({month_abbr})</div>
                <div class="kv-value">{_fmt(rh.get("retail_month"))}</div>
              </div>
              <div class="kv-box">
                <div class="kv-title">Handover ({month_abbr})</div>
                <div class="kv-value">{_fmt(rh.get("handover_month"))}</div>
              </div>
              <div class="kv-box">
                <div class="kv-title">Presold</div>
                <div class="kv-value">{_fmt(pf.get("presold"))}</div>
              </div>
              <div class="kv-box">
                <div class="kv-title">Free</div>
                <div class="kv-value">{_fmt(pf.get("free"))}</div>
              </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        st.caption(f"Kaynak: {perf_path.name}  •  Ay: {month_abbr}")

# ================== Uygulama Akışı ==================
def main():
    # 1) Rakip Karşılaştırma
    price_excel = find_price_excel(DATA_DIR)
    if price_excel is None or not price_excel.exists():
        st.error("`data/` klasöründe fiyat karşılaştırma için bir .xlsx bulunamadı. "
                 f"Öncelik: `{PRICE_FILE_NAME}`.")
    else:
        df_raw = load_price_compare(price_excel)
        build_price_compare_ui(df_raw, price_excel)

    st.markdown("---")

    # 2) Aylık Performans Kutucukları (sadece aylık değerler)
    perf_excel = find_performance_workbook(DATA_DIR)
    build_monthly_performance_ui(perf_excel)

if __name__ == "__main__":
    main()
