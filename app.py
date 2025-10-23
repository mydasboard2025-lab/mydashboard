import streamlit as st
import pandas as pd
import re
from pathlib import Path
from datetime import datetime
import zoneinfo

import numpy as np
import pytz
import glob
import os

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



# === Yeni Bölüm: Satış Performansı Tablosu (Aylık / 3 Aylık / YTD) — Monthly Basis seçimi ===


# -------------------------------------------------------------
# Yardımcılar
# -------------------------------------------------------------
IST_TZ = pytz.timezone("Europe/Istanbul")
MONTHS_EN = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

def _pick_monthly_basis_file(search_dir="data"):
    """
    Sadece adında 'monthly' ve 'basis' geçen Excel dosyalarını (.xlsx/.xlsm) arar (case-insensitive).
    Eğer birden fazla varsa, dosya değiştirilme zamanına göre en güncelini seçer.
    """
    patterns = [
        os.path.join(search_dir, "*monthly*basis*.xlsx"),
        os.path.join(search_dir, "*monthly*basis*.xlsm"),
        os.path.join(search_dir, "*Monthly*Basis*.xlsx"),
        os.path.join(search_dir, "*Monthly*Basis*.xlsm"),
    ]
    candidates = []
    for pat in patterns:
        candidates.extend(glob.glob(pat))
    # unique & sort by mtime desc
    candidates = list({Path(p).resolve() for p in candidates})
    candidates = sorted(candidates, key=lambda p: p.stat().st_mtime, reverse=True)
    return str(candidates[0]) if candidates else None

@st.cache_data(show_spinner=False)
def load_focus_segment_df(file_path: str, sheet_name=0):
    """
    Excel'den D:S aralığını okur.
    - Veriler 10. satırdan itibaren olduğu için skiprows=9 kullanıyoruz.
    - Kolon isimlerini sabit veriyoruz: D=Marka, E=Model, G..R=Jan..Dec, S=YTD
    """
    col_names = ["Marka","Model"] + MONTHS_EN + ["YTD"]
    df = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=None,
        skiprows=9,          # 10. satırdan itibaren veri
        usecols="D:S",       # D..S
        engine="openpyxl" if file_path.lower().endswith((".xlsx",".xlsm",".xltx",".xltm")) else None
    )
    df.columns = col_names

    # String sayıları normalize et (1.234,56 → 1234.56)
    def to_num(x):
        if isinstance(x, str):
            x = x.strip().replace(".", "").replace(",", ".")
            x = re.sub(r"[^\d\.-]", "", x)
        return pd.to_numeric(x, errors="coerce")

    for c in MONTHS_EN + ["YTD"]:
        df[c] = df[c].apply(to_num)

    # Grup ID üretimi: D sütunundaki boşluklar grup ayırıcıdır
    group_id = []
    g = -1
    for _, row in df.iterrows():
        marka = row["Marka"]
        if pd.isna(marka) or (isinstance(marka, str) and marka.strip() == ""):
            g += 1
            group_id.append(np.nan)   # separator satır
        else:
            if len(group_id) == 0 or pd.isna(group_id[-1]):
                g = g if g >= 0 else 0
            group_id.append(g)
    df["group_id"] = group_id

    # Sadece dolu satırlar (separator'lar hariç)
    data_df = df[~df["Marka"].isna()].copy()
    data_df["Marka"] = data_df["Marka"].astype(str).str.strip()
    data_df["Model"] = data_df["Model"].astype(str).str.strip()

    return data_df

def current_month_info():
    """İstanbul saatine göre ay bilgisi ve önceki ay/son 3 ay indeksleri."""
    now = datetime.now(IST_TZ)
    cur_month_num = now.month  # 1..12
    cur_month_idx = cur_month_num - 1
    prev_idx = (cur_month_idx - 1) % 12
    last3 = [ (prev_idx - 2) % 12, (prev_idx - 1) % 12, prev_idx ]
    return cur_month_idx, prev_idx, last3, now

def compute_metrics(df: pd.DataFrame):
    """
    - Aylık Satış: içinde bulunulan ayın bir önceki ayı
    - 3 Aylık Satış: önceki 3 ayın ortalaması
    - YTD Satış (ortalama): S / (içinde bulunulan ay - 1)  (Ocak için S/1)
    """
    cur_idx, prev_idx, last3, _ = current_month_info()
    prev_month_name = MONTHS_EN[prev_idx]
    last3_names = [MONTHS_EN[i] for i in last3]

    work = df.copy()
    work["Aylık Satış"] = work[prev_month_name]
    work["3 Aylık Satış"] = work[last3_names].mean(axis=1, skipna=True)

    denom = max(cur_idx, 1)  # Ocak'ta 0'a bölmeyi engelle
    work["YTD Satış"] = work["YTD"] / denom

    out = work[["Marka","Model","Aylık Satış","3 Aylık Satış","YTD Satış","YTD","group_id"]].copy()
    return out, prev_month_name, last3_names, denom

def style_bmw_first(df: pd.DataFrame):
    """BMW’yi üstte sırala; BMW satırını kalın gösterme styler'la yapılacak."""
    df = df.copy()
    df["__bmw__"] = (df["Marka"].str.upper() == "BMW").astype(int)
    df = df.sort_values(by=["__bmw__","Marka","Model"], ascending=[False, True, True]).drop(columns="__bmw__")
    return df

def format_int(x):
    if pd.isna(x):
        return ""
    try:
        return f"{int(round(x)):,}".replace(",", ".")
    except:
        return str(x)

# -------------------------------------------------------------
# UI: Bölüm
# -------------------------------------------------------------
st.markdown("## Satış Performansı (Aylık / 3 Aylık / YTD) — Monthly Basis")

# Sadece adı 'monthly basis' içeren dosyaları ara
monthly_file = _pick_monthly_basis_file(search_dir="data")
if monthly_file is None:
    st.warning("`data/` klasöründe adı **'monthly basis'** içeren Excel dosyası (.xlsx/.xlsm) bulunamadı.\nÖrn: `Monthly Basis - Focus Segment Retail Comparision 09-2025.xlsm`")
    st.stop()

data_df = load_focus_segment_df(monthly_file)

# Hesaplamalar
calc_df, prev_month_name, last3_names, ytd_denom = compute_metrics(data_df)

# S=0 olanları dışla
calc_df = calc_df[calc_df["YTD"].fillna(0) != 0].copy()

# Filtre: yalnızca BMW modelleri
bmw_models = (calc_df.loc[calc_df["Marka"].str.upper() == "BMW", "Model"]
              .dropna().drop_duplicates().tolist())
if not bmw_models:
    st.info("BMW modeli bulunamadı. Lütfen 'monthly basis' dosyasını kontrol edin.")
    st.stop()

selected_bmw = st.selectbox(
    "BMW Model Filtresi",
    options=bmw_models,
    index=0,
    help="Bir BMW modeli seçtiğinde, o modelin rakip grubundaki satırlar listelenir."
)

# Seçilen BMW modelinin grup kimliği
target_groups = calc_df.loc[
    (calc_df["Marka"].str.upper() == "BMW") & (calc_df["Model"] == selected_bmw),
    "group_id"
].dropna().unique()
if len(target_groups) == 0:
    st.warning("Seçilen BMW modeline ait grup bulunamadı.")
    st.stop()

gid = target_groups[0]
group_view = calc_df[calc_df["group_id"] == gid].copy()

# Görüntü ve sıralama
view = group_view[["Marka","Model","Aylık Satış","3 Aylık Satış","YTD Satış"]].copy()
view = style_bmw_first(view)

# Bilgi notu
cur_idx, prev_idx, last3, now = current_month_info()
st.caption(
    f"Dosya: `{Path(monthly_file).name}` • "
    f"İçinde bulunulan ay: **{MONTHS_EN[cur_idx]}** • "
    f"Aylık Satış = **{MONTHS_EN[prev_idx]}** • "
    f"3 Aylık = {', '.join(MONTHS_EN[i] for i in last3)} • "
    f"YTD Ortalama bölünen: **{ytd_denom}**"
)

# Gösterim: BMW satırını kalın yap
styled = (view.style
    .apply(lambda s: ["font-weight: 700" if (s.name in view.index and view.loc[s.name, "Marka"].upper()=="BMW") else "" for _ in s], axis=1)
    .format({"Aylık Satış": format_int, "3 Aylık Satış": format_int, "YTD Satış": format_int})
)

st.dataframe(styled, use_container_width=True)

# CSV indir
csv_bytes = view.to_csv(index=False).encode("utf-8")
st.download_button(
    "CSV indir (filtrelenmiş)",
    data=csv_bytes,
    file_name=f"satis_performansi_{Path(monthly_file).stem.replace(' ','_')}_{selected_bmw.replace(' ','_')}.csv",
    mime="text/csv"
)
# === /Bölüm Sonu ===
