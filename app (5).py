
import re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from datetime import datetime

# ---------------- Page & Style ----------------
st.set_page_config(
    page_title="September AMS Analysis",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
.main {padding: 0rem 1rem;}
h1 {color:#0b1f44;font-weight:800;letter-spacing:.2px;border-bottom:3px solid #2ecc71;padding-bottom:10px;}
h2 {color:#0b1f44;font-weight:700;margin-top:1.0rem;margin-bottom:.6rem;}
.kpi {background:#fff;border:1px solid #e6e6e6;border-radius:14px;padding:16px; height:96px;}
.k-num {font-size:32px;font-weight:800;color:#0b1f44;line-height:1.0;}
.k-cap {font-size:13px;color:#6b7280;margin-top:4px;}
.dataframe {font-size: 14px !important;}
.stTabs [data-baseweb="tab-list"] {gap: 8px;}
.stTabs [data-baseweb="tab"] {height: 48px;padding-left: 20px;padding-right: 20px;background-color: #f8f9fa;border-radius: 8px 8px 0 0;}
.stTabs [aria-selected="true"] {background-color: #0b1f44;color: white;}
</style>
""", unsafe_allow_html=True)

st.title("September AMS Analysis")

# ---------------- Constants & Keywords ----------------
# Accounts for RP tab (AMS sheet)
RP_ACCOUNTS = {
    "Marken Ltd",
    "QIAGEN GmbH Weekly",
    "Fisher Clinical Services",
    "Agilent Technologies Deutschland GmbH",
    "Patheon Biologics BV",
    "Delpharm Development Leiden BV",
    "Abbott Biologicals BV",
    "Fisher BioServices Netherlands BV",
    "Abbott Healthcare Products BV",
    "UNIVERSAL PICTURES INTERNATIONAL NETHERLANDS",
    "Patheon UK",
    "VERACYTE INC",
    "Tosoh Europe",
    "Exnet Services",
    "Nobel Biocare Distribution Center BV",
}

# Healthcare / Aviation classification keywords
HEALTHCARE_KEYWORDS = [
    'pharma','medical','health','bio','clinical','hospital','diagnostic','therapeut',
    'laborator','patholog','oncolog','genetic','genomic','molecular','vaccine','abbott',
    'marken','fisher','patheon','qiagen','veracyte','delpharm','agilent','biocare','tosoh',
    'biolog','life science','lifescience','specimen','reagent','immuno','assay','curium','medtronic'
]

AVIATION_KEYWORDS = [
    'airline','airport','aviation','aircraft','aerospace','cargo','freight','express','courier',
    'lufthansa','delta','american airlines','british airways','klm','air france','iberia','easyjet',
    'end eavor air','storm aviation','heathrow','spairliners','mnx','tokyo electron'
]

CTRL_REGEX = re.compile(r"\b(agent|del\s*agt|delivery\s*agent|customs|warehouse|w/house)\b", re.I)

# ---------------- Helpers ----------------
def _excel_to_dt(s: pd.Series) -> pd.Series:
    """Parse to datetime; if many NaT, try Excel serial numbers."""
    out = pd.to_datetime(s, errors="coerce")
    if out.isna().mean() > 0.5:
        num  = pd.to_numeric(s, errors="coerce")
        out2 = pd.to_datetime("1899-12-30") + pd.to_timedelta(num, unit="D")
        out  = out.where(~out.isna(), out2)
    return out

def _coerce_num(x):
    return pd.to_numeric(x, errors="coerce").fillna(0)

def _contains_any(text: str, keywords: list[str]) -> bool:
    if not isinstance(text, str): 
        return False
    lower = text.lower()
    return any(k in lower for k in keywords)

def is_healthcare(account_name: str, sheet_name: str | None = None) -> bool:
    if sheet_name == "Aviation SVC":
        return False
    return _contains_any(str(account_name), HEALTHCARE_KEYWORDS)

def is_aviation(account_name: str, sheet_name: str | None = None) -> bool:
    if sheet_name == "Aviation SVC":
        return True
    return _contains_any(str(account_name), AVIATION_KEYWORDS)

def add_otp_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Add POD/Target and OTP columns (Gross/Net)."""
    d = df.copy()
    # Dates
    if "POD DATE/TIME" in d.columns:
        d["_pod"] = _excel_to_dt(d["POD DATE/TIME"])
    else:
        d["_pod"] = pd.NaT
    # Target: UPD DEL preferred, else QDT
    target_col = "UPD DEL" if "UPD DEL" in d.columns else ("QDT" if "QDT" in d.columns else None)
    d["_target"] = _excel_to_dt(d[target_col]) if target_col else pd.NaT

    # QC controllability
    qc_col = "QC NAME" if "QC NAME" in d.columns else None
    if qc_col:
        d["QC_NAME_CLEAN"] = d[qc_col].astype(str)
        d["Is_Controllable"] = d["QC_NAME_CLEAN"].str.contains(CTRL_REGEX, na=False)
    else:
        d["Is_Controllable"] = False

    # Pieces numeric
    if "PIECES" in d.columns:
        d["PIECES"] = _coerce_num(d["PIECES"]).astype(float)
    else:
        d["PIECES"] = 0.0

    # OTP
    ok = d["_pod"].notna() & d["_target"].notna()
    d["On_Time_Gross"] = False
    d.loc[ok, "On_Time_Gross"] = d.loc[ok, "_pod"] <= d.loc[ok, "_target"]
    d["Late"] = ~d["On_Time_Gross"]
    d["On_Time_Net"] = d["On_Time_Gross"] | (d["Late"] & ~d["Is_Controllable"])
    return d

def filter_ams_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Keep rows where DEP or ARR equals 'AMS' (case-insensitive). Also accepts fallback columns DEP STN/ARR STN."""
    cols = {c.lower(): c for c in df.columns}
    dep = cols.get("dep") or cols.get("dep stn") or None
    arr = cols.get("arr") or cols.get("arr stn") or None
    if not dep and not arr:
        return df  # can't filter; keep as-is
    mask = pd.Series(False, index=df.index)
    if dep:
        mask |= df[dep].astype(str).str.upper().str.strip().eq("AMS")
    if arr:
        mask |= df[arr].astype(str).str.upper().str.strip().eq("AMS")
    return df[mask]

def pick_september_years(df: pd.DataFrame) -> list[int]:
    """Return sorted list of years where POD month == 9."""
    if "_pod" not in df.columns:
        return []
    m = df["_pod"].dt.month == 9
    years = sorted(df.loc[m & df["_pod"].notna(), "_pod"].dt.year.dropna().astype(int).unique().tolist())
    return years

def kpis_for_september(df: pd.DataFrame, year: int | None) -> dict:
    """Compute Volume, Pieces, Gross OTP, Net OTP for September (given year)."""
    if df.empty or "_pod" not in df.columns:
        return {"volume": 0, "pieces": 0, "gross": np.nan, "net": np.nan}
    d = df.copy()
    d = d[d["_pod"].notna()]
    d = d[d["_pod"].dt.month == 9]
    if year is not None:
        d = d[d["_pod"].dt.year == year]
    volume = int(len(d))
    pieces = int(d["PIECES"].sum()) if "PIECES" in d.columns else 0
    # Need target for OTP
    base = d.dropna(subset=["_target"])
    if len(base) == 0:
        gross = np.nan
        net = np.nan
    else:
        gross = float(base["On_Time_Gross"].mean() * 100.0)
        net = float(base["On_Time_Net"].mean() * 100.0)
        if not np.isnan(gross) and not np.isnan(net) and net < gross:
            net = gross
    return {"volume": volume, "pieces": pieces, "gross": gross, "net": net}

def gauge(title: str, value: float | int) -> go.Figure:
    val = 0 if (value is None or (isinstance(value, float) and np.isnan(value))) else float(value)
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=val,
        number={"suffix": "%" if "OTP" in title else ""},
        title={"text": title},
        gauge={
            "axis": {"range": [0, 100] if "OTP" in title else [0, max(100, val*1.2 if val else 100)]},
        },
        domain={"x": [0, 1], "y": [0, 1]}
    ))
    fig.update_layout(height=220, margin=dict(l=10, r=10, t=40, b=10))
    return fig

def bar_top_accounts(df: pd.DataFrame, title: str, n: int = 10):
    if df.empty or "ACCT NM" not in df.columns or "_pod" not in df.columns:
        st.info("No account breakdown available.")
        return
    d = df.copy()
    d = d[d["_pod"].notna() & d["_target"].notna()]
    grp = d.groupby("ACCT NM").size().sort_values(ascending=False).head(n)
    if grp.empty:
        st.info("No account breakdown available.")
        return
    fig = go.Figure()
    fig.add_bar(x=grp.values, y=grp.index.tolist(), orientation="h")
    fig.update_layout(
        title=title,
        height=360,
        plot_bgcolor="white",
        margin=dict(l=10, r=10, t=40, b=40),
        xaxis_title="Volume (rows)",
        yaxis_title="Account"
    )
    st.plotly_chart(fig, use_container_width=True)

# ---------------- Sidebar: upload ----------------
with st.sidebar:
    st.markdown("### 📁 Data Upload")
    uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

    st.markdown("---")
    st.markdown("### ⚙️ Options")
    show_debug = st.checkbox("Show debug info", value=False)

if not uploaded:
    st.info("👆 Upload your Excel to get started.")
    st.stop()

# ---------------- Read Excel (lazily per sheet) ----------------
@st.cache_data(show_spinner=False)
def read_sheet(file, sheet_name: str) -> pd.DataFrame:
    try:
        return pd.read_excel(file, sheet_name=sheet_name, engine="openpyxl")
    except Exception:
        return pd.DataFrame()

@st.cache_data(show_spinner=False)
def read_many(file, sheets: list[str]) -> pd.DataFrame:
    frames = []
    for s in sheets:
        df = read_sheet(file, s)
        if not df.empty:
            df["Source_Sheet"] = s
            frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

# --- RP TAB (sheet 'AMS' + fixed accounts) ---
rp_raw = read_sheet(uploaded, "AMS")
if rp_raw.empty:
    st.warning("Sheet 'AMS' not found or empty.")
rp = rp_raw.copy()
if not rp.empty and "ACCT NM" in rp.columns:
    rp["ACCT NM"] = rp["ACCT NM"].astype(str).str.strip()
    rp = rp[rp["ACCT NM"].isin(RP_ACCOUNTS)]
rp = add_otp_columns(rp)

# --- LFS TAB (AMS + Americas International Desk) → AMS in DEP/ARR, Healthcare (exclude aviation) ---
lfs_raw = read_many(uploaded, ["AMS", "Americas International Desk"])
lfs = filter_ams_rows(lfs_raw)
if not lfs.empty and "ACCT NM" in lfs.columns:
    lfs["ACCT NM"] = lfs["ACCT NM"].astype(str)
    lfs = lfs[[not is_aviation(a, r.Source_Sheet) and is_healthcare(a, r.Source_Sheet) 
               for a, r in zip(lfs["ACCT NM"], lfs.itertuples())]]
lfs = add_otp_columns(lfs)

# --- AVS TAB (AMS + Americas International Desk + Aviation SVC) → AMS in DEP/ARR, Aviation (exclude healthcare) ---
avs_raw = read_many(uploaded, ["AMS", "Americas International Desk", "Aviation SVC"])
avs = filter_ams_rows(avs_raw)
if not avs.empty and "ACCT NM" in avs.columns:
    avs["ACCT NM"] = avs["ACCT NM"].astype(str)
    avs = avs[[is_aviation(a, r.Source_Sheet) and not is_healthcare(a, r.Source_Sheet) 
               for a, r in zip(avs["ACCT NM"], avs.itertuples())]]
avs = add_otp_columns(avs)

# Determine September years per tab
rp_years  = pick_september_years(rp)
lfs_years = pick_september_years(lfs)
avs_years = pick_september_years(avs)

# ---------------- Tabs ----------------
tab1, tab2, tab3 = st.tabs(["RP", "LFS", "AVS"])

with tab1:
    st.subheader("RP — Selected AMS Accounts")
    coly, = st.columns(1)
    with coly:
        sel_year = st.selectbox("Select year (September):", options=rp_years or [datetime.now().year], index=(len(rp_years)-1 if rp_years else 0), key="rp_year")
    rp_k = kpis_for_september(rp, sel_year)
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f'<div class="kpi"><div class="k-num">{rp_k["volume"]:,}</div><div class="k-cap">Volume (rows)</div></div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="kpi"><div class="k-num">{rp_k["pieces"]:,}</div><div class="k-cap">Pieces (sum)</div></div>', unsafe_allow_html=True)
    c3.plotly_chart(gauge("Gross OTP", rp_k["gross"]), use_container_width=True)
    c4.plotly_chart(gauge("Net OTP", rp_k["net"]), use_container_width=True)

    st.markdown("---")
    st.markdown("**Top accounts by volume (September)**")
    # Filter rp to the selected September for the breakdown
    if not rp.empty:
        rp_sep = rp[(rp["_pod"].dt.month == 9) & (rp["_pod"].dt.year == (sel_year if sel_year else datetime.now().year))]
        bar_top_accounts(rp_sep, "RP — Top Accounts by Volume (Sep)")

    if show_debug:
        st.caption("Debug: RP sample rows")
        st.dataframe(rp.head(10), use_container_width=True)

with tab2:
    st.subheader("LFS — Healthcare (AMS in DEP/ARR)")
    coly, = st.columns(1)
    with coly:
        sel_year = st.selectbox("Select year (September):", options=lfs_years or [datetime.now().year], index=(len(lfs_years)-1 if lfs_years else 0), key="lfs_year")
    k = kpis_for_september(lfs, sel_year)
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f'<div class="kpi"><div class="k-num">{k["volume"]:,}</div><div class="k-cap">Volume (rows)</div></div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="kpi"><div class="k-num">{k["pieces"]:,}</div><div class="k-cap">Pieces (sum)</div></div>', unsafe_allow_html=True)
    c3.plotly_chart(gauge("Gross OTP", k["gross"]), use_container_width=True)
    c4.plotly_chart(gauge("Net OTP", k["net"]), use_container_width=True)

    st.markdown("---")
    st.markdown("**Top accounts by volume (September)**")
    if not lfs.empty:
        df_sep = lfs[(lfs["_pod"].dt.month == 9) & (lfs["_pod"].dt.year == (sel_year if sel_year else datetime.now().year))]
        bar_top_accounts(df_sep, "LFS — Top Accounts by Volume (Sep)")

    if show_debug:
        st.caption("Debug: LFS sample rows")
        st.dataframe(lfs.head(10), use_container_width=True)

with tab3:
    st.subheader("AVS — Aviation (AMS in DEP/ARR)")
    coly, = st.columns(1)
    with coly:
        sel_year = st.selectbox("Select year (September):", options=avs_years or [datetime.now().year], index=(len(avs_years)-1 if avs_years else 0), key="avs_year")
    k = kpis_for_september(avs, sel_year)
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f'<div class="kpi"><div class="k-num">{k["volume"]:,}</div><div class="k-cap">Volume (rows)</div></div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="kpi"><div class="k-num">{k["pieces"]:,}</div><div class="k-cap">Pieces (sum)</div></div>', unsafe_allow_html=True)
    c3.plotly_chart(gauge("Gross OTP", k["gross"]), use_container_width=True)
    c4.plotly_chart(gauge("Net OTP", k["net"]), use_container_width=True)

    st.markdown("---")
    st.markdown("**Top accounts by volume (September)**")
    if not avs.empty:
        df_sep = avs[(avs["_pod"].dt.month == 9) & (avs["_pod"].dt.year == (sel_year if sel_year else datetime.now().year))]
        bar_top_accounts(df_sep, "AVS — Top Accounts by Volume (Sep)")

    if show_debug:
        st.caption("Debug: AVS sample rows")
        st.dataframe(avs.head(10), use_container_width=True)
