import re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from datetime import datetime
import plotly.express as px

# ---------------- Page & Style ----------------
st.set_page_config(page_title="September AMS Analysis", page_icon="📊",
                   layout="wide", initial_sidebar_state="collapsed")

# Define colors
NAVY  = "#0b1f44"
GOLD  = "#f0b429"
BLUE  = "#1f77b4"
GREEN = "#10b981"
SLATE = "#334155"
GRID  = "#e5e7eb"
RED   = "#dc2626"
EMERALD = "#059669"

st.markdown("""
<style>
.main {padding: 0rem 1rem;}
h1 {color:#0b1f44;font-weight:800;letter-spacing:.2px;border-bottom:3px solid #2ecc71;padding-bottom:10px;}
h2 {color:#0b1f44;font-weight:700;margin-top:1.2rem;margin-bottom:.6rem;}
.kpi {background:#fff;border:1px solid #e6e6e6;border-radius:14px;padding:14px;}
.k-num {font-size:36px;font-weight:800;color:#0b1f44;line-height:1.0;}
.k-cap {font-size:13px;color:#6b7280;margin-top:4px;}
.stTabs [data-baseweb="tab-list"] {gap: 8px;}
.stTabs [data-baseweb="tab"] {height: 50px;padding-left: 24px;padding-right: 24px;background-color: #f8f9fa;border-radius: 8px 8px 0 0;}
.stTabs [aria-selected="true"] {background-color: #0b1f44;color: white;}
</style>
""", unsafe_allow_html=True)

st.title("September AMS Analysis")

# ---------------- Config ----------------
OTP_TARGET = 95

# Radiopharma specific accounts
RP_ACCOUNTS = [
    'Marken Ltd',
    'QIAGEN GmbH Weekly',
    'Fisher Clinical Services',
    'Agilent Technologies Deutschland GmbH',
    'Patheon Biologics BV',
    'Delpharm Development Leiden BV',
    'Abbott Biologicals BV',
    'Fisher BioServices Netherlands BV',
    'Abbott Healthcare Products BV',
    'UNIVERSAL PICTURES INTERNATIONAL NETHERLANDS',
    'Patheon UK',
    'VERACYTE INC',
    'Tosoh Europe',
    'Exnet Services',
    'Nobel Biocare Distribution Center BV'
]

# Healthcare keywords for LFS identification
HEALTHCARE_KEYWORDS = [
    'pharma', 'medical', 'health', 'bio', 'clinical', 'hospital', 'diagnostic',
    'therapeut', 'laborator', 'patholog', 'imaging', 'surgical', 'oncolog',
    'cardio', 'neuro', 'radiol', 'genetic', 'genomic', 'molecular', 'cell',
    'tissue', 'organ', 'transplant', 'vaccine', 'antibod', 'protein', 'peptide',
    'life science', 'lifescience', 'medic', 'therap', 'diagnost', 'clinic',
    'patient', 'treatment', 'disease', 'drug', 'dose', 'isotope', 'radio',
    'nuclear', 'pet', 'spect', 'immuno', 'assay', 'reagent', 'specimen',
    'sample', 'blood', 'plasma', 'serum', 'biobank', 'cryo', 'stem',
    'marken', 'fisher', 'cardinal', 'patheon', 'organox', 'qiagen', 'abbott',
    'tosoh', 'leica', 'sophia', 'cerus', 'sirtex', 'lantheus', 'avid',
    'petnet', 'innervate', 'ndri', 'university', 'institut', 'pentec',
    'sexton', 'atomics', 'curium', 'medtronic', 'catalent', 'delpharm',
    'veracyte', 'eckert', 'ziegler', 'shine', 'altasciences',
    'onkos', 'biolabs', 'biosystem', 'life molecular', 'cerveau', 'meilleur',
    'samsung bio', 'agilent'
]

# Aviation keywords for AVS identification
AVIATION_KEYWORDS = [
    'airline', 'airport', 'cargo', 'freight', 'logistic', 'transport',
    'express', 'disney', 'pictures', 'aviation', 'aircraft', 'aerospace',
    'volaris', 'easyjet', 'lufthansa', 'delta', 'american airlines',
    'british airways', 'nippon', 'aeromexico', 'spairliners', 'universal',
    'paramount', 'productions', 'courier', 'forwarding', 'tmr global',
    'aeroplex', 'nova traffic', 'ups', 'endeavor air',
    'storm aviation', 'adventures', 'hartford', 'tokyo electron', 'slipstick',
    'sealion production', 'heathrow courier', 'macaronesia', 'exnet service',
    'mnx global logistics', 'logical freight', 'concesionaria', 'vuela compania',
    'panasonic avionics'
]

# ---------------- Helper Functions ----------------
def _excel_to_dt(s: pd.Series) -> pd.Series:
    """Robust datetime: parse; if many NaT, try Excel serials."""
    out = pd.to_datetime(s, errors="coerce")
    if out.isna().mean() > 0.5:
        num  = pd.to_numeric(s, errors="coerce")
        out2 = pd.to_datetime("1899-12-30") + pd.to_timedelta(num, unit="D")
        out  = out.where(~out.isna(), out2)
    return out

def _get_target_series(df: pd.DataFrame) -> pd.Series | None:
    if "UPD DEL" in df.columns and df["UPD DEL"].notna().any():
        return df["UPD DEL"]
    if "QDT" in df.columns:
        return df["QDT"]
    return None

def is_healthcare(account_name):
    """Determine if an account is healthcare-related."""
    if not account_name:
        return False
    account_lower = str(account_name).lower()
   
    # Check if any aviation keyword is present (exclude these from healthcare)
    for keyword in AVIATION_KEYWORDS:
        if keyword in account_lower:
            return False
   
    # Check if any healthcare keyword is present
    for keyword in HEALTHCARE_KEYWORDS:
        if keyword in account_lower:
            return True
   
    return False

def is_aviation(account_name):
    """Determine if an account is aviation-related."""
    if not account_name:
        return False
    account_lower = str(account_name).lower()
   
    # Check if any healthcare keyword is present (exclude these from aviation)
    for keyword in HEALTHCARE_KEYWORDS:
        if keyword in account_lower:
            return False
   
    # Check if any aviation keyword is present
    for keyword in AVIATION_KEYWORDS:
        if keyword in account_lower:
            return True
   
    return False

def calculate_otp(df: pd.DataFrame) -> tuple:
    """Calculate Gross and Net OTP percentages.
    Gross OTP: All QC (Quality Control) shipments
    Net OTP: Only controllable shipments (excludes certain QC codes)
    """
    if df.empty:
        return 0.0, 0.0
   
    # Get the target series (UPD DEL or QDT)
    target_col = _get_target_series(df)
    if target_col is None or "POD DATE/TIME" not in df.columns:
        return 0.0, 0.0
   
    # Parse dates
    target_dt = _excel_to_dt(target_col)
    pod_dt = _excel_to_dt(df["POD DATE/TIME"])
   
    # Calculate valid mask (has both target and pod dates)
    valid_mask = target_dt.notna() & pod_dt.notna()
    if not valid_mask.any():
        return 0.0, 0.0
   
    # On-time mask: delivered on or before target date
    ontime_mask = (pod_dt <= target_dt) & valid_mask
   
    # Gross OTP: All shipments with valid dates (all QC)
    gross_otp = (ontime_mask.sum() / valid_mask.sum() * 100) if valid_mask.sum() > 0 else 0.0
   
    # Net OTP: Only controllable shipments
    # Check if QC column exists for filtering
    if 'QC' in df.columns:
        # Controllable: exclude specific QC codes that are not controllable
        # Based on reference code, controllables exclude certain QC patterns
        qc_col = df['QC'].astype(str).str.upper()
       
        # Non-controllable patterns (agent, customs, warehouse issues)
        ctrl_regex = re.compile(r"\b(AGENT|DEL\s*AGT|DELIVERY\s*AGENT|CUSTOMS|WAREHOUSE|W/HOUSE)\b", re.I)
        non_controllable = qc_col.apply(lambda x: bool(ctrl_regex.search(str(x))) if pd.notna(x) else False)
       
        # Controllable mask: valid dates AND not in non-controllable QC codes
        controllable_mask = valid_mask & ~non_controllable
       
        if controllable_mask.sum() > 0:
            net_otp = ((pod_dt <= target_dt) & controllable_mask).sum() / controllable_mask.sum() * 100
        else:
            net_otp = gross_otp  # Fallback to gross if no controllable distinction
    else:
        # If no QC column, Net OTP = Gross OTP
        net_otp = gross_otp
   
    return gross_otp, net_otp

def filter_by_ams(df: pd.DataFrame) -> pd.DataFrame:
    """Filter rows where DEP or ARR contains 'AMS'."""
    if df.empty:
        return df
   
    mask = pd.Series([False] * len(df), index=df.index)
   
    if 'DEP' in df.columns:
        mask |= df['DEP'].astype(str).str.contains('AMS', case=False, na=False)
   
    if 'ARR' in df.columns:
        mask |= df['ARR'].astype(str).str.contains('AMS', case=False, na=False)
   
    return df[mask]

def create_metrics_display(volume, pieces, gross_otp, net_otp, revenue, title):
    """Create a metrics display for a tab."""
    st.markdown(f"### {title} Metrics")
   
    col1, col2, col3, col4, col5 = st.columns(5)
   
    with col1:
        st.markdown(f"""
        <div class="kpi">
            <div class="k-num">{volume:,}</div>
            <div class="k-cap">Volume (Shipments)</div>
        </div>
        """, unsafe_allow_html=True)
   
    with col2:
        st.markdown(f"""
        <div class="kpi">
            <div class="k-num">{pieces:,}</div>
            <div class="k-cap">Total Pieces</div>
        </div>
        """, unsafe_allow_html=True)
   
    with col3:
        st.markdown(f"""
        <div class="kpi">
            <div class="k-num">${revenue:,.2f}</div>
            <div class="k-cap">Total Revenue</div>
        </div>
        """, unsafe_allow_html=True)
   
    with col4:
        color = GREEN if gross_otp >= OTP_TARGET else RED
        st.markdown(f"""
        <div class="kpi">
            <div class="k-num" style="color:{color}">{gross_otp:.1f}%</div>
            <div class="k-cap">Gross OTP (All QC)</div>
        </div>
        """, unsafe_allow_html=True)
   
    with col5:
        color = GREEN if net_otp >= OTP_TARGET else RED
        st.markdown(f"""
        <div class="kpi">
            <div class="k-num" style="color:{color}">{net_otp:.1f}%</div>
            <div class="k-cap">Net OTP (Controllable)</div>
        </div>
        """, unsafe_allow_html=True)

def create_top10_charts(df: pd.DataFrame, title_prefix: str):
    """Create top 10 charts for accounts by volume and revenue."""
    if df.empty or 'ACCT NM' not in df.columns:
        st.warning("No account data available for charts.")
        return
   
    col1, col2 = st.columns(2)
   
    with col1:
        st.markdown(f"#### Top 10 Accounts by Volume")
        # Count orders per account
        volume_by_account = df['ACCT NM'].value_counts().head(10)
       
        if not volume_by_account.empty:
            fig_volume = go.Figure(data=[
                go.Bar(
                    x=volume_by_account.values,
                    y=volume_by_account.index,
                    orientation='h',
                    marker=dict(color=NAVY),
                    text=volume_by_account.values,
                    textposition='auto',
                )
            ])
            fig_volume.update_layout(
                xaxis_title="Number of Orders",
                yaxis_title="Account",
                height=400,
                margin=dict(l=20, r=20, t=40, b=20),
                yaxis={'categoryorder': 'total ascending'}
            )
            st.plotly_chart(fig_volume, use_container_width=True)
        else:
            st.info("No volume data available.")
   
    with col2:
        st.markdown(f"#### Top 10 Accounts by Revenue")
        # Sum revenue per account
        if 'TOTAL CHARGES' in df.columns:
            revenue_by_account = df.groupby('ACCT NM')['TOTAL CHARGES'].sum().sort_values(ascending=False).head(10)
           
            if not revenue_by_account.empty:
                fig_revenue = go.Figure(data=[
                    go.Bar(
                        x=revenue_by_account.values,
                        y=revenue_by_account.index,
                        orientation='h',
                        marker=dict(color=GOLD),
                        text=['${:,.0f}'.format(x) for x in revenue_by_account.values],
                        textposition='auto',
                    )
                ])
                fig_revenue.update_layout(
                    xaxis_title="Total Revenue ($)",
                    yaxis_title="Account",
                    height=400,
                    margin=dict(l=20, r=20, t=40, b=20),
                    yaxis={'categoryorder': 'total ascending'}
                )
                st.plotly_chart(fig_revenue, use_container_width=True)
            else:
                st.info("No revenue data available.")
        else:
            st.info("TOTAL CHARGES column not found in data.")

# ---------------- File Upload ----------------
uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # Read all sheets
        excel_file = pd.ExcelFile(uploaded_file)
        available_sheets = excel_file.sheet_names
       
        st.success(f"✅ Excel file loaded successfully! Found {len(available_sheets)} sheets.")
       
        # ---------------- TAB 1: RP (Radiopharma) ----------------
        # Read only AMS sheet
        if 'AMS' in available_sheets:
            df_ams = pd.read_excel(uploaded_file, sheet_name='AMS')
           
            # Filter for RP accounts
            rp_df = df_ams[df_ams['ACCT NM'].isin(RP_ACCOUNTS)]
           
            # Calculate metrics
            rp_volume = len(rp_df)
            rp_pieces = rp_df['PIECES'].sum() if 'PIECES' in rp_df.columns else 0
            rp_revenue = rp_df['TOTAL CHARGES'].sum() if 'TOTAL CHARGES' in rp_df.columns else 0.0
            rp_gross_otp, rp_net_otp = calculate_otp(rp_df)
        else:
            rp_df = pd.DataFrame()
            rp_volume = 0
            rp_pieces = 0
            rp_revenue = 0.0
            rp_gross_otp = 0.0
            rp_net_otp = 0.0
       
        # ---------------- TAB 2: LFS (Life Sciences) ----------------
        # Read AMS and Americas International Desk sheets
        lfs_dfs = []
        for sheet in ['AMS', 'Americas International Desk']:
            if sheet in available_sheets:
                df = pd.read_excel(uploaded_file, sheet_name=sheet)
                df = filter_by_ams(df)  # Filter by AMS office
                lfs_dfs.append(df)
       
        if lfs_dfs:
            lfs_combined = pd.concat(lfs_dfs, ignore_index=True)
            # Filter for healthcare accounts
            if 'ACCT NM' in lfs_combined.columns:
                lfs_combined['is_healthcare'] = lfs_combined['ACCT NM'].apply(is_healthcare)
                lfs_df = lfs_combined[lfs_combined['is_healthcare'] == True]
            else:
                lfs_df = pd.DataFrame()
           
            # Calculate metrics
            lfs_volume = len(lfs_df)
            lfs_pieces = lfs_df['PIECES'].sum() if 'PIECES' in lfs_df.columns else 0
            lfs_revenue = lfs_df['TOTAL CHARGES'].sum() if 'TOTAL CHARGES' in lfs_df.columns else 0.0
            lfs_gross_otp, lfs_net_otp = calculate_otp(lfs_df)
        else:
            lfs_df = pd.DataFrame()
            lfs_volume = 0
            lfs_pieces = 0
            lfs_revenue = 0.0
            lfs_gross_otp = 0.0
            lfs_net_otp = 0.0
       
        # ---------------- TAB 3: AVS (Aviation Services) ----------------
        # Read AMS, Americas International Desk, and Aviation SVC sheets
        avs_dfs = []
        for sheet in ['AMS', 'Americas International Desk', 'Aviation SVC']:
            if sheet in available_sheets:
                df = pd.read_excel(uploaded_file, sheet_name=sheet)
                df = filter_by_ams(df)  # Filter by AMS office
                avs_dfs.append(df)
       
        if avs_dfs:
            avs_combined = pd.concat(avs_dfs, ignore_index=True)
            # Filter for aviation accounts
            if 'ACCT NM' in avs_combined.columns:
                avs_combined['is_aviation'] = avs_combined['ACCT NM'].apply(is_aviation)
                avs_df = avs_combined[avs_combined['is_aviation'] == True]
            else:
                avs_df = pd.DataFrame()
           
            # Calculate metrics
            avs_volume = len(avs_df)
            avs_pieces = avs_df['PIECES'].sum() if 'PIECES' in avs_df.columns else 0
            avs_revenue = avs_df['TOTAL CHARGES'].sum() if 'TOTAL CHARGES' in avs_df.columns else 0.0
            avs_gross_otp, avs_net_otp = calculate_otp(avs_df)
        else:
            avs_df = pd.DataFrame()
            avs_volume = 0
            avs_pieces = 0
            avs_revenue = 0.0
            avs_gross_otp = 0.0
            avs_net_otp = 0.0
       
        # ---------------- Create Tabs ----------------
        tab1, tab2, tab3 = st.tabs(["📦 RP", "🏥 LFS", "✈️ AVS"])
       
        with tab1:
            st.markdown("## Radiopharma (RP)")
            if not rp_df.empty:
                create_metrics_display(rp_volume, rp_pieces, rp_gross_otp, rp_net_otp, rp_revenue, "September RP")
               
                st.markdown("---")
                create_top10_charts(rp_df, "RP")
               
                with st.expander("📊 View Data Details"):
                    st.markdown(f"**Total Rows:** {len(rp_df):,}")
                    st.markdown(f"**Unique Accounts:** {rp_df['ACCT NM'].nunique()}")
                    st.dataframe(rp_df.head(50))
            else:
                st.warning("⚠️ No RP data found. Please check if the AMS sheet exists and contains the specified accounts.")
       
        with tab2:
            st.markdown("## Life Sciences (LFS)")
            if not lfs_df.empty:
                create_metrics_display(lfs_volume, lfs_pieces, lfs_gross_otp, lfs_net_otp, lfs_revenue, "September LFS")
               
                st.markdown("---")
                create_top10_charts(lfs_df, "LFS")
               
                with st.expander("📊 View Data Details"):
                    st.markdown(f"**Total Rows:** {len(lfs_df):,}")
                    st.markdown(f"**Unique Accounts:** {lfs_df['ACCT NM'].nunique() if 'ACCT NM' in lfs_df.columns else 0}")
                    st.dataframe(lfs_df.head(50))
            else:
                st.warning("⚠️ No LFS data found. Please check if the required sheets exist and contain AMS office data with healthcare accounts.")
       
        with tab3:
            st.markdown("## Aviation Services (AVS)")
            if not avs_df.empty:
                create_metrics_display(avs_volume, avs_pieces, avs_gross_otp, avs_net_otp, avs_revenue, "September AVS")
               
                st.markdown("---")
                create_top10_charts(avs_df, "AVS")
               
                with st.expander("📊 View Data Details"):
                    st.markdown(f"**Total Rows:** {len(avs_df):,}")
                    st.markdown(f"**Unique Accounts:** {avs_df['ACCT NM'].nunique() if 'ACCT NM' in avs_df.columns else 0}")
                    st.dataframe(avs_df.head(50))
            else:
                st.warning("⚠️ No AVS data found. Please check if the required sheets exist and contain AMS office data with aviation accounts.")
       
        # Summary statistics
        with st.expander("📈 Overall Summary"):
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown("#### RP (Radiopharma)")
                st.metric("Volume", f"{rp_volume:,}")
                st.metric("Pieces", f"{rp_pieces:,}")
                st.metric("Revenue", f"${rp_revenue:,.2f}")
            with col2:
                st.markdown("#### LFS (Life Sciences)")
                st.metric("Volume", f"{lfs_volume:,}")
                st.metric("Pieces", f"{lfs_pieces:,}")
                st.metric("Revenue", f"${lfs_revenue:,.2f}")
            with col3:
                st.markdown("#### AVS (Aviation)")
                st.metric("Volume", f"{avs_volume:,}")
                st.metric("Pieces", f"{avs_pieces:,}")
                st.metric("Revenue", f"${avs_revenue:,.2f}")
           
            st.markdown("---")
            total_volume = rp_volume + lfs_volume + avs_volume
            total_pieces = rp_pieces + lfs_pieces + avs_pieces
            total_revenue = rp_revenue + lfs_revenue + avs_revenue
           
            col_t1, col_t2, col_t3 = st.columns(3)
            with col_t1:
                st.markdown(f"**Total Volume (All):** {total_volume:,}")
            with col_t2:
                st.markdown(f"**Total Pieces (All):** {total_pieces:,}")
            with col_t3:
                st.markdown(f"**Total Revenue (All):** ${total_revenue:,.2f}")
   
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        st.exception(e)

else:
    st.info("👆 Please upload an Excel file to begin analysis.")
    st.markdown("""
    ### Expected File Structure:
    - **Sheet: AMS** - Contains radiopharma accounts and general data
    - **Sheet: Americas International Desk** - Additional data for LFS and AVS
    - **Sheet: Aviation SVC** - Aviation-specific data for AVS
   
    ### Required Columns:
    - `ACCT NM` - Account Name
    - `PIECES` - Number of pieces
    - `DEP` - Departure location
    - `ARR` - Arrival location
    - `POD DATE/TIME` - Proof of delivery date/time
    - `UPD DEL` or `QDT` - Target delivery date
    """)
