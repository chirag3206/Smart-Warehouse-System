import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import math
import gspread
from google.oauth2 import service_account
# Add import for autorefresh
from streamlit_autorefresh import st_autorefresh
import hashlib
import time
import sys

# Helper for rerun (Streamlit >=1.18: st.rerun, else st.experimental_rerun)
def rerun_app():
    if hasattr(st, 'rerun'):
        st.rerun()
    elif hasattr(st, 'experimental_rerun'):
        st.experimental_rerun()
    else:
        st.warning('Streamlit rerun is not available. Please upgrade Streamlit.')

def create_device_box(device):
    status_class = "live" if device['Initial Status'] == 'Live' else "offline"
    html = f"""
        <div class="device-box">
            <div class="status-dot {status_class}"></div>
            {device['Model']}
            <div class="popup">
                <strong>Type:</strong> {device['Types']}<br>
                <strong>Status:</strong> {device['Initial Status']}<br>
                <strong>Location:</strong> {device['Camera & NVR(1F or HO)']}<br>
                <strong>PO Date:</strong> {device.get('PO Date', 'N/A')}<br>
                <strong>IP Address:</strong> {device['Camera or NVR IP']}<br>
            </div>
        </div>
    """
    return html

# Set page config
st.set_page_config(
    page_title="Inventory Management System",
    page_icon="🏍️",
    layout="wide"
)

# Add auto-refresh at the very top of the app (5 seconds = 5000 ms)
st_autorefresh(interval=5000, limit=None, key="autorefresh2s")

# Simple header row: Title and emoji-based refresh button, right-aligned
st.title("🏍️ Inventory Management System")

# Load and apply CSS
with open('styles.css') as f:
    css = f.read()
    
# Create API client
credentials = service_account.Credentials.from_service_account_file(
    'inventory-managment-465211-7ba8ecdf5815.json',
    scopes=['https://www.googleapis.com/auth/spreadsheets']
)
client = gspread.authorize(credentials)

# Custom CSS for styled boxes and overlay popups
st.markdown("""
<style>
    /* Increase font sizes for all text elements */
    .stTitle {
        font-size: 5.5rem !important;
    }
    
    .stSubheader {
        font-size: 2.5rem !important;
    }
    
    .stMarkdown {
        font-size: 1.5rem !important;
    }
    
    .stRadio > label {
        font-size: 1.5rem !important;
    }
    
    .stSelectbox > label {
        font-size: 1.5rem !important;
    }
    
    .stButton > button {
        font-size: 1.5rem !important;
    }
    
    .stCheckbox > label {
        font-size: 1.5rem !important;
    }
    
    .stMetric {
        font-size: 1.5rem !important;
    }
    
    .stExpander > summary {
        font-size: 1.5rem !important;
    }
    
    /* Device box styling */
    .device-box {
        background-color: #1e3d59;
        color: white;
        padding: 8px;
        border-radius: 4px;
        cursor: pointer;
        margin: 4px;
        min-height: 60px;
        position: relative;
        font-size: 1.4rem;
        text-align: center;
        width: 100%;
        display: inline-block;
    }
    
    /* Status indicator dot */
    .status-dot {
        position: absolute;
        top: 5px;
        right: 5px;
        width: 10px;
        height: 10px;
        border-radius: 50%;
    }
    
    .status-dot.live {
        background-color: #2ecc71;
        box-shadow: 0 0 5px #2ecc71;
    }
    
    .status-dot.offline {
        background-color: #e74c3c;
        box-shadow: 0 0 5px #e74c3c;
    }
    
    /* Popup styling */
    .popup {
        display: none;
        position: absolute;
        background-color: white;
        border: 1px solid #ddd;
        border-radius: 4px;
        padding: 12px;
        z-index: 1000;
        width: 250px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        color: black;
        text-align: left;
        left: 50%;
        transform: translateX(-50%);
    }
    
    .device-box:hover .popup {
        display: block;
    }
</style>
""", unsafe_allow_html=True)

data_source = st.radio("Select Data Source", ["Google Sheet", "Upload Excel File"], horizontal=True)

# --- Smart Google Sheet Refresh with Caching ---
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1r55Y83e4LV-dN00b2u5dFPnZ9unehUyK4K9d7iANEYo'
SHEET_IDX = 0
@st.cache_resource(ttl=60)
def fetch_gsheet_data():
    sheet = client.open_by_url(SHEET_URL)
    worksheet = sheet.get_worksheet(SHEET_IDX)
    data = worksheet.get_all_records()
    df_gsheet = pd.DataFrame(data)
    return df_gsheet

# Use session state for main DataFrame
df = None
if data_source == "Google Sheet":
    if 'df' not in st.session_state:
        st.session_state.df = fetch_gsheet_data()
    if st.button('🔄 Refresh', key="refresh_btn"):
        st.cache_resource.clear()
        st.session_state.df = fetch_gsheet_data()
    df = st.session_state.df
else:
    # Excel upload logic as before
    uploaded_file = st.file_uploader("\U0001F4C2 Upload Inventory Excel File", type=["xlsx"])
    # Removed st_autorefresh for Excel uploads
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
            if 'PO Date' in df.columns:
                df['Age (Years)'] = pd.to_datetime(df['PO Date'], errors='coerce', dayfirst=True).apply(
                    lambda x: (pd.Timestamp.now() - x).days / 365.25 if pd.notna(x) else None
                )
            else:
                st.warning("'PO Date' column not found in your Excel file. Age calculation will be skipped.")
                df['Age (Years)'] = None
        except Exception as e:
            st.error(f"Error reading the Excel file: {str(e)}")
            st.stop()
    else:
        st.warning("Please upload an Excel file.")
        st.stop()

if df is None or df.empty:
    st.error("No data available.")
    st.stop()

# --- Firmware Update Alert Section ---
if "Firmware available or not" in df.columns:
    # Exclude rows where value is 'No more updates' or 'OK' (case-insensitive, strip spaces)
    firmware_mask = ~df["Firmware available or not"].astype(str).str.strip().str.lower().isin(["no more updates", "ok"])
    firmware_update_df = df[firmware_mask]
    firmware_update_count = len(firmware_update_df)
    if firmware_update_count > 0:
        st.markdown(
            f"<div style='font-size:2.0rem; font-weight:bold; margin-bottom: 0.5em; color:#f39c12;'>"
            f"🔔 Firmware Update: {firmware_update_count}</div>",
            unsafe_allow_html=True
        )
        # Group by location
        plant_office_count = firmware_update_df[firmware_update_df['Camera & NVR(1F or HO)'].astype(str).str.strip().str.upper() == '1F'].shape[0]
        ho_count = firmware_update_df[firmware_update_df['Camera & NVR(1F or HO)'].astype(str).str.strip().str.upper() == 'HO'].shape[0]
        st.markdown(f"""
        • 🏭 **Plant(1F)**: `{plant_office_count}` devices  
        • 🏢 **HO**: `{ho_count}` devices
        """)
        # Firmware Update Expander
        if firmware_update_count > 0:
            with st.expander("🔧 View Devices Needing Firmware Update"):
                firmware_cols = [c for c in ['Area', 'Types', 'Model', 'Camera & NVR(1F or HO)', 'Firmware available or not', 'Initial Status'] if c in firmware_update_df.columns]
                st.dataframe(
                    firmware_update_df[firmware_cols],
                    column_config={
                        'Area': 'Area',
                        'Types': 'Types',
                        'Model': 'Model',
                        'Camera & NVR(1F or HO)': 'Location',
                        'Firmware available or not': 'Firmware Status',
                        'Initial Status': 'Initial Status',
                    },
                    hide_index=True
                )
        else:
            st.success("✅ All devices have updated firmware.")
else:
    st.warning("⚠️ 'Firmware available or not' column not found in your data.")

# --- Devices Requiring Repair Section (all devices, not filtered) ---
repair_devices_df = df[df['Initial Status'].astype(str).str.strip().str.lower() == 'repair']
repair_count = len(repair_devices_df)
st.markdown(
    f"<div style='font-size:2.0rem; font-weight:bold; margin-bottom: 0.5em; color:#e67e22;'>"
    f"⚠️🛠 Repair: {repair_count}</div>",
    unsafe_allow_html=True
)
# Plant/HO breakdown for Repair
plant_repair_count = repair_devices_df[repair_devices_df['Camera & NVR(1F or HO)'].astype(str).str.strip().str.upper() == '1F'].shape[0]
ho_repair_count = repair_devices_df[repair_devices_df['Camera & NVR(1F or HO)'].astype(str).str.strip().str.upper() == 'HO'].shape[0]
st.markdown(f"""
• 🏭 **Plant (1F)**: `{plant_repair_count}` devices  
• 🏢 **HO**: `{ho_repair_count}` devices
""")
if repair_count > 0:
    with st.expander("Show Repair Device Details"):
        repair_cols = [c for c in ['Area', 'Types', 'Model', 'PO Date', 'Device Age (Years)', 'Initial Status'] if c in repair_devices_df.columns]
        st.dataframe(
            repair_devices_df[repair_cols],
            column_config={
                'Area': 'Area',
                'Types': 'Types',
                'Model': 'Model',
                'PO Date': 'PO Date',
                'Device Age (Years)': st.column_config.NumberColumn(
                    'Device Age (Years)', format='%.1f'),
                'Initial Status': 'Initial Status',
            },
            hide_index=True
        )

# --- Not in Use (Discard) Section (all devices, not filtered) ---
stock_devices_df = df[df['Initial Status'].astype(str).str.strip() == 'Discard']
stock_count = len(stock_devices_df)
st.markdown(
    f"<div style='font-size:2.0rem; font-weight:bold; margin-bottom: 0.5em; color:#888;'>"
    f"❌ Not in Use: {stock_count}</div>",
    unsafe_allow_html=True
)
# Plant/HO breakdown for Not in Use
plant_down_count = stock_devices_df[stock_devices_df['Camera & NVR(1F or HO)'].astype(str).str.strip().str.upper() == '1F'].shape[0]
ho_down_count = stock_devices_df[stock_devices_df['Camera & NVR(1F or HO)'].astype(str).str.strip().str.upper() == 'HO'].shape[0]
st.markdown(f"""
• 🏭 **Plant (1F)**: `{plant_down_count}` devices  
• 🏢 **HO**: `{ho_down_count}` devices
""")
if stock_count > 0:
    with st.expander("Show Not in Use (Discard) Device Details"):
        down_cols = [c for c in ['Area', 'Types', 'Model', 'PO Date', 'Device Age (Years)', 'Initial Status'] if c in stock_devices_df.columns]
        st.dataframe(
            stock_devices_df[down_cols],
            column_config={
                'Area': 'Area',
                'Types': 'Types',
                'Model': 'Model',
                'PO Date': 'PO Date',
                'Device Age (Years)': st.column_config.NumberColumn(
                    'Device Age (Years)', format='%.1f'),
                'Initial Status': 'Initial Status',
            },
            hide_index=True
        )

# --- PO Date Age Alerts Section ---
# Safe date conversion
po_dates = pd.to_datetime(df['PO Date'], errors='coerce', dayfirst=True)
now = pd.Timestamp.now()
device_ages = (now - po_dates).dt.total_seconds() / (365.25 * 24 * 60 * 60)

# Add device age to df for filtering
age_df = df.copy()
age_df['Device Age (Years)'] = device_ages

# High Alert: Devices older than 6 years OR within 1 month (1/12 year) of crossing 6 years
high_alert_mask = age_df['Device Age (Years)'] > (6 - 1/12)
high_alert_df = age_df[high_alert_mask].copy()

# Mild Alert: Devices within 6 months (0.5 year) of crossing 6 years (but not in High Alert)
mild_alert_mask = (age_df['Device Age (Years)'] > (6 - 0.5)) & (age_df['Device Age (Years)'] <= (6 - 1/12))
mild_alert_df = age_df[mild_alert_mask].copy()

# Format columns
for alert_df in [high_alert_df, mild_alert_df]:
    alert_df = alert_df.copy()
    alert_df.loc[:, 'Device Age (Years)'] = alert_df['Device Age (Years)'].round(1)
    alert_df.loc[:, 'PO Date'] = pd.to_datetime(alert_df['PO Date'], errors='coerce', dayfirst=True).dt.strftime('%Y-%m-%d')

# High Alert Banner (red if count > 0, green if 0)
high_alert_count = len(high_alert_df)
if high_alert_count > 0:
    high_alert_color = '#d32f2f'
    high_alert_icon = '🔴'
else:
    high_alert_color = '#27ae60'
    high_alert_icon = '🟢'
st.markdown(
    f"<div style='font-size:2.0rem; font-weight:bold; margin-bottom: 0.5em; color:{high_alert_color};'>"
    f"{high_alert_icon} High Alert: {high_alert_count}</div>",
    unsafe_allow_html=True
)
# Plant/HO breakdown for High Alert
plant_high_count = high_alert_df[high_alert_df['Camera & NVR(1F or HO)'].astype(str).str.strip().str.upper() == '1F'].shape[0]
ho_high_count = high_alert_df[high_alert_df['Camera & NVR(1F or HO)'].astype(str).str.strip().str.upper() == 'HO'].shape[0]
st.markdown(f"""
• 🏭 **Plant (1F)**: `{plant_high_count}` devices  
• 🏢 **HO**: `{ho_high_count}` devices
""")
with st.expander("Show High Alert Device Details"):
    st.markdown("""
    <div style='font-size:1.1rem; margin-bottom: 1em;'>Devices with PO Date more than 6 years ago or within 1 month of crossing 6 years. These are considered high risk for replacement or maintenance.</div>
    """, unsafe_allow_html=True)
    if not high_alert_df.empty:
        high_cols = [c for c in ['Area', 'Types', 'Model', 'PO Date', 'Device Age (Years)', 'Initial Status'] if c in high_alert_df.columns]
        st.dataframe(
            high_alert_df[high_cols],
            column_config={
                'Area': 'Area',
                'Types': 'Types',
                'Model': 'Model',
                'PO Date': 'PO Date',
                'Device Age (Years)': st.column_config.NumberColumn(
                    'Device Age (Years)', format='%.1f'),
                'Initial Status': 'Initial Status',
            },
            hide_index=True
        )
    else:
        st.success("No devices in High Alert category.")

# Mild Alert Banner
mild_alert_count = len(mild_alert_df)
st.markdown(
    f"<div style='font-size:2.0rem; font-weight:bold; margin-bottom: 0.5em; color:#fbc02d;'>"
    f"🟡 Mild Alert: {mild_alert_count}</div>",
    unsafe_allow_html=True
)
# Plant/HO breakdown for Mild Alert
plant_mild_count = mild_alert_df[mild_alert_df['Camera & NVR(1F or HO)'].astype(str).str.strip().str.upper() == '1F'].shape[0]
ho_mild_count = mild_alert_df[mild_alert_df['Camera & NVR(1F or HO)'].astype(str).str.strip().str.upper() == 'HO'].shape[0]
st.markdown(f"""
• 🏭 **Plant (1F)**: `{plant_mild_count}` devices  
• 🏢 **HO**: `{ho_mild_count}` devices
""")
with st.expander("Show Mild Alert Device Details"):
    st.markdown("""
    <div style='font-size:1.1rem; margin-bottom: 1em;'>Devices with PO Date with 6 months left for threshold of 6 years. These are within 6 months of the 6-year threshold.</div>
    """, unsafe_allow_html=True)
    if not mild_alert_df.empty:
        mild_cols = [c for c in ['Area', 'Types', 'Model', 'PO Date', 'Device Age (Years)', 'Initial Status'] if c in mild_alert_df.columns]
        st.dataframe(
            mild_alert_df[mild_cols],
            column_config={
                'Area': 'Area',
                'Types': 'Types',
                'Model': 'Model',
                'PO Date': 'PO Date',
                'Device Age (Years)': st.column_config.NumberColumn(
                    'Device Age (Years)', format='%.1f'),
                'Initial Status': 'Initial Status',
            },
            hide_index=True
        )
    else:
        st.success("No devices in Mild Alert category.")

# --- Non-active device breakdown for entire inventory (accurate, case-sensitive) ---
total_devices_all = len(df)
stock_count_all = (df['Initial Status'].str.strip() == 'Discard').sum()
repair_count_all = (df['Initial Status'].str.strip() == 'Repair').sum()
stock_pct_all = (stock_count_all / total_devices_all) * 100 if total_devices_all > 0 else 0
repair_pct_all = (repair_count_all / total_devices_all) * 100 if total_devices_all > 0 else 0
st.markdown(
    f"<div style='font-size:1.6rem; margin-top: 8px; margin-bottom: 8px;'>"
    f"<b>Non-active devices in entire inventory:</b> "
    f"<span style='color:#888;'>❌ Not in Use: {stock_pct_all:.1f}%</span> &nbsp; "
    f"<span style='color:#e67e22;'>🛠 Repair: {repair_pct_all:.1f}%</span></div>",
    unsafe_allow_html=True
)

# Add extra space after the firmware/repair/down/non-active group
st.markdown("<br><br><br>", unsafe_allow_html=True)

# --- Location Filter Group ---
st.markdown("<div style='font-size:2.0rem; font-weight:bold; margin-bottom: 0.5em;'>Location Filter</div>", unsafe_allow_html=True)
# Define location mapping
location_mapping = {
    'Plant': ['1F'],
    'HO': ['HO'],
}
# Get main locations for radio
main_locations = list(location_mapping.keys())
# Create main location filter as radio (not dropdown)
selected_main_location = st.radio(
    "Select Main Location",
    options=main_locations,
    index=0,
    horizontal=True
)

# Filter data for the selected main location
sub_locations = location_mapping[selected_main_location]
filtered_df = df[df['Camera & NVR(1F or HO)'].isin(sub_locations)]

# If Plant is selected, add area location filter (use 'Area' column instead of 'Camera name')
if selected_main_location == 'Plant':
    area_locations = sorted(filtered_df['Area'].unique())
    selected_area_location = st.selectbox(
        "Select Area",
        options=area_locations,
        index=0
    )
    filtered_df = filtered_df[filtered_df['Area'] == selected_area_location]

if not filtered_df.empty:
    # --- Quick Stats Group ---
    st.markdown("### Quick Stats")
    
    # Calculate KPIs
    total_devices = len(filtered_df)
    active_devices = len(filtered_df[filtered_df['Initial Status'] == 'Live'])
    active_percentage = (active_devices / total_devices * 100) if total_devices > 0 else 0
    
    # 2. Warranty/AMC Coverage
    covered_devices = len(filtered_df[
        filtered_df['AMC, Warranty,Not in AMC and warranty'].str.contains('Warranty|AMC', case=False, na=False)
    ])
    coverage_percentage = (covered_devices / total_devices * 100) if total_devices > 0 else 0
    
   
    
    # 4. Most Popular Location
    top_location = filtered_df['Camera & NVR(1F or HO)'].value_counts().idxmax()
    location_count = filtered_df['Camera & NVR(1F or HO)'].value_counts().max()
    
    # Display KPIs in columns
    kpi1, kpi2, kpi3 = st.columns(3)
    
    with kpi1:
        st.metric(
            "Active Devices",
            f"{active_percentage:.1f}%",
            f"{active_devices} of {total_devices}"
        )
    
    with kpi2:
        st.metric(
            "Warranty/AMC Coverage",
            f"{coverage_percentage:.1f}%",
            f"{covered_devices} devices"
        )
    
    with kpi3:
        st.metric(
            "Most Devices At",
            top_location,
            f"{location_count} devices"
        )
    # --- Department-wise Device Count Summary ---
    if 'Camera & NVR(1F or HO)' in df.columns:
        dept_counts = df['Camera & NVR(1F or HO)'].astype(str).str.strip().str.upper().value_counts()
        plant_office_count = dept_counts.get('1F', 0)
        ho_count = dept_counts.get('HO', 0)
        st.markdown(f"""
        <div style='font-size:1.5rem; margin-top: 10px;'>
            <b>🗂 Department-wise Device Count</b><br>
            <span style='color:#1f77b4;'>🏭 Plant Office (1F): {plant_office_count}</span><br>
            <span style='color:#2ca02c;'>🏢 HO: {ho_count}</span>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.warning("'Camera & NVR(1F or HO)' column not found for department-wise count.")

    # --- Device status breakdown for filtered location (accurate, sums to 100%) ---
    total_devices_filtered = len(filtered_df)
    stock_count_filtered = (filtered_df['Initial Status'].str.strip() == 'Discard').sum()
    repair_count_filtered = (filtered_df['Initial Status'].str.strip() == 'Repair').sum()
    live_count_filtered = (filtered_df['Initial Status'].str.strip() == 'Live').sum()

    stock_pct_filtered = (stock_count_filtered / total_devices_filtered) * 100 if total_devices_filtered > 0 else 0
    repair_pct_filtered = (repair_count_filtered / total_devices_filtered) * 100 if total_devices_filtered > 0 else 0
    live_pct_filtered = (live_count_filtered / total_devices_filtered) * 100 if total_devices_filtered > 0 else 0

    st.markdown(
        f"<div style='font-size:1.1rem; margin-top: 8px; margin-bottom: 8px;'>"
        f"<b>Device status in this view:</b> "
        f"<span style='color:#2ecc71;'>🟢 Live: {live_pct_filtered:.1f}%</span> &nbsp; "
        f"<span style='color:#888;'>❌ Not in Use: {stock_pct_filtered:.1f}%</span> &nbsp; "
        f"<span style='color:#e67e22;'>🛠 Repair: {repair_pct_filtered:.1f}%</span></div>",
        unsafe_allow_html=True
    )
    # Add extra space before Select View Mode
    st.markdown("<br><br><br>", unsafe_allow_html=True)

    # Toggle between Grid and Table View
    view_mode = st.radio("Select View Mode", ["Grid View", "Table View"], horizontal=True)
    
    st.subheader(f"Devices at {selected_main_location}")
    
    if view_mode == "Grid View":
        # Display devices in a 5-column grid
        num_devices = len(filtered_df)
        num_cols = 5
        num_rows = math.ceil(num_devices / num_cols)
        
        for row in range(num_rows):
            cols = st.columns(num_cols)
            for col in range(num_cols):
                idx = row * num_cols + col
                if idx < num_devices:
                    device = filtered_df.iloc[idx]
                    with cols[col]:
                        device_html = create_device_box(device)
                        st.markdown(device_html, unsafe_allow_html=True)
    else:  # Table View
        st.dataframe(
            filtered_df,
            column_config={
                "Camera name": "Location",
                "Types": "Device Type",
                "Camera or NVR IP": "IP Address",
                "Initial Status": "Status",
                "Manufacturing Date": "Manufactured On",
                "AMC, Warranty,Not in AMC and warranty": "Coverage Status"
            },
            hide_index=True
        )
    
    # Add a little space before Analytics
    st.markdown("<br><br><br>", unsafe_allow_html=True)
    # --- Analytics ---
    st.subheader("\U0001F4CA Analytics")
    
    col1, col2 = st.columns(2)
    
    with col1:
        status_counts = filtered_df['Initial Status'].value_counts()
        if not status_counts.empty:
            fig_status = px.pie(
                values=status_counts.values,
                names=status_counts.index,
                title=f"Device Status Distribution in {selected_main_location}",
                color_discrete_sequence=px.colors.qualitative.Pastel2  # Second color scheme
            )
            st.plotly_chart(fig_status, use_container_width=True)
    
    with col2:
        types_counts = filtered_df['Types'].value_counts()
        if not types_counts.empty:
            fig_types = px.pie(
                values=types_counts.values,
                names=types_counts.index,
                title=f"Device Types Distribution in {selected_main_location}",
                color_discrete_sequence=px.colors.qualitative.Dark24  # Third color scheme
            )
            st.plotly_chart(fig_types, use_container_width=True)
    

    
    # Warranty Coverage Chart
    with st.expander("🛡️ Warranty Status Summary"):
        # Clean and count warranty statuses
        warranty_counts = filtered_df['AMC, Warranty,Not in AMC and warranty'].value_counts()
        
        fig_warranty = px.pie(
            values=warranty_counts.values,
            names=warranty_counts.index,
            title="Warranty & AMC Coverage"
        )
        
        # Add summary metrics
        wcol1, wcol2, wcol3 = st.columns(3)
        with wcol1:
            st.metric("Under AMC", len(filtered_df[filtered_df['AMC, Warranty,Not in AMC and warranty'].str.contains('AMC', case=False, na=False)]))
        with wcol2:
            st.metric("Under Warranty", len(filtered_df[filtered_df['AMC, Warranty,Not in AMC and warranty'].str.contains('Warranty', case=False, na=False)]))
        with wcol3:
            st.metric("No Coverage", len(filtered_df[filtered_df['AMC, Warranty,Not in AMC and warranty'].str.contains('Not in', case=False, na=False)]))
        
        st.plotly_chart(fig_warranty, use_container_width=True)
    

    
    # Device Age Report Section
    with st.expander("\U0001F4C5 Device Age Report"):
        try:
            # Convert PO Date to datetime and calculate age
            if 'PO Date' in filtered_df.columns:
                filtered_df['PO Date'] = pd.to_datetime(filtered_df['PO Date'], errors='coerce', dayfirst=True)
                filtered_df['Device Age (Years)'] = (pd.Timestamp.now() - filtered_df['PO Date']).dt.total_seconds() / (365.25 * 24 * 60 * 60)
            else:
                st.warning("'PO Date' column not found. Age calculation will be skipped.")
                filtered_df['Device Age (Years)'] = None
            # Calculate average age by device type
            avg_age_by_type = filtered_df.groupby('Types')['Device Age (Years)'].mean().round(1)
            # Create bar chart for average device age
            fig_age = go.Figure(data=[
                go.Bar(
                    x=list(avg_age_by_type.index),
                    y=list(avg_age_by_type.values),
                    text=[f"{age:.1f} years" for age in avg_age_by_type.values],
                    textposition='auto',
                )
            ])
            fig_age.update_layout(
                title="Average Device Age by Type",
                xaxis_title="Device Type",
                yaxis_title="Average Age (Years)",
                showlegend=False
            )
            st.plotly_chart(fig_age, use_container_width=True)
            # Highlight old devices (> 5 years)
            AGE_THRESHOLD = 5  # years
            old_devices = filtered_df[filtered_df['Device Age (Years)'] > AGE_THRESHOLD].copy()
            if not old_devices.empty:
                st.subheader(f"⚠️ Devices Older Than {AGE_THRESHOLD} Years")
                st.markdown(f"**Total aged devices: {len(old_devices)}**")
                # Format the age and date columns
                old_devices['Device Age (Years)'] = old_devices['Device Age (Years)'].round(1)
                old_devices['PO Date'] = old_devices['PO Date'].dt.strftime('%Y-%m-%d')
                # Display old devices
                st.dataframe(
                    old_devices[[
                        'Camera name', 'Types', 'Model', 
                        'PO Date', 'Device Age (Years)', 'Initial Status'
                    ]].sort_values('Device Age (Years)', ascending=False),
                    column_config={
                        "Camera name": "Location",
                        "Types": "Device Type",
                        "PO Date": "PO Date",
                        "Device Age (Years)": st.column_config.NumberColumn(
                            "Age (Years)",
                            help="Device age in years",
                            format="%.1f"
                        ),
                        "Initial Status": "Current Status"
                    },
                    hide_index=True
                )
                # Age distribution summary
                st.subheader("Age Distribution Summary")
                age_ranges = [
                    (0, 2, "0-2 years"),
                    (2, 5, "2-5 years"),
                    (5, float('inf'), "5+ years")
                ]
                age_distribution = []
                for start, end, label in age_ranges:
                    count = len(filtered_df[
                        (filtered_df['Device Age (Years)'] >= start) & 
                        (filtered_df['Device Age (Years)'] < end)
                    ])
                    percentage = (count / len(filtered_df) * 100) if len(filtered_df) > 0 else 0
                    age_distribution.append({
                        "Age Range": label,
                        "Count": count,
                        "Percentage": f"{percentage:.1f}%"
                    })
                st.table(pd.DataFrame(age_distribution))
            else:
                st.success(f"No devices older than {AGE_THRESHOLD} years found.")
        except Exception as e:
            st.warning("Could not generate age report. Please ensure the 'PO Date' column exists and contains valid dates.")
    
    # Recent Changes Log Section
    with st.expander("🔄 Recent Changes Log (Last 7 Days)"):
        try:
            # Convert Last Updated to datetime
            df['Last Updated'] = pd.to_datetime(df['Last Updated'])
            
            # Calculate the date threshold (7 days ago)
            seven_days_ago = pd.Timestamp.now() - pd.Timedelta(days=7)
            
            # Filter recent updates
            recent_updates = df[df['Last Updated'] >= seven_days_ago].copy()
            
            if not recent_updates.empty:
                # Sort by Last Updated (most recent first)
                recent_updates = recent_updates.sort_values('Last Updated', ascending=False)
                
                # New Devices Added
                st.subheader("New Devices Added")
                new_devices = recent_updates[['Camera name', 'Types', 'Model', 'Initial Status', 'Last Updated']]
                
                if not new_devices.empty:
                    # Format the dataframe for display
                    new_devices['Last Updated'] = new_devices['Last Updated'].dt.strftime('%Y-%m-%d %H:%M')
                    st.dataframe(
                        new_devices,
                        column_config={
                            "Camera name": "Location",
                            "Types": "Device Type",
                            "Model": "Model",
                            "Initial Status": "Status",
                            "Last Updated": "Added On"
                        },
                        hide_index=True
                    )
                else:
                    st.markdown("*No new devices added in the last 7 days.*")
                
                # Status Changes
                st.subheader("Status Changes")
                status_changes = recent_updates[
                    (recent_updates['Initial Status'] != "Live")
                ][['Camera name', 'Types', 'Initial Status', 'Last Updated']]
                
                if not status_changes.empty:
                    # Format the dataframe for display
                    status_changes['Last Updated'] = status_changes['Last Updated'].dt.strftime('%Y-%m-%d %H:%M')
                    st.dataframe(
                        status_changes,
                        column_config={
                            "Camera name": "Location",
                            "Types": "Device Type",
                            "Initial Status": "Current Status",
                            "Last Updated": "Changed On"
                        },
                        hide_index=True
                    )
                else:
                    st.markdown("*No status changes in the last 7 days.*")
            else:
                st.markdown("*No updates found in the last 7 days.*")
            
        except Exception as e:
            st.warning("Could not load recent changes. Please ensure the 'Last Updated' column exists and contains valid dates.")
else:
    st.warning("No devices found for the selected location.") 