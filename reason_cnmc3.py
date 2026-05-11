import streamlit as st
import pandas as pd
import io
import re
import xlrd

# Authorized portal reasons (Full List)
FAULT_TYPES = sorted([
    "11 KV Jumper fault", "2G 3G Faults due to TCS antenna swap", "4G cloned node",
    "Aggregation Site Issue", "Backhaul Link Set-up Failure", "Battery backup zero/DG not working",
    "Battery life expired", "Battery not available", "BBNL Dependent Site", "BBNL Media issue",
    "BBNL OFC break", "BBNL-connected cascaded 4G SP sites", "BONDADA Power Issue at Saturation Site",
    "BSNL Rigger Issue", "BTS Burnt", "CEF Card issue", "Ceregon DMW Issue", "CFA issue",
    "CNTx-zone Media issue", "CNTx-zone OFC Break", "COVID-19 Restricted Entry", "CPAN Diversion",
    "CPAN Issue", "CPAN/CPE/ MADM system failure", "CPRI cable issue", "DC POwer Cable faulty",
    "DEPL issue", "DEPL OFC Break", "DEPL power issue for saturation sites", "DG faulty",
    "DG not available", "Diesel Shortage", "DUE TO HUB SITE", "E1 Failure", "EB Pole break",
    "EB Supply Low Voltage", "Electrical fault AVR etc.", "Electricity bill Payment",
    "Electricity Shutdown", "eNB S-1 link down TCS/TEJAS team rectify", "ENodeB Network Unreachable",
    "FALSE ALARM", "False Alarm/BTS working", "Far end LWE site media issue", "Far end LWE site power issue",
    "Far end Mini Link/DMW/UBR media issue", "Far end Power supply failure", "Faulty Battery bank",
    "GPS Issue", "Hardware (BTS)", "HARDWARE (PP)", "HARDWARE THEFTS", "HFCL UBR Issue",
    "Honey bee issue on saturation/LWE site", "IaasP - OFC Break", "IaasP - Site Access issue",
    "IaasP L1 Activity Pending", "IaasP- Any Other", "IaasP- Battery issue", "IaasP- Solar power fail",
    "IaasP- Theft", "Incomplete installation (4G)", "Insufficient Rectifier", "IP Issue",
    "Laasp-Mains Failure", "LT power cable fault", "LVDT fault in PP", "MAAN cluster media issue",
    "MAAN Issue", "MAAN migration Issue", "MADM Issue", "MAHA-IT OFC Fault", "Mains Failure",
    "Mains not available", "Media buildup station failure", "Media issue at HUB site",
    "Media issue on LWE site", "MINILINK FAULT", "Modem Fault", "New Fault Reason", "No 2G Users",
    "No 3G Users", "Nonpayment of Rental Dues", "Other Media Failure", "Other OFC break",
    "Outsource vendor issue", "Owner Issue", "PACE needs to rectify media issue", "Planned Sites",
    "Poor battery backup", "Power Cable Theft", "Power issue at HUB site", "POWER PLANT FAULT",
    "PP Control panel fault", "PTPL power issue on LWE site", "RAC Card issue", "Rectifier faulty",
    "RF Cable Theft", "Road blocked due to snowfall", "RRH Antenna Port-0 VSWR Error",
    "RRH Antenna Port-1 VSWR Error", "RRH Link Down", "Scheduled Lock", "Sector Down",
    "Sector(s) locked", "Shut Down", "Site Deleted from Network", "SITE RESETTING ISSUE",
    "Site Sealed by Local Authority", "Site Swap in Progress", "SOFTWARE ISSUE", "Solar power issue",
    "SSA Media issue", "SSA OFC Break", "State Elec Board Transformer faulty", "STR Cable Cut",
    "SUPERVISION OUTAGE", "TCS core MME down", "TCS/TEJAS team visit requird for 2nd level mtce",
    "TEJAS EMS Issue", "TEJAS GPS ISSUE", "Tejas/TCS needs to rectify hardware issue",
    "Tejas/TCS needs to rectify software issue", "Tempature Issue", "Test Sites",
    "Tested on Nokia/ZTE core planned to shift IX.2 core", "Tower collapsed", "Tower/Shelter Damage",
    "Transformer faults", "Tx team visit reqrd for media loss rectification of TCS eNodeB",
    "USO-Remote site commercially nonviable", "VLAN Media issue", "VSAT Media issue",
    "VSWR Alarm", "Watch and ward issue", "Water Logging Issue", "Zero battery backup"
])


def clean_id(val):
    """Standardizes IDs: removes spaces, uppercase, removes .0 from numeric inputs."""
    if pd.isna(val): return ""
    s = str(val).strip().upper()
    if s.endswith('.0'): s = s[:-2]
    return s


def robust_read(file):
    """Handles different file encodings and formats to prevent crashes."""
    content = file.read()
    file.seek(0)
    try:
        wb = xlrd.open_workbook(file_contents=content, ignore_workbook_corruption=True)
        ws = wb.sheet_by_index(0)
        data = [ws.row_values(i) for i in range(ws.nrows)]
        df = pd.DataFrame(data[1:], columns=data[0])
    except:
        try:
            df = pd.read_html(io.BytesIO(content))[0]
        except:
            df = pd.read_excel(io.BytesIO(content))
    df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]
    return df


st.set_page_config(page_title="BTS Master Log", layout="wide")
st.title("📊 BTS Master Log & JTO Summary")

# 1. Sidebar Configuration
st.sidebar.header("📁 Step 1: Configuration")
lookup_file = st.sidebar.file_uploader("Upload BTSIPID_PKEY Source File", type=["csv", "xlsx"])

# 2. Main Uploads
st.header("Step 2: Upload Data Files")
col1, col2 = st.columns(2)
with col1:
    master_file = st.file_uploader("Existing Master Log", type=["xlsx"])
with col2:
    daily_files = st.file_uploader("New Daily Logs (Adds only empty Fault Types)", type=["xlsx", "xls"],
                                   accept_multiple_files=True)

if st.button("🚀 Process & Generate Report"):
    if not lookup_file:
        st.error("Please upload the BTSIPID_PKEY file in the sidebar!")
    else:
        # --- LOAD LOOKUP DATA (Strict Mapping) ---
        try:
            l_df = pd.read_csv(lookup_file) if lookup_file.name.endswith('.csv') else pd.read_excel(lookup_file)
            l_df.columns = [str(c).strip().upper() for c in l_df.columns]

            # Use BTSIPID as primary key to get correct JTO and SDCA
            mapping_data = l_df[['BTSIPID', 'SDCA', 'JTO INCHARGE']].copy()
            mapping_data['JOIN_KEY'] = mapping_data['BTSIPID'].apply(clean_id)
            mapping_data = mapping_data[['JOIN_KEY', 'SDCA', 'JTO INCHARGE']].drop_duplicates(subset=['JOIN_KEY'],
                                                                                              keep='last')
            mapping_data.rename(columns={'SDCA': 'new_sdca', 'JTO INCHARGE': 'new_jto'}, inplace=True)
        except Exception as e:
            st.error(f"Error reading lookup file. Ensure 'BTSIPID', 'SDCA', and 'JTO INCHARGE' columns exist.")
            st.stop()

        # --- PROCESS LOGS ---
        all_dfs = []
        if master_file:
            all_dfs.append(robust_read(master_file))

        if daily_files:
            for f in daily_files:
                d_df = robust_read(f)
                if d_df is not None:
                    if 'fault_type' not in d_df.columns: d_df['fault_type'] = ""
                    # Keep only rows where fault_type is missing
                    d_df = d_df[d_df['fault_type'].isna() | (d_df['fault_type'].astype(str).str.strip() == "")]

                    # Extract date from filename if possible
                    d_match = re.search(r'(\d{4}-\d{2}-\d{2})', f.name)
                    d_df['source_date'] = d_match.group(1) if d_match else "New"
                    all_dfs.append(d_df)

        if all_dfs:
            combined_df = pd.concat(all_dfs, ignore_index=True)

            # Deduplicate based on IP and Down/Up times
            unique_keys = ['bts_ip_id', 'bts_down_dt', 'bts_up_dt']
            if all(k in combined_df.columns for k in unique_keys):
                combined_df = combined_df.drop_duplicates(subset=unique_keys, keep='first')

                # --- RE-MAP JTO & SDCA BASED ON IP ID ---
                combined_df['match_key'] = combined_df['bts_ip_id'].apply(clean_id)

                # Remove old mapping columns to avoid duplicates
                for col in ['sdca', 'jto_incharge', 'sdca_name', 'Status']:
                    if col in combined_df.columns: combined_df.drop(columns=[col], inplace=True)

                # Merge fresh mapping from Source file
                combined_df = combined_df.merge(mapping_data, left_on='match_key', right_on='JOIN_KEY', how='left')
                combined_df.rename(columns={'new_sdca': 'sdca', 'new_jto': 'jto_incharge'}, inplace=True)
                combined_df.drop(columns=['match_key', 'JOIN_KEY'], inplace=True)

                # --- SUMMARY DASHBOARD ---
                pending_count = combined_df[
                    combined_df['fault_type'].isna() | (combined_df['fault_type'].astype(str).str.strip() == "")].shape[
                    0]
                resolved_count = len(combined_df) - pending_count

                st.subheader("Current Status Summary")
                m1, m2, m3 = st.columns(3)
                m1.metric("Total Cases", len(combined_df))
                m2.metric("Pending (Empty Fault Type)", pending_count, delta_color="inverse")
                m3.metric("Resolved", resolved_count)

                # --- EXCEL EXPORT WITH AUTO-UPDATE FORMULAS ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    # Initialize Status column
                    combined_df['Status'] = ""
                    combined_df.to_excel(writer, index=False, sheet_name='Master_Log')
                    workbook, worksheet = writer.book, writer.sheets['Master_Log']

                    fault_col = combined_df.columns.get_loc('fault_type')
                    status_col = combined_df.columns.get_loc('Status')

                    # Write Excel Formulas for Auto-Updating Status
                    for row_num in range(1, len(combined_df) + 1):
                        fault_cell = xlrd.colname(fault_col) + str(row_num + 1)
                        # Formula: If fault type is blank, status is PENDING, else CLOSED
                        formula = f'=IF(TRIM({fault_cell})="", "PENDING", "CLOSED")'
                        worksheet.write_formula(row_num, status_col, formula)

                    # Add Dropdown list for Fault Type
                    list_sheet = workbook.add_worksheet('Dropdowns')
                    for i, reason in enumerate(FAULT_TYPES):
                        list_sheet.write(i, 0, reason)
                    list_sheet.hide()

                    worksheet.data_validation(1, fault_col, 65000, fault_col, {
                        'validate': 'list',
                        'source': '=Dropdowns!$A$1:$A$' + str(len(FAULT_TYPES))
                    })

                    # Final Formatting
                    header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
                    for col_num, value in enumerate(combined_df.columns.values):
                        worksheet.write(0, col_num, value, header_format)

                    worksheet.set_column(fault_col, fault_col, 35)
                    worksheet.set_column(status_col, status_col, 15)
                    worksheet.set_column(combined_df.columns.get_loc('jto_incharge'),
                                         combined_df.columns.get_loc('jto_incharge'), 20)

                st.success("Report Processed Successfully!")
                st.download_button("📥 Download Final Master Report.xlsx", output.getvalue(), "Master_Report.xlsx")
            else:
                st.error(f"Required columns (bts_ip_id, bts_down_dt, bts_up_dt) missing in daily logs.")