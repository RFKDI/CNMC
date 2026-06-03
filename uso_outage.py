import streamlit as st
import pandas as pd

# Set page configuration
st.set_page_config(page_title="Site Performance Dashboard", layout="wide")

st.title("📊 BTS Site Performance Dashboard")

# 1. File Uploader
uploaded_files = st.file_uploader("Upload multiple CSV files", type=["csv"], accept_multiple_files=True)

if uploaded_files:
    # Merge files
    df_list = []
    for file in uploaded_files:
        df_temp = pd.read_csv(file)
        df_list.append(df_temp)

    df = pd.concat(df_list, ignore_index=True)

    # Ensure date column is datetime format for accurate filtering
    if 'Report Date' in df.columns:
        df['Report Date'] = pd.to_datetime(df['Report Date'])

    # 2. Sidebar Filter: SSAID
    st.sidebar.header("Filter Options")
    ssa_ids = sorted(df['SSAID'].dropna().unique().tolist())
    selected_ssa = st.sidebar.multiselect("Select SSAID", ssa_ids, default=ssa_ids)

    # Apply SSAID Filter
    df_filtered = df[df['SSAID'].isin(selected_ssa)]

    # 3. Main Filter: BTS Name
    bts_names = sorted(df_filtered['BTS Name'].dropna().unique().tolist())
    selected_bts = st.multiselect("Select BTS Name", bts_names, default=bts_names)

    # Apply BTS Name Filter
    if selected_bts:
        df_filtered = df_filtered[df_filtered['BTS Name'].isin(selected_bts)]

    # 4. Date Filter
    if 'Report Date' in df_filtered.columns:
        min_date = df_filtered['Report Date'].min()
        max_date = df_filtered['Report Date'].max()

        # Date Input
        date_range = st.date_input("Select Date Range", [min_date, max_date])

        if len(date_range) == 2:
            start_date, end_date = date_range
            df_filtered = df_filtered[
                (df_filtered['Report Date'].dt.date >= start_date) &
                (df_filtered['Report Date'].dt.date <= end_date)
                ]

    # Display the filtered data
    st.write(f"### Filtered Data ({len(df_filtered)} records)")
    st.dataframe(df_filtered)

    # 5. Download Button
    csv = df_filtered.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Download Filtered Data as CSV",
        data=csv,
        file_name='filtered_site_performance.csv',
        mime='text/csv',
    )
else:
    st.info("Please upload CSV files to proceed.")