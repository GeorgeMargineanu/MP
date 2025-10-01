import streamlit as st
import pandas as pd
import json
import io
import time
import os
from logic import DataProcessor
from excel_styling import style_and_export_excel
from guide import show_app_info

pd.set_option("future.no_silent_downcasting", True)

st.set_page_config(
    page_title="Data Processor",
    page_icon="static/images/logo.png",
    layout="wide"
)

st.sidebar.image("static/images/logo.png", width=120)
st.title("MP Data Processor")
st.caption(
    "Upload multiple Excel/CSV files, process them with your `groups.json` config, "
    "and download a clean, standardized dataset."
)

st.sidebar.header("⚙️ Configuration")
agency_commission = st.sidebar.number_input("Enter the agency commission (as percent, e.g., 1 for 1%)", value=1.0, step=0.1)
agency_commission = agency_commission / 100

# Automatically load groups.json
current_dir = os.path.dirname(os.path.abspath(__file__))
groups_path = os.path.join(current_dir, "groups.json")
with open(groups_path, "r", encoding="utf-8") as f:
    groups = json.load(f)

uploaded_files = st.file_uploader(
    "Upload Excel/CSV files",
    type=["xlsx", "csv"],
    accept_multiple_files=True
)

# ------------------- Session State -------------------
if "final_df" not in st.session_state:
    st.session_state.final_df = None
if "file_objs" not in st.session_state:
    st.session_state.file_objs = []
if "prev_agency_commission" not in st.session_state:
    st.session_state.prev_agency_commission = None

def _read_csv(uploaded_file):
    return pd.read_csv(uploaded_file)

# ------------------- Processing Trigger -------------------
process_button = st.sidebar.button("Process Files")

if process_button and uploaded_files:
    st.session_state.final_df = None
    st.session_state.prev_agency_commission = agency_commission

    file_objs = []

    for uploaded_file in uploaded_files:
        if uploaded_file.type == "text/csv" or uploaded_file.name.endswith(".csv"):
            df = _read_csv(uploaded_file)
            buffer = io.BytesIO()
            df.to_excel(buffer, index=False, engine="openpyxl")
            buffer.seek(0)
        else:
            buffer = io.BytesIO(uploaded_file.read())
            buffer.seek(0)
        file_objs.append((buffer, uploaded_file.name))

    processor = DataProcessor(groups, agency_commission)

    st.subheader("⚙️ Processing Files")
    progress_bar = st.progress(0)
    status_placeholder = st.empty()

    all_results = []
    for i, (buffer, name) in enumerate(file_objs, start=1):
        status_placeholder.info(f"Processing **{name}** ({i}/{len(file_objs)})...")
        df_chunk = processor.process_files([(buffer, name)])
        if not df_chunk.empty:
            all_results.append(df_chunk)
        progress_bar.progress(int(i / len(file_objs) * 100))
        time.sleep(0.2)

    status_placeholder.success("Processing complete!")

    if all_results:
        st.session_state.final_df = pd.concat(all_results, ignore_index=True)
        st.session_state.file_objs = file_objs
    else:
        st.warning("No data could be extracted.")

# ------------------- Display Tabs -------------------
if st.session_state.final_df is not None:
    final_df = st.session_state.final_df
    file_objs = st.session_state.file_objs

    tab1, tab2, tab3 = st.tabs(["🔍 Preview", "📥 Download", "📈 Summary"])

    with tab1:
        st.sidebar.markdown('-----------------')
        options = final_df["__source_file"].unique()
        select_provider = st.sidebar.selectbox("Select the provider for a quick view", options)
        st.write("🔍 Summary for the selected provider :")
        df_per_company = final_df[final_df["__source_file"] == select_provider]
        st.dataframe(df_per_company, use_container_width=True)

    with tab2:
        metadata = {
            "Brand": "",
            "Campaign": "",
            "Version": "",
            "Start": "",
            "End": ""
        }
        output_buffer = style_and_export_excel(final_df, metadata=metadata)

        st.download_button(
            label="Download Excel",
            data=output_buffer,
            file_name="processed_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with tab3:
        st.write("### Quick Summary")
        col1, col2, col3 = st.columns(3)
        col1.metric("📂 Files Processed", len(file_objs))
        col2.metric("📊 Rows Combined", len(final_df))
        col3.metric("🧾 Columns Detected", len(final_df.columns))
else:
    st.info("⬅ Please upload at least one Excel/CSV file and click **Process Files** to begin.")


show_app_info()