import streamlit as st
import pandas as pd
import json
import io
import time
from logic import DataProcessor

st.set_page_config(page_title="Data Processor", page_icon="📊", layout="wide")

st.title("MP Data Processor")
st.caption("Upload multiple Excel/CSV files, process them with your `groups.json` config, and download a clean, standardized dataset.")

# Sidebar for configuration
st.sidebar.header("Configuration")
groups_file = st.sidebar.file_uploader("Upload groups.json", type="json")

# File uploader
uploaded_files = st.file_uploader(
    "Upload Excel/CSV files", 
    type=["xlsx", "csv"], 
    accept_multiple_files=True
)

if groups_file and uploaded_files:
    groups = json.load(groups_file)

    file_objs = []
    for uploaded_file in uploaded_files:
        if uploaded_file.type == "text/csv" or uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
            buffer = io.BytesIO()
            df.to_excel(buffer, index=False, engine="openpyxl")
            buffer.seek(0)
        else:
            buffer = io.BytesIO(uploaded_file.read())
            buffer.seek(0)
        file_objs.append((buffer, uploaded_file.name))

    processor = DataProcessor(groups)

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

    status_placeholder.success("✅ Processing complete!")

    if all_results:
        final_df = pd.concat(all_results, ignore_index=True)

        # Tabs for results
        tab1, tab2, tab3 = st.tabs(["🔍 Preview", "📥 Download", "📈 Summary"])

        with tab1:
            st.write("Here’s a preview of your processed dataset:")
            st.dataframe(final_df.head(50), use_container_width=True)

        with tab2:
            output_buffer = io.BytesIO()
            final_df.to_excel(output_buffer, index=False, engine="openpyxl")
            output_buffer.seek(0)

            st.download_button(
                label="📥 Download Excel",
                data=output_buffer,
                file_name="processed_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with tab3:
            st.write("### Quick Summary")
            st.metric("📂 Files Processed", len(file_objs))
            st.metric("📊 Rows Combined", len(final_df))
            st.metric("🧾 Columns Detected", len(final_df.columns))

    else:
        st.warning("⚠️ No data could be extracted.")
else:
    st.info("⬅️ Please upload a `groups.json` file and at least one Excel/CSV file to begin.")
