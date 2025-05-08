import os
import io
import zipfile

import streamlit as st
import  pandas as pd


from reports_core import filter_master, generate_reports

st.set_page_config(page_title="Grower Reports", layout="wide")

st.title("Grower Report Generator")

#1. Upload Master File
master_file = st.file_uploader("Upload your Return to Grower Report.xlsx", type="xlsx")
if not master_file:
    st.info("Please upload file to begin.")
    st.stop()


#2. Pick Date Range
today = pd.Timestamp("today").normalize()
one_month_ago = today - pd.DateOffset(months=1)
start_date, end_date = st.date_input(
    "Select date range for Packed Date",
    value=(one_month_ago, today)
    )

#3. Read and List Growers
with st.spinner("Loading and filtering file..."):
    tmp_master = "temp_master.xlsx"
    with open(tmp_master, "wb") as f:
        f.write(master_file.read())
    df_master = filter_master(tmp_master, start_date, end_date)

all_growers = sorted(df_master['Grower Name'].unique())
selected = st.multiselect(
    "Select growers to generate a report for",
    options=all_growers,
    default=all_growers
)
if not selected:
    st.warning("Please select at least one grower.")
    st.stop()

#4. Generate & Download
if st.button("Generate Reports"):
    out_dir = "temp_reports"
    paths = generate_reports(
        df_master,
        template_path= "TBC_Grower_Report_Template.xlsx",
        output_dir=out_dir,
        growers=selected
    )

    #bundle into ZIP
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for p in paths:
            zf.write(p, os.path.basename(p))
    zip_buffer.seek(0)

    import datetime
    date_str = datetime.date.today().strftime("%Y.%m.%d")

    st.success(f"Generated {len(paths)} report(s).")
    st.download_button(
        "Download All Reports",
        zip_buffer,
        file_name=f"Grower Reports {date_str}.zip",
        mime="application/zip"
    )