import os
import io
import zipfile

import streamlit as st
import  pandas as pd


from reports_core import filter_master, generate_reports, send_reports

st.set_page_config(page_title="Grower Reports", layout="wide")

st.title("Grower Report Generator")

master_file = st.file_uploader("Upload your Return to Grower Report.xlsx", type="xlsx")

today = pd.Timestamp("today").normalize()
one_month_ago = today - pd.DateOffset(months=1)
start_date, end_date = st.date_input(
    "Select date range for Packed Date",
    value=(one_month_ago, today)
    )

st.markdown(
    "Enter grower names and emails, one per line in the format:\n\n"
    "'Grower Name <email@example.com>'\n"
    "e.g. 'John Smith <JohnSmith@berryco.com>'"
    )

email_input = st.text_area("Grower Email Map", height=150)

if st.button("Generate & Send Reports"):
    if not master_file:
        st.error("Please upload the master Excel file first.")
        st.stop()

    email_map = {}
    for line in email_input.splitlines():
        parts = line.strip().split()
        if len(parts) >=2:
            name = " ".join(parts[:-1])
            addr = parts[-1]
            email_map[name] = addr
    if not email_map:
        st.error("Please enter at least one grower and email.")
        st.stop()

    with st.spinner("Filterting master data..."):
        tmp_path = os.path.join("temp_master.xlsx")
        with open(tmp_path, "wb") as f:
            f.write(master_file.read())
        df = filter_master(tmp_path, start_date, end_date)

    out_dir = "temp_reports"
    paths = generate_reports(df,
                             template_path="TBC_Grower_Report_Template.xlsx",
                             output_dir=out_dir,
                             growers=list(email_map.keys()))
    
    smtp_cfg = {
        "host":"smtp.office365.com",
        "port": 587,
        "user": "sbuss@theberrycollective.com.au",
        "password": st.secrets["smtp_password"],
        "from_address": "marketing@theberrycollective.com.au"
    }


    with st.spinner("Sending emails..."):
        send_reports(paths, email_map, smtp_cfg)

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w") as z:
        for p in paths:
            z.write(p, os.path.basename(p))
    zip_buf.seek(0)
    st.download_button("Download All Reports (ZIP)", zip_buf, file_name="reports.zip")

    st.success("Done! Reports generated, emailed, and ready to download.")