import os
import io
import zipfile
import datetime

import streamlit as st
import  pandas as pd
import requests


from reports_core import filter_master, generate_reports

st.set_page_config(page_title="Grower Reports", layout="wide")

st.title("Grower Report Generator")

#1. Upload Master File
master_file = st.file_uploader("Upload your Return to Grower Report.xlsx", type="xlsx")
if not master_file:
    st.info("Please upload file to begin.")
    st.stop()


#2. Read and List Growers
with st.spinner("Loading file..."):
    tmp_master = "temp_master.xlsx"
    with open(tmp_master, "wb") as f:
        f.write(master_file.read())

    df_master = pd.read_excel(tmp_master, header=1).dropna(how="all")
    df_master['Packed Date'] = pd.to_datetime(
        df_master['Packed Date'],
         dayfirst=True,
         errors='coerce'
    )
    df_master['GrowerName'] = (
        df_master['Supplier'].astype(str)
        .str.split("(", n=1)
        .str[0]
        .str.strip()
    )

#3 Loading grower settings
settings_df = pd.read_excel("grower_settings.xlsx", sheet_name="Filters")
st.markdown("### Grower-specific Filter Settings")
st.dataframe(settings_df, use_container_width=True)


#4. Generate & Download
if st.button("Generate Reports"):
    out_dir = "temp_reports"
    os.makedirs(out_dir,exist_ok=True)

    report_path = []
    today = datetime.date.today()

    for _, row in settings_df.iterrows():
        grower = row["GrowerName"]

        if row["FilterType"] == "Past month":
            start = today - datetime.timedelta(days=30)
        else:
            start = pd.to_datetime(row["CustomStart"]).date()
        end = today

        start_ts = pd.Timestamp(start)
        end_ts = pd.Timestamp(end) + pd.Timedelta(days=1)

        subset = df_master[
            (df_master["GrowerName"] == grower) &
            (df_master["Packed Date"] >= start_ts) &
            (df_master["Packed Date"] < end_ts)
        ]

        subset = subset[ subset["Trays"].fillna(0) !=0]

        paths = generate_reports(
            subset,
            template_path="TBC_Grower_Report_Template.xlsx",
            output_dir=out_dir,
            growers=[grower]
        )
        report_path.extend(paths)

        #4 Parse email recipients
        raw_emails = row["Emails"]
        to_email_list = [e.strip() for e in raw_emails.split(",") if e.strip()]

        webhook = st.secrets["make_webhook_url"]
        for p in paths:
            with open(p, "rb") as f:
                files = {
                    "Report File":(
                        os.path.basename(p),
                        f,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                }
                data = {
                    "grower": grower,
                    "emails": ",".join(to_email_list),
                }

                requests.post(webhook, data=data, files=files)

    #bundle into ZIP
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for p in report_path:
            zf.write(p, os.path.basename(p))
    zip_buffer.seek(0)

    import datetime
    date_str = datetime.date.today().strftime("%Y.%m.%d")

    st.success(f"Generated {len(report_path)} report(s).")
    st.download_button(
        "Download All Reports",
        zip_buffer,
        file_name=f"Grower Reports {date_str}.zip",
        mime="application/zip"
    )