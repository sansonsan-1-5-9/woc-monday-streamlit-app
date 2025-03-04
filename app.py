import streamlit as st
import json
import pandas as pd
import os
from datetime import datetime
from Hjelpeskript.add_days_to_date import add_working_days_with_holidays
from Hjelpeskript.kommune_til_fylke import finn_fylke
from Hjelpeskript.woc_excel_sortfile import split_excel_by_customer_category


# Function to process JSON data
def process_json(json_data):
    rows = extract_data_from_json(json_data)
    df = pd.DataFrame(rows, columns=columns)
    return df


# Streamlit UI
st.set_page_config(page_title="WoC Report Processor", layout="wide")

st.title("📊 WoC Report Processor")
st.write("Upload a JSON file, process it, and download the results as an Excel file.")

# File uploader
uploaded_file = st.file_uploader("📂 Upload JSON File", type="json")

if uploaded_file:
    # Load JSON
    json_data = json.load(uploaded_file)

    # Process Data
    df = process_json(json_data)

    # Show the processed DataFrame
    st.subheader("📑 Extracted Data Preview")
    st.dataframe(df)

    # Save DataFrame to an Excel file
    excel_file = "processed_data.xlsx"
    df.to_excel(excel_file, index=False)

    # Provide download link
    with open(excel_file, "rb") as file:
        st.download_button(
            label="📥 Download Processed Excel File",
            data=file,
            file_name="WoC_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
