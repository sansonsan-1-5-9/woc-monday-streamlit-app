import streamlit as st
import json
import os
import subprocess

st.subheader("🔍 Installed Packages in Streamlit Cloud")
result = subprocess.run(["pip", "list"], capture_output=True, text=True)
st.text(result.stdout)

# Streamlit UI
st.set_page_config(page_title="WoC Report Processor", layout="wide")
st.title("📊 WoC Report Processor")
st.write("Upload a JSON file to process it using the 'woc-to-monday.py' script.")

# File uploader
uploaded_file = st.file_uploader("📂 Upload JSON File", type="json")

if uploaded_file:
    # Save uploaded file to a temporary location
    temp_json_path = "uploaded_file.json"
    with open(temp_json_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.write("Processing the uploaded file...")

    # Run the woc-to-monday.py script
    result = subprocess.run(["python", "woc_to_monday.py", temp_json_path], capture_output=True, text=True)

    # Display output logs
    st.text(result.stdout)
    st.text(result.stderr)

    # List of expected output files
    output_files = ["Monday_Import.xlsx", "Monday_Import - B.xlsx", "Monday_Import - P.xlsx"]

    for output_file in output_files:
        if os.path.exists(output_file):
            with open(output_file, "rb") as file:
                st.download_button(
                    label=f"📥 Download {output_file}",
                    data=file,
                    file_name=output_file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.error(f"Processing failed or {output_file} was not generated.")
