import streamlit as st
import json
import os
import subprocess
import shutil
import zipfile

# Streamlit UI configuration (must be first command)
st.set_page_config(page_title="WoC Report Processor", layout="wide")

# Ensure dependencies are installed
st.subheader("🔍 Checking & Installing Dependencies")
required_packages = ["pandas", "numpy", "openpyxl", "streamlit", "fpdf", "dateutil"]
for package in required_packages:
    subprocess.run(["pip", "install", package])

st.title("📊 WoC JSON-Report Processor")
st.write("Upload a JSON file to process it using the 'woc_to_monday.py' and 'Generere_PDF_fra_JSON.py' scripts.")

# File uploader
uploaded_file = st.file_uploader("📂 Upload JSON File", type="json")

if uploaded_file:
    # Save uploaded file to a temporary location
    temp_json_path = "uploaded_file.json"
    with open(temp_json_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.write("Processing the uploaded file...")

    # Run the woc_to_monday.py script
    result_monday = subprocess.run(["python", "woc_to_monday.py", temp_json_path], capture_output=True, text=True)
    
    # Run the Generere_PDF_fra_JSON.py script
    result_pdf = subprocess.run(["python", "Generere_PDF_fra_JSON.py", temp_json_path], capture_output=True, text=True)
    
    # List of expected output files
    st.subheader("📁 Output Files:")
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
    
    # Handle PDFs - Zip the generated PDFs folder
    pdf_base_dir = "generated_pdfs"  # Adjust this to match the actual output folder
    zip_file_path = "generated_pdfs.zip"
    
    if os.path.exists(pdf_base_dir):
        with zipfile.ZipFile(zip_file_path, "w", zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(pdf_base_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    zipf.write(file_path, os.path.relpath(file_path, pdf_base_dir))
        
        # Provide download button for ZIP file
        with open(zip_file_path, "rb") as zip_file:
            st.download_button(
                label="📥 Download All PDFs (ZIP)",
                data=zip_file,
                file_name="generated_pdfs.zip",
                mime="application/zip"
            )
    else:
        st.error("No PDFs were generated or directory does not exist.")
    
    # Display output logs
    # st.subheader("🔍 Script Output (WoC to Monday)")
    # st.text(result_monday.stdout)

    if result_monday.stderr:
        st.subheader("🚨 Errors (if any) - WoC to Monday")
        st.text(result_monday.stderr)
    
    # st.subheader("🔍 Script Output (Generate PDF)")
    # st.text(result_pdf.stdout)
    
    if (result_pdf.stderr):
        st.subheader("🚨 Errors (if any) - Generate PDF")
        st.text(result_pdf.stderr)
