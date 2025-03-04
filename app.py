import streamlit as st
import json
import pandas as pd
import os
from datetime import datetime
from Hjelpeskript.add_days_to_date import add_working_days_with_holidays
from Hjelpeskript.kommune_til_fylke import finn_fylke
from Hjelpeskript.woc_excel_sortfile import split_excel_by_customer_category

def extract_data_from_json(json_data):
    extracted_data = []

    for entry in json_data:
        if entry.get("supplier", {}).get("contactPersons"):
            continue
        elif not entry.get("wocOrderStatus", {}).lower() in ['accepted', 'received']:
            continue

        # Extract relevant data using your existing functions
        item = extract_item(entry)
        addresse, municipality, fylke, coordsys, x_koordinat, y_koordinat = extract_work_order_details(entry)
        kunde_navn, telefon_nr = extract_contact_info(entry)
        issued_date = entry.get("issuedDate").split("T")[0]
        bookes_innen = add_working_days_with_holidays(issued_date, 5)
        ordre_dato = get_latest_accept_workorder_date(entry.get("activityLog")) or issued_date
        last_transaction_date = format_date(entry.get("modifiedDate"))
        dato_leveranse = format_date(entry.get("deliveryPeriod", {}).get("endDate"))
        ordrenr_leveranse = entry.get("clientOrderId", {}).get("referenceNumber")
        sambandsnummer, _ = extract_service_details(entry, item)
        LU_nummer = extract_lu_number(entry)
        spidernummer = extract_spidernumber(entry)
        woc_type_oppdrag = extract_woc_type_oppdrag(entry)
        orderlines_productId = extract_product_ids(entry)
        prioritert_product_id = get_highest_priority_product(orderlines_productId, woc_type_oppdrag)
        orderinfo_description = entry.get("detailedOrderInformation", {}).get("orderDescription")
        product_description = extract_product_descriptions(entry)
        VULA_nr = extract_vula_numbers(entry, item)
        contract_details = entry.get("contract", {}).get("detailedPurchaseArea")
        customer_category = entry.get("detailedOrderInformation", {}).get("customerCategory")
        oppdrag_kategori = determine_oppdrag_kategori(customer_category, contract_details)
        gpon_p2p_woc = entry.get("areaOfSubject")
        gpon_from_detailed = entry.get("detailedAreaOfSubject")
        status_leveranse = determine_status_leveranse(orderlines_productId, spidernummer, orderinfo_description, contract_details, gpon_p2p_woc, VULA_nr, oppdrag_kategori, item)
        type_oppdrag = determine_type_oppdrag(orderinfo_description, status_leveranse, woc_type_oppdrag, VULA_nr)
        FTTx = determine_fttx(orderlines_productId, spidernummer, contract_details, gpon_p2p_woc, orderinfo_description, status_leveranse, oppdrag_kategori, item)
        gpon_p2p = determine_gpon_p2p(status_leveranse, VULA_nr, gpon_p2p_woc)

        extracted_data.append([
            item, addresse, municipality, fylke, kunde_navn, telefon_nr, issued_date, bookes_innen,
            ordre_dato, None, None, dato_leveranse, None, None, None, ordrenr_leveranse,
            sambandsnummer, "WOC", spidernummer, status_leveranse, contract_details, FTTx,
            oppdrag_kategori, gpon_p2p, gpon_p2p_woc, gpon_from_detailed, type_oppdrag,
            woc_type_oppdrag, LU_nummer, last_transaction_date, prioritert_product_id,
            orderlines_productId, VULA_nr, product_description, coordsys, x_koordinat, y_koordinat,
            orderinfo_description, customer_category
        ])
    
    return extracted_data

columns = [
    "Item", "Adresse", "Kommune", "Fylke", "Kunde", "Telefon", "Issued Date", "Bookes innen",
    "Ordredato", "Entreprenør", "Start arbeid tidligst", "Dato leveranse", "WOC Status", "BC Status", "UE Status",
    "Ordrenummer", "Sambandsnummer", "WOC/connector", "Spidernummer", "Status Leveranse", "kontraktdetaljer",
    "Type FTTx", "Kunde Kategori", "GPON/P2P", "GPON/P2P - WOC", "GPON/AEG - from detailedAreaOfSubject",
    "Type oppdrag", "Type oppdrag WOC", "LU-nummer", "Last Transaction Date", "Hovedprodukt",
    "Produkt ID", "VULA ?", "Beskrivelse av produkt", "Coordsys", "X-koordinat", "Y-koordinat",
    "Orderinfo Description", "Customer Category"
]

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
    rows = extract_data_from_json(json_data)
    df = pd.DataFrame(rows, columns=columns)
    
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
