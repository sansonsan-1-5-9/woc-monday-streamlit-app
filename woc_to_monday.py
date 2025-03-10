import json
import os
import sys
import pandas as pd
from datetime import datetime
from Hjelpeskript.add_days_to_date import add_working_days_with_holidays
from Hjelpeskript.kommune_til_fylke import finn_fylke
from Hjelpeskript.woc_excel_sortfile import split_excel_by_customer_category
from Hjelpeskript.fylke_kommune_entreprenor import finn_entreprenor

# Use the uploaded JSON file passed from Streamlit
if len(sys.argv) > 1:
    json_file_path = sys.argv[1]
else:
    raise FileNotFoundError("No JSON file provided to the script.")

# Define a valid output directory
output_directory = os.path.dirname(os.path.abspath(__file__))

# Ensure the directory exists before saving
if not os.path.exists(output_directory):
    os.makedirs(output_directory, exist_ok=True)

# Define the full file path
target_excel_file = os.path.join(output_directory, "Monday_Import.xlsx")

# Sletter excelfilen om den finnes fra før
try:
    os.remove(target_excel_file)
    print(f"{target_excel_file} er slettet.")
except FileNotFoundError:
    print(f"Filen {target_excel_file} finnes ikke.")
except PermissionError:
    print(f"Du har ikke tilgang til å slette {target_excel_file}.")
except Exception as e:
    print(f"En feil oppstod: {e}")



### DEFINERER FUNKSJONER ###

# Finne Item-navn
def extract_item(entry):
    user1_info = entry.get("detailedOrderInformation", {}).get("user1", {})
    entr_title = entry.get("title")

    if user1_info:
        return user1_info.get("fullName")

    cp_fullname = entry.get("connectionPoint", {}).get("fullName")
    cp_id = entry.get("connectionPoint", {}).get("id")
    detailed_aos = entry.get("detailedAreaOfSubject")
    if entr_title.strip().endswith("OLT"):
        return f"{cp_fullname}-{cp_id}-{detailed_aos}-OLT"

    return f"{cp_fullname}-{cp_id}-{detailed_aos}"

# WorkOrderAddress
def extract_work_order_details(entry):

    work_order_address = entry.get("workOrderAddress", [{}])[0]

    # Henter gateadresse eller matrikkeladresse
    street_address = work_order_address.get("streetAddress", {})
    if street_address:
        municipality = street_address.get("municipalityName")
        adr_step = f"{street_address.get('streetName', '')} {street_address.get('houseNumber', '')}"
        house_char = street_address.get('houseChar', '')
        if house_char:
            adr_step += house_char
        adresse = f"{adr_step}, {street_address.get('city')}, Norge"
        post_nummer = street_address.get('postalCode')
    else:
        cadastral_unit = work_order_address.get('cadastralUnit', {})
        municipality = cadastral_unit.get('municipalityName')
        cadastral_unit_number = cadastral_unit.get('cadastralUnitNumber')
        property_unit_number = cadastral_unit.get('propertyUnitNumber')
        adresse = f"Gnr. {cadastral_unit_number} Bnr. {property_unit_number}"
        post_nummer = cadastral_unit.get('postalCode')

    # Hent fylke basert på kommunenavn
    fylke = finn_fylke(municipality)

    # Hent koordinater
    coordinates = work_order_address.get("coordinates", {})
    coordsys = coordinates.get("system")
    x_koordinat = coordinates.get("x")
    y_koordinat = coordinates.get("y")

    return adresse, post_nummer, municipality, fylke, coordsys, x_koordinat, y_koordinat

# Kundenavn og kontaktinfo
def extract_contact_info(entry):
    """
    Ekstraherer kontaktinformasjon (navn og telefonnummer) fra en work order entry.

    :param entry: Dictionary som inneholder work order data.
    :return: Tuple med (kunde_navn, telefon_nr).
    """
    # Henter kontaktinformasjon fra 'detailedOrderInformation' hvis tilgjengelig
    contact_details = entry.get("detailedOrderInformation", {}).get("user1", {})
    if contact_details:
        contact_persons = contact_details.get("contactPersons", [])
    else:
        contact_persons = entry.get("buyer", {}).get("contactPersons", [])

    # Sjekk om det finnes kontaktpersoner
    if contact_persons:
        kunde_navn = f"{contact_persons[0].get('firstName', '')} {contact_persons[0].get('familyName', '')}".strip()
        telefon_nr = contact_persons[0].get("phone1", "").strip()
    else:
        kunde_navn = "Ukjent"
        telefon_nr = "Ukjent"

    return kunde_navn, telefon_nr

# Sambandsnummer
def extract_service_details(entry, item):
    """
    Ekstraherer sambandsnummer og tilgjengelige ressurser fra detailedOrderInformation.

    :param entry: Dictionary som inneholder ordredata.
    :param item: Valgfritt navn eller ID for logging hvis ingen sambandsnummer finnes.
    :return: Tuple med (sambandsnummer, available_resources).
    """
    service_details = entry.get("detailedOrderInformation", {}).get("serviceDetails", [])
    sambandsnummer = None
    available_resources = []  # Liste for å logge ressurstyper

    if isinstance(service_details, list):
        # PRIORITET 1: Finn første CircuitId
        circuit_ids = [
            d.get("resourceId", "").strip()
            for d in service_details
            if d.get("resourceType", "").strip().lower() == "circuitid"
               and d.get("resourceId")
        ]
        if circuit_ids:
            sambandsnummer = circuit_ids[0]
        else:
            # PRIORITET 2: Finn første CustomerId
            customer_ids = [
                d.get("resourceId", "").strip()
                for d in service_details
                if d.get("resourceType", "").strip().lower() == "customerid"
                   and d.get("resourceId")
            ]
            if customer_ids:
                sambandsnummer = customer_ids[0]

        # Samle andre ressurstyper i available_resources
        available_resources = [
            f"{d.get('resourceType', '').strip()}: {d.get('resourceId', '').strip()}"
            for d in service_details
            if d.get("resourceType", "").strip().lower() not in ["circuitid", "customerid"]
        ]

        # Legg til DG i available_resources istedenfor å bruke det som sambandsnummer
        dg_resources = [
            f"DG: {d.get('resourceId', '').strip()}"
            for d in service_details
            if d.get("resourceType", "").strip().lower() == "dg"
               and d.get("resourceId")
        ]
        available_resources.extend(dg_resources)  # Legg til DG i listen

    # Logging hvis ingen sambandsnummer er funnet
    if not sambandsnummer:
        print(f"\nIngen sambandsnummer funnet for: {item}")
        print("Tilgjengelige nummer:", available_resources)

    return sambandsnummer, available_resources

# LU-Nummer
def extract_lu_number(entry):
    """
    Ekstraherer LU-nummer fra serviceDetails i detailedOrderInformation.

    :param entry: Dictionary som inneholder ordredata.
    :return: LU-nummer (str) eller None hvis ingen finnes.
    """
    service_details = entry.get("detailedOrderInformation", {}).get("serviceDetails", [])

    if isinstance(service_details, list):
        for detail in service_details:
            if detail.get("resourceType") == "LU":
                return detail.get("resourceId", "").strip()  # Returnerer første LU som dukker opp

    return None  # Returnerer None hvis ingen LU-nummer finnes

# Spidernummer
def extract_spidernumber(entry):
    """
    Ekstraherer Spidernummer fra dependentWorkOrders.

    :param entry: Dictionary som inneholder ordredata.
    :return: Spidernummer (str) eller None hvis ingen finnes.
    """
    dependant_work_orders = entry.get("dependentWorkOrders")

    if isinstance(dependant_work_orders, list):
        for dep_order in dependant_work_orders:
            return dep_order.get("workOrderId", "").strip()  # Returnerer første funnet Spidernummer

    return None  # Returnerer None hvis ingen Spidernummer finnes

# Finn alle produkt-ID
def extract_product_ids(entry):
    """
    Ekstraherer alle Produkt-IDer fra orderlines.

    :param entry: Dictionary som inneholder ordredata.
    :return: Liste med Produkt-IDer (list) eller tom liste hvis ingen finnes.
    """
    orderlines_list = entry.get("orderlines", [{}])
    orderlines_productId = []

    if isinstance(orderlines_list, list):
        for line in orderlines_list:
            cp_id = line.get("contractorProductId")
            if cp_id:  # Sjekk at verdien ikke er None eller tom
                orderlines_productId.append(cp_id)

    return orderlines_productId

# Finne produktkode med høyest prioritet
def get_highest_priority_product(product_codes, woc_type_oppdrag):
    try:
        # Les inn Excel-filen
        df = pd.read_excel("Datafiler/WOC_Prioritering_Produktkategorier.xlsx", engine="openpyxl")

        # Filtrer etter de spesifikke produktkodene
        df_filtered = df[df['Produktkode'].isin(product_codes)]

        # Finn raden med høyest prioritet (lavest tall i "Prioritering")
        highest_priority_row = df_filtered.loc[df_filtered['Prioritering'].idxmin()]

        return highest_priority_row['Produktkode'] + ": " + highest_priority_row['Produkt']
    except Exception as e:
        print(f"Produkt ikke i liste: {e}")
        if woc_type_oppdrag:
            return woc_type_oppdrag[0]
        else:
            return product_codes[0]

# Type oppdrag fra WOC
def extract_woc_type_oppdrag(entry):
    """
    Ekstraherer type oppdrag fra WOC basert på orderlines.

    :param entry: Dictionary som inneholder ordredata.
    :return: Liste med type oppdrag (list) eller tom liste hvis ingen finnes.
    """
    orderlines_list = entry.get("orderlines", [])
    woc_type_oppdrag = []

    if isinstance(orderlines_list, list):
        for line in orderlines_list:
            typop_id = line.get("description")
            if typop_id and line.get("isMainProduct") is True:
                woc_type_oppdrag.append(typop_id)

    return woc_type_oppdrag

# Type oppdrag til Monday
def determine_type_oppdrag(orderinfo_description, VULA_nr, prioritert_product_id):
    """
    Bestemmer type oppdrag basert på ordrebeskrivelse, leveransestatus, WOC-type oppdrag og VULA-referanser.

    :param orderinfo_description: Ordrebeskrivelse (str).
    :param VULA_nr: VULA-referansenummer (str eller list).
    :return: Type oppdrag (str) eller None hvis ingen kriterier er oppfylt.
    """
    if orderinfo_description == "BB_ACCESS":
        return "BB-Access"
    elif "LVA1A" in prioritert_product_id: #status_leveranse == "NY FTTH" or
        return "Komplett fortetning"
    elif "LVK0" in prioritert_product_id:
        return "Eksperthjelpen"
    elif "LVK2F" in prioritert_product_id:
        return "Installasjonshjelpen"
    elif VULA_nr == "VULA":
        return "VULA"
    elif VULA_nr == "VULA CDK":
        return "VULA CDK"
    elif "LVT2D" in prioritert_product_id:
        return "AEG"
    elif "LVT1C" in prioritert_product_id:
        return "Leveranse timer - Fiber"
    elif "DLS99" in prioritert_product_id:
        return "DLS99"
    return None

# Beskrivelse av produktet
def extract_product_descriptions(entry):
    """
    Ekstraherer produktbeskrivelser fra serviceDetails i detailedOrderInformation.

    :param entry: Dictionary som inneholder ordredata.
    :return: Liste med produktbeskrivelser (list) eller tom liste hvis ingen finnes.
    """
    service_details = entry.get("detailedOrderInformation", {}).get("serviceDetails", [])
    product_description = []

    if isinstance(service_details, list):
        for line in service_details:
            prod_dscr = line.get("productDescription")
            if prod_dscr:  # Sikrer at None-verdier ikke legges til
                product_description.append(prod_dscr)

    return product_description

# VULA
def extract_vula_numbers(entry, item):
    """
    Ekstraherer VULA-referansenummer fra externalOrderReferences.

    :param entry: Dictionary som inneholder ordredata.
    :param item: Valgfritt navn eller ID for logging hvis ingen VULA-nummer finnes.
    :return: Liste med VULA-referansenummer eller tom liste hvis ingen finnes.
    """
    VULA_list = entry.get("externalOrderReferences", [])
    VULA_nr = []

    if isinstance(VULA_list, list):
        for line in VULA_list:
            ref_Num = line.get("referenceNumber")
            if ref_Num and ("VULA" == ref_Num or "VULA CDK" == ref_Num or "VULA" in ref_Num):
                VULA_nr.append(ref_Num)
    else:
        ref_Num = entry.get("externalOrderReferences", {}).get("referenceNumber")
        if ref_Num:
            VULA_nr.append(ref_Num)
        print(f'\nSjekk om {item} er VULA')

    return VULA_nr

# BEDRIFT eller PRIVAT
def determine_oppdrag_kategori(customer_category, contract_details):
    """
    Bestemmer oppdragkategori basert på kunde- og kontraktsdetaljer.

    :param customer_category: Kundekategori (str) or None.
    :param contract_details: Kontraktsdetaljer (str) or None.
    :return: Oppdragkategori (str) - 'bedrift', 'privat' eller 'denne må sjekkes'.
    """
    # Ensure values are strings to avoid AttributeError
    customer_category = (customer_category or "").lower().strip()
    contract_details = (contract_details or "").lower().strip()

    if any(keyword in customer_category for keyword in ["bedrift", "fttb"]) or any(keyword in contract_details for keyword in ["bedrift", "fttb"]):
        return "bedrift"
    elif "privat" in customer_category or "ftth" in contract_details:
        return "privat"
    else:
        print("Sjekk om denne er bedrift eller privat")
        return "denne må sjekkes"


# Status Leveranse
def determine_status_leveranse(orderlines_productId, spidernummer, orderinfo_description, contract_details, gpon_p2p_woc, VULA_nr, oppdrag_kategori, item):
    if oppdrag_kategori == "bedrift":
        if any("LVLU" in L_nr for L_nr in orderlines_productId) or spidernummer:
            return "Booket"
        elif VULA_nr:
            return "NY wholesale"
        elif orderinfo_description == "BB_ACCESS":
            return "NY FWA"
        elif contract_details == "FTTB" or gpon_p2p_woc == "HELIOS":
            return "NY bedrift"
        else:
            print(f"{item} - Bedriftsoppdrag leveransekode må sjekkes")
            return None
    elif oppdrag_kategori == "privat":
        if any(product_id in orderlines_productId for product_id in ["LVA1A", "LVA1B", "LVA1D", "LVA2F"]):
            return "NY FTTH"
        elif VULA_nr:
            return "NY wholesale"
        elif orderinfo_description == "BB_ACCESS":
            return "NY FWA"
        else : #contract_details == "AEG":
            return "NY privat"
        # else:
        #     print(f"{item} - Privatoppdrag leveransekode må sjekkes")
        #     return None
    else:
        print(f"{item} - Oppdrag kategori må sjekkes")
        return None


# Type FTTx til monday
def determine_fttx(orderlines_productId, spidernummer, contract_details, gpon_p2p_woc, orderinfo_description,status_leveranse, oppdrag_kategori, item):
    """
    Bestemmer FTTx-type basert på ordredata.

    :param orderlines_productId: Liste over produkt-IDer fra orderlines.
    :param spidernummer: Spidernummer (kan være None eller str).
    :param contract_details: Kontraktsdetaljer (str).
    :param gpon_p2p_woc: GPON eller P2P WOC-type (str).
    :param orderinfo_description: Ordrebeskrivelse (str).
    :param status_leveranse: Status for leveranse (str).
    :return: FTTx-type (str) eller None hvis ingen kriterier er oppfylt.
    """
    if oppdrag_kategori == "bedrift":
        if any("LVLU" in L_nr for L_nr in orderlines_productId) or spidernummer:
            return "FTTB Onnet"
        elif contract_details == "FTTB" or gpon_p2p_woc == "HELIOS" or orderinfo_description == "BB_ACCESS":
            return "FTTB Offnet"
        else:
            print(f"{item}: Mangler FTTx")
            return None
    elif oppdrag_kategori == "privat":
        if status_leveranse in ["NY wholesale", "NY FTTH"] or contract_details == "AEG":
            return "FTTH Fortetning"
        else:
            print(f"{item}: Mangler FTTx")
            return "FTTH Fortetning" #None
            # denne må nok utviddes etterhvert for å ta skille på FTTH MDU, GF og SDU
    else:
        print(f"{item}: Mangler FTTx")
        return None



# GPON/P2P
def determine_gpon_p2p(status_leveranse, VULA_nr, gpon_p2p_woc):
    """
    Bestemmer GPON/P2P-type basert på leveransestatus, VULA-referanser og WOC-type.

    :param status_leveranse: Status for leveranse (str).
    :param VULA_nr: VULA-referansenummer (str eller list).
    :param gpon_p2p_woc: WOC-type (str).
    :return: GPON/P2P-type (str) eller None hvis ingen kriterier er oppfylt.
    """
    if status_leveranse == "NY FWA":
        return "Antenne"
    elif status_leveranse == "NY FTTH":
        return "FTTH"
    elif VULA_nr or gpon_p2p_woc == "GPON":
        return "GPON"
    elif gpon_p2p_woc == "LEIDE SAMBAND":
        return "P2P"
    elif status_leveranse == "NY privat":
        return "AEG"
    elif gpon_p2p_woc == "NORDIC CONNECT":
        return "Ruterbytte"

    return None



# Funksjon for å formatere datoer til yyyy-mm-dd
def format_date(date_str):
    if date_str:
        try:
            # Håndterer datoformat med tidssone (f.eks. 2025-01-15T08:00:00+01:00)
            return datetime.fromisoformat(date_str).strftime("%Y-%m-%d")
        except ValueError:
            return date_str  # Returner uendret hvis formatet ikke stemmer
    return None

# Ordredato
def get_latest_accept_workorder_date(entries):
    latest_date = None

    for entry in entries:
        if entry.get('action') == 'AcceptWorkOrder':
            date_str = entry.get('changed')
            if date_str:
                date_obj = datetime.strptime(date_str, "%Y-%m-%dT%H:%M:%S.%fZ")
                if latest_date is None or date_obj > latest_date:
                    latest_date = date_obj

    return latest_date.strftime("%Y-%m-%d") if latest_date else None





def extract_data_from_json(json_data):
    extracted_data = []

    for entry in json_data:
        # Hopp over rader hvor supplier.contactPersons ikke er tom
        if entry.get("supplier", {}).get("contactPersons"):
            continue
        elif not entry.get("wocOrderStatus", {}).lower() in ['accepted', 'received']:
            continue


        # Finne Item-navn
        item = extract_item(entry)

        # Håndterer WorkOrderAddress som liste
        addresse, post_nummer, municipality, fylke, coordsys, x_koordinat, y_koordinat = extract_work_order_details(entry)
        fylke_status = fylke #for å lage en statuskolonne i Monday, kun for importboardet og automation

        # Kundenavn og kontaktinfo
        kunde_navn, telefon_nr = extract_contact_info(entry)

        # Datoer
        issued_date = entry.get("issuedDate").split("T")[0]
        bookes_innen = add_working_days_with_holidays(issued_date, 5)
        start_arbeid_tidligst = format_date(entry.get("deliveryPeriod", {}).get("startDate"))
        ordre_dato = get_latest_accept_workorder_date(entry.get("activityLog"))
        if not ordre_dato:
            print(item)
            ordre_dato = issued_date
        last_transaction_date = format_date(entry.get("modifiedDate"))
        dato_leveranse = format_date(entry.get("deliveryPeriod", {}).get("endDate"))

        # Ordrenummer
        ordrenr_leveranse = entry.get("clientOrderId", {}).get("referenceNumber")

        # Sambandsnummer - Håndterer serviceDetails
        sambandsnummer, available_resources = extract_service_details(entry, item)

        # LU-nummer
        LU_nummer = extract_lu_number(entry)

        # Spidernummer
        spidernummer = extract_spidernumber(entry)

        # Type oppdrag fra WOC
        woc_type_oppdrag = extract_woc_type_oppdrag(entry)

        # Legger til Produkt-ID
        orderlines_productId = extract_product_ids(entry)

        # Finner Produkt-ID med høyeste prioritet
        prioritert_product_id = get_highest_priority_product(orderlines_productId, woc_type_oppdrag)

        orderinfo_description = entry.get("detailedOrderInformation", {}).get("orderDescription")

        # Beskrivelse av produktet
        product_description = extract_product_descriptions(entry)

        # VULA
        VULA_nr = extract_vula_numbers(entry, item)

        # Statusnummer til Monday
        contract_details = entry.get("contract", {}).get("detailedPurchaseArea")

        # Customer Category
        customer_category = entry.get("detailedOrderInformation", {}).get("customerCategory")


        # Oppdragkategori, BEDRIFT eller PRIVAT
        oppdrag_kategori = determine_oppdrag_kategori(customer_category, contract_details)

        # Status Leveranse
        gpon_p2p_woc = entry.get("areaOfSubject")
        gpon_from_detailed = entry.get("detailedAreaOfSubject")
        status_leveranse = determine_status_leveranse(orderlines_productId, spidernummer, orderinfo_description,
                                                      contract_details, gpon_p2p_woc, VULA_nr, oppdrag_kategori, item)

        # Type oppdrag til Monday
        type_oppdrag = determine_type_oppdrag(orderinfo_description, VULA_nr, prioritert_product_id)

        # Type FTTx til Monday
        FTTx = determine_fttx(orderlines_productId, spidernummer, contract_details, gpon_p2p_woc, orderinfo_description, status_leveranse, oppdrag_kategori, item)

        ### GPON/P2P til Monday
        gpon_p2p = determine_gpon_p2p(status_leveranse, VULA_nr, gpon_p2p_woc)

        # Legge til underentreprenør
        under_entreprenor = finn_entreprenor(post_nummer)

        # Statuser
        woc_status = None
        bc_status = None
        ue_status = None
        woc_connector = "WOC"



        extracted_data.append([
            item,
            addresse,
            municipality,
            fylke,
            kunde_navn,
            telefon_nr,
            issued_date,
            bookes_innen,
            ordre_dato,
            under_entreprenor,
            fylke_status,
            start_arbeid_tidligst,
            dato_leveranse,
            woc_status,
            bc_status,
            ue_status,
            ordrenr_leveranse,
            sambandsnummer,
            woc_connector,
            spidernummer,
            status_leveranse,
            contract_details,
            FTTx,
            oppdrag_kategori,
            gpon_p2p,
            gpon_p2p_woc,
            gpon_from_detailed,
            type_oppdrag,
            woc_type_oppdrag,
            LU_nummer,
            last_transaction_date,
            prioritert_product_id,
            orderlines_productId,
            VULA_nr,
            product_description,
            coordsys,
            x_koordinat,
            y_koordinat,
            orderinfo_description,
            customer_category
        ])

    return extracted_data


columns = [
    "Item",
    "Adresse",
    "Kommune",
    "Fylke",
    "Kunde",
    "Telefon",
    "Issued Date",
    "Bookes innen",
    "Ordredato",
    "Entreprenør",
    "Fylke Status",
    "Start arbeid tidligst",
    "Dato leveranse",
    "WOC Status",
    "BC Status",
    "UE Status",
    "Ordrenummer",
    "Sambandsnummer",
    "WOC/connector",
    "Spidernummer",
    "Status Leveranse",
    "kontraktdetaljer",
    "Type FTTx",
    "Kunde Kategori",
    "GPON/P2P",
    "GPON/P2P - WOC",
    "GPON/AEG - from detailedAreaOfSubject",
    "Type oppdrag",
    "Type oppdrag WOC",
    "LU-nummer",
    "Last Transaction Date",
    "Hovedprodukt", #Prioritert produkt
    "Produkt ID",
    "VULA ?",
    "Beskrivelse av produkt",
    "Coordsys",
    "X-koordinat",
    "Y-koordinat",
    "Orderinfo Description",
    "Customer Category"
]

with open(json_file_path, "r", encoding="utf-8") as file:
    json_data = json.load(file)

rows = extract_data_from_json(json_data)



# Save the Excel file safely
df = pd.DataFrame(rows, columns=columns)
with pd.ExcelWriter(target_excel_file, engine="openpyxl", mode="w") as writer:
    df.to_excel(writer, index=False, sheet_name="Data")

split_excel_by_customer_category(target_excel_file)
