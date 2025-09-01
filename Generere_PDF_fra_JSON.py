import json
import sys
import os
import shutil
import pandas as pd
from fpdf import FPDF
from datetime import datetime
from Hjelpeskript.kommune_til_fylke import finn_fylke

# Filstier
# Use the uploaded JSON file passed from Streamlit
if len(sys.argv) > 1:
    json_file_path = sys.argv[1]
else:
    raise FileNotFoundError("No JSON file provided to the script.")

output_folder = "generated_pdfs/"

# Slett eksisterende innhold i mappen hvis den finnes
if os.path.exists(output_folder):
    for filename in os.listdir(output_folder):
        file_path = os.path.join(output_folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)  # Slett fil eller symbolsk lenke
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)  # Slett mappe og alt innhold
        except Exception as e:
            print(f'Feil ved sletting av {file_path}: {e}')

# Opprett mappe for PDF-filer hvis den ikke finnes
os.makedirs(output_folder, exist_ok=True)

def format_date(iso_string):
    """Hjelpefunksjon for å formattere ISO8601-datoer til f.eks. DD.MM.YYYY."""
    if not iso_string:
        return ""
    try:
        dt = datetime.fromisoformat(iso_string.replace("Z", "+00:00"))
        return dt.strftime("%d.%m.%Y %H:%M")
    except ValueError:
        return iso_string  # fallback dersom dato ikke kan parses

def format_date_dd_mm_yyyy(iso_string):
    """Hjelpefunksjon for å formattere ISO8601-datoer til f.eks. DD.MM.YYYY."""
    if not iso_string:
        return ""
    try:
        dt = datetime.fromisoformat(iso_string.replace("Z", "+00:00"))
        return dt.strftime("%d.%m.%Y")  # Endret for kun å vise dato, uten klokkeslett
    except ValueError:
        return iso_string  # fallback dersom dato ikke kan parses

def start_section_if_room(pdf, title, needed_height):
    """
    Sjekker om det er nok plass på siden for hele seksjonen.
    Hvis ikke, legger vi til en ny side før vi skriver tittelen.
    needed_height = antatt høyde i punkter (f.eks. linjer * linjehøyde).
    """
    # PDF har en indre grense for auto page-break;
    # vi kan lese den fra pdf.page_break_trigger
    # pdf.get_y() gir oss nåværende “cursor”-posisjon i høyden.
    if pdf.get_y() + needed_height >= pdf.page_break_trigger:
        pdf.add_page()

    # Nå skriver vi selve seksjonstittelen
    pdf.section_title(title)
    pdf.ln(2)  # litt spacing under tittel

# Finne Produktbeskrivelse fra ID
def find_product_description(produkt_id):
    df = pd.read_excel("Datafiler/WOC_Prioritering_Produktkategorier.xlsx", engine="openpyxl")

    resultat = df[df["Produktkode"].str.lower() == produkt_id.lower()]

    if not resultat.empty:
        return resultat["Produkt"].values[0]
    else:
        print(f"Fant ikke produktbeskrivelse for: {produkt_id}")
        return None


class CustomPDF(FPDF):
    def header(self):
        # Lys grå topplinje
        self.set_fill_color(230, 230, 230)
        self.rect(0, 0, self.w, 15, 'F')

        # Skriv dagens dato og klokkeslett (venstre hjørne)
        self.set_xy(5, 5)  # x=5, y=5 for litt margin
        now_str = datetime.now().strftime("%d.%m.%Y, %H:%M")
        self.set_font("Arial", "", 10)
        self.cell(0, 0, now_str, ln=0, align="L")

        # Skriv overskrift (midtstilt)
        self.set_y(3)
        self.set_font("Arial", "B", 12)
        self.cell(0, 10, "WOC 2 Work Order", 0, 0, "C")

        self.ln(11)  # Flytt "cursor" litt nedover

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", size=8)
        # Legg inn alias {nb}
        self.cell(0, 10, f"Side {self.page_no()}/{{nb}}", 0, 0, "R")

    def section_title(self, title):
        """Hjelpefunksjon for seksjons-overskrift."""
        self.set_font("Arial", "B", 11)
        self.set_fill_color(200, 200, 200)  # litt mørkere grå
        self.cell(0, 8, title, ln=True, align="L", fill=True)

# Les JSON-filen
with open(json_file_path, "r", encoding="utf-8") as file:
    data = json.load(file)

# Generer PDF-er basert på JSON-data
for entry in data:

    # Hopp over rader hvor supplier.contactPersons ikke er tom
    if entry.get("supplier", {}).get("contactPersons"):
        continue
    elif not entry.get("wocOrderStatus", {}).lower() in ['accepted', 'received', 'appointed']:
        continue

    ### Bygg en filnavn-vennlig tittel
    client_order_id = entry["clientOrderId"]["referenceNumber"]
    short_street = ""
    short_house = ""
    address_info = entry.get("workOrderAddress", [])
    if address_info:
        # Tar første "workOrderAddress" hvis den finnes
        street_data = address_info[0]["streetAddress"]
        if street_data:
            short_street = street_data.get("streetName", "")
            short_house = street_data.get("houseNumber", "")
            house_char = street_data.get("houseChar", "")
            if house_char:
                short_house = short_house+house_char
            adresse = f"{short_street} {short_house}"
            municipality_name = street_data.get("municipalityName")
        else:
            work_order_address = entry.get("workOrderAddress", [{}])[0]
            cadastral_unit = work_order_address.get('cadastralUnit', {})
            municipality_name = cadastral_unit.get('municipalityName')
            cadastral_unit_number = cadastral_unit.get('cadastralUnitNumber')
            property_unit_number = cadastral_unit.get('propertyUnitNumber')
            adresse = f"Gnr. {cadastral_unit_number} Bnr. {property_unit_number}"
    else:
        adresse = "finner ikke addresse"
        municipality_name = "Mangler info"

    ### Lage filsti
    fylke = finn_fylke(municipality_name)
    order_type = entry.get("orderType", "")
    print(order_type)
    output_folder = f"generated_pdfs/{fylke}/{order_type}"
    os.makedirs(output_folder, exist_ok=True)

    pdf_filename = f"{client_order_id} {adresse}.pdf"
    pdf_filepath = os.path.join(output_folder, pdf_filename)

    pdf = CustomPDF()
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()

    # ----------------------------------------------------
    # 1) Overskrift / Tittel-linje
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, entry.get("title", ""), ln=True)


    ### Undertekst

    # 1) OrderId = (workOrderId.referenceName)-(workOrderId.referenceNumber)
    work_order_id_ref = entry["workOrderId"]
    order_id_str = f"{work_order_id_ref['referenceName']}-{work_order_id_ref['referenceNumber']}"

    # 2) ClientOrderId = (clientOrderId.referenceName)-(clientOrderId.referenceNumber)
    client_order_id_ref = entry["clientOrderId"]
    client_order_id_str = f"{client_order_id_ref['referenceName']}-{client_order_id_ref['referenceNumber']}"

    # 3) CustomerCategory (detailedOrderInformation.customerCategory)
    customer_category = entry["detailedOrderInformation"].get("customerCategory", "")
    if not customer_category:
        customer_category = "-"

    # 4) CircuitId: finn resourceId i serviceDetails med resourceType="CircuitId"
    circuit_id = ""
    customer_id = ""
    sd_list = entry["detailedOrderInformation"].get("serviceDetails", [])

    if sd_list:
        for sd in sd_list:
            if sd.get("resourceType") == "CircuitId":
                circuit_id = sd.get("resourceId", "")
                break

        # 5) CustomerId: finn resourceId i serviceDetails med resourceType="CustomerId"
        customer_id = ""
        for sd in sd_list:
            if sd.get("resourceType") == "CustomerId":
                customer_id = sd.get("resourceId", "")
                break

    # 6) OrderType
    order_type = entry.get("orderType", "")

    # 7) Area: areaOfSubject
    area = entry.get("areaOfSubject", "")

    # 8) Contract: f.eks. contractType + contractSegment
    contract_type = entry["contract"].get("contractType", "")
    contract_segment = entry["contract"].get("contractSegment", "")
    contract_str = f"{contract_type} {contract_segment}".strip()

    # 9) PurchaseArea: (contract.purchaseArea) (contract.priceRegion)
    purchase_area = entry["contract"].get("purchaseArea", "")
    price_region = entry["contract"].get("priceRegion", "")
    purchase_str = f"{purchase_area} {price_region}".strip()

    # 10) Issued: issuedDate
    issued_str = format_date(entry.get("issuedDate"))

    # 11) Modified: modifiedDate
    modified_str = format_date(entry.get("modifiedDate"))

    # 12) State => wocOrderStatus
    woc_status = entry.get("wocOrderStatus", "")

    # 13) State (CWM): (orderStatus)
    workorder_status = entry.get("orderStatus", "").capitalize()

    # # Undertekst
    pdf.set_font("Arial", "", 8)
    # Definer faste kolonnebredder
    col1_width = 60
    col2_width = 60
    col3_width = 60

    # Genererer og skriver ut hver rad i riktig format med faste kolonnebredder
    pdf.cell(col1_width, 5, f"OrderId: {order_id_str}", border=0)
    pdf.cell(col2_width, 5, f"OrderType: {order_type}", border=0)
    pdf.cell(col3_width, 5, f"Issued: {issued_str}", border=0)
    pdf.ln()

    pdf.cell(col1_width, 5, f"ClientOrderId: {client_order_id_str}", border=0)
    pdf.cell(col2_width, 5, f"Area: {area}", border=0)
    pdf.cell(col3_width, 5, f"Modified: {modified_str}", border=0)
    pdf.ln()

    pdf.cell(col1_width, 5, f"CustomerCategory: {customer_category}", border=0)
    pdf.cell(col2_width, 5, f"Contract: {contract_str}", border=0)
    pdf.cell(col3_width, 5, f"State: {woc_status}", border=0)
    pdf.ln()

    pdf.cell(col1_width, 5, f"CircuitId: {circuit_id}", border=0)
    pdf.cell(col2_width, 5, f"PurchaseArea: {purchase_str}", border=0)
    pdf.cell(col3_width, 5, f"State (CWO): {workorder_status}", border=0)
    pdf.ln()

    pdf.cell(col1_width, 5, f"CustomerId: {customer_id}", border=0)
    pdf.ln()

    # Legg til ekstra mellomrom for bedre lesbarhet
    pdf.ln(3)

    pdf.ln(3)

    # ----------------------------------------------------
    # 2) Kontakt-info (Buyer / Supplier / Contact)
    # Header: "Contact Info"
    #pdf.set_font("Arial", "B", 12)
    pdf.section_title( "Contact Info")
    pdf.ln(2)  # Liten spacing

    buyer = entry.get("buyer", {})
    supplier = entry.get("supplier", {})
    contact_details = entry.get("detailedOrderInformation", {}).get("user1", {})
    if contact_details:
        contact_persons = contact_details.get("contactPersons", [])
    else:
        contact_persons = buyer.get("contactPersons", [])

    # Buyer-seksjon
    pdf.set_font("Arial", "B", 8)  # Bold for "Buyer"
    pdf.cell(0, 3, "Buyer", ln=True)

    pdf.set_font("Arial", "", 10)  # Normal font for detaljer
    pdf.cell(0, 4, buyer.get("companyName", ""), ln=True)
    pdf.cell(0, 4, f"Org.nr: {buyer.get('businessRegistrationNumber', '')}", ln=True)
    pdf.ln(2)  # Ekstra spacing

    # Supplier-seksjon
    pdf.set_font("Arial", "B", 8)  # Bold for "Supplier"
    pdf.cell(0, 3, "Supplier", ln=True)

    pdf.set_font("Arial", "", 10)  # Normal font for detaljer
    pdf.cell(0, 4, supplier.get("companyName", ""), ln=True)
    pdf.cell(0, 4, f"Org.nr: {supplier.get('businessRegistrationNumber', '')}", ln=True)
    pdf.ln(2)  # Ekstra spacing

    # Contact-seksjon
    pdf.set_font("Arial", "B", 8)  # Bold for "Supplier"
    pdf.cell(0, 3, "Contacts", ln=True)

    pdf.set_font("Arial", "", 10)  # Normal font for detaljer
    pdf.cell(0, 4, f"Name: {contact_persons[0].get('firstName', '')} {contact_persons[0].get('familyName', '')}".strip(), ln=True)
    pdf.cell(0, 4, f"Role: {contact_persons[0].get('role', '')}", ln=True)
    pdf.cell(0, 4, f"Phone: {(contact_persons[0].get('phone1') or 'Ukjent').strip() if isinstance(contact_persons[0].get('phone1'), str) else 'Ukjent'}", ln=True)
    pdf.cell(0, 4, f"Email: {contact_persons[0].get('email', '')}", ln=True)
    pdf.ln(2)

    #ISP-seksjon
    pdf.set_font("Arial", "B", 8)  # Bold for "Supplier"
    pdf.cell(0, 3, "ISP", ln=True)

    pdf.set_font("Arial", "", 10)  # Normal font for detaljer
    isp = entry['detailedOrderInformation'].get('isp', [])
    if not isp:
        isp_name = ''
    else:
        isp_name = isp.get('fullName', '')
    pdf.cell(0, 4, f"{isp_name}", ln=True)

    pdf.ln(4)  # Ekstra spacing for å skille fra neste seksjon


    # ----------------------------------------------------
    # 3) WorkOrder Address (gateadresse, poststed osv.)
    pdf.section_title("WorkOrder Address")
    pdf.ln(2)
    user1 = entry.get("detailedOrderInformation", {}).get("user1", {})
    if user1:
        address_data = user1.get("address", {})
        if address_data:
            street_data = address_data.get("streetAddress", {})
    work_order_address = entry.get("workOrderAddress", [{}])[0]
    pdf.set_font("Arial", "", 10)


    if street_data:
        # Overskrift "User1"
        pdf.set_font("Arial", "B", 10)
        pdf.cell(0, 3, "User1", ln=True)

        pdf.set_font("Arial", "", 10)

        if user1:
            org_id = user1.get("organizationId", "-")
            full_name = user1.get("fullName", "")

        pdf.cell(0, 4, f"Organisation Id: {org_id}", ln=True)
        pdf.cell(0, 4, f"FullName: {full_name}", ln=True)

        municipality_number = street_data.get("municipalityNumber", "")
        municipality_name = street_data.get("municipalityName", "")
        county_number = street_data.get("countyNumber", "")
        street_name = street_data.get("streetName", "")
        street_code = street_data.get("streetCode", "")
        house_number = street_data.get("houseNumber", "")
        floor_number = street_data.get("floorNumber", "")
        apartment_number = street_data.get("apartmentNumber", "")
        postal_code = street_data.get("postalCode", "")
        city = street_data.get("city", "")

        pdf.cell(0, 4, f"MunicipalityNumber: {municipality_number}", ln=True)
        pdf.cell(0, 4, f"MunicipalityName: {municipality_name}", ln=True)
        pdf.cell(0, 4, f"CountyNumber: {county_number}", ln=True)
        pdf.cell(0, 4, f"StreetName: {street_name}", ln=True)
        pdf.cell(0, 4, f"StreetCode: {street_code}", ln=True)
        pdf.cell(0, 4, f"HouseNumber: {house_number}", ln=True)
        pdf.cell(0, 4, f"FloorNumber: {floor_number}", ln=True)
        pdf.cell(0, 4, f"ApartmentNumber: {apartment_number}", ln=True)
        pdf.cell(0, 4, f"PostalCode: {postal_code}", ln=True)
        pdf.cell(0, 4, f"City: {city}", ln=True)

        # Koordinater
        coordinates = street_data.get("coordinates", {})
        if coordinates:
            system = coordinates.get("system", "")
            x_val = coordinates.get("x", "")
            y_val = coordinates.get("y", "")

            pdf.cell(0, 5, f"Coordinates ({system}):", ln=True)
            pdf.cell(0, 4, f"    X: {x_val}", ln=True)
            pdf.cell(0, 4, f"    Y: {y_val}", ln=True)
    elif work_order_address:


        street_address = work_order_address.get("streetAddress", {})

        if street_address:
            municipality_number = street_address.get("municipalityNumber","")
            municipality_name = street_address.get("municipalityName","")
            county_number = street_address.get("countyNumber","")
            street_name = street_address.get("streetName", "")
            street_code = street_address.get("streetCode")
            house_number = street_address.get("houseNumber", "")
            house_char = street_address.get("houseChar", "")
            if house_char:
                house_number += house_char
            floor_number = street_address.get("floorNumber", "")
            apartment_number = street_address.get("apartmentNumber", "")
            postal_code = street_address.get("postalCode", "")
            city = street_address.get("city", "")

            pdf.cell(0, 4, f"MunicipalityNumber: {municipality_number}", ln=True)
            pdf.cell(0, 4, f"MunicipalityName: {municipality_name}", ln=True)
            pdf.cell(0, 4, f"CountyNumber: {county_number}", ln=True)
            pdf.cell(0, 4, f"StreetName: {street_name}", ln=True)
            pdf.cell(0, 4, f"StreetCode: {street_code}", ln=True)
            pdf.cell(0, 4, f"HouseNumber: {house_number}", ln=True)
            if floor_number:
                pdf.cell(0, 4, f"FloorNumber: {floor_number}", ln=True)
            if apartment_number:
                pdf.cell(0, 4, f"ApartmentNumber: {apartment_number}", ln=True)
            pdf.cell(0, 4, f"PostalCode: {postal_code}", ln=True)
            pdf.cell(0, 4, f"City: {city}", ln=True)

            # Koordinater
            coordinates = street_address.get("coordinates", {})
            system = coordinates.get("system", "")
            x_val = coordinates.get("x", "")
            y_val = coordinates.get("y", "")

            pdf.cell(0, 5, f"Coordinates ({system}):", ln=True)
            pdf.cell(0, 4, f"    X: {x_val}", ln=True)
            pdf.cell(0, 4, f"    Y: {y_val}", ln=True)

        elif work_order_address.get('cadastralUnit', {}):

            cadastral_unit = work_order_address.get('cadastralUnit', {})
            municipality_number = cadastral_unit.get('municipalityNumber')
            municipality_name = cadastral_unit.get("municipalityName", "")
            county_number = cadastral_unit.get("countyNumber", "")
            postal_code = cadastral_unit.get("postalCode", "")
            city = cadastral_unit.get("city", "")
            cadastral_unit_number = cadastral_unit.get('cadastralUnitNumber')
            property_unit_number = cadastral_unit.get('propertyUnitNumber')
            leasehold_number = cadastral_unit.get('leaseholdNumber')
            condominimum_unit_number = cadastral_unit.get('CondominiumUnitNumbe')

            pdf.cell(0, 4, f"MunicipalityNumber: {municipality_number}", ln=True)
            pdf.cell(0, 4, f"MunicipalityName: {municipality_name}", ln=True)
            pdf.cell(0, 4, f"CountyNumber: {county_number}", ln=True)
            pdf.cell(0, 4, f"PostalCode: {postal_code}", ln=True)
            pdf.cell(0, 4, f"City: {city}", ln=True)
            pdf.cell(0, 4, f"CadastralUnitNumber: {cadastral_unit_number}", ln=True)
            pdf.cell(0, 4, f"PropertyUnitNumber: {property_unit_number}", ln=True)
            pdf.cell(0, 4, f"LeaseholdNumber: {leasehold_number}", ln=True)
            pdf.cell(0, 4, f"CondominiumUnitNumber: {condominimum_unit_number}", ln=True)

            # Koordinater
            coordinates = cadastral_unit.get("coordinates", {})
            system = coordinates.get("system", "")
            x_val = coordinates.get("x", "")
            y_val = coordinates.get("y", "")

            pdf.cell(0, 5, f"Coordinates ({system}):", ln=True)
            pdf.cell(0, 4, f"    X: {x_val}", ln=True)
            pdf.cell(0, 4, f"    Y: {y_val}", ln=True)

        else:
            pdf.cell(0, 3, ("Finner ikke adresse fra fil"), ln=True)

    else:
        pdf.cell(0, 3, "Finner ikke adresse i fil", ln=True)


    pdf.ln(4)

    # ----------------------------------------------------
    # 4) Order Information / Delivery Dates
    start_section_if_room(pdf,"Order Information", 50)

    # -- Description --
    pdf.set_font("Arial", "B", 8)  # Bold, mindre font
    pdf.cell(0, 3, "Description", ln=True)

    pdf.set_font("Arial", "", 10)  # Normal font for tekst
    order_description = entry.get("detailedOrderInformation", {}).get("orderDescription", "")
    for line in order_description.split("\n"):
        pdf.multi_cell(0, 4, line)#, ln=True)
    pdf.ln(2)  # Ekstra spacing

    # -- ConnectionPoint --
    pdf.set_font("Arial", "B", 8)  # Fet, liten font
    pdf.cell(0, 3, "ConnectionPoint", ln=True)

    pdf.set_font("Arial", "", 10)

    connection_point = entry.get("connectionPoint")
    if connection_point is None:
        # Hvis connectionPoint er null i JSON:
        pdf.cell(0, 4, "Ingen ConnectionPoint oppgitt", ln=True)
    else:
        conn_id = connection_point.get("id", "-")
        pdf.cell(0, 4, f"Id: {conn_id}", ln=True)
        if connection_point.get("fullName"):
            conn_full_name = connection_point.get("fullName")
            pdf.cell(0, 4, f"FullName: {conn_full_name}", ln=True)
        if connection_point.get("remark"):
            conn_remark = connection_point.get("remark")

            pdf.multi_cell(0, 4, conn_remark)



    pdf.ln(4)  # Litt ekstra luft før neste seksjon

    # ----------------------------------------------------
    # 5.1) Delivery dates
    start_section_if_room(pdf, "Delivery dates", 30)

    delivery_period = entry.get("deliveryPeriod", {})
    pdf.set_font("Arial", "", 10)  # Normal font for tekst

    pdf.cell(0, 4, f"Start: {format_date_dd_mm_yyyy(delivery_period.get('startDate'))}", ln=True)
    pdf.cell(0, 4, f"Planning completion: {format_date_dd_mm_yyyy(delivery_period.get('planningCompletedDate'))}", ln=True)
    pdf.cell(0, 4, f"Acceptance: {format_date_dd_mm_yyyy(delivery_period.get('acceptanceDate'))}", ln=True)
    pdf.cell(0, 4, f"End: {format_date_dd_mm_yyyy(delivery_period.get('endDate'))}", ln=True)
    pdf.cell(0, 4, f"Ad: {format_date_dd_mm_yyyy(delivery_period.get('adDate'))}", ln=True)

    pdf.ln(4)  # Litt ekstra luft før neste seksjon

    # ----------------------------------------------------
    # 5.2) Appointment
    customer_appointment = entry.get("customerAppointment", {})
    if customer_appointment:
        start_section_if_room(pdf, "Appointment", 30)

        customer_appointment = entry.get("customerAppointment", {})
        pdf.set_font("Arial", "", 10)  # Normal font for tekst

        pdf.cell(0, 4, f"Type: {customer_appointment.get('type')}", ln=True)
        pdf.cell(0, 4, f"From: {format_date_dd_mm_yyyy(customer_appointment.get('fromTime'))}",ln=True)
        pdf.cell(0, 4, f"To: {format_date_dd_mm_yyyy(customer_appointment.get('toTime'))}", ln=True)

        pdf.ln(4)  # Litt ekstra luft før neste seksjon

    # ----------------------------------------------------
    # 6) Service details (tabell)

    service_details = entry["detailedOrderInformation"].get("serviceDetails", [])

    if service_details:

        start_section_if_room(pdf, "Service details", 50)

        for sd in service_details:
            resource_type = sd.get("resourceType", "")

            if resource_type:
                # Lag en seksjonstittel med selve resourceType
                pdf.set_font("Arial", "B", 8)
                pdf.cell(0, 3, resource_type, ln=True)

                pdf.set_font("Arial", "", 10)
                if sd.get('resourceId', ''):
                    pdf.cell(0, 4, f"ResourceId: {sd.get('resourceId', '')}", ln=True)
                pdf.cell(0, 4, f"ResourceType: {resource_type}", ln=True)
                if sd.get('productDescription', ''):
                    pdf.cell(0, 4, f"ProductDescription: {sd.get('productDescription', '')}", ln=True)
                if sd.get('action', ''):
                    pdf.cell(0, 4, f"Action: {sd.get('action', '')}", ln=True)
                if sd.get('speedDown', ''):
                    pdf.cell(0, 4, f"SpeedDown: {sd.get('speedDown', '')}", ln=True)
                if sd.get('speedUp', ''):
                    pdf.cell(0, 4, f"SpeedUp: {sd.get('speedUp', '')}", ln=True)
                pdf.cell(0, 4, f"SpeedDownReduced: {sd.get('speedDownReduced', '')}", ln=True)
                pdf.ln(2)

            # Hvis resource_type er tom strenger etc., ignorerer vi den:
            else:
                pass

        pdf.ln(4)

    # ----------------------------------------------------
    # 7.1) Dependent Work Orders
    dependent_orders = entry.get("dependentWorkOrders", [])
    if dependent_orders:
        # -- Header / seksjon --
        pdf.section_title("DependentWorkOrders")
        pdf.ln(2)

        # -- Kolonne-overskrifter (Bold, 8) --
        pdf.set_font("Arial", "B", 8)
        pdf.cell(25, 4, "WorkorderId", ln=False)
        pdf.cell(30, 4, "ContractorName", ln=False)
        pdf.cell(30, 4, "ContactPerson", ln=False)
        pdf.cell(15, 4, "Role", ln=False)
        pdf.cell(20, 4, "Phone", ln=False)
        pdf.cell(25, 4, "Email", ln=False)
        pdf.cell(35, 4, "PreferredContactChannel", ln=True)  # ln=True = ny linje

        # -- Data-rader (Arial, 10) --
        pdf.set_font("Arial", "", 10)
        for dw in dependent_orders:
            workorder_id = dw.get("workOrderId", "")
            contractor_name = dw.get("contractorName", "")
            contact_person = dw.get("contactPerson")

            # Hvis contactPerson er None, lager vi tomme strenger
            if contact_person:
                role = contact_person.get("role", "")
                phone = contact_person.get("phone1", "")
                email = contact_person.get("email", "")
                preferred = contact_person.get("preferredContactChannel", "")
                # Hvis du vil vise evt. fornavn/etternavn, kan du hente det slik:
                first_name = contact_person.get("firstName", "")
                family_name = contact_person.get("familyName", "")
                contact_name = f"{first_name} {family_name}".strip()
            else:
                # Tom info
                role = phone = email = preferred = contact_name = ""

            # Skriv én rad i tabellen
            pdf.cell(25, 4, workorder_id, ln=False)
            pdf.cell(30, 4, contractor_name or "", ln=False)
            pdf.cell(30, 4, contact_name, ln=False)
            pdf.cell(15, 4, role, ln=False)
            pdf.cell(20, 4, phone, ln=False)
            pdf.cell(25, 4, email, ln=False)
            pdf.cell(35, 4, preferred, ln=True)

        pdf.ln(4)  # Litt ekstra spacing etter tabellen

    # ----------------------------------------------------
    # 7.2) Additional Information
    additional_info = entry.get("detailedOrderInformation", {}).get("additionalInformation", [])
    if additional_info:
        # Header for Additional Information
        pdf.section_title("Additional Information")
        pdf.ln(2)  # Litt spacing

        # Eksempel: Itererer over elementene i additionalInfo-listen
        for info_item in additional_info:
            # F.eks. "Scope information for HP delivery"
            description = info_item.get("description", "")

            # Bold, fontsize=8
            pdf.set_font("Arial", "B", 8)
            pdf.cell(0, 3, description, ln=True)

            # Normal tekst for characteristics
            pdf.set_font("Arial", "", 10)

            # Gå gjennom alle "characteristics"
            characteristics = info_item.get("characteristics", [])
            for c in characteristics:
                name_ = c.get("name", "")
                value_ = c.get("value", "")
                # Eksempel: "HP: Homes Passed, bay=false i Smallworld"
                pdf.cell(0, 4, f"{name_}: {value_}", ln=True)

            pdf.ln(2)  # Avslutt hver info-blokk med litt spacing

        pdf.ln(2)  # Litt ekstra luft før neste seksjon

    # ----------------------------------------------------
    # 7.3) Remarks
    remarks = entry.get("remarks", [])
    if remarks:
        pdf.section_title("Remarks")
        pdf.ln(2)  # Litt mellomrom

        for rm in remarks:
            # Hent initiator, dato og tekst
            initiator = rm.get("initiator", "")
            created_str = format_date(rm.get("createdDate"))  # f.eks. 29.01.2025
            text = rm.get("text", "")

            # Linje 1: (BUYER) 29.01.2025 (bold, fontsize=8)
            pdf.set_font("Arial", "B", 8)
            pdf.cell(0, 3, f"({initiator}) {created_str}", ln=True)

            # Linje 2 (og ev. flere): Remarks‐tekst i normal font
            pdf.set_font("Arial", "", 6)
            # For en liten "bullet", kan vi bruke "* " eller "• "
            pdf.multi_cell(0, 4, f"    - {text}")

            pdf.ln(2)  # Litt luft mellom hver remark

        pdf.ln(2)  # Ekstra spacing før neste seksjon

    # ----------------------------------------------------
    # 9) CPE
    cpe_list = entry.get("detailedOrderInformation", {}).get("cpe", [])
    if cpe_list:
        pdf.section_title("CPE")
        pdf.ln(2)  # Liten spacing

        # Tynn, lys grå strek
        pdf.set_draw_color(200, 200, 200)

        # Overskriftsrad: Bold, font=8, med bunngrenser (border="B")
        pdf.set_font("Arial", "B", 8)
        pdf.cell(60, 5, "Name", border="B", ln=False)
        pdf.cell(60, 5, "Serial Number", border="B", ln=False)
        pdf.cell(40, 5, "On-site pairing", border="B", ln=True)

        # Data-rader: Normal font (Arial 10), med bunngrenser (border="B") for hver rad
        pdf.set_font("Arial", "", 10)
        for cpe_item in cpe_list:
            name = cpe_item.get("name", "")
            serial_number = cpe_item.get("serialNumber", "") or ""
            # Hvis onSitePairing=True → "X", ellers tom
            pairing_check = "X" if cpe_item.get("onSitePairing") else ""

            pdf.cell(60, 5, name, border="B", ln=False)
            pdf.cell(60, 5, serial_number, border="B", ln=False)
            pdf.cell(40, 5, pairing_check, border="B", ln=True)

        pdf.ln(4)  # Ekstra luft før neste seksjon

    # ----------------------------------------------------
    # 10) ExternalOrderReferences
    ext_refs = entry.get("externalOrderReferences", [])
    if ext_refs:
        pdf.section_title("ExternalOrderReferences")
        pdf.ln(2)  # Liten spacing

        # Sett farge for tynne, lysegrå linjer
        pdf.set_draw_color(200, 200, 200)

        # Header-rad (bold, fontsize=8) med bunn-strek (border="B")
        pdf.set_font("Arial", "B", 8)
        pdf.cell(60, 5, "ReferenceName", border="B", ln=False)
        pdf.cell(100, 5, "ReferenceNumber", border="B", ln=True)

        # Rader (normal font, fontsize=10) med kun horisontal strek (border="B")
        pdf.set_font("Arial", "", 10)
        for ref in ext_refs:
            ref_name = ref.get("referenceName", "")
            ref_number = ref.get("referenceNumber", "")

            pdf.cell(60, 5, ref_name, border="B", ln=False)
            pdf.cell(100, 5, ref_number, border="B", ln=True)

        pdf.ln(4)  # Ekstra luft før neste seksjon

    # ----------------------------------------------------
    # 11) Orderlines
    order_lines = entry.get("orderlines", [])
    if order_lines:
        pdf.section_title("Orderlines")
        pdf.ln(2)  # Liten spacing

        # Sorter radene ut fra lineNumber
        order_lines_sorted = sorted(order_lines, key=lambda x: x.get("lineNumber", 0))

        # Tynn, lys grå strek
        pdf.set_draw_color(200, 200, 200)

        # Header-rad (bold, font=8), kun bunnlinje (border="B") for horisontal strek
        pdf.set_font("Arial", "B", 8)
        pdf.cell(15, 5, "LineNo", border="B", ln=False)
        pdf.cell(25, 5, "ProductId", border="B", ln=False)
        pdf.cell(80, 5, "ProductName", border="B", ln=False)
        pdf.cell(30, 5, "Quantity", border="B", ln=False)
        pdf.cell(40, 5, "Project", border="B", ln=True)

        # Data-rader i normal font (Arial 10)
        pdf.set_font("Arial", "", 10)

        for line in order_lines_sorted:
            line_no = str(line.get("lineNumber", ""))
            product_id = line.get("contractorProductId", "")
            # Bruker 'description' som "ProductName"
            if line.get("description"):
                product_name = line.get("description") or ""
            else:
                product_name = str(find_product_description(product_id) or "")
            qty = str(line.get("quantity", 0))  # <-- Konvertert til str
            unit = line.get("unitOfMeasure", "")
            qty_with_unit = f"{qty} {unit}".strip()  # Fjern unødvendig space hvis unit mangler
            project_code = line.get("project")
            if project_code:
                project_code = line.get("project", {}).get("projectCode") or ""

            # Estimer hvor mange linjer `product_name` vil ta opp i en 80-bredde kolonne
            max_width = 80  # Max bredde for ProductName
            text_width = pdf.get_string_width(product_name)
            line_count = max(1, int(text_width / max_width) + 1)
            row_height = line_count * 5  # Justerer høyden basert på linjetallet

            pdf.cell(15, row_height, line_no, border="B", ln=False)
            pdf.cell(25, row_height, product_id, border="B", ln=False)

            # Multicell for `ProductName` for å unngå overlapping
            x, y = pdf.get_x(), pdf.get_y()
            pdf.multi_cell(80, 5, product_name, border="B")
            pdf.set_xy(x + 80, y)  # Sett tilbake X-koordinaten etter MultiCell


            pdf.cell(30, row_height, qty_with_unit, border="B", ln=False)
            pdf.cell(40, row_height, str(project_code), border="B", ln=True)

        pdf.ln(4)  # Litt ekstra spacing før neste seksjon

    # ----------------------------------------------------
    # ----------------------------------------------------
    # ----------------------------------------------------


    ### Lagre PDF
    pdf.output(pdf_filepath)


print(f"PDF-filer er generert i mappen: {output_folder}")


