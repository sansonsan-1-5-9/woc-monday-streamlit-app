import pandas as pd


def split_excel_by_customer_category(input_file):
    """
    Leser en Excel-fil og splitter den i to filer basert på verdien i 'Kunde Kategori'.

    - Rader med 'privat' lagres i 'Monday_Import - P.xlsx'
    - Rader med 'bedrift' lagres i 'Monday_Import - B.xlsx'

    Args:
        input_file (str): Filsti til den opprinnelige Excel-filen.

    Returns:
        None
    """
    # Definer utfilene
    output_file_priv = "C:/Users/OdinSanson/PycharmProjects/Python_Telenor-woc/Monday_Import - P.xlsx"
    output_file_bedrift = "C:/Users/OdinSanson/PycharmProjects/Python_Telenor-woc/Monday_Import - B.xlsx"

    # Les Excel-filen
    df = pd.read_excel(input_file)

    # Fjern eventuelle ledende eller etterfølgende mellomrom i kolonnenavn
    df.columns = df.columns.str.strip()

    # Sjekk om nødvendig kolonne eksisterer
    if "Kunde Kategori" not in df.columns:
        raise ValueError("Kolonnen 'Kunde Kategori' finnes ikke i Excel-filen.")

    # Filtrer dataene
    df_priv = df[df["Kunde Kategori"] == "privat"]
    df_bedrift = df[df["Kunde Kategori"] == "bedrift"]

    # Lagre til nye Excel-filer
    df_priv.to_excel(output_file_priv, index=False)
    df_bedrift.to_excel(output_file_bedrift, index=False)

    print(f"Filene '{output_file_priv}' og '{output_file_bedrift}' er opprettet!")



