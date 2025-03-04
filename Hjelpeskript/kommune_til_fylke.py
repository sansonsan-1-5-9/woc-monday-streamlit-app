import pandas as pd

# Laste inn Excel-filen
file_path = "Datafiler/Kommune_Fylke_Oversikt.xlsx"
df = pd.read_excel(file_path, sheet_name="Ark1")

# Rense kolonnenavn
df.columns = ["Fylkesnavn", "Fylkesnr", "Kommunenavn", "Kommunenr", "Kommunenr_2023"]


# Funksjon for å finne fylke basert på kommunenavn
def finn_fylke(kommune_navn):
    resultat = df[df["Kommunenavn"].str.lower() == kommune_navn.lower()]

    if not resultat.empty:
        return resultat["Fylkesnavn"].values[0]
    else:
        print(f"Fant ikke Fylke for kommune: {kommune_navn}")
        return None



