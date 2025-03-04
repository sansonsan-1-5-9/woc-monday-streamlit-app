import pandas as pd

# Laste inn Excel-filen
file_path = "Datafiler/Kommune_Fylke_Oversikt.xlsx"
df = pd.read_excel(file_path, sheet_name="Postnummerregister")

# Rense kolonnenavn
df.columns = ["Postnummer","Poststed","Kommunenummer","Kommunenavn","Kategori"]


# Funksjon for å finne fylke basert på kommunenavn
def finn_kommune(poststed_navn):
    resultat = df[df["Poststed"].str.lower() == poststed_navn.lower()]

    if not resultat.empty:
        return resultat["Kommunenavn"].values[0]
    else:
        print(f"Fant ikke kommune for poststed: {poststed_navn}")
        return None

def finn_kommune_fra_postnr(poststnr):
    resultat = df[df["Postnummer"] == poststnr]

    if not resultat.empty:
        return resultat["Kommunenavn"].values[0]
    else:
        print(f"Fant ikke kommune for poststed: {poststnr}")
        return None

#print(f"Kommune: {finn_kommune_fra_postnr(3285)}")
# kommune_input = input("Skriv inn kommunenavn: ")
# print(f"Kommune: {finn_kommune(kommune_input)}")
