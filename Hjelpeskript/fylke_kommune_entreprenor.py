import pandas as pd

# Laste inn Excel-filen
file_path = "C:/Users/OdinSanson/PycharmProjects/Python_Telenor-woc/Datafiler/Fordeling_Entreprenor.xlsx"
df = pd.read_excel(file_path, sheet_name="Postnummerregister")

# Rense kolonnenavn
df.columns = ["Fylke", "Kommunenummer", "Kommunenavn", "Postnummer", "Poststed", "Entreprenør"]

# Sikre at 'Postnummer' er en streng
df["Postnummer"] = df["Postnummer"].astype(str)

# Funksjon for å finne entreprenør basert på postnummer
def finn_entreprenor(post_nummer):
    post_nummer = str(post_nummer).strip()  # Sikre at input er en streng og fjern unødvendige mellomrom
    resultat = df[df["Postnummer"] == post_nummer]  # Direkte sammenligning fungerer bedre

    if not resultat.empty:
        return resultat["Entreprenør"].values[0]
    else:
        print(f"Fant ikke entreprenør for postnr. {post_nummer}")
        return None


