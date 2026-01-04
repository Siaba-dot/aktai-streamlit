import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime

# Tarkime, tavo duomenys yra sąraše žodynų formatu
data = [
    {
        "Užsakovas": "Įmonė A",
        "Tiekėjas": "Tiekėjas X",
        "Sutarties numeris": "123",
        "Objekto adresas": "Vilnius, Gedimino pr. 1",
        "Skyrius": "Skyrius 1",
        "Atlikimo data": "2026-01-04",
        "Atlikimo laikotarpis": "2026-01 mėn.",
        "Paslaugos pavadinimas": "Valymo paslauga",
        "Kiekis": 10,
        "Įkainis": 15,
        "Suma": 150
    },
    # Čia galima pridėti daugiau eilučių
]

# Excel failui paruošti naudojame pandas DataFrame
df_header = pd.DataFrame([
    {
        "Užsakovas": row.get("Užsakovas", ""),
        "Tiekėjas": row.get("Tiekėjas", ""),
        "Sutarties numeris": row.get("Sutarties numeris", ""),
        "Objekto adresas": row.get("Objekto adresas", ""),
        "Skyrius": row.get("Skyrius", ""),
        "Atlikimo data": row.get("Atlikimo data", ""),
        "Atlikimo laikotarpis": row.get("Atlikimo laikotarpis", ""),
    } for row in data
])

df_table = pd.DataFrame([
    {
        "Paslaugos pavadinimas": row.get("Paslaugos pavadinimas", ""),
        "Kiekis": row.get("Kiekis", 0),
        "Įkainis": row.get("Įkainis", 0),
        "Suma": row.get("Suma", 0)
    } for row in data
])

# Sukuriame Excel failą
wb = Workbook()
ws = wb.active
ws.title = "Sąskaita"

# Pridedame header
for r in dataframe_to_rows(df_header, index=False, header=True):
    ws.append(r)

# Tarpo eilutė tarp header ir lentelės
ws.append([])

# Pridedame lentelę
for r in dataframe_to_rows(df_table, index=False, header=True):
    ws.append(r)

# Išsaugome failą
excel_filename = f"saskaita_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
wb.save(excel_filename)
print(f"Excel failas sukurtas: {excel_filename}")
