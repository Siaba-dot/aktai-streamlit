def write_act_to_sheet(ws, sheet_name, meta, items, pvm_pct, show_pvm):
    """
    ws: Excel sheet objektas (pvz., xlwt.Workbook().add_sheet)
    meta: dict su header info (Užsakovas, Tiekėjas ir pan.)
    items: list of dict arba pandas Series su lentelės duomenimis
    pvm_pct: PVM procentas
    show_pvm: ar rodyti PVM stulpelį
    """

    # Pradžia rašyti nuo 0 eilutės
    start = 0

    # HEADER
    header_fields = [
        "Užsakovas",
        "Tiekėjas",
        "Sutarties numeris",
        "Objekto adresas",
        "Skyrius",
        "Atlikimo data",
        "Atlikimo laikotarpis"
    ]

    for col, field in enumerate(header_fields):
        value = meta.get(field, "")
        ws.write(start, col, value)
    start += 2  # paliekame vieną tarpo eilutę

    # LENTELĖ
    table_fields = ["Paslaugos pavadinimas", "Kiekis", "Įkainis", "Suma"]
    if show_pvm:
        table_fields.append("PVM")

    # Parašome lentelės header
    for col, field in enumerate(table_fields):
        ws.write(start, col, field)
    start += 1

    # Parašome kiekvieną eilutę
    for i, row in enumerate(items):
        for col, field in enumerate(table_fields):
            # Saugi prieiga prie reikšmių: veiks ir su dict, ir su pandas Series
            if isinstance(row, dict):
                value = row.get(field, "")
            else:  # pandas Series
                value = row[field] if field in row else ""
            
            # Jei reikia PVM stulpelio ir jis nerastas, pridedame
            if field == "PVM" and show_pvm and not value:
                value = float(row.get("Suma", 0)) * pvm_pct / 100

            ws.write(start + i, col, value)
