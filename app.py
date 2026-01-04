
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, NamedStyle
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from decimal import Decimal, ROUND_HALF_UP

# ======= KONFIGŪRA =======
PVM_DEFAULT = Decimal("21.00")
FMT_MONEY = "#,##0.00"
FMT_QTY = "#,##0.00"

ROW_TABLE_START = 9          # Paslaugų lentelės pradžia
MAX_LINES = 40               # Kiek eilučių paruošti su DV (keisk, jei reikia)

def huf(value: str) -> Decimal:
    """Saugi konversija į Decimal su lietuvišku kableliu ir apvalinimu HALF_UP."""
    return (Decimal(str(value).replace(",", "."))
            .quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

def create_named_styles(wb):
    if "Money" not in wb.named_styles:
        st = NamedStyle(name="Money"); st.number_format = FMT_MONEY
        st.alignment = Alignment(horizontal="right"); wb.add_named_style(st)
    if "Qty" not in wb.named_styles:
        st = NamedStyle(name="Qty"); st.number_format = FMT_QTY
        st.alignment = Alignment(horizontal="right"); wb.add_named_style(st)
    if "Label" not in wb.named_styles:
        st = NamedStyle(name="Label"); st.font = Font(bold=True)
        st.alignment = Alignment(horizontal="left"); wb.add_named_style(st)

def set_borders(ws, rng, thick=False):
    side = Side(style="thick" if thick else "thin")
    for row in ws[rng]:
        for c in row:
            c.border = Border(left=side, right=side, top=side, bottom=side)

def build_workbook(header, catalog_rows):
    """
    Sukuria pilną AKTAS darbo knygą.

    header: dict su laukais:
        {'uzsakovas','vykdytojas','sutartis','skyrius','akto_data','metai_atlik':'2026','adresas':''}
        'adresas' paliekamas tuščias (bus parenkamas iš dropdown META!B4)
    catalog_rows: list[dict] su laukais:
        [{'address','service','rate','qty'}]
    """
    wb = Workbook()
    create_named_styles(wb)

    # ----- KATALOGAS (duombazė) -----
    wsK = wb.active
    wsK.title = "KATALOGAS"
    wsK.append(["Adresas", "Paslauga", "Įkainis (be PVM)", "Numatytas kiekis"])
    for r in catalog_rows:
        wsK.append([
            r["address"],
            r["service"],
            float(huf(r["rate"])),
            float(huf(r.get("qty", "0")))
        ])

    # ----- META (žali kintamieji/antraštės) -----
    wsM = wb.create_sheet("META")
    wsM["A1"], wsM["B1"] = "Užsakovas", header.get("uzsakovas","")
    wsM["A2"], wsM["B2"] = "Vykdytojas", header.get("vykdytojas","")
    wsM["A3"], wsM["B3"] = "Sutarties nr.", header.get("sutartis","")
    wsM["A4"], wsM["B4"] = "Objektas / adresas", header.get("adresas","")  # pasirenkama iš DV
    wsM["A5"], wsM["B5"] = "Skyrius", header.get("skyrius","")
    wsM["A6"], wsM["B6"] = "Akto data", header.get("akto_data","")
    wsM["A7"], wsM["B7"] = "Atlikimo metai", header.get("metai_atlik","2026")

    # Pavadinimas pasirinktai adreso reikšmei (patogu formulėms)
    wb.defined_names.append(DefinedName(name="AdrSelected", attr_text="'META'!$B$4"))

    # ----- LISTOS (paieškos laukai + spill) -----
    wsL = wb.create_sheet("LISTOS")
    # Adresų paieška
    wsL["A1"] = "Adresų paieška:"; wsL["B1"] = ""  # vartotojas rašo paieškos tekstą čia
    wsL["A2"] = "=UNIQUE(KATALOGAS!A2:A100000)"   # unikalūs adresai (spill A2#)
    wsL["C1"] = "Filtruoti adresai (DV)"
    wsL["C2"] = "=FILTER(A2#, ISNUMBER(SEARCH(B1, A2#)), \"Nėra atitikmenų\")"

    # Paslaugų paieška
    wsL["E1"] = "Paslaugų paieška:"; wsL["F1"] = ""  # vartotojas rašo paiešką čia
    wsL["E2"] = "=FILTER(KATALOGAS!B2:B100000, KATALOGAS!A2:A100000=AdrSelected, \"Nėra\")"
    wsL["F2"] = "=FILTER(E2#, ISNUMBER(SEARCH(F1, E2#)), \"Nėra\")"

    # Vardiniai pavadinimai DV šaltiniams (spill)
    wb.defined_names.append(DefinedName(name="AdresaiDV",   attr_text="'LISTOS'!$C$2#"))
    wb.defined_names.append(DefinedName(name="PaslaugosDV", attr_text="'LISTOS'!$F$2#"))

    # ----- AKTAS (vartotojui) -----
    ws = wb.create_sheet("AKTAS")
    # A1–A6 „Etiketė: Reikšmė“ iš META
    labels = ["Užsakovas", "Vykdytojas", "Sutarties nr.", "Objektas / adresas", "Skyrius", "Akto data"]
    for i, lab in enumerate(labels, start=1):
        ws[f"A{i}"] = f'="{lab}: " & META!B{i}'
    # A7 – atlikimo data su metais ir „m.“ (mėnesį įrašysi ranka)
    ws["A7"] = '="Atlikimo data: " & META!B7 & " m. "'

    # Lentelės antraštės
    ws["A8"] = "Eil. Nr."; ws["B8"] = "Paslaugos pavadinimas"; ws["C8"] = "Kiekis"
    ws["D8"] = "Įkainis (be PVM)"; ws["E8"] = "Suma (be PVM)"
    set_borders(ws, "A8:E8", thick=True)

    # Paslaugų dropdown + formulės
    dv_service = DataValidation(type="list", formula1="=PaslaugosDV", allow_blank=True)
    ws.add_data_validation(dv_service)
    dv_nonneg = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
    ws.add_data_validation(dv_nonneg)

    for idx in range(MAX_LINES):
        r = ROW_TABLE_START + idx
        # Eil. Nr.
        ws.cell(r, 1).value = idx + 1

        # B = Paslauga (dropdown)
        dv_service.add(ws.cell(r, 2))

        # C = Kiekis (AUTOMATIŠKAI iš KATALOGO pagal pasirinktą paslaugą ir adresą)
        # Jei paslauga nepersirinkta – paliks tuščią; vartotojas gali perrašyti ranka.
        ws.cell(r, 3).value = (
            f'=IFERROR(INDEX(FILTER(KATALOGAS!D2:D100000,'
            f'(KATALOGAS!B2:B100000=B{r})*(KIALOGAS!A2:A100000=AdrSelected)),1),"")'
        ).replace("KIALOGAS", "KATALOGAS")  # apsauga nuo klaidos kopijuojant
        ws.cell(r, 3).style = "Qty"
        dv_nonneg.add(ws.cell(r, 3))

        # D = Įkainis pagal pasirinktą B ir pasirinktą adresą (pirmas atitikimas KATALOGE)
        ws.cell(r, 4).value = (
            f'=IFERROR(INDEX(FILTER(KATALOGAS!C2:C100000,'
            f'(KATALOGAS!B2:B100000=B{r})*(KATALOGAS!A2:A100000=AdrSelected)),1),"")'
        )
        ws.cell(r, 4).style = "Money"
        dv_nonneg.add(ws.cell(r, 4))

        # E = C * D (horizontalios formulės)
        ws.cell(r, 5).value = f"=C{r}*D{r}"
        ws.cell(r, 5).style = "Money"

    set_borders(ws, f"A{ROW_TABLE_START}:E{ROW_TABLE_START+MAX_LINES-1}")

    # Sumų blokas (dešinėje) – kol kas rodo visą MAX_LINES diapazoną
    ws["D12"] = "Suma (be PVM):"; ws["E12"] = f"=SUM(E{ROW_TABLE_START}:E{ROW_TABLE_START+MAX_LINES-1})"
    ws["D13"] = f"PVM {float(PVM_DEFAULT)}%:"; ws["E13"] = f"=E12*{float(PVM_DEFAULT)/100}"
    ws["D14"] = "Suma su PVM:"; ws["E14"] = "=E12+E13"
    for c in ("E12","E13","E14"):
        ws[c].style = "Money"
    set_borders(ws, "D12:E14", thick=True)

    # Data Validation adresui (META!B4)
    dv_addr = DataValidation(type="list", formula1="=AdresaiDV", allow_blank=False)
    wsM.add_data_validation(dv_addr)
    dv_addr.add(wsM["B4"])

    # (nebūtina) galima paslėpti techninius lapus:
    # wsM.sheet_state = "hidden"; wsL.sheet_state = "hidden"; wsK.sheet_state = "hidden"

    return wb


def clean_before_sending(filename_in: str, filename_out: str = None):
    """
    Atidaro esamą xlsx, randa paskutinę užpildytą paslaugų eilutę,
    ištrina TUŠČIAS MAX_LINES eilutes nuo galo ir perrašo sumų formules
    tik į realiai užpildytą diapazoną.
    """
    wb = load_workbook(filename_in)
    ws = wb["AKTAS"]

    start = ROW_TABLE_START
    end_prepared = ROW_TABLE_START + MAX_LINES - 1

    # Paskutinė užpildyta eilutė = turi bent vieną reikšmę B/C/D/E
    last_used = start - 1
    for r in range(start, end_prepared + 1):
        if any(ws.cell(row=r, column=c).value not in (None, "") for c in (2, 3, 4, 5)):
            last_used = r

    # Jei niekas neužpildyta – paliekam tik antraštes
    if last_used < start:
        last_used = start - 1

    # Ištrinam tuščias eilutes žemiau paskutinės užpildytos
    rows_to_delete = end_prepared - last_used
    if rows_to_delete > 0:
        ws.delete_rows(last_used + 1, rows_to_delete)

    # Perrašom sumų formules pagal realų diapazoną
    if last_used >= start:
        ws["E12"].value = f"=SUM(E{start}:E{last_used})"
    else:
        ws["E12"].value = "0"
    ws["E13"].value = f"=E12*{float(PVM_DEFAULT)/100}"
    ws["E14"].value = "=E12+E13"

    # (neprivaloma) galime sutraukti ribas iki realaus korpuso
    # set_borders(ws, f"A{start}:E{last_used if last_used>=start else start}")

    if not filename_out:
        filename_out = filename_in.replace(".xlsx", "_clean.xlsx")
    wb.save(filename_out)
    return filename_out


# ======= PALEIDIMAS (pavyzdys) =======
if __name__ == "__main__":
    header = {
        "uzsakovas": "ANYKŠČIŲ RAJONO SAVIVALDYBĖS ADMINISTRACIJA",
        "vykdytojas": "Corpus A, UAB",
        "sutartis": "6-793/CA-224154",
        "adresas": "",                 # pasirinks iš dropdown META!B4
        "skyrius": "ANA.P.A.J",
        "akto_data": "2026-01-04",
        "metai_atlik": "2026",
    }

    # Pavyzdiniai katalogo įrašai — keisk į savo realius
    catalog = [
        {"address":"J. Biliūno g. 19, Anykščiai","service":"Langų valymo paslauga valant iš abiejų pusių","rate":"2,00","qty":"50,38"},
        {"address":"J. Biliūno g. 19, Anykščiai","service":"Durų rankenų dezinfekavimas","rate":"0,50","qty":"30,00"},
        {"address":"Vilniaus g. 1, Anykščiai","service":"Grindų plovimas","rate":"1,20","qty":"100,00"},
        {"address":"Vilniaus g. 1, Anykščiai","service":"Sanitarinių mazgų valymas","rate":"3,50","qty":"10,00"},
    ]

    # 1) Generuojame aktą su dropdown + paieška + automatiniais kiekiais/įkainiais
    wb = build_workbook(header, catalog)
    base_file = "aktas_dropdown_paieška.xlsx"
    wb.save(base_file)

    # 2) Kai užpildei paslaugas (Excel’e), prieš siuntimą:
    #    paleisk clean_before_sending(base_file) — jis uždarys tuščias eilutes ir perrašys sumų diapazoną.
    # output_file = clean_before_sending(base_file)
    # print("Išvalytas failas:", output_file)
``
