
# -*- coding: utf-8 -*-
import io
from typing import List, Dict, Optional
from decimal import Decimal, ROUND_HALF_UP

import streamlit as st
import pandas as pd

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, NamedStyle
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

# ======= KONFIGŪRA =======
PVM_DEFAULT = Decimal("21.00")
FMT_MONEY = "#,##0.00"
FMT_QTY = "#,##0.00"
ROW_TABLE_START = 9
MAX_LINES = 40
SEP = ","  # jei tavo Excel reikalauja kabliataškio, keisk į ";"

SHEET_DB = "KATALOGAS"
SHEET_META = "META"
SHEET_LIST = "LISTOS"
SHEET_AKT = "AKTAS"

# ======= PAGALBINIAI =======
def huf(value: str) -> Decimal:
    """Decimal su HALF_UP apvalinimu; palaiko lietuvišką kablelį."""
    return (Decimal(str(value).replace(",", "."))
            .quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

def create_named_styles(wb: Workbook) -> None:
    if "Money" not in wb.named_styles:
        stl = NamedStyle(name="Money"); stl.number_format = FMT_MONEY
        stl.alignment = Alignment(horizontal="right"); wb.add_named_style(stl)
    if "Qty" not in wb.named_styles:
        stl = NamedStyle(name="Qty"); stl.number_format = FMT_QTY
        stl.alignment = Alignment(horizontal="right"); wb.add_named_style(stl)

def set_borders(ws, rng: str, thick: bool = False) -> None:
    side = Side(style="thick" if thick else "thin")
    for row in ws[rng]:
        for c in row:
            c.border = Border(left=side, right=side, top=side, bottom=side)

def autosize(ws) -> None:
    for col in ws.columns:
        max_len = 0
        letter = col[0].column_letter
        for cell in col:
            val = str(cell.value) if cell.value is not None else ""
            max_len = max(max_len, len(val))
        ws.column_dimensions[letter].width = min(max_len + 2, 60)

def f_FILTER(range_txt: str, condition_txt: str, ifempty_txt: str) -> str:
    return f"=FILTER({range_txt}{SEP}{condition_txt}{SEP}{ifempty_txt})"

def f_SEARCH(needle: str, haystack: str) -> str:
    return f"ISNUMBER(SEARCH({needle}{SEP}{haystack}))"

# ======= EXCEL GENERAVIMAS =======
def build_workbook(
    header: Dict[str, str],
    catalog_rows: List[Dict[str, str]],
    max_lines: int = MAX_LINES,
    preset_address: Optional[str] = None,
    hide_helper_sheets: bool = False
) -> Workbook:
    wb = Workbook()
    create_named_styles(wb)

    # --- KATALOGAS ---
    wsK = wb.active; wsK.title = SHEET_DB
    wsK.append(["Adresas", "Paslauga", "Įkainis (be PVM)", "Numatytas kiekis"])
    for r in catalog_rows:
        wsK.append([r["address"], r["service"], float(huf(r["rate"])), float(huf(r.get("qty","0")))])
    autosize(wsK)

    # --- META ---
    wsM = wb.create_sheet(SHEET_META)
    wsM["A1"], wsM["B1"] = "Užsakovas", header.get("uzsakovas","")
    wsM["A2"], wsM["B2"] = "Vykdytojas", header.get("vykdytojas","")
    wsM["A3"], wsM["B3"] = "Sutarties nr.", header.get("sutartis","")
    wsM["A4"], wsM["B4"] = "Objektas / adresas", preset_address or header.get("adresas","")
    wsM["A5"], wsM["B5"] = "Skyrius", header.get("skyrius","")
    wsM["A6"], wsM["B6"] = "Akto data", header.get("akto_data","")
    wsM["A7"], wsM["B7"] = "Atlikimo metai", header.get("metai_atlik","2026")
    autosize(wsM)

    # Pavadinimas pasirinktai adreso reikšmei
    wb.defined_names.append(DefinedName(name="AdrSelected", attr_text=f"'{SHEET_META}'!$B$4"))

    # --- LISTOS (paieškos laukai + spill) ---
    wsL = wb.create_sheet(SHEET_LIST)
    wsL["A1"] = "Adresų paieška:"; wsL["B1"] = ""                 # įvedimo laukas
    wsL["A2"] = f"=UNIQUE({SHEET_DB}!A2:A100000)"                 # A2#
    wsL["C1"] = "Filtruoti adresai (DV)"
    wsL["C2"] = f_FILTER("A2#", f_SEARCH("B1", "A2#"), '"Nėra atitikmenų"')

    wsL["E1"] = "Paslaugų paieška:"; wsL["F1"] = ""               # įvedimo laukas
    wsL["E2"] = f_FILTER(f"{SHEET_DB}!B2:B100000",
                         f"{SHEET_DB}!A2:A100000=AdrSelected",
                         '"Nėra"')                                 # E2#
    wsL["F2"] = f_FILTER("E2#", f_SEARCH("F1", "E2#"), '"Nėra"')   # F2#
    autosize(wsL)

    wb.defined_names.append(DefinedName(name="AdresaiDV",   attr_text=f"'{SHEET_LIST}'!$C$2#"))
    wb.defined_names.append(DefinedName(name="PaslaugosDV", attr_text=f"'{SHEET_LIST}'!$F$2#"))

    # --- AKTAS ---
    ws = wb.create_sheet(SHEET_AKT)
    labels = ["Užsakovas", "Vykdytojas", "Sutarties nr.", "Objektas / adresas", "Skyrius", "Akto data"]
    for i, lab in enumerate(labels, start=1):
        ws[f"A{i}"] = f'="{lab}: " & {SHEET_META}!B{i}'
    ws["A7"] = f'="Atlikimo data: " & {SHEET_META}!B7 & " m. "'

    ws["A8"], ws["B8"], ws["C8"], ws["D8"], ws["E8"] = \
        "Eil. Nr.", "Paslaugos pavadinimas", "Kiekis", "Įkainis (be PVM)", "Suma (be PVM)"
    set_borders(ws, "A8:E8", thick=True)

    dv_service = DataValidation(type="list", formula1="=PaslaugosDV", allow_blank=True)
    ws.add_data_validation(dv_service)
    dv_nonneg = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
    ws.add_data_validation(dv_nonneg)

    for idx in range(max_lines):
        r = ROW_TABLE_START + idx
        ws.cell(r, 1).value = idx + 1     # Eil. Nr.
        dv_service.add(ws.cell(r, 2))     # Paslauga (dropdown)

        # C = numatytas kiekis pagal (adresas + paslauga)
        ws.cell(r, 3).value = (
            f'=IFERROR(INDEX(FILTER({SHEET_DB}!D2:D100000'
            f'{SEP}({SHEET_DB}!B2:B100000=B{r})*({SHEET_DB}!A2:A100000=AdrSelected))'
            f'{SEP}1)'
            f'{SEP}"" )'
        )
        ws.cell(r, 3).number_format = FMT_QTY
        dv_nonneg.add(ws.cell(r, 3))

        # D = įkainis pagal (adresas + paslauga)
        ws.cell(r, 4).value = (
            f'=IFERROR(INDEX(FILTER({SHEET_DB}!C2:C100000'
            f'{SEP}({SHEET_DB}!B2:B100000=B{r})*({SHEET_DB}!A2:A100000=AdrSelected))'
            f'{SEP}1)'
            f'{SEP}"" )'
        )
        ws.cell(r, 4).number_format = FMT_MONEY
        dv_nonneg.add(ws.cell(r, 4))

        # E = C * D
        ws.cell(r, 5).value = f"=C{r}*D{r}"
        ws.cell(r, 5).number_format = FMT_MONEY

    set_borders(ws, f"A{ROW_TABLE_START}:E{ROW_TABLE_START+max_lines-1}")

    # Vertikalios sumos (kol kas visam paruoštam diapazonui)
    ws["D12"] = "Suma (be PVM):"
    ws["E12"] = f"=SUM(E{ROW_TABLE_START}:E{ROW_TABLE_START+max_lines-1})"
    ws["D13"] = f"PVM {float(PVM_DEFAULT)}%:"
    ws["E13"] = f"=E12*{float(PVM_DEFAULT)/100}"
    ws["D14"] = "Suma su PVM:"
    ws["E14"] = "=E12+E13"
    for c in ("E12","E13","E14"):
        ws[c].number_format = FMT_MONEY
    set_borders(ws, "D12:E14", thick=True)

    # Adreso dropdown META!B4
    dv_addr = DataValidation(type="list", formula1="=AdresaiDV", allow_blank=False)
    wsM.add_data_validation(dv_addr); dv_addr.add(wsM["B4"])

    if hide_helper_sheets:
        wsM.sheet_state = "hidden"; wsL.sheet_state = "hidden"; wsK.sheet_state = "hidden"

    autosize(ws)
    return wb

# ======= VALYMAS (baituose) =======
def clean_before_sending_bytes(xlsx_bytes: bytes) -> bytes:
    """Grąžina 'švarią' XLSX versiją (ištrintos tuščios eilutės, perrašytos sumos)."""
    bio = io.BytesIO(xlsx_bytes)
    wb = load_workbook(bio)
    ws = wb[SHEET_AKT]

    start = ROW_TABLE_START
    theoretical_end = ROW_TABLE_START + MAX_LINES - 1
    end = min(ws.max_row, theoretical_end)

    last_used = start - 1
    for r in range(start, end + 1):
        if any(ws.cell(row=r, column=c).value not in (None, "") for c in (2, 3, 4, 5)):
            last_used = r

    # Ištrinti tuščias iš galo
    to_delete = end - last_used
    if to_delete > 0:
        ws.delete_rows(last_used + 1, to_delete)

    # Perrašyti sumas
    if last_used >= start:
        ws["E12"].value = f"=SUM(E{start}:E{last_used})"
    else:
        ws["E12"].value = "0"
    ws["E13"].value = f"=E12*{float(PVM_DEFAULT)/100}"
    ws["E14"].value = "=E12+E13"

    # Užtikrinam storas ribas sumų blokui
    side = Side(style="thick")
    for row in ws["D12:E14"]:
        for cell in row:
            cell.border = Border(left=side, right=side, top=side, bottom=side)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# ======= STREAMLIT UI =======
st.set_page_config(page_title="Aktų generatorius", layout="centered")
st.title("Aktų generatorius (Excel)")

with st.expander("Antraštės (A stulpelis: 'Etiketė: reikšmė')", expanded=True):
    c1, c2 = st.columns(2)
    uzsakovas   = c1.text_input("Užsakovas", "ANYKŠČIŲ RAJONO SAVIVALDYBĖS ADMINISTRACIJA")
    vykdytojas  = c2.text_input("Vykdytojas", "Corpus A, UAB")
    sutartis    = c1.text_input("Sutarties nr.", "6-793/CA-224154")
    skyrius     = c2.text_input("Skyrius", "ANA.P.A.J")
    akto_data   = c1.text_input("Akto data", "2026-01-04")
    metai_atlik = c2.text_input("Atlikimo metai", "2026")

with st.expander("KATALOGAS (CSV/XLSX)", expanded=True):
    up = st.file_uploader("Įkelk katalogą su stulpeliais: Adresas, Paslauga, Įkainis, Numatytas kiekis",
                          type=["csv","xlsx"])
    catalog_rows: List[Dict[str, str]] = []
    if up:
        try:
            if up.name.lower().endswith(".csv"):
                df = pd.read_csv(up)
            else:
                df = pd.read_excel(up, engine="openpyxl")
            cols = {c.lower(): c for c in df.columns}
            addr = cols.get("adresas") or "Adresas"
            serv = cols.get("paslauga") or "Paslauga"
            rate = cols.get("įkainis") or cols.get("ikainis") or "Įkainis"
            qty  = cols.get("numatytas kiekis") or "Numatytas kiekis"
            for _, row in df.iterrows():
                catalog_rows.append({
                    "address": str(row[addr]).strip(),
                    "service": str(row[serv]).strip(),
                    "rate":    str(row[rate]).strip(),
                    "qty":     str(row.get(qty, "0")).strip()
                })
            st.success(f"Įkelta {len(catalog_rows)} eil.")
        except Exception as e:
            st.error(f"Katalogo skaitymo klaida: {e}")
    else:
        st.info("Jei neįkelsi, bus naudotas pavyzdinis katalogas.")
        catalog_rows = [
            {"address":"J. Biliūno g. 19, Anykščiai","service":"Langų valymo paslauga valant iš abiejų pusių","rate":"2,00","qty":"50,38"},
            {"address":"J. Biliūno g. 19, Anykščiai","service":"Durų rankenų dezinfekavimas","rate":"0,50","qty":"30,00"},
            {"address":"Vilniaus g. 1, Anykščiai","service":"Grindų plovimas","rate":"1,20","qty":"100,00"},
            {"address":"Vilniaus g. 1, Anykščiai","service":"Sanitarinių mazgų valymas","rate":"3,50","qty":"10,00"},
        ]

# Adreso išankstinis parinkimas (pasirenkama)
unique_addresses = sorted({r["address"] for r in catalog_rows})
preset_on = st.checkbox("Iš anksto įrašyti adresą į META!B4 (pasirinksi čia)", value=False, help="Jei nepažymėsi, adresą pasirinksi Excel faile iš dropdown.")
preset_address = None
if preset_on:
    if unique_addresses:
        preset_address = st.selectbox("Adresas (META!B4)", unique_addresses, index=0)
    else:
        st.warning("Kataloge nerasta adresų — parinkti neišeina.")

# Generavimas
if st.button("Generuoti aktą"):
    header = {
        "uzsakovas": uzsakovas,
        "vykdytojas": vykdytojas,
        "sutartis": sutartis,
        "adresas": "",  # jei nepresetinsi, pasirinksi Excel'e
        "skyrius": skyrius,
        "akto_data": akto_data,
        "metai_atlik": metai_atlik,
    }
    wb = build_workbook(header, catalog_rows, max_lines=MAX_LINES, preset_address=preset_address, hide_helper_sheets=False)
    bio = io.BytesIO(); wb.save(bio); xlsx_bytes = bio.getvalue()

    st.success("Sugeneruota. META!B4 turi adresų dropdown; paslaugos filtruosis pagal adresą.")
    st.download_button("Atsisiųsti aktą (XLSX)", data=xlsx_bytes,
                       file_name="aktas_dropdown_paieška.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    cleaned_bytes = clean_before_sending_bytes(xlsx_bytes)
    st.download_button("Atsisiųsti švarų aktą (be tuščių eilučių)",
                       data=cleaned_bytes,
                       file_name="aktas_dropdown_paieška_clean.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("A7: „Atlikimo data: 2026 m. “ – mėnesį prirašyk ranka. Eilutėse C ir D pildosi automatiškai pagal pasirinktą adresą ir paslaugą; gali perrašyti.")

