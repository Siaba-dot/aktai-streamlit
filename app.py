
# -*- coding: utf-8 -*-
import io
import unicodedata
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

SHEET_DB   = "KATALOGAS"
SHEET_META = "META"
SHEET_LIST = "LISTOS"
SHEET_MAP  = "MAP"
SHEET_AKT  = "AKTAS"

# ======= PAGALBA: diakritikų nuėmimas ir stulpelių atpažinimas =======
def strip_accents(s: str) -> str:
    return ''.join(ch for ch in unicodedata.normalize('NFD', s) if unicodedata.category(ch) != 'Mn')

def norm(s: str) -> str:
    return strip_accents(str(s)).lower().strip()

def detect_delimiter(sample: str) -> str:
    c_semi = sample.count(";"); c_comma = sample.count(",")
    # jei daugiausia ';' - laikom ';', kitaip paliekam pandas numatytą
    return ";" if c_semi > c_comma else None

def read_catalog(uploaded) -> pd.DataFrame:
    """Skaito CSV/XLSX, aptinka skyriklį, sutvarko stulpelių pavadinimus."""
    if uploaded.name.lower().endswith(".csv"):
        head = uploaded.read(4096).decode("utf-8", errors="ignore")
        uploaded.seek(0)
        delim = detect_delimiter(head)
        df = pd.read_csv(uploaded, sep=delim) if delim else pd.read_csv(uploaded)
    else:
        df = pd.read_excel(uploaded, engine="openpyxl")

    # Normalizuoti stulpelių pavadinimus
    df_cols_norm = {c: norm(c) for c in df.columns}
    df = df.rename(columns=df_cols_norm)
    return df

def map_columns(df: pd.DataFrame) -> Dict[str, str]:
    """Grąžina žemėlapį į standartinius laukus (address, service, rate, qty, skyrius, uzsakovas, vykdytojas, sutartis)."""
    col_map = {}
    for c in df.columns:
        cn = norm(c)
        if "adres" in cn:         col_map["address"]   = c
        elif "paslaug" in cn:     col_map["service"]   = c
        elif "ikain" in cn or "įkain" in cn: col_map["rate"] = c
        elif "kiek" in cn:        col_map["qty"]       = c
        elif "skyri" in cn:       col_map["skyrius"]   = c
        elif "uzsakov" in cn or "užsakov" in cn: col_map["uzsakovas"] = c
        elif "vykdytoj" in cn:    col_map["vykdytojas"] = c
        elif "sutart" in cn:      col_map["sutartis"]  = c
    # minimalus rinkinys privalomas
    for req in ("address","service","rate","qty"):
        if req not in col_map:
            raise ValueError(f"Nerastas stulpelis '{req}'. Patikrink CSV/XLSX pavadinimus.")
    # pasirenkami (gali nebūti)
    for opt in ("skyrius","uzsakovas","vykdytojas","sutartis"):
        if opt not in col_map:
            col_map[opt] = None
    return col_map

# ======= STILIUS =======
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
    df: pd.DataFrame,
    col_map: Dict[str, str],
    selection_mode: str,            # "adresas" arba "skyrius"
    selected_key: Optional[str],    # pasirinktas adresas arba skyrius (jei preset)
    selected_services: Optional[List[str]],  # jei pasirinktos konkrečios paslaugos
    header_date: str,               # akto data (tekstinė)
    header_year_exec: str           # atlikimo metai (pvz., "2026")
) -> Workbook:

    wb = Workbook()
    create_named_styles(wb)

    # --- KATALOGAS ---
    wsK = wb.active; wsK.title = SHEET_DB
    wsK.append(["Adresas","Paslauga","Įkainis (be PVM)","Numatytas kiekis","Skyrius","Užsakovas","Vykdytojas","Sutarties nr."])
    for _, row in df.iterrows():
        wsK.append([
            str(row[col_map["address"]]).strip(),
            str(row[col_map["service"]]).strip(),
            float(Decimal(str(row[col_map["rate"]]).replace(",", ".") if col_map["rate"] else "0").quantize(Decimal("0.01"))),
            float(Decimal(str(row[col_map["qty"]]).replace(",", ".") if col_map["qty"] else "0").quantize(Decimal("0.01"))),
            str(row[col_map["skyrius"]]).strip() if col_map["skyrius"] else "",
            str(row[col_map["uzsakovas"]]).strip() if col_map["uzsakovas"] else "",
            str(row[col_map["vykdytojas"]]).strip() if col_map["vykdytojas"] else "",
            str(row[col_map["sutartis"]]).strip() if col_map["sutartis"] else ""
        ])
    autosize(wsK)

    # --- MAP (unikalus adresas/skyrius -> meta laukai) ---
    wsMap = wb.create_sheet(SHEET_MAP)
    wsMap.append(["Adresas","Skyrius","Užsakovas","Vykdytojas","Sutarties nr."])
    # Konsoliduojame pagal unikalų adresą
    addr_grp = df.groupby(col_map["address"]).first()
    for addr, r in addr_grp.iterrows():
        wsMap.append([
            str(addr).strip(),
            str(r[col_map["skyrius"]]).strip() if col_map["skyrius"] else "",
            str(r[col_map["uzsakovas"]]).strip() if col_map["uzsakovas"] else "",
            str(r[col_map["vykdytojas"]]).strip() if col_map["vykdytojas"] else "",
            str(r[col_map["sutartis"]]).strip() if col_map["sutartis"] else ""
        ])
    autosize(wsMap)

    # --- META ---
    wsM = wb.create_sheet(SHEET_META)
    wsM["A1"], wsM["B1"] = "Užsakovas", ""        # TUŠČI
    wsM["A2"], wsM["B2"] = "Vykdytojas", ""       # TUŠČI
    wsM["A3"], wsM["B3"] = "Sutarties nr.", ""    # TUŠČI
    wsM["A4"], wsM["B4"] = "Objektas / adresas", selected_key if (selection_mode=="adresas" and selected_key) else ""
    wsM["A5"], wsM["B5"] = "Skyrius", selected_key if (selection_mode=="skyrius" and selected_key) else ""
    wsM["A6"], wsM["B6"] = "Akto data", header_date
    wsM["A7"], wsM["B7"] = "Atlikimo metai", header_year_exec

    # AdrSelected / SkySelected pavadinimai
    wb.defined_names.append(DefinedName(name="AdrSelected",  attr_text=f"'{SHEET_META}'!$B$4"))
    wb.defined_names.append(DefinedName(name="SkySelected",  attr_text=f"'{SHEET_META}'!$B$5"))

    # META auto‑pildymas iš MAP pagal pasirinktą ADRESĄ arba SKYRIŲ
    # MAP stulpiai: A=Adresas, B=Skyrius, C=Užsakovas, D=Vykdytojas, E=Sutarties nr.
    wsM["B1"] = (
        f'=IFERROR(IF(LEN(B4)>0, INDEX(FILTER({SHEET_MAP}!C2:C100000{SEP}{SHEET_MAP}!A2:A100000=B4),1)'
        f'{SEP}IF(LEN(B5)>0, INDEX(FILTER({SHEET_MAP}!C2:C100000{SEP}{SHEET_MAP}!B2:B100000=B5),1){SEP}"" )){SEP}"" )'
    )
    wsM["B2"] = (
        f'=IFERROR(IF(LEN(B4)>0, INDEX(FILTER({SHEET_MAP}!D2:D100000{SEP}{SHEET_MAP}!A2:A100000=B4),1)'
        f'{SEP}IF(LEN(B5)>0, INDEX(FILTER({SHEET_MAP}!D2:D100000{SEP}{SHEET_MAP}!B2:B100000=B5),1){SEP}"" )){SEP}"" )'
    )
    wsM["B3"] = (
        f'=IFERROR(IF(LEN(B4)>0, INDEX(FILTER({SHEET_MAP}!E2:E100000{SEP}{SHEET_MAP}!A2:A100000=B4),1)'
        f'{SEP}IF(LEN(B5)>0, INDEX(FILTER({SHEET_MAP}!E2:E100000{SEP}{SHEET_MAP}!B2:B100000=B5),1){SEP}"" )){SEP}"" )'
    )

    # --- LISTOS: paieška + dropdown šaltiniai ---
    wsL = wb.create_sheet(SHEET_LIST)
    # Unikalūs adresai ir skyriai
    wsL["A1"] = "Adresų paieška:"; wsL["B1"] = ""
    wsL["A2"] = f"=UNIQUE({SHEET_DB}!A2:A100000)"
    wsL["C1"] = "Filtruoti adresai (DV)"
    wsL["C2"] = f_FILTER("A2#", f_SEARCH("B1", "A2#"), '"Nėra atitikmenų"')

    wsL["E1"] = "Skyrių paieška:"; wsL["F1"] = ""
    wsL["E2"] = f"=UNIQUE({SHEET_DB}!E2:E100000)"
    wsL["G1"] = "Filtruoti skyriai (DV)"
    wsL["G2"] = f_FILTER("E2#", f_SEARCH("F1", "E2#"), '"Nėra"')

    # Paslaugos pagal pasirinktą ADRESĄ ar SKYRIŲ
    wsL["I1"] = "Paslaugų paieška:"; wsL["J1"] = ""
    wsL["I2"] = (
        f'=IF(LEN({SHEET_META}!B4)>0,'
        f'FILTER({SHEET_DB}!B2:B100000{SEP}{SHEET_DB}!A2:A100000={SHEET_META}!B4)'
        f'{SEP}IF(LEN({SHEET_META}!B5)>0,'
        f'FILTER({SHEET_DB}!B2:B100000{SEP}{SHEET_DB}!E2:E100000={SHEET_META}!B5)'
        f'{SEP}"Nėra"))'
    )
    wsL["J2"] = f_FILTER("I2#", f_SEARCH("J1", "I2#"), '"Nėra"')

    # DV pavadinimai
    wb.defined_names.append(DefinedName(name="AdresaiDV",   attr_text=f"'{SHEET_LIST}'!$C$2#"))
    wb.defined_names.append(DefinedName(name="SkyriaiDV",   attr_text=f"'{SHEET_LIST}'!$G$2#"))
    wb.defined_names.append(DefinedName(name="PaslaugosDV", attr_text=f"'{SHEET_LIST}'!$J$2#"))

    # --- AKTAS: antraštė + lentelė ---
    ws = wb.create_sheet(SHEET_AKT)
    # A1–A6: „Etiketė: Reikšmė“ (Užsakovas/Vykdytojas/Sutarties nr./Skyrius pradžioje tušti, pildysis iš MAP)
    labels = ["Užsakovas", "Vykdytojas", "Sutarties nr.", "Objektas / adresas", "Skyrius", "Akto data"]
    for i, lab in enumerate(labels, start=1):
        ws[f"A{i}"] = f'="{lab}: " & {SHEET_META}!B{i}'
    ws["A7"] = f'="Atlikimo data: " & {SHEET_META}!B7 & " m. "'

    # Antraštės
    ws["A8"], ws["B8"], ws["C8"], ws["D8"], ws["E8"] = \
        "Eil. Nr.", "Paslaugos pavadinimas", "Kiekis", "Įkainis (be PVM)", "Suma (be PVM)"
    set_borders(ws, "A8:E8", thick=True)

    # DV: adresai/ skyriai META lape
    dv_addr = DataValidation(type="list", formula1="=AdresaiDV", allow_blank=True)
    dv_sky  = DataValidation(type="list", formula1="=SkyriaiDV", allow_blank=True)
    wsM.add_data_validation(dv_addr); dv_addr.add(wsM["B4"])
    wsM.add_data_validation(dv_sky);  dv_sky.add(wsM["B5"])

    # DV: paslaugos lentelėje
    dv_service = DataValidation(type="list", formula1="=PaslaugosDV", allow_blank=True)
    ws.add_data_validation(dv_service)
    dv_nonneg = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
    ws.add_data_validation(dv_nonneg)

    # Paruoštos eilutės
    for idx in range(MAX_LINES):
        r = ROW_TABLE_START + idx
        ws.cell(r, 1).value = idx + 1
        dv_service.add(ws.cell(r, 2))           # B = paslaugos dropdown (filtruojama pagal adresą/skyrių)
        # C = numatytas kiekis
        ws.cell(r, 3).value = (
            f'=IFERROR(INDEX(FILTER({SHEET_DB}!D2:D100000'
            f'{SEP}({SHEET_DB}!B2:B100000=B{r})*('
            f'IF(LEN({SHEET_META}!B4)>0,{SHEET_DB}!A2:A100000={SHEET_META}!B4,{SHEET_DB}!E2:E100000={SHEET_META}!B5)'
            f')){SEP}1){SEP}"" )'
        )
        ws.cell(r, 3).number_format = FMT_QTY
        dv_nonneg.add(ws.cell(r, 3))

        # D = įkainis
        ws.cell(r, 4).value = (
            f'=IFERROR(INDEX(FILTER({SHEET_DB}!C2:C100000'
            f'{SEP}({SHEET_DB}!B2:B100000=B{r})*('
            f'IF(LEN({SHEET_META}!B4)>0,{SHEET_DB}!A2:A100000={SHEET_META}!B4,{SHEET_DB}!E2:E100000={SHEET_META}!B5)'
            f')){SEP}1){SEP}"" )'
        )
        ws.cell(r, 4).number_format = FMT_MONEY
        dv_nonneg.add(ws.cell(r, 4))

        # E = C * D
        ws.cell(r, 5).value = f"=C{r}*D{r}"
        ws.cell(r, 5).number_format = FMT_MONEY

    set_borders(ws, f"A{ROW_TABLE_START}:E{ROW_TABLE_START+MAX_LINES-1}")

    # Sumos/PVM
    ws["D12"] = "Suma (be PVM):"; ws["E12"] = f"=SUM(E{ROW_TABLE_START}:E{ROW_TABLE_START+MAX_LINES-1})"
    ws["D13"] = f"PVM {float(PVM_DEFAULT)}%:"; ws["E13"] = f"=E12*{float(PVM_DEFAULT)/100}"
    ws["D14"] = "Suma su PVM:"; ws["E14"] = "=E12+E13"
    set_borders(ws, "D12:E14", thick=True)

    autosize(ws); autosize(wsM); autosize(wsL); autosize(wsMap)
    return wb

# ======= VALYMAS (baituose) =======
def clean_before_sending_bytes(xlsx_bytes: bytes) -> bytes:
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

    to_delete = end - last_used
    if to_delete > 0:
        ws.delete_rows(last_used + 1, to_delete)

    if last_used >= start:
        ws["E12"].value = f"=SUM(E{start}:E{last_used})"
    else:
        ws["E12"].value = "0"
    ws["E13"].value = f"=E12*{float(PVM_DEFAULT)/100}"
    ws["E14"].value = "=E12+E13"

    out = io.BytesIO(); wb.save(out)
    return out.getvalue()

# ======= STREAMLIT UI =======
st.set_page_config(page_title="Aktų generatorius (Excel)", layout="centered")
st.title("Aktų generatorius (Excel)")

with st.expander("Antraštės (A stulpelyje: „Etiketė: reikšmė“)", expanded=True):
    c1, c2 = st.columns(2)
    akto_data   = c1.text_input("Akto data", "2026-01-04")
    metai_atlik = c2.text_input("Atlikimo metai", "2026")
    st.caption("Pastaba: Užsakovas / Vykdytojas / Sutarties nr. / Skyrius paliekami tušti ir užsipildys pagal pasirinktą SKYRIŲ arba ADRESĄ.")

with st.expander("KATALOGAS (CSV/XLSX)", expanded=True):
    up = st.file_uploader("Įkelk katalogą su stulpeliais: Adresas, Paslauga, Įkainis, Numatytas kiekis (+ pasirenkamai: Skyrius, Užsakovas, Vykdytojas, Sutarties nr.)",
                          type=["csv","xlsx"])
    if not up:
        st.stop()

try:
    df = read_catalog(up)
    col_map = map_columns(df)
except Exception as e:
    st.error(f"Katalogo skaitymo/atpažinimo klaida: {e}")
    st.stop()

# Parinkimas: pagal Adresą arba pagal Skyrių
mode = st.radio("Pasirinkimas", ["Pagal objekto adresą", "Pagal skyrių"], horizontal=True)
selection_mode = "adresas" if mode.startswith("Pagal objekto adresą") else "skyrius"

# Galimų adresų/skyrių sąrašai
unique_addresses = sorted(set(str(x).strip() for x in df[col_map["address"]].dropna() if str(x).strip()))
unique_depart    = sorted(set(str(x).strip() for x in (df[col_map["skyrius"]] if col_map["skyrius"] else pd.Series(dtype=str)).dropna() if str(x).strip()))

# Parinktas raktas (preset)
selected_key = None
if selection_mode == "adresas":
    if not unique_addresses:
        st.warning("Kataloge nerasta adresų (tušti ar neaptikti stulpeliai).")
    else:
        selected_key = st.selectbox("Objekto adresas", unique_addresses, index=0)
else:
    if not unique_depart:
        st.warning("Kataloge nerasta skyrių (tušti ar neaptikti stulpeliai).")
    else:
        selected_key = st.selectbox("Skyrius", unique_depart, index=0)

# Paslaugų pasirinkimas (multiselect) pagal raktą
if selection_mode == "adresas":
    filter_mask = df[col_map["address"]].astype(str).str.strip() == selected_key
else:
    if col_map["skyrius"]:
        filter_mask = df[col_map["skyrius"]].astype(str).str.strip() == selected_key
    else:
        filter_mask = pd.Series([False]*len(df))

services_pool = sorted(set(df.loc[filter_mask, col_map["service"]].astype(str).str.strip()))
selected_services = st.multiselect("Paslaugos (jei nieko nepasirinksi, bus imtos visos pagal raktą)", services_pool)

if st.button("Generuoti aktą"):
    wb = build_workbook(
        df=df,
        col_map=col_map,
        selection_mode=selection_mode,
        selected_key=selected_key,
        selected_services=selected_services if selected_services else None,
        header_date=akto_data,
        header_year_exec=metai_atlik
    )
    bio = io.BytesIO(); wb.save(bio); xlsx_bytes = bio.getvalue()

    st.success("Sugeneruota. META!B4 (Adresas) ir META!B5 (Skyrius) turi dropdown su paieška. Užsakovas/Vykdytojas/Sutarties nr. užsipildys pagal pasirinktą raktą.")
    st.download_button("Atsisiųsti aktą (XLSX)", data=xlsx_bytes,
                       file_name="aktas_dropdown_paieška.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    cleaned = clean_before_sending_bytes(xlsx_bytes)
    st.download_button("Atsisiųsti švarų aktą (be tuščių eilučių)",
                       data=cleaned,
                       file_name="aktas_dropdown_paieška_clean.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
