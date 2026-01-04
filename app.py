
# -*- coding: utf-8 -*-
"""
Aktų generatorius – vienas lapas „AKTAS“, be papildomų sheet'ų.
Struktūra:
- Viršuje tik data (be „Akto data:“), po jos tuščia eilutė.
- Užsakovas -> Vykdytojas -> (pavadinimas „Atliktų paslaugų aktas“, bold + didesnis, centruotas) -> tuščia eilutė ->
  Sutarties nr. -> tuščia eilutė -> „Atliktų paslaugų data: 2026 m.“ -> Objekto adresas + Skyrius -> lentelė.
- UI: filtras pagal UŽSAKOVĄ (ir, jei yra, VADYBININKĄ), tada adresas ir paslaugų pasirinkimas.
- „Group by“ (pagal paslaugos pavadinimą ir įkainį) – sujungiamos eilutės, sumuojamas kiekis.
- Apačioje: PARAŠŲ BLOKAS dviem stulpeliais (Užsakovas / Vykdytojas), be datų.
"""

import io
import sys
import unicodedata
from typing import List, Dict, Optional
from decimal import Decimal, ROUND_HALF_UP

import streamlit as st
import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, NamedStyle

# ===== Konfigūros =====
PVM_DEFAULT = Decimal("21.00")
FMT_MONEY = "#,##0.00"
FMT_QTY = "#,##0.00"
TITLE_FONT_SIZE = 16  # pavadinimo dydis („Atliktų paslaugų aktas“)

# ===== Pagalbinės =====
def strip_accents(s: str) -> str:
    if s is None:
        return ""
    return "".join(ch for ch in unicodedata.normalize("NFD", str(s)) if unicodedata.category(ch) != "Mn")

def norm(s: str) -> str:
    return strip_accents(s).lower().strip()

def detect_delimiter(sample: str) -> Optional[str]:
    c_semi = sample.count(";")
    c_comma = sample.count(",")
    return ";" if c_semi > c_comma else None

def dec2(v) -> float:
    return float(Decimal(str(v).replace(",", ".")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

def create_named_styles(wb: Workbook) -> None:
    if "Money" not in wb.named_styles:
        stl = NamedStyle(name="Money")
        stl.number_format = FMT_MONEY
        stl.alignment = Alignment(horizontal="right")
        wb.add_named_style(stl)
    if "Qty" not in wb.named_styles:
        stl = NamedStyle(name="Qty")
        stl.number_format = FMT_QTY
        stl.alignment = Alignment(horizontal="right")
        wb.add_named_style(stl)

def set_borders(ws, rng: str, thick: bool = False) -> None:
    side = Side(style="thick" if thick else "thin")
    for row in ws[rng]:
        for c in row:
            c.border = Border(left=side, right=side, top=side, bottom=side)

def autosize(ws) -> None:
    for col in ws.columns:
        try:
            letter = col[0].column_letter
        except Exception:
            continue
        max_len = 0
        for cell in col:
            val = str(cell.value) if cell.value is not None else ""
            max_len = max(max_len, len(val))
        ws.column_dimensions[letter].width = min(max_len + 2, 60)

# ===== Failo skaitymas + stulpelių atpažinimas =====
def read_catalog(uploaded) -> pd.DataFrame:
    """Skaito CSV/XLSX; aptinka skyriklį; normalizuoja antraštes (viduje)."""
    if uploaded.name.lower().endswith(".csv"):
        head = uploaded.read(4096).decode("utf-8", errors="ignore")
        uploaded.seek(0)
        delim = detect_delimiter(head)
        df = pd.read_csv(uploaded, sep=delim) if delim else pd.read_csv(uploaded)
    else:
        df = pd.read_excel(uploaded, engine="openpyxl")
    df.columns = [norm(c) for c in df.columns]
    return df

def map_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    """
    Privaloma: address, service, rate, qty
    Pasirenkama: skyrius, uzsakovas, vykdytojas, sutartis, vadybininkas
    """
    col_map: Dict[str, Optional[str]] = {}
    for c in df.columns:
        cn = norm(c)
        if "adres" in cn: col_map["address"] = c
        elif "paslaug" in cn: col_map["service"] = c
        elif "ikain" in cn or "įkain" in cn: col_map["rate"] = c
        elif any(k in cn for k in ["kiek", "plot", "m2", "m3", "m³", "vnt", "val", "apimt", "sanaud"]): col_map["qty"] = c
        elif "skyri" in cn: col_map["skyrius"] = c
        elif "uzsakov" in cn or "užsakov" in cn: col_map["uzsakovas"] = c
        elif "vykdytoj" in cn: col_map["vykdytojas"] = c
        elif "sutart" in cn: col_map["sutartis"] = c
        elif "vadybin" in cn: col_map["vadybininkas"] = c

    for req in ("address", "service", "rate", "qty"):
        if req not in col_map:
            raise ValueError(f"Nerastas stulpelis '{req}' (trūksta vieno iš: adresas/paslauga/įkainis/kiekis).")

    for opt in ("skyrius", "uzsakovas", "vykdytojas", "sutartis", "vadybininkas"):
        col_map.setdefault(opt, None)
    return col_map

# ===== AKTO kūrimas – 1 lapas =====
def build_workbook_act_single_sheet(
    df: pd.DataFrame,
    col_map: Dict[str, Optional[str]],
    selected_address: str,
    selected_services: List[str],
    akto_data: str,
    group_same: bool,
) -> Workbook:
    """
    Generuoja tik vieną lapą „AKTAS“ su:
    - Viršuje tik data (be „Akto data:“) + tuščia eilutė
    - Užsakovas, Vykdytojas, po to pavadinimas „Atliktų paslaugų aktas“ (bold, didesnis, centruotas), tarpas,
      Sutarties nr., tarpas, „Atliktų paslaugų data: 2026 m.“, Objekto adresas + Skyrius
    - Lentele, sumos, PVM, suma su PVM
    - PARAŠŲ BLOKAS apačioje dviem stulpeliais (kairė – Užsakovas, dešinė – Vykdytojas), be datų
    """
    # Filtruojam pagal adresą
    dfa = df[df[col_map["address"]].astype(str).str.strip() == str(selected_address).strip()].copy()
    if dfa.empty:
        raise ValueError("Pagal pasirinktą adresą įrašų nerasta.")

    # Metaduomenys paimami iš pirmo atitikties įrašo
    uzsakovas = str(dfa.iloc[0][col_map["uzsakovas"]]).strip() if col_map["uzsakovas"] else ""
    vykdytojas = str(dfa.iloc[0][col_map["vykdytojas"]]).strip() if col_map["vykdytojas"] else ""
    sutartis  = str(dfa.iloc[0][col_map["sutartis"]]).strip()  if col_map["sutartis"]  else ""
    skyrius   = str(dfa.iloc[0][col_map["skyrius"]]).strip()   if col_map["skyrius"]   else ""

    # Jei vartotojas nepasirinko konkrečių paslaugų — imame visas to adreso
    if not selected_services:
        selected_services = sorted(set(dfa[col_map["service"]].astype(str).str.strip()))

    # Atrenkame tik pasirinktas paslaugas
    dfa = dfa[dfa[col_map["service"]].astype(str).str.strip().isin([s.strip() for s in selected_services])].copy()
    if dfa.empty:
        raise ValueError("Pasirinktų paslaugų sąrašas tuščias šiam adresui.")

    # Paruošiam workbook
    wb = Workbook()
    create_named_styles(wb)
    ws = wb.active
    ws.title = "AKTAS"

    # ===== Antraštės pagal tvarką =====
    ws["A1"] = akto_data       # viršuje tik data
    ws["A2"] = ""              # tuščia eilutė

    ws["A3"] = f"Užsakovas: {uzsakovas}"
    ws["A4"] = f"Vykdytojas: {vykdytojas}"

    # Pavadinimas po 'Vykdytojas'
    ws.merge_cells("A5:E5")
    ws["A5"] = "Atliktų paslaugų aktas"
    ws["A5"].font = Font(bold=True, size=TITLE_FONT_SIZE)
    ws["A5"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 24

    ws["A6"] = ""  # tarpas po pavadinimo

    ws["A7"] = f"Sutarties nr.: {sutartis}"
    ws["A8"] = ""  # tarpas po sutarties

    ws["A9"]  = "Atliktų paslaugų data: 2026 m."  # fiksuota
    addr_line = f"Objekto adresas: {selected_address}"
    if skyrius:
        addr_line += f", Skyrius: {skyrius}"
    ws["A10"] = addr_line

    ws["A11"] = ""  # tarpas prieš lentelę

    # ===== Paruošiame duomenis lentelei (group by pagal poreikį) =====
    # Normalizuojam laukus
    dfa = dfa.assign(
        __service=dfa[col_map["service"]].astype(str).str.strip(),
        __qty=dfa[col_map["qty"]],
        __rate=dfa[col_map["rate"]],
    )
    # konvertuojam skaitines
    dfa["__qty"]  = dfa["__qty"].apply(dec2)
    dfa["__rate"] = dfa["__rate"].apply(dec2)

    if group_same:
        # Sujungiame vienodas paslaugas su tuo pačiu įkainiu – sumuojam kiekius
        dfg = (
            dfa.groupby(["__service", "__rate"], as_index=False)["__qty"]
               .sum()
               .sort_values(["__service", "__rate"], kind="stable")
        )
        rows_iter = [(r["__service"], r["__qty"], r["__rate"]) for _, r in dfg.iterrows()]
    else:
        # Negrupuojam – paliekam eiles tokia tvarka, kaip faile (stabiliai)
        rows_iter = [(r["__service"], r["__qty"], r["__rate"]) for _, r in dfa.iterrows()]

    # ===== Lentele su paslaugomis =====
    ws["A12"], ws["B12"], ws["C12"], ws["D12"], ws["E12"] = \
        "Eil. Nr.", "Paslaugos pavadinimas", "Kiekis", "Įkainis (be PVM)", "Suma (be PVM)"
    set_borders(ws, "A12:E12", thick=True)
    ws["A12"].font = ws["B12"].font = ws["C12"].font = ws["D12"].font = ws["E12"].font = Font(bold=True)

    r = 13  # duomenų eilučių pradžia
    for i, (service, qty, rate) in enumerate(rows_iter, start=1):
        ws[f"A{r}"] = i
        ws[f"B{r}"] = service
        ws[f"C{r}"] = qty;  ws[f"C{r}"].number_format  = FMT_QTY
        ws[f"D{r}"] = rate; ws[f"D{r}"].number_format = FMT_MONEY
        ws[f"E{r}"] = f"=C{r}*D{r}"; ws[f"E{r}"].number_format = FMT_MONEY
        r += 1

    # ===== Tarpinė suma + PVM + galutinė =====
    ws[f"D{r}"] = "Suma (be PVM):"
    first_data_row = 13
    last_data_row  = r - 1
    if last_data_row >= first_data_row:
        ws[f"E{r}"] = f"=SUM(E{first_data_row}:E{last_data_row})"
    else:
        ws[f"E{r}"] = 0
    ws[f"E{r}"].number_format = FMT_MONEY
    set_borders(ws, f"D{r}:E{r}", thick=True)
    r += 1

    ws[f"D{r}"] = f"PVM {float(PVM_DEFAULT)}%:"
    ws[f"E{r}"] = f"=E{r-1}*{float(PVM_DEFAULT)/100}"; ws[f"E{r}"].number_format = FMT_MONEY
    r += 1

    ws[f"D{r}"] = "Suma su PVM:"
    ws[f"E{r}"] = f"=E{r-2}+E{r-1}"; ws[f"E{r}"].number_format = FMT_MONEY
    set_borders(ws, f"D{r-2}:E{r}", thick=True)
    r += 2  # šiek tiek oro prieš parašų bloką

    # ===== PARAŠŲ BLOKAS APAČIOJE – KAIRĖ/DEŠINĖ (be datų) =====
    # Kairė: Užsakovas (linija per B:D)
    ws[f"A{r}"] = "Užsakovas:"
    ws[f"A{r}"].font = Font(bold=True)
    for col in ("B", "C", "D"):
        ws[f"{col}{r}"].border = Border(bottom=Side(style="thin"))
    ws.row_dimensions[r].height = 22
    ws[f"A{r+1}"] = "Vardas, pavardė / Pareigos"
    ws.row_dimensions[r+1].height = 18

    # Dešinė: Vykdytojas (linija per F:H)
    ws[f"F{r}"] = "Vykdytojas:"
    ws[f"F{r}"].font = Font(bold=True)
    for col in ("G", "H", "I"):
        ws[f"{col}{r}"].border = Border(bottom=Side(style="thin"))
    ws.row_dimensions[r].height = max(ws.row_dimensions[r].height, 22)
    ws[f"F{r+1}"] = "Vardas, pavardė / Pareigos"
    ws.row_dimensions[r+1].height = max(ws.row_dimensions[r+1].height, 18)

    autosize(ws)
    return wb

# ===== STREAMLIT UI =====
st.set_page_config(page_title="AKTAS (vienas lapas)", layout="centered")
st.title("Aktų generatorius (vienas lapas: AKTAS)")
st.caption(f"Python: {sys.version}")

# Katalogo įkėlimas
with st.expander("KATALOGAS (CSV/XLSX)", expanded=True):
    up = st.file_uploader(
        "Įkelk katalogą (privaloma: adresas/paslauga/įkainis/kiekis; pasirenkama: skyrius/užsakovas/vykdytojas/sutartis/vadybininkas)",
        type=["csv", "xlsx"]
    )
    if not up:
        st.info("Įkelk failą, tuomet atsiras pasirinkimai.")
        st.stop()

# Skaitymas + map
try:
    df_raw = read_catalog(up)
    col_map = map_columns(df_raw)
except Exception as e:
    st.error(f"Katalogo klaida: {e}")
    st.stop()

# Darbinė kopija
df = df_raw.copy()

# ===== UŽSAKOVO FILTRAS =====
if col_map.get("uzsakovas"):
    with st.expander("Filtras pagal užsakovą (pasirenkama)", expanded=True):
        uzs_list = sorted(set(str(x).strip() for x in df[col_map["uzsakovas"]].dropna() if str(x).strip()))
        sel_uzs = st.multiselect("Užsakovas(-ai)", uzs_list)
        if sel_uzs:
            df = df[df[col_map["uzsakovas"]].astype(str).str.strip().isin(sel_uzs)]

# (neprivaloma) VADYBININKO FILTRAS
if col_map.get("vadybininkas"):
    with st.expander("Filtras pagal vadybininką (pasirenkama)", expanded=False):
        mgrs = sorted(set(str(x).strip() for x in df[col_map["vadybininkas"]].dropna()))
        sel_mgrs = st.multiselect("Vadybininkas(-ai)", mgrs)
        if sel_mgrs:
            df = df[df[col_map["vadybininkas"]].astype(str).str.strip().isin(sel_mgrs)]

# Adreso pasirinkimas (po filtrų)
addresses = sorted(set(str(x).strip() for x in df[col_map["address"]].dropna() if str(x).strip()))
if not addresses:
    st.error("Po filtrų nerasta adresų.")
    st.stop()

selected_address = st.selectbox("Objekto adresas", addresses, index=0)

# Paslaugų pasirinkimas (tik to adreso ir po filtrų)
services_for_addr = sorted(set(
    df[df[col_map["address"]].astype(str).str.strip() == selected_address][col_map["service"]].astype(str).str.strip()
))
selected_services = st.multiselect(
    "Paslaugos (šiam adresui — pažymėk, kurias įtraukti į aktą)",
    services_for_addr,
    default=services_for_addr  # pagal nutylėjimą – visos
)

# Group by pasirinkimas
group_same = st.checkbox("Sujungti vienodas paslaugas (grupuoti pagal pavadinimą ir įkainį)", value=True)

# Data viršuje (be teksto)
akto_data = st.text_input("Data (rodoma viršuje)", "2026-01-04")

# Generavimo mygtukas
btn = st.button("Generuoti AKTĄ (XLSX)", use_container_width=True, disabled=not bool(selected_address))
if btn:
    try:
        wb = build_workbook_act_single_sheet(df, col_map, selected_address, selected_services, akto_data, group_same)
        bio = io.BytesIO(); wb.save(bio); xlsx = bio.getvalue()
        st.download_button("Atsisiųsti AKTĄ", xlsx, "aktas_vienas_lapas.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.exception(e)
