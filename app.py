
# -*- coding: utf-8 -*-
"""
Aktų generatorius – vienas lapas „AKTAS“, be papildomų sheet'ų.

PAGRINDINĖ LOGIKA:
- Pirma: UŽSAKOVO multiselect (pagal jį formuojami kontraktai)
- Antra: SUTARTIES numerio pasirinkimas (iš pasirinktų užsakovų įrašų)
- Toliau: (jei yra) VADYBININKO filtras -> keli ADRESAI -> GLOBALUS paslaugų filtras ->
         PASLAUGŲ parinkimas kiekvienam adresui atskirai -> Group by -> Data viršuje -> Generavimas

IŠDĖSTYMAS LAPE:
- Viršuje: tik DATA (be „Akto data:“), po jos tuščia eilutė
- „Užsakovas: …“, „Vykdytojas: …“
- „Atliktų paslaugų aktas“ – paryškintas, CENTRUOTAS per A–E
- Tuščia eilutė, „Sutarties nr.: …“, tuščia eilutė
- Kiekvienam adresui:
    „Atliktų paslaugų data: 2026 m.“ ->
    „Objekto adresas …, Skyrius …“ ->
    lentelė: A=Eil. Nr. (siauras), B=Paslaugos (wrap), C=Kiekis, D=Įkainis, E=Suma
- Po visų blokų: Bendra suma (be PVM, PVM, su PVM)
- PARAŠŲ blokas A–E ribose (kairė: Užsakovas, dešinė: Vykdytojas), be datų.

Pastaba: jokių defined_names, jokių kitų lapų – tik XLSX aktas.
"""

import io
import sys
import unicodedata
from typing import List, Dict, Optional, Tuple
from decimal import Decimal, ROUND_HALF_UP

import streamlit as st
import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, NamedStyle

# ===== Konfigūros =====
PVM_DEFAULT = Decimal("21.00")
FMT_MONEY = "#,##0.00"
FMT_QTY   = "#,##0.00"
TITLE_FONT_SIZE = 16  # „Atliktų paslaugų aktas“ pavadinimo dydis

# ===== Pagalbinės =====
def strip_accents(s: str) -> str:
    if s is None:
        return ""
    return "".join(ch for ch in unicodedata.normalize("NFD", str(s)) if unicodedata.category(ch) != "Mn")

def norm(s: str) -> str:
    return strip_accents(s).lower().strip()

def detect_delimiter(sample: str) -> Optional[str]:
    c_semi  = sample.count(";")
    c_comma = sample.count(",")
    return ";" if c_semi > c_comma else None

def dec2(v) -> float:
    return float(Decimal(str(v).replace(",", ".")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

def create_named_styles(wb: Workbook) -> None:
    """
    Suderinama su openpyxl 3.0.x ir 3.1.x:
    - wb.named_styles gali būti sąrašas objektų arba sąrašas eilučių.
    - Pridedame 'Money' ir 'Qty' tik jei jų nėra; saugiai ignoruojame dubliavimą.
    """
    existing: set = set()
    try:
        for s in wb.named_styles:
            # Jei s yra objektas -> s.name; jei s yra eilutė -> pati s
            nm = getattr(s, "name", s)
            existing.add(str(nm))
    except Exception:
        existing = set()

    if "Money" not in existing:
        stl = NamedStyle(name="Money")
        stl.number_format = FMT_MONEY
        stl.alignment = Alignment(horizontal="right")
        try:
            wb.add_named_style(stl)
        except Exception:
            pass  # jei jau yra, tyliai praleidžiam

    if "Qty" not in existing:
        stl = NamedStyle(name="Qty")
        stl.number_format = FMT_QTY
        stl.alignment = Alignment(horizontal="right")
        try:
            wb.add_named_style(stl)
        except Exception:
            pass

def set_borders(ws, rng: str, thick: bool = False) -> None:
    side = Side(style="thick" if thick else "thin")
    for row in ws[rng]:
        for c in row:
            c.border = Border(left=side, right=side, top=side, bottom=side)

def set_table_column_widths(ws) -> None:
    """Fiksuoti pločiai, kad stulpeliai neišsiplėstų neproporcingai ir parašai tilptų A–E ribose."""
    ws.column_dimensions["A"].width = 6    # Eil. Nr. – siauras
    ws.column_dimensions["B"].width = 48   # Paslaugos (wrap)
    ws.column_dimensions["C"].width = 10   # Kiekis
    ws.column_dimensions["D"].width = 12   # Įkainis
    ws.column_dimensions["E"].width = 14   # Suma

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

# ===== Duomenų paruošimas vienam adresui =====
def prepare_rows_for_address(
    dfa: pd.DataFrame,
    service_col: str,
    qty_col: str,
    rate_col: str,
    selected_services_for_addr: Optional[List[str]],
    group_same: bool
):
    dfp = dfa.assign(
        __service=dfa[service_col].astype(str).str.strip(),
        __qty=dfa[qty_col],
        __rate=dfa[rate_col],
    )
    if selected_services_for_addr:
        sel = {s.strip() for s in selected_services_for_addr}
        dfp = dfp[dfp["__service"].isin(sel)]

    dfp["__qty"]  = dfp["__qty"].apply(dec2)
    dfp["__rate"] = dfp["__rate"].apply(dec2)

    if dfp.empty:
        return []

    if group_same:
        dfg = (
            dfp.groupby(["__service", "__rate"], as_index=False)["__qty"]
               .sum()
               .sort_values(["__service", "__rate"], kind="stable")
        )
        return [(r["__service"], r["__qty"], r["__rate"]) for _, r in dfg.iterrows()]
    else:
        return [(r["__service"], r["__qty"], r["__rate"]) for _, r in dfp.iterrows()]

# ===== Vieno adreso blokas lape =====
def write_address_block(
    ws,
    start_row: int,
    address: str,
    department: str,
    rows: List[Tuple[str, float, float]],
) -> Tuple[int, str]:
    r = start_row
    # Data (fiksuota) + adresas
    ws[f"A{r}"] = "Atliktų paslaugų data: 2026 m."; r += 1
    addr_line = f"Objekto adresas: {address}"
    if department:
        addr_line += f", Skyrius: {department}"
    ws[f"A{r}"] = addr_line; r += 1
    ws[f"A{r}"] = ""; r += 1  # tarpas prieš lentelę

    # Lentele header
    ws[f"A{r}"], ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"], ws[f"E{r}"] = \
        "Eil. Nr.", "Paslaugos pavadinimas", "Kiekis", "Įkainis (be PVM)", "Suma (be PVM)"
    set_borders(ws, f"A{r}:E{r}", thick=True)
    for c in (f"A{r}", f"B{r}", f"C{r}", f"D{r}", f"E{r}"):
        ws[c].font = Font(bold=True)
    r += 1

    first_data_row = r
    for i, (service, qty, rate) in enumerate(rows, start=1):
        ws[f"A{r}"] = i
        ws[f"B{r}"] = service
        ws[f"C{r}"] = qty
        ws[f"D{r}"] = rate
        ws[f"E{r}"] = f"=C{r}*D{r}"

        # Formatai ir lygiavimai
        ws[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")  # wrap paslaugoms
        ws[f"C{r}"].number_format = FMT_QTY
        ws[f"D{r}"].number_format = FMT_MONEY
        ws[f"E{r}"].number_format = FMT_MONEY
        r += 1

    last_data_row = r - 1
    # Subtotal
    ws[f"D{r}"] = "Suma (be PVM):"
    ws[f"E{r}"] = f"=SUM(E{first_data_row}:E{last_data_row})" if last_data_row >= first_data_row else 0
    ws[f"E{r}"].number_format = FMT_MONEY
    set_borders(ws, f"D{r}:E{r}", thick=True)
    subtotal_ref = f"E{r}"
    r += 1

    # PVM + su PVM
    ws[f"D{r}"] = f"PVM {float(PVM_DEFAULT)}%:"
    ws[f"E{r}"] = f"={subtotal_ref}*{float(PVM_DEFAULT)/100}"
    ws[f"E{r}"].number_format = FMT_MONEY
    r += 1

    ws[f"D{r}"] = "Suma su PVM:"
    ws[f"E{r}"] = f"={subtotal_ref}+E{r-1}"
    ws[f"E{r}"].number_format = FMT_MONEY
    set_borders(ws, f"D{r-2}:E{r}", thick=True)
    r += 2  # tarpas prieš kitą bloką

    return r, subtotal_ref

# ===== AKTO kūrimas – keli adresai viename lape =====
def build_workbook_act_multi(
    df: pd.DataFrame,
    col_map: Dict[str, Optional[str]],
    contract_no: str,
    manager: Optional[str],
    addresses_selected: List[str],
    service_selection_map: Dict[str, List[str]],
    akto_data: str,
    group_same: bool,
) -> Workbook:
    """
    Generuoja lapą su viršuje: data, Užsakovas, Vykdytojas, pavadinimas, Sutarties nr.
    Toliau – blokai kiekvienam adresui (su per-adresą paslaugų filtrais),
    tada bendra suma ir parašų blokas A–E ribose.
    """
    # Filtras pagal sutartį ir (jei duotas) vadybininką
    dfc = df[df[col_map["sutartis"]].astype(str).str.strip() == contract_no].copy() if col_map.get("sutartis") else df.copy()
    if manager and col_map.get("vadybininkas") and manager != "(visi)":
        dfc = dfc[dfc[col_map["vadybininkas"]].astype(str).str.strip() == manager]

    if dfc.empty:
        raise ValueError("Pagal pasirinktą užsakovą/sutartį/vadybininką įrašų nerasta.")

    # Patvirtiname, kad visi adresai yra dfc
    base_addrs = set(dfc[col_map["address"]].astype(str).str.strip())
    if not addresses_selected:
        raise ValueError("Neparinkti adresai.")
    for a in addresses_selected:
        if a.strip() not in base_addrs:
            raise ValueError(f"Adresas '{a}' nepriklauso pasirinktai sutarčiai/vadybininkui.")

    # Meta (užsakovas/vykdytojas) – imame iš pirmos linijos
    uzsakovas = str(dfc.iloc[0][col_map["uzsakovas"]]).strip() if col_map.get("uzsakovas") else ""
    vykdytojas = str(dfc.iloc[0][col_map["vykdytojas"]]).strip() if col_map.get("vykdytojas") else ""

    # Paruošiam workbook
    wb = Workbook()
    create_named_styles(wb)
    ws = wb.active
    ws.title = "AKTAS"
    set_table_column_widths(ws)

    # ===== Viršutinė antraštė =====
    ws["A1"] = akto_data
    ws["A2"] = ""

    ws["A3"] = f"Užsakovas: {uzsakovas}"
    ws["A4"] = f"Vykdytojas: {vykdytojas}"

    ws.merge_cells("A5:E5")
    ws["A5"] = "Atliktų paslaugų aktas"
    ws["A5"].font = Font(bold=True, size=TITLE_FONT_SIZE)
    ws["A5"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 24

    ws["A6"] = ""
    ws["A7"] = f"Sutarties nr.: {contract_no}"
    ws["A8"] = ""

    # ===== Kiekvieno adreso blokai =====
    r = 9
    subtotal_refs: List[str] = []

    for addr in addresses_selected:
        dfa = dfc[dfc[col_map["address"]].astype(str).str.strip() == addr].copy()
        department = str(dfa.iloc[0][col_map["skyrius"]]).strip() if col_map.get("skyrius") and not dfa.empty else ""
        # Pasirinktos paslaugos šiam adresui
        selected_for_addr = service_selection_map.get(addr, [])
        rows = prepare_rows_for_address(dfa, col_map["service"], col_map["qty"], col_map["rate"], selected_for_addr, group_same)
        if not rows:
            # jei po filtrų šiam adresui nieko neliko – praleidžiam adresą
            continue
        r, subref = write_address_block(ws, r, addr, department, rows)
        subtotal_refs.append(subref)

    # Jei nieko neįrašėm – grąžinam aiškią klaidą
    if not subtotal_refs:
        raise ValueError("Po pasirinktų filtrų/paslaugų adresams nėra jokių eilučių aktui.")

    # ===== Bendra suma iš visų adresų =====
    ws[f"D{r}"] = "Bendra suma (be PVM):"
    ws[f"E{r}"] = f"=SUM({','.join(subtotal_refs)})"
    ws[f"E{r}"].number_format = FMT_MONEY
    set_borders(ws, f"D{r}:E{r}", thick=True)
    r += 1

    ws[f"D{r}"] = f"PVM {float(PVM_DEFAULT)}%:"
    ws[f"E{r}"] = f"=E{r-1}*{float(PVM_DEFAULT)/100}"
    ws[f"E{r}"].number_format = FMT_MONEY
    r += 1

    ws[f"D{r}"] = "Bendra suma su PVM:"
    ws[f"E{r}"] = f"=E{r-2}+E{r-1}"
    ws[f"E{r}"].number_format = FMT_MONEY
    set_borders(ws, f"D{r-2}:E{r}", thick=True)
    r += 2  # tarpas prieš parašų bloką

    # ===== PARAŠŲ BLOKAS APAČIOJE – KAIRĖ/DEŠINĖ (A–E, be datų) =====
    # Kairė: Užsakovas (A–C), Dešinė: Vykdytojas (D–E)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)  # A-C
    ws.cell(row=r, column=1).value = "Užsakovas:"
    ws.cell(row=r, column=1).font  = Font(bold=True)

    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)  # D-E
    ws.cell(row=r, column=4).value = "Vykdytojas:"
    ws.cell(row=r, column=4).font  = Font(bold=True)

    # Parašo linijos
    for c in (1,2,3):  # A-C
        ws.cell(row=r+1, column=c).border = Border(bottom=Side(style="thin"))
    for c in (4,5):    # D-E
        ws.cell(row=r+1, column=c).border = Border(bottom=Side(style="thin"))
    ws.row_dimensions[r+1].height = 22

    # Aprašai po linijomis
    ws.merge_cells(start_row=r+2, start_column=1, end_row=r+2, end_column=3)
    ws.cell(row=r+2, column=1).value = "Vardas, pavardė / Pareigos"

    ws.merge_cells(start_row=r+2, start_column=4, end_row=r+2, end_column=5)
    ws.cell(row=r+2, column=4).value = "Vardas, pavardė / Pareigos"

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

# ===== UŽSAKOVO FILTRAS (pirma – kaip prašei) =====
selected_customers: List[str] = []
if col_map.get("uzsakovas"):
    uzs_list = sorted(set(str(x).strip() for x in df[col_map["uzsakovas"]].dropna() if str(x).strip()))
    selected_customers = st.multiselect("Užsakovas(-ai)", uzs_list)
    if selected_customers:
        df = df[df[col_map["uzsakovas"]].astype(str).str.strip().isin(selected_customers)]
else:
    st.warning("Kataloge nėra „Užsakovas“ stulpelio – filtras nebus taikomas.")

# ===== SUTARTIES NUMERIO PASIRINKIMAS (pagal pasirinktą užsakovą) =====
contract_no = ""
if col_map.get("sutartis"):
    sut_list = sorted(set(str(x).strip() for x in df[col_map["sutartis"]].dropna() if str(x).strip()))
    if not sut_list:
        st.error("Pasirinktam(-iems) užsakovui(-ams) sutarties numerių nerasta.")
        st.stop()
    contract_no = st.selectbox("Sutarties numeris (pagal pasirinktą užsakovą)", sut_list, index=0)
    st.info(f"Pasirinkta sutartis: **{contract_no}**")
    # filtruojame pagal sutartį
    df = df[df[col_map["sutartis"]].astype(str).str.strip() == contract_no]
else:
    contract_no = st.text_input("Sutarties numeris (nėra stulpelio, suvesk ranka)", "")
    if not contract_no:
        st.warning("Įvesk sutarties numerį arba įkelk katalogą su „Sutarties numeris“ stulpeliu.")
    else:
        st.info(f"Įvestas sutarties numeris: **{contract_no}**")

# ===== VADYBININKO FILTRAS (pasirenkamas) =====
manager = None
if col_map.get("vadybininkas"):
    with st.expander("Filtras pagal vadybininką (pasirenkama)", expanded=False):
        mgr_list = sorted(set(str(x).strip() for x in df[col_map["vadybininkas"]].dropna() if str(x).strip()))
        if mgr_list:
            manager = st.selectbox("Vadybininkas (jei reikia)", ["(visi)"] + mgr_list, index=0)
            if manager != "(visi)":
                df = df[df[col_map["vadybininkas"]].astype(str).str.strip() == manager]
        else:
            st.caption("Šiai sutarčiai neturi priskirtų vadybininkų arba laukas tuščias.")

# ===== GLOBALUS PASLAUGŲ FILTRAS (multiselect) =====
global_services = sorted(set(str(x).strip() for x in df[col_map["service"]].dropna() if str(x).strip()))
with st.expander("Globalus filtras pagal paslaugas (pasirenkama)", expanded=False):
    sel_global_services = st.multiselect("Paslaugos (apribos visus adresus)", global_services)
    if sel_global_services:
        df = df[df[col_map["service"]].astype(str).str.strip().isin(sel_global_services)]

# ===== Adresų pasirinkimas (multiselect; po visų filtrų) =====
addresses = sorted(set(str(x).strip() for x in df[col_map["address"]].dropna() if str(x).strip()))
if not addresses:
    st.error("Po filtrų nerasta adresų.")
    st.stop()

addresses_selected = st.multiselect("Adresai (galima pasirinkti kelis)", addresses, default=addresses)

# ===== Per-adresą paslaugų pasirinkimas + skyrių rodymas (prieš generavimą) =====
service_selection_map: Dict[str, List[str]] = {}
if addresses_selected:
    st.markdown("**Pasirinktų adresų skyriai ir paslaugos:**")
    for addr in addresses_selected:
        dfa = df[df[col_map["address"]].astype(str).str.strip() == addr]
        dep = (str(dfa.iloc[0][col_map["skyrius"]]).strip() if col_map.get("skyrius") and not dfa.empty else "")
        st.caption(f"Adresas: **{addr}** — Skyrius: _{dep or 'nenurodytas'}_")
        services_for_addr = sorted(set(dfa[col_map["service"]].astype(str).str.strip()))
        chosen = st.multiselect(f"Paslaugos šiam adresui — {addr}", services_for_addr, default=services_for_addr)
        service_selection_map[addr] = chosen

# ===== Group by pasirinkimas =====
group_same = st.checkbox("Sujungti vienodas paslaugas (grupuoti pagal pavadinimą ir įkainį)", value=True)

# ===== Data viršuje (be teksto) =====
akto_data = st.text_input("Data (rodoma viršuje)", "2026-01-04")

# ===== Generavimas =====
btn = st.button("Generuoti AKTĄ (XLSX)", use_container_width=True, disabled=not bool(addresses_selected))
if btn:
    try:
        wb = build_workbook_act_multi(df_raw, col_map, contract_no or "", manager, addresses_selected,
                                      service_selection_map, akto_data, group_same)
        bio = io.BytesIO(); wb.save(bio); xlsx = bio.getvalue()
        st.download_button("Atsisiųsti AKTĄ", xlsx, "aktas_vienas_lapas.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.exception(e)
