
# app.py
# Streamlit Cloud Atliktų darbų aktų generatorius (tik Excel)
# Paleidimas Cloud'e: "Deploy from GitHub" -> app.py

import io
import re
from datetime import date

import numpy as np
import pandas as pd
import streamlit as st

REQUIRED_COLS = [
    "Skyrius",
    "Objekto adresas",
    "Vykdytojas",
    "Užsakovas",
    "Sutarties numeris",
    "Paslaugos pavadinimas",
    "Plotas (m2)",
    "Įkainis (Eur be PVM)",
    "Suma",
    "Nuoroda į VVS",
    "Galioja iki",
    "Vadybininkas",
]

st.set_page_config(page_title="Aktų generatorius", page_icon="📄", layout="wide")
st.title("📄 Atliktų darbų aktų generatorius (Streamlit Cloud)")

with st.sidebar:
    st.markdown("### Parametrai")
    pvm_tarifas = st.number_input("PVM tarifas, %", min_value=0.0, max_value=100.0, value=21.0, step=0.5)
    rodyti_pvm = st.checkbox("Rodyti PVM eilutę (sumoje)", value=True)
    sujungti_i_viena_faila = st.checkbox("Sujungti viską į vieną .xlsx (atskiri lapai)", value=False)
    st.markdown("---")
    st.caption("Įkelk Excel su stulpeliais kaip šablone. Generavimas vyksta atmintyje.")

uploaded = st.file_uploader("Įkelk Excel (.xlsx) su duomenimis", type=["xlsx"])

def sanitize_filename(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r'[\\/*?:"<>|]', "_", name)
    name = re.sub(r"\\s+", " ", name)
    return name[:120]

@st.cache_data(show_spinner=False)
def read_excel_to_df(file_bytes: bytes) -> pd.DataFrame:
    # Skaitymas iš baitų (Cloud-friendly).
    xl = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    df = xl.parse(xl.sheet_names[0])
    df.columns = [str(c).strip() for c in df.columns]
    if "Galioja iki" in df.columns:
        df["Galioja iki"] = pd.to_datetime(df["Galioja iki"], errors="coerce")
    if "Vadybininkas" in df.columns:
        df["Vadybininkas"] = df["Vadybininkas"].fillna("")
    return df

def validate_cols(df: pd.DataFrame):
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    return missing

def render_header(ws, wb, start_row, meta: dict):
    bold = wb.add_format({"bold": True})
    date_fmt = wb.add_format({"num_format": "yyyy-mm-dd"})
    row = start_row
    pairs = [
        ("Užsakovas", meta.get("Užsakovas", "")),
        ("Vykdytojas", meta.get("Vykdytojas", "")),
        ("Sutarties nr.", meta.get("Sutarties numeris", "")),
        ("Objektas / adresas", meta.get("Objekto adresas", "")),
        ("Skyrius", meta.get("Skyrius", "")),
        ("Vadybininkas", meta.get("Vadybininkas", "")),
        ("Akto data", meta.get("Akto data", date.today())),
    ]
    for label, val in pairs:
        ws.write(row, 0, label, bold)
        if isinstance(val, (pd.Timestamp, date)):
            ws.write_datetime(row, 1, pd.Timestamp(val).to_pydatetime(), date_fmt)
        else:
            ws.write(row, 1, val)
        row += 1
    return row + 1

def write_act_to_sheet(wb, sheet_name: str, meta: dict, items: pd.DataFrame, pvm_pct: float, show_pvm: bool):
    ws = wb.add_worksheet(sheet_name[:31])
    ws.set_column(0, 0, 20)
    ws.set_column(1, 1, 60)
    ws.set_column(2, 4, 15)

    end_header_row = render_header(ws, wb, 0, meta)

    hdr_fmt  = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})
    num_fmt  = wb.add_format({"num_format": "#,##0.00", "border": 1})
    text_fmt = wb.add_format({"border": 1})

    table_headers = ["Eil. Nr.", "Paslaugos pavadinimas", "Kiekis", "Įkainis (be PVM)", "Suma (be PVM)"]
    for col, h in enumerate(table_headers):
        ws.write(end_header_row, col, h, hdr_fmt)

    start = end_header_row + 1
    for i, row in enumerate(items.itertuples(index=False), start=1):
        ws.write(start + i - 1, 0, i, text_fmt)
        ws.write(start + i - 1, 1, getattr(row, "Paslaugos pavadinimas"), text_fmt)
        ws.write_number(start + i - 1, 2, float(getattr(row, "Plotas (m2)")), num_fmt)
        ws.write_number(start + i - 1, 3, float(getattr(row, "Įkainis (Eur be PVM)")), num_fmt)
        ws.write_number(start + i - 1, 4, float(getattr(row, "Suma")), num_fmt)

    last_row = start + len(items) - 1
    suma_range = f"E{start+1}:E{last_row+1}"
    total_row = last_row + 2
    bold     = wb.add_format({"bold": True})
    bold_num = wb.add_format({"bold": True, "num_format": "#,##0.00"})

    ws.write(total_row, 3, "Suma (be PVM):", bold)
    ws.write_formula(total_row, 4, f"=SUM({suma_range})", bold_num)

    suma_su_pvm_row = total_row
    if show_pvm:
        pvm_row = total_row + 1
        ws.write(pvm_row, 3, f"PVM {pvm_pct:.2f}%:", bold)
        ws.write_formula(pvm_row, 4, f"=E{total_row+1}*{pvm_pct/100.0}", bold_num)
        suma_su_pvm_row = pvm_row + 1
        ws.write(suma_su_pvm_row, 3, "Suma su PVM:", bold)
        ws.write_formula(suma_su_pvm_row, 4, f"=E{total_row+1}+E{pvm_row+1}", bold_num)

    if isinstance(meta.get("Nuoroda į VVS"), str) and meta["Nuoroda į VVS"] and meta["Nuoroda į VVS"] != "All VVS":
        ws.write(suma_su_pvm_row + 2, 0, f"VVS: {meta['Nuoroda į VVS']}")

def build_act_filename(meta: dict) -> str:
    base = f"AKTAS_{meta.get('Užsakovas','')}_{meta.get('Sutarties numeris','')}_{meta.get('Objekto adresas','')}"
    name = sanitize_filename(base) + ".xlsx"
    return name

def df_to_items(g: pd.DataFrame) -> pd.DataFrame:
    return g[["Paslaugos pavadinimas", "Plotas (m2)", "Įkainis (Eur be PVM)", "Suma"]].fillna(0).copy()

def generate_acts_zip_in_memory(df: pd.DataFrame, pvm_pct: float, show_pvm: bool, single_file: bool) -> bytes:
    grp_cols = ["Užsakovas", "Sutarties numeris", "Objekto adresas"]
    groups = df.groupby(grp_cols, dropna=False)

    zip_buf = io.BytesIO()

    if single_file:
        xlsx_buf = io.BytesIO()
        with pd.ExcelWriter(xlsx_buf, engine="xlsxwriter") as writer:
            wb = writer.book
            for (uzs, sut, addr), g in groups:
                items = df_to_items(g)
                first = g.iloc[0]
                meta = {
                    "Užsakovas": uzs,
                    "Vykdytojas": first.get("Vykdytojas", ""),
                    "Sutarties numeris": sut,
                    "Objekto adresas": addr,
                    "Skyrius": first.get("Skyrius", ""),
                    "Vadybininkas": first.get("Vadybininkas", ""),
                    "Akto data": date.today(),
                    "Nuoroda į VVS": first.get("Nuoroda į VVS", ""),
                }
                sheet_name = sanitize_filename(f"{uzs} [{sut}]")[:31]
                write_act_to_sheet(wb, sheet_name, meta, items, pvm_pct, show_pvm)
        import zipfile
        with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr("AKTAI_VIENAME.xlsx", xlsx_buf.getvalue())
    else:
        import zipfile
        with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as z:
            for (uzs, sut, addr), g in groups:
                xlsx_buf = io.BytesIO()
                with pd.ExcelWriter(xlsx_buf, engine="xlsxwriter") as writer:
                    wb = writer.book
                    items = df_to_items(g)
                    first = g.iloc[0]
                    meta = {
                        "Užsakovas": uzs,
                        "Vykdytojas": first.get("Vykdytojas", ""),
                        "Sutarties numeris": sut,
                        "Objekto adresas": addr,
                        "Skyrius": first.get("Skyrius", ""),
                        "Vadybininkas": first.get("Vadybininkas", ""),
                        "Akto data": date.today(),
                        "Nuoroda į VVS": first.get("Nuoroda į VVS", ""),
                    }
                    write_act_to_sheet(wb, "AKTAS", meta, items, pvm_pct, show_pvm)
                z.writestr(build_act_filename(meta), xlsx_buf.getvalue())

    zip_buf.seek(0)
    return zip_buf.getvalue()

if uploaded:
    file_bytes = uploaded.read()
    df = read_excel_to_df(file_bytes)

    missing = validate_cols(df)
    if missing:
        st.error(f"Trūksta stulpelių: {', '.join(missing)}")
        st.stop()

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        uzs_f = st.multiselect("Filtruoti pagal Užsakovą", sorted(df["Užsakovas"].dropna().astype(str).unique().tolist()))
    with col2:
        vdyb_f = st.multiselect("Filtruoti pagal Vadybininką", sorted(df["Vadybininkas"].dropna().astype(str).unique().tolist()))
    with col3:
        sky_f = st.multiselect("Filtruoti pagal Skyrių", sorted(df["Skyrius"].dropna().astype(str).unique().tolist()))
    with col4:
        data_nuo = st.date_input("Galioja iki ≥ (pasirinktinai)", value=None)

    dff = df.copy()
    if uzs_f:
        dff = dff[dff["Užsakovas"].astype(str).isin(uzs_f)]
    if vdyb_f:
        dff = dff[dff["Vadybininkas"].astype(str).isin(vdyb_f)]
    if sky_f:
        dff = dff[dff["Skyrius"].astype(str).isin(sky_f)]
    if data_nuo:
        dff = dff[(~dff["Galioja iki"].isna()) & (dff["Galioja iki"].dt.date >= data_nuo)]

    st.success(f"Eilučių po filtrų: {len(dff)}")
    st.dataframe(dff.head(20), use_container_width=True)

    if len(dff) > 0 and st.button("🧾 Generuoti aktus (ZIP)"):
        zip_bytes = generate_acts_zip_in_memory(dff, pvm_tarifas, rodyti_pvm, single_file=sujungti_i_viena_faila)
        st.download_button(
            "⬇️ Parsisiųsti aktus (ZIP)",
            data=zip_bytes,
            file_name="AKTAI.zip",
            mime="application/zip",
        )
else:
    st.info("Įkelk Excel failą, tada parink filtrus ir spausk „Generuoti aktus“.")
