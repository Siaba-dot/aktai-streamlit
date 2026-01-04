# app.py
import io
import re
from datetime import date

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Aktų generatorius", page_icon="📄", layout="wide")
st.caption("build: v2026-01-04-12:10")

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
]

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
    name = re.sub(r"\s+", " ", name)
    return name[:120]

@st.cache_data
def read_excel_to_df(file_bytes: bytes) -> pd.DataFrame:
    xl = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    df = xl.parse(xl.sheet_names[0])
    df.columns = [str(c).strip() for c in df.columns]
    return df

def validate_cols(df: pd.DataFrame):
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    return missing

def df_to_items(g: pd.DataFrame) -> pd.DataFrame:
    cols = ["Paslaugos pavadinimas", "Plotas (m2)", "Įkainis (Eur be PVM)", "Suma"]
    items = g.copy()
    for c in cols:
        if c not in items.columns:
            items[c] = "" if c == "Paslaugos pavadinimas" else 0.0
    for c in ["Plotas (m2)", "Įkainis (Eur be PVM)", "Suma"]:
        items[c] = pd.to_numeric(items[c], errors="coerce").fillna(0.0)
    items["Paslaugos pavadinimas"] = items["Paslaugos pavadinimas"].astype(str)
    return items[cols]

def render_header(ws, wb, start_row, meta: dict):
    bold = wb.add_format({"bold": True})
    row = start_row
    pairs = [
        ("Užsakovas", meta.get("Užsakovas", "")),
        ("Tiekėjas", meta.get("Vykdytojas", "")),
        ("Sutarties nr.", meta.get("Sutarties numeris", "")),
        ("Objekto adresas", meta.get("Objekto adresas", "")),
        ("Skyrius", meta.get("Skyrius", "")),
        ("Atlikimo data", meta.get("Atlikimo data", "")),
        ("Atlikimo laikotarpis", meta.get("Atlikimo laikotarpis", "")),
    ]
    for label, val in pairs:
        ws.write(row, 0, label, bold)
        ws.write(row, 1, val)
        row += 1
    return row + 1

def write_act_to_sheet(wb, sheet_name: str, meta: dict, items: pd.DataFrame, pvm_pct: float, show_pvm: bool):
    ws = wb.add_worksheet(sheet_name[:31])
    ws.set_column(0, 0, 5)
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
    for i, row in enumerate(items.to_dict("records"), start=1):
        ws.write(start + i - 1, 0, i, text_fmt)
        ws.write(start + i - 1, 1, row.get("Paslaugos pavadinimas", ""), text_fmt)
        ws.write_number(start + i - 1, 2, float(row.get("Plotas (m2)", 0.0)), num_fmt)
        ws.write_number(start + i - 1, 3, float(row.get("Įkainis (Eur be PVM)", 0.0)), num_fmt)
        ws.write_number(start + i - 1, 4, float(row.get("Suma", 0.0)), num_fmt)

    last_row = start + len(items) - 1
    suma_range = f"E{start+1}:E{last_row+1}"
    total_row = last_row + 2
    bold     = wb.add_format({"bold": True})
    bold_num = wb.add_format({"bold": True, "num_format": "#,##0.00"})
    ws.write(total_row, 3, "Suma (be PVM):", bold)
    ws.write_formula(total_row, 4, f"=SUM({suma_range})", bold_num)

    if show_pvm:
        pvm_row = total_row + 1
        ws.write(pvm_row, 3, f"PVM {pvm_pct:.2f}%:", bold)
        ws.write_formula(pvm_row, 4, f"=E{total_row+1}*{pvm_pct/100.0}", bold_num)
        suma_su_pvm_row = pvm_row + 1
        ws.write(suma_su_pvm_row, 3, "Suma su PVM:", bold)
        ws.write_formula(suma_su_pvm_row, 4, f"=E{total_row+1}+E{pvm_row+1}", bold_num)

def build_act_filename(meta: dict) -> str:
    base = f"AKTAS_{meta.get('Užsakovas','')}_{meta.get('Sutarties numeris','')}"
    return sanitize_filename(base) + ".xlsx"

def generate_acts_zip_in_memory(df: pd.DataFrame, pvm_pct: float, show_pvm: bool, single_file: bool, atlikimo_data: str, atlikimo_laikotarpis: str) -> bytes:
    import zipfile
    grp_cols = ["Užsakovas", "Sutarties numeris", "Objekto adresas"]
    groups = df.groupby(grp_cols, dropna=False)
    zip_buf = io.BytesIO()

    if single_file:
        xlsx_buf = io.BytesIO()
        with pd.ExcelWriter(xlsx_buf, engine="xlsxwriter") as writer:
            wb = writer.book
            for (uzs, sut, addr), g in groups:
                items = df_to_items(g)
                first = g.iloc[0].to_dict()
                meta = {
                    "Užsakovas": uzs,
                    "Vykdytojas": first.get("Vykdytojas",""),
                    "Sutarties numeris": sut,
                    "Objekto adresas": addr,
                    "Skyrius": first.get("Skyrius",""),
                    "Atlikimo data": atlikimo_data,
                    "Atlikimo laikotarpis": atlikimo_laikotarpis
                }
                sheet_name = sanitize_filename(f"{uzs} [{sut}]")[:31]
                write_act_to_sheet(wb, sheet_name, meta, items, pvm_pct, show_pvm)
        with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr("AKTAI_VIENAME.xlsx", xlsx_buf.getvalue())
    else:
        with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as z:
            for (uzs, sut, addr), g in groups:
                xlsx_buf = io.BytesIO()
                with pd.ExcelWriter(xlsx_buf, engine="xlsxwriter") as writer:
                    wb = writer.book
                    items = df_to_items(g)
                    first = g.iloc[0].to_dict()
                    meta = {
                        "Užsakovas": uzs,
                        "Vykdytojas": first.get("Vykdytojas",""),
                        "Sutarties numeris": sut,
                        "Objekto adresas": addr,
                        "Skyrius": first.get("Skyrius",""),
                        "Atlikimo data": atlikimo_data,
                        "Atlikimo laikotarpis": atlikimo_laikotarpis
                    }
                    write_act_to_sheet(wb, "AKTAS", meta, items, pvm_pct, show_pvm)
                z.writestr(build_act_filename(meta), xlsx_buf.getvalue())

    zip_buf.seek(0)
    return zip_buf.getvalue()

# ------------------------------------------------------------
# Pagrindinis srautas
# ------------------------------------------------------------
if uploaded:
    file_bytes = uploaded.read()
    df = read_excel_to_df(file_bytes)

    missing = validate_cols(df)
    if missing:
        st.error(f"Trūksta stulpelių: {', '.join(missing)}")
        st.stop()

    atlikimo_data = st.text_input("Atlikimo data (YYYY-MM-DD)")
    atlikimo_laikotarpis = st.text_input("Atlikimo laikotarpis (pvz. 2026-01)")

    st.dataframe(df.head(20), use_container_width=True)

    if st.button("🧾 Generuoti aktus (ZIP)"):
        zip_bytes = generate_acts_zip_in_memory(
            df, pvm_tarifas, rodyti_pvm, sujungti_i_viena_faila,
            atlikimo_data, atlikimo_laikotarpis
        )
        st.download_button(
            "⬇️ Parsisiųsti aktus (ZIP)",
            data=zip_bytes,
            file_name="AKTAI.zip",
            mime="application/zip",
        )
else:
    st.info("Įkelk Excel failą, tada įrašyk datas ir spausk „Generuoti aktus“.")
