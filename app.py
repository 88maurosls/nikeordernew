import streamlit as st
import pandas as pd
from io import BytesIO
import re
import os

# --------------------------------
# Utils
# --------------------------------
def extract_order_id(filename: str) -> str:
    match = re.search(r"_(\d+)_", filename)
    return match.group(1) if match else ""

def read_order_xlsx(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0, header=None, dtype=object)
        return df.fillna("")
    except Exception as e:
        st.error(f"Errore lettura file: {e}")
        return None

def find_col(row, label: str):
    for i, v in enumerate(row):
        if isinstance(v, str) and v.strip() == label:
            return i
    return None

def to_int(x) -> int:
    try:
        return int(float(str(x).replace(",", ".")))
    except:
        return 0

def to_money(x) -> float:
    """
    Converte stringhe tipo:
    '55,00 €' -> 55.00
    '1.770,00 €' -> 1770.00
    Se non parsabile -> 0.0
    """
    try:
        s = str(x)
        s = s.replace("\xa0", " ").strip()
        s = s.replace("€", "").strip()
        s = s.replace(".", "")          # separatore migliaia
        s = s.replace(",", ".")         # decimali
        s = re.sub(r"\s+", "", s)
        return float(s) if s else 0.0
    except:
        return 0.0

# --------------------------------
# Parser nuovo template Nike
# --------------------------------
def process_order_details(df: pd.DataFrame, order_id: str, view_option: str):
    rows = []

    current_model = None
    model_name = None
    color_desc = None
    product_type = None
    whs_val = None
    rtl_val = None

    col_size = col_upc = None
    col_richiesti = col_aperti = col_spediti = None
    in_table = False

    for _, r in df.iterrows():
        row = r.tolist()

        # Nuovo articolo
        idx = find_col(row, "Modello/Colore:")
        if idx is not None:
            current_model = str(row[idx + 1]).strip() if idx + 1 < len(row) else ""
            model_name = color_desc = product_type = None
            whs_val = rtl_val = None
            in_table = False

            # A volte WHS è già sulla stessa riga
            idx_whs = find_col(row, "All'ingrosso:")
            if idx_whs is not None and idx_whs + 1 < len(row):
                whs_val = to_money(row[idx_whs + 1])
            continue

        # Metadati
        idx = find_col(row, "Nome modello:")
        if idx is not None:
            model_name = str(row[idx + 1]).strip() if idx + 1 < len(row) else ""

            # A volte RTL è sulla stessa riga
            idx_rtl = find_col(row, "Retail consigliato:")
            if idx_rtl is not None and idx_rtl + 1 < len(row):
                rtl_val = to_money(row[idx_rtl + 1])
            continue

        idx = find_col(row, "Descrizione colore:")
        if idx is not None:
            color_desc = str(row[idx + 1]).strip() if idx + 1 < len(row) else ""
            continue

        idx = find_col(row, "Tipo di prodotto:")
        if idx is not None:
            product_type = str(row[idx + 1]).strip() if idx + 1 < len(row) else ""
            continue

        # Nel caso WHS/RTL fossero su righe dedicate
        idx = find_col(row, "All'ingrosso:")
        if idx is not None and current_model:
            whs_val = to_money(row[idx + 1]) if idx + 1 < len(row) else 0.0
            continue

        idx = find_col(row, "Retail consigliato:")
        if idx is not None and current_model:
            rtl_val = to_money(row[idx + 1]) if idx + 1 < len(row) else 0.0
            continue

        # Header tabella taglie
        if isinstance(row[0], str) and row[0].strip() == "Misura":
            in_table = True
            col_size = 0
            col_upc = 1
            col_richiesti = col_aperti = col_spediti = None

            for i, v in enumerate(row):
                if isinstance(v, str):
                    label = v.strip()
                    if label == "Richiesti:":
                        col_richiesti = i
                    if label == "Aperti:":
                        col_aperti = i
                    if label == "Spediti:":
                        col_spediti = i
            continue

        if not in_table or not current_model:
            continue

        # Fine tabella
        if isinstance(row[0], str) and str(row[0]).startswith("Qtà totale"):
            in_table = False
            continue

        # Riga taglia
        size = str(row[col_size]).strip() if col_size is not None and col_size < len(row) else ""
        if not size:
            continue

        upc = str(row[col_upc]).strip() if col_upc is not None and col_upc < len(row) else ""

        richiesti = to_int(row[col_richiesti]) if col_richiesti is not None and col_richiesti < len(row) else 0
        aperti = to_int(row[col_aperti]) if col_aperti is not None and col_aperti < len(row) else 0
        spediti = to_int(row[col_spediti]) if col_spediti is not None and col_spediti < len(row) else 0

        rows.append([
            current_model,
            size,
            richiesti,
            aperti,
            spediti,
            model_name or "",
            color_desc or "",
            upc,
            product_type or "",
            order_id,
            float(whs_val) if whs_val is not None else 0.0,
            float(rtl_val) if rtl_val is not None else 0.0,
        ])

    df_final = pd.DataFrame(
        rows,
        columns=[
            "Modello/Colore", "Misura", "Richiesti", "Aperti", "Spediti",
            "Nome del modello", "Descrizione colore",
            "Codice a Barre (UPC)", "Tipo di prodotto", "ID_ORDINE",
            "WHS", "RTL"
        ]
    )

    if df_final.empty:
        return None, df_final

    # Codice e Colore
    df_final["Codice"] = df_final["Modello/Colore"].apply(lambda x: str(x).split("-")[0])
    df_final["Colore"] = df_final["Modello/Colore"].apply(
        lambda x: str(x).split("-")[1] if "-" in str(x) else ""
    )

    # Rimuove righe completamente vuote su tutte e tre
    df_final = df_final[
        (df_final["Richiesti"] != 0) | (df_final["Aperti"] != 0) | (df_final["Spediti"] != 0)
    ]

    # Filtro per vista richiesta dall'utente
    qty_col = view_option  # "Richiesti" | "Aperti" | "Spediti"
    df_final = df_final[df_final[qty_col] > 0]

    # Colonne base (WHS/RTL sempre in fondo, ultime in assoluto)
    base_cols = [
        "Modello/Colore", "Descrizione colore", "Codice",
        "Nome del modello", "Tipo di prodotto",
        "Colore", "Misura", "Codice a Barre (UPC)", "ID_ORDINE",
        qty_col, "WHS", "RTL"
    ]

    df_final = df_final[base_cols]

    # Export Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False)

    return output.getvalue(), df_final

# --------------------------------
# UI Streamlit
# --------------------------------
st.title("Nike order details v2")

uploaded_file = st.file_uploader(
    "Carica il file XLSX (Order Details)",
    type="xlsx"
)

if uploaded_file:
    filename = os.path.splitext(uploaded_file.name)[0]
    order_id = st.text_input("ID_ORDINE", extract_order_id(filename))

    # Radio con le tre viste richieste
    view_option_ui = st.radio(
        "Seleziona vista",
        ("RICHIESTI", "APERTI", "SPEDITI"),
        index=1
    )

    # Mappa UI -> nome colonna df
    view_map = {
        "RICHIESTI": "Richiesti",
        "APERTI": "Aperti",
        "SPEDITI": "Spediti"
    }
    view_option = view_map[view_option_ui]

    df = read_order_xlsx(uploaded_file)

    if df is not None and st.button("Elabora"):
        file_out, preview = process_order_details(df, order_id, view_option)

        if file_out is None:
            st.warning("Nessun dato trovato")
        else:
            st.dataframe(preview, use_container_width=True)

            st.download_button(
                "Scarica Excel",
                data=file_out,
                file_name=f"{filename}_processed.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
