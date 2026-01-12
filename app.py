import streamlit as st
import pandas as pd
from io import BytesIO
import re
import os

# --------------------------------
# Utils
# --------------------------------
def extract_order_id(filename):
    match = re.search(r"_(\d+)_", filename)
    return match.group(1) if match else ""

def read_order_xlsx(uploaded_file):
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0, header=None, dtype=object)
        return df.fillna("")
    except Exception as e:
        st.error(f"Errore lettura file: {e}")
        return None

def find_col(row, label):
    for i, v in enumerate(row):
        if isinstance(v, str) and v.strip() == label:
            return i
    return None

# --------------------------------
# Parser nuovo template Nike
# --------------------------------
def process_order_details(df, order_id, view_option):
    rows = []

    current_model = None
    model_name = None
    color_desc = None
    product_type = None

    col_size = col_upc = col_confirmed = col_shipped = None
    in_table = False

    for _, r in df.iterrows():
        row = r.tolist()

        # Nuovo articolo
        idx = find_col(row, "Modello/Colore:")
        if idx is not None:
            current_model = str(row[idx + 1]).strip()
            model_name = color_desc = product_type = None
            in_table = False
            continue

        # Metadati
        idx = find_col(row, "Nome modello:")
        if idx is not None:
            model_name = str(row[idx + 1]).strip()
            continue

        idx = find_col(row, "Descrizione colore:")
        if idx is not None:
            color_desc = str(row[idx + 1]).strip()
            continue

        idx = find_col(row, "Tipo di prodotto:")
        if idx is not None:
            product_type = str(row[idx + 1]).strip()
            continue

        # Header tabella taglie
        if isinstance(row[0], str) and row[0].strip() == "Misura":
            in_table = True
            col_size = 0
            col_upc = 1
            col_confirmed = col_shipped = None

            for i, v in enumerate(row):
                if isinstance(v, str):
                    if v.strip() == "Aperti:":
                        col_confirmed = i
                    if v.strip() == "Spediti:":
                        col_shipped = i
            continue

        if not in_table or not current_model:
            continue

        # Fine tabella
        if isinstance(row[0], str) and row[0].startswith("Qtà totale"):
            in_table = False
            continue

        # Riga taglia
        size = str(row[col_size]).strip()
        if not size:
            continue

        def to_int(x):
            try:
                return int(float(str(x).replace(",", ".")))
            except:
                return 0

        rows.append([
            current_model,
            size,
            to_int(row[col_confirmed]),
            to_int(row[col_shipped]),
            model_name,
            color_desc,
            str(row[col_upc]).strip(),
            product_type,
            order_id
        ])

    df_final = pd.DataFrame(
        rows,
        columns=[
            "Modello/Colore", "Misura", "Confermati", "Spediti",
            "Nome del modello", "Descrizione colore",
            "Codice a Barre (UPC)", "Tipo di prodotto", "ID_ORDINE"
        ]
    )

    if df_final.empty:
        return None, df_final

    # Codice e Colore
    df_final["Codice"] = df_final["Modello/Colore"].apply(lambda x: x.split("-")[0])
    df_final["Colore"] = df_final["Modello/Colore"].apply(
        lambda x: x.split("-")[1] if "-" in x else ""
    )

    # Rimuove righe inutili
    df_final = df_final[(df_final["Confermati"] != 0) | (df_final["Spediti"] != 0)]

    base_cols = [
        "Modello/Colore", "Descrizione colore", "Codice",
        "Nome del modello", "Tipo di prodotto",
        "Colore", "Misura", "Codice a Barre (UPC)", "ID_ORDINE"
    ]

    if view_option == "CONFERMATI":
        df_final = df_final[base_cols + ["Confermati"]]
    else:
        df_final = df_final[base_cols + ["Spediti"]]

    # Export Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False)

    return output.getvalue(), df_final

# --------------------------------
# UI Streamlit
# --------------------------------
st.title("Nike order details")

uploaded_file = st.file_uploader(
    "Carica il file XLSX (Order Details)",
    type="xlsx"
)

if uploaded_file:
    filename = os.path.splitext(uploaded_file.name)[0]
    order_id = st.text_input("ID_ORDINE", extract_order_id(filename))

    view_option = st.radio(
        "Seleziona vista",
        ("CONFERMATI", "SPEDITI")
    )

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
