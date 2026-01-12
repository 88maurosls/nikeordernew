import streamlit as st
import pandas as pd
from io import BytesIO
import re
import os

# -----------------------------
# Helpers
# -----------------------------
def extract_order_id(filename: str) -> str:
    match = re.search(r"_(\d+)_", filename)
    return match.group(1) if match else ""

def read_order_xlsx(uploaded_file) -> pd.DataFrame:
    """
    Legge il nuovo template Nike (sheet 'Order Details') senza header,
    così possiamo cercare le label nelle celle.
    """
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0, header=None, dtype=object)
        df = df.fillna("")
        return df
    except Exception as e:
        st.error(f"Errore lettura XLSX: {e}")
        return None

def _find_first_col_index(row_values, target: str):
    for i, v in enumerate(row_values):
        if isinstance(v, str) and v.strip() == target:
            return i
    return None

def build_price_map(price_file) -> dict:
    """
    Prova a costruire un mapping prezzi da un file listino CSV/XLSX.

    Supporta mapping su:
    - 'Modello/Colore' (es. BV1021-109)
    - oppure 'Codice' (es. BV1021)

    Cerca automaticamente una colonna prezzo plausibile.
    """
    if price_file is None:
        return {}

    try:
        name = price_file.name.lower()
        if name.endswith(".csv"):
            price_df = pd.read_csv(price_file)
        else:
            price_df = pd.read_excel(price_file)

        price_df.columns = [str(c).strip() for c in price_df.columns]

        # colonne chiave possibili
        key_cols_priority = [
            "Modello/Colore", "Modello", "StyleColor", "style_color",
            "SKU", "sku", "Item Code", "ITEM CODE", "Codice"
        ]

        key_col = None
        for c in key_cols_priority:
            if c in price_df.columns:
                key_col = c
                break
        if key_col is None:
            # fallback: prima colonna
            key_col = price_df.columns[0]

        # colonne prezzo possibili
        price_cols_priority = [
            "Prezzo all'ingrosso", "Prezzo all’ingrosso", "Prezzo all'ingrosso (EUR)",
            "Wholesale", "wholesale", "WHL", "whl", "WHS", "prezzo", "price"
        ]
        price_col = None
        for c in price_cols_priority:
            if c in price_df.columns:
                price_col = c
                break

        if price_col is None:
            # fallback: prima colonna numerica
            numeric_candidates = []
            for c in price_df.columns:
                s = pd.to_numeric(price_df[c], errors="coerce")
                if s.notna().sum() > 0:
                    numeric_candidates.append((c, s.notna().sum()))
            numeric_candidates.sort(key=lambda x: x[1], reverse=True)
            if numeric_candidates:
                price_col = numeric_candidates[0][0]

        if price_col is None:
            st.warning("Listino caricato ma non trovo una colonna prezzo utilizzabile.")
            return {}

        # pulizia prezzo (virgole, euro, ecc)
        def to_float(x):
            if x is None:
                return None
            s = str(x).strip().replace("€", "").replace("EUR", "").strip()
            s = s.replace(",", ".")
            try:
                return float(s)
            except:
                return None

        mapping = {}
        for _, r in price_df.iterrows():
            k = str(r.get(key_col, "")).strip()
            if not k:
                continue
            p = to_float(r.get(price_col, None))
            if p is None:
                continue
            mapping[k] = p

        return mapping

    except Exception as e:
        st.warning(f"Impossibile leggere il listino prezzi: {e}")
        return {}

def process_order_details(df: pd.DataFrame, discount_percentage: float, order_id: str,
                          view_option: str, price_map: dict, default_wholesale_price: float | None):
    """
    Parser per il nuovo template.
    - 'Confermati' viene preso da colonna 'Aperti:' (nel template allegato coincide con le qty aperte)
    - 'Spediti' da colonna 'Spediti:'
    """
    new_data = []

    current_model = None
    current_model_name = None
    current_color_description = None
    current_product_type = None

    # colonne della tabella taglie, verranno scoperte quando incontriamo la riga header "Misura"
    col_size = col_upc = col_confirmed = col_shipped = None
    in_sizes_table = False

    for _, row in df.iterrows():
        row_vals = row.tolist()

        # Inizio articolo
        idx_mc = _find_first_col_index(row_vals, "Modello/Colore:")
        if idx_mc is not None:
            current_model = str(row_vals[idx_mc + 1]).strip()
            current_model_name = None
            current_color_description = None
            current_product_type = None
            in_sizes_table = False
            col_size = col_upc = col_confirmed = col_shipped = None
            continue

        # Metadati articolo
        idx_name = _find_first_col_index(row_vals, "Nome modello:")
        if idx_name is not None and current_model:
            current_model_name = str(row_vals[idx_name + 1]).strip()
            continue

        idx_color = _find_first_col_index(row_vals, "Descrizione colore:")
        if idx_color is not None and current_model:
            current_color_description = str(row_vals[idx_color + 1]).strip()
            continue

        idx_type = _find_first_col_index(row_vals, "Tipo di prodotto:")
        if idx_type is not None and current_model:
            current_product_type = str(row_vals[idx_type + 1]).strip()
            continue

        # Header tabella taglie
        if isinstance(row_vals[0], str) and row_vals[0].strip() == "Misura" and current_model:
            # Identifica indici colonne importanti
            in_sizes_table = True
            col_size = 0

            # di solito UPC è col 1
            col_upc = 1

            # Trova colonne "Aperti:" e "Spediti:"
            # (nell'allegato: Aperti: è col 5, Spediti: col 7)
            for i, v in enumerate(row_vals):
                if isinstance(v, str):
                    t = v.strip()
                    if t == "Aperti:":
                        col_confirmed = i
                    if t == "Spediti:":
                        col_shipped = i
            continue

        if not in_sizes_table or not current_model:
            continue

        # Fine tabella taglie
        if isinstance(row_vals[0], str) and row_vals[0].strip().startswith("Qtà totale"):
            in_sizes_table = False
            continue

        # Riga taglia
        size = str(row_vals[col_size]).strip() if col_size is not None else ""
        if not size or size in ["", "Misura"]:
            continue

        upc = str(row_vals[col_upc]).strip() if col_upc is not None else ""

        def safe_int(x):
            try:
                return int(float(str(x).strip().replace(",", ".")))
            except:
                return 0

        confirmed = safe_int(row_vals[col_confirmed]) if col_confirmed is not None else 0
        shipped = safe_int(row_vals[col_shipped]) if col_shipped is not None else 0

        new_data.append([
            current_model,
            size,
            confirmed,
            shipped,
            current_model_name,
            current_color_description,
            upc,
            discount_percentage,
            current_product_type,
            order_id
        ])

    final_df = pd.DataFrame(
        new_data,
        columns=[
            "Modello/Colore", "Misura", "Confermati", "Spediti",
            "Nome del modello", "Descrizione colore", "Codice a Barre (UPC)",
            "Percentuale sconto", "Tipo di prodotto", "ID_ORDINE"
        ]
    )

    if final_df.empty:
        return None, final_df

    # Codice e Colore da Modello/Colore
    def split_code_color(x):
        x = str(x).strip()
        if "-" in x:
            code, _, color = x.partition("-")
            return code, color
        return x, ""

    final_df[["Codice", "Colore"]] = final_df["Modello/Colore"].apply(lambda x: pd.Series(split_code_color(x)))

    # Prezzo all'ingrosso: da price_map (prima su Modello/Colore, poi su Codice), altrimenti default
    def resolve_wholesale(row):
        mc = str(row["Modello/Colore"]).strip()
        code = str(row["Codice"]).strip()

        if mc in price_map:
            return float(price_map[mc])
        if code in price_map:
            return float(price_map[code])
        if default_wholesale_price is not None:
            return float(default_wholesale_price)
        return None

    final_df["Prezzo all'ingrosso"] = final_df.apply(resolve_wholesale, axis=1)

    # Calcoli prezzo finale e totali solo se prezzo presente
    final_df["Prezzo finale"] = pd.to_numeric(final_df["Prezzo all'ingrosso"], errors="coerce") * (
        1 - (pd.to_numeric(final_df["Percentuale sconto"], errors="coerce").fillna(0) / 100.0)
    )

    final_df["TOT CONFERMATI"] = final_df["Prezzo finale"] * pd.to_numeric(final_df["Confermati"], errors="coerce").fillna(0)
    final_df["TOT SPEDITI"] = final_df["Prezzo finale"] * pd.to_numeric(final_df["Spediti"], errors="coerce").fillna(0)

    # Rimozione righe con confermati e spediti entrambi 0
    final_df = final_df[(final_df["Confermati"] != 0) | (final_df["Spediti"] != 0)]

    # Selezione colonne per vista
    base_cols = [
        "Modello/Colore", "Descrizione colore", "Codice", "Nome del modello",
        "Tipo di prodotto", "Colore", "Misura", "Codice a Barre (UPC)", "ID_ORDINE"
    ]

    if view_option == "CONFERMATI":
        final_df = final_df[base_cols + [
            "Confermati", "Prezzo all'ingrosso", "Percentuale sconto", "Prezzo finale", "TOT CONFERMATI"
        ]]
    else:
        final_df = final_df[base_cols + [
            "Spediti", "Prezzo all'ingrosso", "Percentuale sconto", "Prezzo finale", "TOT SPEDITI"
        ]]

    # Export Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False)

    return output.getvalue(), final_df


# -----------------------------
# UI Streamlit
# -----------------------------
st.title("Nike order details")

uploaded_file = st.file_uploader("Carica il nuovo file XLSX (Order Details)", type="xlsx")

if uploaded_file is not None:
    original_filename = os.path.splitext(uploaded_file.name)[0]
    extracted_order_id = extract_order_id(original_filename)
    order_id = st.text_input("ID_ORDINE", value=extracted_order_id)

    view_option = st.radio("Seleziona l'opzione di visualizzazione:", ("CONFERMATI", "SPEDITI"))

    discount_percentage = st.number_input(
        "Inserisci la percentuale di sconto sul prezzo whl",
        min_value=0.0, max_value=100.0, step=0.1, value=0.0
    )

    st.subheader("Prezzo all'ingrosso")
    default_wholesale_price = st.number_input(
        "Prezzo all'ingrosso default (EUR) da applicare se non carichi un listino",
        min_value=0.0, step=0.01, value=0.0
    )
    use_default_price = st.checkbox("Usa questo prezzo default", value=False)
    default_price_value = float(default_wholesale_price) if use_default_price else None

    price_file = st.file_uploader(
        "Opzionale: carica un listino prezzi (CSV o XLSX) con colonne tipo 'Modello/Colore' o 'Codice' + prezzo",
        type=["csv", "xlsx"]
    )

    df = read_order_xlsx(uploaded_file)
    if df is not None:
        price_map = build_price_map(price_file)

        if st.button("Elabora"):
            processed_file, final_df = process_order_details(
                df=df,
                discount_percentage=discount_percentage,
                order_id=order_id,
                view_option=view_option,
                price_map=price_map,
                default_wholesale_price=default_price_value
            )

            if processed_file is None:
                st.warning("Nessun dato estratto. Controlla che il file sia quello corretto (Order Details).")
            else:
                st.write("Anteprima del file elaborato:")
                st.dataframe(final_df, use_container_width=True)

                processed_filename = f"{original_filename}_processed.xlsx"
                st.download_button(
                    label="Scarica il file elaborato",
                    data=processed_file,
                    file_name=processed_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

