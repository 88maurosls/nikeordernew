import streamlit as st
import pandas as pd
from io import BytesIO
import re
import os
from numbers import Real


def normalize(value):
    return re.sub(r'\s+', ' ', str(value).strip()).rstrip(':').strip().casefold()


def find_col(row, *labels):
    wanted = {normalize(label) for label in labels}
    return next((i for i, v in enumerate(row) if normalize(v) in wanted), None)


def cell(row, index):
    return row[index] if index is not None and 0 <= index < len(row) else ''


def text_value(value):
    if isinstance(value, Real) and float(value).is_integer():
        return str(int(value))
    return str(value).strip()


def extract_order_id(filename):
    match = re.search(r'_(\d+)_', filename)
    return match.group(1) if match else ''


def extract_order_id_from_df(df):
    for row in df.itertuples(index=False, name=None):
        idx = find_col(row, 'Numero ordine:')
        if idx is not None:
            value = text_value(cell(row, idx + 1))
            if value:
                return value
    return ''


def read_order_xlsx(uploaded_file):
    try:
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, sheet_name=0, header=None, dtype=object).fillna('')
    except Exception as e:
        st.error(f'Errore lettura file: {e}')
        return None


def to_money(value):
    if isinstance(value, Real):
        return float(value)
    s = re.sub(r'\s+', '', str(value).replace('€', ''))
    if ',' in s:
        if '.' in s and s.rfind('.') > s.rfind(','):
            s = s.replace(',', '')
        else:
            s = s.replace('.', '').replace(',', '.')
    elif re.fullmatch(r'[-+]?\d{1,3}(\.\d{3})+', s):
        s = s.replace('.', '')
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def to_int(value):
    try:
        return int(to_money(value))
    except (ValueError, OverflowError):
        return 0


def process_order_details(df, order_id, view_option):
    rows = []
    current_model = ''
    meta = {}
    columns = {}
    in_table = False
    qty_names = ('Richiesti', 'Aperti', 'Spediti')

    for row in df.itertuples(index=False, name=None):
        idx = find_col(row, 'Modello/Colore:')
        if idx is not None:
            current_model = text_value(cell(row, idx + 1))
            meta = {'Nome del modello': '', 'Descrizione colore': '',
                    'Tipo di prodotto': '', 'WHS': 0.0, 'RTL': 0.0}
            in_table = False
            columns = {}

        if not current_model:
            continue

        # Legge tutti i metadati, anche se presenti sulla stessa riga.
        for key, labels in {
            'Nome del modello': ('Nome modello:', 'Nome prodotto:'),
            'Descrizione colore': ('Descrizione colore:',),
            'Tipo di prodotto': ('Tipo di prodotto:',),
            'WHS': ("All'ingrosso:",),
            'RTL': ('Retail consigliato:',),
        }.items():
            pos = find_col(row, *labels)
            if pos is not None:
                value = cell(row, pos + 1)
                meta[key] = to_money(value) if key in ('WHS', 'RTL') else text_value(value)

        if idx is not None:
            continue

        size_col = find_col(row, 'Misura', 'Taglia/misura', 'Taglia')
        if size_col is not None:
            upc_col = find_col(row, 'UPC', 'Codice a Barre (UPC)')
            columns = {'Misura': size_col, 'UPC': upc_col if upc_col is not None else size_col + 1}
            columns.update({name: find_col(row, name) for name in qty_names})
            in_table = all(columns[name] is not None for name in qty_names)
            continue

        if not in_table:
            continue

        size = text_value(cell(row, columns['Misura']))
        if normalize(size).startswith(('qtà totale', 'quantità totale', 'totale quantità')):
            in_table = False
            continue
        upc = text_value(cell(row, columns['UPC']))
        if not size or not upc:
            continue

        code, _, color = current_model.partition('-')
        record = {
            'Modello/Colore': current_model,
            'Codice': code, 'Colore': color, 'Misura': size,
            'Codice a Barre (UPC)': upc, 'ID_ORDINE': str(order_id),
            **meta,
        }
        record.update({name: to_int(cell(row, columns[name])) for name in qty_names})
        rows.append(record)

    base_cols = [
        'Modello/Colore', 'Descrizione colore', 'Codice',
        'Nome del modello', 'Tipo di prodotto', 'Colore', 'Misura',
        'Codice a Barre (UPC)', 'ID_ORDINE', view_option, 'WHS', 'RTL'
    ]
    if not rows:
        return None, pd.DataFrame(columns=base_cols)

    result = pd.DataFrame(rows)
    result = result.loc[result[view_option] > 0, base_cols].reset_index(drop=True)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result.to_excel(writer, index=False)
        ws = writer.sheets['Sheet1']
        for name in ('Codice', 'Colore', 'Misura', 'Codice a Barre (UPC)', 'ID_ORDINE'):
            for c in list(ws.columns)[result.columns.get_loc(name)][1:]:
                c.number_format = '@'
        for name in ('WHS', 'RTL'):
            for c in list(ws.columns)[result.columns.get_loc(name)][1:]:
                c.number_format = '0.00'
    return output.getvalue(), result


def main():
    st.title('Nike order details v3')
    uploaded_file = st.file_uploader('Carica il file XLSX (Order Details)', type='xlsx')
    if uploaded_file is None:
        return
    filename = os.path.splitext(uploaded_file.name)[0]
    df = read_order_xlsx(uploaded_file)
    if df is None:
        return
    auto_id = extract_order_id_from_df(df) or extract_order_id(filename)
    order_id = st.text_input('ID_ORDINE', value=auto_id, key=f'order_id_{uploaded_file.file_id}')
    view_option = st.radio('Seleziona vista', ('RICHIESTI', 'APERTI', 'SPEDITI'), index=1).capitalize()
    if st.button('Elabora'):
        file_out, preview = process_order_details(df, order_id, view_option)
        st.session_state['nike_result'] = (uploaded_file.file_id, order_id, view_option, file_out, preview)
    saved = st.session_state.get('nike_result')
    if saved and saved[:3] == (uploaded_file.file_id, order_id, view_option):
        _, _, _, file_out, preview = saved
        if file_out is None:
            st.error('Nessun articolo riconosciuto. Verifica il formato del file Nike.')
        elif preview.empty:
            st.warning(f'Nessuna quantità presente nella vista {view_option.upper()}.')
        else:
            st.success(f"{preview['Modello/Colore'].nunique()} articoli | {len(preview)} righe | {int(preview[view_option].sum())} pezzi")
            st.dataframe(preview, use_container_width=True)
            st.download_button('Scarica Excel', data=file_out, file_name=f'{filename}_processed.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == '__main__':
    main()
