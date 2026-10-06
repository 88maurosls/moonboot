import io
import os
import re
from collections import OrderedDict

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill


st.set_page_config(page_title="Moon Boot PDF to Excel", layout="wide")
st.title("Moon Boot PDF to Excel")
st.write("v1.0 by MM")


STATIC_COLS = ["CODICE", "COLORE", "DESCRIZIONE", "PREZZO WHS", "PREZZO RTL"]


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def normalize_size(value: str) -> str:
    value = normalize_text(value).upper()
    value = value.replace("–", "-").replace("—", "-")
    value = re.sub(r"\s*-\s*", "-", value)
    return value


def parse_euro_number(value: str):
    """
    Converte sia formati tipo 88.00 sia formati europei tipo 1.760,00.
    """
    value = normalize_text(value)
    if not value:
        return None

    try:
        if "," in value:
            value = value.replace(".", "").replace(",", ".")
        return float(value)
    except ValueError:
        return None


def group_words_into_rows(words, y_tol=1.6):
    """
    Raggruppa le parole di pdfplumber in righe in base alla coordinata verticale.
    """
    rows = []

    for word in sorted(words, key=lambda x: (x["top"], x["x0"])):
        placed = False
        for row in rows:
            if abs(row["top"] - word["top"]) <= y_tol:
                row["words"].append(word)
                placed = True
                break

        if not placed:
            rows.append({"top": word["top"], "words": [word]})

    for row in rows:
        row["words"] = sorted(row["words"], key=lambda x: x["x0"])
        row["text"] = normalize_text(" ".join(w["text"] for w in row["words"]))

    return rows


def extract_order_number(file_obj):
    """
    Estrae il numero ordine Moon Boot / Tecnica Group.
    Esempio: Proposta d'ordine/Order proposal n. 164059
    """
    file_obj.seek(0)

    patterns = [
        r"Proposta\s+d['’]ordine\s*/?\s*Order\s+proposal\s+n\.?\s*([A-Z0-9-]+)",
        r"Order\s+([A-Z0-9-]+)\s*-\s*page",
        r"Order\s+proposal\s+(?:no\.?|n\.?)\s*:?\s*([A-Z0-9-]+)",
    ]

    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            text = normalize_text(page.extract_text() or "")
            for pattern in patterns:
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    file_obj.seek(0)
                    return match.group(1).strip()

    file_obj.seek(0)
    return ""


def extract_declared_totals(file_obj):
    """
    Legge, quando presenti, i totali dichiarati nella prima pagina.
    Restituisce (qty, value).
    """
    file_obj.seek(0)
    declared_qty = None
    declared_value = None

    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages[:2]:
            text = normalize_text(page.extract_text() or "")
            match = re.search(
                r"Totals\s+Qty\s*:\s*(\d+)\s+Value\s*\(EURO\)\s*:\s*([\d.,]+)",
                text,
                re.IGNORECASE,
            )
            if match:
                declared_qty = int(match.group(1))
                declared_value = parse_euro_number(match.group(2))
                break

    file_obj.seek(0)
    return declared_qty, declared_value


def build_output_filename(uploaded_files, order_number):
    if len(uploaded_files) == 1:
        original_name = uploaded_files[0].name
        base_name = os.path.splitext(original_name)[0]
        if order_number:
            return f"{base_name}_{order_number}.xlsx"
        return f"{base_name}.xlsx"

    if order_number:
        return f"moon_boot_export_{order_number}.xlsx"

    return "moon_boot_export.xlsx"


def parse_header_words(header_words):
    """
    Esempio riga:
    cod: 80D1401680 B003 - art: ICON GLANCE - var: ICON GLANCE PLATINUM
    """
    tokens = [normalize_text(w["text"]) for w in header_words]
    lower = [t.lower() for t in tokens]

    codice = ""
    colore = ""
    articolo = ""
    variante = ""

    if "cod:" in lower:
        idx = lower.index("cod:")
        if idx + 1 < len(tokens):
            codice = tokens[idx + 1]
        if idx + 2 < len(tokens):
            colore = tokens[idx + 2]

    if "art:" in lower:
        art_idx = lower.index("art:")
        var_idx = lower.index("var:") if "var:" in lower else None

        art_end = var_idx if var_idx is not None else len(tokens)
        art_tokens = tokens[art_idx + 1:art_end]
        while art_tokens and art_tokens[-1] == "-":
            art_tokens.pop()
        articolo = normalize_text(" ".join(art_tokens))

        if var_idx is not None:
            variante = normalize_text(" ".join(tokens[var_idx + 1:]))

    descrizione = variante or articolo
    return codice, colore, descrizione


def extract_retail_price(block_words, header_top):
    """
    Trova il prezzo retail dalla riga subito sotto al codice prodotto.
    """
    rows = group_words_into_rows(block_words)

    for row in rows:
        text = row["text"]
        if "retail" in text.lower() and "price" in text.lower():
            match = re.search(r"retail\s+price\s*:\s*([\d.,]+)\s*EUR", text, re.IGNORECASE)
            if match:
                return parse_euro_number(match.group(1))

    # fallback geometrico
    for word in block_words:
        if abs(word["top"] - (header_top + 6)) <= 2.5 and re.fullmatch(r"\d+[.,]\d{2}", word["text"]):
            return parse_euro_number(word["text"])

    return None


def extract_size_columns(block_words, header_top):
    """
    Ricostruisce automaticamente le colonne taglia dalla testata della tabella.

    Nel PDF Moon Boot alcune taglie sono su una riga (35, 36, 37...), mentre
    le taglie aggregate sono spezzate su due righe (es. 35- / 36 -> 35-36).

    Restituisce una lista ordinata di tuple: (taglia, x_center).
    """
    candidates = []

    for word in block_words:
        # la testata taglie si trova tra la riga retail e la riga quantità
        if not (header_top + 10 <= word["top"] <= header_top + 24):
            continue

        center = (word["x0"] + word["x1"]) / 2
        if not (95 <= center < 450):
            continue

        txt = normalize_size(word["text"])
        if not re.fullmatch(r"[A-Z0-9]+-?|[A-Z0-9]+/[A-Z0-9]+", txt):
            continue

        candidates.append({"text": txt, "top": word["top"], "center": center})

    if not candidates:
        return []

    # Raggruppa per colonna X. Le taglie aggregate sono due parole quasi allineate.
    columns = []
    for item in sorted(candidates, key=lambda x: x["center"]):
        placed = False
        for column in columns:
            if abs(column["center"] - item["center"]) <= 3.5:
                column["items"].append(item)
                column["center"] = sum(x["center"] for x in column["items"]) / len(column["items"])
                placed = True
                break

        if not placed:
            columns.append({"center": item["center"], "items": [item]})

    result = []
    for column in sorted(columns, key=lambda x: x["center"]):
        parts = [x["text"] for x in sorted(column["items"], key=lambda x: x["top"])]
        size = normalize_size("".join(parts))

        # Esclude eventuali artefatti non-taglia.
        if not re.fullmatch(r"[A-Z0-9]+(?:-[A-Z0-9]+)?", size):
            continue

        result.append((size, column["center"]))

    return result


def find_quantity_row(block_words, header_top):
    """
    Individua la riga quantità cercando contemporaneamente:
    - il totale paia nella zona x ~ 460
    - il prezzo unitario nella zona x ~ 490
    """
    rows = group_words_into_rows(block_words)

    for row in rows:
        if row["top"] <= header_top + 10:
            continue

        words = row["words"]
        has_total_qty = any(
            450 <= w["x0"] <= 475 and re.fullmatch(r"\d+", normalize_text(w["text"]))
            for w in words
        )
        has_unit_price = any(
            480 <= w["x0"] <= 510 and re.fullmatch(r"\d+[.,]\d{2}", normalize_text(w["text"]))
            for w in words
        )

        if has_total_qty and has_unit_price:
            return words

    return None


def parse_product_block(block_words, header_top):
    header_words = sorted(
        [w for w in block_words if abs(w["top"] - header_top) <= 1.8],
        key=lambda x: x["x0"],
    )

    codice, colore, descrizione = parse_header_words(header_words)
    if not codice:
        return None, []

    prezzo_rtl = extract_retail_price(block_words, header_top)
    size_columns = extract_size_columns(block_words, header_top)
    qty_row = find_quantity_row(block_words, header_top)

    if not qty_row:
        return None, size_columns

    record = {
        "CODICE": codice,
        "COLORE": colore,
        "DESCRIZIONE": descrizione,
        "PREZZO WHS": None,
        "PREZZO RTL": prezzo_rtl,
    }

    # Quantità: abbina ogni numero alla colonna taglia più vicina.
    for word in qty_row:
        txt = normalize_text(word["text"])
        center = (word["x0"] + word["x1"]) / 2

        if 95 <= center < 450 and re.fullmatch(r"\d+", txt) and size_columns:
            nearest_size, nearest_x = min(size_columns, key=lambda x: abs(x[1] - center))
            if abs(nearest_x - center) <= 10:
                record[nearest_size] = record.get(nearest_size, 0) + int(txt)

        if 480 <= word["x0"] <= 510 and re.fullmatch(r"\d+[.,]\d{2}", txt):
            record["PREZZO WHS"] = parse_euro_number(txt)

    return record, size_columns


def parse_pdf(file_obj):
    records = []
    global_size_order = []

    file_obj.seek(0)
    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            words = page.extract_words(
                x_tolerance=2,
                y_tolerance=2,
                keep_blank_chars=False,
                use_text_flow=False,
            )

            # Un prodotto inizia da una parola "cod:".
            product_headers = sorted(
                [w for w in words if normalize_text(w["text"]).lower() == "cod:"],
                key=lambda x: x["top"],
            )

            for idx, header in enumerate(product_headers):
                header_top = header["top"]
                next_top = (
                    product_headers[idx + 1]["top"]
                    if idx + 1 < len(product_headers)
                    else page.height
                )

                block_words = [
                    w for w in words
                    if header_top - 1 <= w["top"] < next_top - 1
                ]

                record, size_columns = parse_product_block(block_words, header_top)

                for size, _ in size_columns:
                    if size not in global_size_order:
                        global_size_order.append(size)

                if record:
                    records.append(record)

    file_obj.seek(0)
    return records, global_size_order


def build_dataframe(all_records, size_order):
    if not all_records:
        return pd.DataFrame()

    # Mantiene l'ordine delle taglie così come appare nel PDF.
    used_sizes = []
    for size in size_order:
        if any(record.get(size, 0) not in (0, "", None) for record in all_records):
            used_sizes.append(size)

    # Aggiunge eventuali taglie trovate nei record ma non nella testata globale.
    for record in all_records:
        for key in record.keys():
            if key not in STATIC_COLS and key not in used_sizes:
                used_sizes.append(key)

    rows = []
    for record in all_records:
        row = {col: record.get(col, "") for col in STATIC_COLS}
        for size in used_sizes:
            row[size] = record.get(size, "")
        rows.append(row)

    return pd.DataFrame(rows, columns=STATIC_COLS + used_sizes)


def calculate_total_qty(df: pd.DataFrame) -> int:
    total_qty = 0
    for col in df.columns:
        if col not in STATIC_COLS:
            total_qty += pd.to_numeric(df[col], errors="coerce").fillna(0).sum()
    return int(total_qty)


def calculate_total_value(df: pd.DataFrame) -> float:
    if df.empty or "PREZZO WHS" not in df.columns:
        return 0.0

    total_value = 0.0
    size_cols = [c for c in df.columns if c not in STATIC_COLS]

    for _, row in df.iterrows():
        qty = 0
        for col in size_cols:
            value = pd.to_numeric(pd.Series([row[col]]), errors="coerce").fillna(0).iloc[0]
            qty += value

        whs = pd.to_numeric(pd.Series([row["PREZZO WHS"]]), errors="coerce").fillna(0).iloc[0]
        total_value += qty * whs

    return round(float(total_value), 2)


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="ORDINE")
        ws = writer.book["ORDINE"]

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        header_fill = PatternFill("solid", fgColor="1F2937")
        header_font = Font(color="FFFFFF", bold=True)

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Prezzi come numeri con 2 decimali.
        for col_name in ["PREZZO WHS", "PREZZO RTL"]:
            if col_name in df.columns:
                col_idx = df.columns.get_loc(col_name) + 1
                for row_idx in range(2, ws.max_row + 1):
                    ws.cell(row=row_idx, column=col_idx).number_format = "0.00"

        # Quantità come interi.
        for col_name in df.columns:
            if col_name not in STATIC_COLS:
                col_idx = df.columns.get_loc(col_name) + 1
                for row_idx in range(2, ws.max_row + 1):
                    ws.cell(row=row_idx, column=col_idx).number_format = "0"
                    ws.cell(row=row_idx, column=col_idx).alignment = Alignment(horizontal="center")

        # Larghezze leggibili senza esplodere il foglio.
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter

            for cell in col:
                cell_value = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(cell_value))

            width = min(max(max_len + 2, 10), 42)
            if col[0].value == "DESCRIZIONE":
                width = min(max(width, 28), 45)
            ws.column_dimensions[col_letter].width = width

    output.seek(0)
    return output.getvalue()


uploaded_files = st.file_uploader(
    "Carica PDF Moon Boot",
    type=["pdf"],
    accept_multiple_files=True,
)


if uploaded_files:
    all_records = []
    all_size_order = []
    first_order_number = ""
    declared_qty_total = 0
    declared_value_total = 0.0
    declared_qty_found = False
    declared_value_found = False

    per_file_checks = []

    with st.spinner("Sto leggendo il PDF Moon Boot e creando l'Excel..."):
        for uploaded_file in uploaded_files:
            try:
                if not first_order_number:
                    first_order_number = extract_order_number(uploaded_file)

                declared_qty, declared_value = extract_declared_totals(uploaded_file)

                uploaded_file.seek(0)
                records, size_order = parse_pdf(uploaded_file)
                all_records.extend(records)

                for size in size_order:
                    if size not in all_size_order:
                        all_size_order.append(size)

                file_df = build_dataframe(records, size_order)
                file_qty = calculate_total_qty(file_df)
                file_value = calculate_total_value(file_df)

                if declared_qty is not None:
                    declared_qty_total += declared_qty
                    declared_qty_found = True

                if declared_value is not None:
                    declared_value_total += declared_value
                    declared_value_found = True

                per_file_checks.append(
                    {
                        "file": uploaded_file.name,
                        "prodotti": len(records),
                        "qty_estratta": file_qty,
                        "qty_pdf": declared_qty,
                        "valore_estratto": file_value,
                        "valore_pdf": declared_value,
                    }
                )

            except Exception as exc:
                st.error(f"Errore su {uploaded_file.name}: {exc}")

    df = build_dataframe(all_records, all_size_order)

    if df.empty:
        st.warning("Non sono riuscito a trovare prodotti nel PDF.")
    else:
        total_qty = calculate_total_qty(df)
        total_value = calculate_total_value(df)
        output_filename = build_output_filename(uploaded_files, first_order_number)

        st.success(f"Prodotti estratti: {len(df)}")
        st.info(f"Totale quantità estratte: {total_qty}")
        st.info(f"Valore WHS estratto: € {total_value:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

        if first_order_number:
            st.info(f"Numero ordine rilevato: {first_order_number}")

        if declared_qty_found:
            if total_qty == declared_qty_total:
                st.success(f"Controllo quantità PDF: OK ({declared_qty_total})")
            else:
                st.warning(
                    f"Controllo quantità PDF: estratte {total_qty}, dichiarate {declared_qty_total}"
                )

        if declared_value_found:
            if abs(total_value - declared_value_total) < 0.01:
                st.success(
                    "Controllo valore PDF: OK (€ "
                    + f"{declared_value_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    + ")"
                )
            else:
                st.warning(
                    "Controllo valore PDF: estratto € "
                    + f"{total_value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    + ", dichiarato € "
                    + f"{declared_value_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                )

        st.dataframe(df, use_container_width=True)

        with st.expander("Controlli per singolo PDF"):
            st.dataframe(pd.DataFrame(per_file_checks), use_container_width=True)

        excel_bytes = dataframe_to_excel_bytes(df)

        st.download_button(
            label="Scarica Excel",
            data=excel_bytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

