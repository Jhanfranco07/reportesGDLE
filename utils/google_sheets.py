import io
import re
import unicodedata

import pandas as pd
import streamlit as st


def normalize_column_name(value):
    """Normalize column labels from Google Sheets without depending on exact accents."""
    text = str(value or "").strip().upper()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_text(value):
    text = str(value or "").strip().upper()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = re.sub(r"\s+", " ", text)
    return text


def parse_money_series(series):
    values = (
        series.astype(str)
        .str.strip()
        .str.replace("S/", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    numeric = pd.to_numeric(values, errors="coerce").fillna(0.0).astype(float)
    cents_mask = values.str.fullmatch(r"\d+").fillna(False) & (numeric.abs() >= 1000)
    numeric.loc[cents_mask] = numeric.loc[cents_mask] / 100
    return numeric


def get_secret_value(key):
    """Read Streamlit secrets while tolerating UTF-8 BOM in local TOML files."""
    if key in st.secrets:
        return st.secrets[key]
    bom_key = f"\ufeff{key}"
    if bom_key in st.secrets:
        return st.secrets[bom_key]
    raise KeyError(key)


def extract_google_sheet_id(value):
    text = str(value or "").strip()
    if not text:
        return ""

    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", text)
    if match:
        return match.group(1)

    match = re.search(r"[?&]id=([a-zA-Z0-9-_]+)", text)
    if match:
        return match.group(1)

    return text


def get_gspread_client(scopes):
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError as exc:
        raise RuntimeError(
            "Faltan dependencias para usar Google Sheets. Instala gspread y google-auth."
        ) from exc

    credentials = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=scopes,
    )
    return gspread.authorize(credentials)


def open_google_worksheet(sheet_id, tab_name, scopes=None):
    required_scopes = scopes or [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    client = get_gspread_client(required_scopes)
    return client.open_by_key(extract_google_sheet_id(sheet_id)).worksheet(tab_name)


def read_google_worksheet_with_rows(sheet_id, tab_name):
    worksheet = open_google_worksheet(sheet_id, tab_name)
    values = worksheet.get_all_values()
    if not values:
        return worksheet, pd.DataFrame(), {}

    headers = [normalize_column_name(header) for header in values[0]]
    column_numbers = {}
    for index, header in enumerate(headers, start=1):
        if header and header not in column_numbers:
            column_numbers[header] = index

    rows = []
    for sheet_row, row_values in enumerate(values[1:], start=2):
        padded = row_values + [""] * (len(headers) - len(row_values))
        record = dict(zip(headers, padded[: len(headers)]))
        record["__SHEET_ROW"] = sheet_row
        rows.append(record)

    return worksheet, pd.DataFrame(rows), column_numbers


@st.cache_data(ttl=600, show_spinner=False)
def load_resoluciones_sheet(tab_name=None):
    """Read the private Google Sheet configured in Streamlit secrets."""
    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseDownload
        from google.oauth2.service_account import Credentials
    except ImportError as exc:
        raise RuntimeError(
            "Faltan dependencias para leer Google Drive/Sheets. Instala gspread, google-auth y google-api-python-client."
        ) from exc

    required_scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    credentials = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=required_scopes,
    )

    client = get_gspread_client(required_scopes)
    sheet_id = get_secret_value("GOOGLE_SHEET_ID")
    tab_name = tab_name or get_secret_value("GOOGLE_SHEET_TAB")
    try:
        worksheet = client.open_by_key(sheet_id).worksheet(tab_name)
        records = worksheet.get_all_records()
        df = pd.DataFrame(records)
    except Exception:
        drive = build("drive", "v3", credentials=credentials)
        metadata = drive.files().get(fileId=sheet_id, fields="mimeType,name").execute()
        mime_type = metadata.get("mimeType", "")
        buffer = io.BytesIO()

        if mime_type == "application/vnd.google-apps.spreadsheet":
            request = drive.files().export_media(
                fileId=sheet_id,
                mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            request = drive.files().get_media(fileId=sheet_id)

        downloader = MediaIoBaseDownload(buffer, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()

        buffer.seek(0)
        df = pd.read_excel(buffer, sheet_name=tab_name)

    df.columns = [normalize_column_name(col) for col in df.columns]
    df = df.loc[:, [bool(col) for col in df.columns]]
    df = df.loc[:, ~df.columns.duplicated()]
    return df


def get_resoluciones_sheet_or_none(tab_name=None, show_warning=True):
    try:
        df = load_resoluciones_sheet(tab_name)
        return df if df is not None and not df.empty else None
    except Exception as exc:
        if tab_name is None:
            for fallback_tab in ["RESOLUCIONES 2026"]:
                try:
                    df = load_resoluciones_sheet(fallback_tab)
                    return df if df is not None and not df.empty else None
                except Exception:
                    pass
        if show_warning:
            tab_text = f" ({tab_name})" if tab_name else ""
            st.warning(f"No se pudo leer Drive{tab_text}. Se usaran datos locales. Detalle: {exc}")
        return None
