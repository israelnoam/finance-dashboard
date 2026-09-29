"""
Google Sheets integration module using gspread.
Connects to Google Sheet named 'Financial_DB' and tab 'MasterData'.
Columns: Date, Business Name, Category, Type, Amount.
"""

import os
import gspread
import pandas as pd
from datetime import datetime
from auth import get_google_credentials

SHEET_NAME = "Financial_DB"
WORKSHEET_NAME = "MasterData"
REQUIRED_COLUMNS = ["Date", "Business Name", "Category", "Type", "Amount", "Notes"]
LOCAL_DATA_FILE = "MasterData.csv"
LOCAL_FLAG_FILE = ".use_local_storage"


def is_local_storage_enabled():
    """Returns True if local storage mode is explicitly active."""
    return os.path.exists(LOCAL_FLAG_FILE)


def enable_local_storage():
    """Enables local storage mode and initializes MasterData.csv if missing."""
    with open(LOCAL_FLAG_FILE, "w", encoding="utf-8") as f:
        f.write("local")
    if not os.path.exists(LOCAL_DATA_FILE):
        df = pd.DataFrame(columns=REQUIRED_COLUMNS)
        df.to_csv(LOCAL_DATA_FILE, index=False, encoding="utf-8-sig")


def disable_local_storage():
    """Disables local storage mode to prefer Google Sheets."""
    if os.path.exists(LOCAL_FLAG_FILE):
        try:
            os.remove(LOCAL_FLAG_FILE)
        except Exception:
            pass


def get_gspread_client():
    """Returns an authorized gspread client using local OAuth2 desktop credentials."""
    creds = get_google_credentials()
    return gspread.authorize(creds)


def get_master_worksheet(client=None):
    """
    Finds and returns the 'MasterData' worksheet from 'Financial_DB'.
    If the spreadsheet exists but worksheet doesn't, creates it with required headers.
    """
    if client is None:
        client = get_gspread_client()

    try:
        spreadsheet = client.open(SHEET_NAME)
    except gspread.SpreadsheetNotFound:
        # Prompt user or create it if not found
        spreadsheet = client.create(SHEET_NAME)

    try:
        worksheet = spreadsheet.worksheet(WORKSHEET_NAME)
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=WORKSHEET_NAME, rows=1000, cols=10)
        worksheet.append_row(REQUIRED_COLUMNS)
        return worksheet

    # Ensure headers are present
    existing_headers = worksheet.row_values(1)
    if not existing_headers:
        worksheet.append_row(REQUIRED_COLUMNS)
    
    return worksheet


def fetch_master_data():
    """
    Fetches all transactions from MasterData worksheet or local MasterData.csv.
    Returns:
        pd.DataFrame with columns: Date, Business Name, Category, Type, Amount, Notes
    """
    if is_local_storage_enabled():
        if not os.path.exists(LOCAL_DATA_FILE):
            df = pd.DataFrame(columns=REQUIRED_COLUMNS)
            df.to_csv(LOCAL_DATA_FILE, index=False, encoding="utf-8-sig")
            return df
        try:
            df = pd.read_csv(LOCAL_DATA_FILE, encoding="utf-8-sig")
        except Exception:
            df = pd.read_csv(LOCAL_DATA_FILE, encoding="utf-8", errors="replace")
    else:
        worksheet = get_master_worksheet()
        records = worksheet.get_all_records()
        if not records:
            return pd.DataFrame(columns=REQUIRED_COLUMNS)
        df = pd.DataFrame(records)

    # Column alias resolution (handle spelling variations or legacy headers)
    rename_map = {
        "Buisness Name": "Business Name",
        "buisness name": "Business Name",
        "business name": "Business Name",
        "Business_Name": "Business Name"
    }
    for old_c, new_c in rename_map.items():
        if old_c in df.columns:
            if new_c not in df.columns or (df[new_c].fillna("").astype(str).str.strip() == "").all():
                df[new_c] = df[old_c]

    # Ensure all required columns are present
    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    # Clean and standardize types
    df["Date_Clean"] = pd.to_datetime(df["Date"], errors="coerce")
    
    # Normalize Amount to float
    def clean_amount(val):
        if pd.isna(val) or val == "":
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)
        val_str = str(val).replace(",", "").replace("₪", "").replace("$", "").strip()
        try:
            return float(val_str)
        except ValueError:
            return 0.0

    df["Amount"] = df["Amount"].apply(clean_amount)
    df["Type"] = df["Type"].astype(str).str.strip().str.capitalize()
    # Normalize Type to Income / Expense
    df["Type"] = df["Type"].apply(lambda t: "Income" if "inc" in t.lower() or "הכנסה" in t else "Expense")
    df["Category"] = df["Category"].fillna("Uncategorized").astype(str).str.strip()
    df["Business Name"] = df["Business Name"].fillna("").astype(str).str.strip()
    df["Notes"] = df["Notes"].fillna("").astype(str).str.strip()

    return df


def append_transactions(new_df: pd.DataFrame):
    """
    Appends a DataFrame of new transactions to the MasterData database (local or Google Sheets).
    new_df must contain: Date, Business Name, Category, Type, Amount, Notes
    """
    if new_df.empty:
        return 0

    if is_local_storage_enabled():
        current_df = fetch_master_data()
        clean_new = []
        for _, row in new_df.iterrows():
            dt = row.get("Date")
            date_str = dt.strftime("%Y-%m-%d") if isinstance(dt, (datetime, pd.Timestamp)) else str(dt)
            clean_new.append({
                "Date": date_str,
                "Business Name": str(row.get("Business Name", "")).strip(),
                "Category": str(row.get("Category", "Uncategorized")).strip(),
                "Type": str(row.get("Type", "Expense")).strip().capitalize(),
                "Amount": round(float(row.get("Amount", 0.0)), 2),
                "Notes": str(row.get("Notes", "")).strip() if pd.notna(row.get("Notes")) else ""
            })
        combined_df = pd.concat([current_df[REQUIRED_COLUMNS], pd.DataFrame(clean_new)], ignore_index=True)
        combined_df.to_csv(LOCAL_DATA_FILE, index=False, encoding="utf-8-sig")
        return len(clean_new)
    else:
        worksheet = get_master_worksheet()
        rows_to_append = []
        for _, row in new_df.iterrows():
            dt = row.get("Date")
            date_str = dt.strftime("%Y-%m-%d") if isinstance(dt, (datetime, pd.Timestamp)) else str(dt)
            b_name = str(row.get("Business Name", "")).strip()
            cat = str(row.get("Category", "Uncategorized")).strip()
            t_type = str(row.get("Type", "Expense")).strip().capitalize()
            amount = float(row.get("Amount", 0.0))
            notes = str(row.get("Notes", "")).strip() if pd.notna(row.get("Notes")) else ""
            rows_to_append.append([date_str, b_name, cat, t_type, round(amount, 2), notes])

        worksheet.append_rows(rows_to_append, value_input_option="USER_ENTERED")
        return len(rows_to_append)


def add_single_transaction(date_val, business_name: str, category: str, trans_type: str, amount: float, notes: str = ""):
    """Appends a single manual transaction to MasterData."""
    single_df = pd.DataFrame([{
        "Date": date_val,
        "Business Name": business_name,
        "Category": category,
        "Type": trans_type,
        "Amount": amount,
        "Notes": notes
    }])
    return append_transactions(single_df) > 0


def overwrite_master_data(df: pd.DataFrame):
    """
    Overwrites the MasterData database (local or Google Sheets) with the provided full DataFrame.
    Preserves headers: Date, Business Name, Category, Type, Amount, Notes.
    """
    if is_local_storage_enabled():
        clean_df = df.copy()
        for col in REQUIRED_COLUMNS:
            if col not in clean_df.columns:
                clean_df[col] = ""
        clean_df[REQUIRED_COLUMNS].to_csv(LOCAL_DATA_FILE, index=False, encoding="utf-8-sig")
        return len(clean_df)
    else:
        worksheet = get_master_worksheet()
        worksheet.clear()
        worksheet.append_row(REQUIRED_COLUMNS)

        if df.empty:
            return 0

        rows_to_append = []
        for _, row in df.iterrows():
            dt = row.get("Date")
            if isinstance(dt, (datetime, pd.Timestamp)):
                date_str = dt.strftime("%Y-%m-%d")
            else:
                date_str = str(dt) if pd.notna(dt) else ""

            b_name = str(row.get("Business Name", "")).strip()
            cat = str(row.get("Category", "Uncategorized")).strip()
            t_type = str(row.get("Type", "Expense")).strip().capitalize()
            try:
                amount = float(row.get("Amount", 0.0))
            except (ValueError, TypeError):
                amount = 0.0
            notes = str(row.get("Notes", "")).strip() if pd.notna(row.get("Notes")) else ""

            rows_to_append.append([date_str, b_name, cat, t_type, round(amount, 2), notes])

        if rows_to_append:
            worksheet.append_rows(rows_to_append, value_input_option="USER_ENTERED")
        return len(rows_to_append)
