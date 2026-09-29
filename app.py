"""
Streamlit Personal Finance Dashboard
Connected to Google Sheet 'Financial_DB' (tab 'MasterData').
Features:
  - OAuth 2.0 Desktop local browser login
  - Dynamic billing cycle (9th of previous month to 8th of current month)
  - Dark mode theme (#0d0d12) with glowing emerald and crimson accents
  - Dynamic KPI cards (Income, Expenses, Net Flow)
  - Interactive Plotly charts (Category Breakdown Pie Chart & Comparison Bar Chart)
  - Category filter & Sort by dropdown
  - Transaction Center modal (@st.dialog) with Income table, Outcome table, & Manual Entry Form
  - Data Import Module: Max Excel (.xlsx) & Isracard PDF (.pdf) parser + st.data_editor + Google Sheets sync
"""

import os
import re
import json
import socket
import asyncio
import concurrent.futures
import urllib.request
from datetime import datetime, date, timezone, timedelta
from dateutil.relativedelta import relativedelta
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from telethon import TelegramClient

import auth
import sheets
import parsers
import importlib
importlib.reload(auth)
importlib.reload(sheets)
importlib.reload(parsers)
try:
    from parsers import clean_and_shorten_business_name
except ImportError:
    from bidi.algorithm import get_display
    def clean_and_shorten_business_name(raw_name: str, apply_bidi: bool = True) -> str:
        s = str(raw_name or "").strip()
        words = [w for w in s.split() if len(w) > 1][:3]
        res = " ".join(words)
        if apply_bidi and any("\u0590" <= c <= "\u05ea" for c in res):
            return get_display(res)
        return res

from styles import get_custom_css

# ==============================================================================
# Telegram User Client Configuration & Strict Security Whitelist
# ==============================================================================
TELEGRAM_API_ID = 30881080
TELEGRAM_API_HASH = "a90f4f5422de616eb7175f120b88f8e2"
TELEGRAM_PHONE = ""  # e.g., "+972501234567" - Paste your phone number here or enter via UI
TELEGRAM_SESSION_NAME = "telegram_user_session"

ALLOWED_CHANNELS = ['behatsdaa', 'fibiinvestmentchennel']


def is_channel_allowed(channel_identifier):
    """
    Validates if a channel identifier (string username or numeric integer ID)
    is explicitly in the ALLOWED_CHANNELS whitelist.
    """
    for allowed in ALLOWED_CHANNELS:
        if channel_identifier == allowed:
            return True, allowed
        if isinstance(allowed, int):
            try:
                c_int = int(channel_identifier)
                if c_int == allowed:
                    return True, allowed
                # Handle cases where channel ID is passed with or without the standard -100 prefix
                if str(abs(allowed)).startswith("100") and int(str(abs(allowed))[3:]) == abs(c_int):
                    return True, allowed
            except (ValueError, TypeError):
                pass
        elif isinstance(channel_identifier, str) and isinstance(allowed, str):
            if channel_identifier.lower().strip().lstrip("@") == allowed.lower().strip().lstrip("@"):
                return True, allowed
    return False, None


def format_telegram_message_link(channel_identifier, message_id):
    """
    Formats the Telegram message link strictly according to channel type:
    - If channel is a numeric ID (starting with -100 or negative int):
      removes the -100 prefix and formats strictly as: https://t.me/c/<stripped_id>/<message_id>
    - For standard string usernames:
      formats strictly as: https://t.me/<username>/<message_id>
    """
    ch_str = str(channel_identifier or "").strip().lstrip("@")
    msg_id_str = str(message_id or "").strip()

    if not ch_str:
        return f"https://t.me/{msg_id_str}" if msg_id_str else "https://t.me"

    # If the channel is a numeric ID (starting with -100)
    if ch_str.startswith("-100"):
        stripped_id = ch_str[4:]  # remove '-100' prefix
        return f"https://t.me/c/{stripped_id}/{msg_id_str}" if msg_id_str else f"https://t.me/c/{stripped_id}"
    elif ch_str.startswith("-"):
        stripped_id = ch_str[1:]
        return f"https://t.me/c/{stripped_id}/{msg_id_str}" if msg_id_str else f"https://t.me/c/{stripped_id}"
    elif ch_str.isdigit() and len(ch_str) > 10 and ch_str.startswith("100"):
        stripped_id = ch_str[3:]  # remove '100' prefix
        return f"https://t.me/c/{stripped_id}/{msg_id_str}" if msg_id_str else f"https://t.me/c/{stripped_id}"
    elif ch_str.isdigit():
        return f"https://t.me/c/{ch_str}/{msg_id_str}" if msg_id_str else f"https://t.me/c/{ch_str}"
    else:
        # Standard string username
        return f"https://t.me/{ch_str}/{msg_id_str}" if msg_id_str else f"https://t.me/{ch_str}"


def run_async(coro):
    """Safely executes an async coroutine in an isolated thread with its own event loop."""
    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as pool:
        return pool.submit(lambda: asyncio.run(coro)).result()


async def fetch_telegram_channel_messages(
    channel_name,
    limit: int = 10,
    search_query: str = None,
    days_back: int = 365,
    api_id: int = TELEGRAM_API_ID,
    api_hash: str = TELEGRAM_API_HASH,
    phone: str = TELEGRAM_PHONE,
    session_name: str = TELEGRAM_SESSION_NAME
):
    """
    Read-only asynchronous function using Telethon as a User Client to fetch
    messages strictly from whitelisted channels ('behatsdaa', 'fibiinvestmentchennel', -1001252653702).
    If search_query is provided, searches origin channel directly via Telegram server search
    for all matching messages up to days_back (1 year by default).
    Enforces strict access control: instantly blocks API calls if channel is not in ALLOWED_CHANNELS.
    Contains strictly read-only retrieval logic (no message sending or joining).
    """
    # Strict Access Control: instantly block API calls if not in ALLOWED_CHANNELS
    is_allowed, resolved_channel = is_channel_allowed(channel_name)
    if not is_allowed:
        return [], f"ACCESS_DENIED: Channel '{channel_name}' is not in ALLOWED_CHANNELS whitelist. API call blocked."

    client = TelegramClient(session_name, api_id, api_hash)
    await client.connect()
    
    if not await client.is_user_authorized():
        await client.disconnect()
        return None, "UNAUTHORIZED"

    try:
        # Handles both string usernames ('behatsdaa') and integer IDs (-1001252653702)
        entity = await client.get_entity(resolved_channel)
        messages_data = []

        # Determine cutoff date if searching (default: 365 days back)
        cutoff_date = None
        if days_back:
            cutoff_date = datetime.now(timezone.utc) - timedelta(days=days_back)

        query_clean = search_query.strip() if search_query and search_query.strip() else None

        # When searching, allow fetching up to limit (e.g. 150) matching messages within 1 year range
        # If no search query, fetch the latest 'limit' messages (default 10)
        iter_kwargs = {"limit": limit}
        if query_clean:
            iter_kwargs["search"] = query_clean

        # Read-only retrieval via iter_messages (no send_message, no joining)
        async for msg in client.iter_messages(entity, **iter_kwargs):
            if msg.date and cutoff_date:
                msg_date = msg.date
                if msg_date.tzinfo is None:
                    msg_date = msg_date.replace(tzinfo=timezone.utc)
                # Since iter_messages is reverse chronological (newest first),
                # once msg.date < cutoff_date, all following messages are also older than cutoff
                if msg_date < cutoff_date:
                    break

            text_val = msg.text or msg.message or ""
            if text_val.strip():
                # Formulate safe direct link to the specific message in Telegram
                channel_username = getattr(entity, "username", None)
                if channel_username:
                    post_link = f"https://t.me/{channel_username}/{msg.id}"
                else:
                    post_link = format_telegram_message_link(resolved_channel, msg.id)

                # Check if forwarded from an origin channel
                origin_link = None
                if getattr(msg, "fwd_from", None):
                    fwd = msg.fwd_from
                    orig_id = getattr(fwd, "channel_post", None)
                    from_id = getattr(fwd, "from_id", None)
                    if orig_id and from_id and hasattr(from_id, "channel_id"):
                        origin_link = f"https://t.me/c/{from_id.channel_id}/{orig_id}"

                channel_title = getattr(entity, "title", str(resolved_channel))

                messages_data.append({
                    "id": msg.id,
                    "date": msg.date,
                    "text": text_val.strip(),
                    "views": getattr(msg, "views", 0) or 0,
                    "channel": str(resolved_channel),
                    "channel_title": channel_title,
                    "link": post_link,
                    "origin_link": origin_link
                })
        await client.disconnect()
        return messages_data, None
    except Exception as exc:
        await client.disconnect()
        return None, str(exc)


async def async_check_telegram_auth(api_id: int = TELEGRAM_API_ID, api_hash: str = TELEGRAM_API_HASH, session_name: str = TELEGRAM_SESSION_NAME):
    client = TelegramClient(session_name, api_id, api_hash)
    await client.connect()
    is_auth = await client.is_user_authorized()
    await client.disconnect()
    return is_auth


async def async_send_telegram_code(phone: str, api_id: int = TELEGRAM_API_ID, api_hash: str = TELEGRAM_API_HASH, session_name: str = TELEGRAM_SESSION_NAME):
    client = TelegramClient(session_name, api_id, api_hash)
    await client.connect()
    sent = await client.send_code_request(phone)
    await client.disconnect()
    return sent.phone_code_hash


async def async_sign_in_telegram(phone: str, code: str, phone_code_hash: str, password: str = None, api_id: int = TELEGRAM_API_ID, api_hash: str = TELEGRAM_API_HASH, session_name: str = TELEGRAM_SESSION_NAME):
    client = TelegramClient(session_name, api_id, api_hash)
    await client.connect()
    try:
        if password:
            await client.sign_in(password=password)
        else:
            await client.sign_in(phone, code, phone_code_hash=phone_code_hash)
        is_auth = await client.is_user_authorized()
        await client.disconnect()
        return is_auth, None
    except Exception as e:
        await client.disconnect()
        return False, str(e)

# Page Configuration
st.set_page_config(
    page_title="Financial Dashboard",
    page_icon="💳",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Apply Theme CSS
if "theme" not in st.session_state:
    st.session_state["theme"] = "dark"

st.markdown(get_custom_css(theme=st.session_state.get("theme", "dark")), unsafe_allow_html=True)

# --- PWA & Apple iOS Standalone Full-Screen App Meta Tags ---
st.markdown("""
    <head>
        <meta name="apple-mobile-web-app-capable" content="yes">
        <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
        <meta name="apple-mobile-web-app-title" content="Finance">
        <meta name="mobile-web-app-capable" content="yes">
        <meta name="theme-color" content="#0d0d12">
        <link rel="apple-touch-icon" href="https://img.icons8.com/fluency/192/wallet.png">
    </head>
    <script>
        // Prevent accidental pull-to-refresh reload on mobile iOS
        document.addEventListener('touchstart', function() {}, {passive: true});
    </script>
""", unsafe_allow_html=True)


# --- Helper Functions ---

def get_local_ip():
    """Resolves current local Wi-Fi / LAN IP address for mobile access."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "192.168.1.219"


def get_billing_cycle_dates(reference_date: date = None):
    """
    Computes billing cycle from the 9th of the previous month to the 8th of the current month.
    """
    if reference_date is None:
        reference_date = date.today()

    prev_month = reference_date - relativedelta(months=1)
    start_date = date(prev_month.year, prev_month.month, 9)
    end_date = date(reference_date.year, reference_date.month, 8)
    
    return start_date, end_date


def generate_cycle_options(reference_date: date = None):
    """Generates options for billing cycle dropdown centered on current active cycle."""
    if reference_date is None:
        reference_date = date.today()
        
    cycles = []
    # Generate 4 past cycles, active cycle, and 1 future cycle
    for offset in range(-4, 2):
        ref_curr = reference_date + relativedelta(months=offset)
        ref_prev = ref_curr - relativedelta(months=1)
        s = date(ref_prev.year, ref_prev.month, 9)
        e = date(ref_curr.year, ref_curr.month, 8)
        
        is_current = (offset == 0)
        tag = " [ACTIVE CYCLE]" if is_current else ""
        label = f"{s.strftime('%b %d, %Y')} – {e.strftime('%b %d, %Y')}{tag}"
        cycles.append({
            "label": label,
            "start": s,
            "end": e,
            "is_current": is_current
        })
        
    return cycles


def format_currency(amount: float) -> str:
    """Formats numeric amount as formatted currency with ₪."""
    return f"₪{amount:,.2f}"


def get_demo_data():
    """Generates realistic demo transactions for preview."""
    today = date.today()
    prev_m = today - relativedelta(months=1)
    data = [
        {"Date": date(today.year, today.month, 1).strftime("%Y-%m-%d"), "Business Name": "משכורת חודשית", "Category": "Salary", "Type": "Income", "Amount": 19500.00, "Notes": "Monthly paycheck"},
        {"Date": date(prev_m.year, prev_m.month, 10).strftime("%Y-%m-%d"), "Business Name": "שופרסל דיל", "Category": "Groceries", "Type": "Expense", "Amount": 582.40, "Notes": "Weekly groceries"},
        {"Date": date(prev_m.year, prev_m.month, 12).strftime("%Y-%m-%d"), "Business Name": "קפה ארומה", "Category": "Going Out", "Type": "Expense", "Amount": 48.00, "Notes": ""},
        {"Date": date(prev_m.year, prev_m.month, 16).strftime("%Y-%m-%d"), "Business Name": "YELLOW פז תחנת דלק", "Category": "Transportation", "Type": "Expense", "Amount": 260.00, "Notes": "Full tank"},
        {"Date": date(prev_m.year, prev_m.month, 20).strftime("%Y-%m-%d"), "Business Name": "זארה קניון TLV", "Category": "Shopping", "Type": "Expense", "Amount": 389.90, "Notes": ""},
        {"Date": date(prev_m.year, prev_m.month, 24).strftime("%Y-%m-%d"), "Business Name": "חברת החשמל לישראל", "Category": "Utilities & Bills", "Type": "Expense", "Amount": 420.30, "Notes": "Bi-monthly electric bill"},
        {"Date": date(prev_m.year, prev_m.month, 27).strftime("%Y-%m-%d"), "Business Name": "סופר פארם", "Category": "Health & Pharmacy", "Type": "Expense", "Amount": 142.50, "Notes": ""},
        {"Date": date(today.year, today.month, 2).strftime("%Y-%m-%d"), "Business Name": "Domo Espresso Bar", "Category": "Going Out", "Type": "Expense", "Amount": 65.00, "Notes": "Meeting coffee"},
        {"Date": date(today.year, today.month, 3).strftime("%Y-%m-%d"), "Business Name": "מיטב דש קרן השתלמות", "Category": "Investment Fund", "Type": "Expense", "Amount": 1500.00, "Notes": "Monthly contribution"},
        {"Date": date(today.year, today.month, 4).strftime("%Y-%m-%d"), "Business Name": "רמי לוי שיווק השקמה", "Category": "Groceries", "Type": "Expense", "Amount": 645.20, "Notes": ""},
        {"Date": date(today.year, today.month, 6).strftime("%Y-%m-%d"), "Business Name": "Netflix", "Category": "Entertainment & Subs", "Type": "Expense", "Amount": 54.90, "Notes": "Subscription"},
        {"Date": date(today.year, today.month, 7).strftime("%Y-%m-%d"), "Business Name": "פנגו כחול לבן", "Category": "Transportation", "Type": "Expense", "Amount": 32.50, "Notes": "Parking"},
    ]
    df = pd.DataFrame(data)
    df["Date_Clean"] = pd.to_datetime(df["Date"], errors="coerce")
    return df


# --- Quick 1-Tap Categorization Assistant Modal Dialog ---

@st.dialog("⚡ Quick 1-Tap Categorization Assistant", width="large")
def show_quick_categorize_dialog(uncat_df: pd.DataFrame, full_df: pd.DataFrame, is_demo_mode: bool = False):
    """
    Mobile-friendly dialog offering 1-tap categorization chips for uncategorized transactions.
    """
    st.markdown("<h4 style='color: #f59e0b; margin-top: 0;'>⚡ Quick 1-Tap Categorizer</h4>", unsafe_allow_html=True)
    st.caption(f"Resolve **{len(uncat_df)}** uncategorized transactions instantly with 1 tap:")

    quick_cats = [
        ("🛍️ Shopping & Out", "Shopping & Going Out"),
        ("🛒 Groceries", "Groceries"),
        ("🚗 Transport", "Transportation"),
        ("💡 Bills", "Utilities & Bills"),
        ("💊 Health", "Health & Pharmacy"),
        ("🍿 Entertainment", "Entertainment & Subs"),
        ("🏷️ Other", "Other")
    ]

    for idx_num, (_, row) in enumerate(uncat_df.iterrows()):
        bname = row.get("Business Name") or "Unspecified Merchant"
        amt = row.get("Amount", 0.0)
        dt = str(row.get("Date", ""))
        row_id = row.get("_row_idx")

        st.markdown(f"""
        <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.08); border-radius: 8px; padding: 0.65rem 0.85rem; margin-top: 0.5rem; margin-bottom: 0.35rem;">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <span style="font-weight:600; font-size:0.95rem; color:#f3f4f6;">{bname}</span>
                <span style="font-family:'JetBrains Mono',monospace; font-weight:700; color:#f43f5e; font-size:1.05rem;">₪{amt:,.2f}</span>
            </div>
            <div style="font-size:0.8rem; color:#94a3b8; margin-top:0.2rem;">📅 {dt}</div>
        </div>
        """, unsafe_allow_html=True)

        c_chips = st.columns(4)
        for c_i, (chip_label, target_cat) in enumerate(quick_cats):
            col_target = c_chips[c_i % 4]
            with col_target:
                if st.button(chip_label, key=f"qcat_{row_id}_{target_cat}_{idx_num}", use_container_width=True):
                    if full_df is not None and "_row_idx" in full_df.columns and row_id in full_df["_row_idx"].values:
                        full_df.loc[full_df["_row_idx"] == row_id, "Category"] = target_cat
                    
                    if is_demo_mode:
                        if "demo_df" in st.session_state:
                            st.session_state["demo_df"].loc[st.session_state["demo_df"]["_row_idx"] == row_id, "Category"] = target_cat
                        st.success(f"Categorized '{bname}' as '{target_cat}'!")
                        st.rerun()
                    else:
                        with st.spinner(f"Saving '{bname}' ➔ '{target_cat}' to Google Sheets..."):
                            try:
                                sheets.overwrite_master_data(full_df)
                                st.success(f"Saved: '{bname}' ➔ '{target_cat}'!")
                                st.cache_data.clear()
                                st.rerun()
                            except Exception as e:
                                st.error(f"Failed to save update: {e}")

        st.markdown("<hr style='margin: 0.5rem 0; border: none; border-top: 1px solid rgba(255,255,255,0.06);'>", unsafe_allow_html=True)


# --- Transaction Center Modal Dialog ---

@st.dialog("Transaction Center", width="large")
def show_transaction_center_dialog(df_current_cycle: pd.DataFrame, is_demo_mode: bool = False, full_df: pd.DataFrame = None):
    """
    Hidden modal popup containing:
      1. Income Table (Editable: edit cells, delete rows, save to MasterData)
      2. Outcome (Expense) Table (Editable: edit cells, delete rows, save to MasterData)
      3. Manual Entry Form
      4. Data Import (Max Excel & Isracard PDF)
    """
    tabs = st.tabs(["💰 Income Table", "💳 Expense Table", "✍️ Manual Entry Form", "📥 Import Statements"])

    category_options = list(parsers.CATEGORY_KEYWORDS.keys()) + ["Uncategorized", "Salary", "Investment Fund", "Other"]
    cols_to_keep = ["Date", "Business Name", "Category", "Type", "Amount", "Notes"]

    # Tab 1: Income Table
    with tabs[0]:
        st.markdown("<h4 style='color: #10b981; margin-top: 0.5rem;'>Income Transactions</h4>", unsafe_allow_html=True)
        st.caption("✏️ Click any cell to edit. Select row(s) and press **Delete** to remove. Click **Save Income Changes** when finished.")
        
        income_df = df_current_cycle[df_current_cycle["Type"] == "Income"].copy()
        for c in ["Date", "Business Name", "Category", "Amount", "Notes"]:
            if c not in income_df.columns:
                income_df[c] = ""
        
        total_inc = income_df["Amount"].sum() if not income_df.empty else 0.0
        st.caption(f"Currently showing **{len(income_df)}** income records totaling **{format_currency(total_inc)}**")

        edited_income_df = st.data_editor(
            income_df,
            column_config={
                "Date": st.column_config.TextColumn("Date", required=True),
                "Business Name": st.column_config.TextColumn("Expense Description", required=True),
                "Category": st.column_config.SelectboxColumn("Category", options=sorted(list(set(category_options))), required=True),
                "Amount": st.column_config.NumberColumn("Amount (₪)", format="₪%.2f", min_value=0.0, required=True),
                "Notes": st.column_config.TextColumn("Notes", required=False),
            },
            column_order=["Date", "Business Name", "Category", "Amount", "Notes"],
            num_rows="dynamic",
            use_container_width=True,
            key="income_table_editor"
        )

        col_save_inc, _ = st.columns([2, 1])
        with col_save_inc:
            if st.button("💾 Save Income Changes to MasterData", type="primary", key="save_income_btn", use_container_width=True):
                # Determine untouched rows outside this cycle's income
                if full_df is not None and "_row_idx" in full_df.columns and "_row_idx" in income_df.columns:
                    orig_cycle_ids = set(income_df["_row_idx"].dropna())
                    remaining_full = full_df[~full_df["_row_idx"].isin(orig_cycle_ids)].copy()
                else:
                    remaining_full = full_df.copy() if full_df is not None else pd.DataFrame(columns=sheets.REQUIRED_COLUMNS)

                new_inc_rows = edited_income_df.copy()
                new_inc_rows["Type"] = "Income"
                
                # Normalize types
                for c in cols_to_keep:
                    if c not in remaining_full.columns:
                        remaining_full[c] = ""
                    if c not in new_inc_rows.columns:
                        new_inc_rows[c] = ""

                clean_full = pd.concat([remaining_full[cols_to_keep], new_inc_rows[cols_to_keep]], ignore_index=True)

                if is_demo_mode:
                    clean_full["Date_Clean"] = pd.to_datetime(clean_full["Date"], errors="coerce")
                    st.session_state["demo_df"] = clean_full
                    st.success("✅ [DEMO MODE] Income table updated successfully!")
                    st.rerun()
                else:
                    with st.spinner("Saving changes to Google Sheets 'MasterData'..."):
                        try:
                            sheets.overwrite_master_data(clean_full)
                            st.success("✅ Income changes saved to Google Sheets successfully!")
                            st.cache_data.clear()
                            st.rerun()
                        except Exception as e:
                            st.error(f"Failed to update Google Sheets: {e}")

    # Tab 2: Outcome / Expenses Table
    with tabs[1]:
        st.markdown("<h4 style='color: #f43f5e; margin-top: 0.5rem;'>Outcome / Expense Transactions</h4>", unsafe_allow_html=True)
        st.caption("✏️ Click any cell to edit. Select row(s) and press **Delete** to remove. Click **Save Expense Changes** when finished.")
        
        expense_df = df_current_cycle[df_current_cycle["Type"] == "Expense"].copy()
        for c in ["Date", "Business Name", "Category", "Amount", "Notes"]:
            if c not in expense_df.columns:
                expense_df[c] = ""
                
        total_exp = expense_df["Amount"].sum() if not expense_df.empty else 0.0
        st.caption(f"Currently showing **{len(expense_df)}** expense records totaling **{format_currency(total_exp)}**")

        edited_expense_df = st.data_editor(
            expense_df,
            column_config={
                "Date": st.column_config.TextColumn("Date", required=True),
                "Business Name": st.column_config.TextColumn("Expense Description", required=True),
                "Category": st.column_config.SelectboxColumn("Category", options=sorted(list(set(category_options))), required=True),
                "Amount": st.column_config.NumberColumn("Amount (₪)", format="₪%.2f", min_value=0.0, required=True),
                "Notes": st.column_config.TextColumn("Notes", required=False),
            },
            column_order=["Date", "Business Name", "Category", "Amount", "Notes"],
            num_rows="dynamic",
            use_container_width=True,
            key="expense_table_editor"
        )

        col_save_exp, _ = st.columns([2, 1])
        with col_save_exp:
            if st.button("💾 Save Expense Changes to MasterData", type="primary", key="save_expense_btn", use_container_width=True):
                if full_df is not None and "_row_idx" in full_df.columns and "_row_idx" in expense_df.columns:
                    orig_cycle_ids = set(expense_df["_row_idx"].dropna())
                    remaining_full = full_df[~full_df["_row_idx"].isin(orig_cycle_ids)].copy()
                else:
                    remaining_full = full_df.copy() if full_df is not None else pd.DataFrame(columns=sheets.REQUIRED_COLUMNS)

                new_exp_rows = edited_expense_df.copy()
                new_exp_rows["Type"] = "Expense"

                for c in cols_to_keep:
                    if c not in remaining_full.columns:
                        remaining_full[c] = ""
                    if c not in new_exp_rows.columns:
                        new_exp_rows[c] = ""

                clean_full = pd.concat([remaining_full[cols_to_keep], new_exp_rows[cols_to_keep]], ignore_index=True)

                if is_demo_mode:
                    clean_full["Date_Clean"] = pd.to_datetime(clean_full["Date"], errors="coerce")
                    st.session_state["demo_df"] = clean_full
                    st.success("✅ [DEMO MODE] Expense table updated successfully!")
                    st.rerun()
                else:
                    with st.spinner("Saving changes to Google Sheets 'MasterData'..."):
                        try:
                            sheets.overwrite_master_data(clean_full)
                            st.success("✅ Expense changes saved to Google Sheets successfully!")
                            st.cache_data.clear()
                            st.rerun()
                        except Exception as e:
                            st.error(f"Failed to update Google Sheets: {e}")

    # Tab 3: Manual Entry Form
    with tabs[2]:
        st.markdown("<h4 style='color: #38bdf8; margin-top: 0.5rem;'>Manual Entry Form</h4>", unsafe_allow_html=True)
        with st.form("manual_entry_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                manual_date = st.date_input("Transaction Date", value=date.today())
                manual_business = st.text_input("Expense Description", placeholder="e.g. Shufersal, Aroma, Salary")
                manual_type = st.selectbox("Type", ["Expense", "Income"], index=0)
            with col2:
                categories = list(parsers.CATEGORY_KEYWORDS.keys()) + ["Uncategorized", "Salary", "Investment Fund", "Other"]
                manual_category = st.selectbox("Category", sorted(list(set(categories))), index=0)
                manual_amount = st.number_input("Amount (₪)", min_value=0.01, step=10.0, format="%.2f")
                manual_notes = st.text_input("Notes", placeholder="Optional remarks...")

            submit_btn = st.form_submit_button("➕ Save Transaction to MasterData", use_container_width=True)
            if submit_btn:
                if not manual_business.strip():
                    st.error("Please enter an Expense Description.")
                elif manual_amount <= 0:
                    st.error("Amount must be greater than 0.")
                elif is_demo_mode:
                    new_entry = pd.DataFrame([{
                        "Date": manual_date.strftime("%Y-%m-%d"),
                        "Business Name": manual_business.strip(),
                        "Category": manual_category.strip(),
                        "Type": manual_type.strip(),
                        "Amount": float(manual_amount),
                        "Notes": manual_notes.strip(),
                        "Date_Clean": pd.to_datetime(manual_date)
                    }])
                    if "demo_df" in st.session_state:
                        st.session_state["demo_df"] = pd.concat([st.session_state["demo_df"], new_entry], ignore_index=True)
                    st.success(f"[DEMO MODE] Added '{manual_business}' ({format_currency(manual_amount)}) to preview.")
                    st.rerun()
                else:
                    with st.spinner("Saving to Google Sheets 'MasterData'..."):
                        try:
                            sheets.add_single_transaction(
                                manual_date,
                                manual_business,
                                manual_category,
                                manual_type,
                                manual_amount,
                                manual_notes
                            )
                            st.success(f"Transaction '{manual_business}' ({format_currency(manual_amount)}) saved successfully!")
                            st.cache_data.clear()
                            st.rerun()
                        except Exception as e:
                            st.error(f"Failed to save transaction: {e}")

    # Tab 4: Data Import Module
    with tabs[3]:
        st.markdown("<h4 style='color: #a78bfa; margin-top: 0.5rem;'>Data Import & Sync Module</h4>", unsafe_allow_html=True)
        st.caption("Upload Max Credit Card (.xlsx) or Isracard Statement (.pdf) to auto-categorize and sync.")
        
        uploaded_file = st.file_uploader(
            "Choose Max (.xlsx) or Isracard (.pdf)",
            type=["xlsx", "pdf"],
            key="modal_file_uploader"
        )

        if uploaded_file is not None:
            if "parsed_staging_df" not in st.session_state or st.session_state.get("last_uploaded_name") != uploaded_file.name:
                with st.spinner(f"Parsing {uploaded_file.name}..."):
                    try:
                        parsed_df = parsers.parse_uploaded_file(uploaded_file)
                        st.session_state["parsed_staging_df"] = parsed_df
                        st.session_state["last_uploaded_name"] = uploaded_file.name
                        st.success(f"Parsed {len(parsed_df)} transactions from {uploaded_file.name}!")
                    except Exception as e:
                        st.error(f"Error parsing file: {e}")
                        st.session_state["parsed_staging_df"] = None

        if st.session_state.get("parsed_staging_df") is not None and not st.session_state["parsed_staging_df"].empty:
            st.markdown("##### Review & Edit Grid")
            st.info("💡 Review and fix 'Uncategorized' entries or edit values before syncing:")

            category_options = list(parsers.CATEGORY_KEYWORDS.keys()) + ["Uncategorized", "Salary", "Investment Fund", "Other"]
            type_options = ["Expense", "Income"]

            edited_df = st.data_editor(
                st.session_state["parsed_staging_df"],
                column_config={
                    "Date": st.column_config.TextColumn(
                        "Date",
                        required=True
                    ),
                    "Business Name": st.column_config.TextColumn(
                        "Expense Description",
                        required=True
                    ),
                    "Category": st.column_config.SelectboxColumn(
                        "Category",
                        options=sorted(list(set(category_options))),
                        required=True
                    ),
                    "Type": st.column_config.SelectboxColumn(
                        "Type",
                        options=type_options,
                        required=True
                    ),
                    "Amount": st.column_config.NumberColumn(
                        "Amount (₪)",
                        format="₪%.2f",
                        min_value=0.0,
                        required=True
                    ),
                    "Notes": st.column_config.TextColumn(
                        "Notes",
                        help="Free text notes or comments (click or double-click to edit anytime)",
                        default="",
                        required=False
                    )
                },
                column_order=["Date", "Business Name", "Category", "Type", "Amount", "Notes"],
                use_container_width=True,
                num_rows="dynamic",
                key="staging_data_editor"
            )

            col_sync, col_discard = st.columns([2, 1])
            with col_sync:
                if st.button("☁️ Sync to Database", type="primary", use_container_width=True):
                    if is_demo_mode:
                        staging_copy = edited_df[cols_to_keep].copy()
                        staging_copy["Date_Clean"] = pd.to_datetime(staging_copy["Date"], errors="coerce")
                        if "demo_df" in st.session_state:
                            st.session_state["demo_df"] = pd.concat([st.session_state["demo_df"], staging_copy], ignore_index=True)
                        st.success(f"[DEMO MODE] Appended {len(edited_df)} rows to preview dataset!")
                        st.session_state["parsed_staging_df"] = None
                        st.session_state["last_uploaded_name"] = None
                        st.rerun()
                    else:
                        with st.spinner("Appending reviewed rows to Google Sheet 'Financial_DB' (MasterData)..."):
                            try:
                                count = sheets.append_transactions(edited_df)
                                st.success(f"Successfully synced {count} transactions to Google Sheets!")
                                st.session_state["parsed_staging_df"] = None
                                st.session_state["last_uploaded_name"] = None
                                st.cache_data.clear()
                                st.rerun()
                            except Exception as e:
                                st.error(f"Sync failed: {e}")
            with col_discard:
                if st.button("Clear Uploaded", use_container_width=True):
                    st.session_state["parsed_staging_df"] = None
                    st.session_state["last_uploaded_name"] = None
                    st.rerun()


TELEGRAM_CACHE_FILE = "telegram_cache.json"


def load_telegram_cache(channel_name: str):
    """Loads cached channel posts and timestamp from local disk."""
    if not os.path.exists(TELEGRAM_CACHE_FILE):
        return None, 0
    try:
        with open(TELEGRAM_CACHE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            ch_data = data.get(channel_name)
            if ch_data:
                ts = ch_data.get("timestamp", 0)
                posts = ch_data.get("posts", [])
                for p in posts:
                    if isinstance(p.get("date"), str):
                        try:
                            p["date"] = datetime.fromisoformat(p["date"])
                        except Exception:
                            p["date"] = datetime.now(timezone.utc)
                return posts, ts
    except Exception:
        pass
    return None, 0


def save_telegram_cache(channel_name: str, posts: list):
    """Saves fetched channel posts to local disk for 0ms cellular retrieval."""
    try:
        data = {}
        if os.path.exists(TELEGRAM_CACHE_FILE):
            try:
                with open(TELEGRAM_CACHE_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                data = {}

        serializable_posts = []
        for p in posts:
            p_copy = dict(p)
            if isinstance(p_copy.get("date"), datetime):
                p_copy["date"] = p_copy["date"].isoformat()
            serializable_posts.append(p_copy)

        data[channel_name] = {
            "timestamp": datetime.now().timestamp(),
            "posts": serializable_posts
        }
        with open(TELEGRAM_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


@st.cache_data(ttl=300)
def fetch_public_telegram_channel(channel_name: str, limit: int = 15):
    """
    Directly fetches public Telegram channel messages via https://t.me/s/<channel_name>
    Accelerated with local disk caching for instantaneous loading on mobile data.
    """
    cached_posts, last_ts = load_telegram_cache(channel_name)
    now_ts = datetime.now().timestamp()

    # If cached less than 30 minutes ago and has enough posts, return immediately
    if cached_posts and (now_ts - last_ts < 1800) and len(cached_posts) >= limit:
        return cached_posts[:limit]

    url = f"https://t.me/s/{channel_name}"
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }
    )
    try:
        html = urllib.request.urlopen(req, timeout=5).read().decode("utf-8", errors="ignore")
    except Exception:
        # Fall back to existing cached posts on network timeout/cellular issues
        if cached_posts:
            return cached_posts[:limit]
        return []

    pattern = re.compile(
        r'data-post="([^"]+)"[^>]*>.*?<div class="tgme_widget_message_text[^"]*"[^>]*>(.*?)</div>.*?<time datetime="([^"]+)"',
        re.DOTALL
    )
    matches = pattern.findall(html)
    results = []
    for data_post, raw_text, time_str in matches:
        parts = data_post.split('/')
        msg_id = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else parts[-1]
        
        # Clean text
        clean = re.sub(r'<br\s*/?>', '\n', raw_text)
        extracted_urls = re.findall(r'<a\s+[^>]*href="([^"]+)"[^>]*>', clean)
        clean = re.sub(r'<[^>]+>', '', clean).strip()

        try:
            dt = datetime.fromisoformat(time_str.replace("Z", "+00:00"))
        except Exception:
            dt = datetime.now(timezone.utc)

        web_links = [u for u in extracted_urls if "t.me" not in u.lower()]

        views = 0
        views_match = re.search(r'class="tgme_widget_message_views">([^<]+)<', raw_text)
        if views_match:
            v_str = views_match.group(1).replace("K", "000").replace(".", "").strip()
            views = int(v_str) if v_str.isdigit() else 0

        results.append({
            "id": msg_id,
            "data_post": data_post,
            "channel": channel_name,
            "text": clean,
            "date": dt,
            "views": views,
            "link": f"https://t.me/{data_post}",
            "tg_app_link": f"tg://resolve?domain={channel_name}&post={msg_id}",
            "web_links": web_links
        })

    results.reverse()
    if results:
        save_telegram_cache(channel_name, results)
        return results[:limit]
    
    if cached_posts:
        return cached_posts[:limit]
    return []


def get_sample_telegram_insights():
    """
    Returns curated, realistic & live messages for Investment Center & Shopping Center.
    First attempts to fetch public live messages from @fibiinvestmentchennel and @behatsdaa.
    Falls back gracefully to verified realistic items with functional links.
    """
    now = datetime.now()

    live_fibi = fetch_public_telegram_channel("fibiinvestmentchennel", limit=20)
    live_beh = fetch_public_telegram_channel("behatsdaa", limit=20)

    hever_sample = [
        {
            "id": "hvr_cinema",
            "date": now - relativedelta(hours=3),
            "channel": "hvr",
            "brand": "מועדון חבר",
            "views": 5400,
            "link": "https://www.hvr.co.il",
            "web_links": ["https://www.cinema-city.co.il"],
            "text": "🎬 מועדון 'חבר' - סרטי קולנוע ופופקורן:\n\nכרטיס לקולנוע (סינמה סיטי, יס פלאנט, הוט סינמה) ב-18 ₪ בלבד כולל שובר פופקורן ושתייה ב-14 ₪!\n\n🍿 המבצע תקף לכל ימות השבוע כולל סופי שבוע וחגים. מימוש מיידי באפליקציית חבר או באתר הרשת."
        },
        {
            "id": "hvr_food",
            "date": now - relativedelta(hours=8),
            "channel": "hvr",
            "brand": "מועדון חבר",
            "views": 6700,
            "link": "https://www.hvr.co.il",
            "web_links": ["https://wolt.com"],
            "text": "🍔 שוברי מסעדות ומשלוחים - מועדון חבר:\n\nWolt, תן ביס ורשתות קפה מובילות (ארומה, לנדוור, גרג) בהנחה קבועה של 20% על כל טעינה.\n\n☕ הקופון נטען מיידית בארנק הדיגיטלי וניתן לשימוש רציף: https://www.hvr.co.il"
        },
        {
            "id": "hvr_tech",
            "date": now - relativedelta(days=1),
            "channel": "hvr",
            "brand": "מועדון חבר",
            "views": 4200,
            "link": "https://www.hvr.co.il",
            "web_links": ["https://ksp.co.il"],
            "text": "💻 מוצרי אלקטרוניקה ומחשבים - מועדון חבר:\n\nהנחה מיוחדת של 12% נוספים ברשת KSP ואייבורי על מחשבים ניידים, מסכים וציוד היקפי עם קוד שובר ייעודי.\n\nפרטים מלאים באתר חבר: https://www.hvr.co.il"
        },
        {
            "id": "hvr_hotel",
            "date": now - relativedelta(days=2),
            "channel": "hvr",
            "brand": "מועדון חבר",
            "views": 5900,
            "link": "https://www.hvr.co.il",
            "web_links": ["https://www.isrotel.co.il"],
            "text": "🏨 נופש ומלונות בארץ - מועדון חבר:\n\nעד 30% הנחה ברשת מלונות ישרוטל ופתאל (אילת, ים המלח וירושלים) לחודשים הקרובים.\n\nכולל ארוחת בוקר וטיפול ספא בהנחה. הזמנות דרך פורטל מועדון חבר: https://www.hvr.co.il"
        }
    ]

    fallback_fibi = [
        {
            "id": 3557,
            "date": now - relativedelta(hours=2),
            "channel": "fibiinvestmentchennel",
            "views": 2450,
            "link": "https://t.me/fibiinvestmentchennel/3557",
            "tg_app_link": "tg://resolve?domain=fibiinvestmentchennel&post=3557",
            "text": "📊 סקירת מאקרו שבועית - בנק הבינלאומי (FIBI):\n\nמדד המחירים לצרכן בארה\"ב הצביע על התמתנות קלה בלחצי האינפלציה. אנו ממליצים לשמור על חשיפה של 60/40 במניות ואג\"ח, עם עדיפות לאפיק השקלי לטווח בינוני.\n\nנקודות מפתח להשקעה:\n• מניות טכנולוגיה (AI & Semiconductors): המשך מומנטום חיובי עם דוחות רבעוניים חזקים.\n• אג\"ח ממשלתי שקלי: תשואות אטרקטיביות במח\"מ 4-6 שנים.\n• דולר/שקל: יציבות סביב 3.65-3.70 ₪."
        },
        {
            "id": 3555,
            "date": now - relativedelta(hours=7),
            "channel": "fibiinvestmentchennel",
            "views": 1890,
            "link": "https://t.me/fibiinvestmentchennel/3555",
            "tg_app_link": "tg://resolve?domain=fibiinvestmentchennel&post=3555",
            "text": "💡 המלצת השקעה - סקטור האנרגיה הירוקה והתשתיות:\n\nבחינת הביצועים מראה כי חברות תשתית בעלות תזרים מזומנים יציב ודיבידנד גבוה מספקות הגנה אפקטיבית בתקופות תנודתיות. אנו ממליצים על שילוב מדדי תשתיות גלובליים כעוגן סולידי בתיק ההשקעות."
        },
        {
            "id": 3554,
            "date": now - relativedelta(days=1),
            "channel": "fibiinvestmentchennel",
            "views": 3120,
            "link": "https://t.me/fibiinvestmentchennel/3554",
            "tg_app_link": "tg://resolve?domain=fibiinvestmentchennel&post=3554",
            "text": "📈 ניתוח ריבית בנק ישראל והשלכותיה על שוק ההון:\n\nההערכה הרווחת היא כי בנק ישראל ימתין עם הורדות ריבית חדות עד להתבהרות המצב הגיאופוליטי. למשקיעים סולידיים: קרנות כספיות ופקדונות לטווח קצר ממשיכים להציע תשואה ריאלית חיובית נטולת סיכון אשראי."
        },
        {
            "id": 3552,
            "date": now - relativedelta(days=2),
            "channel": "fibiinvestmentchennel",
            "views": 2100,
            "link": "https://t.me/fibiinvestmentchennel/3552",
            "tg_app_link": "tg://resolve?domain=fibiinvestmentchennel&post=3552",
            "text": "🌐 קרנות סל (ETFs) מומלצות לחודש הקרוב:\n\n1. S&P 500 Equal Weight (RSP) - פיזור רחב יותר מצמצם תלות בענקיות הטק.\n2. מדדי דיבידנד גלובליים (VIG / SCHD) - תזרים שוטף ויציבות.\n3. אג\"ח קונצרני בדירוג השקעה גבוה (LQD) - מרווח תשואה מעניין מול הממשלתי."
        },
        {
            "id": 3550,
            "date": now - relativedelta(days=3),
            "channel": "fibiinvestmentchennel",
            "views": 1650,
            "link": "https://t.me/fibiinvestmentchennel/3550",
            "tg_app_link": "tg://resolve?domain=fibiinvestmentchennel&post=3550",
            "text": "🔍 טיפ פיננסי שבועי מחלקת המחקר:\n\nאל תנסו לתזמן את נקודות השפל בשוק. היסטורית, המשקיעים שהשיגו את התשואות הגבוהות ביותר הם אלו שהשקיעו בהוראת קבע חודשית קבועה (DCA) במדדים רחבים, ללא קשר לתנודות קצרות טווח."
        }
    ]

    fallback_beh = [
        {
            "id": 403,
            "date": now - relativedelta(hours=1),
            "channel": "behatsdaa",
            "brand": "בהצדעה",
            "views": 4300,
            "link": "https://t.me/behatsdaa/403",
            "tg_app_link": "tg://resolve?domain=behatsdaa&post=403",
            "web_links": ["https://www.behatsdaa.org.il"],
            "text": "🏷️ מבצע בלעדי למשרתי מילואים פעילים ב'בהצדעה'!\n\nתווי קנייה לרשתות המזון והפארם (שופרסל, קרפור, Be פארם) ב-18% הנחה נטענת ישירות לכרטיס!\n\n💳 טוענים 500 ₪ ומשלמים רק 410 ₪. מוגבל ל-2 טעינות לחודש קלנדרי: https://www.behatsdaa.org.il"
        },
        {
            "id": 402,
            "date": now - relativedelta(hours=9),
            "channel": "behatsdaa",
            "brand": "בהצדעה",
            "views": 3800,
            "link": "https://t.me/behatsdaa/402",
            "tg_app_link": "tg://resolve?domain=behatsdaa&post=402",
            "web_links": ["https://www.behatsdaa.org.il"],
            "text": "✈️ נופש חורף ומלונות בארץ ב'בהצדעה':\n\nעד 30% הנחה ברשת מלונות ישרוטל ופתאל (אילת, ים המלח וירושלים) לחודשים הקרובים.\n\n🏨 כולל ארוחת בוקר וטיפול ספא בהנחה. הזמנות דרך פורטל בהצדעה: https://www.behatsdaa.org.il"
        },
        {
            "id": 401,
            "date": now - relativedelta(days=2),
            "channel": "behatsdaa",
            "brand": "בהצדעה",
            "views": 4900,
            "link": "https://t.me/behatsdaa/401",
            "tg_app_link": "tg://resolve?domain=behatsdaa&post=401",
            "web_links": ["https://www.behatsdaa.org.il"],
            "text": "⛽ הנחת דלק ושטיפת רכב ב'בהצדעה':\n\nהנחה קבועה של 35 אגורות לליטר בנזין מתחת למחיר המרבי בתחנות סונול ודור אלון, ו-30% הנחה על שטיפה אוטומטית ברחבי הארץ.\n\nפרטים מלאים: https://www.behatsdaa.org.il"
        }
    ]

    investment_posts = live_fibi if live_fibi else fallback_fibi
    deals_source = live_beh if live_beh else fallback_beh
    for p in deals_source:
        if not p.get("brand"):
            p["brand"] = "בהצדעה"

    shopping_posts = sort_posts_by_date(deals_source + hever_sample)
    return investment_posts, shopping_posts


def format_tg_date(dt):
    """Formats Telegram datetime object or string."""
    if not dt:
        return ""
    if isinstance(dt, (datetime, pd.Timestamp)):
        return dt.strftime("%d/%m/%Y %H:%M")
    return str(dt)[:16]


def sort_posts_by_date(posts):
    """Sorts list of post dictionaries by datetime descending."""
    def get_sort_key(p):
        d = p.get("date")
        if isinstance(d, (datetime, pd.Timestamp)):
            if getattr(d, "tzinfo", None) is not None:
                return d
            return d.replace(tzinfo=timezone.utc)
        return datetime.min.replace(tzinfo=timezone.utc)
    return sorted(posts, key=get_sort_key, reverse=True)


def extract_web_links(text):
    """Finds http/https URLs inside the message text that are not Telegram links."""
    if not text:
        return []
    urls = re.findall(r'https?://[^\s<>"]+|www\.[^\s<>"]+', text)
    clean_urls = []
    for u in urls:
        clean = u.rstrip(".,;!?)\"'")
        if "t.me" not in clean.lower():
            if not clean.startswith("http"):
                clean = "https://" + clean
            clean_urls.append(clean)
    return clean_urls


def get_post_buttons_html(post):
    """
    Renders mobile-friendly, bulletproof action buttons for every post:
    1. For Hever ('חבר'):
       - Direct link to Hever portal (https://www.hvr.co.il)
       - If merchant link exists (e.g. Cinema City, Wolt, KSP), direct deal link
       - If shared on Behatsdaa channel, app + web Telegram links
    2. For Telegram posts:
       - 📱 Open in Telegram App (tg://resolve?domain=...&post=...) -> 1-tap opens Telegram on iPhone!
       - ↗ Web Link (https://t.me/...)
       - 🌐 Open Deal Link (if merchant URL found in text)
    """
    channel = (post.get("channel") or "").strip()
    msg_id = post.get("id")
    brand = post.get("brand", "")
    is_hever = ("חבר" in brand) or (channel == "hvr")
    
    web_links = post.get("web_links") or extract_web_links(post.get("text", ""))
    deal_url = web_links[0] if web_links else ""
    
    buttons = []
    
    if is_hever:
        buttons.append(
            '<a href="https://www.hvr.co.il" target="_blank" style="background: rgba(168, 85, 247, 0.18); color: #c084fc; border: 1px solid rgba(168, 85, 247, 0.4); padding: 6px 14px; border-radius: 9px; text-decoration: none; font-size: 0.84rem; font-weight: 600; display: inline-flex; align-items: center; gap: 5px;">💳 כניסה לאתר חבר</a>'
        )
        if deal_url and "hvr.co.il" not in deal_url:
            buttons.append(
                f'<a href="{deal_url}" target="_blank" style="background: rgba(16, 185, 129, 0.18); color: #34d399; border: 1px solid rgba(16, 185, 129, 0.4); padding: 6px 14px; border-radius: 9px; text-decoration: none; font-size: 0.84rem; font-weight: 600; display: inline-flex; align-items: center; gap: 5px;">🌐 כניסה למבצע</a>'
            )
        if channel == "behatsdaa" and msg_id:
            tg_app = f"tg://resolve?domain=behatsdaa&post={msg_id}"
            buttons.append(
                f'<a href="{tg_app}" style="background: rgba(244, 63, 94, 0.15); color: #fb7185; border: 1px solid rgba(244, 63, 94, 0.35); padding: 6px 14px; border-radius: 9px; text-decoration: none; font-size: 0.84rem; font-weight: 600; display: inline-flex; align-items: center; gap: 5px;">📱 טלגרם</a>'
            )
        return "".join(buttons)

    # Telegram Channel Post (FIBI or Behatsdaa)
    actual_ch = channel if channel else ("fibiinvestmentchennel" if "fibi" in str(post).lower() else "behatsdaa")
    tg_app_link = post.get("tg_app_link") or (f"tg://resolve?domain={actual_ch}&post={msg_id}" if msg_id else f"tg://resolve?domain={actual_ch}")
    tg_web_link = post.get("link") or (f"https://t.me/{actual_ch}/{msg_id}" if msg_id else f"https://t.me/{actual_ch}")

    if deal_url:
        buttons.append(
            f'<a href="{deal_url}" target="_blank" style="background: rgba(16, 185, 129, 0.18); color: #34d399; border: 1px solid rgba(16, 185, 129, 0.4); padding: 6px 14px; border-radius: 9px; text-decoration: none; font-size: 0.84rem; font-weight: 600; display: inline-flex; align-items: center; gap: 5px;">🌐 קישור להטבה</a>'
        )

    btn_color = "#60a5fa" if "fibi" in actual_ch.lower() else "#fb7185"
    btn_bg = "rgba(59, 130, 246, 0.18)" if "fibi" in actual_ch.lower() else "rgba(244, 63, 94, 0.18)"
    btn_border = "rgba(59, 130, 246, 0.4)" if "fibi" in actual_ch.lower() else "rgba(244, 63, 94, 0.4)"

    buttons.append(
        f'<a href="{tg_app_link}" style="background: {btn_bg}; color: {btn_color}; border: 1px solid {btn_border}; padding: 6px 14px; border-radius: 9px; text-decoration: none; font-size: 0.84rem; font-weight: 600; display: inline-flex; align-items: center; gap: 5px;">📱 פתח באפליקציה</a>'
    )
    buttons.append(
        f'<a href="{tg_web_link}" target="_blank" style="background: rgba(255, 255, 255, 0.06); color: #94a3b8; border: 1px solid rgba(255, 255, 255, 0.15); padding: 6px 12px; border-radius: 9px; text-decoration: none; font-size: 0.84rem; font-weight: 500; display: inline-flex; align-items: center; gap: 4px;">↗ Web</a>'
    )

    if "בהצדעה" in brand:
        buttons.append(
            '<a href="https://www.behatsdaa.org.il" target="_blank" style="background: rgba(255, 255, 255, 0.06); color: #94a3b8; border: 1px solid rgba(255, 255, 255, 0.15); padding: 6px 12px; border-radius: 9px; text-decoration: none; font-size: 0.84rem; font-weight: 500; display: inline-flex; align-items: center; gap: 4px;">💳 אתר בהצדעה</a>'
        )

    return "".join(buttons)


def render_market_insights_tab():
    """
    Renders the Market Insights tab powered by Telethon:
      1. Investment Center: 10 latest recommendations from FIBI channel
      2. Shopping Center: 10 latest sales from Behatsdaa and Hever
    """
    st.markdown("""
        <div style="margin-bottom: 1.2rem;">
            <h2 style="font-size: 1.45rem; font-weight: 700; margin-bottom: 0.2rem;">🌐 Market & Consumer Insights</h2>
            <div style="color: #94a3b8; font-size: 0.88rem;">
                Live Telegram intelligence: FIBI Investment Recommendations & Consumer Deals (בהצדעה, חבר)
            </div>
        </div>
    """, unsafe_allow_html=True)

    # Check Telethon User Client Authorization Status
    try:
        is_tg_connected = run_async(async_check_telegram_auth(TELEGRAM_API_ID, TELEGRAM_API_HASH, TELEGRAM_SESSION_NAME))
    except Exception:
        is_tg_connected = False

    # Top Control Bar
    top_c1, top_c2 = st.columns([3, 2])
    with top_c1:
        if is_tg_connected:
            st.success("🟢 **Telegram Client: Connected** as User Client")
        else:
            st.info("ℹ️ **Telegram Client: Not Connected** (Currently displaying preview insights). Connect below to fetch live channel posts.")

    with top_c2:
        btn_col1, btn_col2 = st.columns(2)
        with btn_col1:
            if st.button("🔄 Fetch Latest Posts", type="primary", use_container_width=True):
                if not is_tg_connected:
                    st.warning("Please connect your Telegram account first to fetch live channel posts.")
                else:
                    with st.spinner("Fetching latest 10 messages from whitelisted channels..."):
                        fibi_posts, err1 = run_async(fetch_telegram_channel_messages("fibiinvestmentchennel", limit=10))
                        behatsdaa_posts, err2 = run_async(fetch_telegram_channel_messages("behatsdaa", limit=10))
                        
                        if fibi_posts:
                            st.session_state["live_fibi_posts"] = fibi_posts
                        
                        if behatsdaa_posts:
                            st.session_state["live_deals_posts"] = behatsdaa_posts
                        
                        errors = [e for e in [err1, err2] if e and "UNAUTHORIZED" not in str(e)]
                        if errors:
                            st.warning(f"Fetch notes: {'; '.join(errors)}")
                        else:
                            st.success("Successfully fetched latest posts from all whitelisted channels!")
                            st.rerun()

        with btn_col2:
            if is_tg_connected:
                if st.button("🔓 Disconnect", use_container_width=True):
                    session_file = f"{TELEGRAM_SESSION_NAME}.session"
                    if os.path.exists(session_file):
                        try:
                            os.remove(session_file)
                        except Exception:
                            pass
                    st.session_state.pop("live_fibi_posts", None)
                    st.session_state.pop("live_deals_posts", None)
                    st.success("Disconnected Telegram session.")
                    st.rerun()

    # Telegram Login Expander if not connected
    if not is_tg_connected:
        with st.expander("📲 Connect Your Telegram Account (One-Time Setup)", expanded=False):
            st.caption("Authenticate as a Telegram User Client using your phone number to fetch public channels.")
            
            c_p1, c_p2 = st.columns([3, 1])
            with c_p1:
                phone_val = st.text_input(
                    "Phone Number (with international prefix)",
                    value=st.session_state.get("tg_login_phone", TELEGRAM_PHONE or "+972"),
                    placeholder="+972501234567"
                )
            with c_p2:
                st.write("")
                st.write("")
                if st.button("Send Code", use_container_width=True):
                    if not phone_val.strip() or len(phone_val.strip()) < 8:
                        st.error("Please enter a valid phone number.")
                    else:
                        with st.spinner("Requesting Telegram verification code..."):
                            try:
                                phone_code_hash = run_async(async_send_telegram_code(phone_val.strip(), TELEGRAM_API_ID, TELEGRAM_API_HASH, TELEGRAM_SESSION_NAME))
                                st.session_state["tg_phone_code_hash"] = phone_code_hash
                                st.session_state["tg_login_phone"] = phone_val.strip()
                                st.success(f"Code sent to {phone_val}! Check your Telegram app or SMS.")
                            except Exception as e:
                                st.error(f"Failed to send code: {e}")

            if st.session_state.get("tg_phone_code_hash"):
                c_c1, c_c2, c_c3 = st.columns([2, 2, 1])
                with c_c1:
                    code_val = st.text_input("Verification Code", placeholder="12345")
                with c_c2:
                    pwd_val = st.text_input("2FA Password (if enabled)", type="password", placeholder="Optional 2FA password")
                with c_c3:
                    st.write("")
                    st.write("")
                    if st.button("Verify & Login", type="primary", use_container_width=True):
                        with st.spinner("Signing in to Telegram..."):
                            success, err = run_async(async_sign_in_telegram(
                                st.session_state["tg_login_phone"],
                                code_val.strip(),
                                st.session_state["tg_phone_code_hash"],
                                pwd_val.strip() if pwd_val else None,
                                TELEGRAM_API_ID,
                                TELEGRAM_API_HASH,
                                TELEGRAM_SESSION_NAME
                            ))
                            if success:
                                st.success("Connected successfully! Loading channel posts...")
                                st.session_state.pop("tg_phone_code_hash", None)
                                st.rerun()
                            else:
                                st.error(f"Sign in failed: {err}")

    # Curated sample data fallback
    sample_investment, sample_shopping = get_sample_telegram_insights()

    # Active posts (live or preview)
    fibi_posts = st.session_state.get("live_fibi_posts", sample_investment)
    shopping_posts = st.session_state.get("live_deals_posts", sample_shopping)

    # Sub-tabs for Investment Center and Shopping Center
    subtab_inv, subtab_shop = st.tabs([
        "📈 Investment Center (FIBI)",
        "🛍️ Shopping Center (בהצדעה & חבר)"
    ])

    # 1. Investment Center
    with subtab_inv:
        c_inv_head, c_inv_pop = st.columns([3, 1.3])
        with c_inv_head:
            st.markdown("<h4 style='color: #3b82f6; margin-top: 0.3rem; margin-bottom: 0.2rem;'>📈 FIBI Investment Recommendations</h4>", unsafe_allow_html=True)
            st.caption("Financial analysis from **@fibiinvestmentchennel**")
        
        with c_inv_pop:
            st.write("")
            cur_q_inv = st.session_state.get("active_inv_query")
            pop_label_inv = f"🔍 Filter: {cur_q_inv[:8]}.." if cur_q_inv else "🔍 Search"
            with st.popover(pop_label_inv, use_container_width=True):
                st.markdown("##### 🔍 Search 1-Year FIBI Channel")
                search_inv = st.text_input("Keywords", value=st.session_state.get("search_inv_input", ""), placeholder="e.g. S&P, טכנולוגיה, אג\"ח", key="search_inv_input")
                c_ib1, c_ib2 = st.columns(2)
                with c_ib1:
                    btn_search_inv = st.button("Search", key="btn_search_inv", type="primary", use_container_width=True)
                with c_ib2:
                    btn_clear_inv = st.button("Clear", key="btn_clear_inv", use_container_width=True)

        if btn_clear_inv:
            st.session_state["search_inv_input"] = ""
            st.session_state.pop("active_inv_query", None)
            st.session_state.pop("inv_search_results", None)
            st.rerun()

        query_inv = (search_inv.strip() if 'search_inv' in locals() else st.session_state.get("search_inv_input", "")).strip()
        inv_limit = st.session_state.get("inv_limit", 10)

        if query_inv:
            if btn_search_inv or st.session_state.get("active_inv_query") != query_inv:
                if is_tg_connected:
                    with st.spinner(f"Searching @fibiinvestmentchennel for '{query_inv}'..."):
                        fetched_res, err = run_async(fetch_telegram_channel_messages("fibiinvestmentchennel", search_query=query_inv, days_back=365, limit=150))
                        if err and "UNAUTHORIZED" not in str(err):
                            st.warning(f"Search notice: {err}")
                        st.session_state["inv_search_results"] = fetched_res or []
                        st.session_state["active_inv_query"] = query_inv
                else:
                    kw = query_inv.lower()
                    st.session_state["inv_search_results"] = [p for p in sample_investment if kw in p["text"].lower()]
                    st.session_state["active_inv_query"] = query_inv

            full_inv_results = st.session_state.get("inv_search_results", [])
            display_inv_posts = full_inv_results[:inv_limit]
            
            c_tag1, c_tag2 = st.columns([3, 1])
            with c_tag1:
                st.caption(f"🎯 Showing **{len(display_inv_posts)}** of **{len(full_inv_results)}** results for **'{query_inv}'**.")
            with c_tag2:
                if st.button("✖ Reset Search", key="btn_reset_inv_tag", use_container_width=True):
                    st.session_state["search_inv_input"] = ""
                    st.session_state.pop("active_inv_query", None)
                    st.session_state.pop("inv_search_results", None)
                    st.rerun()
            if not is_tg_connected:
                st.info("ℹ️ Telegram user client not connected. Showing local preview matches. Connect above to search live Telegram channel up to 1 year back.")
        else:
            st.session_state.pop("active_inv_query", None)
            st.session_state.pop("inv_search_results", None)
            display_inv_posts = fibi_posts[:inv_limit]
            st.caption(f"Displaying latest **{len(display_inv_posts)}** recommendations from @fibiinvestmentchennel")

        if not display_inv_posts:
            st.info("No recommendations match your search.")
        else:
            for post in display_inv_posts:
                post_date = format_tg_date(post.get("date"))
                views_count = f"👁️ {post.get('views', 0):,}" if post.get('views') else ""
                buttons_html = get_post_buttons_html(post)

                card_html = f"""
                <div class="insight-card">
                    <div class="insight-header">
                        <span class="insight-badge-fibi">📊 FIBI Investment</span>
                        <span class="insight-date">{post_date}</span>
                    </div>
                    <div class="insight-body">{post['text']}</div>
                    <div class="insight-footer">
                        <span>{views_count}</span>
                        <div style="display: flex; gap: 8px; align-items: center; flex-wrap: wrap;">
                            {buttons_html}
                        </div>
                    </div>
                </div>
                """
                st.markdown(card_html, unsafe_allow_html=True)

            # Load More Button for Investment Center - Always active and accessible
            st.markdown("<div style='margin-top: 1rem; margin-bottom: 1.5rem;'>", unsafe_allow_html=True)
            if st.button("⬇️ Load More Recommendations (+10)", key="btn_load_more_inv", use_container_width=True):
                new_lim = inv_limit + 10
                st.session_state["inv_limit"] = new_lim
                if is_tg_connected and not query_inv:
                    with st.spinner("Fetching more from @fibiinvestmentchennel..."):
                        more_p, _ = run_async(fetch_telegram_channel_messages("fibiinvestmentchennel", limit=new_lim))
                        if more_p:
                            st.session_state["live_fibi_posts"] = more_p
                else:
                    more_public = fetch_public_telegram_channel("fibiinvestmentchennel", limit=new_lim)
                    if more_public:
                        st.session_state["live_fibi_posts"] = more_public
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

    # 2. Shopping Center
    with subtab_shop:
        c_shop_head, c_shop_pop = st.columns([3, 1.3])
        with c_shop_head:
            st.markdown("<h4 style='color: #f43f5e; margin-top: 0.3rem; margin-bottom: 0.2rem;'>🛍️ Consumer Sales & Deals Center</h4>", unsafe_allow_html=True)
            st.caption("Sales, discount vouchers, and perks from **בהצדעה** ומועדון חבר")
        
        with c_shop_pop:
            st.write("")
            cur_q_shop = st.session_state.get("active_shop_query")
            pop_label_shop = "🔍 Filter & Search" if not cur_q_shop else "🔍 Filter (Active)"
            with st.popover(pop_label_shop, use_container_width=True):
                st.markdown("##### 🔍 Search Deals & Filter Club")
                search_shop = st.text_input("Keywords", value=st.session_state.get("search_shop_input", ""), placeholder="e.g. שופרסל, דלק, קולנוע, וולט, ACE", key="search_shop_input")
                brand_filter = st.radio(
                    "Filter Source",
                    options=["All Deals", "בהצדעה (Behatsdaa)", "מועדון חבר (Hever)"],
                    index=["All Deals", "בהצדעה (Behatsdaa)", "מועדון חבר (Hever)"].index(st.session_state.get("shop_brand_filter", "All Deals")),
                    horizontal=True,
                    key="shop_brand_filter_radio"
                )
                st.session_state["shop_brand_filter"] = brand_filter
                c_sb1, c_sb2 = st.columns(2)
                with c_sb1:
                    btn_search_shop = st.button("Apply", key="btn_search_shop", type="primary", use_container_width=True)
                with c_sb2:
                    btn_clear_shop = st.button("Reset", key="btn_clear_shop", use_container_width=True)

        if btn_clear_shop:
            st.session_state["search_shop_input"] = ""
            st.session_state["shop_brand_filter"] = "All Deals"
            st.session_state.pop("active_shop_query", None)
            st.session_state.pop("shop_search_results", None)
            st.rerun()

        brand_filter = st.session_state.get("shop_brand_filter", "All Deals")
        shop_limit = st.session_state.get("shop_limit", 10)
        query_shop = (search_shop.strip() if 'search_shop' in locals() else st.session_state.get("search_shop_input", "")).strip()

        if query_shop or brand_filter != "All Deals":
            shop_state_key = f"{brand_filter}_{query_shop}"
            if btn_search_shop or st.session_state.get("active_shop_query") != shop_state_key:
                if is_tg_connected:
                    with st.spinner(f"Searching origin Telegram channels for '{query_shop}'..."):
                        target_channels = ["behatsdaa"] if brand_filter in ["All Deals", "בהצדעה (Behatsdaa)"] else []
                        
                        all_fetched = []
                        for ch in target_channels:
                            res, err = run_async(fetch_telegram_channel_messages(ch, search_query=query_shop if query_shop else None, days_back=365, limit=150))
                            if res:
                                all_fetched.extend(res)
                        
                        # Add any sample Hever matches if Hever or All Deals selected
                        if brand_filter in ["All Deals", "מועדון חבר (Hever)"]:
                            kw = query_shop.lower()
                            hever_matches = [p for p in sample_shopping if ("חבר" in p.get("brand", "") or p.get("channel") == "behatsdaa") and (not kw or kw in p["text"].lower())]
                            all_fetched.extend(hever_matches)

                        display_shop_posts_raw = sort_posts_by_date(all_fetched)
                        st.session_state["shop_search_results"] = display_shop_posts_raw
                        st.session_state["active_shop_query"] = shop_state_key
                else:
                    # Fallback to local sample filtering
                    kw = query_shop.lower()
                    filtered = [p for p in shopping_posts if not kw or kw in p["text"].lower()]
                    if brand_filter == "בהצדעה (Behatsdaa)":
                        filtered = [p for p in filtered if "בהצדעה" in p.get("brand", "") or p.get("channel") == "behatsdaa"]
                    elif brand_filter == "מועדון חבר (Hever)":
                        filtered = [p for p in filtered if "חבר" in p.get("brand", "") or p.get("channel") == "behatsdaa"]
                    st.session_state["shop_search_results"] = sort_posts_by_date(filtered)
                    st.session_state["active_shop_query"] = shop_state_key

            full_shop_results = st.session_state.get("shop_search_results", [])
            display_shop_posts = full_shop_results[:shop_limit]
            
            c_stag1, c_stag2 = st.columns([3, 1])
            with c_stag1:
                filter_desc = f"'{query_shop}'" if query_shop else brand_filter
                st.caption(f"🎯 Showing **{len(display_shop_posts)}** of **{len(full_shop_results)}** deals ({filter_desc}).")
            with c_stag2:
                if st.button("✖ Reset", key="btn_reset_shop_tag", use_container_width=True):
                    st.session_state["search_shop_input"] = ""
                    st.session_state["shop_brand_filter"] = "All Deals"
                    st.session_state.pop("active_shop_query", None)
                    st.session_state.pop("shop_search_results", None)
                    st.rerun()

            if not is_tg_connected:
                st.info("ℹ️ Telegram user client not connected. Showing local preview matches. Connect above to search live Telegram channels up to 1 year back.")
        else:
            st.session_state.pop("active_shop_query", None)
            st.session_state.pop("shop_search_results", None)
            display_shop_posts = shopping_posts[:shop_limit]
            st.caption(f"Displaying latest **{len(display_shop_posts)}** sales & perks.")

        if not display_shop_posts:
            st.info("No deals match your search criteria.")
        else:
            for post in display_shop_posts:
                post_date = format_tg_date(post.get("date"))
                brand = post.get("brand")
                if not brand:
                    ch_str = str(post.get("channel", ""))
                    if "behatsdaa" in ch_str:
                        brand = "בהצדעה"
                    else:
                        brand = "מועדון חבר"

                badge_class = "insight-badge-deals" if "בהצדעה" in brand else "insight-badge-hever"
                buttons_html = get_post_buttons_html(post)

                card_html = f"""
                <div class="insight-card">
                    <div class="insight-header">
                        <span class="{badge_class}">🏷️ {brand}</span>
                        <span class="insight-date">{post_date}</span>
                    </div>
                    <div class="insight-body">{post['text']}</div>
                    <div class="insight-footer">
                        <span>{views_count}</span>
                        <div style="display: flex; gap: 8px; align-items: center; flex-wrap: wrap;">
                            {buttons_html}
                        </div>
                    </div>
                </div>
                """
                st.markdown(card_html, unsafe_allow_html=True)

            # Load More Button for Shopping Center - Always active and accessible
            st.markdown("<div style='margin-top: 1rem; margin-bottom: 1.5rem;'>", unsafe_allow_html=True)
            if st.button("⬇️ Load More Deals (+10)", key="btn_load_more_shop", use_container_width=True):
                new_shop_lim = shop_limit + 10
                st.session_state["shop_limit"] = new_shop_lim
                if is_tg_connected and not query_shop:
                    with st.spinner("Fetching more deals from Telegram..."):
                        more_d, _ = run_async(fetch_telegram_channel_messages("behatsdaa", limit=new_shop_lim))
                        if more_d:
                            st.session_state["live_deals_posts"] = more_d
                else:
                    more_public_beh = fetch_public_telegram_channel("behatsdaa", limit=new_shop_lim)
                    if more_public_beh:
                        for p in more_public_beh:
                            if not p.get("brand"):
                                p["brand"] = "בהצדעה"
                        st.session_state["live_deals_posts"] = sort_posts_by_date(more_public_beh + sample_shopping)
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)


# --- Main Dashboard Logic ---

def main():
    # Top Header
    st.markdown("""
        <div class="dashboard-header">
            <div>
                <h1 class="dashboard-title">⚡ Financial Dashboard</h1>
                <div class="dashboard-subtitle">Personal Finance Analytics • Google Sheets Integration</div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    # Check Google Credentials (Service Account or OAuth Desktop Secret)
    client_secret_path = auth.find_client_secret_file()
    service_account_path = auth.find_service_account_file()
    if not client_secret_path and not service_account_path and not sheets.is_local_storage_enabled():
        st.error(
            "⚠️ Google Credentials JSON not found! "
            "Please ensure your client_secret_*.json or service_account.json is in the 'json/' directory."
        )
        return

    # Sidebar: Appearance, Account, Mode & Quick Import
    with st.sidebar:
        st.markdown("### 🎨 Appearance")
        current_theme = st.session_state.get("theme", "dark")
        theme_choice = st.radio(
            "Select Theme",
            options=["🌙 Dark Mode", "☀️ Light Mode"],
            index=0 if current_theme == "dark" else 1,
            horizontal=True,
            label_visibility="collapsed"
        )
        selected_theme = "dark" if "Dark" in theme_choice else "light"
        if selected_theme != current_theme:
            st.session_state["theme"] = selected_theme
            st.rerun()

        st.markdown("---")
        st.markdown("### ⚙️ System & Connection")
        is_auth = auth.is_authenticated()
        is_local = sheets.is_local_storage_enabled() and not is_auth
        auth_type = auth.get_auth_type()

        if is_auth:
            if auth_type == "service_account":
                st.success("🟢 Storage: **Google Sheets (Service Account)**")
                st.caption("⚡ 24/7 Headless Mobile Sync Active")
            else:
                st.success("🟢 Storage: **Google Sheets (Financial_DB)**")
            if st.button("🔓 Sign Out / Disconnect", use_container_width=True):
                auth.logout()
                st.session_state.clear()
                st.rerun()
        elif is_local:
            st.success("💾 Storage: **Local Database (MasterData.csv)**")
            if st.button("🌐 Connect Google Sheets", use_container_width=True):
                sheets.disable_local_storage()
                st.session_state["show_storage_setup"] = True
                st.rerun()
        else:
            st.warning("🟠 Storage: **Not Connected**")
            col_sb_a, col_sb_b = st.columns(2)
            with col_sb_a:
                if st.button("🌐 Google Sheets", use_container_width=True):
                    st.session_state["show_storage_setup"] = True
                    st.rerun()
            with col_sb_b:
                if st.button("💾 Local DB", use_container_width=True):
                    sheets.enable_local_storage()
                    st.session_state["show_storage_setup"] = False
                    st.rerun()

        demo_default = (not is_auth and not is_local and not st.session_state.get("show_storage_setup", False))
        demo_mode = st.toggle(
            "🧪 Demo / Preview Mode",
            value=st.session_state.get("demo_mode", demo_default),
            help="Preview dashboard with sample transactions",
            key="demo_toggle"
        )
        st.session_state["demo_mode"] = demo_mode

        st.markdown("---")
        st.markdown("### 📥 Quick Statement Upload")
        st.caption("Upload Max .xlsx or Isracard .pdf:")
        sidebar_file = st.file_uploader("Upload file", type=["xlsx", "pdf"], key="sidebar_uploader")
        if sidebar_file:
            if "parsed_staging_df" not in st.session_state or st.session_state.get("last_uploaded_name") != sidebar_file.name:
                try:
                    df_up = parsers.parse_uploaded_file(sidebar_file)
                    st.session_state["parsed_staging_df"] = df_up
                    st.session_state["last_uploaded_name"] = sidebar_file.name
                    st.success(f"Parsed {len(df_up)} rows! Click 'Transaction Center' to review and sync.")
                except Exception as err:
                    st.error(f"Parse error: {err}")

    main_tab_dash, main_tab_insights = st.tabs(["📊 Financial Dashboard", "🌐 Market Insights"])

    with main_tab_insights:
        render_market_insights_tab()

    with main_tab_dash:
        # Check whether to display the Storage Setup / Google Sign-in screen
        show_setup = st.session_state.get("show_storage_setup", False) or (not is_auth and not is_local and not demo_mode)
        
        if show_setup and not is_auth:
            local_ip = get_local_ip()

            if is_local or demo_mode:
                if st.button("← Back to Dashboard", key="btn_back_dash"):
                    st.session_state["show_storage_setup"] = False
                    st.rerun()

            st.markdown("""
                <div class="chart-box" style="text-align: center; padding: 1.8rem 1.2rem; max-width: 760px; margin: 1.2rem auto 1.5rem auto;">
                    <h2 style="color: #f3f4f6; margin-bottom: 0.4rem; font-size: 1.5rem;">🔗 Choose Your Database Storage</h2>
                    <p style="color: #94a3b8; font-size: 0.92rem; margin-bottom: 0.2rem;">
                        Connect to Google Sheet <b>Financial_DB</b> for 24/7 mobile sync, or use the <b>Local Database (MasterData.csv)</b>.
                    </p>
                </div>
            """, unsafe_allow_html=True)

            # Generate or preserve OAuth authorization URL
            if "oauth_auth_url" not in st.session_state or "oauth_flow" not in st.session_state:
                try:
                    auth_url, auth_state, oauth_flow = auth.get_authorization_url(redirect_uri="http://localhost:8080/")
                    st.session_state["oauth_auth_url"] = auth_url
                    st.session_state["oauth_state"] = auth_state
                    st.session_state["oauth_flow"] = oauth_flow
                except Exception as e:
                    st.error(f"Error initializing Google OAuth: {e}")

            auth_url = st.session_state.get("oauth_auth_url", "")

            col_auth_left, col_auth_right = st.columns([1, 1], gap="large")

            with col_auth_left:
                st.markdown("#### 🌐 Option 1: Google Sheets (Recommended for Mobile)")
                st.caption("Step 1: Open Google Sign-In with your granted account:")
                if auth_url:
                    st.link_button("🌐 Step 1: Open Google Sign-In", auth_url, use_container_width=True)
                
                st.write("")
                st.caption("Step 2: Paste the redirect URL or authorization code:")
                code_input = st.text_input("Paste redirect URL or code:", placeholder="http://localhost:8080/?state=...&code=4/0A...", key="oauth_paste_input")
                
                if st.button("✅ Complete Connection", type="primary", use_container_width=True):
                    if not code_input.strip():
                        st.error("Please paste the redirect URL or code from your browser.")
                    else:
                        with st.spinner("Connecting to Google Sheets..."):
                            try:
                                auth.complete_auth_with_code(code_input.strip(), flow=st.session_state.get("oauth_flow"))
                                sheets.disable_local_storage()
                                st.session_state["show_storage_setup"] = False
                                st.session_state["demo_mode"] = False
                                st.success("Connected successfully to Google Sheets!")
                                st.rerun()
                            except Exception as ex:
                                st.error(f"Authentication failed: {ex}")

            with col_auth_right:
                st.markdown("#### 💾 Option 2: Local Database (MasterData.csv)")
                st.caption("Store all transactions on this PC without Google Cloud:")
                if st.button("🚀 Activate Local Database", use_container_width=True):
                    sheets.enable_local_storage()
                    st.session_state["show_storage_setup"] = False
                    st.session_state["demo_mode"] = False
                    st.success("Local Database activated!")
                    st.rerun()

                st.markdown("---")
                st.markdown("#### 📱 Access From Your iPhone")
                st.markdown(f"""
                When running on your home Wi-Fi, open Safari on your iPhone and visit:
                <div style="background: rgba(16, 185, 129, 0.12); border: 1px solid #10b981; border-radius: 8px; padding: 0.6rem; text-align: center; margin: 0.5rem 0; font-family: monospace; font-size: 1.05rem; color: #34d399;">
                    <b>http://{local_ip}:8501</b>
                </div>
                """, unsafe_allow_html=True)

                if st.button("Preview With Demo Data", use_container_width=True):
                    st.session_state["demo_mode"] = True
                    st.session_state["show_storage_setup"] = False
                    st.rerun()
        else:
            # Fetch Data
            if demo_mode:
                if "demo_df" not in st.session_state:
                    st.session_state["demo_df"] = get_demo_data()
                raw_df = st.session_state["demo_df"].copy()
                st.info("ℹ️ Currently viewing in **Demo Mode**. You can switch to Local Storage or Google Sheets via the sidebar.")
            elif is_local:
                with st.spinner("Loading transactions from local database (MasterData.csv)..."):
                    try:
                        raw_df = sheets.fetch_master_data()
                    except Exception as e:
                        st.error(f"Error accessing local database: {e}")
                        raw_df = pd.DataFrame(columns=sheets.REQUIRED_COLUMNS)
            else:
                with st.spinner("Loading transactions from Google Sheet 'Financial_DB'..."):
                    try:
                        raw_df = sheets.fetch_master_data()
                    except Exception as e:
                        st.error(f"Error accessing Google Sheet 'Financial_DB': {e}")
                        st.info("Make sure the sheet 'Financial_DB' exists in your Google Drive.")
                        raw_df = pd.DataFrame(columns=sheets.REQUIRED_COLUMNS)

            # Assign internal row tracker index to track edited/deleted rows
            if not raw_df.empty:
                raw_df["_row_idx"] = list(range(len(raw_df)))
                if "Category" in raw_df.columns:
                    raw_df["Category"] = raw_df["Category"].replace({"Going Out": "Shopping & Going Out", "Shopping": "Shopping & Going Out", "Education & Kids": "Shopping & Going Out"})
            else:
                raw_df["_row_idx"] = []

            # Billing Cycle Calculation (9th of prev month to 8th of current month, All Time, or Custom Range)
            cycle_options = generate_cycle_options()
            active_idx = next((i for i, c in enumerate(cycle_options) if c["is_current"]), 0)
            cycle_labels = [c["label"] for c in cycle_options] + ["All Time", "📅 Custom Date Range"]

            # Minimalist Top Bar: Subtle Billing Cycle + Transactions modal button
            c_cycle_col, c_btn_col = st.columns([2.8, 1.2], gap="small")

            with c_cycle_col:
                selected_cycle_label = st.selectbox(
                    "Billing Cycle",
                    options=cycle_labels,
                    index=active_idx,
                    label_visibility="collapsed",
                    help="Select active cycle, past cycles, All Time, or a Custom Date Range"
                )

            # If Custom Date Range is selected, render manual range date pickers
            custom_start_val = date.today() - relativedelta(months=1)
            custom_end_val = date.today()
            if selected_cycle_label == "📅 Custom Date Range":
                st.markdown("<div style='margin-top: 0.2rem; margin-bottom: 0.5rem;'>", unsafe_allow_html=True)
                cd_col1, cd_col2 = st.columns(2)
                with cd_col1:
                    custom_start_val = st.date_input("From Date", value=st.session_state.get("custom_start_val", custom_start_val), key="manual_range_start")
                    st.session_state["custom_start_val"] = custom_start_val
                with cd_col2:
                    custom_end_val = st.date_input("To Date", value=st.session_state.get("custom_end_val", custom_end_val), key="manual_range_end")
                    st.session_state["custom_end_val"] = custom_end_val
                st.markdown("</div>", unsafe_allow_html=True)

            # Filter by selected cycle or manual custom date range
            if selected_cycle_label == "All Time":
                df_cycle = raw_df.copy()
            elif selected_cycle_label == "📅 Custom Date Range":
                start_d = pd.to_datetime(custom_start_val)
                end_d = pd.to_datetime(custom_end_val) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
                if not raw_df.empty and "Date_Clean" in raw_df.columns:
                    df_cycle = raw_df[(raw_df["Date_Clean"] >= start_d) & (raw_df["Date_Clean"] <= end_d)].copy()
                else:
                    df_cycle = pd.DataFrame(columns=sheets.REQUIRED_COLUMNS + ["_row_idx"])
            else:
                chosen_cycle = next(c for c in cycle_options if c["label"] == selected_cycle_label)
                start_d = pd.to_datetime(chosen_cycle["start"])
                end_d = pd.to_datetime(chosen_cycle["end"]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
                
                if not raw_df.empty and "Date_Clean" in raw_df.columns:
                    df_cycle = raw_df[(raw_df["Date_Clean"] >= start_d) & (raw_df["Date_Clean"] <= end_d)].copy()
                else:
                    df_cycle = pd.DataFrame(columns=sheets.REQUIRED_COLUMNS + ["_row_idx"])

            with c_btn_col:
                if st.button("📂 Transactions", type="secondary", use_container_width=True):
                    show_transaction_center_dialog(df_cycle, is_demo_mode=demo_mode, full_df=raw_df)

            # Uncategorized Transactions Detection & Quick Assistant Banner
            uncat_mask = (raw_df["Category"].fillna("").str.lower().str.strip() == "uncategorized") if not raw_df.empty else pd.Series([], dtype=bool)
            uncat_count = int(uncat_mask.sum())
            if uncat_count > 0:
                c_uc1, c_uc2 = st.columns([3.1, 1.3])
                with c_uc1:
                    st.markdown(f"""
                    <div style="background: rgba(245, 158, 11, 0.12); border: 1px solid rgba(245, 158, 11, 0.3); border-radius: 8px; padding: 0.45rem 0.8rem; margin-top: 0.2rem;">
                        <span style="color: #fbbf24; font-weight: 600; font-size: 0.86rem;">⚠️ {uncat_count} Uncategorized Transactions detected</span>
                    </div>
                    """, unsafe_allow_html=True)
                with c_uc2:
                    if st.button(f"⚡ 1-Tap Fix ({uncat_count})", type="primary", key="btn_open_quick_uncat", use_container_width=True):
                        show_quick_categorize_dialog(raw_df[uncat_mask].copy(), raw_df, is_demo_mode=demo_mode)

            # Clean filtered_df sorted by newest first
            filtered_df = df_cycle.copy()
            if not filtered_df.empty and "Date_Clean" in filtered_df.columns:
                filtered_df = filtered_df.sort_values(by="Date_Clean", ascending=False)

            # --- TOP SECTION: Glowing KPIs with Cycle-over-Cycle Trend Badges ---
            income_total = filtered_df[filtered_df["Type"] == "Income"]["Amount"].sum() if not filtered_df.empty else 0.0
            expense_total = filtered_df[filtered_df["Type"] == "Expense"]["Amount"].sum() if not filtered_df.empty else 0.0
            net_flow = income_total - expense_total

            net_class = "net-pos" if net_flow >= 0 else "net-neg"
            net_card_class = "kpi-net-positive" if net_flow >= 0 else "kpi-net-negative"

            # Calculate Previous Cycle Baseline for Trend Badges
            prev_cycle = None
            if selected_cycle_label not in ["All Time", "📅 Custom Date Range"]:
                cycle_idx = next((i for i, c in enumerate(cycle_options) if c["label"] == selected_cycle_label), None)
                if cycle_idx is not None and cycle_idx > 0:
                    prev_cycle = cycle_options[cycle_idx - 1]

            prev_income = 0.0
            prev_expense = 0.0
            if prev_cycle and not raw_df.empty and "Date_Clean" in raw_df.columns:
                p_start = pd.to_datetime(prev_cycle["start"])
                p_end = pd.to_datetime(prev_cycle["end"]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
                prev_sub = raw_df[(raw_df["Date_Clean"] >= p_start) & (raw_df["Date_Clean"] <= p_end)]
                prev_income = prev_sub[prev_sub["Type"] == "Income"]["Amount"].sum()
                prev_expense = prev_sub[prev_sub["Type"] == "Expense"]["Amount"].sum()

            prev_net = prev_income - prev_expense

            # Trend Badges for KPIs
            if prev_cycle and prev_expense > 0:
                exp_pct_diff = ((expense_total - prev_expense) / prev_expense) * 100
                if exp_pct_diff <= 0:
                    exp_subtext = f"<span style='color:#10b981; font-weight:600;'>▼ {abs(exp_pct_diff):.1f}% vs last cycle</span>"
                else:
                    exp_subtext = f"<span style='color:#f43f5e; font-weight:600;'>▲ {abs(exp_pct_diff):.1f}% vs last cycle</span>"
            else:
                exp_subtext = f"{len(filtered_df[filtered_df['Type'] == 'Expense'])} debit transactions"

            if prev_cycle and prev_income > 0:
                inc_pct_diff = ((income_total - prev_income) / prev_income) * 100
                if inc_pct_diff >= 0:
                    inc_subtext = f"<span style='color:#10b981; font-weight:600;'>▲ {abs(inc_pct_diff):.1f}% vs last cycle</span>"
                else:
                    inc_subtext = f"<span style='color:#f43f5e; font-weight:600;'>▼ {abs(inc_pct_diff):.1f}% vs last cycle</span>"
            else:
                inc_subtext = f"{len(filtered_df[filtered_df['Type'] == 'Income'])} credit transactions"

            if prev_cycle:
                net_delta = net_flow - prev_net
                if net_delta >= 0:
                    net_subtext = f"<span style='color:#10b981; font-weight:600;'>▲ +₪{net_delta:,.0f} vs last cycle</span>"
                else:
                    net_subtext = f"<span style='color:#f43f5e; font-weight:600;'>▼ -₪{abs(net_delta):,.0f} vs last cycle</span>"
            else:
                net_subtext = 'Savings positive' if net_flow >= 0 else 'Deficit in cycle'

            kpi_html = f"""
            <div class="kpi-container">
                <div class="kpi-card kpi-income">
                    <div class="kpi-label">
                        <span>Total Income</span>
                        <span>⬆</span>
                    </div>
                    <div class="kpi-value income">{format_currency(income_total)}</div>
                    <div class="kpi-subtext">{inc_subtext}</div>
                </div>
                <div class="kpi-card kpi-expense">
                    <div class="kpi-label">
                        <span>Total Expenses</span>
                        <span>⬇</span>
                    </div>
                    <div class="kpi-value expense">{format_currency(expense_total)}</div>
                    <div class="kpi-subtext">{exp_subtext}</div>
                </div>
                <div class="kpi-card {net_card_class}">
                    <div class="kpi-label">
                        <span>Net Flow</span>
                        <span>{'⚖' if net_flow >= 0 else '⚠'}</span>
                    </div>
                    <div class="kpi-value {net_class}">{format_currency(net_flow)}</div>
                    <div class="kpi-subtext">{net_subtext}</div>
                </div>
            </div>
            """
            st.markdown(kpi_html, unsafe_allow_html=True)

            # --- MIDDLE SECTION: Merged Category Breakdown & 1-Tap Drilldown ---
            is_dark = (st.session_state.get("theme", "dark") == "dark")
            chart_font_color = "#94a3b8" if is_dark else "#475569"
            chart_title_color = "#f3f4f6" if is_dark else "#0f172a"
            donut_border = "#14141e" if is_dark else "#ffffff"

            expense_sub_df = filtered_df[filtered_df["Type"] == "Expense"].copy()
            if not expense_sub_df.empty and "Category" in expense_sub_df.columns:
                expense_sub_df["Category"] = expense_sub_df["Category"].replace({"Going Out": "Shopping & Going Out", "Shopping": "Shopping & Going Out", "Education & Kids": "Shopping & Going Out"})

            # Classic Executive Financial Color Palette
            classic_palette = [
                "#2563eb",  # Classic Royal Blue
                "#0d9488",  # Forest Teal
                "#d97706",  # Classic Amber
                "#7c3aed",  # Deep Violet
                "#e11d48",  # Rose Red
                "#0284c7",  # Sky Blue
                "#16a34a",  # Emerald Green
                "#9333ea",  # Purple
                "#ea580c",  # Warm Orange
                "#475569",  # Slate
                "#0891b2",  # Cyan
                "#64748b"   # Muted Grey
            ]

            if expense_sub_df.empty or expense_sub_df["Amount"].sum() == 0:
                st.info("No expense records to display in category breakdown for this filter.")
            else:
                cat_summary = expense_sub_df.groupby("Category")["Amount"].sum().reset_index()
                # Filter out non-positive totals to prevent px.pie errors
                cat_summary = cat_summary[cat_summary["Amount"] > 0].copy()
                cat_summary = cat_summary.sort_values(by="Amount", ascending=False).reset_index(drop=True)
                cat_summary["Pct"] = (cat_summary["Amount"] / expense_total * 100) if expense_total > 0 else 0.0

                if cat_summary.empty:
                    st.info("No expense records to display in category breakdown for this filter.")
                else:
                    # Map each category to a distinct color
                    cat_color_map = {row["Category"]: classic_palette[i % len(classic_palette)] for i, row in cat_summary.iterrows()}

                    # --- TOP 1-TAP CATEGORY PILLS SELECTOR (BREAKDOWN + FILTER) ---
                    active_cat = st.session_state.get("drilldown_category")
                    if active_cat and active_cat not in cat_summary["Category"].values:
                        active_cat = None
                        st.session_state["drilldown_category"] = None

                    cat_options = ["🌟 All"] + list(cat_summary["Category"].unique())
                    target_pill = active_cat if active_cat else "🌟 All"

                    # Label formatter map: stable option keys, rich display text
                    pill_display_map = {"🌟 All": "🌟 All"}
                    for _, row in cat_summary.iterrows():
                        pill_display_map[row["Category"]] = f"{row['Category']} • ₪{row['Amount']:,.0f} ({row['Pct']:.0f}%)"

                    def format_pill_label(opt):
                        return pill_display_map.get(opt, str(opt))

                    pill_key = "cat_pills_mobile"
                    # Prevent StreamlitValueAssignmentError if options changed between cycles
                    if pill_key in st.session_state and st.session_state[pill_key] not in cat_options:
                        st.session_state[pill_key] = target_pill

                    st.markdown("<div style='margin-top:0.3rem; margin-bottom: 0.5rem;'><span style='font-size:0.88rem; font-weight:600; color:#94a3b8;'>🏷️ Spending Breakdown & Filter (Tap to isolate):</span></div>", unsafe_allow_html=True)

                    selected_pill = st.pills(
                        "Category Filter",
                        options=cat_options,
                        default=target_pill,
                        format_func=format_pill_label,
                        key=pill_key,
                        label_visibility="collapsed"
                    )

                    # Detect pill change immediately
                    if selected_pill and selected_pill != target_pill:
                        st.session_state["drilldown_category"] = None if selected_pill == "🌟 All" else selected_pill
                        if "pie_chart_selection" in st.session_state:
                            del st.session_state["pie_chart_selection"]
                        st.rerun()

                    # --- CENTERED DONUT CHART ---
                    st.markdown("<div class='chart-box' style='padding: 0.8rem 0.5rem 0.5rem 0.5rem;'>", unsafe_allow_html=True)
                    st.markdown(f"<h4 style='margin-top:0; margin-bottom:0.2rem; font-size:1.05rem; color:{chart_title_color}; text-align:center;'>🍩 Expense Distribution</h4>", unsafe_allow_html=True)

                    fig_donut = px.pie(
                        cat_summary,
                        values="Amount",
                        names="Category",
                        hole=0.62,
                        color="Category",
                        color_discrete_map=cat_color_map
                    )
                    fig_donut.update_traces(
                        textposition='inside',
                        textinfo='percent',
                        insidetextfont=dict(color="#ffffff", family="Inter, -apple-system, sans-serif", size=11),
                        marker=dict(line=dict(color=donut_border, width=2)),
                        hovertemplate="<b>%{label}</b><br>Amount: ₪%{value:,.2f}<br>Share: %{percent}<extra></extra>"
                    )
                    fig_donut.update_layout(
                        template="plotly_dark" if is_dark else "plotly_white",
                        showlegend=False,
                        height=270,
                        margin=dict(t=10, b=10, l=10, r=10),
                        paper_bgcolor="rgba(0,0,0,0)",
                        plot_bgcolor="rgba(0,0,0,0)",
                        font=dict(color=chart_font_color, family="Inter, -apple-system, sans-serif", size=12),
                        hoverlabel=dict(
                            bgcolor="#1c1c28" if is_dark else "#ffffff",
                            font_color="#f3f4f6" if is_dark else "#0f172a",
                            font_size=12,
                            font_family="Inter, -apple-system, sans-serif",
                            bordercolor=donut_border
                        ),
                        annotations=[
                            dict(
                                text=f"<span style='font-size:10px; font-weight:600; letter-spacing:0.04em; color:{chart_font_color};'>TOTAL EXPENSE</span><br><b style='font-size:18px; color:{chart_title_color};'>₪{expense_total:,.0f}</b>",
                                x=0.5, y=0.5,
                                font_size=13,
                                font_family="Inter, -apple-system, sans-serif",
                                showarrow=False
                            )
                        ]
                    )

                    pie_event = st.plotly_chart(
                        fig_donut,
                        use_container_width=True,
                        on_select="rerun",
                        selection_mode="points",
                        key="pie_chart_selection"
                    )
                    if pie_event and isinstance(pie_event, dict):
                        pts = pie_event.get("selection", {}).get("points", [])
                        if pts:
                            p0 = pts[0]
                            clicked_cat = p0.get("label") or p0.get("customdata")
                            if not clicked_cat and "point_number" in p0 and p0["point_number"] < len(cat_summary):
                                clicked_cat = cat_summary.iloc[p0["point_number"]]["Category"]
                            if clicked_cat and clicked_cat in cat_summary["Category"].values and clicked_cat != active_cat:
                                st.session_state["drilldown_category"] = clicked_cat
                                if pill_key in st.session_state:
                                    st.session_state[pill_key] = clicked_cat
                                st.rerun()

                    st.markdown("</div>", unsafe_allow_html=True)

                    # --- ISOLATED CATEGORY DRILLDOWN TABLE & CARDS ---
                    req_cols = ["Date", "Business Name", "Category", "Amount", "Type", "Notes"]

                    if active_cat:
                        cat_detail_rows = expense_sub_df[expense_sub_df["Category"] == active_cat].copy()
                        for c in req_cols:
                            if c not in cat_detail_rows.columns:
                                cat_detail_rows[c] = ""
                        cat_detail_rows["Amount"] = pd.to_numeric(cat_detail_rows["Amount"], errors="coerce").fillna(0.0)
                        cat_detail_rows["Notes"] = cat_detail_rows["Notes"].fillna("").astype(str)
                        cat_detail_rows["Business Name"] = cat_detail_rows["Business Name"].fillna("Unspecified").astype(str)

                        cat_total_amt = cat_detail_rows["Amount"].sum()
                        cat_share_pct = (cat_total_amt / expense_total * 100) if expense_total > 0 else 0.0

                        if "Date_Clean" in cat_detail_rows.columns:
                            cat_detail_rows = cat_detail_rows.sort_values(by="Date_Clean", ascending=False)

                        cat_badge_color = cat_color_map.get(active_cat, "#3b82f6")

                        st.markdown(f"""
                            <div class="category-drilldown-box" style="border-left-color: {cat_badge_color};">
                                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.8rem; flex-wrap: wrap; gap: 0.6rem;">
                                    <div>
                                        <span style="font-size: 1.2rem; font-weight: 700; color: {cat_badge_color};">📂 {active_cat}</span>
                                        <span style="font-size: 0.88rem; color: #94a3b8; margin-left: 0.6rem;">• {len(cat_detail_rows)} transactions ({cat_share_pct:.1f}% of expenses)</span>
                                    </div>
                                    <div style="font-size: 1.35rem; font-weight: 700; font-family: 'JetBrains Mono', monospace; color: #f43f5e;">
                                        {format_currency(cat_total_amt)}
                                    </div>
                                </div>
                            </div>
                        """, unsafe_allow_html=True)

                        col_clr1, col_clr2 = st.columns([1.5, 3.5])
                        with col_clr1:
                            if st.button("✖ Show All Categories", key="btn_reset_cat_isolation", type="primary", use_container_width=True):
                                st.session_state["drilldown_category"] = None
                                if pill_key in st.session_state:
                                    st.session_state[pill_key] = "🌟 All"
                                if "pie_chart_selection" in st.session_state:
                                    del st.session_state["pie_chart_selection"]
                                st.rerun()

                        # Display complete transaction data via Mobile Cards and Full Data Table
                        tab_tx_cards, tab_tx_table = st.tabs(["📱 Transaction Feed (Mobile)", "📋 Full Data Table"])

                        with tab_tx_cards:
                            for _, tx_row in cat_detail_rows.iterrows():
                                tx_bname = tx_row.get("Business Name") or "Unspecified Merchant"
                                tx_amt = tx_row.get("Amount", 0.0)
                                tx_date = str(tx_row.get("Date", ""))
                                tx_notes = str(tx_row.get("Notes", "")).strip()
                                tx_type = str(tx_row.get("Type", "Expense"))
                                tx_color = "#10b981" if tx_type == "Income" else ("#f43f5e" if is_dark else "#dc2626")

                                notes_html = f"<div style='font-size:0.8rem; color:{chart_font_color}; margin-top:3px;'>📝 {tx_notes}</div>" if tx_notes else ""

                                card_item = f"""
                                <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.08); border-radius: 10px; padding: 0.75rem 0.95rem; margin-bottom: 0.5rem;">
                                    <div style="display:flex; justify-content:space-between; align-items:center;">
                                        <span style="font-weight:600; font-size:0.95rem; color:{chart_title_color};">{tx_bname}</span>
                                        <span style="font-family:'JetBrains Mono',monospace; font-weight:700; font-size:1.02rem; color:{tx_color};">₪{tx_amt:,.2f}</span>
                                    </div>
                                    <div style="display:flex; justify-content:space-between; align-items:center; margin-top:0.35rem; font-size:0.8rem; color:{chart_font_color};">
                                        <span>📅 {tx_date} • <span style="background:rgba(59,130,246,0.15); color:#60a5fa; padding:1px 6px; border-radius:4px; font-size:0.75rem;">{active_cat}</span></span>
                                        <span style="font-size:0.75rem;">{tx_type}</span>
                                    </div>
                                    {notes_html}
                                </div>
                                """
                                st.markdown(card_item, unsafe_allow_html=True)

                        with tab_tx_table:
                            st.dataframe(
                                cat_detail_rows[req_cols],
                                use_container_width=True,
                                column_config={
                                    "Date": st.column_config.TextColumn("Date", width="small"),
                                    "Business Name": st.column_config.TextColumn("Expense Description", width="medium"),
                                    "Category": st.column_config.TextColumn("Category", width="small"),
                                    "Amount": st.column_config.NumberColumn("Amount (₪)", format="₪%.2f", width="small"),
                                    "Type": st.column_config.TextColumn("Type", width="small"),
                                    "Notes": st.column_config.TextColumn("Notes / Details", width="large")
                                },
                                hide_index=True
                            )
                    else:
                        with st.expander("📋 View All Current Cycle Transactions Table", expanded=False):
                            tab_all_cards, tab_all_table = st.tabs(["📱 Transaction Feed (Mobile)", "📋 Full Data Table"])
                            with tab_all_cards:
                                for _, tx_row in filtered_df.iterrows():
                                    tx_bname = tx_row.get("Business Name") or "Unspecified Merchant"
                                    tx_amt = tx_row.get("Amount", 0.0)
                                    tx_date = str(tx_row.get("Date", ""))
                                    tx_cat = tx_row.get("Category", "General")
                                    tx_notes = str(tx_row.get("Notes", "")).strip()
                                    tx_type = str(tx_row.get("Type", "Expense"))
                                    tx_color = "#10b981" if tx_type == "Income" else ("#f43f5e" if is_dark else "#dc2626")

                                    notes_html = f"<div style='font-size:0.8rem; color:{chart_font_color}; margin-top:3px;'>📝 {tx_notes}</div>" if tx_notes else ""

                                    card_item = f"""
                                    <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.08); border-radius: 10px; padding: 0.75rem 0.95rem; margin-bottom: 0.5rem;">
                                        <div style="display:flex; justify-content:space-between; align-items:center;">
                                            <span style="font-weight:600; font-size:0.95rem; color:{chart_title_color};">{tx_bname}</span>
                                            <span style="font-family:'JetBrains Mono',monospace; font-weight:700; font-size:1.02rem; color:{tx_color};">₪{tx_amt:,.2f}</span>
                                        </div>
                                        <div style="display:flex; justify-content:space-between; align-items:center; margin-top:0.35rem; font-size:0.8rem; color:{chart_font_color};">
                                            <span>📅 {tx_date} • <span style="background:rgba(59,130,246,0.15); color:#60a5fa; padding:1px 6px; border-radius:4px; font-size:0.75rem;">{tx_cat}</span></span>
                                            <span style="font-size:0.75rem;">{tx_type}</span>
                                        </div>
                                        {notes_html}
                                    </div>
                                    """
                                    st.markdown(card_item, unsafe_allow_html=True)
                            with tab_all_table:
                                for c in req_cols:
                                    if c not in filtered_df.columns:
                                        filtered_df[c] = ""
                                filtered_df["Amount"] = pd.to_numeric(filtered_df["Amount"], errors="coerce").fillna(0.0)
                                filtered_df["Notes"] = filtered_df["Notes"].fillna("").astype(str)
                                filtered_df["Business Name"] = filtered_df["Business Name"].fillna("Unspecified").astype(str)

                                st.dataframe(
                                    filtered_df[req_cols],
                                    use_container_width=True,
                                    column_config={
                                        "Date": st.column_config.TextColumn("Date", width="small"),
                                        "Business Name": st.column_config.TextColumn("Expense Description", width="medium"),
                                        "Category": st.column_config.TextColumn("Category", width="small"),
                                        "Amount": st.column_config.NumberColumn("Amount (₪)", format="₪%.2f", width="small"),
                                        "Type": st.column_config.TextColumn("Type", width="small"),
                                        "Notes": st.column_config.TextColumn("Notes / Details", width="large")
                                    },
                                    hide_index=True
                                )

    # Footer note
    st.markdown("""
        <div style="text-align: center; color: #475569; font-size: 0.8rem; margin-top: 2rem; border-top: 1px solid #1e1e2d; padding-top: 1rem;">
            ⚡ Personal Finance Dashboard • Connected to Google Sheet <b>Financial_DB</b> (MasterData) • Click <b>Transaction Center</b> to inspect tables, enter transactions, or sync statements.
        </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
