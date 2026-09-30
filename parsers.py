"""
Data Parsing & Categorization Module.
Parses:
  1. Max Credit Card Excel (.xlsx) files via dynamic header row detection
  2. Isracard Statements PDF (.pdf) via strict regex line extraction & disclosure filtering
Applies Hebrew string sanitization, noise stripping, and bidi directionality.
"""

import re
import io
import pandas as pd
from datetime import datetime
import pdfplumber
from bidi.algorithm import get_display

# Keyword category mapping dictionary
CATEGORY_KEYWORDS = {
    "Shopping": [
        "בהצדעה", "כרטיס נטען בהצדעה", "פי הקריון", "קריון", "piitel", "paypal piitel",
        "ביליבונג", "billabong", "זארה", "zara", "castro", "קסטרו", "h&m", "pull&bear", "bershka",
        "terminal x", "טרמינל", "asos", "amazon", "אמזון", "aliexpress",
        "shein", "עלי אקספרס", "nike", "adidas", "איקאה", "ikea", "ksp",
        "אייבורי", "ivory", "באג", "bug", "ace", "הום סנטר", "מגנוליה",
        "פול אנד בר", "חומרי בניי", "שזר", "אמריקן איגל", "american eagle",
        "נונה טכניון", "נונה", "טכניון", "ספרים", "סטימצקי", "צומת ספרים",
        "גן", "צהרון", "מעון", "בית ספר", "חוג", "אוניברסיט", "מכללה",
        "בית הסטודנט", "סטודנט", "דלתא", "פוקס", "fox", "רנואר", "renuar"
    ],
    "Restaurants": [
        "קפה", "domo", "ארומה", "קפית", "גרג", "לנדוור", "rebar", "רולדין",
        "מסעד", "בר", "פיצה", "בורגר", "שווארמה", "גולדה", "golda", "בייקרי",
        "ביסטרו", "קפה קפה", "ארקפה", "arcaffe", "starbucks", "mcdonalds",
        "מקדולנס", "wolt", "תן ביס", "10bis", "tabit", "ontopo", "פאב",
        "חומוס", "מזנון", "סילבה", "פלפלת", "אספרסו", "mosh beach", "נולה סוקס",
        "ברדיצ'ב", "panda wok", "דומינוס", "ג'פניקה", "japanika", "מוזס", "bbb",
        "אגאדיר", "סושי", "נודלס", "בלאק", "מאפיית", "לחם", "קונדיטוריה"
    ],
    "Entertainment": [
        "מכבי חיפה (איצטדיון)", "מכבי חיפה", "איצטדיון", "סמי עופר",
        "netflix", "spotify", "apple.com", "apple", "google", "youtube", "playstation",
        "steam", "cinema", "קולנוע", "סינמה", "יס פלאנט", "yes planet",
        "רב חן", "הוט סינמה", "תיאטרון", "הופעה", "זאפה", "zappa",
        "מלביר", "אייר חיפה", "המבשלה", "חבר שלי", "כרטיסים", "אירוע", "הצגה"
    ],
    "Investment Fund": [
        "קרן השתלמות", "קופת גמל", "קרן נאמנות", "אינטראקטיב", "מיטב",
        "אלטשולר", "הפניקס", "מגדל", "הראל", "פסגות", "אנליסט", "מור",
        "ibkr", "interactive brokers", "etoro", "investment", "השקע",
        "fidelity", "vanguard", "schwab", "fund", "השקעות", "ניירות ערך"
    ],
    "Transportation": [
        "yellow", "דלק", "סונול", "פז", "דור אלון", "מיקה", "ten",
        "רכבת", "אגד", "דן", "מטרופולין", "moovit", "רב קו", "rav kav",
        "gett", "yango", "uber", "חניון", "אחוזות החוף", "פנגו", "pango",
        "סלופארק", "cellopark", "דרך ארץ", "כביש 6", "lime payoff tda7", "lime payoff", "lime"
    ],
    "Groceries": [
        "שופרסל", "רמי לוי", "יוחננוף", "אושר עד", "קרפור", "carrefour",
        "ויקטורי", "מגה", "טיב טעם", "am:pm", "סופר יודה", "מינימרקט",
        "מכולת", "סופרמרקט", "ירקות", "פירות", "קצב", "דגים", "מאפיה",
        "טבע קסטל", "ניצת הדובדבן", "מעדניית"
    ],
    "Health & Pharmacy": [
        "סופר פארם", "סופר-פארם", "super-pharm", "be פארם", "ניו פארם",
        "כללית", "מכבי שירותי בריאות", "קופת חולים מכבי", "מאוחדת", "לאומית",
        "בית מרקחת", "אופטיקה", "רופא", "מרפאה", "דנטל", "שיניים", "שרותי בריאות"
    ],
    "Utilities & Bills": [
        "חברת החשמל", "חשמל", "ארנונה", "עיריית", "מים", "מי אביבים",
        "הוט", "hot", "יס", "yes", "בזק", "bezeq", "פרטנר", "partner",
        "סלקום", "cellcom", "פלאפון", "pelephone", "012", "019", "גז",
        "היפרטרוניקס", "ביטוח", "מ.התחבורה", "שדה תעופה", "נתבג", "נתב\"ג",
        "שלומי בן שיטרית", "שיטרית", "שטרית", "עסקת תשלומים (שיטרית)"
    ]
}

# Footer metadata, legal terms, and credit disclosures to strictly ignore
DISCLOSURE_NOISE_PATTERNS = [
    r"סה[\"״\']?כ",
    r"חיוב\s*לתאריך",
    r"הנחה",
    r"בניכוי",
    r"ריבית",
    r"דמי\s*כרטיס",
    r"מסגרת",
    r"קרדיט",
    r"%",
    r"לתשומת\s*לב",
    r"הודעות",
    r"שטרם\s*נפרעה",
    r"בנק\s*ישראל",
    r"חלקך\s*בחשבון",
    r"תנאי\s*האשראי",
    r"פירוט\s*עסקאות",
    r"סך\s*הכל",
    r"מועד\s*חיוב",
    r"תאריך\s*חיוב",
    r"תקנון",
    r"עמוד\s*\d+\s*מתוך",
    r"דף\s*\d+\s*מתוך"
]

SECTOR_NOISE = [
    "מסעדות/בתי קפה", "מסעדות/קפה", "בתי קפה", "בית קפה", "מסעדות", "שונות", 
    "דלק", "הלבשה", "פארמה", "מחשבים", "מוצרי חשמל", "שרות רפואי", "שירותי רכב", 
    "ביטוח", "תש' רשויות", "מכולת/סופר", "סופר/מכולת", "מכולת סופר", "מכולת", 
    "סופרמרקט", "תרבות", "קניה אינט'", "בניה/שיפוץ", "נופש ותיור"
]

OPERATIONAL_NOISE = [
    r"תש\.?\s*נייד",
    r"לא\s*הוצג[ה]?",
    r"הוצג\s+",
    r"ה\.?\s*קבע",
    r"הוראת\s*קבע",
    r"בBIT\s+",
    r"בע[\"״\']?מ",
    r"גרנד\s*קניון",
    r"קניון",
    r"סניף\s*\S*",
    r"[\*\#\₪\$\:\_\|\-\/]+",
    r"תשלום\s*\d+\s*מתוך\s*\d+",
    r"\d+\s*מתוך\s*\d+"
]


def clean_and_shorten_business_name(raw_name: str, apply_bidi: bool = False) -> str:
    """
    Cleans and shortens business names:
    - Strips operational noise prefixes & suffixes:
      'תש.נייד', 'לא הוצג', 'ה.קבע', 'בע"מ', 'סניף', 'גרנד קניון', 'קניון', sector names
    - Prevents accidental sector words like 'קפה' or 'מכולת סופר' from adhering to merchant names.
    - Fixes separated Latin characters (e.g. 'P AYPAL' -> 'PAYPAL', 'L IME' -> 'LIME')
    - Normalizes messy strings into 2–3 core words (e.g. 'תש.נייד מגנוליה קניון אילת מ' -> 'מגנוליה אילת')
    - Applies bidi.algorithm.get_display() ONLY when apply_bidi=True (used specifically for visual-RTL PDF text).
    """
    if not raw_name or not isinstance(raw_name, str):
        return ""

    s = raw_name.strip()

    # Operational noise removal
    for pat in OPERATIONAL_NOISE:
        s = re.sub(pat, " ", s, flags=re.IGNORECASE)

    # Sector noise removal (both normal and reversed)
    for sec in SECTOR_NOISE:
        s = s.replace(sec, " ")
        s = s.replace(sec[::-1], " ")

    # Strip accidental trailing sector tags
    genuine_cafes = [
        "קפה סילבה", "קפה טכניון", "קפה משחקים", "קפה ברנצו", 
        "קפה קפה", "קפה ארומה", "קפה לנדוור", "קפה גרג", "מילהאוס קפה"
    ]
    if s.endswith(" קפה"):
        prefix = s[:-4].strip()
        if len(prefix) >= 3 and not any(gc in s for gc in genuine_cafes):
            s = prefix
        elif any(gc in prefix for gc in genuine_cafes):
            s = prefix

    if s.endswith(" מכולת סופר"):
        s = s[:-11].strip()
    elif s.endswith(" מכולת"):
        s = s[:-6].strip()

    # Remove standalone numbers (card index, digits)
    s = re.sub(r"\b\d+\b", " ", s)

    # Reconnect separated Latin characters like 'P AYPAL' -> 'PAYPAL', 'L IME' -> 'LIME'
    s = re.sub(r"\b([A-Za-z])\s+([A-Za-z]{2,})\b", r"\1\2", s)

    # Split into words and eliminate single Hebrew noise characters (e.g. trailing 'מ', 'א')
    words = [
        w for w in s.split()
        if len(w) > 1 or (w.isalnum() and len(w) == 1 and not ('\u0590' <= w <= '\u05ea'))
    ]

    if not words:
        words = s.split()

    # Core description extraction: keep 2-3 core words
    core_words = words[:3] if len(words) >= 3 else words
    cleaned = " ".join(core_words).strip()

    if not cleaned:
        cleaned = raw_name.strip()

    # Known business names normalization
    fixes = {
        "KOW ADNAP": "PANDA WOK",
        "WOLT": "WOLT",
        "פז YELLOW אפליקצית": "פז YELLOW",
        "YELLOW פז אפליקצית": "פז YELLOW",
        "איגל קריון גמ": "אמריקן איגל",
        "איצטדיון סמי": "איצטדיון סמי עופר",
        "אייבורי": "אייבורי מחשבים",
        "סעוד הראל כללית": "הראל סיעוד כללית",
        "שמשון מרקט": "סופר שמשון מרקט",
        "שוק אורנים": "סופר שוק אורנים",
        "ורדיה": "סופר ורדיה"
    }
    if cleaned in fixes:
        cleaned = fixes[cleaned]

    # Apply Hebrew bidi rendering ONLY if explicitly requested (e.g. for PDF visual text)
    if apply_bidi and any("\u0590" <= c <= "\u05ea" for c in cleaned):
        return get_display(cleaned)

    return cleaned


def categorize_transaction(business_name: str) -> str:
    """
    Assigns category based on smart keyword matching.
    Supports regular and visual-RTL reversed Hebrew text.
    Explicitly prioritizes:
      - 'שלומי בן שיטרית' -> Utilities & Bills
      - 'פי הקריון בע' -> Shopping
      - 'PAYPAL PIITEL' -> Shopping
      - 'בהצדעה' -> Shopping
      - 'מכבי חיפה (איצטדיון)' / 'איצטדיון' -> Entertainment
      - 'שדה תעופה' -> Utilities & Bills
      - 'ביליבונג' -> Shopping
      - 'LIME PAYOFF TDA7' / 'lime' -> Transportation
      - 'domo' / 'קפה' / 'הפק' -> Restaurants
    Defaults to 'Shopping' (eliminating Uncategorized).
    """
    if not business_name or pd.isna(business_name):
        return "Shopping"

    text = str(business_name).lower().strip()

    # Explicit high-priority specific rules
    if "מכבי חיפה" in text or "איצטדיון" in text or "ןוידטציא" in text or "סמי עופר" in text:
        return "Entertainment"
    if "שדה תעופה" in text or "הפועת הדש" in text or "נתבג" in text or "נתב\"ג" in text:
        return "Utilities & Bills"
    if "שיטרית" in text or "שטרית" in text or "שלומי" in text:
        return "Utilities & Bills"
    if "קריון" in text or "פי הקריון" in text or "ןוירק" in text:
        return "Shopping"
    if "piitel" in text or "פייטל" in text:
        return "Shopping"
    if "בהצדעה" in text or "העדצהב" in text:
        return "Shopping"
    if "ביליבונג" in text or "גנוביליב" in text or "billabong" in text:
        return "Shopping"
    if "lime" in text or "לייל" in text or "ליים" in text:
        return "Transportation"
    if "domo" in text or "קפה" in text or "הפק" in text or "rebar" in text or "רי באר" in text:
        return "Restaurants"

    # Normalize punctuation into spaces for clean tokenization
    normalized_text = re.sub(r"[^\w\s]", " ", text)
    tokens = normalized_text.split()

    for category, keywords in CATEGORY_KEYWORDS.items():
        for kw in keywords:
            kw_clean = kw.lower().strip()
            kw_reversed = kw_clean[::-1] if any("\u0590" <= c <= "\u05ea" for c in kw_clean) else None

            # For short keywords (<= 3 chars, like 'בר', 'דן', 'פז', 'גן', 'ten'), require whole word token match
            if len(kw_clean) <= 3:
                hebrew_prefixes = ["", "ב", "ל", "מ", "ה", "כ", "ו"]
                possible_variants = {prefix + kw_clean for prefix in hebrew_prefixes}
                if kw_reversed:
                    possible_variants.update({prefix + kw_reversed for prefix in hebrew_prefixes})
                if any(tok in possible_variants for tok in tokens):
                    return category
            else:
                # Longer keywords can match as substring or token match
                if kw_clean in text or any(kw_clean in tok for tok in tokens):
                    return category
                if kw_reversed and (kw_reversed in text or any(kw_reversed in tok for tok in tokens)):
                    return category

    return "Shopping"


def parse_max_excel(file_content) -> pd.DataFrame:
    """
    Parses Max Credit Card Excel file.
    1. Reads the Excel file initially with header=None (no headers).
    2. Iterates through rows to locate the index containing 'תאריך עסקה'.
    3. Re-parses the file setting header=header_row_index so columns
       'תאריך עסקה', 'שם בית העסק', and 'סכום חיוב' are resolved correctly.
    4. Cleans and shortens business names and applies bidi directionality.
    """
    if hasattr(file_content, "seek"):
        file_content.seek(0)
    
    # 1. Read initially with no headers
    raw_df = pd.read_excel(file_content, header=None)

    # 2. Iterate through rows to locate index containing 'תאריך עסקה'
    header_row_index = None
    for idx, row in raw_df.iterrows():
        row_str_vals = [str(val).strip() for val in row.values if pd.notna(val)]
        if any("תאריך עסקה" in val for val in row_str_vals):
            header_row_index = idx
            break

    if header_row_index is None:
        # Fallback inspection for variations like 'תאריך רכישה'
        for idx, row in raw_df.iterrows():
            row_str_vals = [str(val).strip() for val in row.values if pd.notna(val)]
            if any("תאריך" in val for val in row_str_vals) and any("סכום" in val for val in row_str_vals):
                header_row_index = idx
                break

    if header_row_index is None:
        raise ValueError("Could not locate header row containing 'תאריך עסקה' in Excel file.")

    # 3. Re-parse the file setting header=header_row_index
    if hasattr(file_content, "seek"):
        file_content.seek(0)
    df = pd.read_excel(file_content, header=header_row_index)

    # Resolve required columns
    date_col = None
    name_col = None
    amount_col = None

    for col in df.columns:
        col_str = str(col).strip()
        if "תאריך עסקה" in col_str or "תאריך רכישה" in col_str or col_str == "תאריך":
            if not date_col:
                date_col = col
        elif "שם בית העסק" in col_str or "שם בית עסק" in col_str or "שם עסק" in col_str or "בית עסק" in col_str:
            if not name_col:
                name_col = col
        elif "סכום חיוב" in col_str or "סכום החיוב" in col_str or "סכום עסקה" in col_str or col_str == "סכום":
            if not amount_col:
                amount_col = col

    if not date_col or not name_col or not amount_col:
        raise ValueError(
            f"Could not resolve required columns ('תאריך עסקה', 'שם בית העסק', 'סכום חיוב') from header row index {header_row_index}. "
            f"Found columns: {list(df.columns)}"
        )

    clean_rows = []
    for _, row in df.iterrows():
        raw_date = row[date_col]
        raw_name = row[name_col]
        raw_amount = row[amount_col]

        if pd.isna(raw_date) or pd.isna(raw_amount) or pd.isna(raw_name):
            continue

        parsed_date = pd.to_datetime(raw_date, errors="coerce", dayfirst=True)
        if pd.isna(parsed_date):
            continue

        date_str = parsed_date.strftime("%Y-%m-%d")
        raw_business = str(raw_name).strip()

        # Parse numeric amount
        try:
            if isinstance(raw_amount, (int, float)):
                amt = float(raw_amount)
            else:
                amt_str = str(raw_amount).replace(",", "").replace("₪", "").replace("$", "").strip()
                if amt_str.endswith("-"):
                    amt = -float(amt_str[:-1])
                else:
                    amt = float(amt_str)
        except ValueError:
            continue

        t_type = "Expense"
        if amt < 0:
            amt = abs(amt)
            t_type = "Income"

        # Clean and shorten business name natively in logical Hebrew order WITHOUT bidi / get_display()
        cleaned_business = clean_and_shorten_business_name(raw_business, apply_bidi=False)

        # Categorize after the text is correctly oriented
        category = categorize_transaction(cleaned_business)
        if category == "Uncategorized":
            category = categorize_transaction(raw_business)

        clean_rows.append({
            "Date": date_str,
            "Business Name": cleaned_business,
            "Category": category,
            "Type": t_type,
            "Amount": round(amt, 2),
            "Notes": ""
        })

    return pd.DataFrame(clean_rows, columns=["Date", "Business Name", "Category", "Type", "Amount", "Notes"])


def is_summary_or_disclosure_line(line_str: str) -> bool:
    """Detects if a line is a summary subtotal, header, or legal disclosure line."""
    if not line_str:
        return True

    # Subtotal lines: סה"כ or כ"הס or סהכ or כהס or חיוב לתאריך
    if "סה\"כ" in line_str or "כ\"הס" in line_str or "סהכ" in line_str or "כהס" in line_str:
        return True
    if "חיוב לתאריך" in line_str or ("בויח" in line_str and ("ךיראתל" in line_str or "לתאריך" in line_str)):
        return True
    
    # Credit terms / disclosure patterns
    for noise_pat in DISCLOSURE_NOISE_PATTERNS:
        if re.search(noise_pat, line_str, flags=re.IGNORECASE):
            return True

    return False


def parse_isracard_pdf(file_content) -> pd.DataFrame:
    """
    Parses Isracard Statement PDF using pdfplumber with strict regex and disclosure filtering:
    1. Filter out footer metadata, legal terms, card fees, and credit disclosures.
    2. Enforce strict regex matching date (DD/MM/YY or DD/MM/YYYY) and numerical amount.
    3. Clean & shorten business names (strip 'תש.נייד', 'לא הוצג', 'ה.קבע', etc.).
    4. Apply bidi.algorithm.get_display() to Hebrew strings for correct rendering.
    5. Ensure 'Domo' or 'קפה' categorizes as 'Going Out'.
    """
    clean_rows = []
    
    # Strict regex for transaction date (DD/MM/YY or DD/MM/YYYY)
    date_pattern = re.compile(r"(\d{1,2}/\d{1,2}/\d{2,4})")
    # Strict regex for decimal amount
    amount_pattern = re.compile(r"(-?\d{1,3}(?:,\d{3})*\.\d{2}-?)")

    if isinstance(file_content, bytes):
        pdf_file = io.BytesIO(file_content)
    else:
        pdf_file = file_content

    with pdfplumber.open(pdf_file) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            text = page.extract_text(layout=False) or ""
            lines = text.split("\n")

            for line in lines:
                line_str = line.strip()
                if is_summary_or_disclosure_line(line_str):
                    continue

                # Remove discount annotations (e.g. הנחה 22.91 ₪) so the discount amount is not mistaken for transaction billing amount
                line_for_amt = re.sub(r"(?:הנחה|החנה)\s*[\₪\$]?\s*\d+(?:\.\d+)?|[\₪\$]?\s*\d+(?:\.\d+)?\s*[\₪\$]?\s*(?:הנחה|החנה)", " ", line_str)

                # 2. Strict regex matching for valid transaction lines
                date_matches = list(date_pattern.finditer(line_str))
                amt_matches = list(amount_pattern.finditer(line_for_amt))

                if not date_matches or not amt_matches:
                    continue

                # Extract primary date
                date_raw = date_matches[0].group(1)
                dt = pd.to_datetime(date_raw, errors="coerce", dayfirst=True)
                if pd.isna(dt):
                    continue

                # In Isracard tables with installments (e.g. 342.00 2,394.00 or 2,062.00 4,123.00),
                # the first amount extracted is סכום החיוב (the monthly billing amount).
                amt_raw = amt_matches[0].group(1)
                is_negative = False
                clean_amt_str = amt_raw.replace(",", "").strip()
                if clean_amt_str.endswith("-"):
                    is_negative = True
                    clean_amt_str = clean_amt_str[:-1]
                elif clean_amt_str.startswith("-"):
                    is_negative = True
                    clean_amt_str = clean_amt_str[1:]

                try:
                    amt_val = float(clean_amt_str)
                except ValueError:
                    continue

                # 3. Extract business name by removing matched date and amounts
                remaining = line_str
                for dm in date_matches:
                    remaining = remaining.replace(dm.group(1), " ")
                for am in amt_matches:
                    remaining = remaining.replace(am.group(1), " ")

                # Strip 4-digit card sequences
                remaining = re.sub(r"\b\d{4}\b", " ", remaining)
                raw_extracted_name = " ".join(remaining.split())

                if not raw_extracted_name or len(raw_extracted_name) < 2:
                    continue

                # Check high priority category keywords
                line_lower = line_str.lower()
                combined_context = f"{raw_extracted_name} {line_str}"
                category = categorize_transaction(combined_context)
                if "domo" in line_lower or "קפה" in line_str or "הפק" in line_str:
                    category = "Going Out"

                # 4. Clean & shorten business name into 2-3 core words + apply get_display()
                cleaned_name = clean_and_shorten_business_name(raw_extracted_name, apply_bidi=True)

                t_type = "Income" if is_negative else "Expense"

                clean_rows.append({
                    "Date": dt.strftime("%Y-%m-%d"),
                    "Business Name": cleaned_name,
                    "Category": category,
                    "Type": t_type,
                    "Amount": round(amt_val, 2),
                    "Notes": ""
                })

    df_result = pd.DataFrame(clean_rows, columns=["Date", "Business Name", "Category", "Type", "Amount", "Notes"])
    if not df_result.empty:
        df_result = df_result.drop_duplicates(subset=["Date", "Business Name", "Amount"])

    return df_result


def parse_uploaded_file(uploaded_file) -> pd.DataFrame:
    """Auto-detects file type (.xlsx or .pdf) and returns parsed DataFrame."""
    filename = uploaded_file.name.lower()
    if filename.endswith(".xlsx") or filename.endswith(".xls"):
        return parse_max_excel(uploaded_file)
    elif filename.endswith(".pdf"):
        return parse_isracard_pdf(uploaded_file)
    else:
        raise ValueError(f"Unsupported file format: {uploaded_file.name}. Please upload .xlsx or .pdf.")
