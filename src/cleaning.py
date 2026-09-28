"""
cleaning.py
------------
Cleaning utilities for the messy e-commerce sales dataset.

Each function targets one column and handles the specific inconsistencies
found during the initial data audit (see notebooks/analysis.ipynb, Section 1).
"""

import datetime
import numpy as np
import pandas as pd


CITY_MAPPING = {
    "cairo": "Cairo", "القاهرة": "Cairo", "القاهره": "Cairo",
    "alexandria": "Alexandria", "الاسكندرية": "Alexandria",
    "الإسكندرية": "Alexandria", "alex": "Alexandria",
    "giza": "Giza", "الجيزة": "Giza", "الجيزه": "Giza",
    "sohag": "Sohag", "سوهاج": "Sohag",
    "mansoura": "Mansoura", "المنصورة": "Mansoura", "المنصوره": "Mansoura",
    "assiut": "Assiut", "asyut": "Assiut", "أسيوط": "Assiut", "اسيوط": "Assiut",
    "tanta": "Tanta", "طنطا": "Tanta",
}

PRODUCT_MAPPING = {
    "T Shirt": "T-Shirt",
    "Tshirt": "T-Shirt",
}

# Formats that can only be read one way. The slash format is day-first
# (D/M/Y) in this export, as in the rest of the file.
DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%d-%b-%Y"]

# Dotted dates are the ambiguous ones: the export writes them as D.M.Y for
# some rows and M.D.Y for others (e.g. "04.07.2026" vs "08.28.2026").
# Order matters: when a value is genuinely ambiguous, the first format wins,
# so day-first is the fallback (it matches every other format in the file).
DOTTED_FORMATS = ["%d.%m.%Y", "%m.%d.%Y"]


def clean_customer_name(value):
    """Strip whitespace and standardize to Title Case."""
    if pd.isna(value):
        return value
    return str(value).strip().title()


def clean_city(value):
    """
    Map every raw city spelling (Arabic/English, mixed case, abbreviations)
    to one standard English name. Unmapped values (including the '-'
    placeholder) become NaN rather than being guessed.
    """
    if pd.isna(value):
        return value
    cleaned = str(value).strip().lower()
    return CITY_MAPPING.get(cleaned, np.nan)


def date_candidates(value):
    """
    Return every valid reading of a raw date value, as a list of datetimes.

    - Already a datetime  -> [value]
    - Unambiguous string  -> one reading (ISO, D/M/Y, D-Mon-Y)
    - Dotted string       -> one or two readings, depending on whether the
      two numbers can be read as both day and month (e.g. "08.28.2026" has
      only one valid reading, "04.07.2026" has two)
    - Missing/unreadable  -> []
    """
    if isinstance(value, datetime.datetime):
        return [value]
    if pd.isna(value):
        return []
    value = str(value).strip()

    for fmt in DATE_FORMATS:
        try:
            return [datetime.datetime.strptime(value, fmt)]
        except ValueError:
            continue

    readings = []
    for fmt in DOTTED_FORMATS:
        try:
            parsed = datetime.datetime.strptime(value, fmt)
        except ValueError:
            continue
        if parsed not in readings:
            readings.append(parsed)
    return readings


def clean_date(value, window=None):
    """
    Parse a date that may already be a datetime object, or a string in one
    of several formats. Returns pd.NaT if no known format matches.

    `window` is an optional (earliest, latest) pair used to settle ambiguous
    dotted dates: the reading that falls inside the window wins. It is built
    from the unambiguous rows of the same file (see load_and_clean), so a
    date like "04.07.2026" is read as 4 July (inside the Jun-Aug data) and
    not 7 April (outside it).

    If both readings are still plausible, the day-first reading is used,
    consistent with the other formats in this export.
    """
    readings = date_candidates(value)
    if not readings:
        return pd.NaT
    if len(readings) == 1 or window is None:
        return readings[0]

    lo, hi = window
    inside = [r for r in readings if lo <= r <= hi]
    return inside[0] if inside else readings[0]


def clean_price(value):
    """
    Extract a numeric price from values that may already be numeric, or may
    be strings with currency labels ('EGP', 'جنيه', '$') and thousands
    separators. The '-' placeholder correctly becomes NaN.
    """
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        value = value.strip().replace(",", "").replace("$", "")
        value = value.replace("EGP", "").replace("جنيه", "").strip()
        try:
            return float(value)
        except ValueError:
            return np.nan
    return np.nan


def clean_phone(value):
    """
    Normalize Egyptian phone numbers to the local 11-digit format
    (e.g. 01012345678), handling the +20 country code, stripped leading
    zeros (numbers that were read as int/float), and the '-' placeholder.
    """
    if pd.isna(value):
        return pd.NA
    value = str(value).strip().replace(" ", "").replace("-", "")
    value = value.replace("+", "")
    if value in ("", "nan"):
        return pd.NA
    if value.startswith("20") and len(value) == 12:
        value = "0" + value[2:]
    elif len(value) == 10:
        value = "0" + value
    return value


def clean_product(value):
    """Standardize casing/spacing and unify T-Shirt spelling variants."""
    if pd.isna(value):
        return value
    cleaned = str(value).strip().title()
    return PRODUCT_MAPPING.get(cleaned, cleaned)


def clean_payment_method(value):
    """
    Standardize case and unify abbreviations (COD, Card). Uses exact
    matching rather than substring replace, since "Card" is a substring of
    the already-correct "Credit Card" and a naive .replace() would corrupt it
    into "Credit Credit Card".
    """
    if pd.isna(value):
        return value
    cleaned = str(value).strip().title()
    if cleaned == "Cod":
        return "Cash On Delivery"
    if cleaned == "Card":
        return "Credit Card"
    return cleaned


def clean_qty(value):
    """
    Convert quantity to numeric, handling the word 'two', the '-'
    placeholder, and stray whitespace.
    """
    if pd.isna(value):
        return np.nan
    if isinstance(value, (int, float)):
        return value
    value = str(value).strip().lower()
    if value == "two":
        return 2
    if value in ("-", ""):
        return np.nan
    try:
        return float(value)
    except ValueError:
        return np.nan


def clean_status(value):
    """Standardize order status casing."""
    if pd.isna(value):
        return value
    return str(value).strip().title()


def load_and_clean(filepath, sheet_name="Sales"):
    """
    Load the raw Excel file and apply the full cleaning pipeline.
    Duplicate rows are dropped only AFTER cleaning, so that rows which
    are duplicates once formatting noise is removed are also caught.

    Returns the cleaned DataFrame with an added 'Revenue' column
    (Qty * Price).
    """
    df = pd.read_excel(filepath, sheet_name=sheet_name)

    df["Customer Name"] = df["Customer Name"].apply(clean_customer_name)
    df["City"] = df["City"].apply(clean_city)

    # Dates: first collect every valid reading of each value. Rows with only
    # one reading are certain; they define the date range of the file, which
    # is then used to settle the ambiguous dotted dates.
    readings = df["Order Date"].apply(date_candidates)
    certain = [r[0] for r in readings if len(r) == 1]
    window = (min(certain), max(certain))
    df["Order Date"] = df["Order Date"].apply(lambda v: clean_date(v, window))

    df["Price"] = df["Price"].apply(clean_price)
    df["Phone"] = df["Phone"].apply(clean_phone)
    df["Product"] = df["Product"].apply(clean_product)
    df["Payment Method"] = df["Payment Method"].apply(clean_payment_method)
    df["Qty"] = df["Qty"].apply(clean_qty)
    df["Status"] = df["Status"].apply(clean_status)

    # Missing-value decisions (documented in README):
    # - Phone / Customer Name: not used in any numeric calculation, so a
    #   readable placeholder is safe and more useful than NaN in a report.
    # - City / Qty / Price: left as NaN on purpose. They feed numeric
    #   aggregations (sums, groupbys) where a placeholder would silently
    #   corrupt the math, so pandas' built-in NaN-skipping is the safer choice.
    df["Phone"] = df["Phone"].fillna("Not provided")
    df["Customer Name"] = df["Customer Name"].fillna("Not Found")

    df = df.drop_duplicates()

    df["Month"] = df["Order Date"].dt.to_period("M")
    df["Revenue"] = df["Qty"] * df["Price"]

    return df
