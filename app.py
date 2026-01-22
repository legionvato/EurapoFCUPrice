import re
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

# ----------------------------
# Config
# ----------------------------
PRICE_LIST_PATH = Path("data/Price LIST FCU Eurapo 4pipe AC.xlsx")

# Input file expected columns:
INPUT_MODEL_COL = "model"
INPUT_ROWS_COL = "rows"

# Output columns added:
OUTPUT_PRICE_COL = "price_eur"
OUTPUT_STATUS_COL = "match_status"

# Pricelist expected columns (in some sheet inside the file):
PL_MODEL_COL = "model"
PL_ROWS_COL = "rows"
PL_TOTAL_COL = "total_price_eur"


# ----------------------------
# Helpers
# ----------------------------
def norm_model(x: str) -> str:
    """
    Normalize model strings to increase match rate:
    - uppercase
    - collapse whitespace
    - replace '-' with space
    - normalize EBH sizes to 'EBH 050' format
    """
    if x is None:
        return ""
    x = str(x).strip().upper()
    x = x.replace("-", " ")
    x = re.sub(r"\s+", " ", x).strip()

    # Normalize EBH formats:
    # "EBH050" -> "EBH 050"
    # "EBH 50" -> "EBH 050"
    m = re.match(r"^(EBH)\s*0*([0-9]{1,3})$", x)
    if m:
        return f"{m.group(1)} {int(m.group(2)):03d}"

    return x


def norm_rows(v):
    """
    Normalize rows to int, or None if empty.
    """
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    s = str(v).strip()
    if s == "":
        return None
    try:
        return int(float(s))
    except Exception:
        return None


@st.cache_data
def load_pricelist(pricelist_path: Path):
    """
    Loads the built-in pricelist and builds a lookup dict keyed by (model_norm, rows_norm).
    Auto-detects which sheet contains required columns.
    """
    if not pricelist_path.exists():
        raise FileNotFoundError(f"Pricelist file not found: {pricelist_path}")

    xls = pd.ExcelFile(pricelist_path)
    required = {PL_MODEL_COL, PL_ROWS_COL, PL_TOTAL_COL}

    chosen_sheet = None
    for sh in xls.sheet_names:
        preview = pd.read_excel(xls, sheet_name=sh, nrows=5)
        if required.issubset(set(preview.columns)):
            chosen_sheet = sh
            break

    if chosen_sheet is None:
        raise ValueError(
            f"No sheet in '{pricelist_path.name}' contains required columns {required}. "
            f"Sheets found: {xls.sheet_names}"
        )

    df = pd.read_excel(xls, sheet_name=chosen_sheet)
    df["model_norm"] = df[PL_MODEL_COL].apply(norm_model)
    df["rows_norm"] = df[PL_ROWS_COL].apply(norm_rows)

    # Build primary lookup: (model_norm, rows_norm) -> total_price_eur
    lookup = {}
    dupes = []
    for r in df.itertuples(index=False):
        key = (getattr(r, "model_norm"), getattr(r, "rows_norm"))
        val = float(getattr(r, PL_TOTAL_COL))
        if key in lookup and lookup[key] != val:
            dupes.append((key, lookup[key], val))
        lookup[key] = val

    return df, lookup, chosen_sheet, dupes


def price_input_df(input_df: pd.DataFrame, lookup: dict):
    """
    Adds price/status columns to input_df by matching (model, rows).
    """
    df = input_df.copy()

    df["model_norm"] = df[INPUT_MODEL_COL].apply(norm_model)
    df["rows_norm"] = df[INPUT_ROWS_COL].apply(norm_rows)

    prices = []
    status = []

    for m, r in zip(df["model_norm"], df["rows_norm"]):
        p = lookup.get((m, r))
        if p is None:
            prices.append(None)
            status.append("NOT_FOUND")
        else:
            prices.append(p)
            status.append("OK")

    df[OUTPUT_PRICE_COL] = prices
    df[OUTPUT_STATUS_COL] = status

    return df


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "PRICED") -> bytes:
    """
    Convert DataFrame to Excel bytes.
    """
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    bio.seek(0)
    return bio.getvalue()


# ----------------------------
# Streamlit UI
# ----------------------------
st.set_page_config(page_title="Eurapo FCU Price Filler", layout="wide")
st.title("Eurapo FCU Price Filler (4-pipe, AC motor, NO EC)")
st.caption("Upload an Excel with columns: model, rows — the app writes the matched price next to it.")

# Load pricelist
try:
    pl_df, lookup, sheet, dupes = load_pricelist(PRICE_LIST_PATH)
    st.success(f"Loaded pricelist: {PRICE_LIST_PATH.name} | Sheet: {sheet} | Rows: {len(pl_df)}")
    if dupes:
        st.warning(f"Warning: found {len(dupes)} duplicate keys with different prices in pricelist. Using last seen value.")
except Exception as e:
    st.error(str(e))
    st.stop()

with st.expander("Show pricelist columns (debug)"):
    st.write(pl_df.head(10))

uploaded = st.file_uploader("Upload your Excel (XLSX)", type=["xlsx"])

if uploaded:
    # Read input
    try:
        in_df = pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"Could not read uploaded file as Excel: {e}")
        st.stop()

    # Validate columns
    missing = [c for c in [INPUT_MODEL_COL, INPUT_ROWS_COL] if c not in in_df.columns]
    if missing:
        st.error(f"Missing required columns in uploaded file: {missing}. Required: ['model', 'rows']")
        st.stop()

    # Price it
    out_df = price_input_df(in_df, lookup)

    # Preview
    left, right = st.columns([2, 1])
    with left:
        st.subheader("Preview (first 50 rows)")
        st.dataframe(out_df[[INPUT_MODEL_COL, INPUT_ROWS_COL, OUTPUT_PRICE_COL, OUTPUT_STATUS_COL]].head(50))

    with right:
        total = len(out_df)
        ok = int((out_df[OUTPUT_STATUS_COL] == "OK").sum())
        nf = int((out_df[OUTPUT_STATUS_COL] == "NOT_FOUND").sum())
        st.subheader("Match summary")
        st.write(f"Total rows: **{total}**")
        st.write(f"Matched: **{ok}**")
        st.write(f"Not found: **{nf}**")

        if nf > 0:
            st.subheader("Not found examples")
            st.dataframe(out_df[out_df[OUTPUT_STATUS_COL] == "NOT_FOUND"][[INPUT_MODEL_COL, INPUT_ROWS_COL]].head(20))

    # Download output Excel
    excel_bytes = to_excel_bytes(out_df, sheet_name="PRICED")
    st.download_button(
        label="Download priced Excel",
        data=excel_bytes,
        file_name="priced_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
