import re
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

# ----------------------------
# Config
# ----------------------------
# Put the master pricelist Excel next to this app (or adjust the path).
PRICE_LIST_PATH = Path("Price_LIST_Eurapo_FINAL_v3.xlsx")

# Master pricelist expected columns:
PL_MODEL_COL = "Model"
PL_ROWS_COL = "Cooling Rows + Heating Row"
PL_BASE_COL = "Base Price"
PL_ROWPRICE_COL = "Row Price"

# Input file expected columns (the user selection file):
IN_MODEL_COL = "Model"
IN_ROWS_COL = "Cooling Rows + Heating Row"

# Output columns added:
OUT_BASE_COL = "Base Price"
OUT_ROWPRICE_COL = "Row Price"
OUT_TOTAL_COL = "Total Price"
OUT_STATUS_COL = "Match Status"


# ----------------------------
# Helpers
# ----------------------------
def norm_model(x) -> str:
    """
    Normalize model strings to increase match rate:
    - uppercase
    - collapse whitespace
    - replace '-' with space
    """
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).strip().upper()
    s = s.replace("-", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def norm_rows(x) -> str:
    """
    Normalize 'Cooling Rows + Heating Row' values like '2+1R'
    - uppercase
    - remove spaces
    """
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).strip().upper()
    s = re.sub(r"\s+", "", s)
    return s


def _find_column(df: pd.DataFrame, wanted: str):
    """Case/space-insensitive column matcher. Returns actual column name or None."""
    w = re.sub(r"\s+", " ", wanted.strip().lower())
    for c in df.columns:
        cc = re.sub(r"\s+", " ", str(c).strip().lower())
        if cc == w:
            return c
    return None


@st.cache_data
def load_pricelist(pricelist_path: Path):
    """
    Loads the pricelist and builds a lookup dict keyed by (model_norm, rows_norm)
    -> (base_price, row_price)
    """
    if not pricelist_path.exists():
        raise FileNotFoundError(f"Pricelist file not found: {pricelist_path}")

    xls = pd.ExcelFile(pricelist_path)

    required = {PL_MODEL_COL, PL_ROWS_COL, PL_BASE_COL, PL_ROWPRICE_COL}
    chosen_sheet = None

    for sh in xls.sheet_names:
        preview = pd.read_excel(xls, sheet_name=sh, nrows=5)
        if all(_find_column(preview, col) is not None for col in required):
            chosen_sheet = sh
            break

    if chosen_sheet is None:
        raise ValueError(
            f"No sheet in '{pricelist_path.name}' contains the required columns: {sorted(required)}. "
            f"Sheets found: {xls.sheet_names}"
        )

    df = pd.read_excel(xls, sheet_name=chosen_sheet)

    # Resolve actual column names in the sheet
    c_model = _find_column(df, PL_MODEL_COL)
    c_rows = _find_column(df, PL_ROWS_COL)
    c_base = _find_column(df, PL_BASE_COL)
    c_rowp = _find_column(df, PL_ROWPRICE_COL)

    df["model_norm"] = df[c_model].apply(norm_model)
    df["rows_norm"] = df[c_rows].apply(norm_rows)

    lookup = {}
    dupes = []

    for _, row in df.iterrows():
        key = (row["model_norm"], row["rows_norm"])
        base = float(row[c_base]) if pd.notna(row[c_base]) else None
        rowp = float(row[c_rowp]) if pd.notna(row[c_rowp]) else None

        if key in lookup and lookup[key] != (base, rowp):
            dupes.append((key, lookup[key], (base, rowp)))
        lookup[key] = (base, rowp)

    return df, lookup, chosen_sheet, dupes


def price_input_df(input_df: pd.DataFrame, lookup: dict):
    """
    Adds Base Price, Row Price, Total Price (= Base + Row), and Match Status columns
    by matching (Model, Cooling Rows + Heating Row).
    """
    df = input_df.copy()

    # Resolve input columns (case/space-insensitive)
    c_model = _find_column(df, IN_MODEL_COL)
    c_rows = _find_column(df, IN_ROWS_COL)

    missing = []
    if c_model is None:
        missing.append(IN_MODEL_COL)
    if c_rows is None:
        missing.append(IN_ROWS_COL)
    if missing:
        raise ValueError(
            f"Missing required columns in uploaded file: {missing}. "
            f"Required columns (exact names preferred): ['{IN_MODEL_COL}', '{IN_ROWS_COL}']"
        )

    df["_model_norm"] = df[c_model].apply(norm_model)
    df["_rows_norm"] = df[c_rows].apply(norm_rows)

    base_out, rowp_out, total_out, status = [], [], [], []

    for m, r in zip(df["_model_norm"], df["_rows_norm"]):
        val = lookup.get((m, r))
        if val is None or val == (None, None):
            base_out.append(None)
            rowp_out.append(None)
            total_out.append(None)
            status.append("NOT_FOUND")
        else:
            base, rowp = val
            base_out.append(base)
            rowp_out.append(rowp)
            if base is None or rowp is None:
                total_out.append(None)
                status.append("INCOMPLETE_PRICE")
            else:
                total_out.append(float(base) + float(rowp))
                status.append("OK")

    df[OUT_BASE_COL] = base_out
    df[OUT_ROWPRICE_COL] = rowp_out
    df[OUT_TOTAL_COL] = total_out
    df[OUT_STATUS_COL] = status

    df = df.drop(columns=["_model_norm", "_rows_norm"], errors="ignore")
    return df


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "PRICED") -> bytes:
    """Convert DataFrame to Excel bytes."""
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    bio.seek(0)
    return bio.getvalue()


# ----------------------------
# Streamlit UI
# ----------------------------
st.set_page_config(page_title="Eurapo FCU Price Filler", layout="wide")

# Your font/theme tweaks stay in your Streamlit config/theme; this app won't override them.
st.title("Eurapo FCU Price Filler")
st.caption(
    "Upload an Excel with columns: 'Model' and 'Cooling Rows + Heating Row' — "
    "the app fills Base Price, Row Price, and computes Total Price (= Base + Row)."
)

# Load pricelist
try:
    pl_df, lookup, sheet, dupes = load_pricelist(PRICE_LIST_PATH)
    st.success(f"Loaded pricelist: {PRICE_LIST_PATH.name} | Sheet: {sheet} | Rows: {len(pl_df)}")
    if dupes:
        st.warning(
            f"Warning: found {len(dupes)} duplicate (Model, Cooling Rows + Heating Row) keys "
            f"with different prices in pricelist. Using the last seen value."
        )
except Exception as e:
    st.error(str(e))
    st.stop()

with st.expander("Show pricelist preview (debug)"):
    st.dataframe(pl_df.head(20))

uploaded = st.file_uploader("Upload your Excel (XLSX)", type=["xlsx"])

if uploaded:
    # Read input
    try:
        in_df = pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"Could not read uploaded file as Excel: {e}")
        st.stop()

    # Price it
    try:
        out_df = price_input_df(in_df, lookup)
    except Exception as e:
        st.error(str(e))
        st.stop()

    # Preview
    left, right = st.columns([2, 1])

    with left:
        st.subheader("Preview (first 50 rows)")
        c_model = _find_column(out_df, IN_MODEL_COL) or IN_MODEL_COL
        c_rows = _find_column(out_df, IN_ROWS_COL) or IN_ROWS_COL
        cols_to_show = [c_model, c_rows, OUT_BASE_COL, OUT_ROWPRICE_COL, OUT_TOTAL_COL, OUT_STATUS_COL]
        cols_to_show = [c for c in cols_to_show if c in out_df.columns]
        st.dataframe(out_df[cols_to_show].head(50))

    with right:
        total = len(out_df)
        ok = int((out_df[OUT_STATUS_COL] == "OK").sum()) if OUT_STATUS_COL in out_df.columns else 0
        nf = int((out_df[OUT_STATUS_COL] == "NOT_FOUND").sum()) if OUT_STATUS_COL in out_df.columns else 0
        inc = int((out_df[OUT_STATUS_COL] == "INCOMPLETE_PRICE").sum()) if OUT_STATUS_COL in out_df.columns else 0

        st.subheader("Match summary")
        st.write(f"Total rows: **{total}**")
        st.write(f"Matched (OK): **{ok}**")
        st.write(f"Not found: **{nf}**")
        if inc:
            st.write(f"Incomplete price rows: **{inc}**")

        if nf > 0:
            st.subheader("Not found examples")
            c_model = _find_column(out_df, IN_MODEL_COL) or IN_MODEL_COL
            c_rows = _find_column(out_df, IN_ROWS_COL) or IN_ROWS_COL
            show_cols = [c for c in [c_model, c_rows, OUT_STATUS_COL] if c in out_df.columns]
            st.dataframe(out_df[out_df[OUT_STATUS_COL] == "NOT_FOUND"][show_cols].head(30))

    # Download output Excel
    excel_bytes = to_excel_bytes(out_df, sheet_name="PRICED")
    st.download_button(
        label="Download priced Excel",
        data=excel_bytes,
        file_name="priced_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
