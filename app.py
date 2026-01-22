import streamlit as st
import pandas as pd
import re
from io import BytesIO
from pathlib import Path

PRICE_LIST_PATH = Path("data/Price LIST FCU Eurapo 4pipe AC.xlsx")
PRICE_SHEET = "MASTER"          # change if your sheet name differs
OUTPUT_PRICE_COLUMN = "price_eur"  # name of the column we write in output

def norm_model(x: str) -> str:
    if x is None:
        return ""
    x = str(x).strip().upper()
    x = re.sub(r"\s+", " ", x)
    x = x.replace("-", " ")
    x = re.sub(r"\s+", " ", x).strip()

    # Normalize EBH formats: "EBH050" / "EBH 50" -> "EBH 050"
    x = re.sub(r"^(EBH)\s*0*([0-9]{1,3})$", lambda m: f"{m.group(1)} {int(m.group(2)):03d}", x)
    return x

def norm_rows(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "":
        return None
    try:
        return int(float(v))
    except:
        return None

@st.cache_data
def load_pricelist():
    df = pd.read_excel(PRICE_LIST_PATH, sheet_name=PRICE_SHEET)
    # expected columns in MASTER: model, rows, total_price_eur
    df["model_norm"] = df["model"].apply(norm_model)
    df["rows_norm"] = df["rows"].apply(norm_rows)

    lookup = {}
    for r in df.itertuples(index=False):
        key = (r.model_norm, r.rows_norm)
        lookup[key] = float(r.total_price_eur)

    return df, lookup

st.title("Eurapo FCU Price Filler (4-pipe, AC motor)")

if not PRICE_LIST_PATH.exists():
    st.error(f"Built-in pricelist not found: {PRICE_LIST_PATH}. Add it to the repo.")
    st.stop()

pl_df, lookup = load_pricelist()
st.caption(f"Loaded built-in pricelist: {PRICE_LIST_PATH.name} ({len(pl_df)} rows)")

uploaded = st.file_uploader("Upload Excel with columns: model, rows", type=["xlsx"])

if uploaded:
    in_df = pd.read_excel(uploaded)

    if "model" not in in_df.columns or "rows" not in in_df.columns:
        st.error("Your uploaded file must contain columns named exactly: model, rows")
        st.stop()

    in_df["model_norm"] = in_df["model"].apply(norm_model)
    in_df["rows_norm"] = in_df["rows"].apply(norm_rows)

    prices = []
    status = []
    for m, r in zip(in_df["model_norm"], in_df["rows_norm"]):
        p = lookup.get((m, r))
        if p is None:
            prices.append(None)
            status.append("NOT_FOUND")
        else:
            prices.append(p)
            status.append("OK")

    in_df[OUTPUT_PRICE_COLUMN] = prices
    in_df["match_status"] = status

    st.subheader("Preview")
    st.dataframe(in_df[["model", "rows", OUTPUT_PRICE_COLUMN, "match_status"]].head(50))

    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        in_df.to_excel(writer, index=False, sheet_name="PRICED")
    out.seek(0)

    st.download_button(
        "Download priced Excel",
        data=out,
        file_name="priced_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
