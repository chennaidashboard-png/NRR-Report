import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="NRR Report", page_icon="📊", layout="wide")

REPORT_TYPES = ["Client", "Region", "Page No.", "Agency", "Industry", "Product", "Billing Center"]
CATEGORIES = ["Commercial", "DIPR", "DAVP", "CINEMA"]
EDITION_MODES = ["Local", "Group", "All"]


def norm_text(v):
    return " ".join(str(v).strip().split())


@st.cache_data
def load_excel(file):
    data = pd.read_excel(file, sheet_name="Data")
    rates = pd.read_excel(file, sheet_name="Sheet2")
    return data, rates


def clean_data(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    required = [
        "Date", "Total Sq.Cms", "Print Center", "Value", "Client", "Region",
        "Page No.", "Agency", "Industry", "Product", "Billing Center"
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("Missing Data-sheet columns: " + ", ".join(missing))

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Total Sq.Cms"] = pd.to_numeric(df["Total Sq.Cms"], errors="coerce").fillna(0.0)
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce").fillna(0.0)

    for c in required:
        if c not in ["Date", "Total Sq.Cms", "Value"]:
            df[c] = df[c].fillna("").astype(str).map(norm_text)

    if "Category" not in df.columns:
        df["Category"] = "Commercial"
    else:
        df["Category"] = df["Category"].fillna("Commercial").astype(str).map(norm_text)
        df.loc[df["Category"].eq(""), "Category"] = "Commercial"

    return df


def clean_rates(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    required = ["Print Center (Edition)", "Card Rate (Rs./Sq.Cm)", "Net Rate (Rs./Sq.Cm)"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("Missing Sheet2 rate-card columns: " + ", ".join(missing))

    df["Print Center (Edition)"] = df["Print Center (Edition)"].fillna("").astype(str).map(norm_text)
    df["Card Rate (Rs./Sq.Cm)"] = pd.to_numeric(df["Card Rate (Rs./Sq.Cm)"], errors="coerce")
    df["Net Rate (Rs./Sq.Cm)"] = pd.to_numeric(df["Net Rate (Rs./Sq.Cm)"], errors="coerce")
    return df.dropna(subset=["Print Center (Edition)", "Card Rate (Rs./Sq.Cm)"])


def split_centers(value):
    """Return published editions from a Print Center cell."""
    text = norm_text(value)
    if not text:
        return []
    # Source files normally use commas. Also accept semicolon/newline separators.
    parts = text.replace(";", ",").replace("\n", ",").split(",")
    return [norm_text(x) for x in parts if norm_text(x)]


def calculate(d, rates, mode):
    """Calculate Local / Group NRR.

    Group rule:
      - Different Billing Center vs Published Center = Group.
      - Same Billing Center with more than one published edition = Group.
      - Group size is allocated by edition card-rate share:
            actual size * edition rate / ALL EDITION CARD-RATE TOTAL
      - The denominator is the constant total card rate of all editions.
      - Group value is NEVER split or multiplied. The complete original value
        is counted exactly once for the ad/group.
    """
    base = d.copy()
    rate_map = {
        norm_text(k).casefold(): float(v)
        for k, v in zip(rates["Print Center (Edition)"], rates["Card Rate (Rs./Sq.Cm)"])
        if pd.notna(v)
    }

    # Group allocation denominator is the dedicated ALL-edition card rate,
    # which is 3668 in the supplied rate sheet. Do not sum the individual
    # edition rows because that would include the ALL row itself and produce
    # the wrong denominator.
    all_rate_rows = rates[
        rates["Print Center (Edition)"].astype(str).str.strip().str.casefold().eq("all")
    ]
    if not all_rate_rows.empty:
        group_rate_pool = float(all_rate_rows.iloc[0]["Card Rate (Rs./Sq.Cm)"])
    else:
        group_rate_pool = 3668.0

    rows = []
    for _, row in base.iterrows():
        published = split_centers(row["Print Center"])
        billing = norm_text(row["Billing Center"])
        actual_size = float(row["Total Sq.Cms"] or 0)
        value = float(row["Value"] or 0)

        if not published:
            scope = "Local" if not billing else "Group"
            if mode in (scope, "All"):
                r = row.to_dict()
                r["Published Edition"] = ""
                r["Allocated Size"] = actual_size
                r["Scope"] = scope
                r["Card Rate"] = np.nan
                rows.append(r)
            continue

        # One edition + same billing center = Local.
        is_local = len(published) == 1 and published[0].casefold() == billing.casefold()
        scope = "Local" if is_local else "Group"

        if mode not in (scope, "All"):
            continue

        if scope == "Local":
            r = row.to_dict()
            r["Published Edition"] = published[0]
            r["Allocated Size"] = actual_size
            r["Scope"] = "Local"
            r["Card Rate"] = rate_map.get(published[0].casefold(), np.nan)
            rows.append(r)
            continue

        # Group size uses the ACTUAL size. No 1716 -> 1713 conversion applies.
        # EXACT GROUP FORMULA:
        # Allocated Size = Actual Size * respective edition Card Rate / 3668
        # The denominator is the dedicated ALL-edition card-rate total.
        valid = [(ed, rate_map.get(ed.casefold(), 0.0)) for ed in published]
        rate_pool = group_rate_pool

        if rate_pool > 0:
            # Keep the original group value only once. The edition rows carry
            # the allocated sizes, while Value=0 on subsequent rows prevents
            # the group value from being multiplied by the number of editions.
            first_value = True
            for ed, rate in valid:
                if rate <= 0:
                    continue
                r = row.to_dict()
                r["Published Edition"] = ed
                r["Allocated Size"] = actual_size * rate / rate_pool
                r["Scope"] = "Group"
                r["Card Rate"] = rate
                r["Group Rate Pool"] = rate_pool
                r["Value"] = value if first_value else 0.0
                first_value = False
                rows.append(r)
        else:
            share = actual_size / len(published) if published else 0
            first_value = True
            for ed in published:
                r = row.to_dict()
                r["Published Edition"] = ed
                r["Allocated Size"] = share
                r["Scope"] = "Group"
                r["Card Rate"] = np.nan
                r["Group Rate Pool"] = 0.0
                r["Value"] = value if first_value else 0.0
                first_value = False
                rows.append(r)

    return pd.DataFrame(rows)


def aggregate(d, report_type):
    if d.empty:
        return pd.DataFrame(columns=[report_type, "Total Sq.Cms", "Total Value", "NRR"])

    field = report_type
    out = d.groupby(field, dropna=False, as_index=False).agg(
        **{
            "Total Sq.Cms": ("Allocated Size", "sum"),
            "Total Value": ("Value", "sum")
        }
    )
    out["NRR"] = np.where(
        out["Total Sq.Cms"] > 0,
        out["Total Value"] / out["Total Sq.Cms"],
        0
    )
    return out.sort_values("Total Value", ascending=False)


def make_xlsx(df, title):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="NRR Report", startrow=2)
        ws = writer.book["NRR Report"]
        ws["A1"] = title
        ws.freeze_panes = "A4"
        for col in ws.columns:
            letter = col[0].column_letter
            width = max(len(str(x.value or "")) for x in col) + 2
            ws.column_dimensions[letter].width = min(35, max(12, width))
    return bio.getvalue()


st.title("NRR Report")

with st.sidebar:
    uploaded = st.file_uploader("Upload NRR Excel file", type=["xlsx"])
    if uploaded is None:
        st.info("Upload NRR Report Datas.xlsx")
        st.stop()

try:
    data, rates = load_excel(uploaded)
    data = clean_data(data)
    rates = clean_rates(rates)
except Exception as e:
    st.error(str(e))
    st.stop()

with st.sidebar:
    category = st.selectbox("Category", CATEGORIES)
    mode = st.selectbox("NRR Based On / Edition", EDITION_MODES)
    report_type = st.selectbox("Report Type", REPORT_TYPES)

    valid_dates = data["Date"].dropna()
    lo = valid_dates.min().date()
    hi = valid_dates.max().date()
    dates = st.date_input("Date Range", (lo, hi), min_value=lo, max_value=hi)

    work = data[data["Category"].eq(category)].copy()
    if isinstance(dates, (tuple, list)) and len(dates) == 2:
        work = work[work["Date"].between(pd.Timestamp(dates[0]), pd.Timestamp(dates[1]))]

    for field in REPORT_TYPES:
        opts = sorted(work[field].dropna().astype(str).unique())
        sel = st.multiselect(field, opts)
        if sel:
            work = work[work[field].isin(sel)]

calc = calculate(work, rates, mode)
report = aggregate(calc, report_type)

st.subheader("NRR Report")

a, b, c = st.columns(3)
a.metric("Total Value", f"₹{calc['Value'].sum():,.2f}" if not calc.empty else "₹0.00")
b.metric("Allocated Size", f"{calc['Allocated Size'].sum():,.2f}" if not calc.empty else "0.00")
overall = calc["Value"].sum() / calc["Allocated Size"].sum() if (not calc.empty and calc["Allocated Size"].sum()) else 0
c.metric("Overall NRR", f"₹{overall:,.2f}")

show = report.copy()
if not show.empty:
    show["Total Sq.Cms"] = show["Total Sq.Cms"].round(2)
    show["Total Value"] = show["Total Value"].round(2)
    show["NRR"] = show["NRR"].round(2)
st.dataframe(show, use_container_width=True, hide_index=True)

st.download_button(
    "⬇️ Download NRR Report",
    make_xlsx(show, "NRR Report"),
    "NRR_Report.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
