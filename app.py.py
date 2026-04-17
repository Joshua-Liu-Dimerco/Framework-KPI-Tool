import io
import pandas as pd
import plotly.express as px
import streamlit as st
import plotly.graph_objects as go

st.set_page_config(page_title="Framework Outbound KPI Tool", layout="wide")

REQUIRED_COLUMNS = [
    "Committed Date",
    "Created Date",
    "Created Time",
    "DocNo",
    "MoveType",
    "TransNo",
    "TransStatus",
    "Country code",
]

STATUS_MAP = {
    "Complete": "Complete",
    "PGI": "PGI",
    "Void": "Void",
}
from datetime import datetime, timedelta

EU_COUNTRIES = [
    "DE", "NL", "FR", "IE", "AT", "ES", "BE", "EE", "LT", "RO",
    "BG", "MT", "CY", "FI", "PL", "PT", "LV", "LU", "CZ", "SK",
    "SE", "HR", "SI", "GR", "DK", "IT"
]

KPI_RULES = {
    1: {"US": 1, "EU": 1, "CA": 2, "GB": 2, "AU": 2, "NO": 2, "SG": 2, "NZ": 2, "CH": 2, "TW": 0},  # Mon
    2: {"US": 0, "EU": 0, "CA": 1, "GB": 1, "AU": 1, "NO": 1, "SG": 1, "NZ": 1, "CH": 1, "TW": 0},  # Tue
    3: {"US": 2, "EU": 2, "CA": 0, "GB": 0, "AU": 0, "NO": 0, "SG": 0, "NZ": 0, "CH": 0, "TW": 0},  # Wed
    4: {"US": 1, "EU": 1, "CA": 6, "GB": 6, "AU": 6, "NO": 6, "SG": 6, "NZ": 6, "CH": 6, "TW": 0},  # Thu
    5: {"US": 0, "EU": 0, "CA": 5, "GB": 5, "AU": 5, "NO": 5, "SG": 5, "NZ": 5, "CH": 5, "TW": 0},  # Fri
    6: {"US": 3, "EU": 3, "CA": 4, "GB": 4, "AU": 4, "NO": 4, "SG": 4, "NZ": 4, "CH": 4, "TW": 2},  # Sat
    7: {"US": 2, "EU": 2, "CA": 3, "GB": 3, "AU": 3, "NO": 3, "SG": 3, "NZ": 3, "CH": 3, "TW": 1},  # Sun
}

HOLIDAYS = [
    # 先放幾個測試值，之後再改成外部表
    # "2026-01-01",
    # "2026-02-28",
]

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]
    return df


@st.cache_data
def load_data(uploaded_file) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, sheet_name="OrderReport")
    df = normalize_columns(df)
    return df


@st.cache_data
def load_special_rules(uploaded_file) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file)
    df.columns = [str(col).strip() for col in df.columns]
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.normalize()
    df["Add Days"] = pd.to_numeric(df["Add Days"], errors="coerce")
    return df


def build_special_rule_dict(special_df: pd.DataFrame) -> dict:
    if special_df is None or special_df.empty:
        return {}

    return {
        row["Date"].date(): int(row["Add Days"])
        for _, row in special_df.dropna(subset=["Date", "Add Days"]).iterrows()
    }


def validate_columns(df: pd.DataFrame):
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    return missing

def get_region(country_code: str) -> str:
    country_code = str(country_code).upper().strip()

    if country_code in EU_COUNTRIES:
        return "EU"

    region_list = ["US", "CA", "GB", "AU", "NO", "SG", "NZ", "CH", "TW"]
    if country_code in region_list:
        return country_code

    return "OTHER"


def get_base_date(created_date, created_time_str, country):
    if pd.isna(created_date):
        return pd.NaT

    base_date = pd.to_datetime(created_date)

    try:
        time_obj = pd.to_datetime(created_time_str, format="%H:%M:%S", errors="coerce")
        if pd.isna(time_obj):
            time_obj = pd.to_datetime(created_time_str, errors="coerce")
    except:
        time_obj = pd.NaT

    if pd.notna(time_obj):
        actual_time = time_obj.time()

        if country == "TW":
            cutoff = pd.to_datetime("11:00:00").time()
        else:
            cutoff = pd.to_datetime("14:00:00").time()

        if actual_time >= cutoff:
            base_date = base_date + pd.Timedelta(days=1)

    return base_date.normalize()


def add_business_holiday_offset(start_date, days_to_add, holidays):
    if pd.isna(start_date):
        return pd.NaT

    result_date = pd.to_datetime(start_date)

    if pd.isna(days_to_add):
        return pd.NaT

    result_date = result_date + pd.Timedelta(days=int(days_to_add))

    holiday_set = {pd.to_datetime(d).date() for d in holidays}

    while result_date.date() in holiday_set:
        result_date = result_date + pd.Timedelta(days=1)

    return result_date

def prepare_data(df: pd.DataFrame, special_rule_dict=None) -> pd.DataFrame:
    df = df.copy()

    if special_rule_dict is None:
        special_rule_dict = {}

    df["Created Date"] = pd.to_datetime(df["Created Date"], errors="coerce")
    df["Committed Date"] = pd.to_datetime(df["Committed Date"], errors="coerce")

    # 加在這裡
    df["945 Day"] = df["Committed Date"].dt.date

    df["Created Time"] = df["Created Time"].astype(str).str.strip()
    df["TransStatus"] = df["TransStatus"].astype(str).str.strip()
    df["Country code"] = df["Country code"].astype(str).str.upper().str.strip()

    df["Status Group"] = df["TransStatus"].map(STATUS_MAP).fillna(df["TransStatus"])
    df["Order Day"] = df["Created Date"].dt.date
    df["Committed Day"] = df["Committed Date"].dt.date

    # KPI logic
    df["Region"] = df["Country code"].apply(get_region)

    df["Base Date"] = df.apply(
    lambda row: get_base_date(
        row["Created Date"],
        row["Created Time"],
        row["Country code"]
    ),
    axis=1
)

    df["Weekday No"] = df["Base Date"].dt.weekday + 1  # Monday=1 ... Sunday=7

    df["Transit Days"] = df.apply(
        lambda row: KPI_RULES.get(row["Weekday No"], {}).get(row["Region"], None),
        axis=1
    )

    df["KPI Failed Date"] = df.apply(
        lambda row: add_business_holiday_offset(row["Base Date"], row["Transit Days"], HOLIDAYS),
        axis=1
    )
    df["945 Day"] = pd.to_datetime(df["Committed Date"], errors="coerce").dt.date
    df["Need Fulfill Day"] = pd.to_datetime(df["KPI Failed Date"], errors="coerce").dt.date

    df["KPI Result"] = df.apply(
    lambda row: "Failed"
    if pd.notna(row["Committed Date"])
       and pd.notna(row["KPI Failed Date"])
       and row["Committed Date"] > row["KPI Failed Date"]
    else "In KPI",
    axis=1
    )

    return df
    
def build_daily_kpi_summary(df: pd.DataFrame) -> pd.DataFrame:
    daily_945 = (
        df.dropna(subset=["945 Day"])
        .groupby("945 Day")
        .agg(total_945=("DocNo", "count"))
        .reset_index()
        .rename(columns={"945 Day": "Report Date"})
    )

    daily_kpi = (
        df.dropna(subset=["Need Fulfill Day"])
        .groupby("Need Fulfill Day")
        .agg(
            need_fulfill=("DocNo", "count"),
            in_kpi=("KPI Result", lambda x: (x == "In KPI").sum()),
            failed=("KPI Result", lambda x: (x == "Failed").sum()),
        )
        .reset_index()
        .rename(columns={"Need Fulfill Day": "Report Date"})
    )

    summary = pd.merge(
        daily_kpi,
        daily_945,
        on="Report Date",
        how="outer"
    ).fillna(0)

    summary["kpi_rate"] = summary.apply(
        lambda row: row["in_kpi"] / row["need_fulfill"]
        if row["need_fulfill"] > 0 else 1,
        axis=1
    )

    summary["Report Date"] = pd.to_datetime(summary["Report Date"])
    summary = summary.sort_values("Report Date")
    summary["Report Date"] = summary["Report Date"].dt.date

    return summary  

def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    summary = (
        df.groupby(["Order Day", "Status Group"], dropna=False)
        .size()
        .reset_index(name="Count")
        .sort_values(["Order Day", "Status Group"])
    )
    return summary


def export_summary_excel(clean_df: pd.DataFrame, summary_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        clean_df.to_excel(writer, sheet_name="Clean_Data", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
    output.seek(0)
    return output.getvalue()

def build_outbound_kpi_chart(summary_df: pd.DataFrame):

    summary_df = summary_df.copy()
    summary_df["Report Date"] = pd.to_datetime(summary_df["Report Date"]).dt.date

    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=summary_df["Report Date"],
        y=summary_df["total_945"],
        name="945"
    ))

    fig.add_trace(go.Bar(
        x=summary_df["Report Date"],
        y=summary_df["need_fulfill"],
        name="Need Fulfill"
    ))

    fig.add_trace(go.Bar(
        x=summary_df["Report Date"],
        y=summary_df["in_kpi"],
        name="In KPI"
    ))

    fig.add_trace(go.Bar(
        x=summary_df["Report Date"],
        y=summary_df["failed"],
        name="Failed"
    ))

    fig.add_trace(go.Scatter(
        x=summary_df["Report Date"],
        y=summary_df["kpi_rate"],
        name="KPI Rate",
        mode="lines+markers",
        yaxis="y2"
    ))

    fig.update_layout(
    title="Outbound KPI",
    barmode="group",
    xaxis=dict(
        title="Date",
        tickformat="%Y-%m-%d"
    ),
    yaxis=dict(title="Volume"),
    yaxis2=dict(
        title="KPI Rate",
        overlaying="y",
        side="right",
        tickformat=".0%"
    ),
    legend=dict(orientation="h"),
    height=600
)

    return fig

st.title("Framework Outbound KPI Tool")

uploaded_file = st.file_uploader("Upload raw data (.xlsx)", type=["xlsx"])
special_rule_file = st.file_uploader("Upload special date rule (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    raw_df = load_data(uploaded_file)

    special_rule_dict = {}

    if special_rule_file is not None:
        special_df = load_special_rules(special_rule_file)
        special_rule_dict = build_special_rule_dict(special_df)

    df = prepare_data(raw_df, special_rule_dict)

    summary_df = build_daily_kpi_summary(df)

    # ===== 日期篩選 =====
    if not summary_df.empty:
        min_date = pd.to_datetime(summary_df["Report Date"]).min().date()
        max_date = pd.to_datetime(summary_df["Report Date"]).max().date()

        date_range = st.date_input(
            "Select date range",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date
        )

        if isinstance(date_range, tuple) and len(date_range) == 2:
            start_date, end_date = date_range

            filtered_summary_df = summary_df[
                (pd.to_datetime(summary_df["Report Date"]).dt.date >= start_date) &
                (pd.to_datetime(summary_df["Report Date"]).dt.date <= end_date)
            ].copy()

            filtered_df = df[
                (pd.to_datetime(df["Need Fulfill Day"]).dt.date >= start_date) &
                (pd.to_datetime(df["Need Fulfill Day"]).dt.date <= end_date)
            ].copy()
        else:
            filtered_summary_df = summary_df.copy()
            filtered_df = df.copy()
    else:
        filtered_summary_df = summary_df.copy()
        filtered_df = df.copy()

    # ===== KPI 指標 =====
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Rows", f"{len(filtered_df):,}")
    col2.metric("Complete", f"{(filtered_df['Status Group'] == 'Complete').sum():,}")
    col3.metric("PGI", f"{(filtered_df['Status Group'] == 'PGI').sum():,}")
    col4.metric("Void", f"{(filtered_df['Status Group'] == 'Void').sum():,}")

    # ===== 圖表 =====
    st.subheader("Outbound KPI Chart")
    fig_kpi = build_outbound_kpi_chart(filtered_summary_df)
    st.plotly_chart(fig_kpi, width="stretch")

    # ===== Summary 表 =====
    display_summary_df = filtered_summary_df.rename(columns={
        "Report Date": "Date",
        "total_945": "945",
        "need_fulfill": "Need Fulfill",
        "in_kpi": "In KPI",
        "failed": "Failed",
        "kpi_rate": "KPI Rate"
    })

    display_summary_df["KPI Rate"] = display_summary_df["KPI Rate"].map(lambda x: f"{x:.2%}")

    st.subheader("Daily KPI Summary")
    st.dataframe(display_summary_df, width="stretch")
    st.subheader("Processed Data Preview")
    display_df = filtered_df.copy()

    date_cols = [
    "Committed Date",
    "Created Date",
    "Base Date",
    "KPI Failed Date"
]

    for col in date_cols:
     if col in display_df.columns:
        display_df[col] = pd.to_datetime(display_df[col], errors="coerce").dt.date

    st.dataframe(display_df, width="stretch")
  

    st.subheader("Raw Data Preview")
    st.dataframe(raw_df, width="stretch")
else:
    st.info("Please upload your raw data Excel file to begin.")