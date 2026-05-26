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

# 倉庫使用率分頁的儲位類型清單
STORAGE_TYPES = ["RCK", "LAR", "SHF", "MED"]


# ===== 檔名日期 / 週數工具 =====
import re as _re_fn

def parse_date_from_filename(filename):
    """從檔名 (例: CR_WMSv3_Framework_Storage_Usage_Report_2026-05-20.xlsx) 解析日期。"""
    if not filename:
        return None
    m = _re_fn.search(r"(\d{4})[-_]?(\d{1,2})[-_]?(\d{1,2})", str(filename))
    if not m:
        return None
    try:
        return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))).date()
    except Exception:
        return None


def get_year_week_label(d):
    """日期 -> 'YYYY Week NN' (ISO 週)。"""
    if d is None:
        return ""
    if isinstance(d, str):
        d = pd.to_datetime(d, errors="coerce")
        if pd.isna(d):
            return ""
        d = d.date()
    iso = d.isocalendar()
    try:
        return f"{iso.year} Week {iso.week:02d}"
    except AttributeError:
        return f"{iso[0]} Week {iso[1]:02d}"


def dedup_files_by_week(pairs):
    """同週多檔保留日期最晚的那筆。"""
    bucket = {}
    for f, d in pairs:
        if d is None:
            continue
        wk = get_year_week_label(d)
        if wk not in bucket or d > bucket[wk]["date"]:
            bucket[wk] = {"file": f, "date": d, "week_label": wk}
    return sorted(bucket.values(), key=lambda x: x["date"])


# ===== Exception (排除訂單) 載入器 =====

def load_exception_orders(uploaded_file):
    """讀取 Exception 檔，回傳訂單號集合 (大寫去空白)。"""
    if uploaded_file is None:
        return set()
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0)
    except Exception:
        return set()
    df.columns = [str(c).strip() for c in df.columns]
    target_col = None
    for col in df.columns:
        cl = str(col).strip().lower()
        if cl in ("order", "order#", "orderno", "order no", "order_no",
                  "docno", "doc no", "doc_no", "transno", "trans no",
                  "inbship no.", "inbship no", "inbship_no",
                  "declaration#", "declaration",
                  "訂單", "訂單號", "訂單號碼", "單號"):
            target_col = col
            break
    if target_col is None and len(df.columns) > 0:
        target_col = df.columns[0]
    if target_col is None or df.empty:
        return set()
    vals = df[target_col].dropna().astype(str).str.strip().str.upper()
    return set(v for v in vals.tolist() if v)

# ===== Inventory Item Master (從上傳的檔案讀取) =====
# 在 Storage Usage 分頁需上傳 Item Master Excel，內含 itemcode → Item Type 對照。
# 預期欄位：itemcode, Item Type （大小寫不拘，會自動正規化）。

@st.cache_data
def load_item_master(uploaded_file) -> dict:
    """讀取 Item Master 檔案並回傳 {itemcode_upper: item_type} dict。

    支援 .xlsx / .csv。會自動偵測 itemcode 與 Item Type 欄位名稱。
    """
    name = getattr(uploaded_file, "name", "") or ""
    if name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        # Excel：先試讀 OrderReport，否則第一個工作表
        try:
            df = pd.read_excel(uploaded_file, sheet_name="OrderReport")
        except Exception:
            df = pd.read_excel(uploaded_file, sheet_name=0)

    df = normalize_columns(df)

    # 自動找 itemcode 欄
    code_col = None
    type_col = None
    for col in df.columns:
        c = str(col).strip().lower()
        if code_col is None and c in ("itemcode", "item code", "item_code", "framework pn", "pn", "料號", "品號"):
            code_col = col
        if type_col is None and c in ("item type", "itemtype", "item_type", "種類", "類別", "type"):
            type_col = col

    if code_col is None or type_col is None:
        return {}

    sub = df[[code_col, type_col]].dropna(how="all").copy()
    sub[code_col] = sub[code_col].astype(str).str.strip().str.upper()
    sub[type_col] = sub[type_col].astype(str).str.strip()

    # 排除像「Total」這種總和列、空字串
    sub = sub[(sub[code_col] != "") & (sub[code_col].str.upper() != "TOTAL")]
    sub = sub[sub[type_col] != ""]

    # 同 itemcode 重複時取第一筆
    sub = sub.drop_duplicates(subset=[code_col], keep="first")

    return dict(zip(sub[code_col], sub[type_col]))


def get_item_type(itemcode, item_master=None) -> str:
    """回傳 itemcode 對應的 Item Type；找不到或未提供 master 時回傳 'Unknown'。"""
    if itemcode is None or not item_master:
        return "Unknown"
    key = str(itemcode).strip().upper()
    if not key:
        return "Unknown"
    return item_master.get(key, "Unknown")


def build_item_master_template_bytes() -> bytes:
    """產生 Item Master 匯入範本 (.xlsx)，含正確欄名與幾筆示範資料。"""
    sample_df = pd.DataFrame([
        {"itemcode": "FRAEXAMPLE01", "Item Type": "Laptop"},
        {"itemcode": "FRAEXAMPLE02", "Item Type": "Keyboard"},
        {"itemcode": "FRAEXAMPLE03", "Item Type": "Accessories"},
        {"itemcode": "FRAEXAMPLE04", "Item Type": "Mainboard"},
        {"itemcode": "FRAEXAMPLE05", "Item Type": "SSD"},
    ])

    instructions_df = pd.DataFrame({
        "說明 / Instructions": [
            "1. 在『OrderReport』分頁填入 itemcode 與 Item Type 兩欄。",
            "2. itemcode 為品號，會自動轉大寫去空白。",
            "3. Item Type 為品項類別，可自訂；常見類別請參考下方清單。",
            "4. 同 itemcode 重複時，會取第一筆。",
            "5. 含『Total』字樣的列會被自動跳過。",
            "",
            "常見 Item Type 範例：",
            "  Laptop / Mainboard / Keyboard / Input Cover / Bezel",
            "  SSD / RAM / Power Adapter / Expansion / Accessories / Desktop",
            "",
            "也可改用以下欄名（系統會自動辨識）：",
            "  itemcode 欄：itemcode / item code / Framework PN / 料號 / 品號",
            "  Item Type 欄：Item Type / 種類 / 類別 / type",
        ]
    })

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sample_df.to_excel(writer, sheet_name="OrderReport", index=False)
        instructions_df.to_excel(writer, sheet_name="說明", index=False)
    output.seek(0)
    return output.getvalue()

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


def add_business_holiday_offset(start_date, days_to_add, holidays, special_rule_dict=None):
    if pd.isna(start_date):
        return pd.NaT

    if pd.isna(days_to_add):
        return pd.NaT

    if special_rule_dict is None:
        special_rule_dict = {}

    result_date = pd.to_datetime(start_date).normalize()

    # 先加原本 KPI 規則天數
    result_date = result_date + pd.Timedelta(days=int(days_to_add))

    holiday_set = {pd.to_datetime(d).date() for d in holidays}

    # 如果落在 holiday，往後推
    while result_date.date() in holiday_set:
        result_date = result_date + pd.Timedelta(days=1)

    # 再套用 special date 額外加天數
    extra_days = special_rule_dict.get(result_date.date(), 0)
    if pd.notna(extra_days) and int(extra_days) != 0:
        result_date = result_date + pd.Timedelta(days=int(extra_days))

        # 加完後如果又碰到 holiday，再往後推
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
        lambda row: add_business_holiday_offset(
            row["Base Date"],
            row["Transit Days"],
            HOLIDAYS,
            special_rule_dict
        ),
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
import re

# ===== Inbound 專用 =====

def normalize_decl_no(value):
    if pd.isna(value):
        return ""

    value = str(value).strip().upper()
    value = re.sub(r"[^A-Z0-9]", "", value)
    return value


@st.cache_data
def load_inbound_raw(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df.columns = [str(col).strip() for col in df.columns]
    return df


@st.cache_data
def load_inbound_mapping(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df.columns = [str(col).strip() for col in df.columns]

    df["報單單號 Clean"] = df["報單單號"].apply(normalize_decl_no)
    df["Inbound Date"] = pd.to_datetime(df["Inbound Date"], errors="coerce")

    return df


def calculate_inbound_kpi_date(inbound_date):
    if pd.isna(inbound_date):
        return pd.NaT

    result_date = pd.to_datetime(inbound_date) + pd.Timedelta(days=1)

    weekday = result_date.weekday()

    if weekday == 5:  # Saturday
        result_date += pd.Timedelta(days=2)
    elif weekday == 6:  # Sunday
        result_date += pd.Timedelta(days=1)

    return result_date.normalize()


def prepare_inbound_data(raw_df, mapping_df):
    df = raw_df.copy()

    df.columns = [str(col).strip() for col in df.columns]

    # 清洗報單
    df["Declaration Clean"] = df["Declaration#"].apply(normalize_decl_no)

    mapping_dict = (
        mapping_df.dropna(subset=["報單單號 Clean", "Inbound Date"])
        .drop_duplicates(subset=["報單單號 Clean"], keep="first")
        .set_index("報單單號 Clean")["Inbound Date"]
        .to_dict()
    )

    df["Mapped Inbound Date"] = df["Declaration Clean"].map(mapping_dict)

    df["Mapping Status"] = df["Mapped Inbound Date"].apply(
        lambda x: "Matched" if pd.notna(x) else "確認報單號碼"
    )

    df["Need Fulfill Date"] = df["Mapped Inbound Date"].apply(calculate_inbound_kpi_date)
    df["KPI Failed Date"] = df["Need Fulfill Date"]

# ⭐ KPI 判斷（這是你缺的核心）
    df["Actual Date"] = pd.to_datetime(df["Date"], errors="coerce")

    df["KPI Result"] = df.apply(
    lambda row: "Failed"
    if pd.notna(row["Actual Date"])
       and pd.notna(row["Need Fulfill Date"])
       and row["Actual Date"] > row["Need Fulfill Date"]
    else "In KPI",
    axis=1
)

    return df

    return df


def build_inbound_kpi_summary(df):

    df = df.copy()

    df["Actual Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date

    summary = (
        df.dropna(subset=["Actual Date"])
        .groupby("Actual Date")
        .agg(
            in_kpi=("KPI Result", lambda x: (x == "In KPI").sum()),
            failed=("KPI Result", lambda x: (x == "Failed").sum())
        )
        .reset_index()
        .rename(columns={"Actual Date": "Report Date"})
    )

    summary["total"] = summary["in_kpi"] + summary["failed"]

    summary["kpi_rate"] = summary.apply(
        lambda row: row["in_kpi"] / row["total"]
        if row["total"] > 0 else 1,
        axis=1
    )

    summary = summary.sort_values("Report Date")

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
def build_inbound_kpi_chart(summary_df):

    fig = go.Figure()

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
        title="Inbound KPI",
        barmode="stack",
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


# ===== Warehouse Utilization (倉庫儲位使用率) =====

def _normalize_loc_code(value) -> str:
    """把儲位代碼統一成大寫去空白，避免比對失敗。"""
    if pd.isna(value):
        return ""
    return str(value).strip().upper()


@st.cache_data
def load_storage_master(uploaded_file) -> pd.DataFrame:
    """讀取儲位總表。預期欄位：Location, Type。"""
    # 先試讀指定的工作表，若不存在則讀第一個
    try:
        df = pd.read_excel(uploaded_file, sheet_name="工作表1")
    except Exception:
        df = pd.read_excel(uploaded_file, sheet_name=0)

    df = normalize_columns(df)

    # 自動找 Location / Type 欄
    loc_col = None
    type_col = None
    for col in df.columns:
        col_lower = str(col).strip().lower()
        if loc_col is None and col_lower in ("location", "locationcode", "儲位", "儲位代碼"):
            loc_col = col
        if type_col is None and col_lower in ("type", "類別", "儲位類別"):
            type_col = col

    if loc_col is None or type_col is None:
        # 無法自動辨識，回傳空表
        return pd.DataFrame(columns=["Location", "Type"])

    df = df[[loc_col, type_col]].rename(columns={loc_col: "Location", type_col: "Type"})
    df["Location"] = df["Location"].apply(_normalize_loc_code)
    df["Type"] = df["Type"].astype(str).str.strip().str.upper()
    df = df[df["Location"] != ""].drop_duplicates(subset=["Location"], keep="first")
    return df.reset_index(drop=True)


@st.cache_data
def load_storage_usage(uploaded_file) -> pd.DataFrame:
    """讀取實際庫存使用報表。預期欄位：LocationCode, itemcode, ItemDesc, Storage, Staging, Defective。"""
    try:
        df = pd.read_excel(uploaded_file, sheet_name="OrderReport")
    except Exception:
        df = pd.read_excel(uploaded_file, sheet_name=0)

    df = normalize_columns(df)

    # 自動找 LocationCode 欄
    loc_col = None
    for col in df.columns:
        if str(col).strip().lower() in ("locationcode", "location", "儲位", "儲位代碼"):
            loc_col = col
            break

    if loc_col is None:
        return pd.DataFrame(columns=["LocationCode"])

    if loc_col != "LocationCode":
        df = df.rename(columns={loc_col: "LocationCode"})

    df["LocationCode"] = df["LocationCode"].apply(_normalize_loc_code)

    # 把 Storage/Staging/Defective 轉成數字（若存在）
    for q_col in ["Storage", "Staging", "Defective"]:
        if q_col in df.columns:
            df[q_col] = pd.to_numeric(df[q_col], errors="coerce").fillna(0)

    df = df[df["LocationCode"] != ""].copy()
    return df.reset_index(drop=True)


def build_utilization_summary(master_df: pd.DataFrame, usage_df: pd.DataFrame):
    """
    回傳:
      type_summary_df  : 每個 Type 的 Total / Used / Empty / Utilization
      overall_summary  : dict, 總儲位使用率
      unmatched_df     : usage_df 中不存在於 master 的儲位 (暫時撿貨儲位，已被略過)
      used_locations   : set, master 中被使用到的 Location
    """
    if master_df is None or master_df.empty:
        empty = pd.DataFrame(columns=["Type", "Total", "Used", "Empty", "Utilization"])
        return empty, {"total": 0, "used": 0, "empty": 0, "utilization": 0.0}, pd.DataFrame(), set()

    master_locations = set(master_df["Location"].unique())

    if usage_df is None or usage_df.empty:
        usage_locations = set()
    else:
        usage_locations = set(usage_df["LocationCode"].unique())

    # 在 usage 但不在 master → 暫時撿貨儲位，略過
    unmatched_locs = usage_locations - master_locations
    if usage_df is not None and not usage_df.empty:
        unmatched_df = usage_df[usage_df["LocationCode"].isin(unmatched_locs)].copy()
    else:
        unmatched_df = pd.DataFrame()

    # 真正算進使用率的「使用中儲位」= master 與 usage 的交集
    used_locations = master_locations & usage_locations

    # 依 Type 統計
    rows = []
    for t in STORAGE_TYPES:
        type_locs = set(master_df.loc[master_df["Type"] == t, "Location"])
        total = len(type_locs)
        used = len(type_locs & used_locations)
        empty = total - used
        rate = (used / total) if total > 0 else 0.0
        rows.append({
            "Type": t,
            "Total": total,
            "Used": used,
            "Empty": empty,
            "Utilization": rate,
        })

    # 處理可能存在的其他 Type（不在 STORAGE_TYPES 清單中的）
    other_types = set(master_df["Type"].unique()) - set(STORAGE_TYPES)
    for t in sorted(other_types):
        type_locs = set(master_df.loc[master_df["Type"] == t, "Location"])
        total = len(type_locs)
        used = len(type_locs & used_locations)
        empty = total - used
        rate = (used / total) if total > 0 else 0.0
        rows.append({
            "Type": t,
            "Total": total,
            "Used": used,
            "Empty": empty,
            "Utilization": rate,
        })

    type_summary_df = pd.DataFrame(rows)

    total_all = len(master_locations)
    used_all = len(used_locations)
    overall_summary = {
        "total": total_all,
        "used": used_all,
        "empty": total_all - used_all,
        "utilization": (used_all / total_all) if total_all > 0 else 0.0,
    }

    return type_summary_df, overall_summary, unmatched_df, used_locations


def build_utilization_chart(type_summary_df: pd.DataFrame):
    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=type_summary_df["Type"],
        y=type_summary_df["Used"],
        name="Used",
        marker_color="#1f77b4",
    ))

    fig.add_trace(go.Bar(
        x=type_summary_df["Type"],
        y=type_summary_df["Empty"],
        name="Empty",
        marker_color="#d3d3d3",
    ))

    fig.add_trace(go.Scatter(
        x=type_summary_df["Type"],
        y=type_summary_df["Utilization"],
        name="Utilization",
        mode="lines+markers+text",
        text=[f"{v:.1%}" for v in type_summary_df["Utilization"]],
        textposition="top center",
        yaxis="y2",
        line=dict(color="#ff7f0e"),
    ))

    fig.update_layout(
        title="Warehouse Utilization by Type",
        barmode="stack",
        xaxis=dict(title="Storage Type"),
        yaxis=dict(title="Locations"),
        yaxis2=dict(
            title="Utilization",
            overlaying="y",
            side="right",
            tickformat=".0%",
            range=[0, 1.1],
        ),
        legend=dict(orientation="h"),
        height=500,
    )

    return fig


def build_weekly_trend_chart(weekly_df, selected_types):
    """每週變化趨勢圖（總使用率版）。

    需求：不依儲位類型分色，每週只顯示一支堆疊柱
    (Used + Empty = 總儲位)，柱子內以白色粗體顯示使用率，
    視覺上更清楚。"""
    fig = go.Figure()
    if weekly_df is None or weekly_df.empty:
        fig.update_layout(title="Weekly Storage Utilization Trend (尚無資料)")
        return fig

    week_order = (
        weekly_df[["Week Label", "Date"]]
        .drop_duplicates().sort_values("Date")["Week Label"].tolist()
    )

    # 仍尊重使用者勾選的儲位類型，但加總成單一數值
    types_to_include = [t for t in (selected_types or []) if t in weekly_df["Type"].unique()]
    if not types_to_include:
        types_to_include = sorted(weekly_df["Type"].unique())

    # 每週彙總成單一筆：Total / Used / Empty / Utilization
    agg_rows = []
    for wk in week_order:
        wk_sub = weekly_df[
            (weekly_df["Week Label"] == wk)
            & (weekly_df["Type"].isin(types_to_include))
        ]
        total = int(wk_sub["Total"].sum())
        used = int(wk_sub["Used"].sum())
        empty = max(total - used, 0)
        rate = (used / total) if total > 0 else 0.0
        agg_rows.append({
            "Week Label": wk,
            "Total": total,
            "Used": used,
            "Empty": empty,
            "Utilization": rate,
        })
    agg_df = pd.DataFrame(agg_rows)

    # Used 柱：在柱子內用白色粗體顯示「Used: 979 / 75.19%」
    fig.add_trace(go.Bar(
        x=agg_df["Week Label"],
        y=agg_df["Used"],
        name="Used",
        marker_color="#1f77b4",
        text=[
            f"Used: {u:,}<br><b>{r:.2%}</b>"
            for u, r in zip(agg_df["Used"], agg_df["Utilization"])
        ],
        textposition="inside",
        insidetextanchor="middle",
        textfont=dict(color="white", size=16, family="Arial Black"),
        hovertemplate=(
            "Week: %{x}<br>Used: %{y:,}"
            "<br>Total: %{customdata[0]:,}"
            "<br>Utilization: %{customdata[1]:.2%}<extra></extra>"
        ),
        customdata=list(zip(agg_df["Total"], agg_df["Utilization"])),
    ))

    # Empty 柱：補滿到總儲位高度（淺灰色）
    fig.add_trace(go.Bar(
        x=agg_df["Week Label"],
        y=agg_df["Empty"],
        name="Empty",
        marker_color="#d3d3d3",
        text=[f"Empty: {v:,}" for v in agg_df["Empty"]],
        textposition="inside",
        insidetextanchor="middle",
        textfont=dict(color="#333", size=12),
        hovertemplate=(
            "Week: %{x}<br>Empty: %{y:,}"
            "<br>Total: %{customdata[0]:,}<extra></extra>"
        ),
        customdata=list(zip(agg_df["Total"])),
    ))

    # 副軸：使用率折線（不放文字，避免和柱子內白字重疊）
    fig.add_trace(go.Scatter(
        x=agg_df["Week Label"],
        y=agg_df["Utilization"],
        name="Utilization",
        mode="lines+markers",
        yaxis="y2",
        line=dict(color="#ff7f0e", width=3),
        marker=dict(size=10, color="#ff7f0e"),
        hovertemplate="Week: %{x}<br>Utilization: %{y:.2%}<extra></extra>",
    ))

    fig.update_layout(
        title="Weekly Storage Utilization Trend (Total)",
        barmode="stack",
        xaxis=dict(
            title="Year Week",
            type="category",
            categoryorder="array",
            categoryarray=week_order,
        ),
        yaxis=dict(title="Total Locations"),
        yaxis2=dict(
            title="Utilization",
            overlaying="y",
            side="right",
            tickformat=".0%",
            range=[0, 1.1],
        ),
        legend=dict(orientation="h"),
        height=550,
    )
    return fig


def export_utilization_excel(
    type_summary_df: pd.DataFrame,
    overall_summary: dict,
    used_detail_df: pd.DataFrame,
    empty_detail_df: pd.DataFrame,
    unmatched_df: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        overall_df = pd.DataFrame([{
            "Total Locations": overall_summary["total"],
            "Used": overall_summary["used"],
            "Empty": overall_summary["empty"],
            "Utilization": overall_summary["utilization"],
        }])
        overall_df.to_excel(writer, sheet_name="Overall", index=False)

        type_summary_df.to_excel(writer, sheet_name="By_Type", index=False)
        used_detail_df.to_excel(writer, sheet_name="Used_Locations", index=False)
        empty_detail_df.to_excel(writer, sheet_name="Empty_Locations", index=False)
        unmatched_df.to_excel(writer, sheet_name="Temp_Picking_Locations", index=False)
    output.seek(0)
    return output.getvalue()


st.title("Framework KPI Tool")

tab1, tab2, tab3, tab4 = st.tabs(["Outbound KPI", "Inbound KPI", "Warehouse Utilization", "Email Report"])

with tab1:

    # ===== Upload =====
    uploaded_file = st.file_uploader("Upload raw data (.xlsx)", type=["xlsx"])
    special_rule_file = st.file_uploader("Upload special date rule (.xlsx)", type=["xlsx"])
    exception_file_out = st.file_uploader(
        "Upload Exception Orders (.xlsx) — 選填",
        type=["xlsx"],
        key="outbound_exception",
        help="檔案需含『Order』欄（會比對 DocNo）。命中的訂單會從 KPI/Failed 計算中排除，但會單獨統計筆數。",
    )

    if uploaded_file is not None:

        # ===== Load Data =====
        raw_df = load_data(uploaded_file)

        special_rule_dict = {}

        if special_rule_file is not None:
            special_df = load_special_rules(special_rule_file)
            special_rule_dict = build_special_rule_dict(special_df)

        exception_orders_out = (
            load_exception_orders(exception_file_out) if exception_file_out is not None else set()
        )

        # ===== Prepare KPI =====
        df_all = prepare_data(raw_df, special_rule_dict)

        if exception_orders_out:
            df_all["Is Exception"] = (
                df_all["DocNo"].astype(str).str.strip().str.upper().isin(exception_orders_out)
            )
        else:
            df_all["Is Exception"] = False

        excluded_total_out = int(df_all["Is Exception"].sum())

        df = df_all[~df_all["Is Exception"]].copy()
        summary_df = build_daily_kpi_summary(df)

        excluded_per_day = pd.DataFrame(columns=["Report Date", "excluded"])
        if excluded_total_out > 0:
            excl_src = df_all[df_all["Is Exception"]].copy()
            excl_src["Report Date"] = pd.to_datetime(excl_src["Committed Date"], errors="coerce").dt.date
            excluded_per_day = (
                excl_src.dropna(subset=["Report Date"])
                .groupby("Report Date").size().reset_index(name="excluded")
            )

        if not summary_df.empty or not excluded_per_day.empty:
            summary_df = pd.merge(summary_df, excluded_per_day, on="Report Date", how="outer").fillna(0)
            if "excluded" not in summary_df.columns:
                summary_df["excluded"] = 0
            summary_df["Report Date"] = pd.to_datetime(summary_df["Report Date"]).dt.date
            summary_df = summary_df.sort_values("Report Date").reset_index(drop=True)

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
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("Total Rows", f"{len(filtered_df):,}")
        col2.metric("Complete", f"{(filtered_df['Status Group'] == 'Complete').sum():,}")
        col3.metric("PGI", f"{(filtered_df['Status Group'] == 'PGI').sum():,}")
        col4.metric("Void", f"{(filtered_df['Status Group'] == 'Void').sum():,}")
        col5.metric("Excluded (Exception)", f"{excluded_total_out:,}",
                    help="由 Exception 檔指定排除、不納入 KPI 計算的訂單筆數。")

        # ===== 圖表 =====
        st.subheader("Outbound KPI Chart")
        fig_kpi = build_outbound_kpi_chart(filtered_summary_df)
        st.plotly_chart(fig_kpi, width="stretch")

        # ===== Summary =====
        rename_map = {
            "Report Date": "Date",
            "total_945": "945",
            "need_fulfill": "Need Fulfill",
            "in_kpi": "In KPI",
            "failed": "Failed",
            "kpi_rate": "KPI Rate",
            "excluded": "Excluded",
        }
        display_summary_df = filtered_summary_df.rename(columns=rename_map)

        if "KPI Rate" in display_summary_df.columns:
            display_summary_df["KPI Rate"] = display_summary_df["KPI Rate"].map(lambda x: f"{x:.2%}")
        if "Excluded" in display_summary_df.columns:
            display_summary_df["Excluded"] = display_summary_df["Excluded"].astype(int)

        preferred_order = ["Date", "945", "Need Fulfill", "In KPI", "Failed", "Excluded", "KPI Rate"]
        ordered_cols = [c for c in preferred_order if c in display_summary_df.columns] + [
            c for c in display_summary_df.columns if c not in preferred_order
        ]
        display_summary_df = display_summary_df[ordered_cols]

        st.subheader("Daily KPI Summary")
        st.dataframe(display_summary_df, width="stretch")

        if excluded_total_out > 0:
            st.subheader("Excluded Orders (Exception)")
            excl_view = df_all[df_all["Is Exception"]].copy()
            for c in ["Committed Date", "Created Date"]:
                if c in excl_view.columns:
                    excl_view[c] = pd.to_datetime(excl_view[c], errors="coerce").dt.date
            keep_cols = [c for c in ["DocNo", "TransNo", "Country code", "TransStatus",
                                     "Created Date", "Committed Date", "Status Group"]
                         if c in excl_view.columns]
            st.write(f"共 **{excluded_total_out:,}** 筆訂單被排除（不納入 KPI 計算）")
            st.dataframe(excl_view[keep_cols] if keep_cols else excl_view, width="stretch")

        # ===== 明細 =====
        st.subheader("Processed Data Preview")
        display_df = filtered_df.copy()

        date_cols = ["Committed Date", "Created Date", "Base Date", "KPI Failed Date"]

        for col in date_cols:
            if col in display_df.columns:
                display_df[col] = pd.to_datetime(display_df[col], errors="coerce").dt.date

        st.dataframe(display_df, width="stretch")

        # ===== Raw =====
        st.subheader("Raw Data Preview")
        st.dataframe(raw_df, width="stretch")

        # ===== 把 Outbound 結果存進 session_state，供 Email Report 分頁讀取 =====
        _ob_need_fulfill_total = int(filtered_summary_df["need_fulfill"].sum()) if "need_fulfill" in filtered_summary_df.columns else 0
        _ob_in_kpi_total = int(filtered_summary_df["in_kpi"].sum()) if "in_kpi" in filtered_summary_df.columns else 0
        _ob_failed_total = int(filtered_summary_df["failed"].sum()) if "failed" in filtered_summary_df.columns else 0
        _ob_total_945 = int(filtered_summary_df["total_945"].sum()) if "total_945" in filtered_summary_df.columns else 0
        _ob_avg_rate = (_ob_in_kpi_total / _ob_need_fulfill_total) if _ob_need_fulfill_total > 0 else 1.0
        st.session_state["outbound_report"] = {
            "summary_df": display_summary_df.copy(),
            "metrics": {
                "Need to Fulfill": _ob_need_fulfill_total,
                "In KPI": _ob_in_kpi_total,
                "Failed": _ob_failed_total,
                "Excluded": int(excluded_total_out),
                "Total 945": _ob_total_945,
                "Avg KPI Rate": f"{_ob_avg_rate:.2%}",
            },
            "chart_fig": fig_kpi,
            "date_range": (
                str(filtered_summary_df["Report Date"].min()) if not filtered_summary_df.empty else "",
                str(filtered_summary_df["Report Date"].max()) if not filtered_summary_df.empty else "",
            ),
        }

    else:
        st.info("Please upload your raw data Excel file to begin.")
with tab2:
    st.subheader("Inbound KPI Module")

    inbound_raw_file = st.file_uploader(
        "Upload Inbound Raw Data (.xlsx)",
        type=["xlsx"],
        key="inbound_raw"
    )

    inbound_mapping_file = st.file_uploader(
        "Upload Declaration Mapping File (.xlsx)",
        type=["xlsx"],
        key="inbound_mapping"
    )

    exception_file_in = st.file_uploader(
        "Upload Exception Orders (.xlsx) — 選填",
        type=["xlsx"],
        key="inbound_exception",
        help="檔案需含『Order』欄（會比對 INBSHIP No.）。命中的訂單會從 KPI/Failed 計算中排除，但會單獨統計筆數。",
    )

    if inbound_raw_file is not None and inbound_mapping_file is not None:

        # ===== Load Data =====
        raw_df = load_inbound_raw(inbound_raw_file)
        mapping_df = load_inbound_mapping(inbound_mapping_file)

        exception_orders_in = (
            load_exception_orders(exception_file_in) if exception_file_in is not None else set()
        )

        # ===== Prepare Data =====
        inbound_df = prepare_inbound_data(raw_df, mapping_df)

        if exception_orders_in and "INBSHIP No." in inbound_df.columns:
            inbound_df["Is Exception"] = (
                inbound_df["INBSHIP No."].astype(str).str.strip().str.upper().isin(exception_orders_in)
            )
        else:
            inbound_df["Is Exception"] = False

        excluded_total_in = int(inbound_df["Is Exception"].sum())
        inbound_df_all = inbound_df.copy()
        inbound_df = inbound_df[~inbound_df["Is Exception"]].copy()

        # ===== 日期篩選 =====
        inbound_df["Actual Date"] = pd.to_datetime(inbound_df["Date"], errors="coerce")
        inbound_df_all["Actual Date"] = pd.to_datetime(inbound_df_all["Date"], errors="coerce")

        if not inbound_df["Actual Date"].dropna().empty:
            min_date = inbound_df["Actual Date"].min().date()
            max_date = inbound_df["Actual Date"].max().date()

            date_range = st.date_input(
                "Select date range",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date,
                key="inbound_date"
            )

            if isinstance(date_range, tuple) and len(date_range) == 2:
                start_date, end_date = date_range

                filtered_inbound_df = inbound_df[
                    (inbound_df["Actual Date"].dt.date >= start_date) &
                    (inbound_df["Actual Date"].dt.date <= end_date)
                ].copy()
                filtered_inbound_all_df = inbound_df_all[
                    (inbound_df_all["Actual Date"].dt.date >= start_date) &
                    (inbound_df_all["Actual Date"].dt.date <= end_date)
                ].copy()
            else:
                filtered_inbound_df = inbound_df.copy()
                filtered_inbound_all_df = inbound_df_all.copy()
        else:
            filtered_inbound_df = inbound_df.copy()
            filtered_inbound_all_df = inbound_df_all.copy()

        excluded_inbound_df = filtered_inbound_all_df[filtered_inbound_all_df["Is Exception"]].copy()
        excluded_total_in_filtered = len(excluded_inbound_df)

        # ===== KPI Summary =====
        inbound_summary_df = build_inbound_kpi_summary(filtered_inbound_df)

        # ===== Failed item list =====
        failed_items_df = filtered_inbound_df[
            filtered_inbound_df["KPI Result"] == "Failed"
        ].copy()

        if not failed_items_df.empty:
            failed_items_df["Mapped Inbound Date"] = pd.to_datetime(
                failed_items_df["Mapped Inbound Date"], errors="coerce"
            ).dt.date

            failed_items_df["Date"] = pd.to_datetime(
                failed_items_df["Date"], errors="coerce"
            ).dt.date

            failed_items_display_df = failed_items_df[
                ["Vendor", "Mapped Inbound Date", "Date", "INBSHIP No.", "Declaration#", "PCS QTY"]
            ].rename(columns={
                "Mapped Inbound Date": "Inbound Date",
                "Date": "Complete Date"
            })
        else:
            failed_items_display_df = pd.DataFrame(
                columns=["Vendor", "Inbound Date", "Complete Date", "INBSHIP No.", "Declaration#", "PCS QTY"]
            )

        # ===== KPI 指標 =====
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("Total Rows", f"{len(filtered_inbound_df):,}")
        col2.metric("Matched", f"{(filtered_inbound_df['Mapping Status'] == 'Matched').sum():,}")
        col3.metric("Need Confirm", f"{(filtered_inbound_df['Mapping Status'] == '確認報單號碼').sum():,}")
        col4.metric("Failed", f"{(filtered_inbound_df['KPI Result'] == 'Failed').sum():,}")
        col5.metric("Excluded (Exception)", f"{excluded_total_in_filtered:,}",
                    help="由 Exception 檔指定排除、不納入 KPI 計算的進貨筆數。")

        # ===== KPI Chart =====
        st.subheader("Inbound KPI Chart")
        if not inbound_summary_df.empty:
            fig = build_inbound_kpi_chart(inbound_summary_df)
            st.plotly_chart(fig, width="stretch")
        else:
            st.info("No KPI data available for the selected date range.")

        # ===== Summary =====
        st.subheader("Inbound KPI Summary")
        if not inbound_summary_df.empty:
            display_inbound_summary_df = inbound_summary_df.copy()

            if not excluded_inbound_df.empty:
                excl_by_date = (
                    excluded_inbound_df.dropna(subset=["Actual Date"])
                    .assign(_d=lambda d: d["Actual Date"].dt.date)
                    .groupby("_d").size().reset_index(name="excluded")
                    .rename(columns={"_d": "Report Date"})
                )
                display_inbound_summary_df = pd.merge(
                    display_inbound_summary_df, excl_by_date,
                    on="Report Date", how="outer"
                ).fillna(0)
            else:
                display_inbound_summary_df["excluded"] = 0

            display_inbound_summary_df["kpi_rate"] = display_inbound_summary_df["kpi_rate"].map(
                lambda x: f"{x:.2%}" if pd.notna(x) and not isinstance(x, str) else x
            )
            display_inbound_summary_df = display_inbound_summary_df.rename(columns={
                "Report Date": "Date",
                "in_kpi": "In KPI",
                "failed": "Failed",
                "total": "Total",
                "kpi_rate": "KPI Rate",
                "excluded": "Excluded",
            })
            if "Excluded" in display_inbound_summary_df.columns:
                display_inbound_summary_df["Excluded"] = display_inbound_summary_df["Excluded"].astype(int)
            preferred_order = ["Date", "In KPI", "Failed", "Total", "Excluded", "KPI Rate"]
            ordered_cols = [c for c in preferred_order if c in display_inbound_summary_df.columns] + [
                c for c in display_inbound_summary_df.columns if c not in preferred_order
            ]
            display_inbound_summary_df = display_inbound_summary_df[ordered_cols].sort_values("Date")
            st.dataframe(display_inbound_summary_df, width="stretch")
        else:
            st.info("No summary data available for the selected date range.")

        if excluded_total_in_filtered > 0:
            st.subheader("Excluded Inbound Orders (Exception)")
            excl_view = excluded_inbound_df.copy()
            for c in ["Date", "Mapped Inbound Date", "Need Fulfill Date"]:
                if c in excl_view.columns:
                    excl_view[c] = pd.to_datetime(excl_view[c], errors="coerce").dt.date
            keep_cols = [c for c in ["Vendor", "Mapped Inbound Date", "Date",
                                     "INBSHIP No.", "Declaration#", "PCS QTY"]
                         if c in excl_view.columns]
            st.write(f"共 **{excluded_total_in_filtered:,}** 筆進貨被排除（不納入 KPI 計算）")
            st.dataframe(excl_view[keep_cols] if keep_cols else excl_view, width="stretch")

        # ===== Failed List =====
        st.subheader("Failed Item List")
        st.dataframe(failed_items_display_df, width="stretch")

        # ===== Need Confirm =====
        st.subheader("Need Confirm List")
        unmatched_df = filtered_inbound_df[
            filtered_inbound_df["Mapping Status"] == "確認報單號碼"
        ].copy()
        st.dataframe(unmatched_df, width="stretch")

        # ===== Processed Data =====
        st.subheader("Processed Data")
        display_df = filtered_inbound_df.copy()

        for col in ["Mapped Inbound Date", "Need Fulfill Date", "KPI Failed Date", "Actual Date"]:
            if col in display_df.columns:
                display_df[col] = pd.to_datetime(display_df[col], errors="coerce").dt.date

        st.dataframe(display_df, width="stretch")

        # ===== 把 Inbound 結果存進 session_state，供 Email Report 分頁讀取 =====
        _inbound_summary_for_email = (
            display_inbound_summary_df.copy()
            if ("display_inbound_summary_df" in dir() and isinstance(display_inbound_summary_df, pd.DataFrame))
            else pd.DataFrame()
        )
        _inbound_fig_for_email = fig if ("fig" in dir() and not inbound_summary_df.empty) else None
        _in_total = int(len(filtered_inbound_df))
        _in_in_kpi = int((filtered_inbound_df["KPI Result"] == "In KPI").sum())
        _in_failed = int((filtered_inbound_df["KPI Result"] == "Failed").sum())
        _in_avg_rate = (_in_in_kpi / (_in_in_kpi + _in_failed)) if (_in_in_kpi + _in_failed) > 0 else 1.0
        st.session_state["inbound_report"] = {
            "summary_df": _inbound_summary_for_email,
            "metrics": {
                "Total Inbound Shipment": _in_total,
                "In KPI": _in_in_kpi,
                "Failed": _in_failed,
                "Excluded": int(excluded_total_in_filtered),
                "Avg KPI Rate": f"{_in_avg_rate:.2%}",
            },
            "chart_fig": _inbound_fig_for_email,
            "failed_items_df": failed_items_display_df.copy() if isinstance(failed_items_display_df, pd.DataFrame) else pd.DataFrame(),
        }

    else:
        st.info("Please upload both files.")

with tab3:
    st.subheader("Warehouse Utilization Module")
    st.caption(
        "上傳「儲位總表」、「實際庫存使用報表」與「Item Master」後，系統會計算總儲位使用率與 RCK / LAR / SHF / MED 的個別使用率。\n"
        "Item Master 用來把 itemcode 對應到 Item Type，方便依品項類別觀察使用率；若不上傳則所有品項會標示為 Unknown，仍可看總攬。\n"
        "若實際使用報表中的儲位不在儲位總表內，將視為暫時撿貨儲位並從計算中略過。"
    )

    storage_master_file = st.file_uploader(
        "Upload Storage Master (儲位總表 .xlsx)",
        type=["xlsx"],
        key="storage_master",
    )

    storage_usage_files = st.file_uploader(
        "Upload Storage Usage Reports (實際庫存報表 .xlsx，可一次上傳多週)",
        type=["xlsx"],
        key="storage_usage",
        accept_multiple_files=True,
        help=(
            "可一次上傳多個檔案；系統會從檔名（例如 CR_WMSv3_Framework_Storage_Usage_Report_2026-05-20.xlsx）"
            "解析日期並轉成年度週數。同一週若有多個檔案，將以日期最晚的那筆為準。"
        ),
    )

    item_master_file = st.file_uploader(
        "Upload Item Master (Inventory 對照表 .xlsx / .csv) — 選填",
        type=["xlsx", "csv"],
        key="item_master",
        help="預期欄位：itemcode 與 Item Type（也接受 Framework PN / 種類 / 類別 等常見欄名）。每次重新整理頁面都需要重新上傳。",
    )

    # ===== Item Master 匯入範本下載 =====
    st.download_button(
        label="📥 下載 Item Master 匯入範本 (.xlsx)",
        data=build_item_master_template_bytes(),
        file_name="item_master_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="下載空白範本，內含正確欄名與幾筆示範資料；填好後即可上傳。",
        key="dl_item_master_template",
    )

    if storage_master_file is not None and storage_usage_files:

        # ===== Load Master =====
        master_df = load_storage_master(storage_master_file)

        # ===== 解析每個 usage 檔案的日期 / 週數 =====
        parsed_files = []
        files_without_date = []
        for f in storage_usage_files:
            d = parse_date_from_filename(getattr(f, "name", ""))
            if d is None:
                files_without_date.append(getattr(f, "name", "(unknown)"))
            else:
                parsed_files.append((f, d))

        if files_without_date:
            st.warning(
                "以下檔名無法解析日期，將被略過：\n- "
                + "\n- ".join(files_without_date)
                + "\n請確保檔名包含 YYYY-MM-DD 格式（例如 ..._2026-05-20.xlsx）。"
            )

        weekly_files = dedup_files_by_week(parsed_files)

        # 讀 Item Master（選填）
        if item_master_file is not None:
            item_master_dict = load_item_master(item_master_file)
            if not item_master_dict:
                st.warning(
                    "Item Master 上傳了但無法讀到 itemcode / Item Type 欄位，"
                    "請確認檔案格式（需要兩欄：itemcode、Item Type）。將以 Unknown 處理。"
                )
        else:
            item_master_dict = {}

        latest_entry = weekly_files[-1] if weekly_files else None
        usage_df = load_storage_usage(latest_entry["file"]) if latest_entry is not None else pd.DataFrame()

        if master_df.empty:
            st.error("無法從儲位總表讀到 Location/Type 欄位，請確認檔案格式。")
        elif not weekly_files:
            st.error("沒有任何 Storage Usage 檔案可分析，請確認檔名包含日期。")
        elif usage_df.empty:
            st.error("無法從實際庫存報表讀到 LocationCode 欄位，請確認檔案格式。")
        else:
            # ===== 套用 Item Master 為 usage_df 加上 Item Type =====
            usage_df = usage_df.copy()
            if "itemcode" in usage_df.columns:
                if item_master_dict:
                    usage_df["Item Type"] = usage_df["itemcode"].apply(
                        lambda x: get_item_type(x, item_master_dict)
                    )
                    matched_cnt = (usage_df["Item Type"] != "Unknown").sum()
                    st.caption(
                        f"Item Master 載入成功：共 **{len(item_master_dict):,}** 筆對照；"
                        f"實際使用報表中匹配 **{matched_cnt:,}** / {len(usage_df):,} 筆。"
                    )
                else:
                    usage_df["Item Type"] = "Unknown"
                    st.info("尚未上傳 Item Master，所有品項暫時標記為 Unknown。")
            else:
                # 若沒有 itemcode 欄位則整批標記為 Unknown，仍可顯示總攬
                usage_df["Item Type"] = "Unknown"
                st.warning("實際庫存報表中找不到 itemcode 欄位，無法依 Item Type 進一步分類。")

            # ===== Item Type 多選器（連動整個 tab3）=====
            st.markdown("### 觀察範圍篩選")
            usage_item_types_all = sorted(usage_df["Item Type"].dropna().unique())
            selected_item_types = st.multiselect(
                "選擇要分析的 Item Type (可複選；不選代表「全部」)",
                options=usage_item_types_all,
                default=usage_item_types_all,
                key="storage_item_type_multiselect",
                help="可選一或多個 Item Type；下方『每週變化量』與『依 Item Type 觀察』兩個區塊都會跟著連動。不選任何項目時，預設視為全部。",
            )

            if not selected_item_types:
                effective_item_types = usage_item_types_all
                is_full_scope = True
            else:
                effective_item_types = selected_item_types
                is_full_scope = set(selected_item_types) == set(usage_item_types_all)

            if is_full_scope:
                scope_label = "全部 (All Item Types)"
            elif len(effective_item_types) == 1:
                scope_label = f"Item Type = {effective_item_types[0]}"
            else:
                preview = ", ".join(effective_item_types[:3])
                tail = " ..." if len(effective_item_types) > 3 else ""
                scope_label = f"Item Type ({len(effective_item_types)} 種): {preview}{tail}"

            if is_full_scope:
                filtered_usage_df = usage_df.copy()
            else:
                filtered_usage_df = usage_df[usage_df["Item Type"].isin(effective_item_types)].copy()

            st.caption(f"目前 scope：**{scope_label}** ｜ 對應 itemcode 列數: {len(filtered_usage_df):,}")
            st.markdown("---")

            # ===== 每週變化趨勢 (跨週) =====
            st.markdown("### 每週變化量 (Weekly Trend)")
            weekly_info_rows = [
                {
                    "Year Week": w["week_label"],
                    "File Date": w["date"].strftime("%Y-%m-%d"),
                    "File Name": getattr(w["file"], "name", ""),
                }
                for w in weekly_files
            ]
            st.caption(
                f"共偵測到 **{len(weekly_files):,}** 週資料（同週多檔僅保留最晚日期）。"
                f"最新一週：**{weekly_files[-1]['week_label']}**（檔案日期 {weekly_files[-1]['date']:%Y-%m-%d}）。"
                f" 目前 Item Type scope：{scope_label}"
            )
            st.dataframe(pd.DataFrame(weekly_info_rows), width="stretch")

            available_types = list(STORAGE_TYPES) + sorted(set(master_df["Type"].unique()) - set(STORAGE_TYPES))
            selected_storage_types = st.multiselect(
                "選擇要觀察的儲位類型 (可複選)",
                options=available_types,
                default=list(STORAGE_TYPES),
                key="weekly_storage_type_multiselect",
                help="可選擇單一或多種儲位類型；圖表會即時連動，副座標的『合計使用率』也會依所選類型重新計算。",
            )

            if not selected_storage_types:
                st.info("請至少選擇一個儲位類型來顯示趨勢圖。")
                weekly_trend_df = pd.DataFrame(
                    columns=["Week Label", "Date", "Type", "Total", "Used", "Empty", "Utilization"]
                )
            else:
                weekly_rows = []
                _weekly_data_raw = []  # 為 Email Report 客製化保留每週「未過濾」的 usage_df
                for entry in weekly_files:
                    wk_usage_df_full = load_storage_usage(entry["file"]).copy()
                    if "itemcode" in wk_usage_df_full.columns and item_master_dict:
                        wk_usage_df_full["Item Type"] = wk_usage_df_full["itemcode"].apply(
                            lambda x: get_item_type(x, item_master_dict)
                        )
                    else:
                        wk_usage_df_full["Item Type"] = "Unknown"

                    # 保留未過濾版本供 tab4 重新計算
                    _weekly_data_raw.append({
                        "week_label": entry["week_label"],
                        "date": entry["date"],
                        "usage_df": wk_usage_df_full.copy(),
                    })

                    # 套用目前 scope 過濾
                    if not is_full_scope:
                        wk_usage_df = wk_usage_df_full[wk_usage_df_full["Item Type"].isin(effective_item_types)]
                    else:
                        wk_usage_df = wk_usage_df_full

                    wk_summary_df, _ovw, _unm, _used = build_utilization_summary(master_df, wk_usage_df)
                    for _, row in wk_summary_df.iterrows():
                        weekly_rows.append({
                            "Week Label": entry["week_label"],
                            "Date": entry["date"],
                            "Type": row["Type"],
                            "Total": int(row["Total"]),
                            "Used": int(row["Used"]),
                            "Empty": int(row["Empty"]),
                            "Utilization": float(row["Utilization"]),
                        })
                weekly_trend_df = pd.DataFrame(weekly_rows)

                # 存 raw 資料到 session_state，讓 Email Report 分頁可以重新依 Item Type 計算
                st.session_state["storage_raw"] = {
                    "master_df": master_df.copy(),
                    "all_item_types": list(usage_item_types_all),
                    "weekly_data": _weekly_data_raw,
                    "STORAGE_TYPES": list(STORAGE_TYPES),
                }

                fig_weekly = build_weekly_trend_chart(weekly_trend_df, selected_storage_types)
                st.plotly_chart(fig_weekly, width="stretch")

                if not weekly_trend_df.empty:
                    pivot_used = (
                        weekly_trend_df[weekly_trend_df["Type"].isin(selected_storage_types)]
                        .pivot_table(index="Week Label", columns="Type", values="Used", aggfunc="sum")
                        .fillna(0).astype(int)
                    )
                    week_order = (
                        weekly_trend_df[["Week Label", "Date"]]
                        .drop_duplicates().sort_values("Date")["Week Label"].tolist()
                    )
                    pivot_used = pivot_used.reindex(week_order)

                    rate_per_week = []
                    for wk in week_order:
                        wk_sub = weekly_trend_df[
                            (weekly_trend_df["Week Label"] == wk)
                            & (weekly_trend_df["Type"].isin(selected_storage_types))
                        ]
                        total = wk_sub["Total"].sum()
                        used = wk_sub["Used"].sum()
                        rate_per_week.append(used / total if total > 0 else 0.0)
                    pivot_used["Total Used"] = pivot_used.sum(axis=1)
                    pivot_used["Utilization"] = rate_per_week

                    diff = pivot_used[["Total Used"]].diff().rename(columns={"Total Used": "WoW Δ Used"})
                    util_diff = pivot_used[["Utilization"]].diff().rename(columns={"Utilization": "WoW Δ Util"})
                    summary_table = pivot_used.join(diff).join(util_diff)

                    display_table = summary_table.copy()
                    display_table["Utilization"] = display_table["Utilization"].map(lambda x: f"{x:.2%}")
                    display_table["WoW Δ Util"] = display_table["WoW Δ Util"].map(
                        lambda x: "" if pd.isna(x) else f"{x:+.2%}"
                    )
                    display_table["WoW Δ Used"] = display_table["WoW Δ Used"].map(
                        lambda x: "" if pd.isna(x) else f"{int(x):+d}"
                    )

                    st.markdown("#### 每週各類型 Used 數量與週對週變化")
                    st.dataframe(display_table, width="stretch")

            st.markdown("---")

            # ===== 總體 (不分 Item Type) 使用率：作為「總攬」 =====
            type_summary_all_df, overall_summary_all, unmatched_df, used_locations_all = (
                build_utilization_summary(master_df, usage_df)
            )

            st.markdown(f"### 總攬 (All Item Types) — 最新週: {weekly_files[-1]['week_label']}")

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("總儲位數", f"{overall_summary_all['total']:,}")
            col2.metric("已使用", f"{overall_summary_all['used']:,}")
            col3.metric("空儲位", f"{overall_summary_all['empty']:,}")
            col4.metric("總使用率", f"{overall_summary_all['utilization']:.2%}")

            # 總攬：每個 Location Type 的使用率指標
            type_cols = st.columns(len(STORAGE_TYPES))
            for i, t in enumerate(STORAGE_TYPES):
                row = type_summary_all_df[type_summary_all_df["Type"] == t]
                if not row.empty:
                    r = row.iloc[0]
                    type_cols[i].metric(
                        label=f"{t}  ({int(r['Used'])}/{int(r['Total'])})",
                        value=f"{r['Utilization']:.2%}",
                        delta=f"Empty: {int(r['Empty'])}",
                        delta_color="off",
                    )
                else:
                    type_cols[i].metric(label=t, value="N/A")

            # ===== Item Type × Location Type Breakdown (總攬交叉分析) =====
            st.markdown("#### Item Type × Location Type 使用儲位數")
            st.caption("每一格代表「該 Item Type 在該 Location Type 中所佔用的儲位數量」。")

            # 用 master_df 取得 Location → Type 對照
            loc_to_type = dict(zip(master_df["Location"], master_df["Type"]))

            # 排除暫時撿貨儲位
            usage_in_master_df = usage_df[usage_df["LocationCode"].isin(set(master_df["Location"]))].copy()
            usage_in_master_df["Loc Type"] = usage_in_master_df["LocationCode"].map(loc_to_type)

            # 每個 (Item Type, Loc Type) 對應到的「不重複 Location 數量」
            breakdown_long = (
                usage_in_master_df.dropna(subset=["Loc Type"])
                .groupby(["Item Type", "Loc Type"])["LocationCode"]
                .nunique()
                .reset_index(name="Used Locations")
            )

            if not breakdown_long.empty:
                # 透視成矩陣：列 = Item Type, 欄 = Location Type
                ordered_loc_types = STORAGE_TYPES + sorted(
                    set(breakdown_long["Loc Type"]) - set(STORAGE_TYPES)
                )
                breakdown_pivot = (
                    breakdown_long.pivot(index="Item Type", columns="Loc Type", values="Used Locations")
                    .fillna(0)
                    .astype(int)
                )
                # 重排欄位順序
                breakdown_pivot = breakdown_pivot.reindex(
                    columns=[c for c in ordered_loc_types if c in breakdown_pivot.columns]
                )
                breakdown_pivot["Total"] = breakdown_pivot.sum(axis=1)
                breakdown_pivot = breakdown_pivot.sort_values("Total", ascending=False)
                st.dataframe(breakdown_pivot, width="stretch")
            else:
                st.info("沒有可顯示的 Item Type × Location Type 資料。")

            # ===== 依 Item Type 觀察（使用上方多選器的選擇）=====
            st.markdown("---")
            st.markdown("### 依 Item Type 觀察")

            if is_full_scope:
                type_summary_df = type_summary_all_df
                overall_summary = overall_summary_all
                used_locations = used_locations_all
            else:
                type_summary_df, overall_summary, _unmatched_ignore, used_locations = (
                    build_utilization_summary(master_df, filtered_usage_df)
                )

            st.markdown(f"**Scope**: {scope_label} ｜ 對應 itemcode 列數: {len(filtered_usage_df):,}")

            # ===== 篩選後的總體指標 =====
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("總儲位數", f"{overall_summary['total']:,}")
            col2.metric(
                "該 Item Type 佔用儲位",
                f"{overall_summary['used']:,}",
            )
            col3.metric(
                "該 Item Type 未佔用儲位",
                f"{overall_summary['empty']:,}",
            )
            col4.metric(
                "佔用率",
                f"{overall_summary['utilization']:.2%}",
                help="該 Item Type 佔用的儲位數 / 倉庫總儲位數",
            )

            # ===== Per-Location-Type 指標 (RCK / LAR / SHF / MED) =====
            st.subheader(f"Utilization by Location Type — {scope_label}")
            type_cols = st.columns(len(STORAGE_TYPES))
            for i, t in enumerate(STORAGE_TYPES):
                row = type_summary_df[type_summary_df["Type"] == t]
                if not row.empty:
                    r = row.iloc[0]
                    type_cols[i].metric(
                        label=f"{t}  ({int(r['Used'])}/{int(r['Total'])})",
                        value=f"{r['Utilization']:.2%}",
                        delta=f"Empty: {int(r['Empty'])}",
                        delta_color="off",
                    )
                else:
                    type_cols[i].metric(label=t, value="N/A")

            # ===== 圖表 =====
            st.subheader("Utilization Chart")
            fig_util = build_utilization_chart(type_summary_df)
            st.plotly_chart(fig_util, width="stretch")

            # ===== Summary 表 =====
            st.subheader("Utilization Summary Table")
            display_summary = type_summary_df.copy()
            display_summary["Utilization"] = display_summary["Utilization"].map(lambda x: f"{x:.2%}")
            st.dataframe(display_summary, width="stretch")

            # ===== 空儲位明細 (相對於目前 scope) =====
            st.subheader("Empty Locations (該 scope 下未被佔用的儲位)")
            empty_detail_df = master_df[~master_df["Location"].isin(used_locations)].copy()
            empty_detail_df = empty_detail_df.sort_values(["Type", "Location"]).reset_index(drop=True)
            st.write(f"共 **{len(empty_detail_df):,}** 個空儲位")
            st.dataframe(empty_detail_df, width="stretch")

            # ===== 已使用儲位明細 (含庫存量, 套用 Item Type 篩選) =====
            st.subheader(f"Used Locations — {scope_label}")
            usage_agg_cols = [c for c in ["Storage", "Staging", "Defective"] if c in filtered_usage_df.columns]
            if usage_agg_cols:
                usage_agg = (
                    filtered_usage_df.groupby("LocationCode")[usage_agg_cols]
                    .sum()
                    .reset_index()
                )
            else:
                usage_agg = filtered_usage_df[["LocationCode"]].drop_duplicates().reset_index(drop=True)

            used_detail_df = master_df[master_df["Location"].isin(used_locations)].merge(
                usage_agg, left_on="Location", right_on="LocationCode", how="left"
            )
            if "LocationCode" in used_detail_df.columns:
                used_detail_df = used_detail_df.drop(columns=["LocationCode"])
            used_detail_df = used_detail_df.sort_values(["Type", "Location"]).reset_index(drop=True)
            st.dataframe(used_detail_df, width="stretch")

            # ===== 暫時撿貨儲位 (在 usage 但不在 master)，固定使用全量 usage =====
            st.subheader("Temporary Picking Locations (略過計算的儲位)")
            if unmatched_df.empty:
                st.info("沒有暫時撿貨儲位 — 所有實際使用的儲位都能對應到儲位總表。")
            else:
                unmatched_unique = (
                    unmatched_df["LocationCode"].drop_duplicates().sort_values().reset_index(drop=True)
                )
                st.write(f"共 **{len(unmatched_unique):,}** 個儲位不在儲位總表中（已從使用率計算中略過）")
                st.dataframe(unmatched_df, width="stretch")

            # ===== Download =====
            st.subheader("Export")
            excel_bytes = export_utilization_excel(
                type_summary_df=type_summary_df,
                overall_summary=overall_summary,
                used_detail_df=used_detail_df,
                empty_detail_df=empty_detail_df,
                unmatched_df=unmatched_df,
            )
            if is_full_scope:
                file_name_scope = "all"
            elif len(effective_item_types) == 1:
                file_name_scope = effective_item_types[0].lower().replace(" ", "_")
            else:
                file_name_scope = f"{len(effective_item_types)}_types"
            st.download_button(
                label=f"Download Utilization Report ({scope_label}) (.xlsx)",
                data=excel_bytes,
                file_name=f"warehouse_utilization_{file_name_scope}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # ===== 把 Storage 結果存進 session_state，供 Email Report 分頁讀取 =====
            _per_type_for_email = type_summary_all_df.copy()
            if not _per_type_for_email.empty and "Utilization" in _per_type_for_email.columns:
                _per_type_for_email["Utilization"] = _per_type_for_email["Utilization"].map(
                    lambda x: f"{x:.2%}" if pd.notna(x) else ""
                )
            st.session_state["storage_report"] = {
                "overall": overall_summary_all,
                "latest_week": weekly_files[-1]["week_label"],
                "latest_date": weekly_files[-1]["date"].strftime("%Y-%m-%d"),
                "per_type_df": _per_type_for_email,
                "weekly_trend_df": weekly_trend_df.copy() if isinstance(weekly_trend_df, pd.DataFrame) else pd.DataFrame(),
                "weekly_fig": fig_weekly if "fig_weekly" in dir() else None,
                "scope_label": scope_label,
            }

    else:
        st.info("請同時上傳「儲位總表」與一個（或多個）「實際庫存報表」以開始計算。")
# =====================================================================
#  Email Report Tab
#  整合三個分頁的 KPI / 圖表為一份可直接在 Outlook 信件內文觀看的報告
# =====================================================================

import base64 as _b64
import mimetypes as _mt
from datetime import datetime as _dt
from email.mime.multipart import MIMEMultipart as _MIMEMultipart
from email.mime.text import MIMEText as _MIMEText
from email.mime.image import MIMEImage as _MIMEImage
from email.utils import make_msgid as _make_msgid, formatdate as _formatdate

DEFAULT_EMAIL_TO = "joshua_j_liu@dimerco.com"
DEFAULT_EMAIL_CC = ""
DEFAULT_EMAIL_SUBJECT = "[Framework] Weekly Outbound / Inbound / Warehouse KPI Report"

# ====== Excel-style 表格樣式 (Outlook Word engine 友善：全 inline styles, table-based) ======
_TABLE_STYLE = (
    "border-collapse:collapse;border-spacing:0;"
    "font-family:'Segoe UI',Arial,Helvetica,sans-serif;font-size:13px;"
    "width:100%;margin:8px 0 14px;"
    "border:1px solid #b7c3cf;"
)
_TH_STYLE = (
    "background:#1f4e78;color:#ffffff;padding:8px 10px;"
    "border:1px solid #1f4e78;text-align:left;font-weight:600;"
)
_TD_STYLE = "padding:6px 10px;border:1px solid #d0d7de;color:#1f2937;"
_TD_NUM_STYLE = _TD_STYLE + "text-align:right;font-variant-numeric:tabular-nums;"
_ROW_ALT = "#f4f7fb"


def _td(value, num=False, bg=None, extra=""):
    """產生 <td> HTML，支援數字靠右對齊與背景色。"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        text = ""
    elif isinstance(value, float):
        text = f"{value:,.2f}"
    elif isinstance(value, int):
        text = f"{value:,}"
    else:
        text = str(value)
    style = _TD_NUM_STYLE if num else _TD_STYLE
    if bg:
        style += f"background:{bg};"
    if extra:
        style += extra
    return f'<td style="{style}">{text}</td>'


def _excel_table_html(df, max_rows=200, numeric_cols=None):
    """把 DataFrame 渲染成 Excel 樣式 HTML 表格 (Outlook 相容)。"""
    if df is None or len(df) == 0:
        return '<p style="color:#9ca3af;font-style:italic;">（無資料）</p>'
    df = df.head(max_rows)
    if numeric_cols is None:
        numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    cols = [str(c) for c in df.columns]
    head = "".join(f'<th style="{_TH_STYLE}">{c}</th>' for c in cols)
    rows_html = []
    for i, (_, row) in enumerate(df.iterrows()):
        bg = _ROW_ALT if i % 2 else "#ffffff"
        cells = "".join(
            _td(row[c], num=(c in numeric_cols), bg=bg) for c in df.columns
        )
        rows_html.append(f"<tr>{cells}</tr>")
    return (
        f'<table style="{_TABLE_STYLE}">'
        f'<thead><tr>{head}</tr></thead>'
        f'<tbody>{"".join(rows_html)}</tbody>'
        f"</table>"
    )


def _bar_cell_html(value_pct, color="#1f77b4", width_px=180, label=None):
    """畫一格內嵌的水平 bar (用 nested table，Outlook 相容)。
    value_pct 介於 0~1。label 為 bar 內文字。"""
    try:
        v = max(0.0, min(1.0, float(value_pct)))
    except Exception:
        v = 0.0
    used_w = int(round(v * width_px))
    empty_w = max(width_px - used_w, 0)
    label_html = ""
    if label:
        label_html = (
            f'<span style="position:relative;left:6px;color:#ffffff;'
            f'font-size:11px;font-weight:bold;">{label}</span>'
        )
    return (
        f'<table style="border-collapse:collapse;border-spacing:0;'
        f'width:{width_px}px;height:18px;border:1px solid #b7c3cf;background:#eef2f7;">'
        f'<tr>'
        f'<td style="width:{used_w}px;height:18px;background:{color};padding:0;line-height:18px;">'
        f'{label_html}</td>'
        f'<td style="width:{empty_w}px;height:18px;background:#eef2f7;padding:0;"></td>'
        f"</tr></table>"
    )


# ====== 三個區塊各自的 HTML 圖表 (純表格，Outlook 相容) ======

def _to_rate_float(v):
    """把任意輸入 (數字 / 字串 / NaN) 統一轉成 0~1 的使用率 float。
    支援：1.0 / 0.95 / "100.00%" / "95.50%" / "0.95" / None / NaN
    回傳介於 0.0 ~ 1.0 (不會超出範圍)。
    """
    try:
        if v is None:
            return 0.0
        if isinstance(v, (int, float)):
            if pd.isna(v):
                return 0.0
            f = float(v)
        else:
            s = str(v).replace("%", "").replace(",", "").strip()
            if not s or s.lower() == "nan":
                return 0.0
            f = float(s)
        # 若數值 > 1.0001，視為百分比 (e.g. 95.5 來自 "95.50%")
        if f > 1.0001:
            f = f / 100.0
        # 限制範圍
        if f < 0:
            f = 0.0
        elif f > 1.0:
            f = 1.0
        return f
    except (ValueError, TypeError):
        return 0.0


def _to_int_safe(v, default=0):
    """穩健地把任意輸入轉成 int。"""
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return default
        if isinstance(v, (int, float)):
            return int(v)
        s = str(v).replace(",", "").strip()
        if not s or s.lower() == "nan":
            return default
        return int(float(s))
    except (ValueError, TypeError):
        return default


def _chart_outbound_html(summary_df):
    """Outbound: 顯示每日 945 / Need Fulfill / In KPI / Failed / KPI Rate
    以 stacked-bar 風格呈現 (In KPI 藍 + Failed 紅)。"""
    if summary_df is None or summary_df.empty:
        return '<p style="color:#9ca3af;font-style:italic;">（無圖表資料）</p>'
    df = summary_df.copy()
    # 統一欄名
    rename = {
        "Report Date": "Date", "total_945": "945", "need_fulfill": "Need Fulfill",
        "in_kpi": "In KPI", "failed": "Failed", "kpi_rate": "KPI Rate",
        "excluded": "Excluded",
    }
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})
    if "Date" not in df.columns:
        return '<p style="color:#9ca3af;font-style:italic;">（無圖表資料）</p>'

    rows_html = []
    # 用 helper 安全取得每列 Need Fulfill 最大值
    need_vals = [_to_int_safe(v) for v in df.get("Need Fulfill", pd.Series([1]))]
    max_need = max(need_vals) if need_vals else 1
    if max_need <= 0:
        max_need = 1
    for i, (_, row) in enumerate(df.iterrows()):
        bg = _ROW_ALT if i % 2 else "#ffffff"
        need = _to_int_safe(row.get("Need Fulfill"))
        in_kpi = _to_int_safe(row.get("In KPI"))
        failed = _to_int_safe(row.get("Failed"))
        rate = _to_rate_float(row.get("KPI Rate"))
        # bar: 整條長度按 need / max_need；內部分 in_kpi / failed
        bar_total_w = 220
        bar_w = int(round((need / max_need) * bar_total_w)) if max_need else 0
        if need > 0:
            in_w = int(round((in_kpi / need) * bar_w))
            fail_w = max(bar_w - in_w, 0)
        else:
            in_w = fail_w = 0
        empty_w = max(bar_total_w - bar_w, 0)
        bar_html = (
            f'<table style="border-collapse:collapse;border-spacing:0;'
            f'width:{bar_total_w}px;height:16px;border:1px solid #b7c3cf;background:#eef2f7;">'
            f'<tr>'
            f'<td style="width:{in_w}px;height:16px;background:#1f77b4;padding:0;"></td>'
            f'<td style="width:{fail_w}px;height:16px;background:#d62728;padding:0;"></td>'
            f'<td style="width:{empty_w}px;height:16px;background:#eef2f7;padding:0;"></td>'
            f'</tr></table>'
        )
        rate_color = "#16a34a" if rate >= 0.95 else ("#eab308" if rate >= 0.85 else "#dc2626")
        excluded = _to_int_safe(row.get("Excluded"))
        cells = (
            _td(row.get("Date"), bg=bg)
            + _td(row.get("945", 0), num=True, bg=bg)
            + _td(need, num=True, bg=bg)
            + _td(in_kpi, num=True, bg=bg)
            + _td(failed, num=True, bg=bg)
            + _td(excluded, num=True, bg=bg)
            + f'<td style="{_TD_NUM_STYLE}background:{bg};color:{rate_color};font-weight:600;">{rate:.2%}</td>'
            + f'<td style="{_TD_STYLE}background:{bg};">{bar_html}</td>'
        )
        rows_html.append(f"<tr>{cells}</tr>")
    head_cols = ["Date", "945", "Need Fulfill", "In KPI", "Failed", "Excluded", "KPI Rate", "Viz (In KPI / Failed)"]
    head = "".join(f'<th style="{_TH_STYLE}">{c}</th>' for c in head_cols)
    legend = (
        '<div style="font-size:11px;color:#6b7280;margin-bottom:4px;">'
        '<span style="display:inline-block;width:10px;height:10px;background:#1f77b4;margin-right:4px;"></span>In KPI'
        '<span style="display:inline-block;width:10px;height:10px;background:#d62728;margin:0 4px 0 12px;"></span>Failed</div>'
    )
    return (
        legend
        + f'<table style="{_TABLE_STYLE}"><thead><tr>{head}</tr></thead>'
        + f'<tbody>{"".join(rows_html)}</tbody></table>'
    )


def _chart_inbound_html(summary_df):
    """Inbound: 顯示每日 In KPI / Failed / Total / KPI Rate 與 stacked bar。"""
    if summary_df is None or summary_df.empty:
        return '<p style="color:#9ca3af;font-style:italic;">（無圖表資料）</p>'
    df = summary_df.copy()
    rename = {
        "Report Date": "Date", "in_kpi": "In KPI", "failed": "Failed",
        "total": "Total", "kpi_rate": "KPI Rate", "excluded": "Excluded",
    }
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})
    if "Date" not in df.columns:
        return '<p style="color:#9ca3af;font-style:italic;">（無圖表資料）</p>'

    rows_html = []
    total_vals = [_to_int_safe(v) for v in df.get("Total", pd.Series([1]))]
    max_total = max(total_vals) if total_vals else 1
    if max_total <= 0:
        max_total = 1
    for i, (_, row) in enumerate(df.iterrows()):
        bg = _ROW_ALT if i % 2 else "#ffffff"
        total = _to_int_safe(row.get("Total"))
        in_kpi = _to_int_safe(row.get("In KPI"))
        failed = _to_int_safe(row.get("Failed"))
        rate = _to_rate_float(row.get("KPI Rate"))
        bar_total_w = 220
        bar_w = int(round((total / max_total) * bar_total_w)) if max_total else 0
        if total > 0:
            in_w = int(round((in_kpi / total) * bar_w))
            fail_w = max(bar_w - in_w, 0)
        else:
            in_w = fail_w = 0
        empty_w = max(bar_total_w - bar_w, 0)
        bar_html = (
            f'<table style="border-collapse:collapse;border-spacing:0;'
            f'width:{bar_total_w}px;height:16px;border:1px solid #b7c3cf;background:#eef2f7;">'
            f'<tr>'
            f'<td style="width:{in_w}px;height:16px;background:#1f77b4;padding:0;"></td>'
            f'<td style="width:{fail_w}px;height:16px;background:#d62728;padding:0;"></td>'
            f'<td style="width:{empty_w}px;height:16px;background:#eef2f7;padding:0;"></td>'
            f'</tr></table>'
        )
        rate_color = "#16a34a" if rate >= 0.95 else ("#eab308" if rate >= 0.85 else "#dc2626")
        excluded = _to_int_safe(row.get("Excluded"))
        cells = (
            _td(row.get("Date"), bg=bg)
            + _td(in_kpi, num=True, bg=bg)
            + _td(failed, num=True, bg=bg)
            + _td(total, num=True, bg=bg)
            + _td(excluded, num=True, bg=bg)
            + f'<td style="{_TD_NUM_STYLE}background:{bg};color:{rate_color};font-weight:600;">{rate:.2%}</td>'
            + f'<td style="{_TD_STYLE}background:{bg};">{bar_html}</td>'
        )
        rows_html.append(f"<tr>{cells}</tr>")
    head_cols = ["Date", "In KPI", "Failed", "Total", "Excluded", "KPI Rate", "Viz (In KPI / Failed)"]
    head = "".join(f'<th style="{_TH_STYLE}">{c}</th>' for c in head_cols)
    legend = (
        '<div style="font-size:11px;color:#6b7280;margin-bottom:4px;">'
        '<span style="display:inline-block;width:10px;height:10px;background:#1f77b4;margin-right:4px;"></span>In KPI'
        '<span style="display:inline-block;width:10px;height:10px;background:#d62728;margin:0 4px 0 12px;"></span>Failed</div>'
    )
    return (
        legend
        + f'<table style="{_TABLE_STYLE}"><thead><tr>{head}</tr></thead>'
        + f'<tbody>{"".join(rows_html)}</tbody></table>'
    )


def _chart_storage_weekly_html(weekly_trend_df, selected_types=None):
    """Storage 累加歷史：每週 Total / Used / Empty / Util% + WoW Δ + stacked bar"""
    if weekly_trend_df is None or weekly_trend_df.empty:
        return '<p style="color:#9ca3af;font-style:italic;">（無圖表資料）</p>'
    df = weekly_trend_df.copy()
    if selected_types:
        df = df[df["Type"].isin(selected_types)]
    # 依週彙總
    agg = df.groupby(["Week Label", "Date"], as_index=False).agg(
        Total=("Total", "sum"), Used=("Used", "sum")
    )
    agg["Empty"] = agg["Total"] - agg["Used"]
    agg["Utilization"] = agg.apply(
        lambda r: (r["Used"] / r["Total"]) if r["Total"] > 0 else 0.0, axis=1
    )
    agg = agg.sort_values("Date").reset_index(drop=True)
    agg["WoW_Used"] = agg["Used"].diff()
    agg["WoW_Util"] = agg["Utilization"].diff()

    rows_html = []
    for i, row in agg.iterrows():
        bg = _ROW_ALT if i % 2 else "#ffffff"
        total = int(row["Total"])
        used = int(row["Used"])
        empty = int(row["Empty"])
        util = float(row["Utilization"])
        # bar: Used (藍) + Empty (淺灰)
        bar_total_w = 240
        used_w = int(round(util * bar_total_w))
        empty_w = bar_total_w - used_w
        bar_html = (
            f'<table style="border-collapse:collapse;border-spacing:0;'
            f'width:{bar_total_w}px;height:18px;border:1px solid #b7c3cf;">'
            f'<tr>'
            f'<td style="width:{used_w}px;height:18px;background:#1f77b4;padding:0;text-align:center;'
            f'color:#ffffff;font-weight:700;font-size:11px;line-height:18px;">{util:.2%}</td>'
            f'<td style="width:{empty_w}px;height:18px;background:#d3d3d3;padding:0;"></td>'
            f'</tr></table>'
        )
        wow_used = row.get("WoW_Used")
        wow_util = row.get("WoW_Util")
        wow_used_text = "" if pd.isna(wow_used) else (f"+{int(wow_used):,}" if wow_used >= 0 else f"{int(wow_used):,}")
        wow_util_text = "" if pd.isna(wow_util) else (f"+{wow_util:.2%}" if wow_util >= 0 else f"{wow_util:.2%}")
        wow_used_color = "#16a34a" if (not pd.isna(wow_used) and wow_used >= 0) else "#dc2626"
        wow_util_color = "#16a34a" if (not pd.isna(wow_util) and wow_util >= 0) else "#dc2626"
        cells = (
            _td(row["Week Label"], bg=bg)
            + _td(total, num=True, bg=bg)
            + _td(used, num=True, bg=bg)
            + _td(empty, num=True, bg=bg)
            + f'<td style="{_TD_NUM_STYLE}background:{bg};font-weight:600;">{util:.2%}</td>'
            + f'<td style="{_TD_NUM_STYLE}background:{bg};color:{wow_used_color};">{wow_used_text}</td>'
            + f'<td style="{_TD_NUM_STYLE}background:{bg};color:{wow_util_color};">{wow_util_text}</td>'
            + f'<td style="{_TD_STYLE}background:{bg};">{bar_html}</td>'
        )
        rows_html.append(f"<tr>{cells}</tr>")

    head_cols = ["Year Week", "Total", "Used", "Empty", "Util%", "WoW Δ Used", "WoW Δ Util%", "視覺化 Utilization"]
    head = "".join(f'<th style="{_TH_STYLE}">{c}</th>' for c in head_cols)
    legend = (
        '<div style="font-size:11px;color:#6b7280;margin-bottom:4px;">'
        '<span style="display:inline-block;width:10px;height:10px;background:#1f77b4;margin-right:4px;"></span>Used'
        '<span style="display:inline-block;width:10px;height:10px;background:#d3d3d3;margin:0 4px 0 12px;"></span>Empty</div>'
    )
    return (
        legend
        + f'<table style="{_TABLE_STYLE}"><thead><tr>{head}</tr></thead>'
        + f'<tbody>{"".join(rows_html)}</tbody></table>'
    )


def _chart_storage_by_type_html(weekly_trend_df, selected_types=None):
    """每個 Storage Type 各週 Util% 變化 + 最新一週的 WoW Δ Used / Δ Util%。"""
    if weekly_trend_df is None or weekly_trend_df.empty:
        return '<p style="color:#9ca3af;font-style:italic;">（無歷史資料）</p>'
    df = weekly_trend_df.copy()
    if selected_types:
        df = df[df["Type"].isin(selected_types)]
    if df.empty:
        return '<p style="color:#9ca3af;font-style:italic;">（無歷史資料）</p>'
    week_order = (
        df[["Week Label", "Date"]].drop_duplicates().sort_values("Date")["Week Label"].tolist()
    )
    pivot_util = (
        df.pivot_table(index="Type", columns="Week Label", values="Utilization", aggfunc="mean")
        .reindex(columns=week_order)
        .fillna(0)
    )
    pivot_used = (
        df.pivot_table(index="Type", columns="Week Label", values="Used", aggfunc="sum")
        .reindex(columns=week_order)
        .fillna(0)
        .astype(int)
    )
    has_wow = len(week_order) >= 2
    head_cols = ["Type"] + list(pivot_util.columns)
    if has_wow:
        head_cols += ["WoW Δ Used", "WoW Δ Util%"]
    head_cols += ["Latest Week"]
    head = "".join(f'<th style="{_TH_STYLE}">{c}</th>' for c in head_cols)
    rows_html = []
    for i, typ in enumerate(pivot_util.index):
        bg = _ROW_ALT if i % 2 else "#ffffff"
        srow_util = pivot_util.loc[typ]
        srow_used = pivot_used.loc[typ]
        cells = _td(typ, bg=bg)
        for c in pivot_util.columns:
            v = float(srow_util[c])
            cells += f'<td style="{_TD_NUM_STYLE}background:{bg};">{v:.2%}</td>'
        if has_wow:
            wow_used = int(srow_used.iloc[-1]) - int(srow_used.iloc[-2])
            wow_util = float(srow_util.iloc[-1]) - float(srow_util.iloc[-2])
            used_color = "#16a34a" if wow_used >= 0 else "#dc2626"
            util_color = "#16a34a" if wow_util >= 0 else "#dc2626"
            cells += f'<td style="{_TD_NUM_STYLE}background:{bg};color:{used_color};">{wow_used:+,}</td>'
            cells += f'<td style="{_TD_NUM_STYLE}background:{bg};color:{util_color};">{wow_util:+.2%}</td>'
        latest_v = float(srow_util.iloc[-1]) if len(srow_util) else 0.0
        cells += f'<td style="{_TD_STYLE}background:{bg};">{_bar_cell_html(latest_v, color="#1f77b4", width_px=160, label=f"{latest_v:.1%}")}</td>'
        rows_html.append(f"<tr>{cells}</tr>")
    return (
        f'<table style="{_TABLE_STYLE}">'
        f'<thead><tr>{head}</tr></thead>'
        f'<tbody>{"".join(rows_html)}</tbody>'
        f"</table>"
    )


def _compute_storage_trend(storage_raw, selected_item_types=None):
    """從 storage_raw (master_df + 每週原始 usage_df) 重新計算指定 Item Type 子集的 storage trend。

    參數:
      selected_item_types: None 表示「全部 Item Types」；list 則只取對應 Item Type 的列。

    回傳 dict 包含 overall / weekly_trend_df / per_type_df / latest_week / latest_date。
    """
    if not storage_raw:
        return None
    master_df = storage_raw.get("master_df")
    weekly_data = storage_raw.get("weekly_data", [])
    if master_df is None or not weekly_data:
        return None

    weekly_rows = []
    latest_overall = None
    latest_per_type_df = None
    for entry in weekly_data:
        usage_df = entry["usage_df"]
        if selected_item_types is not None:
            usage_df = usage_df[usage_df["Item Type"].isin(selected_item_types)]
        type_summary_df, overall, _unm, _used = build_utilization_summary(master_df, usage_df)
        for _, row in type_summary_df.iterrows():
            weekly_rows.append({
                "Week Label": entry["week_label"],
                "Date": entry["date"],
                "Type": row["Type"],
                "Total": int(row["Total"]),
                "Used": int(row["Used"]),
                "Empty": int(row["Empty"]),
                "Utilization": float(row["Utilization"]),
            })
        latest_overall = overall
        latest_per_type_df = type_summary_df.copy()

    # 把 latest_per_type_df 的 Utilization 轉成百分比字串以便 Excel 表格顯示
    if latest_per_type_df is not None and "Utilization" in latest_per_type_df.columns:
        latest_per_type_df["Utilization"] = latest_per_type_df["Utilization"].map(
            lambda x: f"{x:.2%}" if pd.notna(x) else ""
        )
    return {
        "overall": latest_overall or {},
        "weekly_trend_df": pd.DataFrame(weekly_rows),
        "per_type_df": latest_per_type_df if latest_per_type_df is not None else pd.DataFrame(),
        "latest_week": weekly_data[-1]["week_label"],
        "latest_date": weekly_data[-1]["date"].strftime("%Y-%m-%d"),
    }


def _render_storage_subsection(report_data, header_text):
    """渲染一個 Storage 子區塊 (All Item Types 或 Custom)。"""
    if not report_data:
        return ""
    parts = []
    parts.append(_sub_header(header_text))
    overall = report_data.get("overall", {})
    latest_wk = report_data.get("latest_week", "")
    latest_dt = report_data.get("latest_date", "")
    parts.append(
        f'<div style="color:#6b7280;font-size:12px;margin:2px 0 6px;">'
        f'Latest Week: {latest_wk} (file date {latest_dt})</div>'
    )
    overview = {
        "Total Locations": int(overall.get("total", 0)),
        "Used": int(overall.get("used", 0)),
        "Empty": int(overall.get("empty", 0)),
        "Latest Week Utilization": f"{overall.get('utilization', 0):.2%}",
    }
    parts.append(_metric_cards_html(overview))
    weekly_trend_df = report_data.get("weekly_trend_df", pd.DataFrame())
    parts.append(_sub_header("Weekly Utilization Trend"))
    parts.append(_chart_storage_weekly_html(weekly_trend_df))
    parts.append(_sub_header("Per Storage Type Weekly Utilization"))
    parts.append(_chart_storage_by_type_html(weekly_trend_df))
    parts.append(_sub_header(f"Latest Week ({latest_wk}) Per Type"))
    parts.append(_excel_table_html(report_data.get("per_type_df", pd.DataFrame())))
    return "".join(parts)


def _metric_cards_html(metrics):
    if not metrics:
        return ""
    card = (
        "display:inline-block;background:#f3f4f6;border:1px solid #e5e7eb;"
        "border-radius:8px;padding:10px 16px;margin:4px 6px 4px 0;min-width:115px;"
        "font-family:'Segoe UI',Arial,Helvetica,sans-serif;"
    )
    label_s = "font-size:11px;color:#6b7280;display:block;"
    val_s = "font-size:18px;font-weight:bold;color:#111827;display:block;margin-top:4px;"
    parts = []
    for k, v in metrics.items():
        v_str = f"{v:,}" if isinstance(v, int) else str(v)
        parts.append(
            f'<div style="{card}">'
            f'<span style="{label_s}">{k}</span>'
            f'<span style="{val_s}">{v_str}</span>'
            "</div>"
        )
    return f'<div style="margin:6px 0 12px;">{"".join(parts)}</div>'


def _section_header(num, title):
    return (
        f'<h2 style="color:#1f4e78;font-size:17px;margin:22px 0 8px;'
        f'border-bottom:2px solid #1f4e78;padding-bottom:4px;'
        f'font-family:\'Segoe UI\',Arial,Helvetica,sans-serif;">'
        f"{num}. {title}</h2>"
    )


def _sub_header(text):
    return (
        f'<h3 style="color:#374151;font-size:14px;margin:14px 0 6px;'
        f'font-family:\'Segoe UI\',Arial,Helvetica,sans-serif;">{text}</h3>'
    )


def build_email_html_report(to_addr, cc_addr, subject, intro_text, custom_item_types=None):
    """整合三個分頁的 KPI / 圖表為單頁 HTML 報告 (Outlook 相容)。
    custom_item_types: None 或 list — Storage 區塊將同時輸出 All Item Types 與 Custom Item Types。
    """
    outbound = st.session_state.get("outbound_report")
    inbound = st.session_state.get("inbound_report")
    storage = st.session_state.get("storage_report")
    storage_raw = st.session_state.get("storage_raw")

    now_str = _dt.now().strftime("%Y-%m-%d %H:%M")
    parts = []
    parts.append('<div style="font-family:\'Segoe UI\',Arial,Helvetica,sans-serif;'
                 'color:#1f2937;background:#ffffff;padding:8px 4px;max-width:1100px;">')
    parts.append(
        f'<h1 style="color:#111827;font-size:22px;margin:0 0 6px;'
        f'border-bottom:3px solid #1f4e78;padding-bottom:6px;">{subject or ""}</h1>'
    )
    meta_html = (
        f'<div style="color:#6b7280;font-size:12px;margin-bottom:14px;">'
        f'Generated: {now_str}　|　To: {to_addr if to_addr else "(unspecified)"}'
    )
    if cc_addr:
        meta_html += f'　|　Cc: {cc_addr}'
    meta_html += "</div>"
    parts.append(meta_html)
    if intro_text:
        parts.append(
            f'<div style="background:#fff7ed;border-left:4px solid #f59e0b;'
            f'padding:10px 14px;border-radius:4px;margin:10px 0 18px;font-size:13px;">'
            f'{intro_text}</div>'
        )

    # ===== Outbound (single merged table) =====
    parts.append(_section_header(1, "Outbound KPI"))
    if outbound:
        rng = outbound.get("date_range") or ("", "")
        parts.append(
            f'<div style="color:#6b7280;font-size:12px;margin:2px 0 6px;">'
            f'Date Range: {rng[0]} ~ {rng[1]}</div>'
        )
        parts.append(_metric_cards_html(outbound.get("metrics", {})))
        parts.append(_sub_header("Daily Summary & Visualization"))
        parts.append(_chart_outbound_html(outbound.get("summary_df", pd.DataFrame())))
    else:
        parts.append('<p style="color:#9ca3af;">尚未在 Outbound KPI 分頁上傳檔案。</p>')

    # ===== Inbound (single merged table) =====
    parts.append(_section_header(2, "Inbound KPI"))
    if inbound:
        parts.append(_metric_cards_html(inbound.get("metrics", {})))
        parts.append(_sub_header("Daily Summary & Visualization"))
        parts.append(_chart_inbound_html(inbound.get("summary_df", pd.DataFrame())))
        failed = inbound.get("failed_items_df", pd.DataFrame())
        if isinstance(failed, pd.DataFrame) and not failed.empty:
            parts.append(_sub_header("Failed Item List"))
            parts.append(_excel_table_html(failed, max_rows=50))
    else:
        parts.append('<p style="color:#9ca3af;">尚未在 Inbound KPI 分頁上傳檔案。</p>')

    # ===== Warehouse Utilization (All Item Types + Custom Item Types) =====
    parts.append(_section_header(3, "Warehouse Utilization"))
    if storage_raw:
        # 3.1 All Item Types
        all_report = _compute_storage_trend(storage_raw, selected_item_types=None)
        if all_report:
            parts.append(_render_storage_subsection(all_report, "3.1  All Item Types"))

        # 3.2 Custom Item Types (only if user provided a non-empty, non-full subset)
        all_types = storage_raw.get("all_item_types", [])
        if custom_item_types and 0 < len(custom_item_types) < len(all_types):
            custom_report = _compute_storage_trend(storage_raw, selected_item_types=custom_item_types)
            if custom_report:
                preview = ", ".join(custom_item_types[:5])
                tail = f" (+{len(custom_item_types)-5} more)" if len(custom_item_types) > 5 else ""
                sub_title = f"3.2  Custom Item Types — {preview}{tail}"
                parts.append(_render_storage_subsection(custom_report, sub_title))
        elif custom_item_types and len(custom_item_types) == len(all_types):
            parts.append(
                f'<p style="color:#6b7280;font-size:12px;font-style:italic;margin-top:6px;">'
                f'(Custom Item Types 與 All Item Types 相同，已合併顯示於 3.1)</p>'
            )
    elif storage:
        # Fallback：tab3 沒存 raw 資料，只有 storage_report 一份
        overview = {
            "Total Locations": int(storage.get("overall", {}).get("total", 0)),
            "Used": int(storage.get("overall", {}).get("used", 0)),
            "Empty": int(storage.get("overall", {}).get("empty", 0)),
            "Latest Week Utilization": f"{storage.get('overall', {}).get('utilization', 0):.2%}",
        }
        latest_wk = storage.get("latest_week", "")
        latest_dt = storage.get("latest_date", "")
        parts.append(
            f'<div style="color:#6b7280;font-size:12px;margin:2px 0 6px;">'
            f'Latest Week: {latest_wk} (file date {latest_dt})</div>'
        )
        parts.append(_metric_cards_html(overview))
        weekly_trend_df = storage.get("weekly_trend_df", pd.DataFrame())
        parts.append(_sub_header("Weekly Utilization Trend"))
        parts.append(_chart_storage_weekly_html(weekly_trend_df))
        parts.append(_sub_header("Per Storage Type Weekly Utilization"))
        parts.append(_chart_storage_by_type_html(weekly_trend_df))
        parts.append(_sub_header(f"Latest Week ({latest_wk}) Per Type"))
        parts.append(_excel_table_html(storage.get("per_type_df", pd.DataFrame())))
    else:
        parts.append('<p style="color:#9ca3af;">尚未在 Warehouse Utilization 分頁上傳檔案。</p>')

    parts.append(
        f'<div style="color:#9ca3af;font-size:11px;margin-top:30px;'
        f'border-top:1px solid #e5e7eb;padding-top:8px;">'
        f'此報告由 Framework KPI Tool 自動產生於 {now_str}。'
        f'若需更新內容，請在前三個分頁重新上傳對應資料檔，再回到 Email Report 分頁重新產生。'
        f"</div>"
    )
    parts.append("</div>")
    return "".join(parts)


def build_full_html_document(subject, body_html):
    """完整 HTML 文件（含 DOCTYPE / head），供下載 .html 用。"""
    return (
        "<!DOCTYPE html>\n"
        "<html lang=\"zh-Hant\"><head><meta charset=\"utf-8\">"
        f"<title>{subject}</title></head><body>"
        f"{body_html}"
        "</body></html>"
    )


def build_eml_bytes(to_addr, cc_addr, subject, html_body, from_addr=None):
    """產生 .eml 信件檔位元組。
    加上 X-Unsent: 1 header，雙擊在 Outlook 中會以「草稿」模式開啟。
    結構：multipart/alternative ( text/plain + text/html )。
    """
    msg = _MIMEMultipart("alternative")
    msg["Subject"] = subject or ""
    if from_addr:
        msg["From"] = from_addr
    msg["To"] = to_addr or ""
    if cc_addr:
        msg["Cc"] = cc_addr
    msg["Date"] = _formatdate(localtime=True)
    msg["X-Unsent"] = "1"  # Outlook 會將此 EML 開為「未寄出」草稿
    msg["MIME-Version"] = "1.0"

    plain_fallback = (
        "您的郵件用戶端不支援 HTML 顯示。\n"
        "請改用支援 HTML 的郵件程式 (例如 Microsoft Outlook) 開啟此信件。"
    )
    msg.attach(_MIMEText(plain_fallback, "plain", "utf-8"))
    msg.attach(_MIMEText(html_body, "html", "utf-8"))
    return msg.as_bytes()


with tab4:
    st.subheader("Email Report — 一鍵產生可直接在 Outlook 寄發的 KPI 報告")
    st.caption(
        "完成前三個分頁的資料上傳後，回到本頁可一次性匯出整合報告。"
        " 推薦下載 **.eml 信件檔**：雙擊後 Outlook 會自動以「草稿」模式開啟，"
        "整份 HTML 報告（含表格 + 視覺化圖表）會內嵌在信件本文，"
        "你只需檢查收件人後按「傳送」即可。"
    )

    ready_out = "outbound_report" in st.session_state
    ready_in = "inbound_report" in st.session_state
    ready_st = "storage_report" in st.session_state
    storage_raw_avail = "storage_raw" in st.session_state
    c1, c2, c3 = st.columns(3)
    c1.metric("Outbound KPI", "已就緒" if ready_out else "未準備")
    c2.metric("Inbound KPI", "已就緒" if ready_in else "未準備")
    c3.metric("Warehouse Utilization", "已就緒" if ready_st else "未準備")

    # ===== Item Type 多選 (Storage 客製化用) =====
    custom_item_types = None
    if storage_raw_avail:
        _storage_raw = st.session_state["storage_raw"]
        _all_types = _storage_raw.get("all_item_types", [])
        if _all_types:
            st.markdown("#### Storage — 客製化 Item Type")
            st.caption(
                "報告 Storage 區塊會分成 **3.1 All Item Types** 與 **3.2 Custom Item Types** 兩部分。"
                "下方可勾選要納入「Custom」分析的 Item Type；若全選或全不選，則只顯示 3.1。"
            )
            custom_item_types = st.multiselect(
                "選擇要納入 Custom 子區塊的 Item Type",
                options=_all_types,
                default=_all_types,
                key="email_custom_item_types",
                help="多選；可清空或全選代表只顯示 All Item Types。",
            )

    if not (ready_out or ready_in or ready_st):
        st.info("請先到前三個分頁上傳資料，回到此頁即可產生整合報告。")
    else:
        st.markdown("#### 收件人 / 主旨設定")
        col_a, col_b = st.columns(2)
        with col_a:
            email_to = st.text_input(
                "To (收件人，多個請用逗號或分號分隔)",
                value=DEFAULT_EMAIL_TO,
                key="email_to",
                help="預設值已填入；可直接編輯。",
            )
        with col_b:
            email_cc = st.text_input(
                "CC (副本，多個請用逗號或分號分隔)",
                value=DEFAULT_EMAIL_CC,
                key="email_cc",
            )

        email_subject = st.text_input(
            "Subject (主旨)",
            value=DEFAULT_EMAIL_SUBJECT + "  " + _dt.now().strftime("%Y-%m-%d"),
            key="email_subject",
        )
        email_intro = st.text_area(
            "Intro / Note (前言，可選)",
            value="Hi team,\n\n附上本週 Framework Outbound / Inbound / Warehouse KPI 報告，請查收。",
            key="email_intro",
            height=110,
        )

        if st.button("產生報告", type="primary", key="btn_build_email"):
            body_html = build_email_html_report(
                to_addr=email_to.strip(),
                cc_addr=email_cc.strip(),
                subject=email_subject.strip(),
                intro_text=email_intro.replace("\n", "<br>") if email_intro else "",
                custom_item_types=custom_item_types,
            )
            full_html = build_full_html_document(email_subject.strip(), body_html)
            eml_bytes = build_eml_bytes(
                to_addr=email_to.strip(),
                cc_addr=email_cc.strip(),
                subject=email_subject.strip(),
                html_body=full_html,
            )
            st.session_state["last_email_html"] = full_html
            st.session_state["last_email_body_html"] = body_html
            st.session_state["last_email_eml"] = eml_bytes
            st.success(
                f"報告已產生：HTML {len(full_html):,} 字元　|　EML {len(eml_bytes):,} bytes。"
                "在下方下載 .eml，雙擊即可在 Outlook 開啟草稿信件。"
            )

        body_html = st.session_state.get("last_email_body_html", "")
        full_html = st.session_state.get("last_email_html", "")
        eml_bytes = st.session_state.get("last_email_eml", b"")

        if body_html:
            st.markdown("#### 報告預覽 (Outlook 中將看到的內容)")
            with st.expander("展開 / 收合預覽", expanded=True):
                st.components.v1.html(body_html, height=900, scrolling=True)

            stamp = _dt.now().strftime("%Y%m%d_%H%M")
            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                st.download_button(
                    label="下載 .eml 信件檔 (推薦，雙擊直接在 Outlook 開草稿)",
                    data=eml_bytes,
                    file_name=f"framework_kpi_report_{stamp}.eml",
                    mime="message/rfc822",
                    key="dl_email_eml",
                    type="primary",
                    help=(
                        "下載後雙擊 .eml 檔，Outlook 會自動以草稿模式開啟，"
                        "整份報告含表格與圖表都會在信件本文中。檢查收件人後按『傳送』即可。"
                    ),
                )
            with col_dl2:
                st.download_button(
                    label="下載 HTML 報告 (.html) — 備用",
                    data=full_html.encode("utf-8"),
                    file_name=f"framework_kpi_report_{stamp}.html",
                    mime="text/html",
                    key="dl_email_html",
                    help="若 .eml 開啟異常，可改下載 .html，用瀏覽器開後 Ctrl+A 全選複製貼到 Outlook。",
                )

            st.info(
                "**使用流程**：\n"
                "1. 點上方「下載 .eml 信件檔」\n"
                "2. 雙擊下載的 .eml 檔 → Outlook 自動開啟一封草稿\n"
                "3. 報告內容（表格 + 圖表）已內嵌在信件本文，收件人與主旨也已填好\n"
                "4. 確認無誤後按「傳送」即可"
            )
