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
            special_rule_dict if row["Country code"] == "TW" else {}
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

tab1, tab2, tab3 = st.tabs(["Outbound KPI", "Inbound KPI", "Warehouse Utilization"])

with tab1:

    # ===== Upload =====
    uploaded_file = st.file_uploader("Upload raw data (.xlsx)", type=["xlsx"])
    special_rule_file = st.file_uploader("Upload special date rule (.xlsx)", type=["xlsx"])

    if uploaded_file is not None:

        # ===== Load Data =====
        raw_df = load_data(uploaded_file)

        special_rule_dict = {}

        if special_rule_file is not None:
            special_df = load_special_rules(special_rule_file)
            special_rule_dict = build_special_rule_dict(special_df)

        # ===== Prepare KPI =====
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

        # ===== Summary =====
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

    if inbound_raw_file is not None and inbound_mapping_file is not None:

        # ===== Load Data =====
        raw_df = load_inbound_raw(inbound_raw_file)
        mapping_df = load_inbound_mapping(inbound_mapping_file)

        # ===== Prepare Data =====
        inbound_df = prepare_inbound_data(raw_df, mapping_df)

        # ===== 日期篩選 =====
        inbound_df["Actual Date"] = pd.to_datetime(inbound_df["Date"], errors="coerce")

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
            else:
                filtered_inbound_df = inbound_df.copy()
        else:
            filtered_inbound_df = inbound_df.copy()

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
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Rows", f"{len(filtered_inbound_df):,}")
        col2.metric("Matched", f"{(filtered_inbound_df['Mapping Status'] == 'Matched').sum():,}")
        col3.metric("Need Confirm", f"{(filtered_inbound_df['Mapping Status'] == '確認報單號碼').sum():,}")
        col4.metric("Failed", f"{(filtered_inbound_df['KPI Result'] == 'Failed').sum():,}")

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
            display_inbound_summary_df["kpi_rate"] = display_inbound_summary_df["kpi_rate"].map(lambda x: f"{x:.2%}")
            display_inbound_summary_df = display_inbound_summary_df.rename(columns={
                "Report Date": "Date",
                "in_kpi": "In KPI",
                "failed": "Failed",
                "total": "Total",
                "kpi_rate": "KPI Rate"
            })
            st.dataframe(display_inbound_summary_df, width="stretch")
        else:
            st.info("No summary data available for the selected date range.")

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

    storage_usage_file = st.file_uploader(
        "Upload Storage Usage Report (實際庫存報表 .xlsx)",
        type=["xlsx"],
        key="storage_usage",
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

    if storage_master_file is not None and storage_usage_file is not None:

        # ===== Load =====
        master_df = load_storage_master(storage_master_file)
        usage_df = load_storage_usage(storage_usage_file)

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

        if master_df.empty:
            st.error("無法從儲位總表讀到 Location/Type 欄位，請確認檔案格式。")
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

            # ===== 總體 (不分 Item Type) 使用率：作為「總攬」 =====
            type_summary_all_df, overall_summary_all, unmatched_df, used_locations_all = (
                build_utilization_summary(master_df, usage_df)
            )

            st.markdown("### 總攬 (All Item Types)")

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

            # ===== Item Type 篩選器 =====
            st.markdown("---")
            st.markdown("### 依 Item Type 觀察")

            # 收集出現過的 Item Types：以 usage 中真正出現的優先，加上 master 中已知的
            usage_item_types = sorted(usage_df["Item Type"].dropna().unique())
            item_type_options = ["全部 (All)"] + usage_item_types

            selected_item_type = st.selectbox(
                "選擇要分析的 Item Type",
                options=item_type_options,
                index=0,
                key="storage_item_type_filter",
                help="選『全部 (All)』可查看跨所有 Item Type 的使用率；選擇單一 Item Type 可看該品項在各 Location Type 的儲位佔用率。",
            )

            if selected_item_type == "全部 (All)":
                filtered_usage_df = usage_df
                type_summary_df = type_summary_all_df
                overall_summary = overall_summary_all
                used_locations = used_locations_all
                scope_label = "全部 (All Item Types)"
            else:
                filtered_usage_df = usage_df[usage_df["Item Type"] == selected_item_type].copy()
                type_summary_df, overall_summary, _unmatched_ignore, used_locations = (
                    build_utilization_summary(master_df, filtered_usage_df)
                )
                scope_label = f"Item Type = {selected_item_type}"

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
            file_name_scope = (
                "all" if selected_item_type == "全部 (All)"
                else selected_item_type.lower().replace(" ", "_")
            )
            st.download_button(
                label=f"Download Utilization Report ({scope_label}) (.xlsx)",
                data=excel_bytes,
                file_name=f"warehouse_utilization_{file_name_scope}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    else:
        st.info("請同時上傳「儲位總表」與「實際庫存報表」以開始計算。")
