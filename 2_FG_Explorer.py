import io
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from dateutil.relativedelta import relativedelta

DEFAULT_FG_PATH = Path(
    r"C:\Users\ashalaby\OneDrive - Halwani Bros\Planning - Sources\new view 2023\FP module 23.xlsb"
)
PRIMARY_SHEETS = ("Data", "Data2")
TEXT_COLUMNS = {
    "itemnumber",
    "itemname",
    "factory",
    "family",
    "unit",
    "brand",
    "category",
    "subfamily",
    "subfamilyname",
    "group",
    "subgroup",
    "market",
    "storagetype",
    "rawtype",
}
DC_LOCATION_MAP = {
    "10th of Ramadan": {"lat": 30.2920, "lon": 31.7500},
    "Ramadan": {"lat": 30.2920, "lon": 31.7500},
    "Cairo": {"lat": 30.0444, "lon": 31.2357},
    "Alex": {"lat": 31.2001, "lon": 29.9187},
    "Alexandria": {"lat": 31.2001, "lon": 29.9187},
    "Gharbeya": {"lat": 30.7865, "lon": 31.0004},
    "Tanta": {"lat": 30.7865, "lon": 31.0004},
    "Upper Egypt": {"lat": 25.6872, "lon": 32.6396},
    "Luxor": {"lat": 25.6872, "lon": 32.6396},
    "Giza": {"lat": 30.0131, "lon": 31.2089},
    "Beheira": {"lat": 31.0341, "lon": 30.4682},
    "Damanhour": {"lat": 31.0341, "lon": 30.4682},
    "Zagazig": {"lat": 30.5877, "lon": 31.5020},
    "Isma3leya": {"lat": 30.6043, "lon": 32.2723},
}
METRIC_ORDER = [
    "CurOS",
    "CurAPP",
    "CurAP",
    "CurOST",
    "CurST",
    "CurAS",
    "CurFOC",
    "Oh",
    "NextOS",
]
MONTHLY_METRIC_CODES = ["BP_CUR", "BP_PRV", "AS"]
MONTHLY_METRIC_LABELS = {
    "BP_CUR": "Current BP",
    "BP_PRV": "Previous BP",
    "AS": "Actual Sales",
}
WEEKLY_METRICS = ["CurST", "CurAS", "CurAPP", "CurAP"]
MONTH_ABBR = {
    "jan": 1,
    "feb": 2,
    "mar": 3,
    "apr": 4,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "aug": 8,
    "sep": 9,
    "oct": 10,
    "nov": 11,
    "dec": 12,
}
METRIC_LABELS = {
    "CurOS": "Opening Stock (CurOS)",
    "CurAPP": "Production Plan (CurAPP)",
    "CurAP": "Actual Production (CurAP)",
    "CurOST": "Original Sales Target (CurOST)",
    "CurST": "Sales Target (CurST)",
    "CurAS": "Actual Sales (CurAS)",
    "CurFOC": "Free of Charge (CurFOC)",
    "Oh": "On Hand (OH)",
    "NextOS": "Next Opening Stock (NextOS)",
}
WEEKLY_METRIC_LABELS = {
    "CurST": "Sales Target",
    "CurAS": "Actual Sales",
    "CurAPP": "Production Plan",
    "CurAP": "Actual Production",
}
METRIC_COLOR_MAP = {
    "CurOS": "#f7dc6f",  # OS - light yellow
    "CurAPP": "#8e5b32",  # APP - brown
    "CurAP": "#ff7f0e",  # AP - orange
    "CurOST": "#5dade2",  # Original ST - cool blue
    "CurST": "#1f77b4",  # ST - blue
    "CurAS": "#2ecc71",  # AS - green
    "CurFOC": "#af7ac5",  # FOC - purple accent
    "Oh": "#f1c40f",  # OH - yellow
    "NextOS": "#f9e79f",  # Next OS - lighter yellow
    "OS": "#f7dc6f",
    "ST": "#1f77b4",
    "AS": "#2ecc71",
    "APP": "#8e5b32",
    "AP": "#ff7f0e",
    "OST": "#5dade2",
    "RFC": "#5dade2",
    "BP_CUR": "#1f77b4",
    "BP_PRV": "#95a5a6",
}
METRIC_CANDIDATES = {
    "CurOS": ["CurOS", "CurSOH", "CurStockOnHand", "CurOpeningStock"],
    "CurAPP": ["CurAPP"],
    "CurAP": ["CurAP"],
    "CurOST": ["CurOST", "CurOrigST", "CurOriginalST"],
    "CurST": ["CurST"],
    "CurAS": ["CurAS"],
    "CurFOC": ["CurFOC", "FOC", "CurFOCQty"],
    "Oh": ["Oh", "OH", "OnHand", "StockOnHand", "SOH"],
    "NextOS": ["NextOS", "NextSOH", "NextStockOnHand"],
}
WEEK_RANGE = range(1, 6)
XLSXWRITER_AVAILABLE = True
TARGET_YEAR_SUFFIX = "26"
MONTH_SEQUENCE = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTH_LABELS = [f"{month}{TARGET_YEAR_SUFFIX}" for month in MONTH_SEQUENCE]
BP_COLUMN_PATTERN = re.compile(r"^(?P<prefix>.*?)(?P<month>[A-Za-z]{3})(?P<year>\d{2})BP$", re.IGNORECASE)
AS_COLUMN_PATTERN = re.compile(r"^(?P<prefix>.*?)(?P<month>[A-Za-z]{3})(?P<year>\d{2})AS$", re.IGNORECASE)


def _normalize_name(name: str | int) -> str:
    return str(name).strip().lower().replace(" ", "").replace("-", "").replace("_", "")


def _find_column(columns: pd.Index, candidates: List[str]) -> str | None:
    columns_map = {_normalize_name(col): col for col in columns}
    for candidate in candidates:
        normalized_candidate = _normalize_name(candidate)
        actual = columns_map.get(normalized_candidate)
        if actual is not None:
            return actual
    return None


def _find_all(columns: pd.Index, pattern: str) -> List[str]:
    normalized_pattern = _normalize_name(pattern)
    return [col for col in columns if normalized_pattern in _normalize_name(col)]


def _unique_sorted_values(df: pd.DataFrame, column: str | None, *, exclude_all: bool = True) -> List[str]:
    if column is None or column not in df.columns:
        return []
    series = df[column].dropna().astype(str).str.strip()
    values = {
        value
        for value in series
        if value and (not exclude_all or value.lower() != "all")
    }
    return sorted(values)


def _calculate_dioh(df: pd.DataFrame) -> float | None:
    oh_col = metric_column_map.get("Oh")
    demand_col = metric_column_map.get("CurAS")
    if not oh_col or not demand_col:
        return None
    if oh_col not in df.columns or demand_col not in df.columns:
        return None

    total_oh = pd.to_numeric(df[oh_col], errors="coerce").sum()
    total_demand = pd.to_numeric(df[demand_col], errors="coerce").sum()
    if total_oh <= 0 or total_demand <= 0:
        return None

    today = datetime.today()
    first_day = today.replace(day=1)
    next_month = first_day + relativedelta(months=1)
    days_in_month = (next_month - first_day).days or 30

    avg_daily_demand = total_demand / days_in_month if days_in_month else None
    if not avg_daily_demand or avg_daily_demand <= 0:
        return None

    return total_oh / avg_daily_demand


def _build_weekly_series(row: pd.Series) -> Tuple[List[str], Dict[str, List[float]]]:
    labels: List[str] = []
    series_map: Dict[str, List[float]] = {metric: [] for metric in WEEKLY_METRICS}

    for week in WEEK_RANGE:
        week_label = f"W{week}"
        has_data = False
        weekly_values: Dict[str, float] = {}
        for metric in WEEKLY_METRICS:
            candidates = [f"{metric}W{week}"]
            column_name = _find_column(row.index, candidates)
            value = row[column_name] if column_name else np.nan
            if pd.notna(value):
                has_data = True
            weekly_values[metric] = value
        if has_data:
            labels.append(week_label)
            for metric, value in weekly_values.items():
                series_map[metric].append(value)

    return labels, series_map


def _build_monthly_series(row: pd.Series) -> Tuple[List[str], Dict[str, List[float]]]:
    label_data: Dict[str, Dict[str, float]] = {}
    label_order: Dict[str, int] = {}

    for column in row.index:
        normalized = _normalize_name(column)
        for metric_code in MONTHLY_METRIC_CODES:
            metric_norm = metric_code.lower()
            if not normalized.endswith(metric_norm):
                continue
            prefix = column[: -len(metric_code)]
            if not prefix:
                continue
            label, order = _resolve_prefix_label(prefix)
            if label.upper().startswith("YTD"):
                continue
            if metric_code == "BP" and label not in MONTH_LABELS:
                continue
            label_order[label] = order
            label_data.setdefault(label, {})[metric_code] = row[column]
            break

    if not label_order:
        return [], {metric: [] for metric in MONTHLY_METRIC_CODES}

    ordered_labels = [label for label, _ in sorted(label_order.items(), key=lambda item: item[1])]
    series_map: Dict[str, List[float]] = {metric: [] for metric in MONTHLY_METRIC_CODES}
    for label in ordered_labels:
        values = label_data.get(label, {})
        for metric in MONTHLY_METRIC_CODES:
            series_map[metric].append(values.get(metric, np.nan))

    return ordered_labels, series_map


def _build_monthly_series(row: pd.Series) -> Tuple[List[str], Dict[str, List[float]]]:
    month_data: Dict[str, Dict[str, float]] = {
        label: {metric: np.nan for metric in MONTHLY_METRIC_CODES}
        for label in MONTH_LABELS
    }

    for column in row.index:
        col_name = str(column).strip()
        match = BP_COLUMN_PATTERN.match(col_name)
        if not match:
            as_match = AS_COLUMN_PATTERN.match(col_name)
            if not as_match:
                continue
            month = as_match.group("month").title()
            year = as_match.group("year")
            month_label = f"{month}{year}"
            if month_label not in MONTH_LABELS:
                continue

            value = pd.to_numeric(row[column], errors="coerce")
            if pd.isna(value):
                continue

            month_entry = month_data.setdefault(month_label, {metric: np.nan for metric in MONTHLY_METRIC_CODES})
            existing = month_entry.get("AS")
            month_entry["AS"] = value if pd.isna(existing) else existing + value
            continue

        month = match.group("month").title()
        year = match.group("year")
        month_label = f"{month}{year}"
        if month_label not in MONTH_LABELS:
            continue

        prefix = (match.group("prefix") or "").strip().lower()
        metric_key = "BP_PRV" if prefix.startswith("prv") else "BP_CUR"

        value = row[column]
        month_entry = month_data.setdefault(month_label, {metric: np.nan for metric in MONTHLY_METRIC_CODES})
        numeric_value = pd.to_numeric(value, errors="coerce")
        if pd.isna(numeric_value):
            continue
        existing = month_entry.get(metric_key)
        month_entry[metric_key] = numeric_value if pd.isna(existing) else existing + numeric_value

    labels: List[str] = []
    series_map: Dict[str, List[float]] = {metric: [] for metric in MONTHLY_METRIC_CODES}

    for month_label in MONTH_LABELS:
        month_entry = month_data.get(month_label)
        if not month_entry:
            continue
        values = [month_entry.get(metric, np.nan) for metric in MONTHLY_METRIC_CODES]
        if all(pd.isna(value) for value in values):
            continue
        labels.append(month_label)
        for metric in MONTHLY_METRIC_CODES:
            series_map[metric].append(month_entry.get(metric, np.nan))

    return labels, series_map


def _filter_monthly_series(series_map: Dict[str, List[float]]) -> Tuple[Dict[str, List[float]], List[str]]:
    filtered: Dict[str, List[float]] = {}
    order: List[str] = []
    for metric in MONTHLY_METRIC_CODES:
        values = series_map.get(metric, [])
        if values and any(pd.notna(value) for value in values):
            filtered[metric] = values
            order.append(metric)
    return filtered, order


def dataframe_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes | None:
    global XLSXWRITER_AVAILABLE
    if df is None or df.empty or not XLSXWRITER_AVAILABLE:
        return None
    buffer = io.BytesIO()
    try:
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    except ImportError:
        st.error("❌ Error: Please install 'xlsxwriter' using `pip install xlsxwriter`")
        XLSXWRITER_AVAILABLE = False
        return None
    return buffer.getvalue()


def _resolve_prefix_label(prefix: str) -> Tuple[str, int]:
    prefix_clean = prefix.strip()
    prefix_lower = prefix_clean.lower()
    if prefix_lower.startswith("prv"):
        suffix = prefix_clean[3:]
        order = -int(suffix) if suffix.isdigit() else -1
        label = f"PRV{suffix}" if suffix else "PRV"
        return label, order
    if prefix_lower.startswith("cur"):
        suffix = prefix_clean[3:]
        label = f"CUR{suffix.upper()}" if suffix else "CUR"
        return label, 0
    if prefix_lower.startswith("next"):
        suffix = prefix_clean[4:]
        order = int(suffix) if suffix.isdigit() else 1
        label = f"NEXT{suffix}" if suffix else "NEXT"
        return label, order

    match = re.match(r"(?P<month>[A-Za-z]{3})(?P<year>\d{2})(?P<suffix>[A-Za-z]*)", prefix_clean)
    if match:
        month_name = match.group("month").lower()
        year = int(match.group("year")) + 2000
        suffix = match.group("suffix").upper()
        month_num = MONTH_ABBR.get(month_name)
        if month_num:
            base = pd.Timestamp.today().normalize().replace(day=1)
            target = pd.Timestamp(year=year, month=month_num, day=1)
            order = (target.year - base.year) * 12 + (target.month - base.month)
        else:
            order = 100
        label = f"{match.group('month').title()}{match.group('year')}{suffix}" if suffix else f"{match.group('month').title()}{match.group('year')}"
        return label, order

    return prefix_clean.upper(), 100


def _render_stacked_bar(
    labels: List[str],
    series_map: Dict[str, List[float]],
    title: str,
    order: List[str],
    label_map: Dict[str, str]
) -> None:
    fig = go.Figure()
    for metric in order:
        values = series_map.get(metric, [])
        if not values or all(pd.isna(values)):
            continue
        color = METRIC_COLOR_MAP.get(metric)
        fig.add_trace(
            go.Bar(
                x=labels,
                y=values,
                name=label_map.get(metric, metric),
                marker={"color": color} if color else None,
            )
        )
    if not fig.data:
        st.info("No data available for the selected view.")
        return
    fig.update_layout(
        title=title,
        barmode="stack",
        xaxis_title="Period",
        yaxis_title="Quantity",
        height=420,
        hovermode="x unified"
    )
    st.plotly_chart(fig, use_container_width=True)


def _format_tonnage(value: float | int | str) -> str:
    try:
        numeric_value = float(value)
    except (TypeError, ValueError):
        return "-"

    if pd.isna(numeric_value):
        return "-"

    return f"{numeric_value:,.0f} Tons"


def _render_line_chart(
    labels: List[str],
    series_map: Dict[str, List[float]],
    title: str,
    order: List[str],
    label_map: Dict[str, str]
) -> None:
    fig = go.Figure()
    for metric in order:
        values = series_map.get(metric, [])
        if not values or all(pd.isna(values)):
            continue
        color = METRIC_COLOR_MAP.get(metric)
        fig.add_trace(
            go.Scatter(
                x=labels,
                y=values,
                name=label_map.get(metric, metric),
                mode="lines+markers",
                line={"color": color} if color else None,
                marker={"color": color} if color else None,
            )
        )
    if not fig.data:
        st.info("No data available for the selected view.")
        return
    fig.update_layout(
        title=title,
        xaxis_title="Period",
        yaxis_title="Quantity",
        height=420,
        hovermode="x unified"
    )
    st.plotly_chart(fig, use_container_width=True)


def _format_days(value: float) -> str:
    if value is None or pd.isna(value):
        return "—"
    return f"{value:,.1f} days"


@st.cache_data(show_spinner=False)
def load_fg_dataset(file_bytes: bytes | None) -> pd.DataFrame:
    source = io.BytesIO(file_bytes) if file_bytes else DEFAULT_FG_PATH
    try:
        sheets = pd.read_excel(source, sheet_name=list(PRIMARY_SHEETS), engine="pyxlsb")
    except FileNotFoundError:
        raise
    except ValueError:
        source = io.BytesIO(file_bytes) if file_bytes else DEFAULT_FG_PATH
        sheets = pd.read_excel(source, sheet_name=[PRIMARY_SHEETS[0]], engine="pyxlsb")

    data_df = sheets.get("Data", pd.DataFrame())
    data2_df = sheets.get("Data2", pd.DataFrame())

    if data_df.empty:
        return pd.DataFrame()

    for df in (data_df, data2_df):
        if df is None or df.empty:
            continue
        df.columns = [str(col).strip() for col in df.columns]
        if "ItemNumber" in df.columns:
            df["ItemNumber"] = df["ItemNumber"].astype(str).str.strip()

    base_df = data_df.copy()

    if data2_df is not None and not data2_df.empty:
        data2_unique = data2_df.drop_duplicates(subset=["ItemNumber"])
        merged_df = base_df.merge(
            data2_unique,
            on="ItemNumber",
            how="left",
            suffixes=("", "_d2")
        )
    else:
        merged_df = base_df

    merged_df.columns = [str(col).strip() for col in merged_df.columns]
    merged_df = merged_df.dropna(how="all")

    for column in merged_df.columns:
        normalized = _normalize_name(column)
        if normalized in TEXT_COLUMNS:
            merged_df[column] = merged_df[column].astype(str).str.strip()
        else:
            merged_df[column] = pd.to_numeric(merged_df[column], errors="ignore")

    return merged_df


st.set_page_config(page_title="FG Explorer", layout="wide")
st.title("🏷️ FG Explorer")
st.caption(
    "Loads the default finished goods workbook automatically. Upload a different file to override."
)

uploaded_file = st.file_uploader(
    "Upload alternative FG XLSB file",
    type=["xlsb"],
    help="Optional: upload another finished goods workbook."
)

file_bytes = uploaded_file.getvalue() if uploaded_file else None

try:
    with st.spinner("Loading FG data..."):
        fg_df = load_fg_dataset(file_bytes)
except FileNotFoundError:
    st.error(
        "Default FG workbook not found at `C:/Users/ashalaby/OneDrive - Halwani Bros/Planning - Sources/new view 2023/FP module 23.xlsb`."
    )
    st.stop()
except Exception as exc:
    st.error(f"Unable to read FG workbook: {exc}")
    st.stop()

if fg_df.empty:
    st.warning("FG workbook is empty or missing the expected 'Data' sheet.")
    st.stop()

factory_col = _find_column(fg_df.columns, ["Factory"])
item_col = _find_column(fg_df.columns, ["ItemNumber"])
name_col = _find_column(fg_df.columns, ["ItemName"])
family_col = _find_column(fg_df.columns, ["Family", "SubFamily", "SubFamilyName"])
curst_col = _find_column(fg_df.columns, ["CurST"])
curas_col = _find_column(fg_df.columns, ["CurAS"])
curapp_col = _find_column(fg_df.columns, ["CurAPP"])
curap_col = _find_column(fg_df.columns, ["CurAP"])
oh_col = _find_column(
    fg_df.columns,
    ["CurSOH", "CurOH", "SOH", "CurStockOnHand", "StockOnHand", "OH"]
)
storagetype_col = _find_column(
    fg_df.columns,
    ["StorageType", "Storage Type", "Storage_Type", "Storage"]
)
rawtype_col = _find_column(
    fg_df.columns,
    ["RawType", "Raw Type", "Raw_Type", "Raw"]
)

required_columns = {
    "Factory column": factory_col,
    "ItemNumber column": item_col,
    "Family column": family_col,
    "CurST column": curst_col,
    "CurAS column": curas_col,
    "OH column": oh_col,
}
missing_columns = [label for label, column in required_columns.items() if column is None]
if missing_columns:
    st.error(
        "Missing required column(s): " + ", ".join(missing_columns)
        + ". Upload a workbook containing the expected columns (Factory, ItemNumber, CurST, CurAS, OH/StockOnHand)."
    )
    st.stop()

st.markdown("### Dataset overview")
st.caption(
    f"Loaded **{len(fg_df):,}** rows from "
    f"{'uploaded file' if uploaded_file else 'default workbook'}."
)

if name_col:
    unique_items = fg_df[[item_col, name_col]].drop_duplicates().shape[0]
else:
    unique_items = fg_df[item_col].nunique()
st.caption(f"Unique FG items: {unique_items:,}")

numeric_cols = fg_df.select_dtypes(include=[np.number]).columns
fg_df[numeric_cols] = fg_df[numeric_cols].fillna(0)

# Drop rows missing factory or family values
fg_df = fg_df[
    fg_df[factory_col].astype(str).str.strip().ne("")
    & fg_df[factory_col].notna()
]
fg_df = fg_df[
    fg_df[family_col].astype(str).str.strip().ne("")
    & fg_df[family_col].notna()
]

market_col = _find_column(fg_df.columns, ["Market", "LocExp", "LOCEXP", "Channel"])
if market_col is None:
    market_col = "_Market"
    fg_df[market_col] = "All"

metric_column_map: Dict[str, str | None] = {
    metric: _find_column(fg_df.columns, candidates)
    for metric, candidates in METRIC_CANDIDATES.items()
}

st.markdown("### Filters")
filter_stage_df = fg_df.copy()

row_one = st.columns(3)
row_two = st.columns(3)

with row_one[0]:
    storage_options = ["All"] + _unique_sorted_values(filter_stage_df, storagetype_col)
    selected_storage = st.selectbox(
        "Storage type",
        storage_options,
        index=0,
        key="fg_storage_select",
    )
    if storagetype_col and selected_storage != "All":
        filter_stage_df = filter_stage_df[
            filter_stage_df[storagetype_col].astype(str).str.strip().eq(str(selected_storage).strip())
        ]

with row_one[1]:
    raw_options = ["All"] + _unique_sorted_values(filter_stage_df, rawtype_col)
    selected_raw = st.selectbox(
        "Raw type",
        raw_options,
        index=0,
        key="fg_raw_select",
    )
    if rawtype_col and selected_raw != "All":
        filter_stage_df = filter_stage_df[
            filter_stage_df[rawtype_col].astype(str).str.strip().eq(str(selected_raw).strip())
        ]

with row_one[2]:
    market_options = ["All"] + _unique_sorted_values(filter_stage_df, market_col)
    selected_market = st.selectbox(
        "Market",
        market_options,
        index=0,
        key="fg_market_select",
    )
    if selected_market != "All":
        filter_stage_df = filter_stage_df[
            filter_stage_df[market_col].astype(str).str.strip().eq(str(selected_market).strip())
        ]

with row_two[0]:
    factory_options = ["All"] + _unique_sorted_values(filter_stage_df, factory_col)
    selected_factory = st.selectbox(
        "Factory",
        factory_options,
        index=0,
        key="fg_factory_select",
    )
    if selected_factory != "All":
        filter_stage_df = filter_stage_df[
            filter_stage_df[factory_col].astype(str).str.strip().eq(str(selected_factory).strip())
        ]

with row_two[1]:
    family_options = ["All"] + _unique_sorted_values(filter_stage_df, family_col)
    selected_family = st.selectbox(
        "Family",
        family_options,
        index=0,
        key="fg_family_select",
    )
    if selected_family != "All" and family_col:
        filter_stage_df = filter_stage_df[
            filter_stage_df[family_col].astype(str).str.strip().eq(str(selected_family).strip())
        ]

with row_two[2]:
    item_options = ["All"]
    item_display_to_value: Dict[str, str] = {}
    if item_col and item_col in filter_stage_df.columns:
        if name_col and name_col in filter_stage_df.columns:
            item_subset = (
                filter_stage_df[[item_col, name_col]]
                .dropna()
                .drop_duplicates()
                .assign(
                    display=lambda df: df[item_col].astype(str).str.zfill(6)
                    + " - "
                    + df[name_col].astype(str)
                )
            )
            displays = sorted(item_subset["display"].tolist())
            item_options += displays
            item_display_to_value = {
                display: str(value)
                for display, value in zip(item_subset["display"], item_subset[item_col])
            }
        else:
            item_values = _unique_sorted_values(filter_stage_df, item_col, exclude_all=False)
            item_options += item_values
            item_display_to_value = {value: value for value in item_values}
    selected_item_filter = st.selectbox(
        "Items",
        item_options,
        index=0,
        key="fg_item_filter",
    )
    if selected_item_filter != "All" and item_display_to_value:
        target_item = item_display_to_value.get(selected_item_filter)
        if target_item is not None:
            filter_stage_df = filter_stage_df[
                filter_stage_df[item_col].astype(str).str.strip().eq(str(target_item).strip())
            ]

filtered_df = filter_stage_df.copy()

if filtered_df.empty:
    st.warning("No data matches the current filters. Adjust factory or market selection.")
    st.stop()


def _compute_kpis(df: pd.DataFrame) -> Dict[str, Tuple[int, pd.DataFrame]]:
    planning_col = _find_column(df.columns, ["Planning"])
    planning_series = df[planning_col].astype(str).str.upper().str.strip() if planning_col else pd.Series([], dtype=str)

    curst_col = metric_column_map.get("CurST")
    ssqty_col = _find_column(df.columns, ["SSQty"])
    oh_column = _find_column(df.columns, ["OH", "StockOnHand", "CurSOH", "CurOH", "OnHand"])

    numeric_curst = pd.to_numeric(df[curst_col], errors="coerce") if curst_col and curst_col in df.columns else pd.Series([], dtype=float)
    numeric_ssqty = pd.to_numeric(df[ssqty_col], errors="coerce") if ssqty_col and ssqty_col in df.columns else pd.Series([], dtype=float)
    numeric_oh = pd.to_numeric(df[oh_column], errors="coerce") if oh_column and oh_column in df.columns else pd.Series([], dtype=float)

    detail_columns = [
        col
        for col in {item_col, name_col, planning_col, curst_col, ssqty_col, oh_column}
        if col and col in df.columns
    ]

    def _detail(mask: pd.Series) -> Tuple[int, pd.DataFrame]:
        if mask is None or mask.empty:
            return 0, pd.DataFrame(columns=detail_columns)
        subset = df.loc[mask, detail_columns].copy() if detail_columns else df.loc[mask].copy()
        return int(mask.sum()), subset

    mts_mask = planning_series.eq("MTS") if not planning_series.empty else pd.Series([], dtype=bool)
    mto_mask = planning_series.eq("MTO") if not planning_series.empty else pd.Series([], dtype=bool)

    oh_lt_ss_mask = pd.Series([], dtype=bool)
    if not numeric_oh.empty and not numeric_ssqty.empty:
        oh_lt_ss_mask = numeric_oh < numeric_ssqty

    zero_oh_mask = pd.Series([], dtype=bool)
    if not numeric_oh.empty and not numeric_curst.empty:
        zero_oh_mask = ((planning_series == "MTS") | (numeric_curst > 0)) & (numeric_curst > 0) & (numeric_oh == 0)

    results: Dict[str, Tuple[int, pd.DataFrame]] = {}
    results["MTS"] = _detail(mts_mask)
    results["MTO"] = _detail(mto_mask)
    results["OH_lt_SSQty"] = _detail(oh_lt_ss_mask)
    results["Demand_or_MTS_zero_OH"] = _detail(zero_oh_mask)
    return results


def _toggle_kpi_visibility(key: str) -> None:
    show_key = f"kpi_show_{key}"
    st.session_state[show_key] = not st.session_state.get(show_key, False)


kpi_data = _compute_kpis(filtered_df)
dioh_value = _calculate_dioh(filtered_df)
st.caption("Click a KPI card's button to toggle its detailed list.")
kpi_cols = st.columns(5)

with kpi_cols[0]:
    st.metric("DIOH", _format_days(dioh_value) if dioh_value is not None else "—")

kpi_labels = [
    ("MTS SKUs", "MTS"),
    ("MTO SKUs", "MTO"),
    ("< SSQty", "OH_lt_SSQty"),
    ("OOS", "Demand_or_MTS_zero_OH"),
]
for col, (label, key) in zip(kpi_cols[1:], kpi_labels):
    with col:
        count_value, _ = kpi_data.get(key, (0, pd.DataFrame()))
        st.metric(label, f"{count_value:,}")
        if st.button("View details", key=f"kpi_btn_{key}", use_container_width=True):
            _toggle_kpi_visibility(key)

for label, key in kpi_labels:
    show_key = f"kpi_show_{key}"
    if st.session_state.get(show_key):
        _, detail_df = kpi_data.get(key, (0, pd.DataFrame()))
        st.markdown(f"#### {label} details")
        if detail_df.empty:
            st.info("No matching items for this KPI.")
        else:
            st.dataframe(detail_df, use_container_width=True)


def _build_group_series(df: pd.DataFrame, group_col: str) -> Tuple[List[str], Dict[str, List[float]]]:
    labels: List[str] = []
    series_map: Dict[str, List[float]] = {metric: [] for metric in METRIC_ORDER}
    grouped = df.groupby(group_col, dropna=False)
    for group_value, group_df in grouped:
        label = str(group_value)
        labels.append(label)
        for metric in METRIC_ORDER:
            column_name = metric_column_map.get(metric)
            if column_name is None:
                series_map[metric].append(np.nan)
            else:
                total_value = pd.to_numeric(group_df[column_name], errors="coerce").sum()
                series_map[metric].append(total_value)
    return labels, series_map


aggregated_row = filtered_df.sum(numeric_only=True)

labels_by_market, series_by_market = _build_group_series(filtered_df, market_col)
_render_stacked_bar(
    labels_by_market,
    series_by_market,
    "Market totals (Cur metrics stack)",
    METRIC_ORDER,
    METRIC_LABELS
)

weekly_labels, weekly_series = _build_weekly_series(aggregated_row)
_render_stacked_bar(
    weekly_labels,
    weekly_series,
    "Cur month weekly breakdown",
    WEEKLY_METRICS,
    WEEKLY_METRIC_LABELS
)

monthly_labels, monthly_series = _build_monthly_series(aggregated_row)
monthly_filtered_series, monthly_order = _filter_monthly_series(monthly_series)
_render_line_chart(
    monthly_labels,
    monthly_filtered_series,
    "Monthly trajectory",
    monthly_order,
    MONTHLY_METRIC_LABELS
)


# ===========================
# DC stock map
# ===========================

dc_columns = [
    col
    for col in fg_df.columns
    if _normalize_name(col).startswith("dc") or _normalize_name(col).endswith("dc")
]
if dc_columns:
    dc_totals = (
        filtered_df[dc_columns]
        .apply(pd.to_numeric, errors="coerce")
        .sum()
    )
    dc_totals = dc_totals[dc_totals > 0]

    if not dc_totals.empty:
        st.markdown("## 🗺️ Stock by DC")
        dc_records: List[Dict[str, float | str]] = []
        for column, value in dc_totals.items():
            display_name = str(column).strip()
            normalized_name = _normalize_name(display_name)
            location_info = None
            for key, coords in DC_LOCATION_MAP.items():
                if _normalize_name(key) == normalized_name:
                    location_info = coords
                    break
                if _normalize_name(key) in normalized_name:
                    location_info = coords
                    break
                if key in display_name:
                    location_info = coords
                    break
            if location_info is None:
                continue
            dc_records.append(
                {
                    "Location": display_name,
                    "Stock": float(value),
                    "lat": location_info["lat"],
                    "lon": location_info["lon"],
                }
            )

        if dc_records:
            dc_df = pd.DataFrame(dc_records)
            selection_options = ["All"] + sorted(dc_df["Location"].unique())
            selected_location = st.selectbox(
                "Select DC location",
                options=selection_options,
                index=0,
                key="fg_dc_location_select",
            )
            if selected_location != "All":
                dc_display_df = dc_df[dc_df["Location"].eq(selected_location)]
            else:
                dc_display_df = dc_df

            fig_map = go.Figure(
                go.Scattergeo(
                    lon=dc_display_df["lon"],
                    lat=dc_display_df["lat"],
                    text=[f"{row.Location}<br>Stock: {row.Stock:,.0f}" for row in dc_display_df.itertuples()],
                    mode="markers",
                    marker=dict(
                        size=np.clip(dc_display_df["Stock"] / dc_display_df["Stock"].max() * 40, 6, 40),
                        color=dc_display_df["Stock"],
                        colorscale="Oranges",
                        reversescale=False,
                        colorbar=dict(title="Stock"),
                        sizemode="diameter",
                    ),
                )
            )
            fig_map.update_layout(
                title="DC Stock Distribution",
                geo=dict(
                    scope="africa",
                    projection=dict(type="mercator"),
                    center=dict(lat=27.0, lon=31.0),
                    lonaxis=dict(range=[24, 34]),
                    lataxis=dict(range=[22, 32]),
                    showland=True,
                    landcolor="#f0f0f0",
                    showcountries=True,
                    countrycolor="#999999",
                ),
                margin=dict(l=0, r=0, t=40, b=0),
            )
            st.plotly_chart(fig_map, use_container_width=True)
            st.dataframe(
                dc_display_df[["Location", "Stock"]]
                .sort_values("Stock", ascending=False)
                .style.format({"Stock": "{:,.0f}"}),
                use_container_width=True,
            )

            st.caption("Quick branch view:")
            st.map(
                dc_display_df.rename(columns={"lat": "Latitude", "lon": "Longitude"}),
                latitude="Latitude",
                longitude="Longitude",
                zoom=6,
                size=100,
            )


# ===========================
# Insights
# ===========================

st.markdown("## 🔎 Insights")

total_curst = pd.to_numeric(filtered_df[metric_column_map.get("CurST")], errors="coerce").sum() if metric_column_map.get("CurST") else 0
total_curas = pd.to_numeric(filtered_df[metric_column_map.get("CurAS")], errors="coerce").sum() if metric_column_map.get("CurAS") else 0
total_oh = pd.to_numeric(filtered_df[metric_column_map.get("CurOS")], errors="coerce").sum() if metric_column_map.get("CurOS") else 0
total_curapp = pd.to_numeric(filtered_df[metric_column_map.get("CurAPP")], errors="coerce").sum() if metric_column_map.get("CurAPP") else 0

insights: List[str] = []

if total_curst and total_curas:
    gap = total_curas - total_curst
    direction = "above" if gap >= 0 else "below"
    insights.append(
        f"- Actual sales are {_format_tonnage(abs(gap))} {direction} the sales target across the filtered selection."
    )

if total_curst and total_oh:
    dioh = (total_oh / total_curst) * 26
    insights.append(f"- Current DIOH stands at {_format_days(dioh)} for the selected scope.")

if total_curapp and total_curas:
    prod_gap = total_curapp - total_curas
    if abs(prod_gap) > 0:
        direction = "higher" if prod_gap >= 0 else "lower"
        insights.append(
            f"- Production plan is {_format_tonnage(abs(prod_gap))} {direction} than actual sales, signalling potential adjustments."
        )

if not insights:
    st.info("No insights available. Provide more data or adjust filters.")
else:
    for line in insights:
        st.markdown(line)
