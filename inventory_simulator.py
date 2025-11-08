import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from pathlib import Path
from datetime import datetime
from dateutil.relativedelta import relativedelta
import io
import re

st.set_page_config(page_title="Inventory Simulator", layout="wide")

# ===========================
# 1. Load Data
# ===========================
DEFAULT_DATA_PATH = Path(r"C:\Users\ashalaby\OneDrive - Halwani Bros\Planning - Sources\Materials.xlsb")
SHEET_NAME = "Data"
XLSXWRITER_AVAILABLE = True

@st.cache_data
def load_data(file_bytes: bytes | None = None):
    data_source = io.BytesIO(file_bytes) if file_bytes else DEFAULT_DATA_PATH
    df = pd.read_excel(data_source, sheet_name=SHEET_NAME, engine="pyxlsb")

    # Prepare ItemNumber and remove blank rows
    df["ItemNumber"] = df["ItemNumber"].astype(str).str.strip()
    valid_items_mask = df["ItemNumber"].ne("") & df["ItemNumber"].str.lower().ne("nan")
    df = df.loc[valid_items_mask].copy()
    df["ItemNumber"] = df["ItemNumber"].str.zfill(6)
    text_cols = ["ItemName", "Factory", "Storagetype", "Family", "RawType", "Unit"]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).fillna("")

    # Convert numeric columns and fill missing values with 0
    numeric_cols = ["OH", "Cost", "Next3M", "MINQTY", "SSDays", "FixedSSQty"] 
    months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    for m in months:
        numeric_cols += [f"{m}APP", f"{m}ST", f"{m}AS", f"{m}AP"]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


@st.cache_data
def load_calc_help(data_bytes: bytes | None = None, calc_bytes: bytes | None = None):
    """Load CalcHelp sheet either from explicit upload or from the main workbook."""

    def _read_excel(source_obj, sheet_name):
        engines = ["pyxlsb", None]
        errors = []
        for engine in engines:
            try:
                if hasattr(source_obj, "seek"):
                    source_obj.seek(0)
                return pd.read_excel(source_obj, sheet_name=sheet_name, engine=engine)
            except ValueError as exc:
                errors.append(exc)
            except ImportError:
                continue
        if errors:
            raise errors[-1]
        raise RuntimeError(f"Unable to read sheet '{sheet_name}'.")

    if calc_bytes:
        source = io.BytesIO(calc_bytes)
    elif data_bytes:
        source = io.BytesIO(data_bytes)
    else:
        source = DEFAULT_DATA_PATH

    for sheet_name in ("CalcHelp", "calchelp"):
        try:
            return _read_excel(source, sheet_name)
        except Exception:
            continue

    st.warning("CalcHelp sheet not found; inventory simulation will fall back to static logic.")
    return pd.DataFrame()


# ===========================
# 1.a Data Source Selection
# ===========================
st.sidebar.header("📂 Data Source")
uploaded_file = st.sidebar.file_uploader("Upload XLSB file", type=["xlsb"], help="Upload a replacement data file in XLSB format.")
uploaded_bytes = uploaded_file.getvalue() if uploaded_file else None
df = load_data(uploaded_bytes)
if uploaded_file:
    st.sidebar.success(f"Using uploaded file: {uploaded_file.name}")
else:
    st.sidebar.caption(f"Using default file: {DEFAULT_DATA_PATH.name}")

st.sidebar.markdown("---")
st.sidebar.subheader("🧮 CalcHelp Guidance")
calc_help_file = st.sidebar.file_uploader(
    "Upload CalcHelp sheet",
    type=["xlsb", "xlsx"],
    help="Optional: override the CalcHelp sheet used for monthly inventory logic.",
    key="calc_help_uploader"
)
calc_help_bytes = calc_help_file.getvalue() if calc_help_file else None
calc_help_df = load_calc_help(uploaded_bytes, calc_help_bytes)
if calc_help_file:
    st.sidebar.success(f"Using CalcHelp from: {calc_help_file.name}")
else:
    st.sidebar.caption("Using CalcHelp bundled with the active materials workbook.")

# ===========================
# Utility Helpers
# ===========================
def format_magnitude(value, suffix=""):
    """Format numeric values using thousands separators and M suffix for millions."""
    try:
        numeric_value = float(value)
    except (TypeError, ValueError):
        return "-"

    if np.isnan(numeric_value):
        return "-"

    if abs(numeric_value) >= 1_000_000:
        formatted = f"{numeric_value / 1_000_000:,.2f}M"
    else:
        formatted = f"{numeric_value:,.0f}"

    return f"{formatted}{suffix}"

def format_percentage(value, precision=1):
    """Format a ratio as percentage string with fixed precision."""
    try:
        numeric_value = float(value)
    except (TypeError, ValueError):
        return "-"

    if np.isnan(numeric_value):
        return "-"

    return f"{numeric_value * 100:.{precision}f}%"

def normalize_unit(unit_value):
    """Standardize unit strings, returning None for blanks/NaN."""
    if unit_value is None or (isinstance(unit_value, float) and np.isnan(unit_value)):
        return None
    unit_str = str(unit_value).strip()
    if not unit_str or unit_str.lower() == "nan":
        return None
    return unit_str

def _unique_preserve(sequence):
    seen = set()
    ordered = []
    for entry in sequence:
        if entry and entry not in seen:
            ordered.append(entry)
            seen.add(entry)
    return ordered


MONTH_ABBR_TO_NUM = {
    month: idx
    for idx, month in enumerate(["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], start=1)
}


def parse_season_window(season_text: str | None):
    """Parse a textual season window (e.g., 'Feb May') into (start_month, end_month)."""
    if not season_text:
        return None
    tokens = re.findall(r"[A-Za-z]{3}", str(season_text))
    if not tokens:
        return None
    start_token = tokens[0].title()
    end_token = tokens[-1].title()
    start_idx = MONTH_ABBR_TO_NUM.get(start_token[:3])
    end_idx = MONTH_ABBR_TO_NUM.get(end_token[:3])
    if not start_idx or not end_idx:
        return None
    return start_idx, end_idx


def is_month_in_season(month_number: int, season_window):
    """Return True if a month number falls within the given seasonal window."""
    if season_window is None:
        return True
    start_idx, end_idx = season_window
    if start_idx is None or end_idx is None:
        return True
    if start_idx <= end_idx:
        return start_idx <= month_number <= end_idx
    return month_number >= start_idx or month_number <= end_idx


def dataframe_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes | None:
    """Convert a DataFrame to Excel bytes, caching xlsxwriter availability."""
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


def render_excel_download_button(
    df: pd.DataFrame,
    label: str,
    filename_prefix: str,
    key: str,
    sheet_name: str = "Sheet1"
):
    excel_bytes = dataframe_to_excel_bytes(df, sheet_name=sheet_name)
    if excel_bytes:
        st.download_button(
            label=label,
            data=excel_bytes,
            file_name=f"{filename_prefix}_{today.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=key
        )


def parse_calc_instructions(calc_df: pd.DataFrame) -> list[dict]:
    """Transform CalcHelp sheet into sequential instruction steps."""
    if calc_df is None or calc_df.empty:
        return []

    df_norm = calc_df.copy()
    df_norm.columns = [str(col).strip() for col in df_norm.columns]

    start_col = next((c for c in df_norm.columns if c.lower().startswith("start")), None)
    target_col = next((c for c in df_norm.columns if c.lower().startswith("calc")), None)
    if not start_col or not target_col:
        return []

    op_cols = []
    for col in df_norm.columns:
        lower_col = col.lower()
        if col in {start_col, target_col}:
            continue
        if "plus" in lower_col:
            op_cols.append(("plus", col))
        elif "minus" in lower_col:
            op_cols.append(("minus", col))

    steps: list[dict] = []
    for _, row in df_norm.iterrows():
        start_ref = str(row.get(start_col, "")).strip()
        target_label = str(row.get(target_col, "")).strip()
        if not target_label:
            continue

        operations = []
        for op_type, col in op_cols:
            value = row.get(col, "")
            if pd.isna(value):
                continue
            ref = str(value).strip()
            if not ref or ref in {"0", "0.0"}:
                continue
            operations.append((op_type, ref))

        steps.append({"start": start_ref, "target": target_label, "ops": operations})

    return steps


def _calc_label_to_month_index(label: str | None) -> int | None:
    if not label:
        return None
    match = re.search(r"M(\d+)", str(label), flags=re.IGNORECASE)
    if not match:
        return None
    month_number = int(match.group(1))
    return month_number + 1  # Closing of M0 feeds Month_1, etc.


def parse_column_reference(reference: str | None):
    if not reference:
        return None
    ref = reference.strip()
    pattern = r"^(?P<month>[A-Za-z]{3})(?P<year>\d{2})?(?P<suffix>BP|APP|ST|AS|AP)(?P<tag>CY|NY)?$"
    match = re.match(pattern, ref, flags=re.IGNORECASE)
    if match:
        month_name = match.group("month").title()
        explicit_year = match.group("year")
        suffix = match.group("suffix").upper()
        tag = match.group("tag")

        # Determine month offset relative to today
        target_year = None
        if explicit_year:
            target_year = explicit_year
        elif tag:
            target_year = datetime.now().strftime("%y") if tag.upper() == "CY" else (datetime.now() + relativedelta(years=1)).strftime("%y")

        month_info = {
            "name": month_name,
            "year": target_year or datetime.now().strftime("%y"),
            "year_tag": tag.upper() if tag else None,
        }
        return suffix, month_info

    # Match labels like closing of M3
    if reference.lower().startswith("closing"):
        return None

    return None

def build_month_column_candidates(month_info, suffix, current_year_code):
    """Return ordered list of candidate source columns for a given month/suffix."""
    month_name = month_info["name"]
    month_year = month_info["year"]
    year_tag = month_info.get("year_tag")
    candidates = []

    if suffix == "ST":
        if year_tag:
            candidates.append(f"{month_name}ST{year_tag}")
        if month_year != current_year_code:
            candidates.append(f"{month_name}{month_year}ST")
        candidates.append(f"{month_name}ST")
        if month_year == current_year_code:
            candidates.append(f"{month_name}{month_year}ST")
    elif suffix == "APP":
        if year_tag:
            candidates.extend([f"{month_name}BP{year_tag}", f"{month_name}APP{year_tag}"])
        if month_year != current_year_code:
            candidates.extend([f"{month_name}{month_year}BP", f"{month_name}{month_year}APP"])
        candidates.extend([f"{month_name}BP", f"{month_name}APP"])
        if month_year == current_year_code:
            candidates.extend([f"{month_name}{month_year}BP", f"{month_name}{month_year}APP"])
    elif suffix in ("AS", "AP"):
        if year_tag:
            candidates.append(f"{month_name}{suffix}{year_tag}")
        if month_year != current_year_code:
            candidates.append(f"{month_name}{month_year}{suffix}")
        candidates.append(f"{month_name}{suffix}")
        if month_year == current_year_code:
            candidates.append(f"{month_name}{month_year}{suffix}")
    else:
        if year_tag:
            candidates.append(f"{month_name}{suffix}{year_tag}")
        candidates.append(f"{month_name}{suffix}")

    return _unique_preserve(candidates)

def resolve_existing_column(df_input, candidates):
    """Return actual column object from DataFrame matching first candidate name."""
    columns_map = {str(col): col for col in df_input.columns}
    for candidate in candidates:
        actual_col = columns_map.get(candidate)
        if actual_col is not None:
            return actual_col
    return None

# ===========================
# 2. Date Logic
# ===========================
today = datetime.now()
months_sequence = []
for i in range(12):
    target_date = today + relativedelta(months=i)
    months_sequence.append(
        {
            "code": target_date.strftime("%b%y"),
            "name": target_date.strftime("%b"),
            "year": target_date.strftime("%y"),
            "full_date": target_date,
            "year_tag": "CY" if target_date.year == today.year else "NY"
        }
    )
current_month_idx = 0  # Index for the current month using AS and AP
current_year_code = today.strftime("%y")


def _prioritize_candidates(candidate_list, preferred_suffix):
    ordered = []
    seen = set()
    preferred_suffix = preferred_suffix.upper()
    for cand in candidate_list:
        cand_str = str(cand)
        if cand_str not in seen and cand_str.upper().endswith(preferred_suffix):
            ordered.append(cand_str)
            seen.add(cand_str)
    for cand in candidate_list:
        cand_str = str(cand)
        if cand_str not in seen:
            ordered.append(cand_str)
            seen.add(cand_str)
    return ordered


def get_requirement_series(df_input, month_info, basis="APP", include_current_extras=False):
    basis = basis.upper()
    candidates = build_month_column_candidates(month_info, "APP", current_year_code)
    extras = []
    if include_current_extras:
        extras = [f"Cur{basis}"]
        fallback_basis = "BP" if basis == "APP" else "APP"
        extras.append(f"Cur{fallback_basis}")
    combined_candidates = extras + candidates
    prioritized = _prioritize_candidates(combined_candidates, basis)
    resolved_col = resolve_existing_column(df_input, prioritized)
    if resolved_col is not None:
        series = pd.to_numeric(df_input[resolved_col], errors="coerce").fillna(0)
        return series, str(resolved_col)
    return pd.Series(0, index=df_input.index, dtype=float), None


def build_requirement_plan(df_input, basis="APP"):
    if df_input is None or df_input.empty:
        return []

    plan = []
    for idx, month_info in enumerate(months_sequence):
        series, column_name = get_requirement_series(
            df_input,
            month_info,
            basis=basis,
            include_current_extras=(idx == 0)
        )
        plan.append({
            "month_idx": idx + 1,
            "month": month_info,
            "series": series,
            "source_column": column_name,
        })
    return plan


def get_existing_supply_series(df_input, month_info, include_current=False):
    candidates = build_month_column_candidates(month_info, "ST", current_year_code)
    if include_current:
        candidates = ["CurST"] + candidates
    resolved_col = resolve_existing_column(df_input, candidates)
    if resolved_col is not None:
        series = pd.to_numeric(df_input[resolved_col], errors="coerce").fillna(0)
        return series, str(resolved_col)
    return pd.Series(0, index=df_input.index, dtype=float), None


def build_existing_supply_plan(df_input):
    if df_input is None or df_input.empty:
        return []

    plan = []
    for idx, month_info in enumerate(months_sequence):
        series, column_name = get_existing_supply_series(
            df_input,
            month_info,
            include_current=(idx == 0)
        )
        plan.append({
            "month_idx": idx + 1,
            "month": month_info,
            "series": series,
            "source_column": column_name
        })
    return plan


def compute_static_ss_reference(df_input, bases):
    if df_input is None or df_input.empty or not bases:
        return None

    aggregate_series = []
    for basis in bases:
        basis_plan = build_requirement_plan(df_input, basis=basis)
        if not basis_plan:
            continue
        monthly_series = [entry["series"] for entry in basis_plan if not entry["series"].empty]
        if not monthly_series:
            continue
        combined = pd.concat(monthly_series, axis=1)
        aggregate_series.append(combined.mean(axis=1))

    if not aggregate_series:
        return None

    combined_reference = pd.concat(aggregate_series, axis=1).mean(axis=1)
    return combined_reference


def compute_supply_schedule(
    df_inventory: pd.DataFrame,
    requirement_plan: list[dict],
    safety_multiplier: float = 1.0,
    buffer_days: float = 0.0,
    round_to_minqty: bool = True,
    ss_mode: str = "dynamic",
    static_reference: pd.Series | None = None,
    existing_supply_plan: list[dict] | None = None
):
    if df_inventory is None or df_inventory.empty or not requirement_plan:
        return pd.DataFrame(), pd.DataFrame()

    schedule_df = df_inventory[[
        col for col in [
            "ItemNumber", "ItemName", "Unit", "Factory", "Cost", "SSDays",
            "MINQTY", "Month_0", "FixedSSQty", "Season"
        ] if col in df_inventory.columns
    ]].copy()

    numeric_cols = ["SSDays", "MINQTY", "Month_0", "Cost", "FixedSSQty"]
    for col in numeric_cols:
        if col in schedule_df.columns:
            schedule_df[col] = pd.to_numeric(schedule_df[col], errors="coerce").fillna(0)

    if "SSDays" not in schedule_df.columns:
        schedule_df["SSDays"] = 0.0
    if "MINQTY" not in schedule_df.columns:
        schedule_df["MINQTY"] = 0.0
    if "FixedSSQty" not in schedule_df.columns:
        schedule_df["FixedSSQty"] = 0.0
    if "Month_0" not in schedule_df.columns:
        schedule_df["Month_0"] = 0.0
    if "Cost" not in schedule_df.columns:
        schedule_df["Cost"] = 0.0
    if "Season" not in schedule_df.columns:
        schedule_df["Season"] = ""

    season_windows = schedule_df["Season"].apply(parse_season_window)
    seasonal_backlog = pd.Series(0, index=schedule_df.index, dtype=float)

    current_stock = schedule_df["Month_0"].astype(float)
    minqty_series = schedule_df["MINQTY"].astype(float)
    ssdays_series = schedule_df["SSDays"].astype(float)
    fixed_ss_series = schedule_df["FixedSSQty"].astype(float)
    cost_series = schedule_df["Cost"].astype(float)

    zero_series = pd.Series(0, index=schedule_df.index, dtype=float)

    if ss_mode.lower() == "static" and static_reference is not None:
        static_reference = pd.to_numeric(static_reference, errors="coerce").reindex(schedule_df.index).fillna(0)
    else:
        static_reference = None

    existing_supply_map: dict[int, pd.Series] = {}
    if existing_supply_plan:
        for entry in existing_supply_plan:
            series = pd.to_numeric(entry.get("series", zero_series), errors="coerce")
            existing_supply_map[entry.get("month_idx", 0)] = series.reindex(schedule_df.index).fillna(0)

    summary_rows = []

    for month_entry in requirement_plan:
        month_idx = month_entry["month_idx"]
        month_info = month_entry["month"]
        month_label = month_info["code"]
        calendar_month = month_info["full_date"].month
        demand_series = month_entry["series"].astype(float)

        if static_reference is not None:
            ss_base = static_reference
        else:
            ss_base = demand_series

        calculated_ss = ((ssdays_series + buffer_days) * ss_base / 26.0) * safety_multiplier
        ss_qty = calculated_ss.where(fixed_ss_series <= 0, fixed_ss_series)
        ss_qty = ss_qty.fillna(0)

        pre_scheduled_supply = existing_supply_map.get(month_idx, zero_series)
        pre_scheduled_supply = pre_scheduled_supply.astype(float).fillna(0)

        base_requirement = demand_series + ss_qty
        total_requirement = base_requirement + seasonal_backlog
        available_stock = current_stock + pre_scheduled_supply
        unmet_requirement = (total_requirement - available_stock).clip(lower=0)

        season_mask = season_windows.apply(lambda window: is_month_in_season(calendar_month, window))
        supply_candidate = unmet_requirement.where(season_mask, 0).copy()

        if round_to_minqty:
            minqty_positive = minqty_series.where(minqty_series > 0)
            valid_mask = supply_candidate.gt(0) & minqty_positive.notna()
            if valid_mask.any():
                supply_candidate.loc[valid_mask] = (
                    np.ceil(
                        (supply_candidate[valid_mask] / minqty_positive[valid_mask]).replace([np.inf, -np.inf], 0)
                    ) * minqty_positive[valid_mask]
                )
        seasonal_supply = supply_candidate.fillna(0)

        total_supply = pre_scheduled_supply + seasonal_supply
        available_post_supply = available_stock + seasonal_supply
        closing_stock = (available_post_supply - demand_series).clip(lower=0)
        seasonal_backlog = (total_requirement - available_post_supply).clip(lower=0)

        demand_value = demand_series * cost_series
        closing_value = closing_stock * cost_series
        ss_value = ss_qty * cost_series
        pre_supply_value = pre_scheduled_supply * cost_series
        new_supply_value = seasonal_supply * cost_series
        total_supply_value = total_supply * cost_series

        schedule_df[f"Demand_M{month_idx}"] = demand_series
        schedule_df[f"ExistingSupply_M{month_idx}"] = pre_scheduled_supply
        schedule_df[f"NewSupply_M{month_idx}"] = seasonal_supply
        schedule_df[f"Supply_M{month_idx}"] = total_supply
        schedule_df[f"Closing_M{month_idx}"] = closing_stock
        schedule_df[f"SS_Target_M{month_idx}"] = ss_qty
        schedule_df[f"ExistingSupplyValue_M{month_idx}"] = pre_supply_value
        schedule_df[f"NewSupplyValue_M{month_idx}"] = new_supply_value
        schedule_df[f"SupplyValue_M{month_idx}"] = total_supply_value
        schedule_df[f"DemandValue_M{month_idx}"] = demand_value
        schedule_df[f"ClosingValue_M{month_idx}"] = closing_value
        schedule_df[f"SSValue_M{month_idx}"] = ss_value

        summary_rows.append({
            "month_idx": month_idx,
            "Month": month_label,
            "DemandQty": demand_series.sum(),
            "ExistingSupplyQty": pre_scheduled_supply.sum(),
            "NewSupplyQty": seasonal_supply.sum(),
            "SupplyQty": total_supply.sum(),
            "ClosingQty": closing_stock.sum(),
            "SSTargetQty": ss_qty.sum(),
            "DemandValue": demand_value.sum(),
            "ExistingSupplyValue": pre_supply_value.sum(),
            "NewSupplyValue": new_supply_value.sum(),
            "SupplyValue": total_supply_value.sum(),
            "ClosingValue": closing_value.sum(),
            "SSTargetValue": ss_value.sum()
        })

        current_stock = closing_stock

    summary_df = pd.DataFrame(summary_rows).sort_values("month_idx")

    return schedule_df, summary_df

# ===========================
# 3. Sidebar Filters
# ===========================
st.sidebar.header("🔍 Filters")

# SKU search (temporarily disables other filters)
search_options = [f"{str(row['ItemNumber']).zfill(6)} - {row['ItemName']}" 
                  for _, row in df[["ItemNumber","ItemName"]].drop_duplicates().iterrows()]
item_search = st.sidebar.selectbox("Search for SKU", options=[""]+sorted(search_options))

df_filtered = df.copy()
search_active = False
if item_search:
    search_active = True
    item_code = item_search.split(" - ")[0]
    df_filtered = df_filtered[df_filtered["ItemNumber"]==item_code]
    st.sidebar.success("✅ SKU selected")

# Other filters
if not search_active:
    st.sidebar.markdown("---")
    st.sidebar.subheader("Or use filters:")
    
    all_option = "All"

    # Factory filter
    factories = [all_option]+sorted(df["Factory"].unique())
    selected_factory = st.sidebar.multiselect("Select factory", factories, default=[all_option])
    if all_option not in selected_factory:
        df_filtered = df_filtered[df_filtered["Factory"].isin(selected_factory)]

    # Storage type filter (depends on previous selection)
    storages = [all_option]+sorted(df_filtered["Storagetype"].unique())
    selected_storage = st.sidebar.multiselect("Select storage type", storages, default=[all_option])
    if all_option not in selected_storage:
        df_filtered = df_filtered[df_filtered["Storagetype"].isin(selected_storage)]

    # Family filter
    families = [all_option]+sorted(df_filtered["Family"].unique())
    selected_family = st.sidebar.multiselect("Select family", families, default=[all_option])
    if all_option not in selected_family:
        df_filtered = df_filtered[df_filtered["Family"].isin(selected_family)]

    # Raw material type filter
    rawtypes = [all_option]+sorted(df_filtered["RawType"].unique())
    selected_rawtype = st.sidebar.multiselect("Select raw material type", rawtypes, default=[all_option])
    if all_option not in selected_rawtype:
        df_filtered = df_filtered[df_filtered["RawType"].isin(selected_rawtype)]

# Additional options
time_grouping = st.sidebar.radio("📆 Time view:", ["Monthly","Quarterly","Yearly"], index=0)
use_next3m = st.sidebar.checkbox("📈 Use Next 3 Months forecast (Next3M)", value=True)

# ===========================
# 4. Calculations
# ===========================
def calculate_monthly_inventory_static(df_input):
    df_calc = df_input.copy()
    df_calc["OH"] = pd.to_numeric(df_calc.get("OH", 0), errors="coerce").fillna(0)
    df_calc["Cost"] = pd.to_numeric(df_calc.get("Cost", 0), errors="coerce").fillna(0)
    df_calc["Month_0"] = df_calc["OH"]
    df_calc["Month_0_value"] = df_calc["Month_0"] * df_calc["Cost"]

    def get_column_series(df_source, candidates):
        resolved_col = resolve_existing_column(df_source, candidates)
        if resolved_col is not None:
            df_source.loc[:, resolved_col] = pd.to_numeric(df_source[resolved_col], errors="coerce").fillna(0)
            return df_source[resolved_col]
        return pd.Series(0, index=df_source.index, dtype=float)

    for i, month_info in enumerate(months_sequence):
        month_col = f"Month_{i+1}"
        month_val = f"{month_col}_value"
        st_candidates = build_month_column_candidates(month_info, "ST", current_year_code)
        app_candidates = build_month_column_candidates(month_info, "APP", current_year_code)
        as_candidates = build_month_column_candidates(month_info, "AS", current_year_code)
        ap_candidates = build_month_column_candidates(month_info, "AP", current_year_code)

        previous_month_col = f"Month_{i}"

        if i == current_month_idx:
            st_series = get_column_series(df_calc, ["CurST"] + st_candidates)
            app_series = get_column_series(df_calc, ["CurAPP"] + app_candidates)
            as_series = get_column_series(df_calc, ["CurAS"] + as_candidates)
            ap_series = get_column_series(df_calc, ["CurAP"] + ap_candidates)
            df_calc[month_col] = df_calc[previous_month_col] + (st_series - as_series) - (app_series - ap_series)
        else:
            st_series = get_column_series(df_calc, st_candidates)
            app_series = get_column_series(df_calc, app_candidates)
            st_app_diff = st_series - app_series

            if use_next3m and "Next3M" in df_calc.columns and i < 3:
                next3m_series = pd.to_numeric(df_calc["Next3M"], errors="coerce").fillna(0)
                df_calc[month_col] = df_calc[previous_month_col] + next3m_series
            else:
                df_calc[month_col] = df_calc[previous_month_col] + st_app_diff

        df_calc[month_val] = df_calc[month_col] * df_calc["Cost"]

    return df_calc


def calculate_monthly_inventory(df_input, calc_help):
    df_calc = df_input.copy()
    df_calc["OH"] = pd.to_numeric(df_calc.get("OH", 0), errors="coerce").fillna(0)
    df_calc["Cost"] = pd.to_numeric(df_calc.get("Cost", 0), errors="coerce").fillna(0)
    df_calc["Month_0"] = df_calc["OH"]
    df_calc["Month_0_value"] = df_calc["Month_0"] * df_calc["Cost"]

    steps = parse_calc_instructions(calc_help)
    if not steps:
        return calculate_monthly_inventory_static(df_input)

    zero_series = pd.Series(0, index=df_calc.index, dtype=float)
    calculated: dict[str, pd.Series] = {"OH": df_calc["OH"], "Month_0": df_calc["Month_0"]}

    def resolve_reference(reference: str | None) -> pd.Series:
        if reference is None:
            return zero_series
        ref = reference.strip()
        if not ref or ref.lower() == "nan" or ref in {"0", "0.0"}:
            return zero_series

        if ref.lower() == "next3m" and not use_next3m:
            return zero_series

        cached = calculated.get(ref)
        if cached is not None:
            return cached

        if ref.lower() == "next3m" and "Next3M" in df_calc.columns:
            series = pd.to_numeric(df_calc["Next3M"], errors="coerce").fillna(0)
            calculated[ref] = series
            return series

        if ref in df_calc.columns:
            df_calc.loc[:, ref] = pd.to_numeric(df_calc[ref], errors="coerce").fillna(0)
            series = df_calc[ref]
            calculated[ref] = series
            return series

        buildable = parse_column_reference(ref)
        if buildable is not None:
            kind, month_info = buildable
            candidates = build_month_column_candidates(month_info, kind, current_year_code)
            resolved_col = resolve_existing_column(df_calc, candidates)
            if resolved_col is not None:
                df_calc.loc[:, resolved_col] = pd.to_numeric(df_calc[resolved_col], errors="coerce").fillna(0)
                series = df_calc[resolved_col]
                calculated[ref] = series
                return series

        return zero_series

    for step in steps:
        start_series = resolve_reference(step.get("start"))
        result_series = start_series.copy()
        for op_type, ref in step.get("ops", []):
            operand = resolve_reference(ref)
            if op_type == "plus":
                result_series = result_series + operand
            elif op_type == "minus":
                result_series = result_series - operand

        target_label = step.get("target")
        if not target_label:
            continue

        month_index = _calc_label_to_month_index(target_label)
        if month_index is not None:
            month_col = f"Month_{month_index}"
            df_calc[month_col] = result_series
            calculated[target_label] = result_series
            calculated[month_col] = result_series
            df_calc[f"{month_col}_value"] = result_series * df_calc["Cost"]
        else:
            safe_label = target_label.strip().replace(" ", "_")
            df_calc[safe_label] = result_series
            calculated[target_label] = result_series

    # Ensure value columns exist for any produced Month_X series
    columns_map = {str(col): col for col in df_calc.columns}
    for col in list(df_calc.columns):
        col_str = str(col)
        if col_str.startswith("Month_") and not col_str.endswith("_value"):
            value_col_str = f"{col_str}_value"
            if value_col_str not in columns_map:
                df_calc[value_col_str] = df_calc[col] * df_calc["Cost"]
                columns_map[value_col_str] = value_col_str

    return df_calc

df_with_months = calculate_monthly_inventory(df_filtered, calc_help_df)

def group_months(df_input, mode="Monthly"):
    df_grouped = df_input.copy()
    # Since inventory is cumulative, quarter/year values use the closing month of the period
    if mode=="Quarterly":
        quarters_months = {"Q1": "Month_3", "Q2": "Month_6", "Q3": "Month_9", "Q4": "Month_12"}
        for q, last_month in quarters_months.items():
            df_grouped[q] = df_grouped[last_month] 
            df_grouped[f"{q}_value"] = df_grouped[f"{last_month}_value"]
            
    elif mode=="Yearly":
        df_grouped["Year_Total"] = df_grouped["Month_12"]
        df_grouped["Year_Total_value"] = df_grouped["Month_12_value"]
        
    return df_grouped

df_grouped = group_months(df_with_months, mode=time_grouping)

def render_inventory_dashboard():
    # ===========================
    # 5. Dashboard
    # ===========================
    st.title("📦 Inventory Simulator")
    st.markdown(f"**Date:** {today.strftime('%d %B %Y')}")

    # استخدام الفاصلة للألوف في الأرقام
    col1,col2,col3,col4,col5 = st.columns(5)
    total_oh_value = (df_filtered['OH']*df_filtered['Cost']).sum()
    predicted_value_total = df_with_months['Month_1_value'].sum()
    unique_item_count = df_filtered["ItemNumber"].nunique()

    units_clean = []
    if "Unit" in df_filtered.columns:
        units_clean = [normalize_unit(u) for u in df_filtered["Unit"].unique()]
        units_clean = [u for u in units_clean if u]

    single_unit_selected = len(set(units_clean)) == 1 and units_clean
    unit_label = units_clean[0] if single_unit_selected else None

    can_sum_quantities = search_active or unique_item_count == 1 or single_unit_selected

    predicted_qty_total = df_with_months['Month_1'].sum() if can_sum_quantities else np.nan
    share_predicted = (predicted_value_total / total_oh_value) if total_oh_value else np.nan

    # DIOH (Days Inventory on Hand) calculation
    dioh_days = None
    dioh_basis_label = None
    if can_sum_quantities:
        cur_month_info = months_sequence[0]
        cur_candidates = ["CurAPP", "CurBP"] + build_month_column_candidates(cur_month_info, "APP", current_year_code)
        resolved_cur = resolve_existing_column(df_filtered, cur_candidates)
        if resolved_cur is not None:
            cur_series = pd.to_numeric(df_filtered[resolved_cur], errors="coerce").fillna(0)
            cur_demand_total = cur_series.sum()
            oh_qty_total = pd.to_numeric(df_filtered["OH"], errors="coerce").fillna(0).sum()
            if cur_demand_total > 0 and oh_qty_total > 0:
                dioh_days = (oh_qty_total / cur_demand_total) * 26
                dioh_basis_label = str(resolved_cur)

    if dioh_days is None:
        cost_series = pd.to_numeric(df_filtered.get("Cost", 0), errors="coerce").fillna(0)
        requirement_preview = build_requirement_plan(df_filtered, basis="APP")
        if requirement_preview:
            next_month_plan = requirement_preview[0]
            demand_series_val = next_month_plan["series"].reindex(df_filtered.index).fillna(0)
            demand_value = (demand_series_val * cost_series).sum()
            if demand_value > 0 and total_oh_value > 0:
                dioh_days = (total_oh_value / demand_value) * 26
                dioh_basis_label = f"{next_month_plan.get('source_column') or next_month_plan['month']['code']} (value-based)"


    with col1:
        st.metric("📦 Filtered SKUs", format_magnitude(len(df_filtered)))
    with col2:
        st.metric("💰 Current OH Value", format_magnitude(total_oh_value, " EGP"))
        total_value_caption = format_percentage(1.0) if total_oh_value else "-"
        st.caption(f"{total_value_caption} of filtered OH value")
    with col3:
        st.metric("🔮 Projected Value", format_magnitude(predicted_value_total, " EGP"))
        share_caption = format_percentage(share_predicted) if total_oh_value else "-"
        st.caption(f"{share_caption} vs. current OH value")
    with col4:
        if can_sum_quantities:
            qty_suffix = f" {unit_label}" if unit_label else " units"
            st.metric("📦 Projected Quantity", format_magnitude(predicted_qty_total, qty_suffix))
        else:
            st.metric("📦 Projected Quantity", "—")
            if single_unit_selected:
                st.caption("Select a specific SKU to view quantity")
            else:
                st.caption("Select a single unit or SKU to view quantity")

    with col5:
        if dioh_days is not None and np.isfinite(dioh_days):
            st.metric("📆 Inventory Days", f"{dioh_days:,.1f} days")
            if dioh_basis_label:
                st.caption(f"Based on {dioh_basis_label} demand")
        else:
            st.metric("📆 Inventory Days", "—")
            st.caption("Requires consistent demand data")

    # ===========================
    # 5.a Factory OH Cards
    # ===========================
    factory_cards_df = df_filtered.copy()
    if not factory_cards_df.empty:
        factory_cards_df["Factory"] = factory_cards_df["Factory"].fillna("Unspecified")
        factory_cards_df["OH_value"] = factory_cards_df["OH"] * factory_cards_df["Cost"]
        factory_cards_summary = factory_cards_df.groupby("Factory", as_index=False).agg({
            "OH_value": "sum",
            "OH": "sum"
        }).sort_values("OH_value", ascending=False)

        total_factory_oh_value = factory_cards_summary["OH_value"].sum()

        st.markdown("### 🏭 Factory OH Value")

        card_chunk_size = 4
        for start in range(0, len(factory_cards_summary), card_chunk_size):
            chunk = factory_cards_summary.iloc[start:start + card_chunk_size]
            cols = st.columns(len(chunk))
            for col, (_, row) in zip(cols, chunk.iterrows()):
                with col:
                    share = row['OH_value'] / total_factory_oh_value if total_factory_oh_value else np.nan
                    delta_text = format_percentage(share) if not np.isnan(share) else "-"
                    st.metric(
                        label=f"🏭 {row['Factory']}",
                        value=format_magnitude(row['OH_value'], " EGP"),
                        delta=delta_text
                    )
                    st.caption("Share of filtered OH value")
    else:
        st.info("No factory data available to display.")

    return dioh_days, dioh_basis_label

def render_abc_analysis():
    # ===========================
    # 6. ABC Analysis
    # ===========================
    st.markdown("---")
    st.subheader("🧮 ABC Inventory Classification")

    abc_enabled = st.checkbox("Enable ABC analysis on filtered data", value=False)
    if not abc_enabled:
        return

    if len(months_sequence) < 2:
        st.warning("Insufficient monthly history to compute next-month consumption.")
        return

    next_month_info = months_sequence[1]
    next_month_candidates = build_month_column_candidates(next_month_info, "APP", current_year_code)
    resolved_next_month_col = resolve_existing_column(df_filtered, next_month_candidates)
    if resolved_next_month_col is None:
        available_app_cols = sorted({
            str(col) for col in df_filtered.columns
            if str(col).endswith(("APP", "BP"))
        })
        tried_labels = ", ".join(next_month_candidates)
        st.warning(
            "No matching consumption column found for next month. "
            f"Tried: {tried_labels or 'none'}. Available APP/BP columns: {', '.join(available_app_cols) or 'none'}."
        )
        return

    abc_base = df_filtered.copy()
    next_month_series_name = str(resolved_next_month_col)
    abc_base["NextMonthAPP"] = pd.to_numeric(abc_base[resolved_next_month_col], errors="coerce").fillna(0)
    abc_base["ConsumptionValue"] = abc_base["NextMonthAPP"] * abc_base["Cost"]
    abc_base["OH_value"] = abc_base["OH"] * abc_base["Cost"]

    if abc_base["ConsumptionValue"].sum() <= 0:
        st.info("Consumption value is zero for the filtered selection; unable to build ABC profile.")
        return

    group_columns = ["ItemNumber", "ItemName"]
    if "Factory" in abc_base.columns:
        group_columns.append("Factory")
    if "Unit" in abc_base.columns:
        group_columns.append("Unit")

    agg_dict = {
        "ConsumptionValue": "sum",
        "NextMonthAPP": "sum",
        "OH": "sum",
        "OH_value": "sum",
        "Cost": "mean"
    }
    for param in ["LT", "MINQTY", "SSDays"]:
        if param in abc_base.columns:
            agg_dict[param] = "first"

    abc_summary = (
        abc_base.groupby(group_columns, dropna=False)
        .agg(agg_dict)
        .reset_index()
    )

    total_consumption_value = abc_summary["ConsumptionValue"].sum()
    abc_summary["Share"] = abc_summary["ConsumptionValue"] / total_consumption_value
    abc_summary = abc_summary.sort_values("ConsumptionValue", ascending=False).reset_index(drop=True)
    abc_summary["CumulativeShare"] = abc_summary["Share"].cumsum()

    conditions = [
        abc_summary["CumulativeShare"] <= 0.8,
        abc_summary["CumulativeShare"] <= 0.95
    ]
    choices = ["A", "B"]
    abc_summary["ABC_Class"] = np.select(conditions, choices, default="C")

    if "SSDays" in abc_summary.columns:
        abc_summary["SSDays"] = pd.to_numeric(abc_summary["SSDays"], errors="coerce")
        abc_summary["SSQty"] = abc_summary["SSDays"].fillna(0) * (abc_summary["NextMonthAPP"] / 26.0)
    else:
        abc_summary["SSQty"] = np.nan

    abc_summary["GapToSS"] = abc_summary["OH"] - abc_summary["SSQty"]

    top_abc_a = abc_summary[abc_summary["ABC_Class"] == "A"].copy()

    col_abc1, col_abc2, col_abc3 = st.columns(3)
    with col_abc1:
        st.metric("A-class SKUs", f"{len(top_abc_a):,}")
    with col_abc2:
        value_share = top_abc_a["Share"].sum() if not top_abc_a.empty else np.nan
        st.metric("A-class value share", format_percentage(value_share))
    with col_abc3:
        avg_lt = top_abc_a["LT"].mean() if ("LT" in top_abc_a.columns and not top_abc_a.empty) else np.nan
        lt_text = f"{avg_lt:.1f} days" if pd.notna(avg_lt) else "N/A"
        st.metric("Avg. lead time (A)", lt_text)

    insights = []
    if not top_abc_a.empty:
        total_a_value = top_abc_a["OH_value"].sum()
        insights.append(f"- {len(top_abc_a):,} SKU(s) in class A cover {format_percentage(top_abc_a['Share'].sum())} of consumption value; prioritize supply adherence for these items.")
        if np.isfinite(total_a_value):
            insights.append(f"- Current OH value for A-class items totals {format_magnitude(total_a_value, ' EGP')}. Review replenishment cadence to stay within safety stock targets.")
    if "GapToSS" in abc_summary.columns and not top_abc_a.empty:
        excess_df = top_abc_a[top_abc_a["GapToSS"] > 0]
        short_df = top_abc_a[top_abc_a["GapToSS"] < 0]
        if not excess_df.empty:
            excess_units = excess_df["GapToSS"].sum()
            insights.append(f"- {len(excess_df)} A-class SKU(s) exceed safety stock by {format_magnitude(excess_units, ' units')}; adjust MINQTY releases or defer purchases.")
        if not short_df.empty:
            short_units = abs(short_df["GapToSS"].sum())
            insights.append(f"- {len(short_df)} A-class SKU(s) fall below safety stock by {format_magnitude(short_units, ' units')}; expedite supply or review lead times.")

    insights.append(
        f"- Safety stock quantity computed as SSDays × ({next_month_series_name} ÷ 26). "
        "Align planning parameters (LT, MINQTY, SSDays) to control inventory levels."
    )

    st.markdown("**Operational insights**")
    for tip in insights:
        st.markdown(tip)

    display_cols = group_columns + [
        "ConsumptionValue", "Share", "CumulativeShare", "ABC_Class",
        "NextMonthAPP", "SSQty", "GapToSS", "OH", "OH_value", "Cost"
    ]
    for optional_col in ["LT", "MINQTY", "SSDays"]:
        if optional_col in abc_summary.columns and optional_col not in display_cols:
            display_cols.append(optional_col)

    abc_display = abc_summary[display_cols]
    formatters = {
        "ConsumptionValue": "{:,.0f}",
        "Share": "{:.2%}",
        "CumulativeShare": "{:.2%}",
        "NextMonthAPP": "{:,.0f}",
        "SSQty": "{:,.0f}",
        "GapToSS": "{:,.0f}",
        "OH": "{:,.0f}",
        "OH_value": "{:,.0f}",
        "Cost": "{:,.2f}"
    }
    optional_formatters = {"LT": "{:,.0f}", "MINQTY": "{:,.0f}", "SSDays": "{:,.0f}"}
    for key, fmt in optional_formatters.items():
        if key in abc_display.columns:
            formatters[key] = fmt

    st.dataframe(
        abc_display.style.format(formatters, na_rep="-"),
        use_container_width=True
    )
    render_excel_download_button(
        abc_display,
        "📥 Download ABC Summary (Excel)",
        "inventory_abc_summary",
        key="download_abc_summary",
        sheet_name="ABC"
    )

    if not top_abc_a.empty:
        st.markdown("**Items covering ~80% of consumption value (Class A)**")
        top_display = top_abc_a[display_cols]
        top_formatters = {k: v for k, v in formatters.items() if k in top_display.columns}
        st.dataframe(
            top_display.style.format(top_formatters, na_rep="-"),
            use_container_width=True
        )
        render_excel_download_button(
            top_display,
            "📥 Download ABC Class A (Excel)",
            "inventory_abc_class_a",
            key="download_abc_class_a",
            sheet_name="ABC_Class_A"
        )

    st.caption(f"ABC analysis uses column '{next_month_series_name}' for next-month consumption.")

dioh_days, dioh_basis_label = render_inventory_dashboard()
render_abc_analysis()

# ===========================
# 7. Summary Table
# ===========================
if not search_active:
    grouping_cols = ["Factory","Storagetype","Family","RawType"]
    summary_base = df_filtered.copy()
    summary_base["OH_value"] = summary_base["OH"] * summary_base["Cost"]
    summary = summary_base.groupby(grouping_cols, as_index=False).agg({
        "ItemNumber": "nunique",
        "OH_value": "sum"
    }).rename(columns={"ItemNumber": "Unique SKUs"})

    month1_values = df_with_months.groupby(grouping_cols, as_index=False)[["Month_1_value"]].sum()
    summary = summary.merge(month1_values, on=grouping_cols, how="left")
    summary = summary.rename(columns={
        "OH_value": "OH Value (EGP)",
        "Month_1_value": "Month_1 Value (EGP)"
    })
    summary["Month_1 Value (EGP)"] = summary["Month_1 Value (EGP)"].fillna(0)

    st.dataframe(
        summary.style.format({
            "Unique SKUs": "{:,.0f}",
            "OH Value (EGP)": "{:,.0f}",
            "Month_1 Value (EGP)": "{:,.0f}"
        }),
        use_container_width=True
    )
    render_excel_download_button(
        summary,
        "📥 Download Summary (Excel)",
        "inventory_summary",
        key="download_summary_excel",
        sheet_name="Summary"
    )

# ===========================
# 8. Supply Planning Scenarios
# ===========================
st.markdown("---")
st.subheader("🚚 Supply Scheduling to Protect Safety Stock")

with st.expander("Configure supply planning scenario", expanded=False):
    col_basis, col_ssmult, col_buffer, col_round = st.columns([2,2,2,2])

    with col_basis:
        supply_basis = st.selectbox(
            "Requirement basis",
            options=["APP", "BP"],
            index=0,
            help="Choose which requirement column (APP vs BP) to use for planning future supply."
        )

    with col_ssmult:
        safety_multiplier = st.slider(
            "Safety stock multiplier",
            min_value=0.5,
            max_value=2.0,
            value=1.0,
            step=0.1,
            help="Increase to add extra buffer beyond SSDays-derived safety stock."
        )

    with col_buffer:
        additional_ss_days = st.number_input(
            "Extra SS days",
            min_value=0.0,
            max_value=30.0,
            value=0.0,
            step=1.0,
            help="Apply an additional buffer (in days) to the SSDays before converting to quantity."
        )

    with col_round:
        round_to_minqty = st.checkbox(
            "Round supply to MINQTY",
            value=True,
            help="If enabled, supply quantities are rounded up to the nearest MINQTY per SKU."
        )

    col_ssmode, col_ssbasis = st.columns([2,3])
    with col_ssmode:
        ss_mode = st.selectbox(
            "Safety stock calculation",
            options=["Dynamic per month", "Static average"],
            index=0,
            help="Dynamic recalculates safety stock from each month's demand; static uses an average demand reference."
        )

    with col_ssbasis:
        static_basis_options = st.multiselect(
            "Static SS reference (use averages of)",
            options=["APP", "BP"],
            default=["APP", "BP"],
            help="Select which requirement columns to average when using static safety stock mode."
        )

    use_existing_supply = st.checkbox(
        "Incorporate existing ST (pre-scheduled supply)",
        value=False,
        help="If enabled, existing ST columns are treated as committed supply before calculating additional quantities."
    )

    st.caption(
        "Supply calculations use the filtered dataset only. Safety stock target is SSDays × demand/26. "
        "Demand basis columns are resolved dynamically based on the selected APP/BP preference."
    )

requirement_plan = build_requirement_plan(df_filtered, basis=supply_basis)
ss_reference_series = None
ss_mode_key = "dynamic"
if ss_mode == "Static average" and static_basis_options:
    ss_reference_series = compute_static_ss_reference(df_filtered, bases=static_basis_options)
    ss_mode_key = "static"

existing_supply_plan = build_existing_supply_plan(df_filtered) if use_existing_supply else None

schedule_detail, schedule_summary = compute_supply_schedule(
    df_with_months,
    requirement_plan,
    safety_multiplier=safety_multiplier,
    buffer_days=additional_ss_days,
    round_to_minqty=round_to_minqty,
    ss_mode=ss_mode_key,
    static_reference=ss_reference_series,
    existing_supply_plan=existing_supply_plan
)

if schedule_summary.empty:
    st.info("Insufficient data to produce a supply schedule for the current filters.")
else:
    display_mode = st.radio(
        "Display mode",
        options=["Quantity", "Value"],
        index=0,
        horizontal=True,
        help="Toggle between unit-based and value-based supply planning views."
    )

    is_value_mode = display_mode == "Value"

    # DIOH indicators for supply view (always quantity-based)
    summary_sorted = schedule_summary.sort_values("month_idx").reset_index(drop=True)
    planned_dioh_days = None
    planned_dioh_label = None
    if not summary_sorted.empty:
        last_row = summary_sorted.iloc[-1]
        planned_base_demand = np.nan
        if len(summary_sorted) > 1:
            planned_base_demand = summary_sorted.iloc[-1]["DemandQty"]
        else:
            planned_base_demand = last_row["DemandQty"]

        closing_qty = last_row.get("ClosingQty", np.nan)
        if pd.notna(closing_qty) and planned_base_demand and planned_base_demand > 0:
            planned_dioh_days = (closing_qty / planned_base_demand) * 26
            planned_dioh_label = last_row.get("Month")

    dioh_cols = st.columns(2)
    with dioh_cols[0]:
        if dioh_days is not None and np.isfinite(dioh_days):
            st.markdown(f"🕒 **Current DIOH:** {dioh_days:,.1f} days")
        else:
            st.markdown("🕒 **Current DIOH:** —")

    with dioh_cols[1]:
        if planned_dioh_days is not None and np.isfinite(planned_dioh_days):
            label_suffix = f" ({planned_dioh_label})" if planned_dioh_label else ""
            st.markdown(f"🔄 **Planned DIOH{label_suffix}:** {planned_dioh_days:,.1f} days")
        else:
            st.markdown("🔄 **Planned DIOH:** —")

    st.markdown("### Supply vs Demand Trajectory")
    supply_chart = go.Figure()
    demand_field = "DemandValue" if is_value_mode else "DemandQty"
    existing_supply_field = "ExistingSupplyValue" if is_value_mode else "ExistingSupplyQty"
    new_supply_field = "NewSupplyValue" if is_value_mode else "NewSupplyQty"
    total_supply_field = "SupplyValue" if is_value_mode else "SupplyQty"
    ss_field = "SSTargetValue" if is_value_mode else "SSTargetQty"

    supply_chart.add_trace(
        go.Bar(
            x=schedule_summary["Month"],
            y=schedule_summary[demand_field],
            name="Demand",
            marker_color="#EF553B"
        )
    )
    if use_existing_supply:
        supply_chart.add_trace(
            go.Bar(
                x=schedule_summary["Month"],
                y=schedule_summary[existing_supply_field],
                name="Existing supply",
                marker_color="#636EFA"
            )
        )
    supply_chart.add_trace(
        go.Bar(
            x=schedule_summary["Month"],
            y=schedule_summary[new_supply_field if use_existing_supply else total_supply_field],
            name="Calculated supply" if use_existing_supply else "Supply",
            marker_color="#00CC96"
        )
    )
    supply_chart.add_trace(
        go.Scatter(
            x=schedule_summary["Month"],
            y=schedule_summary[ss_field],
            name="Safety stock target",
            mode="lines+markers",
            line=dict(color="#636EFA", dash="dash"),
            hovertemplate=(
                "<b>%{x}</b><br>SS Target: %{y:,.0f}" + (" EGP" if is_value_mode else " units") + "<extra></extra>"
            )
        )
    )
    supply_chart.add_trace(
        go.Scatter(
            x=schedule_summary["Month"],
            y=schedule_summary["ClosingValue" if is_value_mode else "ClosingQty"],
            name="Projected closing stock",
            mode="lines+markers",
            line=dict(color="#FFA15A"),
            hovertemplate=(
                "<b>%{x}</b><br>Closing: %{y:,.0f}" + (" EGP" if is_value_mode else " units") + "<extra></extra>"
            )
        )
    )
    supply_chart.update_layout(
        barmode="group",
        yaxis_title="Value (EGP)" if is_value_mode else "Quantity",
        xaxis_title="Month",
        height=500,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    st.plotly_chart(supply_chart, use_container_width=True)

    st.markdown("### Scenario Overview")
    scenario_cols = st.columns([2,1])
    with scenario_cols[0]:
        summary_download_df = schedule_summary.copy()
        if is_value_mode:
            columns_to_show = [
                "Month", "DemandValue", "SupplyValue", "ClosingValue", "SSTargetValue"
            ]
            format_cols = ["DemandValue", "SupplyValue", "ClosingValue", "SSTargetValue"]
        else:
            columns_to_show = [
                "Month", "DemandQty", "SupplyQty", "ClosingQty", "SSTargetQty"
            ]
            format_cols = ["DemandQty", "SupplyQty", "ClosingQty", "SSTargetQty"]
        summary_download_df = summary_download_df[columns_to_show]
        summary_display = summary_download_df.set_index("Month").copy()
        summary_display[format_cols] = summary_display[format_cols].applymap(lambda x: f"{x:,.0f}")
        st.dataframe(summary_display, use_container_width=True)
        render_excel_download_button(
            summary_download_df,
            "📥 Download Scenario Overview (Excel)",
            "supply_scenario_overview",
            key="download_scenario_overview",
            sheet_name="Scenario"
        )

    with scenario_cols[1]:
        st.markdown("#### Totals")
        if is_value_mode:
            st.metric("Total supply value", format_magnitude(schedule_summary["SupplyValue"].sum(), " EGP"))
            st.metric("Projected closing value", format_magnitude(schedule_summary["ClosingValue"].sum(), " EGP"))
        else:
            st.metric("Total planned supply", format_magnitude(schedule_summary["SupplyQty"].sum(), " units"))
            st.metric("Projected closing stock", format_magnitude(schedule_summary["ClosingQty"].sum(), " units"))

    st.markdown("### Detailed Supply Plan")
    base_columns = [col for col in [
        "ItemNumber", "ItemName", "Unit", "Factory", "Month_0"
    ] if col in schedule_detail.columns]

    if is_value_mode:
        value_cols = [col for col in schedule_detail.columns if col.endswith("Value")]
        detail_columns = base_columns + value_cols
    else:
        quantity_prefixes = ("Demand_", "ExistingSupply_", "NewSupply_", "Supply_", "Closing_", "SS_Target_")
        detail_columns = base_columns + [
            col for col in schedule_detail.columns
            if col.startswith(quantity_prefixes) and not col.endswith("Value")
        ]

    detail_df = schedule_detail[detail_columns].copy()

    def _format_cell(value):
        if isinstance(value, (int, float)):
            return f"{value:,.0f}"
        return value

    st.dataframe(
        detail_df.style.format(_format_cell, na_rep="-"),
        use_container_width=True,
        height=400
    )
    render_excel_download_button(
        detail_df,
        "📥 Download Supply Detail (Excel)",
        "supply_plan_detail",
        key="download_supply_detail",
        sheet_name="Supply_Detail"
    )

# ===========================
# 10. Detailed Table
# ===========================
st.markdown("---")
st.subheader("📋 Detailed Table")

# Columns required for calculations
detail_cols = ["ItemNumber","ItemName","Factory","Storagetype","Family","RawType","Cost","Month_0","Month_0_value","MINQTY","FixedSSQty","SSDays"]
rename_dict = {
    "ItemNumber":"SKU Code","ItemName":"Item Name","Factory":"Factory","Storagetype":"Storage Type",
    "Family":"Family","RawType":"Raw Material Type","Cost":"Cost (EGP)",
    "Month_0":"OH (Quantity)","Month_0_value":"OH (EGP)",
    "MINQTY":"Minimum Lot (MINQTY)","FixedSSQty":"Fixed SS Qty","SSDays":"SS Days"
}

if time_grouping=="Monthly":
    for i,m in enumerate(months_sequence):
        detail_cols += [f"Month_{i+1}",f"Month_{i+1}_value"]
        rename_dict[f"Month_{i+1}"] = f"{m['code']} (Quantity)"
        rename_dict[f"Month_{i+1}_value"] = f"{m['code']} (EGP)"
elif time_grouping=="Quarterly":
    # Use quarter closing values (end of each quarter)
    detail_cols += ["Month_3","Month_3_value","Month_6","Month_6_value","Month_9","Month_9_value","Month_12","Month_12_value"]
    rename_dict.update({
        "Month_3":"Q1 Close (Quantity)","Month_3_value":"Q1 Close (EGP)",
        "Month_6":"Q2 Close (Quantity)","Month_6_value":"Q2 Close (EGP)",
        "Month_9":"Q3 Close (Quantity)","Month_9_value":"Q3 Close (EGP)",
        "Month_12":"Q4 Close (Quantity)","Month_12_value":"Q4 Close (EGP)"
    })
else:  # Yearly
    # Use year-end closing values
    detail_cols += ["Month_12","Month_12_value"]
    rename_dict.update({
        "Month_12":"Year Close (Quantity)","Month_12_value":"Year Close (EGP)"
    })

df_display = df_with_months[detail_cols].rename(columns=rename_dict)

st.dataframe(
    # Format numeric columns with thousands separators
    df_display.style.format({
        col:"{:,.0f}" if "Quantity" in col or "OH (Quantity)" in col else "{:,.2f}"
        for col in df_display.columns if col not in ["SKU Code","Item Name","Factory","Storage Type","Family","Raw Material Type"]
    }).format({
        "SKU Code": lambda x: str(x).zfill(6) if pd.notna(x) and str(x).isdigit() else str(x)
    }),
    use_container_width=True,
    height=400
)

render_excel_download_button(
    df_display,
    "📥 Download Detailed Table (Excel)",
    "inventory_simulation",
    key="download_detailed_table_excel",
    sheet_name="Inventory"
)


# ===========================
# 11. Export Options
# ===========================
st.markdown("---")
col_export1,col_export2,col_export3 = st.columns([2,2,1])

with col_export1:
    render_excel_download_button(
        df_display,
        "📥 Download Excel",
        "inventory_simulation",
        key="download_detailed_table_excel_sidebar",
        sheet_name="Inventory"
    )

    st.download_button("📥 Download CSV", data=df_display.to_csv(index=False,encoding="utf-8-sig"),
                       file_name=f"inventory_simulation_{today.strftime('%Y%m%d')}.csv", mime="text/csv")

with col_export2:
    st.info(f"📊 Total records: {len(df_display):,}")

with col_export3:
    st.success(f"✅ {today.strftime('%b %Y')}")