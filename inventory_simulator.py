import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from pathlib import Path
from datetime import datetime
from dateutil.relativedelta import relativedelta
import io
import re

FG_DEFAULT_FG_PATH = Path(
    r"C:\Users\ashalaby\OneDrive - Halwani Bros\Planning - Sources\new view 2023\FP module 23.xlsb"
)
BOM_DEFAULT_PATH = Path(
    r"C:\Users\ashalaby\OneDrive - Halwani Bros\Planning - Sources\RECIPE(ver. Dec 2018).XLSX"
)
FG_PRIMARY_SHEETS = ("Data", "Data2")
FG_TEXT_COLUMNS = {
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
FG_METRIC_ORDER = [
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
FG_METRIC_LABELS = {
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
FG_MONTHLY_METRIC_CODES = ["BP_CUR", "BP_PRV", "AS"]
FG_MONTHLY_METRIC_LABELS = {
    "BP_CUR": "Current BP",
    "BP_PRV": "Previous BP",
    "AS": "Actual Sales",
}
FG_WEEKLY_METRICS = ["CurST", "CurAS", "CurAPP", "CurAP"]
FG_MONTH_ABBR = {
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
FG_WEEKLY_METRIC_LABELS = {
    "CurST": "Sales Target",
    "CurAS": "Actual Sales",
    "CurAPP": "Production Plan",
    "CurAP": "Actual Production",
}
FG_METRIC_COLOR_MAP = {
    "CurOS": "#f7dc6f",
    "CurAPP": "#8e5b32",
    "CurAP": "#ff7f0e",
    "CurOST": "#5dade2",
    "CurST": "#1f77b4",
    "CurAS": "#2ecc71",
    "CurFOC": "#af7ac5",
    "Oh": "#f1c40f",
    "NextOS": "#f9e79f",
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
FG_METRIC_CANDIDATES = {
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
FG_WEEK_RANGE = range(1, 6)
FG_TARGET_YEAR_SUFFIX = "26"
FG_MONTH_SEQUENCE = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
FG_MONTH_LABELS = [f"{month}{FG_TARGET_YEAR_SUFFIX}" for month in FG_MONTH_SEQUENCE]
FG_BP_COLUMN_PATTERN = re.compile(r"^(?P<prefix>.*?)(?P<month>[A-Za-z]{3})(?P<year>\d{2})BP$", re.IGNORECASE)
FG_AS_COLUMN_PATTERN = re.compile(r"^(?P<prefix>.*?)(?P<month>[A-Za-z]{3})(?P<year>\d{2})AS$", re.IGNORECASE)
FG_DC_LOCATION_MAP = {
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


@st.cache_data(show_spinner=False)
def load_fg_dataset(file_bytes: bytes | None) -> pd.DataFrame:
    source = io.BytesIO(file_bytes) if file_bytes else FG_DEFAULT_FG_PATH
    try:
        sheets = pd.read_excel(source, sheet_name=list(FG_PRIMARY_SHEETS), engine="pyxlsb")
    except FileNotFoundError:
        raise
    except ValueError:
        source = io.BytesIO(file_bytes) if file_bytes else FG_DEFAULT_FG_PATH
        sheets = pd.read_excel(source, sheet_name=[FG_PRIMARY_SHEETS[0]], engine="pyxlsb")

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
        normalized = str(column).strip().lower().replace(" ", "").replace("-", "").replace("_", "")
        if normalized in FG_TEXT_COLUMNS:
            merged_df[column] = merged_df[column].astype(str).str.strip()
        else:
            merged_df[column] = pd.to_numeric(merged_df[column], errors="ignore")

    return merged_df


def render_accent_subheader(text: str) -> None:
    st.markdown(
        f"""
        <h3 style='
            font-size: 1.25rem;
            color: #5e4ae3;
            letter-spacing: 0.05em;
            text-transform: uppercase;
            font-family: "Segoe UI", "Helvetica Neue", Arial, sans-serif;
            font-weight: 600;
            margin-top: 0.5rem;
            margin-bottom: 0.5rem;
        '>{text}</h3>
        """,
        unsafe_allow_html=True,
    )


def _ensure_page_banner_styles() -> None:
    if st.session_state.get("_hero_banner_css_injected"):
        return
    st.session_state["_hero_banner_css_injected"] = True
    st.markdown(
        """
        <style>
            @keyframes heroGradientShift {
                0% { transform: rotate(0deg) scale(1); }
                50% { transform: rotate(1.5deg) scale(1.05); }
                100% { transform: rotate(0deg) scale(1); }
            }

            @keyframes heroIconFloat {
                0%, 100% { transform: translateY(0); }
                50% { transform: translateY(-6px); }
            }

            .hero-banner {
                position: relative;
                overflow: hidden;
                border-radius: 20px;
                padding: 2.5rem clamp(1.5rem, 4vw, 3rem);
                margin-bottom: 1.75rem;
                color: #ffffff;
                background: var(
                    --hero-gradient,
                    linear-gradient(120deg, #8e2de2 0%, #4a00e0 50%, #00c6ff 100%)
                );
                box-shadow: 0 18px 45px rgba(20, 20, 43, 0.25);
                border: 1px solid rgba(255, 255, 255, 0.25);
            }

            .hero-banner::before {
                content: "";
                position: absolute;
                inset: -60% -30% -40% -30%;
                background: radial-gradient(circle at center, rgba(255, 255, 255, 0.25), transparent 70%);
                animation: heroGradientShift 16s ease-in-out infinite;
                opacity: 0.7;
            }

            .hero-banner__content {
                position: relative;
                z-index: 1;
                display: flex;
                flex-direction: column;
                gap: 0.75rem;
            }

            .hero-banner__title {
                display: flex;
                align-items: center;
                gap: 0.75rem;
                font-size: clamp(1.6rem, 2.4vw, 2.4rem);
                font-weight: 700;
                letter-spacing: 0.02em;
            }

            .hero-banner__icon {
                font-size: clamp(1.9rem, 3vw, 2.6rem);
                animation: heroIconFloat 4.5s ease-in-out infinite;
                filter: drop-shadow(0 6px 8px rgba(0, 0, 0, 0.25));
            }

            .hero-banner__subtitle {
                font-size: clamp(0.95rem, 1.5vw, 1.1rem);
                font-weight: 500;
                opacity: 0.85;
            }

            .hero-banner__meta {
                display: inline-flex;
                gap: 0.65rem;
                align-items: center;
                padding: 0.55rem 1.2rem;
                border-radius: 999px;
                background: rgba(255, 255, 255, 0.18);
                border: 1px solid rgba(255, 255, 255, 0.3);
                backdrop-filter: blur(8px);
                font-size: 0.9rem;
            }

            @media (max-width: 640px) {
                .hero-banner {
                    padding: 2rem 1.5rem;
                }
            }

            [data-testid="stSidebar"] > div:first-child {
                background: linear-gradient(180deg, #f7f7f7 0%, #e2e2e2 100%);
                color: #1f1f1f;
                padding: 1.5rem 1rem 2rem;
                border-right: 1px solid rgba(0, 0, 0, 0.08);
                box-shadow: inset -1px 0 0 rgba(255, 255, 255, 0.3);
            }

            [data-testid="stSidebar"] > div:first-child * {
                color: inherit !important;
            }

            [data-testid="stSidebar"] [data-testid="stMarkdown"] a {
                color: #3949ab !important;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_hero_banner(
    title: str,
    subtitle: str | None = None,
    *,
    icon: str | None = None,
    meta: str | None = None,
    gradient: str | None = None,
) -> None:
    _ensure_page_banner_styles()
    gradient_style = f" style=\"--hero-gradient: {gradient}\"" if gradient else ""
    icon_html = f"<span class='hero-banner__icon'>{icon}</span>" if icon else ""
    subtitle_html = (
        f"<p class='hero-banner__subtitle'>{subtitle}</p>" if subtitle else ""
    )
    meta_html = (
        f"<span class='hero-banner__meta'>{meta}</span>" if meta else ""
    )
    st.markdown(
        f"""
        <div class='hero-banner'{gradient_style}>
            <div class='hero-banner__content'>
                <div class='hero-banner__title'>
                    {icon_html}
                    <span>{title}</span>
                </div>
                {subtitle_html}
                {meta_html}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _canonical_item_code(value: str | int | float | None) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    if text.isdigit():
        stripped = text.lstrip("0")
        return stripped or "0"
    return text


def _get_factory_icon(factory_name: str) -> str:
    """Return an icon for a factory based on its name."""
    if not factory_name or factory_name == "All":
        return "🏭"
    factory_lower = str(factory_name).strip().lower()
    if "wip" in factory_lower:
        return "⚙️"
    elif "meat" in factory_lower or "لحم" in factory_lower:
        return "🥩"
    elif "chicken" in factory_lower or "دجاج" in factory_lower or "poultry" in factory_lower:
        return "🍗"
    elif "halawa" in factory_lower or "حلاوة" in factory_lower or "halva" in factory_lower:
        return "🌰"
    elif "tahina" in factory_lower or "طحينة" in factory_lower or "tahini" in factory_lower or "tehena" in factory_lower:
        return "🌰"
    elif "jam" in factory_lower or "مربى" in factory_lower:
        return "🍓"
    elif "juice" in factory_lower or "عصير" in factory_lower:
        return "🧃"
    elif "tost" in factory_lower or "toast" in factory_lower or "توست" in factory_lower or "bread" in factory_lower or "خبز" in factory_lower:
        return "🍞"
    elif "fruit" in factory_lower or "فواكه" in factory_lower:
        return "🍎"
    elif "jeddah" in factory_lower or "جدة" in factory_lower:
        return "🏭"
    elif "riyadh" in factory_lower or "رياض" in factory_lower:
        return "🏢"
    elif "dammam" in factory_lower or "دمام" in factory_lower:
        return "🏗️"
    elif "cairo" in factory_lower or "قاهرة" in factory_lower:
        return "🏛️"
    elif "alex" in factory_lower or "اسكندرية" in factory_lower:
        return "⚓"
    elif "unspecified" in factory_lower or "unknown" in factory_lower:
        return "❓"
    else:
        return "🏭"


def _get_family_icon(family_name: str) -> str:
    """Return an icon for a product family based on its name."""
    if not family_name or family_name == "All":
        return "📦"
    family_lower = str(family_name).strip().lower()
    # Meat products
    if "beef" in family_lower or "لحم بقري" in family_lower or "burger" in family_lower or "برجر" in family_lower:
        return "🥩"
    elif "sausage" in family_lower or "سجق" in family_lower or "hotdog" in family_lower:
        return "🌭"
    elif "mortadella" in family_lower or "مرتديلا" in family_lower:
        return "🥓"
    # Chicken products
    elif "chicken" in family_lower or "دجاج" in family_lower or "nugget" in family_lower or "ناجتس" in family_lower:
        return "🍗"
    # Sesame products
    elif "halawa" in family_lower or "حلاوة" in family_lower or "halva" in family_lower:
        return "🌰"
    elif "tahina" in family_lower or "طحينة" in family_lower or "tahini" in family_lower:
        return "🌰"
    # Spreads and preserves
    elif "jam" in family_lower or "مربى" in family_lower or "marmalade" in family_lower:
        return "🍓"
    elif "honey" in family_lower or "عسل" in family_lower:
        return "🍯"
    elif "chocolate" in family_lower or "شوكولاتة" in family_lower or "nutella" in family_lower:
        return "🍫"
    # Beverages
    elif "juice" in family_lower or "عصير" in family_lower:
        return "🧃"
    elif "drink" in family_lower or "مشروب" in family_lower or "beverage" in family_lower:
        return "🥤"
    # Bakery
    elif "bread" in family_lower or "خبز" in family_lower or "toast" in family_lower or "توست" in family_lower:
        return "🍞"
    elif "cake" in family_lower or "كيك" in family_lower or "pastry" in family_lower:
        return "🎂"
    # Dairy
    elif "cheese" in family_lower or "جبن" in family_lower:
        return "🧀"
    elif "butter" in family_lower or "زبدة" in family_lower:
        return "🧈"
    # Vegetables and fruits
    elif "vegetable" in family_lower or "خضار" in family_lower:
        return "🥬"
    elif "fruit" in family_lower or "فواكه" in family_lower:
        return "🍎"
    # Snacks
    elif "snack" in family_lower or "سناك" in family_lower or "chip" in family_lower:
        return "🍿"
    # Packaging
    elif "packaging" in family_lower or "تغليف" in family_lower or "package" in family_lower:
        return "📦"
    else:
        return "📦"


def _get_rawtype_icon(raw_type: str) -> str:
    """Return an icon for raw material type based on its name."""
    if not raw_type or raw_type == "All":
        return "🧱"
    raw_lower = str(raw_type).strip().lower()
    # Meat and poultry
    if "meat" in raw_lower or "لحم" in raw_lower or "beef" in raw_lower:
        return "🥩"
    elif "chicken" in raw_lower or "دجاج" in raw_lower or "poultry" in raw_lower:
        return "🍗"
    # Grains and seeds
    elif "sesame" in raw_lower or "سمسم" in raw_lower:
        return "🌰"
    elif "wheat" in raw_lower or "قمح" in raw_lower or "flour" in raw_lower or "دقيق" in raw_lower:
        return "🌾"
    elif "grain" in raw_lower or "حبوب" in raw_lower or "cereal" in raw_lower:
        return "🌾"
    # Fruits and vegetables
    elif "fruit" in raw_lower or "فواكه" in raw_lower or "strawberry" in raw_lower or "orange" in raw_lower:
        return "🍎"
    elif "vegetable" in raw_lower or "خضار" in raw_lower or "tomato" in raw_lower:
        return "🥬"
    # Dairy
    elif "milk" in raw_lower or "حليب" in raw_lower or "dairy" in raw_lower:
        return "🥛"
    elif "cheese" in raw_lower or "جبن" in raw_lower:
        return "🧀"
    elif "butter" in raw_lower or "زبدة" in raw_lower or "cream" in raw_lower:
        return "🧈"
    # Sweeteners
    elif "sugar" in raw_lower or "سكر" in raw_lower:
        return "🍬"
    elif "honey" in raw_lower or "عسل" in raw_lower:
        return "🍯"
    # Oils and fats
    elif "oil" in raw_lower or "زيت" in raw_lower or "fat" in raw_lower:
        return "🫠"
    # Spices and seasonings
    elif "spice" in raw_lower or "توابل" in raw_lower or "salt" in raw_lower or "ملح" in raw_lower:
        return "🧂"
    # Packaging materials
    elif "packaging" in raw_lower or "تغليف" in raw_lower or "carton" in raw_lower or "plastic" in raw_lower or "can" in raw_lower:
        return "📦"
    elif "bottle" in raw_lower or "زجاجة" in raw_lower or "jar" in raw_lower:
        return "🍾"
    # Chemicals and additives
    elif "chemical" in raw_lower or "كيميائي" in raw_lower or "additive" in raw_lower:
        return "⚗️"
    else:
        return "🧱"


@st.cache_data(show_spinner=False)
def load_bom_workbook(file_bytes: bytes | None = None, *, allow_default: bool = True) -> tuple[pd.DataFrame, pd.DataFrame]:
    if file_bytes:
        source = io.BytesIO(file_bytes)
    elif allow_default:
        if not BOM_DEFAULT_PATH.exists():
            raise FileNotFoundError(
                "Default BOM workbook not found. Please upload a BOM workbook."
            )
        source = BOM_DEFAULT_PATH
    else:
        raise FileNotFoundError("BOM workbook not provided.")

    sheets = pd.read_excel(source, sheet_name=["Bom Data", "Recipe Data"], dtype=str)
    bom_df = sheets.get("Bom Data", pd.DataFrame())
    recipe_df = sheets.get("Recipe Data", pd.DataFrame())

    for frame in (bom_df, recipe_df):
        if frame is None or frame.empty:
            continue
        frame.columns = [str(col).strip() for col in frame.columns]
        for enforce_col in ("BomVersion_ItemId", "ItemId"):
            if enforce_col in frame.columns:
                frame[enforce_col] = frame[enforce_col].astype(str).str.strip()

    return bom_df, recipe_df


def render_fg_explorer(materials_df: pd.DataFrame | None = None) -> None:
    render_hero_banner(
        title="FG Explorer",
        subtitle="Analyze finished goods by cascading filters and KPIs.",
        icon="🏷️",
        meta="Turn The Planning Up ⬆️💡⬆️",
        gradient="linear-gradient(120deg, #ff7eb3 0%, #ff758c 50%, #ffae70 100%)",
    )

    def _normalize_name(name: str | int) -> str:
        return str(name).strip().lower().replace(" ", "").replace("-", "").replace("_", "")

    def _find_column(columns: pd.Index, candidates: list[str]) -> str | None:
        lookup = {_normalize_name(col): col for col in columns}
        for candidate in candidates:
            actual = lookup.get(_normalize_name(candidate))
            if actual is not None:
                return actual
        return None

    def _unique_sorted_values(df: pd.DataFrame, column: str | None) -> list[str]:
        if column is None or column not in df.columns:
            return []
        series = df[column].dropna().astype(str).str.strip()
        return sorted({value for value in series if value and value.lower() != "all"})

    def _resolve_prefix_label(prefix: str) -> tuple[str, int]:
        prefix_lower = prefix.strip().lower()
        if prefix_lower.startswith("prv"):
            suffix = prefix[3:]
            return (f"PRV{suffix}" if suffix else "PRV", -int(suffix) if suffix.isdigit() else -1)
        if prefix_lower.startswith("cur"):
            suffix = prefix[3:]
            return (f"CUR{suffix.upper()}" if suffix else "CUR", 0)
        if prefix_lower.startswith("next"):
            suffix = prefix[4:]
            return (f"NEXT{suffix}" if suffix else "NEXT", int(suffix) if suffix.isdigit() else 1)
        match = re.match(r"(?P<month>[A-Za-z]{3})(?P<year>\d{2})(?P<suffix>[A-Za-z]*)", prefix.strip())
        if match:
            month = match.group("month").lower()
            year = int(match.group("year")) + 2000
            month_num = FG_MONTH_ABBR.get(month)
            order = 100
            if month_num:
                base = pd.Timestamp.today().normalize().replace(day=1)
                target = pd.Timestamp(year=year, month=month_num, day=1)
                order = (target.year - base.year) * 12 + (target.month - base.month)
            suffix = match.group("suffix").upper()
            label = f"{match.group('month').title()}{match.group('year')}{suffix}" if suffix else f"{match.group('month').title()}{match.group('year')}"
            return label, order
        return prefix.upper(), 100

    def _build_weekly_series(row: pd.Series) -> tuple[list[str], dict[str, list[float]]]:
        labels: list[str] = []
        series_map: dict[str, list[float]] = {metric: [] for metric in FG_WEEKLY_METRICS}
        for week in FG_WEEK_RANGE:
            week_label = f"W{week}"
            has_data = False
            cached: dict[str, float] = {}
            for metric in FG_WEEKLY_METRICS:
                column_name = _find_column(row.index, [f"{metric}W{week}"])
                value = row[column_name] if column_name else np.nan
                if pd.notna(value):
                    has_data = True
                cached[metric] = value
            if has_data:
                labels.append(week_label)
                for metric in FG_WEEKLY_METRICS:
                    series_map[metric].append(cached[metric])
        return labels, series_map

    def _build_monthly_series(row: pd.Series) -> tuple[list[str], dict[str, list[float]]]:
        month_data = {label: {metric: np.nan for metric in FG_MONTHLY_METRIC_CODES} for label in FG_MONTH_LABELS}
        for column in row.index:
            col_name = str(column).strip()
            bp_match = FG_BP_COLUMN_PATTERN.match(col_name)
            if bp_match:
                month_label = f"{bp_match.group('month').title()}{bp_match.group('year')}"
                if month_label not in FG_MONTH_LABELS:
                    continue
                metric_key = "BP_PRV" if (bp_match.group("prefix") or "").strip().lower().startswith("prv") else "BP_CUR"
                numeric_value = pd.to_numeric(row[column], errors="coerce")
                if pd.isna(numeric_value):
                    continue
                entry = month_data[month_label]
                entry[metric_key] = numeric_value if pd.isna(entry[metric_key]) else entry[metric_key] + numeric_value
                continue
            as_match = FG_AS_COLUMN_PATTERN.match(col_name)
            if as_match:
                month_label = f"{as_match.group('month').title()}{as_match.group('year')}"
                if month_label not in FG_MONTH_LABELS:
                    continue
                numeric_value = pd.to_numeric(row[column], errors="coerce")
                if pd.isna(numeric_value):
                    continue
                entry = month_data[month_label]
                entry["AS"] = numeric_value if pd.isna(entry["AS"]) else entry["AS"] + numeric_value
        labels: list[str] = []
        series_map = {metric: [] for metric in FG_MONTHLY_METRIC_CODES}
        for month_label in FG_MONTH_LABELS:
            values = month_data.get(month_label)
            if not values:
                continue
            if all(pd.isna(values[metric]) for metric in FG_MONTHLY_METRIC_CODES):
                continue
            labels.append(month_label)
            for metric in FG_MONTHLY_METRIC_CODES:
                series_map[metric].append(values[metric])
        return labels, series_map

    def _filter_monthly_series(series_map: dict[str, list[float]]) -> tuple[dict[str, list[float]], list[str]]:
        filtered: dict[str, list[float]] = {}
        order: list[str] = []
        for metric in FG_MONTHLY_METRIC_CODES:
            values = series_map.get(metric, [])
            if values and any(pd.notna(v) for v in values):
                filtered[metric] = values
                order.append(metric)
        return filtered, order

    def _calculate_dioh(df: pd.DataFrame, metric_column_map: dict[str, str | None]) -> float | None:
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
        average_daily_demand = total_demand / days_in_month if days_in_month else None
        if not average_daily_demand or average_daily_demand <= 0:
            return None
        return total_oh / average_daily_demand

    def _build_group_series(
        df: pd.DataFrame,
        group_col: str | None,
        metric_column_map: dict[str, str | None],
    ) -> tuple[list[str], dict[str, list[float]]]:
        if not group_col or group_col not in df.columns:
            return [], {metric: [] for metric in FG_METRIC_ORDER}

        labels: list[str] = []
        series_map: dict[str, list[float]] = {metric: [] for metric in FG_METRIC_ORDER}
        grouped = df.groupby(group_col, dropna=False)
        for group_value, group_df in grouped:
            label = str(group_value) if pd.notna(group_value) and str(group_value).strip() else "Unspecified"
            labels.append(label)
            for metric in FG_METRIC_ORDER:
                column_name = metric_column_map.get(metric)
                if column_name and column_name in group_df.columns:
                    total_value = pd.to_numeric(group_df[column_name], errors="coerce").fillna(0).sum()
                else:
                    total_value = 0.0
                series_map[metric].append(total_value)
        return labels, series_map

    def _compute_kpis(
        df: pd.DataFrame,
        metric_column_map: dict[str, str | None],
        item_col: str | None,
        name_col: str | None,
        factory_col: str | None,
        family_col: str | None,
        market_col: str | None,
    ) -> dict[str, tuple[int, pd.DataFrame]]:
        base_index = df.index
        empty_bool = pd.Series(False, index=base_index)

        planning_col = _find_column(df.columns, ["Planning", "PlanningType", "PlanType"])
        if planning_col and planning_col in df.columns:
            planning_series = df[planning_col].astype(str).str.upper().str.strip()
        else:
            planning_series = pd.Series("", index=base_index)

        curst_col = metric_column_map.get("CurST")
        if curst_col and curst_col in df.columns:
            numeric_curst = pd.to_numeric(df[curst_col], errors="coerce").fillna(0)
        else:
            numeric_curst = pd.Series(0, index=base_index, dtype=float)

        oh_col = metric_column_map.get("Oh")
        if oh_col and oh_col in df.columns:
            numeric_oh = pd.to_numeric(df[oh_col], errors="coerce").fillna(0)
        else:
            numeric_oh = pd.Series(0, index=base_index, dtype=float)

        ssqty_col = _find_column(df.columns, ["SSQty", "SafetyStockQty", "SS Qty"])
        if ssqty_col and ssqty_col in df.columns:
            numeric_ssqty = pd.to_numeric(df[ssqty_col], errors="coerce").fillna(0)
        else:
            numeric_ssqty = pd.Series(0, index=base_index, dtype=float)

        detail_columns: list[str] = []
        for candidate in [item_col, name_col, planning_col, factory_col, family_col, market_col, curst_col, ssqty_col, oh_col]:
            if candidate and candidate in df.columns and candidate not in detail_columns:
                detail_columns.append(candidate)

        def _detail(mask: pd.Series | None) -> tuple[int, pd.DataFrame]:
            if mask is None or mask.empty:
                return 0, pd.DataFrame(columns=detail_columns)
            prepared_mask = mask.reindex(base_index, fill_value=False)
            count = int(prepared_mask.sum())
            if not detail_columns:
                return count, df.loc[prepared_mask].copy()
            return count, df.loc[prepared_mask, detail_columns].copy()

        mts_mask = planning_series.eq("MTS") if not planning_series.empty else empty_bool
        mto_mask = planning_series.eq("MTO") if not planning_series.empty else empty_bool

        oh_lt_ss_mask = empty_bool
        if ssqty_col:
            oh_lt_ss_mask = (numeric_oh < numeric_ssqty) & (numeric_ssqty > 0)

        curst_gt0_oh0_mask = (numeric_curst > 0) & (numeric_oh == 0)

        return {
            "MTS": _detail(mts_mask.fillna(False)),
            "MTO": _detail(mto_mask.fillna(False)),
            "OH_lt_SSQty": _detail(oh_lt_ss_mask.fillna(False)),
            "CurST_gt0_OH0": _detail(curst_gt0_oh0_mask.fillna(False)),
        }

    def _render_stacked_bar(labels: list[str], series_map: dict[str, list[float]], title: str, order: list[str]) -> None:
        fig = go.Figure()
        for metric in order:
            values = series_map.get(metric, [])
            if not values or all(pd.isna(values)):
                continue
            fig.add_trace(
                go.Bar(
                    x=labels,
                    y=values,
                    name=FG_METRIC_LABELS.get(metric, metric),
                    marker={"color": FG_METRIC_COLOR_MAP.get(metric)},
                )
            )
        if not fig.data:
            st.info("No data available for the selected view.")
            return
        fig.update_layout(title=title, barmode="stack", height=420, hovermode="x unified")
        st.plotly_chart(fig, use_container_width=True)

    def _render_line_chart(labels: list[str], series_map: dict[str, list[float]], title: str, order: list[str]) -> None:
        fig = go.Figure()
        for metric in order:
            values = series_map.get(metric, [])
            if not values or all(pd.isna(values)):
                continue
            color = FG_METRIC_COLOR_MAP.get(metric)
            fig.add_trace(
                go.Scatter(
                    x=labels,
                    y=values,
                    name=FG_MONTHLY_METRIC_LABELS.get(metric, metric),
                    mode="lines+markers",
                    line={"color": color} if color else None,
                )
            )
        if not fig.data:
            st.info("No data available for the selected view.")
            return
        fig.update_layout(title=title, height=420, hovermode="x unified")
        st.plotly_chart(fig, use_container_width=True)

    # Use FG file from sidebar
    # fg_bytes is already defined in sidebar
    try:
        with st.spinner("Loading FG data..."):
            fg_df = load_fg_dataset(fg_bytes)
    except FileNotFoundError:
        st.error("Default FG workbook not found. Upload another file.")
        return
    except Exception as exc:
        st.error(f"Unable to read FG workbook: {exc}")
        return

    if fg_df.empty:
        st.warning("FG workbook is empty or missing expected sheets.")
        return

    factory_col = _find_column(fg_df.columns, ["Factory"])
    item_col = _find_column(fg_df.columns, ["ItemNumber"])
    name_col = _find_column(fg_df.columns, ["ItemName"])
    family_col = _find_column(fg_df.columns, ["Family", "SubFamily", "SubFamilyName"])
    oh_col = _find_column(fg_df.columns, ["CurSOH", "CurOH", "SOH", "CurStockOnHand", "StockOnHand", "OH"])
    market_col = _find_column(fg_df.columns, ["Market", "LocExp", "LOCEXP", "Channel"])
    if market_col is None:
        market_col = "_Market"
        fg_df[market_col] = "All"

    metric_column_map = {
        metric: _find_column(fg_df.columns, candidates)
        for metric, candidates in FG_METRIC_CANDIDATES.items()
    }

    numeric_cols = fg_df.select_dtypes(include=[np.number]).columns
    fg_df[numeric_cols] = fg_df[numeric_cols].fillna(0)

    if factory_col:
        fg_df = fg_df[
            fg_df[factory_col].astype(str).str.strip().ne("") & fg_df[factory_col].notna()
        ]
    if family_col:
        fg_df = fg_df[
            fg_df[family_col].astype(str).str.strip().ne("") & fg_df[family_col].notna()
        ]

    st.markdown("### Filters")
    filtered_stage = fg_df.copy()
    row_one = st.columns(3)
    row_two = st.columns(3)

    with row_one[0]:
        storage_col = _find_column(fg_df.columns, ["StorageType", "Storage", "Storage Type"])
        storage_options = ["All"] + _unique_sorted_values(filtered_stage, storage_col)
        storage_choice = st.selectbox("Storage type", storage_options, index=0)
        if storage_col and storage_choice != "All":
            filtered_stage = filtered_stage[filtered_stage[storage_col].astype(str).str.strip().eq(storage_choice)]

    with row_one[1]:
        raw_col = _find_column(fg_df.columns, ["RawType", "Raw"])
        raw_options = ["All"] + _unique_sorted_values(filtered_stage, raw_col)
        raw_display_options = [f"{_get_rawtype_icon(r)} {r}" for r in raw_options]
        raw_choice_display = st.selectbox("Raw type", raw_display_options, index=0)
        raw_choice = raw_options[raw_display_options.index(raw_choice_display)]
        if raw_col and raw_choice != "All":
            filtered_stage = filtered_stage[filtered_stage[raw_col].astype(str).str.strip().eq(raw_choice)]

    with row_one[2]:
        market_options = ["All"] + _unique_sorted_values(filtered_stage, market_col)
        market_choice = st.selectbox("Market", market_options, index=0)
        if market_choice != "All":
            filtered_stage = filtered_stage[filtered_stage[market_col].astype(str).str.strip().eq(market_choice)]

    with row_two[0]:
        # Add CSS for larger icons in selectbox
        st.markdown("""
            <style>
            div[data-baseweb="select"] span {
                font-size: 1.2em !important;
            }
            </style>
        """, unsafe_allow_html=True)
        factory_options = ["All"] + _unique_sorted_values(filtered_stage, factory_col)
        factory_display_options = [f"{_get_factory_icon(f)} {f}" for f in factory_options]
        factory_choice_display = st.selectbox("Factory", factory_display_options, index=0)
        factory_choice = factory_options[factory_display_options.index(factory_choice_display)]
        if factory_choice != "All":
            filtered_stage = filtered_stage[filtered_stage[factory_col].astype(str).str.strip().eq(factory_choice)]

    with row_two[1]:
        family_options = ["All"] + _unique_sorted_values(filtered_stage, family_col)
        family_display_options = [f"{_get_family_icon(f)} {f}" for f in family_options]
        family_choice_display = st.selectbox("Family", family_display_options, index=0)
        family_choice = family_options[family_display_options.index(family_choice_display)]
        if family_choice != "All" and family_col:
            filtered_stage = filtered_stage[filtered_stage[family_col].astype(str).str.strip().eq(family_choice)]

    with row_two[2]:
        item_options = ["All"]
        item_display_to_value: dict[str, str] = {}
        if item_col and item_col in filtered_stage.columns:
            if name_col and name_col in filtered_stage.columns:
                item_subset = (
                    filtered_stage[[item_col, name_col]]
                    .dropna()
                    .drop_duplicates()
                    .assign(display=lambda df: df[item_col].astype(str).str.zfill(6) + " - " + df[name_col].astype(str))
                )
                displays = sorted(item_subset["display"].tolist())
                item_options += displays
                item_display_to_value = dict(zip(item_subset["display"], item_subset[item_col].astype(str)))
            else:
                item_values = _unique_sorted_values(filtered_stage, item_col)
                item_options += item_values
                item_display_to_value = {value: value for value in item_values}
        item_choice = st.selectbox("Items", item_options, index=0)
        if item_choice != "All" and item_display_to_value:
            target_item = item_display_to_value.get(item_choice)
            if target_item is not None:
                filtered_stage = filtered_stage[filtered_stage[item_col].astype(str).str.strip().eq(target_item)]

    filtered_df = filtered_stage.copy()
    if filtered_df.empty:
        st.warning("No data matches the current filters.")
        return

    kpi_data = _compute_kpis(
        filtered_df,
        metric_column_map,
        item_col,
        name_col,
        factory_col,
        family_col,
        market_col,
    )
    dioh_value = _calculate_dioh(filtered_df, metric_column_map)
    st.caption("Click a KPI card's button to toggle its detailed list.")
    kpi_cols = st.columns(5)
    with kpi_cols[0]:
        st.metric("DIOH", f"{dioh_value:,.1f} days" if dioh_value is not None else "—")

    kpi_labels = [
        ("📦 MTS SKUs", "MTS"),
        ("🔧 MTO SKUs", "MTO"),
        ("⚠️ < SSQty", "OH_lt_SSQty"),
        ("🚨 OOS (CurST>0, OH=0)", "CurST_gt0_OH0"),
    ]
    for col, (label, key) in zip(kpi_cols[1:], kpi_labels):
        with col:
            count_value, _ = kpi_data.get(key, (0, pd.DataFrame()))
            st.metric(label, f"{count_value:,}")
            if st.button("View details", key=f"fg_kpi_btn_{key}"):
                st.session_state[f"fg_show_{key}"] = not st.session_state.get(f"fg_show_{key}", False)

    # Collect all KPI detail dataframes for shortage analysis
    all_kpi_fgs: list[pd.DataFrame] = []
    for label, key in kpi_labels:
        _, detail_df = kpi_data.get(key, (0, pd.DataFrame()))
        if not detail_df.empty:
            all_kpi_fgs.append(detail_df)
        
        if st.session_state.get(f"fg_show_{key}"):
            st.markdown(f"#### {label} details")
            if detail_df.empty:
                st.info("No matching items for this KPI.")
            else:
                st.dataframe(detail_df, use_container_width=True)

    aggregated_row = filtered_df.sum(numeric_only=True)
    weekly_labels, weekly_series = _build_weekly_series(aggregated_row)
    _render_stacked_bar(weekly_labels, weekly_series, "Cur month weekly breakdown", FG_WEEKLY_METRICS)

    monthly_labels, monthly_series = _build_monthly_series(aggregated_row)
    monthly_filtered_series, monthly_order = _filter_monthly_series(monthly_series)
    _render_line_chart(monthly_labels, monthly_filtered_series, "Monthly trajectory", monthly_order)

    dc_columns = [
        col
        for col in filtered_df.columns
        if _normalize_name(col).startswith("dc") or _normalize_name(col).endswith("dc")
    ]
    if dc_columns:
        dc_totals = filtered_df[dc_columns].apply(pd.to_numeric, errors="coerce").sum()
        dc_totals = dc_totals[dc_totals > 0]
        if not dc_totals.empty:
            st.markdown("## 🗺️ Stock by DC")
            records: list[dict[str, float | str]] = []
            for column, value in dc_totals.items():
                display_name = str(column).strip()
                normalized = _normalize_name(display_name)
                coords = None
                for key, loc in FG_DC_LOCATION_MAP.items():
                    key_norm = _normalize_name(key)
                    if key_norm == normalized or key_norm in normalized or key in display_name:
                        coords = loc
                        break
                if coords is None:
                    continue
                records.append({
                    "Location": display_name,
                    "Stock": float(value),
                    "lat": coords["lat"],
                    "lon": coords["lon"],
                })
            if records:
                dc_df = pd.DataFrame(records)
                options = ["All"] + sorted(dc_df["Location"].unique())
                choice = st.selectbox("Select DC location", options=options, index=0, key="fg_dc_location")
                display_df = dc_df if choice == "All" else dc_df[dc_df["Location"].eq(choice)]
                if not display_df.empty:
                    fig_map = go.Figure(
                        go.Scattergeo(
                            lon=display_df["lon"],
                            lat=display_df["lat"],
                            text=[f"{row.Location}<br>Stock: {row.Stock:,.0f}" for row in display_df.itertuples()],
                            mode="markers",
                            marker=dict(
                                size=np.clip(display_df["Stock"] / display_df["Stock"].max() * 40, 6, 40),
                                color=display_df["Stock"],
                                colorscale="Oranges",
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
                        display_df[["Location", "Stock"]].sort_values("Stock", ascending=False).style.format({"Stock": "{:,.0f}"}),
                        use_container_width=True,
                    )
                    st.caption("Quick branch view:")
                    st.map(
                        display_df.rename(columns={"lat": "Latitude", "lon": "Longitude"}),
                        latitude="Latitude",
                        longitude="Longitude",
                        zoom=6,
                        size=100,
                    )

    # Combine all KPI FGs for shortage analysis
    combined_shortage_fgs = pd.DataFrame()
    if all_kpi_fgs:
        combined_shortage_fgs = pd.concat(all_kpi_fgs, ignore_index=True)
        if item_col in combined_shortage_fgs.columns:
            combined_shortage_fgs = combined_shortage_fgs.drop_duplicates(subset=[item_col])
    
    if not combined_shortage_fgs.empty:
        st.markdown("### ⚠️ Materials with Insufficient OH")

        bom_df = recipe_df = pd.DataFrame()
        try:
            bom_df, recipe_df = load_bom_workbook(None, allow_default=True)
        except FileNotFoundError:
            st.warning("Default BOM workbook not found. Upload the BOM in the BOM Calculator view to enable this insight.")
        except Exception as exc:  # noqa: BLE001
            st.warning(f"Unable to load BOM workbook: {exc}")

        if bom_df.empty:
            st.info("BOM data unavailable; cannot evaluate component shortages.")
        else:
            item_code_col = item_col if item_col in combined_shortage_fgs.columns else None
            item_name_col = name_col if name_col in combined_shortage_fgs.columns else None

            materials_stock: dict[str, float] = {}
            if materials_df is not None and not materials_df.empty:
                stock_code_col = _find_column(
                    materials_df.columns,
                    ["ItemNumber", "Item Number", "Item", "Item No", "ItemCode", "Item Code"],
                )
                stock_qty_col = _find_column(
                    materials_df.columns,
                    ["OH", "CurOH", "StockOnHand", "SOH", "CurrentOH"],
                )
                factory_col = _find_column(
                    materials_df.columns,
                    ["Factory", "Plant", "Location"],
                )
                if stock_code_col and stock_qty_col:
                    # Filter out WIP factory rows
                    filtered_materials = materials_df.copy()
                    if factory_col:
                        filtered_materials = filtered_materials[
                            ~filtered_materials[factory_col].astype(str).str.strip().str.upper().eq("WIP")
                        ]
                    
                    stock_series = filtered_materials[[stock_code_col, stock_qty_col]].copy()
                    stock_series[stock_code_col] = stock_series[stock_code_col].astype(str).str.strip()
                    stock_series[stock_qty_col] = pd.to_numeric(
                        stock_series[stock_qty_col], errors="coerce"
                    ).fillna(0.0)
                    materials_stock = {}
                    for _, entry in stock_series.iterrows():
                        raw_code = str(entry[stock_code_col]).strip()
                        qty = float(entry[stock_qty_col])
                        if not raw_code:
                            continue
                        canon_code = _canonical_item_code(raw_code)
                        for key in {raw_code, canon_code}:
                            if not key:
                                continue
                            existing = materials_stock.get(key)
                            materials_stock[key] = qty if existing is None else max(existing, qty)

            if not materials_stock:
                st.info("Materials stock (ItemNumber & OH) not available; please ensure the main materials workbook is loaded.")
            else:
                shortage_rows: list[dict[str, float | str | bool]] = []
                components_lookup: dict[str, pd.DataFrame] = {}
                skipped_fgs: list[str] = []

                def _resolve_demand(row: pd.Series) -> float:
                    demand_keys = [
                        "CurST",
                        next((c for c in row.index if str(c).lower().endswith("curst")), None),
                        "CurAPP",
                        next((c for c in row.index if str(c).lower() in {"demand", "forecast"}), None),
                    ]
                    for key in demand_keys:
                        if key and key in row.index:
                            val = pd.to_numeric(row.get(key), errors="coerce")
                            if pd.notna(val) and val > 0:
                                return float(val)
                    return 0.0

                for _, fg_row in combined_shortage_fgs.iterrows():
                    fg_code = str(fg_row.get(item_code_col, "")).strip() if item_code_col else ""
                    if not fg_code:
                        continue
                    fg_name = str(fg_row.get(item_name_col, "")).strip() if item_name_col else fg_code
                    demand_qty = _resolve_demand(fg_row)
                    shortage_info = compute_fg_material_shortage(
                        fg_code,
                        demand_qty,
                        bom_df,
                        recipe_df,
                        materials_stock,
                    )
                    if shortage_info is None:
                        continue

                    if not shortage_info.get("coverage_evaluated", True):
                        skipped_fgs.append(fg_code)
                        continue

                    max_feasible_qty = shortage_info["max_feasible_qty"]
                    shortfall_qty = shortage_info["shortfall_qty"]
                    if not np.isfinite(max_feasible_qty):
                        max_feasible_qty_display = "∞"
                    else:
                        max_feasible_qty_display = max(max_feasible_qty, 0.0)

                    shortage_rows.append(
                        {
                            "FG Code": fg_code,
                            "FG Name": fg_name,
                            "Required Qty": shortage_info["required_qty"],
                            "Max Feasible Qty": max_feasible_qty_display,
                            "Shortfall Qty": shortfall_qty,
                            "Shortage?": shortage_info["shortage_flag"],
                            "Limiting Components": shortage_info["limiting_summary"],
                        }
                    )
                    components_lookup[fg_code] = shortage_info["components_table"]

                if not shortage_rows:
                    st.info("✅ No material shortages detected for OOS finished goods.")
                else:
                    # Collect all shortage components across all FGs
                    all_shortage_comps = []
                    for fg_code in components_lookup:
                        comp_df = components_lookup[fg_code]
                        if "is_short" in comp_df.columns:
                            shortage_comps = comp_df[comp_df["is_short"] == True].copy()
                            if not shortage_comps.empty:
                                shortage_comps["FG_Code"] = fg_code
                                all_shortage_comps.append(shortage_comps)
                    
                    if all_shortage_comps:
                        combined_shortages = pd.concat(all_shortage_comps, ignore_index=True)
                        
                        # Enrich with Note from materials_df
                        if materials_df is not None and not materials_df.empty:
                            note_col = _find_column(materials_df.columns, ["Note", "Notes", "Remarks", "Comment", "Comments"])
                            if note_col:
                                mat_notes = materials_df[[stock_code_col, note_col]].copy()
                                mat_notes.columns = ["material_code", "note"]
                                mat_notes["material_code"] = mat_notes["material_code"].astype(str).str.strip()
                                mat_notes["note"] = mat_notes["note"].astype(str).fillna("")
                                # Create lookup with canonical codes
                                note_lookup = {}
                                for _, row in mat_notes.iterrows():
                                    raw = row["material_code"]
                                    canon = _canonical_item_code(raw)
                                    note_text = row["note"]
                                    for key in {raw, canon}:
                                        if key and note_text:
                                            note_lookup[key] = note_text
                                combined_shortages["note"] = combined_shortages["component_code"].apply(
                                    lambda c: note_lookup.get(_canonical_item_code(str(c)), "")
                                )
                        
                        shortage_display = combined_shortages[["FG_Code", "component_code", "component_name", "required_qty", "available_oh", "gap_qty"]].copy()
                        if "note" in combined_shortages.columns:
                            shortage_display["note"] = combined_shortages["note"]
                        
                        for col in ["required_qty", "available_oh", "gap_qty"]:
                            if col in shortage_display.columns:
                                shortage_display[col] = pd.to_numeric(shortage_display[col], errors="coerce")
                        shortage_display["shortage_qty"] = -shortage_display["gap_qty"]
                        shortage_display = shortage_display.drop(columns=["gap_qty"])
                        
                        rename_map = {
                            "FG_Code": "FG Code",
                            "component_code": "Material Code",
                            "component_name": "Material Name",
                            "required_qty": "Required",
                            "available_oh": "Available OH",
                            "shortage_qty": "Shortage"
                        }
                        if "note" in shortage_display.columns:
                            rename_map["note"] = "Note"
                        shortage_display = shortage_display.rename(columns=rename_map)
                        
                        styled = shortage_display.style.format({
                            "Required": "{:,.2f}",
                            "Available OH": "{:,.2f}",
                            "Shortage": "{:,.2f}"
                        })
                        
                        # Highlight Note column with yellow background
                        if "Note" in shortage_display.columns:
                            styled = styled.set_properties(
                                subset=["Note"],
                                **{"background-color": "#fff3cd", "color": "#856404", "white-space": "pre-wrap", "max-width": "300px"}
                            )
                        
                        st.dataframe(
                            styled,
                            use_container_width=True,
                        )
                    else:
                        st.info("✅ All materials have sufficient OH.")


def _is_recipe_indicator(name: str) -> bool:
    if not isinstance(name, str):
        return False
    normalized = name.strip().lower()
    if not normalized:
        return False
    return any(token in normalized for token in ("طبخة", "recipe", "mix", "batch"))


def _extract_fg_components(
    fg_code: str,
    bom_df: pd.DataFrame,
    recipe_df: pd.DataFrame,
) -> tuple[pd.DataFrame, float]:
    required_cols = {"BomVersion_ItemId", "ItemId", "ItemName", "BOMQty"}
    if bom_df is None or bom_df.empty or not required_cols.issubset(bom_df.columns):
        return pd.DataFrame(), np.nan

    fg_code_norm = str(fg_code).strip()
    if not fg_code_norm:
        return pd.DataFrame(), np.nan

    subset = bom_df[
        bom_df["BomVersion_ItemId"].astype(str).str.strip() == fg_code_norm
    ].copy()
    if subset.empty:
        return pd.DataFrame(), np.nan

    subset["ItemName"] = subset["ItemName"].astype(str).fillna("")
    recipe_mask = subset["ItemName"].apply(_is_recipe_indicator)
    packaging_rows = subset[~recipe_mask]
    recipe_rows = subset[recipe_mask]

    base_recipe_qty = (
        pd.to_numeric(recipe_rows["BOMQty"], errors="coerce").dropna().iloc[0]
        if not recipe_rows.empty
        else 1.0
    )
    if pd.isna(base_recipe_qty) or base_recipe_qty <= 0:
        base_recipe_qty = 1.0

    recipe_code = (
        str(recipe_rows["ItemId"].iloc[0]).strip()
        if not recipe_rows.empty
        else None
    )

    components: list[dict[str, str | float]] = []
    for _, row in packaging_rows.iterrows():
        qty_value = pd.to_numeric(row.get("BOMQty", 0), errors="coerce")
        qty = float(qty_value) if pd.notna(qty_value) else 0.0
        if qty <= 0:
            continue
        components.append(
            {
                "component_code": str(row.get("ItemId", "")).strip(),
                "component_name": str(row.get("ItemName", "")).strip(),
                "component_type": "Packaging",
                "unit": str(row.get("UnitId", "")).strip(),
                "qty_per_batch": qty,
            }
        )

    if recipe_code and recipe_df is not None and not recipe_df.empty:
        if not {"BomVersion_ItemId", "ItemId", "ItemName", "BOMQty"}.issubset(recipe_df.columns):
            recipe_subset = pd.DataFrame()
        else:
            recipe_subset = recipe_df[
                recipe_df["BomVersion_ItemId"].astype(str).str.strip() == recipe_code
            ].copy()
        if not recipe_subset.empty:
            for _, row in recipe_subset.iterrows():
                qty_value = pd.to_numeric(row.get("BOMQty", 0), errors="coerce")
                qty = float(qty_value) if pd.notna(qty_value) else 0.0
                if qty <= 0:
                    continue
                components.append(
                    {
                        "component_code": str(row.get("ItemId", "")).strip(),
                        "component_name": str(row.get("ItemName", "")).strip(),
                        "component_type": "Raw",
                        "unit": str(row.get("UnitId", "")).strip(),
                        "qty_per_batch": qty,
                    }
                )

    components_df = pd.DataFrame(components)
    return components_df, float(base_recipe_qty)


def compute_fg_material_shortage(
    fg_code: str,
    demand_qty: float,
    bom_df: pd.DataFrame,
    recipe_df: pd.DataFrame,
    materials_stock: dict[str, float],
) -> dict | None:
    components_df, base_recipe_qty = _extract_fg_components(fg_code, bom_df, recipe_df)
    if components_df.empty:
        return None

    base_recipe_qty = base_recipe_qty if base_recipe_qty and base_recipe_qty > 0 else 1.0
    required_qty = float(demand_qty) if demand_qty and demand_qty > 0 else base_recipe_qty
    num_batches = required_qty / base_recipe_qty if base_recipe_qty else 0.0

    evaluated = components_df.copy()
    evaluated["component_code"] = evaluated["component_code"].astype(str).str.strip()

    evaluated["_match_code"] = evaluated["component_code"].apply(_canonical_item_code)

    excluded_components: list[str] = []
    if materials_stock:
        known_codes = {str(code).strip() for code in materials_stock.keys()}
        recognized_mask = evaluated["_match_code"].isin(known_codes)
        if not recognized_mask.all():
            excluded_components = (
                evaluated.loc[~recognized_mask, "component_code"].dropna().unique().tolist()
            )
            evaluated = evaluated.loc[recognized_mask].copy()
    else:
        recognized_mask = pd.Series(True, index=evaluated.index)

    coverage_evaluated = not evaluated.empty
    if not coverage_evaluated:
        return {
            "required_qty": required_qty,
            "max_feasible_qty": float("inf"),
            "shortfall_qty": 0.0,
            "shortage_flag": False,
            "limiting_summary": "",
            "components_table": pd.DataFrame(),
            "excluded_components": excluded_components,
            "coverage_evaluated": False,
        }

    evaluated["required_qty"] = evaluated["qty_per_batch"] * num_batches
    if not materials_stock:
        evaluated["available_oh"] = 0.0
    else:
        evaluated["available_oh"] = (
            evaluated["_match_code"].map(materials_stock).fillna(0.0)
        )
    evaluated["gap_qty"] = evaluated["available_oh"] - evaluated["required_qty"]
    evaluated["is_short"] = evaluated["gap_qty"] < 0

    limiting = evaluated[evaluated["is_short"]]
    if not limiting.empty:
        limiting_sorted = (
            limiting.assign(shortfall=lambda df: -df["gap_qty"])
            .sort_values("shortfall", ascending=False)
        )
        limiting_summary = ", ".join(
            f"{row.component_name or row.component_code}: {row.shortfall:,.0f}"
            for _, row in limiting_sorted.head(3).iterrows()
        )
    else:
        limiting_summary = ""

    feasible_batches: list[float] = []
    for _, row in evaluated.iterrows():
        per_batch = row["qty_per_batch"]
        if per_batch is None or per_batch <= 0:
            continue
        feasible_batches.append(row["available_oh"] / per_batch)

    if feasible_batches:
        max_batches = min(feasible_batches)
    else:
        max_batches = float("inf")

    max_finished_qty = max_batches * base_recipe_qty if np.isfinite(max_batches) else float("inf")
    shortage_flag = not limiting.empty and (not np.isfinite(max_finished_qty) or max_finished_qty < required_qty)
    shortfall_qty = (
        max(required_qty - max_finished_qty, 0.0)
        if shortage_flag and np.isfinite(max_finished_qty)
        else 0.0
    )

    return {
        "required_qty": required_qty,
        "max_feasible_qty": max_finished_qty,
        "shortfall_qty": shortfall_qty,
        "shortage_flag": bool(shortage_flag),
        "limiting_summary": limiting_summary,
        "components_table": evaluated.drop(columns=["_match_code"], errors="ignore"),
        "excluded_components": excluded_components,
        "coverage_evaluated": True,
    }


def render_bom_calculator(materials_df: pd.DataFrame) -> None:
    render_hero_banner(
        title="Affordability + Availability",
        subtitle="Assess BOM cost impact and material coverage in one place.",
        icon="🚩",
        meta="Turn The Planning Up ⬆️💡⬆️",
        gradient="linear-gradient(120deg, #6a11cb 0%, #2575fc 100%)",
    )
    # Use BOM file from sidebar
    # bom_bytes is already defined in sidebar
    using_default_bom = uploaded_bom is None

    if using_default_bom:
        if BOM_DEFAULT_PATH.exists():
            st.caption(f"يتم استخدام الملف الافتراضي: {BOM_DEFAULT_PATH.name}")
        else:
            st.warning("⚠️ لم يتم العثور على ملف BOM افتراضي، يرجى رفع ملف BOM.")

    apply_active_filter = st.checkbox(
        "✅ استخدم فقط الأصناف المفعلة (Active = Yes)",
        value=True,
        key="bom_active_filter",
    )

    try:
        bom_df, recipe_df = load_bom_workbook(bom_bytes, allow_default=using_default_bom)
    except FileNotFoundError as exc:
        st.error(str(exc))
        st.stop()
    except Exception as exc:  # noqa: BLE001
        st.error(f"Unable to read BOM workbook: {exc}")
        st.stop()

    if bom_df.empty:
        st.error("ورقة 'Bom Data' غير موجودة أو فارغة.")
        st.stop()
    if recipe_df.empty:
        st.warning("ورقة 'Recipe Data' غير موجودة أو فارغة، سيتم الاعتماد على بيانات BOM فقط.")

    required_bom_cols = {"BomVersion_ItemId", "ItemName", "ItemId", "UnitId", "BOMQty", "Name"}
    missing_bom_cols = [col for col in required_bom_cols if col not in bom_df.columns]
    if missing_bom_cols:
        st.error(
            "ملف BOM يفتقد للأعمدة التالية: " + ", ".join(missing_bom_cols)
        )
        st.stop()

    recipe_required = {"BomVersion_ItemId", "ItemName", "ItemId", "UnitId", "BOMQty"}
    missing_recipe_cols = [col for col in recipe_required if col not in recipe_df.columns]

    bom_df = bom_df.copy()
    recipe_df = recipe_df.copy()

    if apply_active_filter and "Active" in bom_df.columns:
        bom_df = bom_df[bom_df["Active"].astype(str).str.strip().str.upper() == "YES"].copy()
    if apply_active_filter and "Active" in recipe_df.columns:
        recipe_df = recipe_df[recipe_df["Active"].astype(str).str.strip().str.upper() == "YES"].copy()

    cost_df = materials_df.copy()
    if cost_df.empty:
        st.warning("⚠️ لم يتم تحميل بيانات Materials، سيتم الحساب بدون تكاليف أو أرصدة.")
    else:
        cost_df["ItemNumber"] = cost_df["ItemNumber"].astype(str).str.strip()
        if "ItemName" in cost_df.columns:
            cost_df["ItemName"] = cost_df["ItemName"].astype(str).str.strip()
        for col in ("Cost", "OH"):
            if col not in cost_df.columns:
                cost_df[col] = 0
            cost_df[col] = pd.to_numeric(cost_df[col], errors="coerce").fillna(0.0)
        if apply_active_filter and "Active" in cost_df.columns:
            cost_df = cost_df[cost_df["Active"].astype(str).str.strip().str.upper() == "YES"].copy()

    products_list = (
        bom_df[["Name", "BomVersion_ItemId"]]
        .drop_duplicates(subset=["BomVersion_ItemId"])
        .dropna(subset=["BomVersion_ItemId"])
    )
    product_display = {
        f"{row['Name']} ({row['BomVersion_ItemId']})": row["BomVersion_ItemId"]
        for _, row in products_list.iterrows()
        if pd.notna(row["Name"])
    }

    if not product_display:
        st.info("لا توجد منتجات متاحة بعد تطبيق المرشحات الحالية.")
        st.stop()

    st.subheader("🔍 ابحث عن المنتج التام")
    search_query = st.text_input("اكتب اسم المنتج أو كود المنتج:", key="bom_product_search")

    if search_query:
        q = search_query.strip()
        filtered_by_code = [k for k, v in product_display.items() if q in str(v)]
        filtered_by_name = [k for k in product_display if q.lower() in k.lower()]
        filtered_products = list(dict.fromkeys(filtered_by_code + filtered_by_name))
    else:
        filtered_products = list(product_display.keys())

    selected_display = None
    if filtered_products:
        if len(filtered_products) == 1 and search_query:
            selected_display = filtered_products[0]
        else:
            selected_display = st.selectbox(
                "اختر المنتج من القائمة:",
                options=filtered_products,
                key="bom_product_select",
            )

    if not selected_display:
        st.info("يرجى اختيار منتج تام لعرض تفاصيل BOM.")
        return

    selected_product_code = product_display.get(selected_display)
    batch_qty = st.number_input(
        "🔢 أدخل كمية الإنتاج المطلوبة (كجم)",
        min_value=0.0,
        value=1000.0,
        step=100.0,
        key="bom_batch_qty",
    )

    if batch_qty <= 0:
        st.warning("⚠️ الرجاء إدخال كمية إنتاج أكبر من صفر.")
        return

    bom_items = bom_df[
        bom_df["BomVersion_ItemId"].astype(str) == str(selected_product_code)
    ].copy()
    if bom_items.empty:
        st.warning("لا توجد مكونات BOM مرتبطة بالمنتج المختار.")
        return

    recipe_row = bom_items[bom_items["ItemName"].str.contains("طبخة", case=False, na=False)]
    if recipe_row.empty:
        base_recipe_qty = 1.0
        num_batches = batch_qty
        recipe_code = None
    else:
        base_recipe_qty = (
            pd.to_numeric(recipe_row["BOMQty"], errors="coerce").fillna(0.0).iloc[0]
        )
        base_recipe_qty = base_recipe_qty if base_recipe_qty > 0 else 1.0
        recipe_code = str(recipe_row["ItemId"].iloc[0]).strip()
        num_batches = batch_qty / base_recipe_qty

    results: list[dict[str, float | str]] = []

    bom_pack = bom_items[~bom_items["ItemName"].str.contains("طبخة", case=False, na=False)].copy()
    for _, row in bom_pack.iterrows():
        qty_value = pd.to_numeric(row.get("BOMQty", 0), errors="coerce")
        qty = float(qty_value) if pd.notna(qty_value) else 0.0
        results.append(
            {
                "النوع": "تعبئة وتغليف",
                "اسم الصنف": str(row.get("ItemName", "")).strip(),
                "كود الصنف": str(row.get("ItemId", "")).strip(),
                "الوحدة": str(row.get("UnitId", "")).strip(),
                "الكمية الأصلية (لكل طبخة)": qty,
                "الكمية المطلوبة": qty * num_batches,
            }
        )

    if recipe_code and not recipe_df.empty and not missing_recipe_cols:
        recipe_items = recipe_df[
            recipe_df["BomVersion_ItemId"].astype(str) == str(recipe_code)
        ].copy()
        for _, row in recipe_items.iterrows():
            qty_value = pd.to_numeric(row.get("BOMQty", 0), errors="coerce")
            qty = float(qty_value) if pd.notna(qty_value) else 0.0
            results.append(
                {
                    "النوع": "مواد خام",
                    "اسم الصنف": str(row.get("ItemName", "")).strip(),
                    "كود الصنف": str(row.get("ItemId", "")).strip(),
                    "الوحدة": str(row.get("UnitId", "")).strip(),
                    "الكمية الأصلية (لكل طبخة)": qty,
                    "الكمية المطلوبة": qty * num_batches,
                }
            )

    if not results:
        st.info("لا توجد مكونات قابلة للعرض بعد تطبيق المرشحات.")
        return

    combined_df = pd.DataFrame(results)
    combined_df["الكمية الأصلية (لكل طبخة)"] = pd.to_numeric(
        combined_df["الكمية الأصلية (لكل طبخة)"], errors="coerce"
    ).fillna(0.0)
    combined_df["الكمية المطلوبة"] = pd.to_numeric(
        combined_df["الكمية المطلوبة"], errors="coerce"
    ).fillna(0.0)
    combined_df["كود الصنف"] = combined_df["كود الصنف"].astype(str).str.strip()

    if not cost_df.empty and "ItemNumber" in cost_df.columns:
        merged = combined_df.merge(
            cost_df[["ItemNumber", "ItemName", "Cost", "OH"]],
            how="left",
            left_on="كود الصنف",
            right_on="ItemNumber",
        )
        merged["تكلفة الوحدة"] = pd.to_numeric(merged["Cost"], errors="coerce").fillna(0.0)
        merged["الرصيد المتاح (OH)"] = pd.to_numeric(merged["OH"], errors="coerce").fillna(0.0)
        merged["إجمالي التكلفة"] = merged["الكمية المطلوبة"] * merged["تكلفة الوحدة"]
        merged.drop(columns=["ItemNumber", "ItemName", "Cost", "OH"], inplace=True)
        combined_df = merged
    else:
        combined_df["تكلفة الوحدة"] = 0.0
        combined_df["الرصيد المتاح (OH)"] = 0.0
        combined_df["إجمالي التكلفة"] = 0.0

    total_cost = combined_df["إجمالي التكلفة"].sum()
    st.dataframe(combined_df, use_container_width=True)

    req_per_batch = combined_df[[
        "كود الصنف",
        "اسم الصنف",
        "النوع",
        "الوحدة",
        "الكمية الأصلية (لكل طبخة)",
        "الرصيد المتاح (OH)",
    ]].copy()
    req_per_batch = req_per_batch[req_per_batch["الكمية الأصلية (لكل طبخة)"] > 0]
    req_per_batch["الطبخات الممكنة لكل مكوّن"] = req_per_batch.apply(
        lambda r: (r["الرصيد المتاح (OH)"] / r["الكمية الأصلية (لكل طبخة)"])
        if r["الكمية الأصلية (لكل طبخة)"] > 0
        else np.nan,
        axis=1,
    )
    max_batches_possible = (
        req_per_batch["الطبخات الممكنة لكل مكوّن"].replace([np.inf], pd.NA).min()
        if not req_per_batch.empty
        else 0.0
    )
    max_batches_possible = float(max_batches_possible) if pd.notna(max_batches_possible) else 0.0
    max_finished_qty_kg = max_batches_possible * base_recipe_qty

    st.subheader("🧮 أقصى كمية ممكن إنتاجها بناءً على الأرصدة")
    c1, c2 = st.columns(2)
    with c1:
        st.metric("🏭 أقصى عدد طبخات ممكن", f"{max_batches_possible:,.2f}")
    with c2:
        st.metric("📦 أقصى كمية منتج تام ممكنة (كجم)", f"{max_finished_qty_kg:,.2f}")

    st.subheader("⛔ المكونات المقيِّدة للإنتاج")
    limiting = req_per_batch.sort_values(
        "الطبخات الممكنة لكل مكوّن",
        ascending=True,
    ).head(10)
    st.dataframe(limiting, use_container_width=True)

    st.subheader("📝 Insights")
    insights: list[str] = []
    if max_batches_possible == 0:
        insights.append(
            "لا توجد طبخات ممكنة حاليًا بسبب نفاد بعض المكونات أو عدم توفر رصيد OH."
        )
    else:
        insights.append(
            f"أقصى إنتاج ممكن: {max_batches_possible:,.2f} طبخة = {max_finished_qty_kg:,.2f} كجم من المنتج التام."
        )
    if not limiting.empty:
        lim_names = ", ".join(limiting.head(3)["اسم الصنف"].tolist())
        insights.append(f"أكثر ثلاث مكونات تقييدًا: {lim_names}.")
    if "تكلفة الوحدة" in combined_df.columns and not combined_df.empty:
        cost_by_type = combined_df.groupby("النوع", as_index=False)["إجمالي التكلفة"].sum()
        if not cost_by_type.empty:
            high_cost_share = cost_by_type.sort_values("إجمالي التكلفة", ascending=False).iloc[0]
            insights.append(
                f"أعلى مساهمة في التكلفة تأتي من فئة: {high_cost_share['النوع']}"
                f" بقيمة {high_cost_share['إجمالي التكلفة']:,.2f} جنيه."
            )
        top_cost_component = combined_df.sort_values("إجمالي التكلفة", ascending=False).iloc[0]
        insights.append(
            f"أعلى مكوّن تكلفة: {top_cost_component['اسم الصنف']}"
            f" بإجمالي {top_cost_component['إجمالي التكلفة']:,.2f} جنيه."
        )
    if base_recipe_qty > 0:
        insights.append(f"وزن الطبخة الأساسي: {base_recipe_qty:,.2f} كجم.")

    for idx, txt in enumerate(insights, 1):
        st.write(f"- **Insight {idx}:** {txt}")

    output = io.BytesIO()
    timestamp = datetime.now()
    try:
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            combined_df.to_excel(writer, index=False, sheet_name="جميع المكونات")
            req_per_batch.to_excel(writer, index=False, sheet_name="تحليل الأرصدة")
            limiting.to_excel(writer, index=False, sheet_name="مكونات مقيِّدة")
            summary = pd.DataFrame(
                [
                    {
                        "المنتج": selected_display,
                        "كود المنتج": selected_product_code,
                        "رقم الطبخة": recipe_code if recipe_code else "غير متوفر",
                        "الكمية المطلوبة (كجم)": batch_qty,
                        "وزن الطبخة الأساسي (كجم)": base_recipe_qty,
                        "عدد الطبخات المحسوبة": round(num_batches, 4),
                        "أقصى عدد طبخات ممكن": round(max_batches_possible, 4),
                        "أقصى كمية منتج تام ممكنة (كجم)": round(max_finished_qty_kg, 2),
                        "إجمالي تكلفة الإنتاج": round(total_cost, 2),
                        "تاريخ الإنشاء": timestamp.strftime("%Y-%m-%d %H:%M:%S"),
                    }
                ]
            )
            summary.to_excel(writer, index=False, sheet_name="الملخص")
    except ImportError:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            combined_df.to_excel(writer, index=False, sheet_name="جميع المكونات")
            req_per_batch.to_excel(writer, index=False, sheet_name="تحليل الأرصدة")
            limiting.to_excel(writer, index=False, sheet_name="مكونات مقيِّدة")
            summary = pd.DataFrame(
                [
                    {
                        "المنتج": selected_display,
                        "كود المنتج": selected_product_code,
                        "رقم الطبخة": recipe_code if recipe_code else "غير متوفر",
                        "الكمية المطلوبة (كجم)": batch_qty,
                        "وزن الطبخة الأساسي (كجم)": base_recipe_qty,
                        "عدد الطبخات المحسوبة": round(num_batches, 4),
                        "أقصى عدد طبخات ممكن": round(max_batches_possible, 4),
                        "أقصى كمية منتج تام ممكنة (كجم)": round(max_finished_qty_kg, 2),
                        "إجمالي تكلفة الإنتاج": round(total_cost, 2),
                        "تاريخ الإنشاء": timestamp.strftime("%Y-%m-%d %H:%M:%S"),
                    }
                ]
            )
            summary.to_excel(writer, index=False, sheet_name="الملخص")
    output.seek(0)

    product_name = str(selected_display).split("(")[0].strip().replace(" ", "_")[:30]
    filename = f"BOM_{product_name}_إنتاج_{int(batch_qty)}كجم.xlsx"
    st.download_button(
        label="⬇️ تحميل النتائج كملف Excel",
        data=output.getvalue(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="bom_download_results",
    )

    st.markdown("---")
    st.subheader("🔍 تحليل خامة بناءً على الرصيد المتاح")

    if cost_df.empty:
        st.info("لا تتوفر بيانات تكلفة لتحليل الخامات.")
        return

    def classify_usage(name: str) -> str:
        if pd.isna(name):
            return "Other"
        name = str(name).strip()
        if "تصدير" in name:
            return "Export"
        if "محلي" in name or "محلى" in name:
            return "Local"
        return "Other"

    def calculate_fg_total_cost(fg_code: str) -> float:
        fg_bom = bom_df[bom_df["BomVersion_ItemId"] == fg_code][["ItemId", "BOMQty", "ItemName"]]
        fg_recipe_links = fg_bom[
            fg_bom["ItemName"].str.contains("طبخة", case=False, na=False)
        ]["ItemId"].tolist()
        fg_recipe = recipe_df[recipe_df["BomVersion_ItemId"].isin(fg_recipe_links)][["ItemId", "BOMQty"]]

        all_components = pd.concat(
            [fg_bom[["ItemId", "BOMQty"]], fg_recipe], ignore_index=True
        )
        all_components = all_components.merge(
            cost_df[["ItemNumber", "Cost"]],
            left_on="ItemId",
            right_on="ItemNumber",
            how="left",
        )
        all_components["Cost"] = pd.to_numeric(all_components["Cost"], errors="coerce").fillna(0.0)
        all_components["BOMQty"] = pd.to_numeric(
            all_components["BOMQty"], errors="coerce"
        ).fillna(0.0)
        all_components["comp_total"] = all_components["BOMQty"] * all_components["Cost"]
        return float(all_components["comp_total"].sum())

    component_query = st.text_input(
        "اكتب اسم الخامة أو جزء منه (من ملف Materials):",
        key="bom_component_search",
    )
    if not component_query:
        return

    suggestions = cost_df[
        cost_df["ItemName"].str.contains(component_query, case=False, na=False)
    ]
    if suggestions.empty:
        st.warning("⚠️ لا توجد خامات مطابقة للبحث")
        return

    component_selected = st.selectbox(
        "اختر الخامة:",
        options=suggestions["ItemName"].tolist(),
        key="bom_component_select",
    )
    if not component_selected:
        return

    comp_row = suggestions[suggestions["ItemName"] == component_selected].iloc[0]
    comp_code = str(comp_row["ItemNumber"])
    comp_oh = float(comp_row["OH"])
    comp_cost = float(comp_row.get("Cost", 0.0))

    st.info(
        f"✅ الخامة المختارة: {component_selected} | الكود: {comp_code} | "
        f"الرصيد المتاح: {comp_oh} | تكلفة الوحدة: {comp_cost}"
    )

    results_rows: list[dict[str, float | str | None]] = []

    used_in_bom = bom_df[bom_df["ItemId"] == comp_code][["BomVersion_ItemId", "BOMQty", "Name"]]
    for _, row in used_in_bom.iterrows():
        fg_code = str(row["BomVersion_ItemId"])
        fg_name = str(row["Name"])
        qty_value = pd.to_numeric(row["BOMQty"], errors="coerce")
        qty_per_batch = float(qty_value) if pd.notna(qty_value) else 0.0
        if qty_per_batch <= 0:
            continue
        max_batches = comp_oh / qty_per_batch if qty_per_batch else 0
        recipe_row_fg = bom_df[
            (bom_df["BomVersion_ItemId"] == fg_code)
            & (bom_df["Name"].str.contains("طبخة", case=False, na=False))
        ]
        base_qty = (
            pd.to_numeric(recipe_row_fg["BOMQty"], errors="coerce").fillna(0.0).iloc[0]
            if not recipe_row_fg.empty
            else 1.0
        )
        max_finished_qty = max_batches * base_qty
        fg_total_cost = calculate_fg_total_cost(fg_code)
        cost_in_product = qty_per_batch * comp_cost
        share = (cost_in_product / fg_total_cost * 100) if fg_total_cost > 0 else None

        results_rows.append(
            {
                "نوع الخامة": "Pack",
                "المنتج": fg_name,
                "كود المنتج": fg_code,
                "الكمية لكل طبخة": qty_per_batch,
                "أقصى عدد طبخات ممكنة": round(max_batches, 2),
                "أقصى كمية منتج تام ممكنة (كجم)": round(max_finished_qty, 2),
                "تصنيف الاستخدام": classify_usage(fg_name),
                "تكلفة الخامة في المنتج": cost_in_product,
                "إجمالي تكلفة المنتج": fg_total_cost,
                "نسبة تكلفة الخامة من المنتج (%)": share,
            }
        )

    used_in_recipe = recipe_df[recipe_df["ItemId"] == comp_code][["BomVersion_ItemId", "BOMQty", "Name"]]
    for _, row in used_in_recipe.iterrows():
        recipe_code = str(row["BomVersion_ItemId"])
        recipe_name = str(row["Name"])
        qty_value = pd.to_numeric(row["BOMQty"], errors="coerce")
        qty_per_batch = float(qty_value) if pd.notna(qty_value) else 0.0
        if qty_per_batch <= 0:
            continue
        max_batches = comp_oh / qty_per_batch if qty_per_batch else 0
        recipe_row_recipe = recipe_df[
            (recipe_df["BomVersion_ItemId"] == recipe_code)
            & (recipe_df["Name"].str.contains("طبخة", case=False, na=False))
        ]
        base_qty = (
            pd.to_numeric(recipe_row_recipe["BOMQty"], errors="coerce").fillna(0.0).iloc[0]
            if not recipe_row_recipe.empty
            else 1.0
        )
        max_finished_qty = max_batches * base_qty

        fg_links = (
            bom_df[bom_df["ItemId"] == recipe_code][["BomVersion_ItemId", "Name"]]
            .drop_duplicates()
        )
        if fg_links.empty:
            fg_total_cost = 0.0
            cost_in_product = qty_per_batch * comp_cost
            results_rows.append(
                {
                    "نوع الخامة": "Raw",
                    "المنتج": recipe_name,
                    "كود المنتج": recipe_code,
                    "الكمية لكل طبخة": qty_per_batch,
                    "أقصى عدد طبخات ممكنة": round(max_batches, 2),
                    "أقصى كمية منتج تام ممكنة (كجم)": round(max_finished_qty, 2),
                    "تصنيف الاستخدام": classify_usage(recipe_name),
                    "تكلفة الخامة في المنتج": cost_in_product,
                    "إجمالي تكلفة المنتج": fg_total_cost,
                    "نسبة تكلفة الخامة من المنتج (%)": None,
                }
            )
        else:
            for _, fg in fg_links.iterrows():
                fg_code = str(fg["BomVersion_ItemId"])
                fg_name = str(fg["Name"])
                fg_total_cost = calculate_fg_total_cost(fg_code)
                cost_in_product = qty_per_batch * comp_cost
                share = (cost_in_product / fg_total_cost * 100) if fg_total_cost > 0 else None
                results_rows.append(
                    {
                        "نوع الخامة": "Raw",
                        "المنتج": fg_name,
                        "كود المنتج": fg_code,
                        "الكمية لكل طبخة": qty_per_batch,
                        "أقصى عدد طبخات ممكنة": round(max_batches, 2),
                        "أقصى كمية منتج تام ممكنة (كجم)": round(max_finished_qty, 2),
                        "تصنيف الاستخدام": classify_usage(fg_name),
                        "تكلفة الخامة في المنتج": cost_in_product,
                        "إجمالي تكلفة المنتج": fg_total_cost,
                        "نسبة تكلفة الخامة من المنتج (%)": share,
                    }
                )

    if not results_rows:
        st.info("❕ هذه الخامة غير مستخدمة في أي منتج تام أو طبخة")
        return

    results_df = pd.DataFrame(results_rows)
    if comp_oh > 0:
        results_df["نسبة استهلاك الخامة (%)"] = (
            results_df["الكمية لكل طبخة"] / comp_oh
        ) * 100
    else:
        results_df["نسبة استهلاك الخامة (%)"] = np.nan

    st.subheader("📊 نتائج تحليل الخامة")
    st.dataframe(results_df, use_container_width=True)

    insights_components: list[str] = []
    if results_df["نسبة تكلفة الخامة من المنتج (%)"].notna().any():
        top_share = results_df.sort_values(
            "نسبة تكلفة الخامة من المنتج (%)",
            ascending=False,
        ).iloc[0]
        insights_components.append(
            f"🔎 أعلى اعتماد على الخامة: {top_share['المنتج']} "
            f"تمثل {top_share['نسبة تكلفة الخامة من المنتج (%)']:.2f}% من إجمالي تكلفته."
        )

    top_qty = results_df.sort_values("الكمية لكل طبخة", ascending=False).iloc[0]
    insights_components.append(
        f"📦 أكثر منتج يستهلك الخامة ككمية مطلقة: {top_qty['المنتج']} "
        f"بكمية {top_qty['الكمية لكل طبخة']:.2f} لكل طبخة."
    )

    top_prod = results_df.sort_values(
        "أقصى كمية منتج تام ممكنة (كجم)", ascending=False
    ).iloc[0]
    insights_components.append(
        f"🔝 أكثر منتج يمكن إنتاجه من الخامة: {top_prod['المنتج']} "
        f"بكمية {top_prod['أقصى كمية منتج تام ممكنة (كجم)']:.2f} كجم."
    )

    usage_summary = results_df.groupby("تصنيف الاستخدام")[
        "أقصى كمية منتج تام ممكنة (كجم)"
    ].sum()
    for usage, qty in usage_summary.items():
        insights_components.append(f"🌍 إجمالي استخدام {usage}: {qty:,.2f} كجم")

    st.subheader("📝 Insights")
    for idx, txt in enumerate(insights_components, 1):
        st.write(f"- **Insight {idx}:** {txt}")


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
    df = pd.read_excel(data_source, sheet_name=SHEET_NAME, engine="pyxlsb", dtype={"ItemNumber": str})

    # Prepare ItemNumber and remove blank rows
    if "ItemNumber" in df.columns:
        df["ItemNumber"] = df["ItemNumber"].astype(str).str.strip()
        valid_items_mask = df["ItemNumber"].ne("") & df["ItemNumber"].str.lower().ne("nan")
        df = df.loc[valid_items_mask].copy()
    else:
        # Fallback if ItemNumber column doesn't exist
        df = df.dropna(how="all")
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
st.sidebar.header("📂 Data Sources")

# 1. Materials XLSB
st.sidebar.subheader("📦 Materials Data")
uploaded_materials = st.sidebar.file_uploader(
    "Upload Materials XLSB",
    type=["xlsb"],
    help="Main materials inventory data file",
    key="materials_uploader"
)
materials_bytes = uploaded_materials.getvalue() if uploaded_materials else None
try:
    df = load_data(materials_bytes)
except FileNotFoundError:
    df = pd.DataFrame()
    st.sidebar.warning("⚠️ Default Materials file not found")
except Exception as exc:
    df = pd.DataFrame()
    st.sidebar.error(f"Error loading Materials: {exc}")

if uploaded_materials:
    st.sidebar.success(f"✓ {uploaded_materials.name}")
elif not df.empty:
    st.sidebar.caption(f"Using: {DEFAULT_DATA_PATH.name}")

# 2. CalcHelp Sheet
st.sidebar.subheader("🧮 CalcHelp Sheet")
calc_help_file = st.sidebar.file_uploader(
    "Upload CalcHelp",
    type=["xlsb", "xlsx"],
    help="Monthly inventory calculation logic",
    key="calc_help_uploader"
)
calc_help_bytes = calc_help_file.getvalue() if calc_help_file else None
calc_help_df = load_calc_help(materials_bytes, calc_help_bytes)
if calc_help_file:
    st.sidebar.success(f"✓ {calc_help_file.name}")
else:
    st.sidebar.caption("Using bundled CalcHelp")

# 3. FG (Finished Goods) XLSB
st.sidebar.subheader("🏷️ Finished Goods Data")
uploaded_fg = st.sidebar.file_uploader(
    "Upload FG XLSB",
    type=["xlsb"],
    help="Finished goods data for FG Explorer",
    key="fg_uploader"
)
fg_bytes = uploaded_fg.getvalue() if uploaded_fg else None
if uploaded_fg:
    st.sidebar.success(f"✓ {uploaded_fg.name}")
else:
    st.sidebar.caption("Using default FG file")

# 4. BOM/Recipe XLSX
st.sidebar.subheader("📋 BOM & Recipe Data")
uploaded_bom = st.sidebar.file_uploader(
    "Upload BOM/Recipe XLSX",
    type=["xlsx"],
    help="Bill of Materials and Recipe data",
    key="bom_uploader"
)
bom_bytes = uploaded_bom.getvalue() if uploaded_bom else None
if uploaded_bom:
    st.sidebar.success(f"✓ {uploaded_bom.name}")
else:
    st.sidebar.caption(f"Using: {BOM_DEFAULT_PATH.name}")

st.sidebar.markdown("---")
app_view = st.sidebar.radio(
    "📑 Select view",
    ("Inventory Simulator", "FG Explorer", "BOM Calculator"),
    index=0,
    key="main_app_view",
)

if app_view == "FG Explorer":
    render_fg_explorer(df)
    st.stop()
elif app_view == "BOM Calculator":
    render_bom_calculator(df)
    st.stop()

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

# Check if Materials data is loaded
if df.empty:
    st.sidebar.info("📦 Please upload Materials data to use filters")
    df_filtered = df.copy()
    search_active = False
else:
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

# Check if Materials data is loaded before calculations
if df.empty:
    df_with_months = pd.DataFrame()
    df_grouped = pd.DataFrame()
else:
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
    render_hero_banner(
        title="Inventory Simulator",
        subtitle=f"Daily snapshot · {today.strftime('%d %B %Y')}",
        icon="📦",
        meta="Turn The Planning Up ⬆️💡⬆️",
        gradient="linear-gradient(120deg, #00c6ff 0%, #0072ff 50%, #0041ff 100%)",
    )
    
    # Check if Materials data is loaded
    if df.empty:
        st.info("📦 Please upload Materials data to start using the Inventory Simulator")
        st.markdown("""
        ### How to get started:
        1. Upload your Materials XLSB file using the sidebar
        2. Optionally upload CalcHelp, FG, and BOM files
        3. Use filters to analyze your inventory
        """)
        return

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

        st.markdown("### Factory OH Value")

        card_chunk_size = 4
        for start in range(0, len(factory_cards_summary), card_chunk_size):
            chunk = factory_cards_summary.iloc[start:start + card_chunk_size]
            cols = st.columns(len(chunk))
            for col, (_, row) in zip(cols, chunk.iterrows()):
                with col:
                    share = row['OH_value'] / total_factory_oh_value if total_factory_oh_value else np.nan
                    delta_text = format_percentage(share) if not np.isnan(share) else "-"
                    factory_name = row['Factory']
                    icon = _get_factory_icon(factory_name)
                    # Use larger icon in metric label
                    st.markdown(f"<h3 style='font-size: 1.5em; margin: 0;'>{icon}</h3>", unsafe_allow_html=True)
                    st.metric(
                        label=factory_name,
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