"""
Customer service dashboard (Streamlit page).

Migrated from the JS dashboard "سجل عقود الوحدات والأقساط" (index.html / app.js / style.css).
Adds a Project filter (All Projects / Gardenia 3 / Gardenia Town) that is applied to the data
BEFORE any KPI, chart or table is calculated.

Data is never hardcoded: it is read from DATA_DIR (see CONFIGURATION).
"""
from __future__ import annotations

import calendar
import os
import re
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from theme import ACCENT_GOLD, ACCENT_TEAL, NAVY, NAVY_LIGHT, render_header

# =====================================================================
# CONFIGURATION  (change paths / projects here only)
# =====================================================================
# <project root>/data  (this file lives in <project root>/views/)
DATA_DIR = Path(os.environ.get("CUSTOMER_SERVICE_DATA_DIR", Path(__file__).resolve().parent.parent / "data"))

# Preferred source: the combined file produced by Power Query (Gardenia 3 + Gardenia Town).
COMBINED_FILE = DATA_DIR / "Customer_Service_All.xlsx"
COMBINED_SHEET = 0

# Fallback source (used only if COMBINED_FILE does not exist): one entry per project.
#   contracts    -> contract register (contract / handover dates, floor, area, value ...)
#   installments -> optional installment follow-up sheet (remaining installments + notes),
#                   matched to the contracts by Building + Unit.
PROJECT_SOURCES = {
    "Gardenia 3": {
        "contracts": DATA_DIR / "G3-DB-Customers.xlsx",
        "installments": DATA_DIR / "dataexcel.xlsx",
    },
}

G3_PROJECT = "Gardenia 3"
TOWN_PROJECT = "Gardenia Town"
PROJECTS = [G3_PROJECT, TOWN_PROJECT]

# Gardenia Town = overdue-installments (arrears) report: a different dataset from the G3 contracts register.
TOWN_GLOBS = ("gtown_customers*.xlsx", "g_town-installments-delayed*.xlsx")  # newest matching file in DATA_DIR is used
CURRENCY = "EGP"                                 # label only; change if the amounts are in another currency
ALL_PROJECTS = "All Projects"
ENDING_SOON_DAYS = 90
ARABIC_DIGITS = False  # True -> show ٠١٢٣ like the original JS (ar-EG) dashboard
DEPARTMENT = "Customer Service"

# =====================================================================
# CONSTANTS
# =====================================================================
REQUIRED_COLUMNS = [
    "project", "client_name", "building_no", "unit_no", "area_sqm",
    "contract_date", "handover_date", "remaining_installments", "notes",
]
OPTIONAL_COLUMNS = ["floor", "unit_value", "delivery_term", "contract_notes"]

CONTRACT_COLS = {  # Arabic header in the contract register -> canonical name
    "أسم العميل": "client_name", "اسم العميل": "client_name",
    "تاريخ التعاقد": "contract_date", "رقم الوحدة": "unit_no", "رقم العمارة": "building_no",
    "الدور": "floor", "المساحة": "area_sqm", "قيمة الوحدة": "unit_value",
    "تاريخ الإستلام": "handover_date", "مهلة التسليم": "delivery_term", "ملاحظات": "contract_notes",
}
INSTALLMENT_COLS = {
    "العمارة": "building_no", "الوحدة": "unit_no",
    "عدد الأقساط المتبقية": "remaining_installments", "ملاحظات": "notes",
}

PROJECT_ALIASES = {
    "gardenia 3": "Gardenia 3", "gardenia3": "Gardenia 3", "g3": "Gardenia 3", "جاردينيا 3": "Gardenia 3",
    "gardenia town": "Gardenia Town", "gardenia-town": "Gardenia Town", "town": "Gardenia Town",
    "جاردينيا تاون": "Gardenia Town",
}
FLOOR_FIXES = {"الارضي": "الأرضي", "الاول": "الأول", "الاولى": "الأولى"}  # spelling variants only

STATUS_ORDER = ["منتظم", "ملتزم", "إلى حد ما منتظم", "غير منتظم", "غير ملتزم", "متأخر", "أخرى"]
STATUS_COLORS = {
    "منتظم": "#27a85a", "ملتزم": "#17a398", "إلى حد ما منتظم": "#ee9716",
    "غير منتظم": "#d94343", "غير ملتزم": "#9b2c2c", "متأخر": "#7141c7", "أخرى": "#9aa1ae",
}
BUCKET_ORDER = ["0 (fully paid)", "1–6", "7–12", "13–24", "More than 24"]
DANGER = "#C0563B"

# =====================================================================
# PURE HELPERS (ported 1:1 from app.js where applicable)
# =====================================================================
def parse_number(value) -> float:
    """JS parseNumber: strip everything but digits . - ; blank/invalid -> NaN (JS returned 0, see notes)."""
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return float("nan")
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = re.sub(r"[^\d.\-]", "", str(value).replace(",", ""))
    try:
        return float(cleaned)
    except ValueError:
        return float("nan")


def get_status(notes) -> str:
    """JS getStatus(notes): same keywords, same priority order."""
    text = "" if notes is None or (not isinstance(notes, str) and pd.isna(notes)) else str(notes)
    text = re.sub(r"\s+", " ", text).strip().lower()
    if not text:
        return "أخرى"
    if "غير ملتزم" in text:
        return "غير ملتزم"
    if any(k in text for k in ("متأخر", "متاخر", "متأخرات")):
        return "متأخر"
    if "غير منتظم" in text or "مش منتظم" in text:
        return "غير منتظم"
    if any(k in text for k in ("إلى حد ما منتظم", "الى حد ما منتظم", "الي حد ما منتظم",
                               "إلى حد ما ملتزم", "الى حد ما ملتزم", "الي حد ما ملتزم")):
        return "إلى حد ما منتظم"
    if "ملتزم" in text:
        return "ملتزم"
    if "منتظم" in text:
        return "منتظم"
    return "أخرى"


def _digits(text: str) -> str:
    return text.translate(str.maketrans("0123456789,.", "٠١٢٣٤٥٦٧٨٩٬٫")) if ARABIC_DIGITS else text


def fmt_int(x) -> str:
    return "—" if pd.isna(x) else _digits(f"{x:,.0f}")


def fmt_dec(x, digits: int = 1) -> str:
    return "—" if pd.isna(x) else _digits(f"{x:,.{digits}f}")


def _key_part(value) -> str:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    text = re.sub(r"\s+", " ", str(value)).strip()
    return text.split(".")[0] if re.fullmatch(r"\d+\.0+", text) else text


def _ending_soon_mask(dates: pd.Series, today: pd.Timestamp) -> pd.Series:
    return dates.between(today, today + pd.Timedelta(days=ENDING_SOON_DAYS))  # NaT -> False


def _unit_label(df: pd.DataFrame) -> pd.Series:
    return "عمارة " + df["building_no"] + " – وحدة " + df["unit_no"]


# =====================================================================
# DATA LOADING
# =====================================================================
def _read_table(path: Path, anchor: str) -> pd.DataFrame:
    """Read a sheet whose header row is not necessarily the first row (finds the row containing `anchor`)."""
    raw = pd.read_excel(path, header=None, dtype=object)
    header_idx = next(
        (i for i, row in raw.iterrows() if anchor in [str(v).strip() for v in row.values]), None
    )
    if header_idx is None:
        raise ValueError(f"Header column '{anchor}' not found in {path.name}")
    table = raw.iloc[header_idx + 1:].copy()
    table.columns = [str(c).strip() for c in raw.iloc[header_idx]]
    return table.reset_index(drop=True)


def _select_renamed(table: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    renamed = table.rename(columns=mapping)
    wanted = list(dict.fromkeys(mapping.values()))
    renamed = renamed.loc[:, ~renamed.columns.duplicated()]
    return renamed[[c for c in wanted if c in renamed.columns]]


def _load_project_sources(project: str, cfg: dict) -> tuple[pd.DataFrame, list[str]]:
    notes: list[str] = []
    contracts = _select_renamed(_read_table(cfg["contracts"], "تاريخ التعاقد"), CONTRACT_COLS)
    contracts = contracts.dropna(subset=["building_no", "unit_no"], how="all").reset_index(drop=True)
    contracts["_key"] = contracts["building_no"].map(_key_part) + "|" + contracts["unit_no"].map(_key_part)

    installments_path = cfg.get("installments")
    if installments_path and Path(installments_path).exists():
        inst = _select_renamed(_read_table(Path(installments_path), "عدد الأقساط المتبقية"), INSTALLMENT_COLS)
        inst = inst.dropna(subset=["building_no", "unit_no"], how="all")
        inst["_key"] = inst["building_no"].map(_key_part) + "|" + inst["unit_no"].map(_key_part)
        dup = int(inst["_key"].duplicated().sum())
        inst_unique = inst.drop_duplicates("_key", keep="first")
        merged = contracts.merge(
            inst_unique[["_key", "remaining_installments", "notes"]], on="_key", how="left"
        )
        unmatched = int(merged["remaining_installments"].isna().sum())
        orphan = len(set(inst_unique["_key"]) - set(contracts["_key"]))
        notes.append(f"{project}: {len(contracts)} عقد؛ {unmatched} عقد بدون بيانات أقساط مطابقة؛ "
                     f"{orphan} وحدة في ملف الأقساط بدون عقد؛ {dup} وحدة مكررة في ملف الأقساط (أُخذ أول سجل).")
        contracts = merged
    else:
        contracts["remaining_installments"] = pd.NA
        contracts["notes"] = pd.NA
        notes.append(f"{project}: لا يوجد ملف أقساط — الأقساط والحالة غير متاحة.")

    contracts["project"] = project
    return contracts.drop(columns="_key"), notes


def _load_combined(path: Path) -> tuple[pd.DataFrame, list[str]]:
    df = pd.read_excel(path, sheet_name=COMBINED_SHEET)
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"{path.name} is missing required columns: {', '.join(missing)}")
    return df, [f"المصدر: {path.name}"]


@st.cache_data(show_spinner="جاري تحميل البيانات...")
def load_data(signature: tuple) -> tuple[pd.DataFrame, list[str]]:
    """`signature` (file paths + modified times) only exists so the cache refreshes when files change."""
    if COMBINED_FILE.exists():
        df, notes = _load_combined(COMBINED_FILE)
    else:
        frames, notes = [], []
        for project, cfg in PROJECT_SOURCES.items():
            if Path(cfg["contracts"]).exists():
                frame, frame_notes = _load_project_sources(project, cfg)
                frames.append(frame)
                notes += frame_notes
        if not frames:
            raise FileNotFoundError("No data source found")
        df = pd.concat(frames, ignore_index=True)
    return clean_data(df), notes


def _source_signature() -> tuple:
    paths = [COMBINED_FILE] + [Path(p) for cfg in PROJECT_SOURCES.values() for p in cfg.values()]
    return tuple((str(p), p.stat().st_mtime if p.exists() else None) for p in paths)


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Safe, meaning-preserving cleaning only (whitespace, spelling variants, types). Dates are NOT altered."""
    df = df.copy()
    for col in OPTIONAL_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA

    for col in ["project", "client_name", "building_no", "unit_no", "floor", "notes", "contract_notes"]:
        df[col] = df[col].map(_key_part)
    df["project"] = df["project"].map(lambda p: PROJECT_ALIASES.get(p.lower(), p))
    df["floor"] = df["floor"].replace(FLOOR_FIXES)

    for col in ["contract_date", "handover_date"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")  # already real dates in the source
    df["area_sqm"] = df["area_sqm"].map(parse_number)
    df["unit_value"] = pd.to_numeric(df["unit_value"], errors="coerce")
    df["remaining_installments"] = df["remaining_installments"].map(parse_number)

    df["status_category"] = df["notes"].map(get_status)
    df["unit_label"] = _unit_label(df)
    df["search_text"] = (df["client_name"] + " " + df["notes"] + " " + df["contract_notes"]).str.lower()
    return df


# =====================================================================
# FILTERING
# =====================================================================
def _scope(df: pd.DataFrame, project: str) -> pd.DataFrame:
    return df if project == ALL_PROJECTS else df[df["project"] == project]


def _sorted_buildings(values) -> list[str]:
    return sorted({v for v in values if v}, key=lambda v: (0, int(v)) if v.isdigit() else (1, v))


def render_filters(df: pd.DataFrame) -> dict:
    """Draw the filter bar; options are derived from the selected project so they never show stale choices."""
    project = G3_PROJECT  # the Gardenia 3 / Gardenia Town switch is the tab at the top of the page
    scope = _scope(df, project)
    k = project  # widget keys include the project so selections never clash when options change

    c1, c2, c3, c4 = st.columns(4)
    present = [s for s in STATUS_ORDER if s in set(scope["status_category"])]
    statuses = c1.multiselect("Status", present, key=f"cs_status_{k}", placeholder="All statuses")
    buildings = c2.multiselect("Building", _sorted_buildings(scope["building_no"]),
                               key=f"cs_building_{k}", placeholder="All buildings")
    floors = c3.multiselect("Floor", sorted({f for f in scope["floor"] if f}),
                            key=f"cs_floor_{k}", placeholder="All floors")
    years = sorted(scope["handover_date"].dropna().dt.year.unique())
    year = c4.selectbox("Handover year", ["All"] + [int(y) for y in years], key=f"cs_year_{k}")

    c5, c6 = st.columns([3, 1])
    search = c5.text_input("Search client name or notes", key=f"cs_search_{k}", placeholder="Type to search…")
    soon_only = c6.checkbox(f"Handover within {ENDING_SOON_DAYS} days only", key=f"cs_soon_{k}")
    return {"project": project, "statuses": statuses, "buildings": buildings, "floors": floors,
            "year": year, "search": search.strip().lower(), "soon_only": soon_only}


def filter_data(df: pd.DataFrame, f: dict, today: pd.Timestamp) -> pd.DataFrame:
    """Project and every other filter are applied here, before ANY metric is computed."""
    out = _scope(df, f["project"])
    if f["statuses"]:
        out = out[out["status_category"].isin(f["statuses"])]
    if f["buildings"]:
        out = out[out["building_no"].isin(f["buildings"])]
    if f["floors"]:
        out = out[out["floor"].isin(f["floors"])]
    if f["year"] != "All":
        out = out[out["handover_date"].dt.year == f["year"]]
    if f["soon_only"]:
        out = out[_ending_soon_mask(out["handover_date"], today)]
    if f["search"]:
        out = out[out["search_text"].str.contains(f["search"], regex=False, na=False)]
    return out


# =====================================================================
# KPIs
# =====================================================================
def calculate_kpis(df: pd.DataFrame, today: pd.Timestamp) -> dict:
    total = len(df)
    paid_off = int((df["remaining_installments"] == 0).sum())
    return {
        "total": total,
        "area": df["area_sqm"].sum(),
        "avg_installments": df["remaining_installments"].mean(),  # unknown values are skipped
        "buildings": len(df.loc[df["building_no"] != "", ["project", "building_no"]].drop_duplicates()),
        "completion": paid_off / total * 100 if total else 0.0,
        "ending_soon": int(_ending_soon_mask(df["handover_date"], today).sum()),
    }


def render_kpis(k: dict) -> None:
    """KPI row using st.metric so the cards pick up the shared theme styling."""
    cards = [
        ("Total Contracts", fmt_int(k["total"]), "All registered contracts"),
        ("Total Area (m²)", fmt_int(k["area"]), "Combined unit area"),
        ("Avg. Remaining Installments", fmt_dec(k["avg_installments"]),
         "Average across contracts with installment data"),
        ("Buildings", fmt_int(k["buildings"]), "Distinct buildings"),
        ("Fully Paid", fmt_int(k["completion"]) + "%", "Contracts with no remaining installments"),
        (f"Handover in {_digits(str(ENDING_SOON_DAYS))} Days", fmt_int(k["ending_soon"]),
         "Contracts with an upcoming handover date"),
    ]
    for start in (0, 3):  # two rows of three so labels and numbers are never truncated
        for col, (label, value, hint) in zip(st.columns(3), cards[start:start + 3]):
            col.metric(label, value, help=hint)


# =====================================================================
# CHARTS
# =====================================================================
def _style(fig, height: int = 300):
    fig.update_layout(
        height=height, margin=dict(l=8, r=8, t=8, b=8), paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)", font=dict(size=13, color=NAVY), legend=dict(title=None),
    )
    fig.update_xaxes(showgrid=False, title=None)
    fig.update_yaxes(gridcolor="#E4E8EA", title=None)
    return fig


def _chart_card(title: str, subtitle: str, fig=None, empty_msg: str = "Not enough data for this chart") -> None:
    with st.container(border=True):
        st.markdown(f"**{title}**")
        st.caption(subtitle)
        if fig is None:
            st.info(empty_msg)
        else:
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def _status_fig(df):
    counts = df["status_category"].value_counts().rename_axis("Status").reset_index(name="Contracts")
    fig = px.pie(counts, names="Status", values="Contracts", hole=0.62, color="Status",
                 color_discrete_map=STATUS_COLORS)
    fig.update_traces(textinfo="value+percent", sort=False)
    return _style(fig).update_layout(legend=dict(orientation="v", x=1.0))


def _building_fig(df):
    multi = df["project"].nunique() > 1
    label = (df["project"] + " · " + df["building_no"]) if multi else df["building_no"].replace("", "N/A")
    counts = df.assign(label=label).groupby(["label", "project"]).size().reset_index(name="Units")
    order = counts.groupby("label")["Units"].sum().sort_values(ascending=False).index.tolist()
    fig = px.bar(counts, x="label", y="Units", color="project" if multi else None, text="Units",
                 category_orders={"label": order}, color_discrete_sequence=[NAVY, ACCENT_GOLD],
                 labels={"label": "Building", "project": "Project"})
    fig.update_traces(textposition="outside", cliponaxis=False)
    if not multi:
        fig.update_traces(marker_color=NAVY)
    return _style(fig).update_xaxes(type="category")


def _bucket(n: float) -> str | None:
    if pd.isna(n):
        return None
    if n == 0:
        return "0 (fully paid)"
    return "1–6" if n <= 6 else "7–12" if n <= 12 else "13–24" if n <= 24 else "More than 24"


def _installment_fig(df):
    buckets = df["remaining_installments"].map(_bucket).dropna()
    if buckets.empty:
        return None
    counts = buckets.value_counts().reindex(BUCKET_ORDER, fill_value=0).rename_axis("Installments").reset_index(name="Contracts")
    fig = px.bar(counts, x="Installments", y="Contracts", text="Contracts")
    fig.update_traces(marker_color=NAVY_LIGHT, textposition="outside", cliponaxis=False)
    return _style(fig)


def _quarter_fig(df):
    dates = df["handover_date"].dropna()
    if dates.empty:
        return None
    counts = dates.dt.to_period("Q").value_counts().sort_index()
    data = pd.DataFrame({"Quarter": [f"{p.year} Q{p.quarter}" for p in counts.index], "Contracts": counts.values})
    fig = px.bar(data, x="Quarter", y="Contracts", text="Contracts")
    fig.update_traces(marker_color=ACCENT_TEAL, textposition="outside", cliponaxis=False)
    return _style(fig).update_xaxes(type="category")


def _top_installments_fig(df):
    top = df[df["remaining_installments"] > 0].nlargest(8, "remaining_installments")
    if top.empty:
        return None
    fig = px.bar(top, x="remaining_installments", y="unit_label", orientation="h",
                 text="remaining_installments", hover_data={"project": True, "unit_label": False},
                 labels={"remaining_installments": "Remaining installments", "project": "Project"})
    fig.update_traces(marker_color=DANGER, textposition="outside", cliponaxis=False)
    return _style(fig).update_yaxes(autorange="reversed", gridcolor="rgba(0,0,0,0)")


def render_charts(df: pd.DataFrame) -> None:
    c1, c2 = st.columns(2)
    with c1:
        _chart_card("Contract Status", "Based on the status recorded in the notes", _status_fig(df))
    with c2:
        _chart_card("Remaining Installments", "Contracts grouped by installments left", _installment_fig(df),
                    "No installment data for the selected contracts")
    _chart_card("Units per Building", "Distribution of contracts across buildings", _building_fig(df))
    c3, c4 = st.columns(2)
    with c3:
        _chart_card("Handovers by Quarter", "Contracts by handover quarter", _quarter_fig(df),
                    "No handover dates for the selected contracts")
    with c4:
        _chart_card("Top 8 Units by Remaining Installments", "Units with the most installments left",
                    _top_installments_fig(df), "No remaining installments for the selected contracts")


# =====================================================================
# TABLES & INSIGHTS
# =====================================================================
def _mini_table(df: pd.DataFrame, columns: dict, empty_msg: str) -> None:
    if df.empty:
        st.caption(empty_msg)
        return
    view = df[list(columns)].rename(columns=columns)
    st.dataframe(view, hide_index=True, use_container_width=True, height=250, column_config={
        "Handover": st.column_config.DateColumn(format="YYYY-MM-DD"),
        "Remaining": st.column_config.NumberColumn(format="%d"),
    })


def render_summary_tables(df: pd.DataFrame, today: pd.Timestamp) -> None:
    with st.container(border=True):
        st.markdown("**Upcoming Handovers**")
        st.caption("Next handover dates (past dates excluded)")
        upcoming = df[df["handover_date"] >= today].nsmallest(6, "handover_date")
        _mini_table(upcoming, {"client_name": "Client", "building_no": "Building", "unit_no": "Unit",
                               "handover_date": "Handover", "remaining_installments": "Remaining"},
                    "No upcoming handovers")
    with st.container(border=True):
        st.markdown("**Highest Remaining Installments**")
        st.caption("Top 6 contracts")
        top = df[df["remaining_installments"] > 0].nlargest(6, "remaining_installments")
        _mini_table(top, {"client_name": "Client", "building_no": "Building", "unit_no": "Unit",
                          "remaining_installments": "Remaining"}, "No remaining installments")


def render_insights(df: pd.DataFrame) -> None:
    b = df.loc[df["building_no"] != ""].groupby(["project", "building_no"]).size().sort_values(ascending=False)
    largest = df.loc[df["area_sqm"].idxmax()] if df["area_sqm"].notna().any() else None
    top_status = df["status_category"].value_counts()
    hi = df.loc[df["remaining_installments"].idxmax()] if df["remaining_installments"].notna().any() else None
    items = [
        ("Largest Building", f"Building {b.index[0][1]}" if len(b) else "—", f"{fmt_int(b.iloc[0])} units" if len(b) else ""),
        ("Largest Unit by Area", f"{fmt_int(largest['area_sqm'])} m²" if largest is not None else "—",
         largest["unit_label"] if largest is not None else ""),
        ("Most Common Status", top_status.index[0] if len(top_status) else "—",
         f"{fmt_int(top_status.iloc[0])} contracts" if len(top_status) else ""),
        ("Most Installments Left", fmt_int(hi["remaining_installments"]) if hi is not None else "—",
         hi["unit_label"] if hi is not None else ""),
    ]
    for start in (0, 2):
        for col, (label, value, sub) in zip(st.columns(2), items[start:start + 2]):
            col.metric(label, value)
            col.caption(sub)


def render_data_table(df: pd.DataFrame, today: pd.Timestamp) -> None:
    st.subheader("Contract Register")
    columns = {
        "project": "Project", "client_name": "Client", "building_no": "Building", "unit_no": "Unit",
        "floor": "Floor", "area_sqm": "Area (m²)", "remaining_installments": "Remaining Installments",
        "contract_date": "Contract Date", "handover_date": "Handover Date",
        "status_category": "Status", "notes": "Notes",
    }
    view = df.sort_values(["project", "building_no", "unit_no"]).reset_index(drop=True)
    soon = _ending_soon_mask(view["handover_date"], today)
    view = view[list(columns)].rename(columns=columns)
    styled = view.style.apply(
        lambda row: ["background-color: rgba(192,86,59,.10)" if soon[row.name] else ""] * len(row), axis=1)
    st.caption(f"{fmt_int(len(view))} contracts · highlighted rows hand over within {ENDING_SOON_DAYS} days")
    st.dataframe(styled, hide_index=True, use_container_width=True, height=460, column_config={
        "Contract Date": st.column_config.DateColumn(format="YYYY-MM-DD"),
        "Handover Date": st.column_config.DateColumn(format="YYYY-MM-DD"),
        "Area (m²)": st.column_config.NumberColumn(format="%d"),
        "Remaining Installments": st.column_config.NumberColumn(format="%d"),
    })
    st.download_button("⬇ Export CSV", view.to_csv(index=False).encode("utf-8-sig"),
                       file_name="contracts_filtered.csv", mime="text/csv")


# =====================================================================
# GARDENIA TOWN - delayed installments
# =====================================================================
TOWN_COLS = {  # Arabic header -> canonical name. The phone-number column is intentionally NOT loaded.
    "الاسم": "client_name", "المشروع": "sub_project", "العمارة": "building_no", "الوحدة": "unit_no",
    "المتأخرات": "arrears", "الشهور": "delay_period", "عدد الشهور": "months_delayed", "ملاحظات": "notes",
}
ARABIC_MONTHS = {
    "يناير": 1, "فبراير": 2, "مارس": 3, "ابريل": 4, "مايو": 5, "يونيو": 6,
    "يوليو": 7, "اغسطس": 8, "سبتمبر": 9, "اكتوبر": 10, "نوفمبر": 11, "ديسمبر": 12,
}
# First matching rule wins, so the most serious situations come first. Keyword based: review if notes change style.
FOLLOWUP_RULES = [
    ("Resale / termination", ("انذار فسخ", "ريسيل", "resale")),
    ("Contact problem", ("الرقم غلط", "no whatsapp")),
    ("Promised to pay", ("اكسبشن", "هيتم السداد", "هننزل الشيك", "بيدفع بشيك")),  # before "paid": "هيتم السداد" contains "تم السداد"
    ("Marked paid", ("paid", "تم السداد")),
    ("Message sent", ("whatsapp",)),
]
NO_NOTE = "No follow-up note"
FOLLOWUP_ORDER = ["Resale / termination", "Contact problem", "Promised to pay", "Message sent", "Marked paid", NO_NOTE, "Other"]
AMOUNT_BANDS = [(0, 50_000, "Up to 50K"), (50_000, 100_000, "50K–100K"), (100_000, 150_000, "100K–150K"),
                (150_000, 250_000, "150K–250K"), (250_000, float("inf"), "Over 250K")]
DUE_DAY_GROUPS = [(1, 5, "1–5"), (6, 10, "6–10"), (11, 15, "11–15"), (16, 20, "16–20"),
                  (21, 25, "21–25"), (26, 31, "26–31")]
DUE_DAY_UNSPECIFIED = "Not specified"
MONTH_BUCKETS = ["Not recorded (0)", "1 month", "2 months", "3+ months"]


def _norm_ar(text: str) -> str:
    return text.translate(str.maketrans("أإآ", "ااا")).lower()


def _parse_delay_period(label: str) -> tuple[list[int], float]:
    """'15/اكتوبر' -> ([10], 15); 'مايو - اغسطس' -> ([5, 8], nan). Month names only: the file has no year."""
    text = _norm_ar(label)
    months = sorted({num for name, num in ARABIC_MONTHS.items() if name in text})
    day = re.search(r"(\d{1,2})\s*/", text)
    return months, float(day.group(1)) if day else float("nan")


def classify_followup(note) -> str:
    text = _norm_ar(note) if isinstance(note, str) else ""
    if not text.strip():
        return NO_NOTE
    for status, keywords in FOLLOWUP_RULES:
        if any(k in text for k in keywords):
            return status
    return "Other"


def _months_bucket(n: float) -> str:
    if pd.isna(n) or n <= 0:
        return MONTH_BUCKETS[0]
    return MONTH_BUCKETS[1] if n == 1 else MONTH_BUCKETS[2] if n == 2 else MONTH_BUCKETS[3]


def _amount_band(x: float) -> str | None:
    if pd.isna(x):
        return None
    return next(label for lo, hi, label in AMOUNT_BANDS if lo <= x <= hi) if x > 0 else AMOUNT_BANDS[0][2]


def _due_day_group(day: float) -> str:
    if pd.isna(day):
        return DUE_DAY_UNSPECIFIED
    return next((label for lo, hi, label in DUE_DAY_GROUPS if lo <= day <= hi), DUE_DAY_UNSPECIFIED)


def fmt_money(x) -> str:
    if pd.isna(x):
        return "—"
    return f"{x / 1e6:,.2f}M" if abs(x) >= 1e6 else f"{x / 1e3:,.0f}K" if abs(x) >= 1e3 else f"{x:,.0f}"


def _find_town_file() -> Path | None:
    files = sorted({f for pattern in TOWN_GLOBS for f in DATA_DIR.glob(pattern)},
                   key=lambda f: f.stat().st_mtime, reverse=True)
    return files[0] if files else None


def _town_signature() -> tuple:
    path = _find_town_file()
    return (str(path), path.stat().st_mtime) if path else (None, None)


def clean_town_data(df: pd.DataFrame) -> pd.DataFrame:
    """Whitespace/type cleaning only. Amounts, months and notes are never altered."""
    df = df.copy()
    for col in ["client_name", "sub_project", "building_no", "unit_no", "notes", "delay_period"]:
        df[col] = df[col].map(_key_part)
    df["arrears"] = df["arrears"].map(parse_number)
    df["months_delayed"] = df["months_delayed"].map(parse_number)
    parsed = df["delay_period"].map(_parse_delay_period)
    df["oldest_month"] = parsed.map(lambda p: p[0][0] if p[0] else float("nan"))
    df["latest_month"] = parsed.map(lambda p: p[0][-1] if p[0] else float("nan"))
    df["due_day"] = parsed.map(lambda p: p[1])
    df["followup_status"] = df["notes"].map(classify_followup)
    df["months_bucket"] = df["months_delayed"].map(_months_bucket)
    df["amount_band"] = df["arrears"].map(_amount_band)
    df["unit_label"] = df["sub_project"] + " · Bldg " + df["building_no"] + " · Unit " + df["unit_no"]
    df["search_text"] = (df["client_name"] + " " + df["notes"]).str.lower()
    return df


@st.cache_data(show_spinner="Loading Gardenia Town data...")
def load_town_data(signature: tuple) -> tuple[pd.DataFrame, list[str]]:
    path = _find_town_file()
    if path is None:
        raise FileNotFoundError("Gardenia Town file not found")
    table = _select_renamed(_read_table(path, "المتأخرات"), TOWN_COLS)
    missing = [c for c in TOWN_COLS.values() if c not in table.columns]
    if missing:
        raise ValueError(f"{path.name} is missing columns: {', '.join(missing)}")
    table = table.dropna(subset=["client_name", "arrears"], how="all").reset_index(drop=True)
    df = clean_town_data(table)
    df["project"] = TOWN_PROJECT
    return df, town_quality_notes(df, path.name)


def town_quality_notes(df: pd.DataFrame, filename: str) -> list[str]:
    dup_units = int(df.duplicated(["sub_project", "building_no", "unit_no"], keep=False).sum())
    combined_units = int((~df["unit_no"].str.fullmatch(r"\d+")).sum())
    multi_unit_clients = int((df["client_name"].value_counts() > 1).sum())
    return [
        f"{len(df)} delayed accounts loaded from {filename}.",
        "Phone numbers are in the file but are not loaded or shown anywhere on the page.",
        f"{int((df['followup_status'] == 'Marked paid').sum())} accounts are marked 'paid' in the notes but still carry "
        f"an arrears amount — confirm whether they should be on the list.",
        f"{int((df['months_delayed'] == 0).sum())} accounts show 0 delayed months but still have arrears.",
        f"{combined_units} units are combined units written like '3+13' (kept as text).",
        f"{dup_units} rows share the same sub-project, building and unit with another row; "
        f"{multi_unit_clients} clients appear on more than one row. Nothing was removed.",
        f"{int(df['due_day'].isna().sum())} rows have no day in the delay-period column, and the file has no year, "
        "so days overdue cannot be calculated.",
        f"Follow-up status is read from the notes by keyword; {int((df['followup_status'] == NO_NOTE).sum())} "
        "accounts have no note at all.",
    ]


# ---- filtering / KPIs
def render_town_filters(df: pd.DataFrame) -> dict:
    c1, c2, c3, c4 = st.columns(4)
    sub = c1.multiselect("Sub-project", sorted(df["sub_project"].unique()), key="gt_sub", placeholder="All")
    bld = c2.multiselect("Building", _sorted_buildings(df["building_no"]), key="gt_bld", placeholder="All buildings")
    months = c3.multiselect("Months delayed", [b for b in MONTH_BUCKETS if b in set(df["months_bucket"])],
                            key="gt_months", placeholder="All")
    status = c4.multiselect("Follow-up status", [x for x in FOLLOWUP_ORDER if x in set(df["followup_status"])],
                            key="gt_status", placeholder="All")
    c5, c6 = st.columns([1, 2])
    lo, hi = float(df["arrears"].min()), float(df["arrears"].max())
    amount = c5.slider("Arrears amount", lo, hi, (lo, hi), step=1000.0, key="gt_amount") if lo < hi else (lo, hi)
    search = c6.text_input("Search client name or notes", key="gt_search", placeholder="Type to search…")
    return {"sub": sub, "bld": bld, "months": months, "status": status, "amount": amount,
            "bounds": (lo, hi), "search": search.strip().lower()}


def filter_town_data(df: pd.DataFrame, f: dict) -> pd.DataFrame:
    out = df
    if f["sub"]:
        out = out[out["sub_project"].isin(f["sub"])]
    if f["bld"]:
        out = out[out["building_no"].isin(f["bld"])]
    if f["months"]:
        out = out[out["months_bucket"].isin(f["months"])]
    if f["status"]:
        out = out[out["followup_status"].isin(f["status"])]
    if f["amount"] != f["bounds"]:
        out = out[out["arrears"].between(*f["amount"])]
    if f["search"]:
        out = out[out["search_text"].str.contains(f["search"], regex=False, na=False)]
    return out


def calculate_town_kpis(df: pd.DataFrame) -> dict:
    total = df["arrears"].sum()
    return {
        "accounts": len(df), "total": total, "avg": df["arrears"].mean(), "median": df["arrears"].median(),
        "multi": int((df["months_delayed"] >= 2).sum()),
        "top10_share": df["arrears"].nlargest(10).sum() / total * 100 if total else 0.0,
    }


def render_town_kpis(k: dict) -> None:
    cards = [
        ("Delayed Accounts", fmt_int(k["accounts"]), "Accounts with overdue installments"),
        (f"Total Arrears ({CURRENCY})", fmt_money(k["total"]), "Sum of all overdue amounts"),
        ("Average Arrears", fmt_money(k["avg"]), "Per delayed account"),
        ("Median Arrears", fmt_money(k["median"]), "Typical account: half owe less, half owe more"),
        ("Delayed 2+ Months", fmt_int(k["multi"]), "Accounts that missed more than one month"),
        ("Top 10 Share", fmt_int(k["top10_share"]) + "%", "Share of arrears held by the 10 largest balances"),
    ]
    for start in (0, 3):
        for col, (label, value, hint) in zip(st.columns(3), cards[start:start + 3]):
            col.metric(label, value, help=hint)


# ---- insights (generated from whatever is currently filtered)
def town_insights(df: pd.DataFrame, report_month: int | None = None) -> list[str]:
    """`report_month` = current month of the report, taken from the unfiltered data so filters do not change it."""
    total, n = df["arrears"].sum(), len(df)
    if n == 0 or total <= 0:
        return []
    out = []
    mean, median = df["arrears"].mean(), df["arrears"].median()
    line = (f"**{fmt_int(n)} delayed accounts** owe **{fmt_money(total)} {CURRENCY}** in total "
            f"(median {fmt_money(median)}, average {fmt_money(mean)}).")
    if mean > median * 1.15:
        line += " A few large balances pull the average above the typical account."
    out.append(line)

    big = df[df["arrears"] > 250_000]
    top10 = df["arrears"].nlargest(10).sum() / total * 100
    out.append(f"**Concentration:** the 10 largest balances hold {top10:.0f}% of arrears"
               + (f"; the {len(big)} accounts above 250K hold {big['arrears'].sum() / total * 100:.0f}%." if len(big) else "."))

    if df["sub_project"].nunique() > 1:
        g = df.groupby("sub_project")["arrears"].agg(["sum", "mean", "count"]).sort_values("sum", ascending=False)
        top, hi_avg = g.index[0], g["mean"].idxmax()
        text = (f"**{top}** has the largest arrears ({fmt_money(g['sum'].iloc[0])}, "
                f"{g['sum'].iloc[0] / total * 100:.0f}% of the total, {int(g['count'].iloc[0])} accounts)")
        text += (f", while **{hi_avg}** has the higher average balance ({fmt_money(g.loc[hi_avg, 'mean'])})."
                 if hi_avg != top else ", and also the higher average balance.")
        out.append(text)

    multi = df[df["months_delayed"] >= 2]
    one = int((df["months_delayed"] == 1).sum())
    if len(multi):
        out.append(f"**Duration:** {one / n * 100:.0f}% of accounts are late by a single month, but the "
                   f"{len(multi)} accounts delayed 2+ months hold {multi['arrears'].sum() / total * 100:.0f}% of arrears — "
                   f"they should be handled first.")
    elif one:
        out.append(f"**Duration:** {one / n * 100:.0f}% of accounts are late by a single month; none is delayed longer.")

    if report_month:
        carried = df[df["oldest_month"] < report_month]
        if len(carried):
            out.append(f"**{len(carried)} accounts** carry delays from before {calendar.month_name[report_month]}, "
                       f"worth {fmt_money(carried['arrears'].sum())} ({carried['arrears'].sum() / total * 100:.0f}% of arrears).")

    b = df.groupby(["sub_project", "building_no"])["arrears"].agg(["sum", "count"]).sort_values("sum", ascending=False)
    if len(b):
        (sub, bld), row = b.index[0], b.iloc[0]
        out.append(f"**Hotspot:** {sub} building {bld} has the highest arrears ({fmt_money(row['sum'])} across "
                   f"{int(row['count'])} accounts, {row['sum'] / total * 100:.0f}% of the total).")

    no_note = int((df["followup_status"] == NO_NOTE).sum())
    paid = int((df["followup_status"] == "Marked paid").sum())
    escal = int((df["followup_status"] == "Resale / termination").sum())
    follow = f"**Follow-up gaps:** {no_note / n * 100:.0f}% of accounts ({no_note}) have no follow-up note"
    follow += f"; {paid} marked 'paid' still carry arrears — verify them" if paid else ""
    follow += f"; {escal} are already at resale/termination stage." if escal else "."
    out.append(follow)
    return out


# ---- charts
def _town_bar(data, x, y, color, text_fmt=None, horizontal=False, labels=None):
    fig = px.bar(data, x=y if horizontal else x, y=x if horizontal else y, orientation="h" if horizontal else "v",
                 text=data["_text"] if "_text" in data else y, labels=labels or {})
    fig.update_traces(marker_color=color, textposition="outside", cliponaxis=False)
    fig = _style(fig)
    if horizontal:
        fig.update_yaxes(autorange="reversed", gridcolor="rgba(0,0,0,0)")
    else:
        fig.update_xaxes(type="category")
    return fig


def _t_subproject_fig(df):
    g = df.groupby("sub_project").agg(Arrears=("arrears", "sum"), Accounts=("arrears", "size")).reset_index()
    g["_text"] = g["Arrears"].map(fmt_money)
    return _town_bar(g, "sub_project", "Arrears", NAVY, labels={"sub_project": "Sub-project"})


def _t_months_fig(df):
    c = df["months_bucket"].value_counts().reindex(MONTH_BUCKETS, fill_value=0).rename_axis("Months delayed").reset_index(name="Accounts")
    return _town_bar(c, "Months delayed", "Accounts", ACCENT_GOLD)


def _t_band_fig(df):
    order = [b[2] for b in AMOUNT_BANDS]
    c = df["amount_band"].value_counts().reindex(order, fill_value=0).rename_axis("Arrears size").reset_index(name="Accounts")
    return _town_bar(c, "Arrears size", "Accounts", NAVY_LIGHT)


def _t_oldest_fig(df):
    m = df["oldest_month"].dropna().astype(int)
    if m.empty:
        return None
    c = m.value_counts().sort_index()
    data = pd.DataFrame({"Oldest delayed month": [calendar.month_abbr[i] for i in c.index], "Accounts": c.values})
    return _town_bar(data, "Oldest delayed month", "Accounts", ACCENT_TEAL)


def _t_buildings_fig(df):
    g = df.groupby(["sub_project", "building_no"])["arrears"].sum().nlargest(10).reset_index()
    g["Building"] = g["sub_project"] + " · " + g["building_no"]
    g["_text"] = g["arrears"].map(fmt_money)
    return _town_bar(g, "Building", "arrears", NAVY, labels={"arrears": "Arrears"})


def _t_top_accounts_fig(df):
    top = df.nlargest(10, "arrears").copy()
    top["_text"] = top["arrears"].map(fmt_money)
    return _town_bar(top, "unit_label", "arrears", DANGER, horizontal=True,
                     labels={"arrears": "Arrears", "unit_label": ""})


def _t_followup_fig(df):
    c = df["followup_status"].value_counts().reindex(FOLLOWUP_ORDER).dropna().rename_axis("Status").reset_index(name="Accounts")
    return _town_bar(c, "Status", "Accounts", ACCENT_GOLD, horizontal=True, labels={"Status": ""})


def _t_dueday_fig(df):
    order = [g[2] for g in DUE_DAY_GROUPS] + [DUE_DAY_UNSPECIFIED]
    c = df["due_day"].map(_due_day_group).value_counts().reindex(order, fill_value=0).rename_axis("Day of month").reset_index(name="Accounts")
    return _town_bar(c, "Day of month", "Accounts", NAVY_LIGHT)


def render_town_charts(df: pd.DataFrame) -> None:
    c1, c2 = st.columns(2)
    with c1:
        _chart_card("Arrears by Sub-project", f"Total overdue amount ({CURRENCY})", _t_subproject_fig(df))
    with c2:
        _chart_card("How Long Accounts Are Delayed", "Accounts by number of delayed months", _t_months_fig(df))
    c3, c4 = st.columns(2)
    with c3:
        _chart_card("Arrears Size", "Accounts by overdue amount", _t_band_fig(df))
    with c4:
        _chart_card("Oldest Delayed Month", "First month of delay recorded per account (file has no year)",
                    _t_oldest_fig(df), "No delay months to show")
    _chart_card("Top 10 Buildings by Arrears", "Where the overdue money is concentrated", _t_buildings_fig(df))
    c5, c6 = st.columns(2)
    with c5:
        _chart_card("Top 10 Accounts by Arrears", "Largest single balances (by unit)", _t_top_accounts_fig(df))
    with c6:
        _chart_card("Follow-up Status", "Read from the notes column", _t_followup_fig(df))
    c7, c8 = st.columns(2)
    with c7:
        _chart_card("Day Given in Delay Period", "Day of month written next to the delay month (when specified)",
                    _t_dueday_fig(df))
    with c8, st.container(border=True):
        st.markdown("**Escalation List**")
        st.caption("Accounts delayed 2 months or more, largest first")
        esc = df[df["months_delayed"] >= 2].sort_values("arrears", ascending=False)
        if esc.empty:
            st.caption("No accounts delayed 2+ months.")
        else:
            st.dataframe(esc[["client_name", "unit_label", "months_delayed", "arrears"]].rename(columns={
                "client_name": "Client", "unit_label": "Unit", "months_delayed": "Months", "arrears": "Arrears"}),
                hide_index=True, use_container_width=True, height=250,
                column_config={"Arrears": st.column_config.NumberColumn(format="%d"),
                               "Months": st.column_config.NumberColumn(format="%d")})


def render_town_table(df: pd.DataFrame) -> None:
    st.subheader("Delayed Accounts")
    columns = {"client_name": "Client", "sub_project": "Sub-project", "building_no": "Building", "unit_no": "Unit",
               "arrears": "Arrears", "months_delayed": "Months Delayed", "delay_period": "Delay Period",
               "followup_status": "Follow-up Status", "notes": "Notes"}
    view = df.sort_values("arrears", ascending=False)[list(columns)].rename(columns=columns)
    st.caption(f"{fmt_int(len(view))} accounts, sorted by arrears")
    st.dataframe(view, hide_index=True, use_container_width=True, height=460, column_config={
        "Arrears": st.column_config.NumberColumn(format="%d"),
        "Months Delayed": st.column_config.NumberColumn(format="%d")})
    st.download_button("⬇ Export CSV", view.to_csv(index=False).encode("utf-8-sig"),
                       file_name="gardenia_town_delayed.csv", mime="text/csv")


def render_town_tab() -> None:
    try:
        df, notes = load_town_data(_town_signature())
    except FileNotFoundError:
        st.error(f"Gardenia Town file not found. Put a file named like '{TOWN_GLOBS[0]}' in {DATA_DIR}.")
        return
    except ValueError as exc:
        st.error(f"Problem with the Gardenia Town file structure: {exc}")
        return
    data = filter_town_data(df, render_town_filters(df))
    if data.empty:
        st.info("No accounts match the selected filters.")
        return
    render_town_kpis(calculate_town_kpis(data))
    st.write("")
    with st.container(border=True):
        st.markdown("**Key Insights**")
        report_month = int(df["oldest_month"].max()) if df["oldest_month"].notna().any() else None
        for line in town_insights(data, report_month):
            st.markdown(f"- {line}")
    st.write("")
    render_town_charts(data)
    st.write("")
    render_town_table(data)
    with st.expander("Data quality"):
        for note in notes:
            st.write(f"- {note}")


# =====================================================================
# PAGE
# =====================================================================
# Page config, global CSS and the logo are handled once in app.py / theme.py.
def render_gardenia3_tab() -> None:
    try:
        df, notes = load_data(_source_signature())
    except FileNotFoundError:
        st.error(f"Data file not found. Expected {COMBINED_FILE} or the project files listed in PROJECT_SOURCES.")
        return
    except ValueError as exc:
        st.error(f"Problem with the data file structure: {exc}")
        return

    today = pd.Timestamp.today().normalize()
    filters = render_filters(df)
    data = filter_data(df, filters, today)
    if data.empty:
        st.info("No records match the selected filters.")
        return

    render_kpis(calculate_kpis(data, today))
    st.write("")
    render_charts(data)
    st.write("")
    render_summary_tables(data, today)
    st.write("")
    render_insights(data)
    st.write("")
    render_data_table(data, today)

    with st.expander("Data quality"):
        for note in notes:
            st.write(note)


def main() -> None:
    render_header(DEPARTMENT)
    tab_g3, tab_town = st.tabs([G3_PROJECT, TOWN_PROJECT])
    with tab_g3:
        render_gardenia3_tab()
    with tab_town:
        render_town_tab()


main()