"""
Production-ready Streamlit Attendance Dashboard
Supports:
  - Excel (.xlsx, .xls) attendance-register exports
  - CSV attendance files
  - Browser file-picker via Streamlit st.file_uploader
Designed to reproduce the structure and visual language of the supplied
First Year CSE attendance HTML dashboard.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Any, Optional

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# ---------------------------------------------------------------------------
# Page configuration
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Attendance Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ---------------------------------------------------------------------------
# Theme / CSS
# ---------------------------------------------------------------------------

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:opsz,wght@8..60,500;8..60,600;8..60,700&family=Inter:wght@400;500;600;700;800&display=swap');

:root{
    --navy:#16233d;
    --navy2:#1f3355;
    --cream:#f6f4ee;
    --card:#ffffff;
    --line:#e4dfd2;
    --gold:#b08d2b;
    --gold-soft:#f1e6c8;
    --brick:#a3402f;
    --brick-soft:#f6e0da;
    --olive:#4d7a55;
    --olive-soft:#e2ede2;
    --ink:#20242c;
    --ink-soft:#5b6472;
}

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background: var(--cream);
    color: var(--ink);
}

.block-container {
    max-width: 1280px;
    padding-top: 1.2rem;
    padding-bottom: 3rem;
}

.hero {
    background: linear-gradient(180deg, var(--navy) 0%, var(--navy2) 100%);
    color: #f4f1e8;
    padding: 26px 32px 22px;
    border-radius: 0 0 8px 8px;
    margin: -1.2rem -0.5rem 1.7rem -0.5rem;
    border-bottom: 6px solid rgba(176,141,43,.65);
}

.hero .eyebrow {
    color: #d4b65e;
    font-size: 11px;
    font-weight: 800;
    letter-spacing: .14em;
    text-transform: uppercase;
    margin-bottom: 6px;
}

.hero h1 {
    font-family: 'Source Serif 4', serif;
    font-size: 30px;
    font-weight: 600;
    margin: 0 0 8px 0;
}

.hero .meta {
    display: flex;
    flex-wrap: wrap;
    gap: 10px 24px;
    font-size: 12.5px;
    color: #cdd3e0;
}
.hero .meta b { color: #f4f1e8; }

.section-title {
    font-family: 'Source Serif 4', serif;
    font-size: 20px;
    font-weight: 600;
    color: var(--navy);
    border-bottom: 1px solid var(--line);
    padding-bottom: 8px;
    margin: 1.5rem 0 .9rem;
}

.section-note {
    color: var(--ink-soft);
    font-size: 12px;
    margin-top: -6px;
    margin-bottom: 10px;
}

.kpi {
    background: var(--card);
    border: 1px solid var(--line);
    border-radius: 8px;
    padding: 15px 16px 13px;
    min-height: 96px;
    box-shadow: 0 1px 1px rgba(22,35,61,.02);
}
.kpi .num {
    font-family: 'Source Serif 4', serif;
    font-size: 29px;
    line-height: 1.1;
    font-weight: 600;
}
.kpi .label {
    color: var(--ink-soft);
    font-size: 11.5px;
    margin-top: 6px;
}
.kpi.navy .num { color: var(--navy); }
.kpi.gold .num { color: var(--gold); }
.kpi.brick .num { color: var(--brick); }
.kpi.olive .num { color: var(--olive); }

.panel {
    background: var(--card);
    border: 1px solid var(--line);
    border-radius: 8px;
    padding: 15px 18px 10px;
}

.insight {
    background: #fff;
    border-left: 4px solid var(--gold);
    border-top: 1px solid var(--line);
    border-right: 1px solid var(--line);
    border-bottom: 1px solid var(--line);
    border-radius: 6px;
    padding: 11px 14px;
    margin-bottom: 8px;
    font-size: 12.5px;
    line-height: 1.5;
}

[data-testid="stDataFrame"] {
    border: 1px solid var(--line);
}

div[data-testid="stExpander"] {
    border: 1px solid var(--line);
    border-radius: 7px;
    background: #fff;
}

.stButton > button {
    border-radius: 6px;
}

.small-muted {
    color: var(--ink-soft);
    font-size: 11.5px;
}

footer { visibility: hidden; }

@media (max-width: 900px) {
    .hero { padding: 20px 18px; }
    .hero h1 { font-size: 25px; }
}
</style>
""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Constants / helpers
# ---------------------------------------------------------------------------

TH = "TH"
PR = "PR"

# Known subject abbreviations used by the supplied register format.
SUBJECT_MAP = {
    "M-I": ("Engineering Maths-I", TH),
    "EP": ("Engineering Physics", TH),
    "CP": ("Computer Programming", TH),
    "EM": ("Engineering Mechanics", TH),
    "EP-L": ("Physics Lab", PR),
    "CP-L": ("CP Lab", PR),
    "EM-L": ("Mechanics Lab", PR),
    "WP": ("Workshop Practice", PR),
    "IWT-L": ("IWT Lab", PR),
    "IWT": ("IWT (Theory)", TH),
    "PC": ("Prog. for Comp. (PC)", TH),
    "CC": ("Co-Curricular (CC)", PR),
    "PC-L": ("PC Lab", PR),
}


@dataclass
class AttendanceData:
    students: pd.DataFrame
    subjects: pd.DataFrame
    department: str = "Applied Science & Humanities"
    academic_year: str = ""
    semester: str = ""
    year: str = ""
    division: str = ""
    period: str = ""
    institution: str = "Mauli Group of Institution's College of Engineering and Technology, Shegaon"


def clean_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def title_case_name(value: Any) -> str:
    text = clean_text(value)
    if not text:
        return ""
    return " ".join(part[:1].upper() + part[1:].lower() for part in text.split())


def to_pct(value: Any) -> Optional[float]:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    text = clean_text(value)
    if not text or text in {"-", "—", "NA", "N/A", "nan"}:
        return None
    text = text.replace("%", "")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    number = float(match.group())
    # Decimal fractions are occasionally supplied by CSV exports.
    if 0 <= number <= 1:
        return number * 100
    return min(max(number, 0), 100)


def first_nonempty(values) -> str:
    for value in values:
        text = clean_text(value)
        if text:
            return text
    return ""


def status_of(value: Any) -> str:
    if value is None or pd.isna(value):
        return "unknown"
    value = float(value)
    if value < 50:
        return "critical"
    if value < 75:
        return "warning"
    return "good"


def status_label(value: Any) -> str:
    return {"critical": "Critical", "warning": "At Risk", "good": "Good", "unknown": "—"}[
        status_of(value)
    ]


def status_color(value: Any) -> str:
    return {
        "critical": "#a3402f",
        "warning": "#c79a1f",
        "good": "#4d7a55",
        "unknown": "#999999",
    }[status_of(value)]


def extract_metadata(raw: pd.DataFrame) -> dict[str, str]:
    text = " ".join(
        clean_text(x) for x in raw.iloc[:10, :].to_numpy().flatten()
        if clean_text(x)
    )

    def find(pattern: str) -> str:
        match = re.search(pattern, text, flags=re.I)
        return clean_text(match.group(1)) if match else ""

    return {
        "department": find(r"Department\s*:?\s*([^|]+?)(?=\s+Academic Year|$)") or
                      find(r"Department\s*:?\s*(.+?)(?=\s+Academic Year|$)"),
        "academic_year": find(r"Academic Year\s*:?\s*([0-9]{4}\s*-\s*[0-9]{2,4})"),
        "semester": find(r"Semester\s*:?\s*([A-Za-z0-9]+)"),
        "year": find(r"(?<!Academic )\bYear\s*:?\s*(.+?)(?=\s+Division|$)"),
        "division": find(r"Division\s*:?\s*([A-Za-z0-9_-]+)"),
    }


def extract_period(raw: pd.DataFrame) -> str:
    for i in range(min(12, len(raw))):
        row = " ".join(clean_text(x) for x in raw.iloc[i].tolist() if clean_text(x))
        match = re.search(
            r"from\s+(.+?)\s+to\s+(.+?)(?:$|\s{2,})",
            row,
            flags=re.I,
        )
        if match:
            return f"{match.group(1).strip()} – {match.group(2).strip()}"
    return ""


def identify_register_columns(raw: pd.DataFrame) -> list[dict[str, Any]]:
    """
    Parse the supplied Excel register's 3-column subject groups:
      Pre / [blank] / Per
    using row 6 as the subject header and row 8 as the subheader.
    """
    subject_headers = raw.iloc[6].tolist()
    subheaders = raw.iloc[8].tolist()

    specs: list[dict[str, Any]] = []

    # Known supplied register layout starts subject groups at column 3.
    current_header = ""
    for col in range(3, raw.shape[1]):
        header = clean_text(subject_headers[col])
        sub = clean_text(subheaders[col])

        # In the supplied Excel register the subject name is written only
        # in the first column of each 3-column group, while "Per" is in the
        # third column. Carry the most recent non-empty subject header forward.
        if header:
            current_header = header

        if not current_header or sub.lower() != "per":
            continue

        # Avoid Overall columns.
        if "overall" in header.lower():
            continue

        normalized = current_header.replace("\n", " ").replace("–", "-")
        normalized = re.sub(r"\s+", " ", normalized)

        # Header examples:
        # 1AL100BS - M-I- CP - TH
        # 1AL104ES - EP-L- CP - PR
        code_match = re.search(
            r"-\s*([A-Za-z0-9]+(?:-[A-Za-z0-9]+)?)\s*-\s*CP\s*-\s*(TH|PR)",
            normalized,
            flags=re.I,
        )
        if not code_match:
            # More tolerant fallback.
            code_match = re.search(
                r"-\s*([A-Za-z0-9]+(?:-[A-Za-z0-9]+)?)\s*-.*?\b(TH|PR)\b",
                normalized,
                flags=re.I,
            )
        if not code_match:
            continue

        code = code_match.group(1).upper()
        typ = code_match.group(2).upper()
        subject_name = SUBJECT_MAP.get(code, (code, typ))[0]

        specs.append(
            {
                "percent_col": col,
                "code": code,
                "name": subject_name,
                "type": typ,
            }
        )

    # De-duplicate subject columns while preserving order.
    seen = set()
    unique = []
    for spec in specs:
        key = (spec["name"], spec["type"])
        if key not in seen:
            seen.add(key)
            unique.append(spec)
    return unique


def locate_student_rows(raw: pd.DataFrame) -> pd.DataFrame:
    """Return rows that have a numeric-looking roll number."""
    result = raw.copy()
    roll = result.iloc[:, 1].map(clean_text)
    mask = roll.str.fullmatch(r"\d+")
    return result.loc[mask].copy()


def parse_excel_register(raw: pd.DataFrame) -> AttendanceData:
    if raw.shape[1] < 5 or raw.shape[0] < 10:
        raise ValueError(
            "The Excel file does not look like an attendance register. "
            "At least 10 rows and 5 columns were expected."
        )

    metadata = extract_metadata(raw)
    period = extract_period(raw)
    specs = identify_register_columns(raw)
    if not specs:
        raise ValueError(
            "Could not identify subject attendance columns in the Excel register. "
            "The file may use a different register layout."
        )

    rows = locate_student_rows(raw)
    if rows.empty:
        raise ValueError("No student rows with numeric Roll No. were found.")

    records = []
    for _, row in rows.iterrows():
        roll = clean_text(row.iloc[1])
        name = clean_text(row.iloc[2])
        if not name:
            continue

        subject_values = {}
        for spec in specs:
            subject_values[spec["name"]] = to_pct(row.iloc[spec["percent_col"]])

        # Use source-provided overall values when available; otherwise calculate.
        overall_th = to_pct(row.iloc[42]) if raw.shape[1] > 42 else None
        overall_pr = to_pct(row.iloc[43]) if raw.shape[1] > 43 else None
        overall = to_pct(row.iloc[44]) if raw.shape[1] > 44 else None

        th_values = [
            subject_values[s["name"]]
            for s in specs
            if s["type"] == TH and subject_values[s["name"]] is not None
        ]
        pr_values = [
            subject_values[s["name"]]
            for s in specs
            if s["type"] == PR and subject_values[s["name"]] is not None
        ]

        if overall_th is None and th_values:
            overall_th = float(np.mean(th_values))
        if overall_pr is None and pr_values:
            overall_pr = float(np.mean(pr_values))
        if overall is None:
            combined = [x for x in [overall_th, overall_pr] if x is not None]
            overall = float(np.mean(combined)) if combined else None

        records.append(
            {
                "roll": roll,
                "name": name,
                "overall_th": overall_th,
                "overall_pr": overall_pr,
                "overall": overall,
                **subject_values,
            }
        )

    students = pd.DataFrame(records)
    if students.empty:
        raise ValueError("No valid student attendance records were extracted.")

    subject_rows = []
    faculty_row_index = 76 if len(raw) > 76 else None
    average_row_index = 75 if len(raw) > 75 else None

    for spec in specs:
        values = pd.to_numeric(students[spec["name"]], errors="coerce")
        avg = float(values.mean()) if values.notna().any() else None

        # The supplied register contains an official Average row and faculty row.
        if average_row_index is not None:
            source_avg = to_pct(raw.iloc[average_row_index, max(0, spec["percent_col"] - 2)])
            if source_avg is not None:
                avg = source_avg

        faculty = ""
        if faculty_row_index is not None:
            faculty = clean_text(raw.iloc[faculty_row_index, max(0, spec["percent_col"] - 2)])

        subject_rows.append(
            {
                "name": spec["name"],
                "type": spec["type"],
                "faculty": faculty,
                "class_avg": avg,
            }
        )

    subjects = pd.DataFrame(subject_rows)

    return AttendanceData(
        students=students,
        subjects=subjects,
        department=metadata.get("department") or "Applied Science & Humanities",
        academic_year=metadata.get("academic_year", ""),
        semester=metadata.get("semester", ""),
        year=metadata.get("year", ""),
        division=metadata.get("division", ""),
        period=period,
    )


def normalize_csv_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Convert common CSV layouts to the internal schema.

    Accepted examples:
      Roll No, Name, Engineering Maths-I, Engineering Physics, ...
      roll, name, Maths TH, Physics TH, Physics Lab PR, ...
      Student Name / Roll Number variants.
    """
    work = df.copy()
    work.columns = [clean_text(c) or f"Column_{i+1}" for i, c in enumerate(work.columns)]

    lower_map = {c.lower().strip(): c for c in work.columns}

    def find_col(candidates: list[str]) -> Optional[str]:
        for candidate in candidates:
            if candidate in lower_map:
                return lower_map[candidate]
        for c in work.columns:
            lc = c.lower().strip()
            if any(candidate in lc for candidate in candidates):
                return c
        return None

    roll_col = find_col(["roll no", "roll number", "roll", "rollno", "sr no", "sr.no"])
    name_col = find_col(["name of the student", "student name", "name", "student"])

    if roll_col is None or name_col is None:
        raise ValueError(
            "CSV must contain identifiable Roll No and Student Name columns. "
            "For example: 'Roll No' and 'Name'."
        )

    out = pd.DataFrame(
        {
            "roll": work[roll_col].map(clean_text),
            "name": work[name_col].map(clean_text),
        }
    )

    overall_th_col = find_col(["overall th", "theory %", "overall theory"])
    overall_pr_col = find_col(["overall pr", "practical %", "overall practical"])
    overall_col = find_col(["overall att", "overall %", "overall attendance", "overall"])

    if overall_th_col:
        out["overall_th"] = work[overall_th_col].map(to_pct)
    if overall_pr_col:
        out["overall_pr"] = work[overall_pr_col].map(to_pct)
    if overall_col:
        out["overall"] = work[overall_col].map(to_pct)

    # Identify subject columns from remaining columns.
    reserved = {roll_col, name_col, overall_th_col, overall_pr_col, overall_col}
    subjects = []

    for col in work.columns:
        if col in reserved:
            continue

        lc = col.lower()
        if not re.search(r"(th|pr|theory|practical|lab|lecture)", lc):
            continue

        typ = PR if re.search(r"\b(pr|practical|lab)\b", lc) else TH

        # Strip common type labels for a cleaner subject display.
        subject_name = re.sub(
            r"[\s_\-]*(?:\(|\[)?(?:TH|PR|Theory|Practical|Lab)(?:\)|\])?[\s_\-]*$",
            "",
            col,
            flags=re.I,
        ).strip(" -_")
        if not subject_name:
            subject_name = col

        out[subject_name] = work[col].map(to_pct)
        subjects.append((subject_name, typ))

    # Calculate missing overall values.
    if "overall_th" not in out.columns:
        th_cols = [name for name, typ in subjects if typ == TH and name in out.columns]
        if th_cols:
            out["overall_th"] = out[th_cols].mean(axis=1, skipna=True)
    if "overall_pr" not in out.columns:
        pr_cols = [name for name, typ in subjects if typ == PR and name in out.columns]
        if pr_cols:
            out["overall_pr"] = out[pr_cols].mean(axis=1, skipna=True)
    if "overall" not in out.columns:
        available = [c for c in ["overall_th", "overall_pr"] if c in out.columns]
        if available:
            out["overall"] = out[available].mean(axis=1, skipna=True)

    if "overall" not in out.columns:
        raise ValueError(
            "Could not identify or calculate an Overall Attendance column in the CSV."
        )

    # Remove empty/non-student rows.
    out = out[out["name"].str.len() > 0].copy()
    out = out[out["roll"].str.len() > 0].copy()

    subject_rows = []
    for name, typ in subjects:
        if name in out.columns:
            avg = pd.to_numeric(out[name], errors="coerce").mean()
            subject_rows.append(
                {"name": name, "type": typ, "faculty": "", "class_avg": avg}
            )

    return AttendanceData(
        students=out,
        subjects=pd.DataFrame(subject_rows),
        department="",
        academic_year="",
        semester="",
        year="",
        division="",
        period="",
    )


def parse_uploaded_file(uploaded_file) -> AttendanceData:
    suffix = uploaded_file.name.lower().rsplit(".", 1)[-1]

    if suffix == "csv":
        # utf-8-sig handles common Excel-generated CSVs.
        raw = pd.read_csv(uploaded_file, encoding="utf-8-sig")
        return normalize_csv_columns(raw)

    if suffix in {"xlsx", "xls"}:
        data = uploaded_file.getvalue()
        engine = "openpyxl" if suffix == "xlsx" else None
        raw = pd.read_excel(io.BytesIO(data), header=None, engine=engine)
        return parse_excel_register(raw)

    raise ValueError("Unsupported file type. Please upload CSV, XLSX, or XLS.")


# ---------------------------------------------------------------------------
# Visuals
# ---------------------------------------------------------------------------

def make_distribution_chart(students: pd.DataFrame) -> go.Figure:
    vals = pd.to_numeric(students["overall"], errors="coerce").dropna()
    labels = ["Below 50%", "50–75%", "75–85%", "85–95%", "95–100%"]
    counts = [
        int((vals < 50).sum()),
        int(((vals >= 50) & (vals < 75)).sum()),
        int(((vals >= 75) & (vals < 85)).sum()),
        int(((vals >= 85) & (vals < 95)).sum()),
        int((vals >= 95).sum()),
    ]
    colors = ["#a3402f", "#c79a1f", "#8fae52", "#4d7a55", "#2f6b47"]

    fig = go.Figure(
        go.Bar(
            x=labels,
            y=counts,
            marker_color=colors,
            marker_line_width=0,
            text=counts,
            textposition="outside",
            hovertemplate="%{x}<br>%{y} students<extra></extra>",
        )
    )
    fig.update_layout(
        height=300,
        margin=dict(l=20, r=10, t=10, b=40),
        plot_bgcolor="white",
        paper_bgcolor="white",
        showlegend=False,
        yaxis=dict(dtick=1, gridcolor="#eee6d4", title=None),
        xaxis=dict(showgrid=False, title=None),
    )
    return fig


def make_subject_chart(subjects: pd.DataFrame) -> go.Figure:
    sub = subjects.dropna(subset=["class_avg"]).copy()
    sub = sub.sort_values("class_avg")

    colors = ["#16233d" if t == TH else "#b08d2b" for t in sub["type"]]
    fig = go.Figure(
        go.Bar(
            x=sub["class_avg"],
            y=sub["name"],
            orientation="h",
            marker_color=colors,
            marker_line_width=0,
            text=[f"{v:.1f}%" for v in sub["class_avg"]],
            textposition="outside",
            hovertemplate="%{y}<br>%{x:.2f}% average<extra></extra>",
        )
    )
    fig.update_layout(
        height=max(300, 28 * len(sub) + 80),
        margin=dict(l=10, r=55, t=10, b=30),
        plot_bgcolor="white",
        paper_bgcolor="white",
        showlegend=False,
        xaxis=dict(range=[0, 105], gridcolor="#eee6d4", title=None),
        yaxis=dict(showgrid=False, title=None, tickfont=dict(size=10)),
    )
    return fig


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------

def main() -> None:
    st.sidebar.markdown("## Attendance Dashboard")
    st.sidebar.caption(
        "Upload an attendance register exported as Excel or CSV. "
        "The dashboard is generated dynamically from the uploaded file."
    )

    uploaded = st.sidebar.file_uploader(
        "Upload CSV / Excel file",
        type=["csv", "xlsx", "xls"],
        help="The supplied Excel register format is supported directly.",
    )

    if uploaded is None:
        st.markdown(
            """
            <div class="hero">
              <div class="eyebrow">Attendance Dashboard</div>
              <h1>Upload an Attendance Register</h1>
              <div class="meta">
                <span>Supported: CSV, XLSX, XLS</span>
                <span>Dashboard generated from uploaded data</span>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.info(
            "Use the file picker in the left sidebar to select your attendance CSV or Excel file."
        )
        st.markdown(
            """
            **The dashboard includes**
            - Class snapshot KPIs
            - Attendance-band distribution
            - Subject-wise theory/practical averages
            - Top 5 and bottom 5 students
            - Searchable/filterable student register
            - Click-to-expand student subject details
            - Faculty & subject summary
            - Automatic data-quality validation
            """
        )
        return

    try:
        with st.spinner("Reading attendance register and preparing dashboard..."):
            data = parse_uploaded_file(uploaded)
    except Exception as exc:
        st.error("The uploaded file could not be processed.")
        st.exception(exc)
        st.stop()

    students = data.students.copy()
    subjects = data.subjects.copy()

    # Numeric cleanup.
    for col in ["overall_th", "overall_pr", "overall"]:
        if col in students:
            students[col] = pd.to_numeric(students[col], errors="coerce")

    students = students[students["overall"].notna()].copy()
    if students.empty:
        st.error("No valid overall attendance values were found.")
        st.stop()

    avg = float(students["overall"].mean())
    critical = int((students["overall"] < 50).sum())
    warning = int(((students["overall"] >= 50) & (students["overall"] < 75)).sum())
    good = int((students["overall"] >= 75).sum())
    top_student = students.sort_values("overall", ascending=False).iloc[0]
    bottom_student = students.sort_values("overall", ascending=True).iloc[0]

    # Header.
    year_text = data.year or "First Year"
    division_text = data.division or "—"
    st.markdown(
        f"""
        <div class="hero">
          <div class="eyebrow">
            Attendance Dashboard · {data.institution}
          </div>
          <h1>{year_text} — Division {division_text}</h1>
          <div class="meta">
            <span><b>Department:</b> {data.department or "—"}</span>
            <span><b>Academic Year:</b> {data.academic_year or "—"}</span>
            <span><b>Semester:</b> {data.semester or "—"}</span>
            <span><b>Period:</b> {data.period or "—"}</span>
            <span><b>Strength:</b> {len(students)} students</span>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # KPIs.
    st.markdown('<div class="section-title">Class Snapshot</div>', unsafe_allow_html=True)
    kcols = st.columns(5)
    kpi_data = [
        (str(len(students)), "Total Students Enrolled", "navy"),
        (f"{avg:.1f}%", "Class Average Attendance", "gold"),
        (str(critical + warning), "Students Below 75% (At Risk)", "brick"),
        (str(good), "Students at 75% or Above", "olive"),
        (f"{top_student.overall:.1f}%", f"Highest — {title_case_name(top_student['name'])}", "gold"),
    ]
    for col, (num, label, accent) in zip(kcols, kpi_data):
        with col:
            st.markdown(
                f'<div class="kpi {accent}"><div class="num">{num}</div>'
                f'<div class="label">{label}</div></div>',
                unsafe_allow_html=True,
            )

    # Attendance patterns.
    st.markdown(
        '<div class="section-title">Attendance Patterns <span style="font-family:Inter;font-size:11px;color:#b08d2b;">OVERALL %</span></div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="section-note">Distribution and subject-wise comparison</div>',
        unsafe_allow_html=True,
    )
    c1, c2 = st.columns([1.15, 1])
    with c1:
        st.markdown(
            '<div class="panel"><b>Distribution of Students by Attendance Band</b>'
            '<div class="small-muted">How students spread across attendance ranges</div></div>',
            unsafe_allow_html=True,
        )
        st.plotly_chart(make_distribution_chart(students), use_container_width=True, config={"displayModeBar": False})
    with c2:
        st.markdown(
            '<div class="panel"><b>Subject-wise Average Attendance</b>'
            '<div class="small-muted">Theory (TH) vs Practical (PR) class average per subject</div></div>',
            unsafe_allow_html=True,
        )
        if not subjects.empty:
            st.plotly_chart(make_subject_chart(subjects), use_container_width=True, config={"displayModeBar": False})
            st.caption("■ Theory   ■ Practical")
        else:
            st.info("Subject-level attendance columns were not available.")

    # Insights.
    st.markdown('<div class="section-title">Key Insights</div>', unsafe_allow_html=True)
    insights = []

    if critical:
        insights.append(
            f"**{critical} student(s)** are in the critical zone below 50% overall attendance."
        )
    if warning:
        insights.append(
            f"**{warning} student(s)** are in the at-risk band of 50–75% overall attendance."
        )
    if good:
        insights.append(
            f"**{good} student(s)** have overall attendance of 75% or above."
        )

    if not subjects.empty:
        low_subject = subjects.dropna(subset=["class_avg"]).sort_values("class_avg").iloc[0]
        high_subject = subjects.dropna(subset=["class_avg"]).sort_values("class_avg", ascending=False).iloc[0]
        insights.append(
            f"Lowest subject average: **{low_subject['name']} ({low_subject['class_avg']:.1f}%)**; "
            f"highest: **{high_subject['name']} ({high_subject['class_avg']:.1f}%)**."
        )

    if critical + warning:
        risk_pct = 100 * (critical + warning) / len(students)
        insights.append(
            f"**{risk_pct:.1f}%** of the class is below the 75% attendance threshold and may require attention."
        )

    if bottom_student["overall"] < 50:
        insights.append(
            f"Lowest overall attendance is **{title_case_name(bottom_student['name'])} "
            f"(Roll {bottom_student['roll']}) at {bottom_student['overall']:.1f}%**."
        )

    for text in insights:
        st.markdown(f'<div class="insight">{text}</div>', unsafe_allow_html=True)

    # Extremes.
    st.markdown('<div class="section-title">Extremes</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-note">Best & most at-risk attendance</div>',
        unsafe_allow_html=True,
    )
    e1, e2 = st.columns(2)
    sorted_students = students.sort_values("overall", ascending=False)
    with e1:
        st.markdown("#### Top 5 — Highest Overall Attendance")
        top5 = sorted_students.head(5)[["roll", "name", "overall"]].copy()
        top5["name"] = top5["name"].map(title_case_name)
        top5["overall"] = top5["overall"].map(lambda x: f"{x:.1f}%")
        st.dataframe(top5.rename(columns={"roll": "Roll No", "name": "Name", "overall": "Overall"}),
                     hide_index=True, use_container_width=True)
    with e2:
        st.markdown("#### Bottom 5 — Lowest Overall Attendance")
        bottom5 = sorted_students.tail(5).sort_values("overall")[["roll", "name", "overall"]].copy()
        bottom5["name"] = bottom5["name"].map(title_case_name)
        bottom5["overall"] = bottom5["overall"].map(lambda x: f"{x:.1f}%")
        st.dataframe(bottom5.rename(columns={"roll": "Roll No", "name": "Name", "overall": "Overall"}),
                     hide_index=True, use_container_width=True)

    # Student register.
    st.markdown('<div class="section-title">Student Register</div>', unsafe_allow_html=True)
    left, right = st.columns([1.5, 1])
    with left:
        search = st.text_input("Search by name or roll no.", placeholder="e.g. 24 or RUTUJA")
    with right:
        filter_label = st.selectbox(
            "Attendance filter",
            ["All", "Critical <50%", "At Risk 50–75%", "Good ≥75%"],
        )

    filtered = students.copy()
    if search.strip():
        q = search.strip().lower()
        filtered = filtered[
            filtered["name"].str.lower().str.contains(q, na=False)
            | filtered["roll"].astype(str).str.lower().str.contains(q, na=False)
        ]

    if filter_label != "All":
        wanted = {
            "Critical <50%": "critical",
            "At Risk 50–75%": "warning",
            "Good ≥75%": "good",
        }[filter_label]
        filtered = filtered[filtered["overall"].map(status_of) == wanted]

    sort_col = st.selectbox(
        "Sort register by",
        ["Roll No", "Name", "Overall Attendance", "Theory Attendance", "Practical Attendance"],
        horizontal=True,
    )
    sort_map = {
        "Roll No": "roll",
        "Name": "name",
        "Overall Attendance": "overall",
        "Theory Attendance": "overall_th",
        "Practical Attendance": "overall_pr",
    }
    sort_key = sort_map[sort_col]
    if sort_key == "roll":
        filtered["_roll_num"] = pd.to_numeric(filtered["roll"], errors="coerce")
        filtered = filtered.sort_values("_roll_num")
    else:
        filtered = filtered.sort_values(sort_key, na_position="last")

    st.caption(f"Showing {len(filtered)} of {len(students)} students")

    display = filtered[["roll", "name", "overall_th", "overall_pr", "overall"]].copy()
    display["name"] = display["name"].map(title_case_name)
    display["Theory %"] = display["overall_th"].map(lambda x: f"{x:.1f}%" if pd.notna(x) else "—")
    display["Practical %"] = display["overall_pr"].map(lambda x: f"{x:.1f}%" if pd.notna(x) else "—")
    display["Overall %"] = display["overall"].map(lambda x: f"{x:.1f}%" if pd.notna(x) else "—")
    display["Status"] = display["overall"].map(status_label)

    display = display.rename(columns={"roll": "Roll No", "name": "Name"})
    display = display[["Roll No", "Name", "Theory %", "Practical %", "Overall %", "Status"]]

    st.dataframe(
        display,
        hide_index=True,
        use_container_width=True,
        height=500,
    )

    # Expandable individual student details.
    if len(filtered) > 0:
        st.markdown("#### Student Detail")
        selected_roll = st.selectbox(
            "Select a student to view subject-wise attendance",
            filtered["roll"].astype(str).tolist(),
            format_func=lambda r: title_case_name(
                students.loc[students["roll"].astype(str) == str(r), "name"].iloc[0]
            ) + f" (Roll {r})",
        )
        student = students[students["roll"].astype(str) == str(selected_roll)].iloc[0]

        subject_details = []
        for _, subject in subjects.iterrows():
            subject_name = subject["name"]
            if subject_name in students.columns:
                value = student.get(subject_name)
                subject_details.append(
                    {
                        "Subject": subject_name,
                        "Type": subject["type"],
                        "Attendance": f"{value:.1f}%" if pd.notna(value) else "—",
                        "Status": status_label(value),
                    }
                )

        if subject_details:
            st.dataframe(
                pd.DataFrame(subject_details),
                hide_index=True,
                use_container_width=True,
            )

    # Faculty / subject summary.
    st.markdown('<div class="section-title">Faculty & Subject Summary</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-note">Lecture-wise class averages, as recorded in the register</div>',
        unsafe_allow_html=True,
    )
    if not subjects.empty:
        faculty_display = subjects.copy()
        faculty_display["class_avg"] = faculty_display["class_avg"].map(
            lambda x: f"{x:.2f}%" if pd.notna(x) else "—"
        )
        faculty_display = faculty_display.rename(
            columns={
                "name": "Subject",
                "type": "Type",
                "faculty": "Faculty",
                "class_avg": "Class Average",
            }
        )
        st.dataframe(
            faculty_display[["Subject", "Type", "Faculty", "Class Average"]],
            hide_index=True,
            use_container_width=True,
        )

    st.markdown(
        '<div class="small-muted" style="text-align:center;margin-top:28px;">'
        'Generated from the uploaded attendance register. Figures reflect data as recorded in the source file.'
        '</div>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
