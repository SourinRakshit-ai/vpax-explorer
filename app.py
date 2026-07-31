"""
VPAX Semantic Model Explorer
============================
A small web app that documents the semantic model behind a Power BI report
from its .vpax metadata export (produced by DAX Studio or Tabular Editor).

It surfaces:
  * the report screens/pages in the model,
  * an ER-style relationship diagram per screen, laid out as a Star,
    Snowflake or Galaxy schema with crow's-foot cardinality and filter
    direction,
  * tables, columns, measures (with a best-practice DAX rewrite),
    calculated columns, relationships and the SQL behind each Power Query.

Everything is derived dynamically from the uploaded file - no table, screen
or column names are hard-coded - so it works with any .vpax export.

Made by Sourin.

Run with:
    streamlit run app.py
"""

from __future__ import annotations

import importlib
import subprocess
import sys
from pathlib import Path

# --------------------------------------------------------------------------
# Auto-install missing dependencies from requirements.txt
# --------------------------------------------------------------------------
# Lets the app be started with a plain `python app.py` / double-click on a
# machine that has never had its packages installed, instead of failing with
# "ModuleNotFoundError". It only installs what's actually missing, so it's a
# no-op (fast) once the environment is already set up. This has no effect on
# Streamlit Community Cloud, which installs requirements.txt itself before
# the app ever runs.
_REQUIREMENTS_FILE = Path(__file__).resolve().parent / "requirements.txt"

# A package's pip name and its importable module name aren't always the
# same - list the exceptions here.
_IMPORT_NAME_OVERRIDES = {
    "streamlit": "streamlit",
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "xlsxwriter": "xlsxwriter",
}


def _ensure_requirements_installed() -> None:
    if not _REQUIREMENTS_FILE.exists():
        return

    packages = []
    for line in _REQUIREMENTS_FILE.read_text(encoding="utf-8").splitlines():
        line = line.split("#", 1)[0].strip()
        if line:
            packages.append(line)
    if not packages:
        return

    def _pkg_name(requirement: str) -> str:
        # Strip version pins like "pandas>=2.0" / "pandas==2.1.0" down to "pandas".
        for sep in ("==", ">=", "<=", "~=", ">", "<", "!="):
            if sep in requirement:
                return requirement.split(sep, 1)[0].strip()
        return requirement.strip()

    missing = []
    for requirement in packages:
        pkg = _pkg_name(requirement)
        import_name = _IMPORT_NAME_OVERRIDES.get(pkg, pkg.replace("-", "_"))
        try:
            importlib.import_module(import_name)
        except ImportError:
            missing.append(requirement)

    if not missing:
        return

    print(f"[VPAX Explorer] Installing missing packages: {', '.join(missing)} ...", file=sys.stderr)
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--quiet", "--disable-pip-version-check", *missing]
        )
        importlib.invalidate_caches()
    except Exception as exc:  # noqa: BLE001
        print(
            f"[VPAX Explorer] Automatic install failed ({exc}). "
            f"Please run manually:  pip install -r requirements.txt",
            file=sys.stderr,
        )


_ensure_requirements_installed()

import html
import io
import json
import math
import re
import traceback
import zipfile
from collections import deque
from contextlib import contextmanager
from functools import lru_cache
from typing import Any, Dict, List, Optional, Set, Tuple

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# --------------------------------------------------------------------------
# Guard against being launched the wrong way
# --------------------------------------------------------------------------
# This file is a Streamlit *script*, not a plain Python program - it has to
# be handed to the Streamlit server (`streamlit run app.py`), which re-runs
# it top-to-bottom on every interaction. Launched as `python app.py` instead,
# there's no server and no browser tab, and Streamlit's own st.stop()/widget
# calls silently no-op instead of halting the script - so execution falls
# through every `st.stop()` and crashes later with a confusing NameError.
# Catch that case immediately with a clear, actionable message.
if not st.runtime.exists():
    print(
        "\n"
        "This is a Streamlit app - it can't be run with a plain `python` command.\n"
        "Start it with:\n\n"
        f"    streamlit run \"{__file__}\"\n\n"
        "That launches a local web server and opens the app in your browser.\n",
        file=sys.stderr,
    )
    sys.exit(1)

st.set_page_config(
    page_title="VPAX Semantic Model Explorer",
    page_icon="🧩",
    layout="wide",
    initial_sidebar_state="expanded",
)

MAX_DIAGRAM_COLUMNS = 12      # columns drawn inside one table box
MAX_BRIDGE_TABLES = 60        # cap on the O(n^2) join-path search
MAX_DIAGRAM_TABLES = 40       # past this a diagram is unreadable and slow to render
AUTHOR = "Made by Sourin"


# ==========================================================================
# Styling
# ==========================================================================

def inject_css() -> None:
    st.markdown(
        """
        <style>
          .block-container { padding-top: 2.2rem; padding-bottom: 1rem; max-width: 1500px; }

          .app-hero {
            background: linear-gradient(120deg, #1f3a5f 0%, #2d6ca8 55%, #3f9ad1 100%);
            border-radius: 14px; padding: 1.4rem 1.8rem; margin-bottom: 1.2rem;
            color: #ffffff; box-shadow: 0 6px 20px rgba(31,58,95,.22);
          }
          .app-hero h1 { margin: 0; font-size: 1.75rem; font-weight: 700; letter-spacing:-.4px; color:#fff; }
          .app-hero p  { margin: .35rem 0 0; opacity: .92; font-size: .95rem; }
          .app-hero .badge {
            display:inline-block; margin-top:.7rem; padding:.2rem .7rem; border-radius:999px;
            background: rgba(255,255,255,.18); font-size:.78rem; letter-spacing:.3px;
          }

          .stTabs [data-baseweb="tab-list"] { gap: 4px; border-bottom: 1px solid #e3e8ef; }
          .stTabs [data-baseweb="tab"] {
            height: 42px; padding: 0 16px; background: #fbfcfd;
            border-radius: 8px 8px 0 0; font-size: .88rem; font-weight: 500;
            color: #64748b !important;
          }
          .stTabs [aria-selected="true"] {
            background: #e2e8f0 !important; border-bottom: 2px solid #94a3b8 !important;
          }
          .stTabs [aria-selected="true"] * { color: #1e293b !important; font-weight: 600 !important; }

          /* Every colour below is explicit: inheriting theme colours onto our
             own light cards makes the text vanish under a dark theme. */
          div[data-testid="stMetric"] {
            background: #ffffff; border: 1px solid #dbe3ec; border-radius: 10px;
            padding: .85rem 1rem; box-shadow: 0 1px 3px rgba(16,24,40,.06);
          }
          div[data-testid="stMetricLabel"],
          div[data-testid="stMetricLabel"] * {
            color: #475569 !important; font-weight: 600 !important;
            font-size: .82rem !important; letter-spacing: .2px;
          }
          div[data-testid="stMetricValue"],
          div[data-testid="stMetricValue"] * {
            font-size: 1.55rem !important; color: #1f3a5f !important; font-weight: 700 !important;
          }

          section[data-testid="stSidebar"] { background: #f4f7fb; border-right: 1px solid #dbe3ec; }
          section[data-testid="stSidebar"] * { color: #1e293b; }
          .sb-title {
            font-size: .78rem; font-weight: 700; letter-spacing: .6px;
            text-transform: uppercase; color: #64748b; margin: .2rem 0 .5rem;
          }
          .sb-stat {
            display: flex; justify-content: space-between; align-items: baseline;
            background: #ffffff; border: 1px solid #dbe3ec; border-radius: 8px;
            padding: .5rem .75rem; margin-bottom: .4rem;
          }
          .sb-stat .k { color: #475569; font-size: .85rem; font-weight: 500; }
          .sb-stat .v { color: #1f3a5f; font-size: 1.05rem; font-weight: 700; }
          .sb-model {
            background: #e8f0fa; border: 1px solid #c7dbf0; border-radius: 8px;
            padding: .5rem .7rem; margin-bottom: .7rem;
            color: #1f3a5f; font-size: .82rem; font-weight: 600; word-break: break-word;
          }

          .app-footer {
            margin-top: 2rem; padding: 1rem 0; border-top: 1px solid #e3e8ef;
            text-align: center; color: #64748b; font-size: .85rem;
          }
          .app-footer b { color: #2d6ca8; }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ==========================================================================
# vpax parsing
# ==========================================================================

def _read_json_member(z: zipfile.ZipFile, name: str) -> Any:
    return json.loads(z.read(name).decode("utf-8-sig", errors="ignore"))


def _ensure_columns(df: pd.DataFrame, columns: List[str]) -> pd.DataFrame:
    """Guarantee a frame has the given columns, even when it has no rows.

    pd.DataFrame([]) has *no* columns, so a model with (say) no relationships
    would raise KeyError on df["From Table"] instead of simply being empty.
    Fixing the schema here keeps every downstream consumer simple.
    """
    out = df.copy() if not df.empty else pd.DataFrame(columns=columns)
    for col in columns:
        if col not in out.columns:
            out[col] = pd.Series(dtype="object")
    return out


def _is_measure_group(table: Dict[str, Any]) -> bool:
    """True if a table exists only to host measures (i.e. a report screen)."""
    if not table.get("measures"):
        return False
    data_columns = [
        c for c in table.get("columns", [])
        if (c.get("type") or "").lower() != "rownumber"
    ]
    return not data_columns


def _build_screens(bim_tables: List[Dict[str, Any]]) -> pd.DataFrame:
    measure_groups = [t for t in bim_tables if _is_measure_group(t)]
    source = measure_groups or [t for t in bim_tables if t.get("measures")]

    rows = []
    for t in source:
        name = (t.get("name") or "").strip()
        match = re.match(r"^(\d+)\s*(.*)$", name)
        order, display = (int(match.group(1)), match.group(2).strip()) if match else (None, name)
        rows.append({
            "Order": order,
            "Screen / Page Name": display or name,
            "Table Name": name,
            "Measure Count": len(t.get("measures") or []),
            "Hidden": bool(t.get("isHidden", False)),
        })

    df = _ensure_columns(pd.DataFrame(rows), [
        "Order", "Screen / Page Name", "Table Name", "Measure Count", "Hidden",
    ])
    if not df.empty:
        df = df.sort_values(by=["Order", "Screen / Page Name"], na_position="last").reset_index(drop=True)
    return df


def _looks_page_organised(bim_tables: List[Dict[str, Any]]) -> bool:
    """Heuristic: does this model group its measures one-table-per-report-page?

    A .vpax stores no report-page metadata, so pages can only be *guessed*,
    and only when the model follows that convention. Two signals are needed:

      * more than one measure-only table, and
      * nearly all measures live in those tables.

    The second check matters: a model with a couple of generic containers
    ("Measure Table", "OSA Measure Table") plus many measures sitting on data
    tables is organised by subject area, not by page - calling those "screens"
    would be misleading.
    """
    groups = [t for t in bim_tables if _is_measure_group(t)]
    if len(groups) < 2:
        return False
    in_groups = sum(len(t.get("measures") or []) for t in groups)
    total = sum(len(t.get("measures") or []) for t in bim_tables)
    return bool(total) and in_groups / total >= 0.8


def _build_relationships(bim: Dict[str, Any], vpa: Dict[str, Any]) -> pd.DataFrame:
    stats = {
        r.get("RelationshipName"): r
        for r in (vpa.get("Relationships") or [])
        if isinstance(r, dict) and r.get("RelationshipName")
    }
    rows = []
    for r in (bim.get("model") or {}).get("relationships") or []:
        if not isinstance(r, dict):
            continue
        # A relationship missing either endpoint can't be drawn or reasoned
        # about, and would otherwise surface as a None-named phantom table.
        if not r.get("fromTable") or not r.get("toTable"):
            continue
        s = stats.get(r.get("name"), {})
        cross = r.get("crossFilteringBehavior") or s.get("CrossFilteringBehavior") or "singleDirection"
        is_active = r.get("isActive", s.get("IsActive", True))
        rows.append({
            "From Table": r.get("fromTable"),
            "From Column": r.get("fromColumn"),
            "From Cardinality": s.get("FromCardinalityType") or ("One" if r.get("fromCardinality") == "one" else "Many"),
            "To Table": r.get("toTable"),
            "To Column": r.get("toColumn"),
            "To Cardinality": s.get("ToCardinalityType") or ("Many" if r.get("toCardinality") == "many" else "One"),
            "Cross Filter Direction": "Both" if str(cross).lower().startswith("both") else "Single",
            "Active": bool(is_active) if is_active is not None else True,
        })
    # Some models define no relationships at all (flat / wide-table designs),
    # so pin the schema rather than returning a column-less frame.
    return _ensure_columns(pd.DataFrame(rows), [
        "From Table", "From Column", "From Cardinality",
        "To Table", "To Column", "To Cardinality",
        "Cross Filter Direction", "Active",
    ])


def extract_sql(m_expression: str) -> str:
    """Pull just the SQL out of a Power Query M expression.

    The SQL sits in a `query = "..."` parameter. It's an M string literal, so
    an embedded double quote is escaped by doubling it - scan accordingly
    rather than searching for a fixed end marker.
    """
    if not m_expression:
        return ""
    match = re.search(r"query\s*=\s*\"", m_expression)
    if not match:
        return ""
    out: List[str] = []
    i = match.end()
    while i < len(m_expression):
        ch = m_expression[i]
        if ch == '"':
            if i + 1 < len(m_expression) and m_expression[i + 1] == '"':
                out.append('"')
                i += 2
                continue
            break
        out.append(ch)
        i += 1
    return "".join(out).strip()


@st.cache_data(show_spinner=False)
def load_report_pages(pbix_bytes: bytes) -> Dict[str, Any]:
    """Read the real report pages out of a .pbix.

    A .vpax holds the semantic model only - it has no notion of report pages.
    The page list lives in the `Report/Layout` part of the .pbix, together
    with the tables and fields each visual binds to.
    """
    with zipfile.ZipFile(io.BytesIO(pbix_bytes)) as z:
        layout_name = next((n for n in z.namelist() if n.rsplit("/", 1)[-1] == "Layout"
                            and n.startswith("Report")), None)
        if layout_name is None:
            raise ValueError(
                "No 'Report/Layout' part found - this doesn't look like a .pbix "
                "report (a .pbip/PBIR project stores pages as separate files)."
            )
        raw = z.read(layout_name)

    # Report/Layout is UTF-16LE in classic .pbix files, UTF-8 in some newer ones.
    for encoding in ("utf-16-le", "utf-8-sig", "utf-8"):
        try:
            data = json.loads(raw.decode(encoding, errors="ignore"))
            break
        except json.JSONDecodeError:
            continue
    else:
        raise ValueError("Could not decode Report/Layout as JSON.")

    if not isinstance(data, dict):
        raise ValueError("Report/Layout has an unexpected shape (expected a JSON object).")

    rows, fields_by_page = [], {}
    for section in data.get("sections") or []:
        if not isinstance(section, dict):
            continue
        try:
            cfg = json.loads(section.get("config") or "{}")
            if not isinstance(cfg, dict):
                cfg = {}
        except (json.JSONDecodeError, TypeError):
            cfg = {}
        name = section.get("displayName") or section.get("name") or "(unnamed page)"
        containers = [vc for vc in (section.get("visualContainers") or []) if isinstance(vc, dict)]
        blob = " ".join(
            [str(vc.get("config") or "") for vc in containers]
            + [json.dumps(section.get("filters", ""), default=str)]
        )
        entities = set(re.findall(r'"Entity"\s*:\s*"([^"]+)"', blob))
        properties = set(re.findall(r'"Property"\s*:\s*"([^"]+)"', blob))
        fields_by_page[name] = {"entities": entities, "properties": properties}
        rows.append({
            "Order": pd.to_numeric(section.get("ordinal"), errors="coerce"),
            "Screen / Page Name": name,
            "Visuals": len(containers),
            # visibility == 1 marks hidden pages: tooltips, drill-throughs, scratch pages.
            "Hidden": cfg.get("visibility") == 1,
        })

    # Pin the schema so a report with no readable sections stays an *empty*
    # frame with the right columns instead of a column-less one that would
    # raise KeyError the moment the UI touches pages["Hidden"].
    pages = _ensure_columns(pd.DataFrame(rows), [
        "Order", "Screen / Page Name", "Visuals", "Hidden",
    ])
    if not pages.empty:
        pages["Hidden"] = pages["Hidden"].fillna(False).astype(bool)
        pages = pages.sort_values("Order", na_position="last").reset_index(drop=True)
    return {"pages": pages, "fields_by_page": fields_by_page}


def tables_used_by_page(page_name: str, report: Dict[str, Any], model: Dict[str, Any]) -> Set[str]:
    """Model tables a report page touches.

    Starts from the entities its visuals bind to, then expands any measure-group
    entity into the tables its measures actually read, via their DAX.
    """
    info = report["fields_by_page"].get(page_name, {})
    known = set(model["all_table_names"])
    entities = {e for e in info.get("entities", set()) if e in known}
    used = set(entities)

    measure_groups = set(model["screens"]["Table Name"]) if not model["screens"].empty else set()
    measures_df = model["measures"]
    wanted = info.get("properties", set())
    for group in entities & measure_groups:
        rows = measures_df[
            (measures_df["TableName"] == group) & (measures_df["MeasureName"].isin(wanted))
        ]
        for expr in rows["MeasureExpression"]:
            used |= find_referenced_tables(expr or "", model["all_table_names"])

    return used - measure_groups


@st.cache_data(show_spinner=False)
def load_model(file_bytes: bytes) -> Dict[str, Any]:
    if not file_bytes:
        raise ValueError("The uploaded file is empty.")

    required = ("DaxVpaView.json", "Model.bim", "DaxModel.json")
    try:
        z = zipfile.ZipFile(io.BytesIO(file_bytes))
    except zipfile.BadZipFile as exc:
        raise ValueError(
            "This file isn't a readable .vpax archive - it may be corrupt, "
            "truncated, or not actually a .vpax export."
        ) from exc

    with z:
        names = set(z.namelist())
        missing = [n for n in required if n not in names]
        if missing:
            raise ValueError(
                f"This doesn't look like a standard .vpax file - missing part(s): "
                f"{', '.join(missing)}. Found: {sorted(names)[:10]}"
            )
        try:
            vpa = _read_json_member(z, "DaxVpaView.json")
            bim = _read_json_member(z, "Model.bim")
            dax_model = _read_json_member(z, "DaxModel.json")
        except json.JSONDecodeError as exc:
            raise ValueError(f"A part of this .vpax contains invalid JSON: {exc}") from exc

    if not isinstance(vpa, dict) or not isinstance(bim, dict):
        raise ValueError("Unexpected .vpax contents - the metadata parts are not JSON objects.")
    if not isinstance(dax_model, dict):
        dax_model = {}

    # `or {}` rather than a .get default: these keys can be present-but-null.
    bim_tables = [
        t for t in ((bim.get("model") or {}).get("tables") or [])
        if isinstance(t, dict)
    ]

    def _records(key: str) -> pd.DataFrame:
        """DataFrame from a VPA section, tolerating null/!list/!dict payloads."""
        payload = vpa.get(key)
        if not isinstance(payload, list):
            return pd.DataFrame()
        return pd.DataFrame([r for r in payload if isinstance(r, dict)])

    # --- Tables (drop Description) ---
    tables_df = _records("Tables")
    if "Description" in tables_df.columns:
        tables_df = tables_df.drop(columns=["Description"])

    # --- Columns (drop DisplayFolder / Description / FormatString) ---
    raw_columns = _records("Columns")
    columns_df = raw_columns.drop(
        columns=[c for c in ("DisplayFolder", "Description", "FormatString") if c in raw_columns.columns]
    )

    # --- Measures (keep only useful fields) ---
    measures_df = _records("Measures")
    if not measures_df.empty:
        measures_df = measures_df[
            [c for c in ("TableName", "MeasureName", "MeasureExpression", "DataType", "FormatString")
             if c in measures_df.columns]
        ]
    measures_df = _ensure_columns(measures_df, ["TableName", "MeasureName", "MeasureExpression"])

    # --- Calculated columns ---
    calc_cols_df = pd.DataFrame()
    if not raw_columns.empty and "ColumnType" in raw_columns.columns:
        calc = raw_columns[raw_columns["ColumnType"].isin(["Calculated", "CalculatedTableColumn"])].copy()
        calc_cols_df = calc[
            [c for c in ("TableName", "ColumnName", "ColumnExpression", "ColumnType", "DataType")
             if c in calc.columns]
        ]
    calc_cols_df = _ensure_columns(calc_cols_df, ["TableName", "ColumnName", "ColumnExpression"])

    # --- Power Query: SQL only ---
    pq_rows = []
    for t in bim_tables:
        for p in t.get("partitions") or []:
            if not isinstance(p, dict):
                continue
            src = p.get("source")
            if not isinstance(src, dict):
                continue
            if src.get("type") == "m" and src.get("expression"):
                expression = src.get("expression")
                # M expressions are sometimes stored as a list of lines.
                if isinstance(expression, list):
                    expression = "\n".join(str(x) for x in expression)
                pq_rows.append({
                    "TableName": t.get("name"),
                    "PartitionName": p.get("name"),
                    "Mode": p.get("mode"),
                    "SQL": extract_sql(str(expression)),
                })

    # --- Lookups used by the DAX best-practice rewriter ---
    measure_names: Set[str] = set()
    column_tables: Dict[str, Set[str]] = {}
    columns_by_table: Dict[str, List[str]] = {}
    for t in bim_tables:
        tname = t.get("name")
        if not tname:
            continue
        for m in t.get("measures") or []:
            if isinstance(m, dict) and m.get("name"):
                measure_names.add(m["name"])
        visible = []
        for c in t.get("columns") or []:
            if not isinstance(c, dict):
                continue
            cname = c.get("name")
            if not cname or (c.get("type") or "").lower() == "rownumber":
                continue
            column_tables.setdefault(cname, set()).add(tname)
            if not c.get("isHidden"):
                visible.append(cname)
        columns_by_table[tname] = visible

    return {
        "tables": tables_df,
        "columns": columns_df,
        "measures": measures_df,
        "calc_columns": calc_cols_df,
        "relationships": _build_relationships(bim, vpa),
        "power_query": _ensure_columns(pd.DataFrame(pq_rows),
                                       ["TableName", "PartitionName", "Mode", "SQL"]),
        "screens": _build_screens(bim_tables),
        "page_like": _looks_page_organised(bim_tables),
        "all_table_names": sorted({t["name"] for t in bim_tables if t.get("name")}),
        "columns_by_table": columns_by_table,
        "measure_names": measure_names,
        "column_tables": column_tables,
        "model_name": str(dax_model.get("ModelName") or ""),
    }


# ==========================================================================
# DAX formatting / best-practice rewrite
# ==========================================================================

DAX_KEYWORDS = {"VAR", "RETURN", "IN", "NOT", "AND", "OR", "TRUE", "FALSE", "BLANK"}
BREAKING_FUNCTIONS = {
    "CALCULATE", "CALCULATETABLE", "SUMX", "AVERAGEX", "MINX", "MAXX", "COUNTX",
    "FILTER", "SUMMARIZE", "ADDCOLUMNS", "IF", "SWITCH", "COALESCE", "DIVIDE",
    "TREATAS", "SELECTCOLUMNS", "GROUPBY", "UNION", "EXCEPT", "INTERSECT",
}


def _tokenize_dax(expr: str) -> List[Tuple[str, str]]:
    """Split DAX into (kind, text) tokens, keeping strings/refs intact."""
    tokens: List[Tuple[str, str]] = []
    i, n = 0, len(expr)
    while i < n:
        ch = expr[i]
        if ch == '"':  # string literal
            j = i + 1
            while j < n:
                if expr[j] == '"':
                    if j + 1 < n and expr[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            tokens.append(("string", expr[i:j + 1]))
            i = j + 1
        elif ch == "'":  # quoted table name
            j = expr.find("'", i + 1)
            j = n - 1 if j == -1 else j
            tokens.append(("table", expr[i:j + 1]))
            i = j + 1
        elif ch == "[":  # column / measure reference
            j = expr.find("]", i)
            j = n - 1 if j == -1 else j
            tokens.append(("ref", expr[i:j + 1]))
            i = j + 1
        elif ch == "-" and expr.startswith("--", i):  # line comment
            j = expr.find("\n", i)
            j = n if j == -1 else j
            tokens.append(("comment", expr[i:j]))
            i = j
        elif ch == "/" and expr.startswith("/*", i):  # block comment
            j = expr.find("*/", i)
            j = n if j == -1 else j + 2
            tokens.append(("comment", expr[i:j]))
            i = j
        elif ch.isspace():
            j = i
            while j < n and expr[j].isspace():
                j += 1
            tokens.append(("ws", expr[i:j]))
            i = j
        elif ch.isalpha() or ch == "_":
            j = i
            while j < n and (expr[j].isalnum() or expr[j] in "_."):
                j += 1
            tokens.append(("ident", expr[i:j]))
            i = j
        elif ch.isdigit():
            j = i
            while j < n and (expr[j].isdigit() or expr[j] == "."):
                j += 1
            tokens.append(("number", expr[i:j]))
            i = j
        else:
            tokens.append(("punct", ch))
            i += 1
    return tokens


def _significant(tokens: List[Tuple[str, str]], idx: int, step: int) -> Optional[Tuple[str, str]]:
    """Nearest non-whitespace/comment token in the given direction."""
    j = idx + step
    while 0 <= j < len(tokens):
        if tokens[j][0] not in ("ws", "comment"):
            return tokens[j]
        j += step
    return None


def optimize_dax(expr: str, measure_names: Set[str], column_tables: Dict[str, Set[str]]) -> Tuple[str, str]:
    """Rewrite DAX following well-known best practices.

    Applies only changes that are safe and verifiable from model metadata:
      * measure references are un-qualified   ('T'[M]  ->  [M])
      * column references are fully qualified ([C]     ->  'T'[C], when unambiguous)
      * function names are upper-cased
      * a whole-expression division is converted to DIVIDE()
      * the result is re-indented

    Returns (rewritten_dax, notes).
    """
    if not expr or not expr.strip():
        return "", ""

    tokens = _tokenize_dax(expr)
    notes: List[str] = []
    out: List[Tuple[str, str]] = []

    i = 0
    while i < len(tokens):
        kind, text = tokens[i]

        if kind == "table":
            nxt = _significant(tokens, i, 1)
            if nxt and nxt[0] == "ref":
                ref_name = nxt[1][1:-1]
                if ref_name in measure_names:
                    # Best practice: don't table-qualify measures.
                    out.append(("ref", nxt[1]))
                    notes.append(f"Un-qualified measure reference [{ref_name}]")
                    j = tokens.index(nxt, i + 1)
                    i = j + 1
                    continue
            out.append((kind, text))
            i += 1
            continue

        if kind == "ref":
            prev = _significant(tokens, i, -1)
            already_qualified = prev is not None and (
                prev[0] == "table" or (prev[0] == "ident" and prev[1] not in DAX_KEYWORDS)
            )
            ref_name = text[1:-1]
            if not already_qualified and ref_name not in measure_names:
                owners = column_tables.get(ref_name, set())
                if len(owners) == 1:
                    owner = next(iter(owners))
                    out.append(("table", f"'{owner}'"))
                    out.append((kind, text))
                    notes.append(f"Qualified column reference '{owner}'[{ref_name}]")
                    i += 1
                    continue
                if len(owners) > 1:
                    notes.append(
                        f"[{ref_name}] is ambiguous - it exists on {len(owners)} tables; qualify it manually"
                    )
            out.append((kind, text))
            i += 1
            continue

        if kind == "ident":
            nxt = _significant(tokens, i, 1)
            if nxt and nxt[0] == "punct" and nxt[1] == "(" and text.upper() != text:
                out.append((kind, text.upper()))
                notes.append("Upper-cased DAX function names")
                i += 1
                continue
            if text.upper() in DAX_KEYWORDS and text.upper() != text:
                out.append((kind, text.upper()))
                notes.append("Upper-cased DAX keywords")
                i += 1
                continue

        out.append((kind, text))
        i += 1

    # Whole-expression division -> DIVIDE(), which handles divide-by-zero.
    rewritten = _maybe_apply_divide(out, notes)

    formatted = _format_dax_tokens(_tokenize_dax(rewritten))

    if re.search(r"IFERROR\s*\(", formatted, re.I):
        notes.append("IFERROR() found - prefer DIVIDE() or COALESCE(), which are faster")
    if re.search(r"\bFILTER\s*\(\s*'?[A-Za-z_]", formatted) and "ALL" not in formatted.upper():
        notes.append("FILTER() over a whole table - consider filtering a single column instead")
    if "/" in "".join(t for k, t in out if k == "punct") and "DIVIDE" not in formatted.upper():
        notes.append("Division operator used - consider DIVIDE() to handle division by zero")

    seen: Set[str] = set()
    unique_notes = [n for n in notes if not (n in seen or seen.add(n))]
    return formatted, "; ".join(unique_notes)


def _maybe_apply_divide(tokens: List[Tuple[str, str]], notes: List[str]) -> str:
    """Convert `A / B` to `DIVIDE(A, B)` only when it is the entire expression."""
    depth = 0
    split_at = None
    for idx, (kind, text) in enumerate(tokens):
        if kind == "punct":
            if text == "(":
                depth += 1
            elif text == ")":
                depth -= 1
            elif text == "/" and depth == 0:
                if split_at is not None:  # more than one top-level division
                    return "".join(t for _, t in tokens)
                split_at = idx

    if split_at is None:
        return "".join(t for _, t in tokens)

    left = "".join(t for _, t in tokens[:split_at]).strip()
    right = "".join(t for _, t in tokens[split_at + 1:]).strip()
    if not left or not right or "\n" in left or "\n" in right:
        return "".join(t for _, t in tokens)

    notes.append("Replaced '/' with DIVIDE() to guard against division by zero")
    return f"DIVIDE({left}, {right})"


def _format_dax_tokens(tokens: List[Tuple[str, str]]) -> str:
    """Re-indent DAX: one argument per line inside breaking functions."""
    parts: List[str] = []
    indent = 0
    stack: List[bool] = []  # True when the paren belongs to a breaking function

    def newline() -> None:
        parts.append("\n" + "    " * indent)

    prev_significant: Optional[str] = None
    for idx, (kind, text) in enumerate(tokens):
        if kind == "ws":
            if "\n" in text and parts and not parts[-1].endswith("\n"):
                continue  # existing newlines are re-created by the formatter
            if parts and not parts[-1].endswith((" ", "\n", "(")):
                parts.append(" ")
            continue

        if kind == "comment":
            # A `--` comment runs to end of line, so the next token MUST start
            # on a new line - otherwise the comment swallows real code and the
            # rewritten DAX becomes invalid.
            if parts and not parts[-1].endswith(("\n", " ")):
                parts.append(" ")
            parts.append(text.rstrip())
            if text.lstrip().startswith("--"):
                newline()
        elif kind == "punct" and text == "(":
            is_breaking = (prev_significant or "").upper() in BREAKING_FUNCTIONS
            stack.append(is_breaking)
            parts.append("(")
            if is_breaking:
                indent += 1
                newline()
        elif kind == "punct" and text == ")":
            was_breaking = stack.pop() if stack else False
            if was_breaking:
                indent = max(0, indent - 1)
                newline()
            parts.append(")")
        elif kind == "punct" and text == "{":
            # A brace list like IN {"a","b"} is a single value - never split it.
            stack.append(False)
            parts.append("{")
        elif kind == "punct" and text == "}":
            if stack:
                stack.pop()
            parts.append("}")
        elif kind == "punct" and text == ",":
            parts.append(",")
            if stack and stack[-1]:
                newline()
            else:
                parts.append(" ")
        elif kind == "ident" and text.upper() in ("VAR", "RETURN"):
            if parts and "".join(parts).strip():
                parts.append("\n" + "    " * indent)
            parts.append(text)
            parts.append(" ")
        elif kind == "punct" and text in "=<>+-*/&":
            # Binary operators read better with surrounding spaces.
            if parts and not parts[-1].endswith((" ", "\n")):
                parts.append(" ")
            parts.append(text)
            parts.append(" ")
        else:
            if parts and parts[-1].endswith(","):
                parts.append(" ")
            parts.append(text)

        if kind not in ("ws", "comment"):
            prev_significant = text

    return _tidy_outside_literals("".join(parts)).strip()


def _tidy_outside_literals(text: str) -> str:
    """Apply whitespace/operator clean-ups, but never inside a literal.

    Table names and string literals may legitimately contain runs of spaces
    (e.g. 'LocationDim  Connect with ...'). Rewriting those would point the
    DAX at a table that does not exist, so quoted spans are copied verbatim.
    """
    out: List[str] = []
    buf: List[str] = []

    def flush() -> None:
        if not buf:
            return
        chunk = "".join(buf)
        # Collapse mid-line space runs; leading indentation must survive.
        chunk = re.sub(r"(?<=\S)[ \t]{2,}", " ", chunk)
        chunk = re.sub(r"[ \t]+\n", "\n", chunk)
        chunk = re.sub(r"\n{3,}", "\n\n", chunk)
        chunk = re.sub(r"\(\s*\n\s*\)", "()", chunk)
        # Re-pair comparison operators split by the tokenizer.
        chunk = re.sub(r"<\s*=", "<=", chunk)
        chunk = re.sub(r">\s*=", ">=", chunk)
        chunk = re.sub(r"<\s*>", "<>", chunk)
        out.append(chunk)
        buf.clear()

    i, n = 0, len(text)
    while i < n:
        ch = text[i]
        if ch == "'":  # quoted table name
            flush()
            j = text.find("'", i + 1)
            j = n - 1 if j == -1 else j
            out.append(text[i:j + 1])
            i = j + 1
        elif ch == '"':  # string literal ("" escapes a quote)
            flush()
            j = i + 1
            while j < n:
                if text[j] == '"':
                    if j + 1 < n and text[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            j = n - 1 if j >= n else j
            out.append(text[i:j + 1])
            i = j + 1
        else:
            buf.append(ch)
            i += 1
    flush()
    return "".join(out)


def _same_dax(a: str, b: str) -> bool:
    """True if two DAX snippets differ only by formatting.

    The rewriter always re-indents and re-spaces operators, so comparing raw
    strings - or even whitespace-collapsed strings - would flag every row as
    'changed'. Comparing the non-whitespace token streams ignores layout
    entirely and leaves only genuine edits (renamed refs, DIVIDE, casing).
    """
    def tokens(s: str) -> List[str]:
        return [t for k, t in _tokenize_dax(str(s or "")) if k not in ("ws", "comment")]
    return tokens(a) == tokens(b)


@st.cache_data(show_spinner=False)
def add_optimized_dax(df: pd.DataFrame, expr_col: str, measure_names: Set[str],
                      column_tables: Dict[str, Set[str]]) -> pd.DataFrame:
    """Append best-practice DAX + notes columns to a measures/columns frame.

    Rows whose rewrite is equivalent to the original are left blank, and if
    nothing in the frame changed at all both columns are omitted entirely.
    """
    if df.empty or expr_col not in df.columns:
        return df

    result = df.copy()
    rewritten, notes = [], []
    for expr in result[expr_col]:
        new_expr, note = optimize_dax(str(expr or ""), measure_names, column_tables)
        if _same_dax(new_expr, expr):
            rewritten.append("")   # nothing to suggest for this row
            notes.append("")
        else:
            rewritten.append(new_expr)
            notes.append(note)

    if not any(rewritten):
        return result  # no row benefits - don't show the columns at all

    result["DAX (Best Practice)"] = rewritten
    result["Best-Practice Notes"] = notes
    return result


# ==========================================================================
# Schema shape detection + ER diagram
# ==========================================================================

def classify_schema(tables: Set[str], rel_df: pd.DataFrame) -> Dict[str, Any]:
    """Identify fact/dimension tables and whether it's a Star/Snowflake/Galaxy."""
    edges = rel_df[rel_df["From Table"].isin(tables) & rel_df["To Table"].isin(tables)]

    many_side, one_side = {}, {}
    for _, r in edges.iterrows():
        many_side[r["From Table"]] = many_side.get(r["From Table"], 0) + 1
        one_side[r["To Table"]] = one_side.get(r["To Table"], 0) + 1

    # A fact sits on the "many" end of at least two relationships; a dimension
    # is anything that is looked up (the "one" end).
    facts = {t for t, c in many_side.items() if c >= 2}
    if not facts and many_side:
        top = max(many_side.values())
        facts = {t for t, c in many_side.items() if c == top}
    dims = {t for t in tables if t in one_side and t not in facts}

    # Distance from the nearest fact - a dimension joined to another dimension
    # (depth >= 2) is what makes a schema a snowflake rather than a star.
    depth = {t: 0 for t in facts}
    frontier = set(facts)
    while frontier:
        nxt = set()
        for _, r in edges.iterrows():
            a, b = r["From Table"], r["To Table"]
            for x, y in ((a, b), (b, a)):
                if x in frontier and y not in depth:
                    depth[y] = depth[x] + 1
                    nxt.add(y)
        frontier = nxt
    for t in tables:
        depth.setdefault(t, 0)

    max_depth = max(depth.values()) if depth else 0
    if len(facts) > 1:
        shape = "Galaxy"
    elif max_depth > 1:
        shape = "Snowflake"
    elif facts:
        shape = "Star"
    else:
        shape = "Flat"

    return {
        "facts": facts, "dims": dims, "depth": depth,
        "shape": shape, "max_depth": max_depth, "edges": edges,
    }


def _dot_id(name: Any) -> str:
    """Quote a name so it is always a valid DOT identifier.

    Power BI table names may contain quotes or backslashes; dropping them
    straight into `"{name}"` would terminate the identifier early and produce
    a DOT file that fails to parse (a blank diagram), so escape them.
    """
    return '"' + str(name).replace("\\", "\\\\").replace('"', '\\"') + '"'


def _crow(cardinality: str) -> str:
    """Crow's-foot arrow: 'many' gets the crow, 'one' gets the tee."""
    return "crow" if not str(cardinality).lower().startswith("one") else "tee"


def _card_symbol(cardinality: str) -> str:
    """Compact cardinality marker, as shown in the Power BI model view."""
    return "1" if str(cardinality).lower().startswith("one") else "*"


def _bold_slack(text: str, point_size: float = 12.0) -> str:
    """Non-breaking spaces that reserve the extra width bold text needs.

    Graphviz measures a label with the *regular* Helvetica metrics but emits
    font-weight="bold" into the SVG, so the browser then draws text ~6% wider
    than the box that was reserved for it - which is what clips long table
    names off at the right edge. Padding the label with a proportional number
    of spaces widens the measured box to match what actually gets drawn.
    """
    space_w = 0.278 * point_size      # Helvetica space advance
    avg_char_w = 0.556 * point_size   # Helvetica average advance
    extra_px = 0.06 * len(text) * avg_char_w
    return "&nbsp;" * max(1, math.ceil(extra_px / space_w))


def _edge_chip(text: str, colour: str = "#1e293b", size: int = 11) -> str:
    """An edge label on an opaque chip.

    Bare edge text sits directly on top of whatever it overlaps - a
    relationship line, or a neighbouring table's column list - which is what
    made the cardinality markers unreadable. Drawing it on a white cell with
    a hairline border keeps it legible wherever the layout engine puts it.
    """
    return (
        f'<<TABLE BORDER="0" CELLBORDER="1" CELLSPACING="0" CELLPADDING="2" '
        f'COLOR="#cbd5e1" BGCOLOR="#ffffff">'
        f'<TR><TD><FONT POINT-SIZE="{size}" COLOR="{colour}"><B>{text}</B></FONT></TD></TR>'
        f"</TABLE>>"
    )


def key_columns_by_table(rel_df: pd.DataFrame) -> Dict[str, Set[str]]:
    """Columns that actually participate in relationships, per table."""
    keys: Dict[str, Set[str]] = {}
    for _, r in rel_df.iterrows():
        keys.setdefault(r["From Table"], set()).add(r["From Column"])
        keys.setdefault(r["To Table"], set()).add(r["To Column"])
    return keys


def build_er_dot(
    tables: Set[str],
    rel_df: pd.DataFrame,
    columns_by_table: Dict[str, List[str]],
    info: Dict[str, Any],
    detail: str = "keys",
    engine: str = "dot",
) -> str:
    """Build a compact ER-style Graphviz diagram (crow's foot + filter direction).

    `detail` controls how much of each table is drawn:
      "names" - table name only (most compact)
      "keys"  - only the columns used in relationships (default)
      "all"   - every visible column, capped at MAX_DIAGRAM_COLUMNS
    """
    facts, dims = info["facts"], info["dims"]
    edges = info["edges"]
    keys = key_columns_by_table(rel_df)

    if not tables:
        return "digraph SemanticModel {}"

    root = next(iter(sorted(facts & tables))) if (facts & tables) else next(iter(sorted(tables)))

    # Tighter spacing for the compact modes - big boxes need more room.
    spread = 1.0 if detail == "names" else (1.35 if detail == "keys" else 1.9)

    # Helvetica (a core PostScript font) has built-in width metrics in
    # Graphviz/viz.js. Naming a system font like "Segoe UI" instead makes
    # the renderer *guess* each character's width, and on machines where
    # that font isn't installed (e.g. any Mac) the guess is wrong - table
    # headers get sized too narrow and the name is clipped. Helvetica avoids
    # that entirely, on every platform.
    FONT = "Helvetica"

    lines = ["digraph SemanticModel {"]
    if engine == "twopi":
        # Generous ranksep: the edges need to be long enough for the
        # cardinality and filter-direction chips to sit in clear space
        # rather than piling up on top of each other near the nodes.
        lines.append(f'  graph [layout=twopi, root={_dot_id(root)}, ranksep="{2.35 * spread:.2f} equally", '
                     'overlap=false, splines=true, '
                     'bgcolor="transparent", pad=0.35];')
    elif engine in ("neato", "fdp"):
        lines.append(f'  graph [layout={engine}, overlap=false, splines=true, '
                     f'sep="+{int(10 * spread)}", '
                     'bgcolor="transparent", pad=0.35];')
    else:
        # splines=ortho cannot render edge labels, so use curved splines.
        lines.append(f'  graph [rankdir=LR, splines=spline, nodesep={0.35 * spread:.2f}, '
                     f'ranksep={0.9 * spread:.2f}, bgcolor="transparent", pad=0.35];')
    lines.append(f'  node [shape=plaintext, fontname="{FONT}", fontsize=9, margin=0];')
    lines.append(f'  edge [fontname="{FONT}", fontsize=11, color="#94a3b8", fontcolor="#334155"];')

    for table in sorted(tables):
        is_fact = table in facts
        is_dim = table in dims
        if is_fact:
            header_bg, border = "#1f3a5f", "#1f3a5f"
        elif is_dim:
            header_bg, border = "#2d6ca8", "#2d6ca8"
        else:
            header_bg, border = "#64748b", "#64748b"

        if detail == "names":
            shown, extra = [], 0
        elif detail == "keys":
            table_keys = keys.get(table, set())
            shown = [c for c in columns_by_table.get(table, []) if c in table_keys]
            # Relationship columns are often hidden, so fall back to the keys.
            shown = shown or sorted(table_keys)
            extra = 0
        else:
            cols = columns_by_table.get(table, [])
            shown = cols[:MAX_DIAGRAM_COLUMNS]
            extra = len(cols) - len(shown)

        rows = "".join(
            f'<TR><TD ALIGN="LEFT" BGCOLOR="#ffffff">'
            f'<FONT POINT-SIZE="10" COLOR="#334155">{html.escape(c)}</FONT></TD></TR>'
            for c in shown
        )
        if extra > 0:
            rows += (f'<TR><TD ALIGN="LEFT" BGCOLOR="#ffffff"><FONT POINT-SIZE="9" '
                     f'COLOR="#94a3b8"><I>… +{extra} more</I></FONT></TD></TR>')

        # A little breathing room around the name keeps long table names
        # from ever touching the box edge (the visible cause of "cut off"
        # text), and ROUNDED corners give the box a cleaner, less spreadsheet-y look.
        label = (
            f'<<TABLE BORDER="1.4" CELLBORDER="0" CELLSPACING="0" CELLPADDING="6" '
            f'STYLE="ROUNDED" COLOR="{border}" BGCOLOR="#ffffff">'
            f'<TR><TD ALIGN="LEFT" BGCOLOR="{header_bg}" CELLPADDING="7">'
            f'<FONT COLOR="#ffffff" POINT-SIZE="12"><B>{html.escape(table)}'
            f'{_bold_slack(table)}</B></FONT>'
            f'</TD></TR>'
            f"{rows}</TABLE>>"
        )
        # NB: the root node is set via the graph-level `root` attribute only.
        # Also setting `root=true` on the node crashes the twopi engine.
        lines.append(f'  {_dot_id(table)} [label={label}];')

    for _, r in edges.iterrows():
        # A relationship pointing at a table that isn't drawn would make
        # Graphviz invent an empty node for it, so skip those. Self-joins
        # are skipped too: they add a loop that carries no readable meaning.
        if r["From Table"] not in tables or r["To Table"] not in tables:
            continue
        if r["From Table"] == r["To Table"]:
            continue
        both = r["Cross Filter Direction"] == "Both"
        colour = "#2563eb" if both else "#7c8aa0"
        style = "solid" if r["Active"] else "dashed"
        # Every relationship carries three facts, each with its own encoding:
        #   cardinality      - crow's foot / bar glyph, plus a "*" or "1" chip
        #   filter direction - "<->" (both) or "->" (single) chip, and colour
        #   active/inactive  - solid vs dashed line
        # The chips are drawn on an opaque background (see _edge_chip) so they
        # stay readable even when the layout puts them over a line or a box.
        tip = (
            f'{r["From Table"]}[{r["From Column"]}] -> {r["To Table"]}[{r["To Column"]}] | '
            f'{r["Cross Filter Direction"]} filter | {"Active" if r["Active"] else "Inactive"}'
        )
        tail_chip = _edge_chip(_card_symbol(r["From Cardinality"]))
        head_chip = _edge_chip(_card_symbol(r["To Cardinality"]))
        dir_chip = _edge_chip("&#8596;" if both else "&#8594;",
                              colour="#2563eb" if both else "#475569", size=12)
        lines.append(
            f'  {_dot_id(r["From Table"])} -> {_dot_id(r["To Table"])} ['
            f'dir=both, '
            f'arrowtail={_crow(r["From Cardinality"])}, '
            f'arrowhead={_crow(r["To Cardinality"])}, '
            f'color="{colour}", style={style}, penwidth={2.2 if both else 1.6}, '
            f'taillabel={tail_chip}, headlabel={head_chip}, label={dir_chip}, '
            f'labeldistance=1.35, labelangle=16, '
            f'tooltip="{html.escape(tip)}", labeltooltip="{html.escape(tip)}", '
            f'edgetooltip="{html.escape(tip)}", arrowsize=1.15];'
        )

    lines.append("}")
    return "\n".join(lines)


def static_diagram_panel(dot: str, engine: str, filename: str, max_height: int = 420) -> None:
    """Render a compact, static ER diagram with a PNG download button.

    No pan/zoom/drag controls - just the diagram, scaled to fit the width of
    its column, plus one button. Uses viz.js in the browser to lay the DOT
    out and rasterise it, so PNG export works even with no Graphviz binary
    installed locally.
    """
    # Escaping "</" prevents a stray </script> in a table name from closing
    # the inline script tag early; "<\/" is equivalent inside a JS string.
    payload = json.dumps({"dot": dot, "engine": engine, "filename": filename}).replace("</", r"<\/")
    components.html(
        """
<div style="font-family:Segoe UI,system-ui,sans-serif">
  <div id="wrap" style="border:1px solid #dbe3ec;border-radius:8px;background:#fff;
       max-height:__MAXH__px;overflow:auto;padding:8px"></div>
  <div style="margin-top:8px;display:flex;align-items:center;gap:10px">
    <button id="png" class="b" disabled>⬇ Download PNG</button>
    <span id="msg" style="color:#64748b;font-size:12px">Rendering…</span>
  </div>
</div>
<style>
  .b{background:#2d6ca8;color:#fff;border:0;border-radius:6px;padding:6px 14px;
     font-size:12px;font-weight:600;cursor:pointer}
  .b:hover{background:#1f3a5f}
  .b:disabled{background:#cbd5e1;cursor:not-allowed}
</style>
<script src="https://cdn.jsdelivr.net/npm/viz.js@2.1.2/viz.js"></script>
<script src="https://cdn.jsdelivr.net/npm/viz.js@2.1.2/full.render.js"></script>
<script>
const CFG = __PAYLOAD__;
const msg = document.getElementById('msg');
const wrap = document.getElementById('wrap');
const btn = document.getElementById('png');

function download(blob, name){
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}

if (typeof Viz === 'undefined') {
  msg.textContent = 'Renderer unavailable offline.';
} else {
  new Viz().renderSVGElement(CFG.dot, {engine: CFG.engine}).then(el => {
    // Scale to the column's width instead of the diagram's native size -
    // this is what keeps it compact rather than sprawling off-screen.
    el.style.width = '100%';
    el.style.height = 'auto';
    el.style.display = 'block';
    wrap.appendChild(el);
    msg.textContent = '';
    btn.disabled = false;

    btn.onclick = () => {
      // Rasterise at the diagram's true (unscaled) size, at 2x, for a crisp PNG.
      const xml = new XMLSerializer().serializeToString(el);
      const vb = el.viewBox.baseVal;
      const w = (vb && vb.width) || el.getBoundingClientRect().width;
      const h = (vb && vb.height) || el.getBoundingClientRect().height;
      const img = new Image();
      img.onload = () => {
        const c = document.createElement('canvas');
        c.width = w * 2; c.height = h * 2;
        const ctx = c.getContext('2d');
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, c.width, c.height);
        ctx.drawImage(img, 0, 0, c.width, c.height);
        c.toBlob(b => download(b, CFG.filename + '.png'), 'image/png');
      };
      img.onerror = () => { msg.textContent = 'PNG export failed.'; };
      img.src = 'data:image/svg+xml;base64,' +
                btoa(unescape(encodeURIComponent(xml)));
    };
  }).catch(err => { msg.textContent = 'Render error: ' + err; });
}
</script>
        """.replace("__PAYLOAD__", payload).replace("__MAXH__", str(max_height)),
        height=max_height + 70,
    )


# ==========================================================================
# Screen -> tables resolution
# ==========================================================================

@lru_cache(maxsize=64)
def _table_matchers(table_names: Tuple[str, ...]) -> List[Tuple[str, "re.Pattern[str]"]]:
    """Compiled 'is this table referenced?' patterns, one per table.

    Built once per model and cached: the naive version recompiled two regexes
    for every (table x expression) pair, which on a large model is hundreds of
    thousands of compiles and makes the screen tabs visibly slow.
    """
    matchers = []
    for name in table_names:
        if not name:
            continue
        parts = [r"'" + re.escape(name) + r"'"]
        if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name):
            parts.append(r"(?<![A-Za-z0-9_])" + re.escape(name) + r"\s*\[")
        matchers.append((name, re.compile("|".join(parts))))
    return matchers


def find_referenced_tables(expression: str, table_names: List[str]) -> Set[str]:
    """Tables a DAX expression references ('Quoted Name'[Col] or Unquoted[Col])."""
    if not expression:
        return set()
    return {
        name for name, pattern in _table_matchers(tuple(table_names))
        if pattern.search(expression)
    }


def tables_used_by_screen(screen_table: str, measures_df: pd.DataFrame, table_names: List[str]) -> Set[str]:
    exprs = measures_df.loc[measures_df["TableName"] == screen_table, "MeasureExpression"]
    used: Set[str] = set()
    for expr in exprs:
        used |= find_referenced_tables(expr or "", table_names)
    used.discard(screen_table)
    return used


def expand_with_neighbours(tables: Set[str], rel_df: pd.DataFrame) -> Set[str]:
    expanded = set(tables)
    for _, r in rel_df.iterrows():
        if r["From Table"] in tables:
            expanded.add(r["To Table"])
        if r["To Table"] in tables:
            expanded.add(r["From Table"])
    return expanded


def _adjacency(rel_df: pd.DataFrame) -> Dict[str, Set[str]]:
    """Undirected table graph built from the model's relationships."""
    adj: Dict[str, Set[str]] = {}
    for _, r in rel_df.iterrows():
        adj.setdefault(r["From Table"], set()).add(r["To Table"])
        adj.setdefault(r["To Table"], set()).add(r["From Table"])
    return adj


def bridge_tables(tables: Set[str], rel_df: pd.DataFrame, max_hops: int = 1) -> Set[str]:
    """Add the tables needed to actually connect the given ones.

    Measures often reference a fact and a dimension that are joined only
    *through* another table (e.g. two facts sharing a conformed dimension).
    Those intermediates never appear in the DAX, so without them the diagram
    shows disconnected boxes and no joins. This walks the shortest path
    between each pair and pulls in the tables along the way.

    `max_hops` is how many intermediate tables a path may contain.
    """
    adj = _adjacency(rel_df)
    present = [t for t in sorted(tables) if t in adj]
    result = set(tables)
    limit = max_hops + 1  # path length in edges

    # This is a BFS per pair, so cost grows with the square of the table
    # count. Past a few dozen tables the diagram is unreadable anyway, so
    # cap the work rather than letting a huge model hang the page.
    if len(present) > MAX_BRIDGE_TABLES:
        present = present[:MAX_BRIDGE_TABLES]

    for i, source in enumerate(present):
        for target in present[i + 1:]:
            # Breadth-first search, so the first hit is a shortest path.
            queue = deque([[source]])
            seen = {source}
            while queue:
                path = queue.popleft()
                if path[-1] == target:
                    result.update(path)
                    break
                if len(path) > limit:
                    continue
                for neighbour in adj.get(path[-1], ()):
                    if neighbour not in seen:
                        seen.add(neighbour)
                        queue.append(path + [neighbour])
    return result


def fact_subject_areas(model: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Group a model's real tables around each fact table.

    Some vpax files have no per-page measure tables at all (every measure
    sits in one shared 'Measure Table') and no matching .pbix is available,
    so there's no way to recover real screen names - the vpax format simply
    doesn't store them. In that case there's nothing meaningful to call a
    'screen', so instead we fall back to something the model *does* contain:
    its own join structure. A table on the 'many' side of two or more
    relationships is treated as a fact, and each fact plus the tables
    directly joined to it becomes one "subject area" - a stand-in for a
    dashboard page that reflects how the data is actually organised.
    """
    rel_df = model["relationships"]
    measure_groups = set(model["screens"]["Table Name"]) if not model["screens"].empty else set()
    real_tables = set(model["all_table_names"]) - measure_groups
    if rel_df.empty or not real_tables:
        return []
    info = classify_schema(real_tables, rel_df)
    areas = []
    for fact in sorted(info["facts"]):
        tables = expand_with_neighbours({fact}, rel_df) & real_tables
        areas.append({"name": fact, "tables": tables})
    return areas


# ==========================================================================
# Rendering helpers
# ==========================================================================

def _slug(name: Any, fallback: str = "export") -> str:
    """Filesystem-safe download name.

    Guards the case where a name is entirely punctuation (or empty), which
    would otherwise produce a file called just ".csv".
    """
    slug = re.sub(r"\W+", "_", str(name)).strip("_").lower()
    return slug[:80] or fallback


def _stringify_cell(x: Any) -> str:
    if isinstance(x, (list, dict, tuple)):
        return json.dumps(x, default=str)
    try:
        if pd.isna(x):
            return ""
    except (TypeError, ValueError):
        pass
    return str(x)


def _safe(df: pd.DataFrame) -> pd.DataFrame:
    safe = df.copy()
    for col in safe.columns:
        if safe[col].dtype == object:
            safe[col] = safe[col].apply(_stringify_cell)
    return safe


def _excel_engine() -> Optional[str]:
    """First installed Excel writer engine, or None if neither is available."""
    for module, engine in (("openpyxl", "openpyxl"), ("xlsxwriter", "xlsxwriter")):
        try:
            __import__(module)
            return engine
        except ImportError:
            continue
    return None


EXCEL_ENGINE = _excel_engine()


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    """Serialise a frame to .xlsx. Caller must check EXCEL_ENGINE first."""
    if EXCEL_ENGINE is None:
        raise RuntimeError("No Excel writer installed (pip install openpyxl)")
    # Excel sheet names cannot contain : \ / ? * [ ] and cap at 31 chars.
    safe_sheet = re.sub(r"[:\\/?*\[\]]", "-", sheet_name)[:31] or "Sheet1"
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine=EXCEL_ENGINE) as writer:
        _safe(df).to_excel(writer, index=False, sheet_name=safe_sheet)
    return buf.getvalue()


def show_table(df: pd.DataFrame, name: str, height: int = 380, key: str = "") -> None:
    """Render a dataframe with CSV + XLSX download buttons."""
    safe = _safe(df)
    st.dataframe(safe, width="stretch", height=height, hide_index=True)

    slug = _slug(name)
    col1, col2, _ = st.columns([1, 1, 6])
    with col1:
        st.download_button(
            "⬇ CSV", safe.to_csv(index=False).encode("utf-8"),
            file_name=f"{slug}.csv", mime="text/csv",
            key=f"csv_{key or slug}", width="stretch",
        )
    with col2:
        if EXCEL_ENGINE is None:
            st.button(
                "⬇ Excel", key=f"xlsx_{key or slug}", disabled=True,
                width="stretch",
                help="Excel export needs an writer library — run: pip install openpyxl",
            )
        else:
            st.download_button(
                "⬇ Excel", to_excel_bytes(safe, name),
                file_name=f"{slug}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"xlsx_{key or slug}", width="stretch",
            )


# ==========================================================================
# UI
# ==========================================================================

inject_css()

st.markdown(
    """
    <div class="app-hero">
      <h1>🧩 VPAX Semantic Model Explorer</h1>
      <p>Explore the semantic model behind a Power BI report — screens, ER diagrams,
         measures, lineage and source SQL — straight from a .vpax export.</p>
      <span class="badge">Made by Sourin</span>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown('<div class="sb-title">⚙️ Data source</div>', unsafe_allow_html=True)
    with st.expander("📁 Upload .vpax file", expanded=True):
        uploaded = st.file_uploader("Choose a .vpax file", type=["vpax"], label_visibility="collapsed")
        st.caption(
            "Export from **DAX Studio** (Advanced ➜ Export Metadata) or "
            "**Tabular Editor** (File ➜ Export ➜ Metadata)."
        )
    with st.expander("🖥️ Add report pages (.pbix) — optional"):
        pbix_file = st.file_uploader("Choose a .pbix file", type=["pbix"], label_visibility="collapsed")
        st.caption(
            "A .vpax holds the **semantic model only** and carries no report pages. "
            "Add the .pbix to list the real pages and see the model per page."
        )

if not uploaded:
    st.info("👈 Upload a **.vpax** file from the sidebar to get started.")
    st.markdown(f'<div class="app-footer">{AUTHOR}</div>', unsafe_allow_html=True)
    st.stop()

try:
    with st.spinner("Parsing model metadata…"):
        model = load_model(uploaded.getvalue())
except Exception as exc:  # noqa: BLE001
    st.error(f"Could not parse this file as a vpax model: {exc}")
    st.stop()

report: Optional[Dict[str, Any]] = None
if pbix_file is not None:
    try:
        report = load_report_pages(pbix_file.getvalue())
    except Exception as exc:  # noqa: BLE001
        st.warning(f"Could not read report pages from that .pbix: {exc}", icon="⚠️")

with st.sidebar:
    st.markdown('<div class="sb-title">📊 Model at a glance</div>', unsafe_allow_html=True)
    if model["model_name"]:
        st.markdown(
            f'<div class="sb-model">{html.escape(str(model["model_name"]))}</div>',
            unsafe_allow_html=True,
        )
    sidebar_stats = [
        ("Tables", len(model["tables"])),
        ("Measures", len(model["measures"])),
        ("Calculated columns", len(model["calc_columns"])),
        ("Relationships", len(model["relationships"])),
        ("Report screens", len(model["screens"])),
    ]
    st.markdown(
        "".join(
            f'<div class="sb-stat"><span class="k">{k}</span><span class="v">{v}</span></div>'
            for k, v in sidebar_stats
        ),
        unsafe_allow_html=True,
    )
    if EXCEL_ENGINE is None:
        st.warning("Excel export disabled — run `pip install openpyxl`.", icon="⚠️")
    st.markdown(f'<div class="sb-title" style="margin-top:1rem">{AUTHOR}</div>', unsafe_allow_html=True)

use_pages = report is not None and not report["pages"].empty
try:
    subject_areas = [] if use_pages or model["page_like"] else fact_subject_areas(model)
except Exception:  # noqa: BLE001 - a fallback view must never break the app
    subject_areas = []
show_subject_areas = not use_pages and not model["page_like"] and bool(subject_areas)


@contextmanager
def tab_guard(tab_name: str):
    """Keep one failing tab from taking down the whole app.

    Every tab renders independently, so an unexpected shape in one part of a
    model should degrade that tab only - the other seven still work, and the
    error is reported in place with enough detail to act on.
    """
    try:
        yield
    except Exception as exc:  # noqa: BLE001
        st.error(f"Something went wrong rendering **{tab_name}**: `{type(exc).__name__}: {exc}`")
        with st.expander("Technical details"):
            st.code(traceback.format_exc(), language="text")
        st.caption(
            "The other tabs are unaffected. If this looks like a bug, the details above "
            "identify exactly where it happened."
        )


tabs = st.tabs([
    "🖥️ Dashboard Screens", "🌐 Semantic Model by Screen", "🗂️ Tables",
    "🧱 Columns / Schema", "📐 Measures", "🧮 Calculated Columns",
    "🔗 Relationships", "🛢️ Power Query (SQL)",
])

# --- Dashboard Screens ----------------------------------------------------
with tabs[0]:
    with tab_guard('Dashboard Screens'):
        st.subheader("Dashboard screens / pages")
        screens_df = model["screens"]

        if show_subject_areas:
            st.info(
                "**No real page names are available for this file.** A .vpax export carries "
                "no report-page metadata at all - it holds the semantic model only - and this "
                "model doesn't organise its measures one-table-per-page either, so there is "
                "nothing to reliably read a page name from.\n\n"
                "Instead, here's the model broken down by its own join structure: each **fact "
                "table** (one on the 'many' side of two or more relationships) plus the tables "
                "directly joined to it, as a stand-in for a subject area / page.",
                icon="ℹ️",
            )
            c1, c2, c3 = st.columns(3)
            c1.metric("Subject areas", len(subject_areas))
            c2.metric("Fact tables", len(subject_areas))
            c3.metric("Tables covered", len({t for a in subject_areas for t in a["tables"]}))
            st.markdown("")
            areas_df = pd.DataFrame([
                {
                    "Subject Area (fact table)": a["name"],
                    "Related Tables": len(a["tables"]) - 1,
                }
                for a in subject_areas
            ])
            show_table(areas_df, "Subject Areas", height=280, key="subject_areas")
        elif report is not None:
            pages = report["pages"]
            visible_pages = pages[~pages["Hidden"]]
            st.success(
                f"Report pages read from the .pbix — {len(visible_pages)} visible "
                f"of {len(pages)} total.", icon="✅",
            )
            c1, c2, c3 = st.columns(3)
            c1.metric("Visible pages", len(visible_pages))
            c2.metric("Hidden pages", int(pages["Hidden"].sum()))
            c3.metric("Total visuals", int(visible_pages["Visuals"].sum()))
            st.caption(
                "Hidden pages are tooltip, drill-through and scratch pages — they exist in the "
                "file but users never navigate to them."
            )
            show_hidden_pages = st.checkbox("Show hidden pages too", key="show_hidden_pages")
            view = pages if show_hidden_pages else visible_pages
            show_table(
                view[["Screen / Page Name", "Visuals"]].reset_index(drop=True),
                "Report Pages", height=360, key="pages",
            )
        elif screens_df.empty:
            st.info(
                "No measure groups were found in this model, so report pages cannot be inferred. "
                "A .vpax export contains the semantic model only — it holds no report page or "
                "visual layout metadata."
            )
        else:
            if not model["page_like"]:
                st.warning(
                    "**This model does not split measures by report page.** All measures live in a "
                    f"single table (`{screens_df.iloc[0]['Table Name']}`), so there is nothing to "
                    "break down per screen.\n\n"
                    "A .vpax carries **no report-page metadata at all** — page names simply are not "
                    "in the file. To list your actual pages, upload the matching **.pbix** in the "
                    "sidebar (*Add report pages*). Without it, pages can only be guessed when a model "
                    "happens to store each page's measures in its own table "
                    "(e.g. `1 Sales Overview`, `2 Product Detail`).",
                    icon="ℹ️",
                )
            visible = screens_df[~screens_df["Hidden"]]
            label = "Screens" if model["page_like"] else "Measure groups"
            all_measures = model["measures"]
            duplicate_measures = 0
            if not all_measures.empty:
                names = all_measures["MeasureName"].astype(str).str.strip().str.casefold()
                duplicate_measures = int(names.duplicated(keep=False).sum())
            c1, c2, c3 = st.columns(3)
            c1.metric(label, len(visible))
            c2.metric("Total measures", int(visible["Measure Count"].sum()))
            c3.metric(
                "Duplicate measures", duplicate_measures,
                help="Measures that share the exact same name with at least one other measure "
                     "elsewhere in the model — a sign the same logic may have been copy-pasted "
                     "under an identical name instead of reused.",
            )
            st.markdown("")
            show_table(
                visible[["Screen / Page Name", "Measure Count"]].reset_index(drop=True),
                "Dashboard Screens", height=330, key="screens",
            )

# --- Semantic Model by Screen ---------------------------------------------
with tabs[1]:
    with tab_guard('Semantic Model by Screen'):
        st.subheader("Semantic model per screen")
        screens_df = model["screens"]
        rel_df = model["relationships"]

        if rel_df.empty:
            st.warning(
                "**This model defines no relationships between tables**, so there are no joins to "
                "draw. Diagrams below will show the tables a screen uses as separate boxes. This is "
                "normal for flat / wide-table models where each table is queried independently.",
                icon="⚠️",
            )

        skip_main = (not use_pages) and (not show_subject_areas) and (
            screens_df.empty or model["measures"].empty
        )
        labels: List[str] = []
        loop_items: List[Any] = []

        if skip_main:
            st.info(
                "This model has no measure groups to resolve, and no relationships to group tables "
                "by either. Upload the matching **.pbix** in the sidebar to break the model down by "
                "real report page."
            )
        else:
            if use_pages:
                st.caption(
                    "Each page's tables come from the fields its visuals bind to (read from the "
                    ".pbix), then drawn as an ER diagram — crow's foot marks the *many* side, a bar "
                    "marks the *one* side, and blue two-headed edges are bi-directional filters."
                )
            elif show_subject_areas:
                st.caption(
                    "This file has no per-page measure tables and no .pbix was supplied, so there "
                    "are no real screen names to read. Each tab below is instead a **subject area** "
                    "taken from the model's own join structure — a fact table plus everything "
                    "directly joined to it — drawn as an ER diagram."
                )
            else:
                st.caption(
                    "Each screen's tables are resolved by parsing its measures' DAX, then drawn as an "
                    "ER diagram — crow's foot marks the *many* side, a bar marks the *one* side, and "
                    "blue two-headed edges are bi-directional filters."
                )

            st.caption(
                "Diagrams are always laid out radially, with fact table(s) at the centre "
                "surrounded by their dimensions — the clearest read for a star/snowflake model."
            )
            opt2, opt3 = st.columns(2)
            with opt2:
                detail_choice = st.selectbox(
                    "Table detail",
                    ["Key columns only", "Table names only", "All columns"],
                    key="detail_choice",
                    help="Fewer columns means smaller boxes and a much more compact diagram.",
                )
            with opt3:
                bridge_choice = st.selectbox(
                    "Show joins through",
                    ["1 joining table", "2 joining tables", "Direct joins only"],
                    key="bridge_choice",
                    help="Measures often reference tables that are joined only through another "
                         "table (e.g. two facts sharing a dimension). Those intermediates never "
                         "appear in the DAX, so without them the diagram has no joins to draw.",
                )
            include_neighbours = st.checkbox(
                "Also include every directly-related table", value=False, key="incl_nb",
                help="Broader still — pulls in all neighbours, e.g. dimensions used only by slicers.",
            )
            detail = {
                "Key columns only": "keys",
                "Table names only": "names",
                "All columns": "all",
            }[detail_choice]
            max_hops = {"Direct joins only": 0, "1 joining table": 1, "2 joining tables": 2}[bridge_choice]

            def _numbered(row: "pd.Series", offset: int = 0) -> str:
                """'3. Sales Overview' when the row carries a usable order number."""
                name = str(row["Screen / Page Name"])
                try:
                    if pd.notna(row["Order"]):
                        return f"{int(row['Order']) + offset}. {name}"
                except (TypeError, ValueError):
                    pass
                return name

            if use_pages:
                source_df = report["pages"][~report["pages"]["Hidden"]].reset_index(drop=True)
                labels = [_numbered(r, 1) for _, r in source_df.iterrows()]
                loop_items = list(source_df.iterrows())
            elif show_subject_areas:
                labels = [f"📦 {a['name']}" for a in subject_areas]
                loop_items = list(enumerate(subject_areas))
            else:
                source_df = screens_df[~screens_df["Hidden"]].reset_index(drop=True)
                labels = [_numbered(r) for _, r in source_df.iterrows()]
                loop_items = list(source_df.iterrows())

            # Two report pages may legitimately share a display name; identical
            # tab labels are confusing and can collide, so disambiguate them.
            seen_labels: Dict[str, int] = {}
            unique_labels = []
            for raw_label in labels:
                text = (str(raw_label).strip() or "(unnamed)")
                seen_labels[text] = seen_labels.get(text, 0) + 1
                unique_labels.append(text if seen_labels[text] == 1 else f"{text} ({seen_labels[text]})")
            labels = unique_labels

        # st.tabs() raises if handed an empty list, which is reachable whenever
        # every page/screen is hidden - show a plain message instead.
        if not skip_main and not labels:
            st.info("Every screen in this model is marked hidden, so there is nothing to draw.")

        if labels:
            for sub_tab, (idx, item) in zip(st.tabs(labels), loop_items):
                with sub_tab:
                    if show_subject_areas:
                        screen_table = item["name"]
                        referenced = item["tables"]
                    elif use_pages:
                        row = item
                        screen_table = row["Screen / Page Name"]
                        referenced = tables_used_by_page(row["Screen / Page Name"], report, model)
                    else:
                        row = item
                        screen_table = row["Table Name"]
                        referenced = tables_used_by_screen(
                            screen_table, model["measures"], model["all_table_names"]
                        )
                    used = set(referenced)
                    if include_neighbours:
                        used = expand_with_neighbours(used, rel_df)
                    if max_hops:
                        used = bridge_tables(used, rel_df, max_hops=max_hops)
                    measure_groups = set(screens_df["Table Name"])
                    used = {t for t in used if t not in measure_groups}

                    if not used:
                        st.info(
                            "This page doesn't bind to any model table — typically a text, image "
                            "or navigation page."
                            if use_pages else
                            "No model tables could be resolved from this screen's DAX."
                        )
                        continue

                    info = classify_schema(used, rel_df)
                    edges = info["edges"]

                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("Tables", len(used))
                    m2.metric("Relationships", len(edges))
                    m3.metric("Fact tables", len(info["facts"]))
                    m4.metric("Schema shape", info["shape"])

                    if edges.empty:
                        st.warning(
                            "These tables have no relationships between them. Try raising "
                            "**Show joins through**, or check the **Relationships** tab.",
                            icon="⚠️",
                        )

                    # A table with zero edges to anything else in this group just
                    # floats on the diagram with nothing pointing at it, which
                    # reads as clutter rather than information - drop it from the
                    # picture and name it separately instead.
                    joined = set(edges["From Table"]) | set(edges["To Table"]) if not edges.empty else set()
                    isolated = sorted(used - joined)
                    diagram_tables = used - set(isolated)

                    # A diagram with hundreds of boxes is unreadable and can take
                    # the browser renderer a long time, so draw the best-connected
                    # tables and say plainly that the rest were left out.
                    trimmed = 0
                    if len(diagram_tables) > MAX_DIAGRAM_TABLES:
                        degree = {t: 0 for t in diagram_tables}
                        for _, e in edges.iterrows():
                            for endpoint in (e["From Table"], e["To Table"]):
                                if endpoint in degree:
                                    degree[endpoint] += 1
                        ranked = sorted(diagram_tables, key=lambda t: (-degree.get(t, 0), t))
                        trimmed = len(diagram_tables) - MAX_DIAGRAM_TABLES
                        diagram_tables = set(ranked[:MAX_DIAGRAM_TABLES])
                        st.info(
                            f"This screen touches {len(used)} tables — showing the "
                            f"{MAX_DIAGRAM_TABLES} most connected ones and omitting {trimmed} "
                            "so the diagram stays readable. The full list is in the "
                            "**Relationships** section below.",
                            icon="ℹ️",
                        )

                    # Always radial: facts sit at the centre, dimensions surround them.
                    engine = "twopi"

                    if diagram_tables:
                        dot = build_er_dot(diagram_tables, rel_df, model["columns_by_table"], info,
                                           detail=detail, engine=engine)

                        legend_col, _ = st.columns([3, 1])
                        with legend_col:
                            st.markdown(
                                '<div style="font-size:12px;color:#64748b;display:flex;gap:16px;'
                                'flex-wrap:wrap;margin-bottom:4px">'
                                '<span><b style="color:#1f3a5f">■</b> Fact</span>'
                                '<span><b style="color:#2d6ca8">■</b> Dimension</span>'
                                '<span><b style="color:#64748b">■</b> Other / bridge</span>'
                                '<span><b>*</b> many &nbsp;·&nbsp; <b>1</b> one '
                                '(crow\'s foot / bar)</span>'
                                '<span><b style="color:#475569">→</b> filter flows one way '
                                '(the <i>one</i> side filters the <i>many</i> side)</span>'
                                '<span><b style="color:#2563eb">↔</b> filters flow both ways</span>'
                                '<span>‑ ‑ ‑ inactive relationship</span>'
                                '</div>',
                                unsafe_allow_html=True,
                            )

                        # Compact static diagram with a PNG download button — no pan/zoom UI.
                        static_diagram_panel(dot, engine, filename=_slug(f"{screen_table}_er_diagram"))

                    if isolated:
                        st.caption(
                            "Not directly joined to anything else in this group, so left out of "
                            "the diagram above (still counted in Tables, and listed under "
                            "Relationships/Measures below): " + ", ".join(f"`{t}`" for t in isolated)
                        )

                    with st.expander(f"Relationships ({len(edges)})"):
                        if edges.empty:
                            st.write("No relationships between this screen's tables.")
                        else:
                            show_table(edges.reset_index(drop=True), f"{screen_table} relationships",
                                       height=300, key=f"rel_{idx}")

                    screen_measures = model["measures"][model["measures"]["TableName"] == screen_table]
                    keep = [c for c in ("MeasureName", "MeasureExpression", "FormatString") if c in screen_measures.columns]
                    with st.expander(f"Measures on this screen ({len(screen_measures)})"):
                        show_table(screen_measures[keep].reset_index(drop=True),
                                   f"{screen_table} measures", height=330, key=f"meas_{idx}")

# --- Tables ---------------------------------------------------------------
with tabs[2]:
    with tab_guard('Tables'):
        st.subheader("Tables in the model")
        show_table(model["tables"], "Tables", height=420, key="tables")

# --- Columns / Schema -----------------------------------------------------
with tabs[3]:
    with tab_guard('Columns / Schema'):
        st.subheader("Columns per table")
        schema_df = model["columns"]
        if schema_df.empty:
            st.info("No column metadata found.")
        elif "TableName" not in schema_df.columns:
            # Filtering needs the column; without it just show what we have.
            show_table(schema_df.reset_index(drop=True), "Columns", height=460, key="columns")
        else:
            # key=str: a model can mix numeric and text table names, and
            # sorting those against each other raises TypeError.
            options = sorted(schema_df["TableName"].dropna().unique().tolist(), key=str)
            chosen = st.multiselect("Filter by table", options=options, key="schema_filter")
            view = schema_df[schema_df["TableName"].isin(chosen)] if chosen else schema_df
            show_table(view.reset_index(drop=True), "Columns", height=460, key="columns")

# --- Measures -------------------------------------------------------------
with tabs[4]:
    with tab_guard('Measures'):
        st.subheader("DAX measures")
        meas_df = model["measures"]
        if meas_df.empty:
            st.info("No measures found in this model.")
        else:
            enriched = add_optimized_dax(meas_df, "MeasureExpression", model["measure_names"], model["column_tables"])
            changed = int((enriched["DAX (Best Practice)"] != "").sum()) if "DAX (Best Practice)" in enriched.columns else 0
            st.caption(
                "**DAX (Best Practice)** applies safe, verifiable fixes: measures un-qualified, "
                "columns fully qualified, functions upper-cased, and whole-expression division "
                "converted to DIVIDE(). Rows already following best practice are left blank — "
                f"{changed} of {len(enriched)} measures have a suggestion. Review before applying."
            )
            search = st.text_input("Search measure name or expression…", key="measure_search")
            view = enriched
            if search:
                view = enriched[
                    enriched["MeasureName"].str.contains(search, case=False, na=False)
                    | enriched["MeasureExpression"].str.contains(search, case=False, na=False)
                ]
            show_table(view.reset_index(drop=True), "Measures", height=460, key="measures")

# --- Calculated Columns ---------------------------------------------------
with tabs[5]:
    with tab_guard('Calculated Columns'):
        st.subheader("Calculated columns")
        cols_df = model["calc_columns"]
        if cols_df.empty:
            st.info("No calculated columns found in this model.")
        else:
            enriched = add_optimized_dax(cols_df, "ColumnExpression", model["measure_names"], model["column_tables"])
            show_table(enriched.reset_index(drop=True), "Calculated Columns", height=460, key="calccols")

# --- Relationships --------------------------------------------------------
with tabs[6]:
    with tab_guard('Relationships'):
        st.subheader("Relationships")
        rel_df = model["relationships"]
        if rel_df.empty:
            st.info("No relationships found in this model.")
        else:
            c1, c2, c3 = st.columns(3)
            c1.metric("Relationships", len(rel_df))
            c2.metric("Bi-directional", int((rel_df["Cross Filter Direction"] == "Both").sum()))
            c3.metric("Inactive", int((~rel_df["Active"]).sum()))
            st.markdown("")
            show_table(rel_df, "Relationships", height=460, key="rels")

# --- Power Query ----------------------------------------------------------
with tabs[7]:
    with tab_guard('Power Query (SQL)'):
        st.subheader("Source SQL behind each table")
        pq_df = model["power_query"]
        if pq_df.empty:
            st.info("No Power Query expressions found in this model.")
        else:
            sql_len = pq_df["SQL"].fillna("").astype(str).str.len()
            with_sql = pq_df[sql_len > 0]
            st.caption(f"{len(with_sql)} of {len(pq_df)} partitions embed a SQL query.")
            choice = st.selectbox(
                "View the full SQL for a table",
                ["(all — summary table)"]
                + sorted(with_sql["TableName"].dropna().unique().tolist(), key=str),
            )
            if choice.startswith("(all"):
                show_table(pq_df.reset_index(drop=True), "Power Query SQL", height=460, key="pq")
            else:
                matches = with_sql.loc[with_sql["TableName"] == choice, "SQL"]
                sql = str(matches.iloc[0]) if not matches.empty else ""
                st.code(sql or "-- no SQL found for this table", language="sql")
                st.download_button(
                    "⬇ Download SQL", sql.encode("utf-8"),
                    file_name=f"{_slug(choice, 'query')}.sql",
                    mime="text/plain", key="dl_sql", disabled=not sql,
                )

st.markdown(
    f'<div class="app-footer">🧩 VPAX Semantic Model Explorer &nbsp;·&nbsp; <b>{AUTHOR}</b></div>',
    unsafe_allow_html=True,
)
