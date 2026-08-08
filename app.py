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
import difflib
from collections import deque
from contextlib import contextmanager
from functools import lru_cache
from typing import Any, Dict, List, Optional, Set, Tuple

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from twbx_convert import render_twbx_conversion_page

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
    page_icon=":material/account_tree:",
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

          /* Navigation now uses st.segmented_control (native tab-like
             widget), not a CSS-hacked radio - no custom selector rules
             needed here. See the "Navigation" section below for how
             programmatic jumps (the scorecard's "Open ->" buttons) still
             work by pre-setting st.session_state before the widget renders. */

          /* Scorecard */
          .score-ring {
            width: 108px; height: 108px; border-radius: 50%; display: flex;
            align-items: center; justify-content: center; flex-direction: column;
            font-weight: 800; color: #fff; margin: 0 auto;
          }
          .score-ring .num { font-size: 1.9rem; line-height: 1; }
          .score-ring .lbl { font-size: .62rem; letter-spacing: .5px; text-transform: uppercase; opacity: .85; }
          .issue-card {
            background: #ffffff; border: 1px solid #dbe3ec; border-left: 5px solid #94a3b8;
            border-radius: 10px; padding: .8rem 1rem; margin-bottom: .6rem;
            box-shadow: 0 1px 3px rgba(16,24,40,.05);
          }
          .issue-card.sev-High { border-left-color: #dc2626; }
          .issue-card.sev-Medium { border-left-color: #d97706; }
          .issue-card.sev-Low { border-left-color: #64748b; }
          .issue-card.sev-Clean { border-left-color: #16a34a; }
          .issue-card .title { font-weight: 700; color: #1e293b; font-size: .95rem; }
          .issue-card .desc { color: #64748b; font-size: .82rem; margin-top: .15rem; }
          .badge-pill {
            display:inline-block; padding: .05rem .55rem; border-radius: 999px;
            font-size: .72rem; font-weight: 700; margin-left: .4rem;
          }
          .badge-pill.High { background:#fee2e2; color:#991b1b; }
          .badge-pill.Medium { background:#fef3c7; color:#92400e; }
          .badge-pill.Low { background:#e2e8f0; color:#475569; }
          .badge-pill.Clean { background:#dcfce7; color:#166534; }
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


def _build_date_tables(bim_tables: List[Dict[str, Any]]) -> pd.DataFrame:
    """Flag tables marked as a date table, and their designated date column.

    A .vpax carries model *metadata* only, not row data, so this can confirm
    a table is marked `dataCategory: "Time"` and has a plausible date/key
    column - it cannot verify the dates are actually contiguous.
    """
    rows = []
    for t in bim_tables:
        name = t.get("name")
        if not name:
            continue
        is_time = str(t.get("dataCategory") or "").lower() == "time"
        columns = [c for c in t.get("columns") or [] if isinstance(c, dict)]
        key_cols = [c.get("name") for c in columns if c.get("isKey") and c.get("name")]
        date_cols = [
            c.get("name") for c in columns
            if str(c.get("dataType") or "").lower() in ("datetime", "date") and c.get("name")
        ]
        if not is_time and not date_cols:
            continue
        rows.append({
            "Table": name,
            "Marked As Date Table": is_time,
            "Key Column": ", ".join(key_cols),
            "Date-Typed Columns": ", ".join(date_cols),
        })
    return _ensure_columns(pd.DataFrame(rows), [
        "Table", "Marked As Date Table", "Key Column", "Date-Typed Columns",
    ])


def _build_roles(bim: Dict[str, Any], table_names: List[str]) -> pd.DataFrame:
    """RLS roles and their per-table filter expressions.

    `filterExpression` is DAX, so the same table-reference resolver used
    everywhere else in this app (`find_referenced_tables`) tells us which
    tables a role's filter actually touches.
    """
    rows = []
    for role in (bim.get("model") or {}).get("roles") or []:
        if not isinstance(role, dict):
            continue
        role_name = role.get("name") or ""
        perms = [p for p in (role.get("tablePermissions") or []) if isinstance(p, dict)]
        if not perms:
            rows.append({"Role": role_name, "Table": "", "Filter Expression": "", "Tables Referenced": ""})
            continue
        for p in perms:
            expr = p.get("filterExpression") or ""
            referenced = find_referenced_tables(expr, table_names) if expr else set()
            rows.append({
                "Role": role_name,
                "Table": p.get("name") or "",
                "Filter Expression": expr,
                "Tables Referenced": ", ".join(sorted(referenced)),
            })
    return _ensure_columns(pd.DataFrame(rows), ["Role", "Table", "Filter Expression", "Tables Referenced"])


def _build_perspectives(bim: Dict[str, Any]) -> pd.DataFrame:
    """Flatten perspectives into one row per (perspective, table, object)."""
    rows = []
    for persp in (bim.get("model") or {}).get("perspectives") or []:
        if not isinstance(persp, dict):
            continue
        persp_name = persp.get("name") or ""
        for pt in persp.get("perspectiveTables") or []:
            if not isinstance(pt, dict):
                continue
            table_name = pt.get("name") or ""
            cols = [c for c in (pt.get("perspectiveColumns") or []) if isinstance(c, dict) and c.get("name")]
            meas = [m for m in (pt.get("perspectiveMeasures") or []) if isinstance(m, dict) and m.get("name")]
            for c in cols:
                rows.append({"Perspective": persp_name, "Table": table_name, "Object": c["name"], "Object Type": "Column"})
            for m in meas:
                rows.append({"Perspective": persp_name, "Table": table_name, "Object": m["name"], "Object Type": "Measure"})
            if not cols and not meas:
                rows.append({"Perspective": persp_name, "Table": table_name, "Object": "", "Object Type": "Table"})
    return _ensure_columns(pd.DataFrame(rows), ["Perspective", "Table", "Object", "Object Type"])


_M_ESCAPES = {"lf": "\n", "cr": "\r", "tab": "\t", "#": "#"}


def _decode_m_escapes(text: str) -> str:
    """Turn M's `#(lf)` / `#(tab)` escape sequences back into real characters.

    Power BI writes native queries on one physical line with the newlines
    encoded, so without this the SQL arrives as an unreadable single line and
    any parser choking on it would be our fault, not the query's.
    """
    def sub(match: "re.Match[str]") -> str:
        parts = match.group(1).split(",")
        out = []
        for p in parts:
            p = p.strip().lower()
            if p in _M_ESCAPES:
                out.append(_M_ESCAPES[p])
            elif re.fullmatch(r"[0-9a-f]{4}|[0-9a-f]{8}", p):
                try:
                    out.append(chr(int(p, 16)))
                except ValueError:
                    return match.group(0)
            else:
                return match.group(0)     # unknown escape - leave it alone
        return "".join(out)

    return re.sub(r"#\(([^)]*)\)", sub, text)


def _m_string_literals(m_expression: str) -> List[str]:
    """Every string literal in an M expression, with `""` unescaped."""
    out: List[str] = []
    i, n = 0, len(m_expression)
    while i < n:
        if m_expression[i] != '"':
            i += 1
            continue
        i += 1
        buf: List[str] = []
        while i < n:
            ch = m_expression[i]
            if ch == '"':
                if i + 1 < n and m_expression[i + 1] == '"':
                    buf.append('"')
                    i += 2
                    continue
                i += 1
                break
            buf.append(ch)
            i += 1
        out.append("".join(buf))
    return out


_SQL_START_RE = re.compile(r"^\s*(with|select)\b", re.I)


def extract_sql(m_expression: str) -> str:
    """Pull just the SQL out of a Power Query M expression.

    There are several shapes in the wild - `Sql.Database(..., [Query="..."])`,
    `Value.NativeQuery(source, "...")`, `Odbc.Query(..., "...")` - and rather
    than pattern-match each connector, this scans every M string literal and
    keeps the ones that actually parse as a query. That way a connector we
    haven't seen still works. When a table's query is assembled from more than
    one literal, the longest genuine query wins.
    """
    if not m_expression:
        return ""
    best = ""
    for literal in _m_string_literals(m_expression):
        sql = _decode_m_escapes(literal).strip()
        if len(sql) < 12:
            continue
        looks_like_sql = bool(_SQL_START_RE.match(sql)) or (
            re.search(r"\bselect\b", sql, re.I) and re.search(r"\bfrom\b", sql, re.I)
        )
        if looks_like_sql and len(sql) > len(best):
            best = sql
    return best


def _read_legacy_layout(z: zipfile.ZipFile) -> Optional[Dict[str, Any]]:
    """Read the classic single `Report/Layout` blob, if this .pbix has one."""
    layout_name = next(
        (n for n in z.namelist() if n.rsplit("/", 1)[-1] == "Layout" and n.startswith("Report")), None
    )
    if layout_name is None:
        return None
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
    return data


def _maybe_json(value: Any) -> Any:
    """Parse a value that might be JSON embedded in a string.

    The classic Layout format stores each visual's whole definition as a JSON
    *string* inside a JSON document, sometimes two levels deep. Trying to
    parse anything that looks like an object or array is how the walker below
    reaches those inner definitions without special-casing each key.
    """
    if not isinstance(value, str):
        return value
    s = value.strip()
    if not s or s[0] not in "{[":
        return value
    try:
        return json.loads(s)
    except (json.JSONDecodeError, ValueError):
        return value


def _collect_field_refs(
    node: Any,
    aliases: Optional[Dict[str, str]] = None,
    out: Optional[List[Tuple[str, str, str]]] = None,
    depth: int = 0,
) -> List[Tuple[str, str, str]]:
    """Every (table, field, kind) a chunk of report JSON binds to.

    Power BI writes field references two ways. Sometimes the table is named
    inline (`SourceRef.Entity`), and sometimes the reference points at an
    alias (`SourceRef.Source`) declared in the enclosing query's `From`
    clause - so aliases have to be carried down the tree as we descend, with
    inner scopes shadowing outer ones. Both the classic Layout blob and the
    newer PBIR visual.json use the same shapes, which is why one walker
    serves both.
    """
    if out is None:
        out = []
    if depth > 60:          # cyclical or pathological nesting guard
        return out
    node = _maybe_json(node)

    if isinstance(node, list):
        for item in node:
            _collect_field_refs(item, aliases, out, depth + 1)
        return out
    if not isinstance(node, dict):
        return out

    scope = dict(aliases or {})
    from_clause = node.get("From")
    if isinstance(from_clause, list):
        for src in from_clause:
            if isinstance(src, dict) and src.get("Name") and src.get("Entity"):
                scope[str(src["Name"])] = str(src["Entity"])

    def resolve(expr: Any) -> str:
        """Table name out of a SourceRef, following an alias if needed."""
        expr = _maybe_json(expr)
        if not isinstance(expr, dict):
            return ""
        ref = expr.get("SourceRef")
        if isinstance(ref, dict):
            if ref.get("Entity"):
                return str(ref["Entity"])
            if ref.get("Source"):
                return scope.get(str(ref["Source"]), "")
        # Hierarchy levels wrap another expression one layer deeper.
        for key in ("Expression", "Hierarchy"):
            inner = expr.get(key)
            if isinstance(inner, (dict, str)):
                found = resolve(inner)
                if found:
                    return found
        return ""

    for key, kind in (("Column", "Column"), ("Measure", "Measure")):
        ref = _maybe_json(node.get(key))
        if isinstance(ref, dict) and ref.get("Property"):
            out.append((resolve(ref.get("Expression")), str(ref["Property"]), kind))

    hier = _maybe_json(node.get("HierarchyLevel"))
    if isinstance(hier, dict) and hier.get("Level"):
        out.append((resolve(hier.get("Expression")), str(hier["Level"]), "Hierarchy level"))

    for value in node.values():
        _collect_field_refs(value, scope, out, depth + 1)
    return out


def _visual_binding_rows(page: str, visual_id: str, visual_type: str, definition: Any) -> List[Dict[str, str]]:
    seen: Set[Tuple[str, str, str]] = set()
    rows: List[Dict[str, str]] = []
    for table, field, kind in _collect_field_refs(definition):
        if not field:
            continue
        key = (table, field, kind)
        if key in seen:
            continue
        seen.add(key)
        rows.append({
            "Page": page, "Visual": visual_id, "Visual Type": visual_type or "(unknown)",
            "Kind": kind, "Table": table, "Field": field,
        })
    return rows


def _pages_from_legacy_layout(data: Dict[str, Any]) -> Tuple[pd.DataFrame, Dict[str, Dict[str, Set[str]]]]:
    """Pages + field bindings from the classic single-blob Layout format."""
    rows, fields_by_page = [], {}
    binding_rows: List[Dict[str, str]] = []
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

        # Per-visual, table-qualified bindings. The page-level entity/property
        # sets above stay as they are (other views depend on them) - these are
        # the paired references the broken-visual check needs.
        for idx, vc in enumerate(containers):
            cfg_obj = _maybe_json(vc.get("config"))
            vtype, vid = "", f"visual {idx + 1}"
            if isinstance(cfg_obj, dict):
                vid = str(cfg_obj.get("name") or vid)
                sv = cfg_obj.get("singleVisual")
                if isinstance(sv, dict):
                    vtype = str(sv.get("visualType") or "")
            binding_rows.extend(
                _visual_binding_rows(name, vid, vtype, [cfg_obj, _maybe_json(vc.get("filters"))])
            )
        binding_rows.extend(
            _visual_binding_rows(name, "(page filters)", "filter", _maybe_json(section.get("filters")))
        )
        rows.append({
            "Order": pd.to_numeric(section.get("ordinal"), errors="coerce"),
            "Screen / Page Name": name,
            "Visuals": len(containers),
            # visibility == 1 marks hidden pages: tooltips, drill-throughs, scratch pages.
            "Hidden": cfg.get("visibility") == 1,
        })

    pages = _ensure_columns(pd.DataFrame(rows), ["Order", "Screen / Page Name", "Visuals", "Hidden"])
    if not pages.empty:
        pages["Hidden"] = pages["Hidden"].fillna(False).astype(bool)
        pages = pages.sort_values("Order", na_position="last").reset_index(drop=True)
    return pages, fields_by_page, _bindings_frame(binding_rows)


BINDING_COLUMNS = ["Page", "Visual", "Visual Type", "Kind", "Table", "Field"]


def _bindings_frame(rows: List[Dict[str, str]]) -> pd.DataFrame:
    df = _ensure_columns(pd.DataFrame(rows), BINDING_COLUMNS)
    if df.empty:
        return df
    return df.drop_duplicates().reset_index(drop=True)


def _pages_from_pbir(z: zipfile.ZipFile) -> Optional[Tuple[pd.DataFrame, Dict[str, Dict[str, Set[str]]], pd.DataFrame]]:
    """Pages + field bindings from the newer PBIR project format.

    Recent Power BI Desktop versions export reports as a folder tree
    (`Report/definition/pages/<id>/page.json` + `.../visuals/<id>/visual.json`)
    instead of one `Report/Layout` blob - a plain .pbix can carry either.
    Returns None if this .pbix has no PBIR page definitions either.
    """
    norm = {n.replace("\\", "/"): n for n in z.namelist()}
    index_key = next((k for k in norm if k.endswith("Report/definition/pages/pages.json")), None)
    if index_key is None:
        return None

    index = _read_json_member(z, norm[index_key])
    order = list(index.get("pageOrder") or [])
    prefix = index_key.rsplit("/", 1)[0] + "/"  # ".../Report/definition/pages/"
    if not order:
        # No explicit order recorded - fall back to whatever page folders exist.
        order = sorted({
            k[len(prefix):].split("/", 1)[0]
            for k in norm if k.startswith(prefix) and k.endswith("/page.json")
        })

    rows, fields_by_page = [], {}
    binding_rows: List[Dict[str, str]] = []
    for ordinal, page_id in enumerate(order):
        page_key = f"{prefix}{page_id}/page.json"
        if page_key not in norm:
            continue
        page = _read_json_member(z, norm[page_key])
        name = page.get("displayName") or page.get("name") or page_id

        visual_prefix = f"{prefix}{page_id}/visuals/"
        visual_keys = [
            k for k in norm if k.startswith(visual_prefix) and k.rsplit("/", 1)[-1] == "visual.json"
        ]
        blob = " ".join(z.read(norm[k]).decode("utf-8-sig", errors="ignore") for k in visual_keys)
        entities = set(re.findall(r'"Entity"\s*:\s*"([^"]+)"', blob))
        properties = set(re.findall(r'"Property"\s*:\s*"([^"]+)"', blob))
        fields_by_page[name] = {"entities": entities, "properties": properties}

        for k in visual_keys:
            try:
                vdef = _read_json_member(z, norm[k])
            except (json.JSONDecodeError, KeyError, UnicodeDecodeError):
                continue
            vis = vdef.get("visual") if isinstance(vdef.get("visual"), dict) else {}
            vid = str(vdef.get("name") or k.rsplit("/", 2)[-2])
            binding_rows.extend(
                _visual_binding_rows(name, vid, str(vis.get("visualType") or ""), vdef)
            )

        rows.append({
            "Order": ordinal,
            "Screen / Page Name": name,
            "Visuals": len(visual_keys),
            # A normal visible page simply has no "visibility" key; tooltip/
            # drill-through/hidden pages carry an explicit non-empty value
            # (e.g. "HiddenInViewMode").
            "Hidden": bool(page.get("visibility")),
        })

    pages = _ensure_columns(pd.DataFrame(rows), ["Order", "Screen / Page Name", "Visuals", "Hidden"])
    if not pages.empty:
        pages["Hidden"] = pages["Hidden"].fillna(False).astype(bool)
        pages = pages.sort_values("Order", na_position="last").reset_index(drop=True)
    return pages, fields_by_page, _bindings_frame(binding_rows)


def _read_pbir_report_json(z: zipfile.ZipFile) -> Optional[Dict[str, Any]]:
    norm = {n.replace("\\", "/"): n for n in z.namelist()}
    key = next((k for k in norm if k.endswith("Report/definition/report.json")), None)
    return _read_json_member(z, norm[key]) if key else None


def extract_report_theme(z: zipfile.ZipFile, report_defs: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Resolve the report's active theme to its actual JSON bytes.

    The theme reference (`themeCollection.baseTheme` in PBIR, or a top-level
    `theme`/`resourcePackages` in the classic Layout blob) only names the
    theme and points at a `resourcePackages` entry for its file path - the
    same `StaticResources` layout holds both Power BI's built-in ("shared")
    themes and any theme a user uploaded ("registered") themselves, so both
    are resolved the same way here.
    """
    theme_ref: Optional[Dict[str, Any]] = None
    resource_items: Dict[str, Dict[str, str]] = {}

    for rd in report_defs:
        if not isinstance(rd, dict):
            continue
        tc = rd.get("themeCollection")
        if theme_ref is None and isinstance(tc, dict) and isinstance(tc.get("baseTheme"), dict):
            theme_ref = tc["baseTheme"]
        if theme_ref is None and isinstance(rd.get("theme"), dict) and rd["theme"].get("name"):
            theme_ref = {"name": rd["theme"]["name"], "type": "SharedResources"}
        for pkg in rd.get("resourcePackages") or []:
            if not isinstance(pkg, dict):
                continue
            for item in pkg.get("items") or []:
                if isinstance(item, dict) and item.get("name"):
                    resource_items[item["name"]] = {
                        "path": item.get("path", ""),
                        "package_type": pkg.get("type") or "",
                    }

    if theme_ref is None or not theme_ref.get("name"):
        return {"found": False}

    theme_name = theme_ref["name"]
    is_custom = str(theme_ref.get("type") or "").lower() != "sharedresources"

    norm = {n.replace("\\", "/"): n for n in z.namelist()}
    candidate_paths = []
    item = resource_items.get(theme_name)
    if item:
        base = "RegisteredResources" if str(item["package_type"]).lower() != "sharedresources" else "SharedResources"
        candidate_paths.append(f"Report/StaticResources/{base}/{item['path']}".replace("\\", "/"))
    # Fall back to the conventional built-in location if the report
    # definition didn't list the theme in resourcePackages (older schemas).
    candidate_paths.append(f"Report/StaticResources/SharedResources/BaseThemes/{theme_name}.json")

    matched_key = None
    for path in candidate_paths:
        if path in norm:
            matched_key = norm[path]
            break
    if matched_key is None:
        # Last resort: any part whose filename matches, wherever it lives.
        matched_key = next((n for n in norm if n.replace("\\", "/").endswith(f"/{theme_name}.json")), None)
        matched_key = norm.get(matched_key, matched_key) if matched_key else None

    if matched_key is None:
        return {"found": False, "name": theme_name, "is_custom": is_custom}

    raw = z.read(matched_key)
    try:
        theme_json = json.loads(raw.decode("utf-8-sig"))
    except (json.JSONDecodeError, UnicodeDecodeError):
        theme_json = json.loads(raw.decode("utf-8", errors="ignore"))

    return {"found": True, "name": theme_name, "is_custom": is_custom, "json": theme_json, "raw_bytes": raw}


@st.cache_data(show_spinner=False)
def load_report_pages(pbix_bytes: bytes) -> Dict[str, Any]:
    """Read the real report pages (and active theme) out of a .pbix.

    A .vpax holds the semantic model only - it has no notion of report pages
    or theming. Pages live either in a single `Report/Layout` blob (classic
    format) or under `Report/definition/pages/...` (newer PBIR format,
    produced by recent Power BI Desktop versions) - this reads whichever one
    the file actually has.
    """
    with zipfile.ZipFile(io.BytesIO(pbix_bytes)) as z:
        legacy = _read_legacy_layout(z)
        if legacy is not None:
            pages, fields_by_page, bindings = _pages_from_legacy_layout(legacy)
        else:
            pbir = _pages_from_pbir(z)
            if pbir is None:
                raise ValueError(
                    "No 'Report/Layout' part or 'Report/definition/pages' found - "
                    "this doesn't look like a .pbix report."
                )
            pages, fields_by_page, bindings = pbir

        report_defs = [d for d in (legacy, _read_pbir_report_json(z)) if d is not None]
        theme = extract_report_theme(z, report_defs)

    return {"pages": pages, "fields_by_page": fields_by_page,
            "bindings": bindings, "theme": theme}


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

    # --- Tables ---
    # Description is kept (not dropped): health checks and the data
    # dictionary export both need it, and show_table renders whatever
    # columns are present, so there's no display-side reason to hide it.
    tables_df = _records("Tables")

    # --- Columns ---
    # DisplayFolder/Description/FormatString are kept for the same reason.
    raw_columns = _records("Columns")
    columns_df = raw_columns.copy()

    # --- Measures (keep only useful fields) ---
    measures_df = _records("Measures")
    if not measures_df.empty:
        measures_df = measures_df[
            [c for c in ("TableName", "MeasureName", "MeasureExpression", "DataType",
                         "FormatString", "Description", "DisplayFolder")
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

    all_table_names = sorted({t["name"] for t in bim_tables if t.get("name")})

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
        "all_table_names": all_table_names,
        "columns_by_table": columns_by_table,
        "measure_names": measure_names,
        "column_tables": column_tables,
        "model_name": str(dax_model.get("ModelName") or ""),
        "date_tables": _build_date_tables(bim_tables),
        "roles": _build_roles(bim, all_table_names),
        "perspectives": _build_perspectives(bim),
        # Raw TOM kept so the newer modules (Fabric readiness, display-folder
        # taxonomy, model compare) can read properties VPA never exposes -
        # isHidden, displayFolder, partition mode, hierarchies, lineage tags.
        "bim_tables": bim_tables,
        # The whole TOM document, so the cleanup module can emit a modified
        # Model.bim rather than only describing what to delete.
        "bim": bim,
        # Physical storage detail, used by the compression advisor. Present in
        # DAX Studio exports; absent in some Tabular Editor ones, hence the
        # tolerant _records() reader and the .empty checks downstream.
        "segments": _records("ColumnsSegments"),
        "col_hierarchies": _records("ColumnsHierarchies"),
        "user_hierarchies": _records("UserHierarchies"),
    }


@lru_cache(maxsize=1)
def build_sample_vpax_bytes() -> bytes:
    """A small, self-contained star-schema model for the "try a sample" button.

    New users shouldn't need to go export a real .vpax before they can see
    what this app does. This model is deliberately a little imperfect (a
    bi-directional relationship, an unused column, an inconsistent measure
    name, a missing description) so every audit tab - Model Health, Naming
    Conventions, Unused Objects - has something real to show, not just an
    empty "all clear".
    """
    bim = {
        "model": {
            "tables": [
                {
                    "name": "Date", "dataCategory": "Time",
                    "columns": [
                        {"name": "DateKey", "dataType": "int64", "isKey": True},
                        {"name": "Date", "dataType": "dateTime"},
                        {"name": "Year", "dataType": "int64"},
                        {"name": "Month", "dataType": "string"},
                    ],
                    "measures": [],
                    "partitions": [{"name": "Date", "mode": "import",
                                    "source": {"type": "m", "expression": 'let\n  q = "SELECT * FROM dim_date"\nin\n  q'}}],
                },
                {
                    "name": "Product",
                    "columns": [
                        {"name": "ProductKey", "dataType": "int64", "isKey": True},
                        {"name": "Product Name", "dataType": "string"},
                        {"name": "Category", "dataType": "string"},
                    ],
                    "measures": [],
                    "partitions": [{"name": "Product", "mode": "import",
                                    "source": {"type": "m", "expression": 'let\n  q = "SELECT * FROM dim_product"\nin\n  q'}}],
                },
                {
                    "name": "Customer",
                    "columns": [
                        {"name": "CustomerKey", "dataType": "int64", "isKey": True},
                        {"name": "Customer Name", "dataType": "string"},
                        {"name": "Region", "dataType": "string"},
                        # Deliberately unreferenced - the sample lets Unused
                        # Objects show a real finding on first use.
                        {"name": "internal_notes", "dataType": "string"},
                    ],
                    "measures": [],
                    "partitions": [{"name": "Customer", "mode": "import",
                                    "source": {"type": "m", "expression": 'let\n  q = "SELECT * FROM dim_customer"\nin\n  q'}}],
                },
                {
                    "name": "Sales",
                    "columns": [
                        {"name": "DateKey", "dataType": "int64"},
                        {"name": "ProductKey", "dataType": "int64"},
                        {"name": "CustomerKey", "dataType": "int64"},
                        {"name": "Quantity", "dataType": "int64"},
                        {"name": "Unit Price", "dataType": "double"},
                    ],
                    "measures": [
                        {"name": "Total Sales", "expression": "SUMX(Sales, Sales[Quantity] * Sales[Unit Price])"},
                        {"name": "Total Quantity", "expression": "SUM(Sales[Quantity])"},
                        {"name": "Sales YTD", "expression": "TOTALYTD([Total Sales], 'Date'[Date])"},
                        # Inconsistent casing on purpose, for Naming Conventions.
                        {"name": "total_orders", "expression": "COUNTROWS(Sales)"},
                    ],
                    "partitions": [{"name": "Sales", "mode": "import",
                                    "source": {"type": "m", "expression": 'let\n  q = "SELECT * FROM fact_sales"\nin\n  q'}}],
                },
                # A field parameter nothing binds to - disconnected, so no other
                # check in the app can see it, which is the whole point of
                # having a dedicated one.
                {
                    "name": "Metric Selector",
                    "columns": [
                        {"name": "Metric Selector", "dataType": "string", "type": "calculatedTableColumn"},
                        {"name": "Metric Fields", "dataType": "string", "type": "calculatedTableColumn",
                         "isHidden": True,
                         "extendedProperties": [
                             {"type": "json", "name": "ParameterMetadata",
                              "value": {"version": 3, "kind": 2}},
                         ]},
                        {"name": "Metric Order", "dataType": "int64", "type": "calculatedTableColumn",
                         "isHidden": True},
                    ],
                    "measures": [],
                    "partitions": [{"name": "Metric Selector", "mode": "import", "source": {
                        "type": "calculated",
                        "expression": '{\n  ("Sales", NAMEOF(\'Sales\'[Total Sales]), 0),\n'
                                      '  ("Quantity", NAMEOF(\'Sales\'[Total Quantity]), 1)\n}',
                    }}],
                },
            ],
            "relationships": [
                {"name": "r1", "fromTable": "Sales", "fromColumn": "DateKey",
                 "toTable": "Date", "toColumn": "DateKey", "crossFilteringBehavior": "singleDirection",
                 "isActive": True, "fromCardinality": "many", "toCardinality": "one"},
                {"name": "r2", "fromTable": "Sales", "fromColumn": "ProductKey",
                 "toTable": "Product", "toColumn": "ProductKey", "crossFilteringBehavior": "singleDirection",
                 "isActive": True, "fromCardinality": "many", "toCardinality": "one"},
                # Bi-directional on purpose, for Model Health.
                {"name": "r3", "fromTable": "Sales", "fromColumn": "CustomerKey",
                 "toTable": "Customer", "toColumn": "CustomerKey", "crossFilteringBehavior": "both",
                 "isActive": True, "fromCardinality": "many", "toCardinality": "one"},
            ],
            "roles": [
                # Secures a dimension - the textbook-correct pattern. The
                # filter flows down to Sales.
                {"name": "Regional Manager",
                 "tablePermissions": [{"name": "Customer", "filterExpression": "Customer[Region] = \"West\""}]},
                # Secures the fact instead. Looks configured, but the filter
                # can't travel back up to the dimensions - exactly the trapped-
                # filter case the RLS simulator exists to catch.
                {"name": "Sales Rep (misconfigured)",
                 "tablePermissions": [{"name": "Sales", "filterExpression": "Sales[Quantity] > 0"}]},
            ],
        },
    }
    # Folders on three of the four measures, so the taxonomy view shows both a
    # real folder structure and one measure that fell out of it.
    measure_folders = {"Total Sales": "Sales", "Total Quantity": "Sales", "Sales YTD": "Time Intelligence"}
    row_counts = {"Date": 1461, "Product": 5200, "Customer": 84000, "Sales": 3_200_000}

    def _table_expression(t: Dict[str, Any]) -> str:
        for p in t.get("partitions") or []:
            src = (p or {}).get("source") or {}
            if str(src.get("type")) == "calculated":
                return str(src.get("expression") or "")
        return ""

    vpa = {
        "Tables": [
            {"TableName": t["name"], "RowsCount": row_counts.get(t["name"], 100),
             "Description": "", "IsReferenced": True,
             "TableExpression": _table_expression(t)}
            for t in bim["model"]["tables"]
        ],
        "Columns": [
            {
                "TableName": t["name"], "ColumnName": c["name"], "ColumnType": "Data",
                "DataType": c["dataType"].capitalize(),
                "IsHidden": False, "DisplayFolder": "", "Description": "", "FormatString": "",
                # Key columns left visible on purpose so the taxonomy check has
                # a real "join key in the field list" finding to report.
                "IsAvailableInMDX": True, "IsKey": bool(c.get("isKey")),
                "IsRowNumber": False, "Encoding": "VALUE" if c["dataType"] == "int64" else "HASH",
            }
            for t in bim["model"]["tables"] for c in t["columns"]
        ],
        "Measures": [
            {"TableName": t["name"], "MeasureName": m["name"],
             "MeasureExpression": m["expression"], "FormatString": "",
             "Description": "", "DisplayFolder": measure_folders.get(m["name"], "")}
            for t in bim["model"]["tables"] for m in t["measures"]
        ],
        "Relationships": [],
    }
    dax_model = {"ModelName": "Sample Retail Model.pbix"}

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("Model.bim", json.dumps(bim))
        z.writestr("DaxVpaView.json", json.dumps(vpa))
        z.writestr("DaxModel.json", json.dumps(dax_model))
    return buf.getvalue()


# ==========================================================================
# VertiPaq / Model Size analysis
# ==========================================================================
# DAX Studio / Tabular Editor VertiPaq-Analyzer exports don't use one fixed
# set of field names across tool versions, and no sample .vpax carrying
# these stats is checked into this repo - so nothing here assumes a closed
# schema. Every candidate name is checked for presence before use, and a
# missing field simply means that sub-result is omitted, never an error.

_VPA_SIZE_FIELD_CANDIDATES = (
    "TotalSize", "TableSize", "DataSize", "DictionarySize", "ColumnSize",
)
_VPA_CARDINALITY_CANDIDATES = ("Cardinality", "ColumnCardinality")
_VPA_ENCODING_CANDIDATES = ("Encoding", "ColumnEncoding")


def detect_vpa_size_columns(df: pd.DataFrame) -> List[str]:
    """Which VertiPaq-Analyzer stat fields actually exist on this frame."""
    if df.empty:
        return []
    candidates = _VPA_SIZE_FIELD_CANDIDATES + _VPA_CARDINALITY_CANDIDATES + _VPA_ENCODING_CANDIDATES
    return [c for c in candidates if c in df.columns]


def _name_column(df: pd.DataFrame, *candidates: str) -> Optional[str]:
    return next((c for c in candidates if c in df.columns), None)


def _bytes_to_mb(n: Any) -> Optional[float]:
    """Bytes -> MB, rounded to 2dp. One column, one unit, no ambiguity."""
    try:
        return round(float(n) / (1024.0 * 1024.0), 2)
    except (TypeError, ValueError):
        return None


_SYSTEM_ROWNUMBER_RE = re.compile(r"^RowNumber-[0-9A-Fa-f-]+$")


def _user_facing_columns(columns_df: pd.DataFrame) -> pd.DataFrame:
    """Drop VertiPaq-internal RowNumber pseudo-columns.

    Every table gets an auto-generated row-number column that exists purely
    for the storage engine - it's never something a modeler created or would
    reference, and showing it (as `RowNumber-<guid>`) in size/unused/naming
    analyses is just noise. `columns_by_table`/`column_tables` (built from
    Model.bim) already exclude these; `model["columns"]` (built from
    DaxVpaView.json) doesn't, so callers that iterate it need this filter.
    """
    if columns_df.empty:
        return columns_df
    mask = pd.Series(True, index=columns_df.index)
    if "ColumnType" in columns_df.columns:
        mask &= columns_df["ColumnType"].astype(str).str.lower() != "rownumber"
    if "IsRowNumber" in columns_df.columns:
        mask &= ~columns_df["IsRowNumber"].fillna(False).astype(bool)
    if "ColumnName" in columns_df.columns:
        mask &= ~columns_df["ColumnName"].astype(str).str.match(_SYSTEM_ROWNUMBER_RE)
    return columns_df[mask]


def build_model_size_summary(model: Dict[str, Any]) -> Dict[str, Any]:
    """Aggregate whatever VertiPaq size/cardinality stats are present.

    Degrades field-by-field: an export without VertiPaq stats (or one that
    names them differently) yields fewer sub-results, never a broken table.
    Every size figure is converted to a single MB column (compressed
    in-memory VertiPaq size, not file size) - one unit, no ambiguity.
    """
    tables_df = model["tables"]
    columns_df = _user_facing_columns(model["columns"])
    summary: Dict[str, Any] = {"available": False}

    table_name_col = _name_column(tables_df, "TableName", "Name")
    table_size_col = _name_column(tables_df, *_VPA_SIZE_FIELD_CANDIDATES)
    if table_name_col and table_size_col:
        sized = tables_df[[table_name_col, table_size_col]].dropna(subset=[table_size_col])
        if not sized.empty:
            summary["available"] = True
            total_bytes = float(pd.to_numeric(sized[table_size_col], errors="coerce").sum())
            summary["total_model_size_mb"] = _bytes_to_mb(total_bytes)
            top = sized.assign(**{table_size_col: pd.to_numeric(sized[table_size_col], errors="coerce")})
            top = top.sort_values(table_size_col, ascending=False).head(20)
            top[table_size_col] = top[table_size_col].apply(_bytes_to_mb)
            top = top.rename(columns={table_name_col: "Table", table_size_col: "Size (MB)"})
            summary["top_tables"] = top.reset_index(drop=True)

    col_name_cols = [c for c in (_name_column(columns_df, "TableName"), _name_column(columns_df, "ColumnName")) if c]
    col_size_col = _name_column(columns_df, *_VPA_SIZE_FIELD_CANDIDATES)
    if col_name_cols and col_size_col:
        sized = columns_df[col_name_cols + [col_size_col]].dropna(subset=[col_size_col])
        if not sized.empty:
            summary["available"] = True
            top = sized.assign(**{col_size_col: pd.to_numeric(sized[col_size_col], errors="coerce")})
            top = top.sort_values(col_size_col, ascending=False).head(20)
            top[col_size_col] = top[col_size_col].apply(_bytes_to_mb)
            top = top.rename(columns={col_size_col: "Size (MB)"})
            summary["top_columns"] = top.reset_index(drop=True)

    cardinality_col = _name_column(columns_df, *_VPA_CARDINALITY_CANDIDATES)
    if col_name_cols and cardinality_col:
        sized = columns_df[col_name_cols + [cardinality_col]].dropna(subset=[cardinality_col])
        if not sized.empty:
            summary["available"] = True
            summary["top_cardinality"] = (
                sized.assign(**{cardinality_col: pd.to_numeric(sized[cardinality_col], errors="coerce")})
                .sort_values(cardinality_col, ascending=False)
                .head(20)
                .rename(columns={cardinality_col: "Distinct Values"})
                .reset_index(drop=True)
            )

    data_col = "DataSize" if "DataSize" in columns_df.columns else None
    dict_col = "DictionarySize" if "DictionarySize" in columns_df.columns else None
    if col_name_cols and data_col and dict_col:
        breakdown = columns_df[col_name_cols + [data_col, dict_col]].dropna(subset=[data_col, dict_col], how="all")
        if not breakdown.empty:
            summary["available"] = True
            breakdown = breakdown.copy()
            breakdown[data_col] = breakdown[data_col].apply(_bytes_to_mb)
            breakdown[dict_col] = breakdown[dict_col].apply(_bytes_to_mb)
            breakdown = breakdown.rename(
                columns={data_col: "Data Size (MB)", dict_col: "Dictionary Size (MB)"}
            ).reset_index(drop=True)
            summary["dict_vs_data"] = breakdown

    encoding_col = _name_column(columns_df, *_VPA_ENCODING_CANDIDATES)
    if encoding_col:
        counts = columns_df[encoding_col].dropna()
        if not counts.empty:
            summary["available"] = True
            vc = counts.value_counts().reset_index()
            vc.columns = ["Encoding", "Column Count"]
            summary["encoding_breakdown"] = vc

    return summary


# ==========================================================================
# DAX tokenizer (shared infrastructure for every DAX-aware feature)
# ==========================================================================

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


def find_referenced_measures(expr: str, measure_names: Set[str]) -> Set[str]:
    """Measure names a DAX expression calls, e.g. `[Total Sales] * 1.1`.

    A `[Ref]` token is a measure reference when its bare name is in
    `measure_names` - the same test `find_referenced_columns` below uses to
    rule measure references *out* when resolving column references.
    """
    if not expr:
        return set()
    found = set()
    for kind, text in _tokenize_dax(expr):
        if kind != "ref":
            continue
        name = text[1:-1]
        if name in measure_names:
            found.add(name)
    return found


def find_referenced_columns(
    expr: str, column_tables: Dict[str, Set[str]], measure_names: Set[str]
) -> Set[Tuple[str, str]]:
    """(table, column) pairs a DAX expression unambiguously references.

    A bare `[Ref]` token could be a column or a measure name; only tokens
    that resolve to exactly one owning table are counted, since an
    ambiguous name can't be attributed to a specific table without
    inspecting how it's qualified at each use site.
    """
    if not expr:
        return set()
    found: Set[Tuple[str, str]] = set()
    for kind, text in _tokenize_dax(expr):
        if kind != "ref":
            continue
        name = text[1:-1]
        if name in measure_names:
            continue
        owners = column_tables.get(name, set())
        if len(owners) == 1:
            found.add((next(iter(owners)), name))
    return found


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
# Model analysis: unused objects, impact analysis, measure dependencies
# ==========================================================================

def referenced_columns(model: Dict[str, Any]) -> Set[Tuple[str, str]]:
    """(table, column) pairs referenced by any measure, calc column, or relationship.

    A static reference scan: it can't see Power BI report visuals (a .vpax
    carries none) or RLS filter expressions - fold `model["roles"]`'s filter
    expressions in too if those should also count as "used".
    """
    column_tables = model["column_tables"]
    measure_names = model["measure_names"]
    found: Set[Tuple[str, str]] = set()

    for expr in model["measures"]["MeasureExpression"]:
        found |= find_referenced_columns(str(expr or ""), column_tables, measure_names)
    for expr in model["calc_columns"]["ColumnExpression"]:
        found |= find_referenced_columns(str(expr or ""), column_tables, measure_names)

    rel_df = model["relationships"]
    for _, r in rel_df.iterrows():
        if r["From Table"] and r["From Column"]:
            found.add((r["From Table"], r["From Column"]))
        if r["To Table"] and r["To Column"]:
            found.add((r["To Table"], r["To Column"]))

    return found


def find_unused_columns(model: Dict[str, Any]) -> pd.DataFrame:
    """Columns no measure, calculated column, or relationship references.

    A column name shared by more than one table can't be attributed to a
    specific table from an unqualified [Ref] alone, so those are reported as
    ambiguous rather than silently marked used or unused. VertiPaq's internal
    RowNumber pseudo-columns are excluded - they're never real modeling
    objects, so "unused" doesn't mean anything for them.
    """
    columns_df = _user_facing_columns(model["columns"])
    if columns_df.empty or "TableName" not in columns_df.columns or "ColumnName" not in columns_df.columns:
        return _ensure_columns(pd.DataFrame(), ["Table", "Column", "Status", "Severity"])

    used = referenced_columns(model)
    used_names = {c for _, c in used}
    column_tables = model["column_tables"]

    rows = []
    for _, row in columns_df[["TableName", "ColumnName"]].dropna().iterrows():
        table, col = row["TableName"], row["ColumnName"]
        if (table, col) in used:
            # Confirmed for this exact table, e.g. via a relationship
            # endpoint - never downgrade this to "ambiguous".
            status = "Referenced"
        elif len(column_tables.get(col, set())) > 1 and col in used_names:
            status = "Referenced somewhere (table ambiguous)"
        else:
            status = "Likely unused"
        # Low, not Medium: this scan can't see report visuals, so a "likely
        # unused" column is a lead to investigate, never a safe-to-delete
        # verdict. Ambiguous rows are Info so they don't inflate the score.
        severity = {"Referenced": "Info",
                    "Referenced somewhere (table ambiguous)": "Info",
                    "Likely unused": "Low"}[status]
        rows.append({"Table": table, "Column": col, "Status": status, "Severity": severity})

    return _ensure_columns(pd.DataFrame(rows), ["Table", "Column", "Status", "Severity"])


def _column_ref_matcher(table: str, column: str) -> "re.Pattern[str]":
    """Match an explicit, table-qualified reference to one exact column.

    Column names are frequently reused across tables (e.g. a fact and a
    dimension both having a "Key" column), so matching by column name alone
    can't tell which table a bare `[Key]` means. Only the qualified forms
    (`'Table'[Column]` or `Table[Column]`) are counted - deliberately
    conservative, since an unqualified reference elsewhere can't be safely
    attributed to one table without a real DAX engine resolving row context.
    """
    t, c = re.escape(table), re.escape(column)
    parts = [rf"'{t}'\s*\[{c}\]"]
    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", table):
        parts.append(rf"(?<![A-Za-z0-9_]){t}\s*\[{c}\]")
    return re.compile("|".join(parts))


def impact_of(
    target_kind: str, target_name: str, model: Dict[str, Any], table_name: Optional[str] = None
) -> Dict[str, pd.DataFrame]:
    """Everything that references a table or column, for pre-change review.

    `target_kind` is "table" or "column". For "column", `table_name` pins
    down *which* table's column is meant, since column names aren't unique
    across tables - see `_column_ref_matcher`.
    """
    measures_df, calc_df, rel_df = model["measures"], model["calc_columns"], model["relationships"]
    table_names = model["all_table_names"]

    if target_kind == "table":
        def touches(expr: Any) -> bool:
            expr = str(expr or "")
            return bool(expr) and target_name in find_referenced_tables(expr, table_names)

        rel_hits = rel_df[
            (rel_df["From Table"] == target_name) | (rel_df["To Table"] == target_name)
        ].reset_index(drop=True)
        neighbours = sorted(_adjacency(rel_df).get(target_name, set()))
    else:
        pattern = _column_ref_matcher(table_name or "", target_name)

        def touches(expr: Any) -> bool:
            expr = str(expr or "")
            return bool(expr) and bool(pattern.search(expr))

        rel_hits = rel_df[
            ((rel_df["From Table"] == table_name) & (rel_df["From Column"] == target_name))
            | ((rel_df["To Table"] == table_name) & (rel_df["To Column"] == target_name))
        ].reset_index(drop=True)
        neighbours = []

    measure_hits = measures_df[measures_df["MeasureExpression"].apply(touches)].reset_index(drop=True)
    calc_hits = calc_df[calc_df["ColumnExpression"].apply(touches)].reset_index(drop=True)

    related_df = pd.DataFrame({"Related Table": neighbours}) if neighbours else _ensure_columns(
        pd.DataFrame(), ["Related Table"]
    )

    return {
        "measures": measure_hits,
        "calc_columns": calc_hits,
        "relationships": rel_hits,
        "related_tables": related_df,
    }


def build_measure_graph(model: Dict[str, Any]) -> Dict[str, Set[str]]:
    """{measure name: {measures it calls}}, built from every measure's DAX."""
    measure_names = model["measure_names"]
    graph: Dict[str, Set[str]] = {name: set() for name in measure_names}
    meas_df = model["measures"]
    if "MeasureName" not in meas_df.columns:
        return graph
    for _, row in meas_df[["MeasureName", "MeasureExpression"]].dropna(subset=["MeasureName"]).iterrows():
        name = row["MeasureName"]
        calls = find_referenced_measures(str(row["MeasureExpression"] or ""), measure_names) - {name}
        graph.setdefault(name, set())
        graph[name] |= calls
    return graph


def find_cycles(graph: Dict[str, Set[str]]) -> List[List[str]]:
    """Circular measure-reference chains, via DFS with a recursion stack.

    A measure that (directly or transitively) calls itself can never
    evaluate - always a modeling bug worth surfacing loudly.
    """
    WHITE, GRAY, BLACK = 0, 1, 2
    color: Dict[str, int] = {n: WHITE for n in graph}
    stack: List[str] = []
    cycles: List[List[str]] = []

    def visit(node: str) -> None:
        color[node] = GRAY
        stack.append(node)
        for nxt in graph.get(node, set()):
            state = color.get(nxt, WHITE)
            if state == WHITE:
                visit(nxt)
            elif state == GRAY:
                idx = stack.index(nxt)
                cycles.append(stack[idx:] + [nxt])
        stack.pop()
        color[node] = BLACK

    for node in list(graph):
        if color.get(node, WHITE) == WHITE:
            visit(node)
    return cycles


def build_measure_dependency_table(graph: Dict[str, Set[str]]) -> pd.DataFrame:
    """One row per measure: what it calls, and what calls it back.

    This is the plain "which measure is used to build which" answer - the
    diagram is a visual on top of the same data, not a replacement for it.
    """
    used_by: Dict[str, Set[str]] = {n: set() for n in graph}
    for name, calls in graph.items():
        for called in calls:
            used_by.setdefault(called, set()).add(name)

    rows = []
    for name in sorted(graph):
        rows.append({
            "Measure": name,
            "Depends On": ", ".join(sorted(graph.get(name, set()))),
            "Used By": ", ".join(sorted(used_by.get(name, set()))),
        })
    return _ensure_columns(pd.DataFrame(rows), ["Measure", "Depends On", "Used By"])


def build_measure_dependency_dot(graph: Dict[str, Set[str]], focus: Optional[str] = None) -> str:
    """DOT for a directed measure-calls-measure graph.

    Same colour language as the ER diagrams (`#1f3a5f` primary / `#2d6ca8`
    secondary / `#64748b` neutral) - just simple labeled nodes/edges, since
    a measure graph has no crow's-foot cardinality to draw.
    """
    if focus:
        keep = {focus} | graph.get(focus, set())
        keep |= {n for n, calls in graph.items() if focus in calls}
        nodes = keep
        edges = [(a, b) for a, calls in graph.items() if a in keep for b in calls if b in keep]
    else:
        nodes = set(graph)
        edges = [(a, b) for a, calls in graph.items() for b in calls]

    cyclic_edges = set()
    for cycle in find_cycles(graph):
        for a, b in zip(cycle, cycle[1:]):
            cyclic_edges.add((a, b))

    lines = [
        "digraph G {",
        'rankdir="LR"; bgcolor="#ffffff";',
        'node [shape=box, style="rounded,filled", fontname="Helvetica", fontsize=11, '
        'fillcolor="#2d6ca8", fontcolor="#ffffff", color="#1f3a5f"];',
        'edge [color="#64748b", fontname="Helvetica", fontsize=9];',
    ]
    for n in sorted(nodes):
        fill = "#1f3a5f" if n == focus else "#2d6ca8"
        lines.append(f"{_dot_id(n)} [label={_dot_id(n)}, fillcolor=\"{fill}\"];")
    for a, b in edges:
        colour = "#dc2626" if (a, b) in cyclic_edges else "#64748b"
        lines.append(f"{_dot_id(a)} -> {_dot_id(b)} [color=\"{colour}\"];")
    lines.append("}")
    return "\n".join(lines)


# ==========================================================================
# Model health checks + naming lint
# ==========================================================================

_ROW_CONTEXT_FUNCTIONS = (
    "CALCULATE", "RELATED", "RELATEDTABLE", "EARLIER", "EARLIEST",
    "FILTER", "ALL", "ALLEXCEPT", "USERELATIONSHIP",
)


def run_health_checks(model: Dict[str, Any]) -> pd.DataFrame:
    """Rules-engine checklist: {Rule, Object, Severity, Message, How to Fix}.

    Every rule only fires when the data it needs is actually present in
    this model export - e.g. the cardinality rule is skipped (not guessed)
    when no VertiPaq Cardinality field was captured. `Object` is the where
    (the exact table/column/measure); `Message` is the what-and-why;
    `How to Fix` is the concrete action - kept as three separate columns so
    none of the three gets lost inside a wall of prose.
    """
    rows: List[Dict[str, str]] = []

    def add(rule: str, obj: str, severity: str, message: str, fix: str) -> None:
        rows.append({"Rule": rule, "Object": obj, "Severity": severity,
                     "Message": message, "How to Fix": fix})

    meas_df_all = model["measures"]
    if "MeasureName" in meas_df_all.columns:
        names = meas_df_all["MeasureName"].dropna().astype(str)
        for name in sorted(names[names.duplicated(keep=False)].unique()):
            add(
                "Duplicate measure name", name, "High",
                "Another measure elsewhere in the model shares this exact name — likely "
                "copy-pasted logic under an identical name instead of reused, and confusing "
                "for anyone trying to report against it.",
                "Check Investigate ➜ Impact Analysis on this name first to see what each copy "
                "actually feeds, then rename or delete the redundant one and repoint anything "
                "that referenced it.",
            )

    measure_graph = build_measure_graph(model)
    for cycle in find_cycles(measure_graph):
        add(
            "Circular measure reference", " → ".join(cycle), "High",
            "These measures call each other in a loop and can never fully evaluate — this is a "
            "correctness bug, not a style preference.",
            "Open each measure in this chain (Explore ➜ Measures) and rewrite one of them so it "
            "no longer calls a measure that, directly or through others, calls it back — usually "
            "by inlining that dependency's logic or basing it on a shared base measure instead.",
        )

    rel_df = model["relationships"]
    for _, r in rel_df[rel_df["Cross Filter Direction"] == "Both"].iterrows():
        add(
            "Bi-directional relationship", f'{r["From Table"]} ↔ {r["To Table"]}', "Medium",
            "Bi-directional filtering can cause ambiguous or double-counted results — confirm it's intentional.",
            "In Power BI Desktop: Model view ➜ select the relationship ➜ set Cross filter "
            "direction to Single, unless a specific measure genuinely needs the reverse filter — "
            "in which case use CROSSFILTER() inside just that measure instead of leaving the "
            "relationship bi-directional for the whole model.",
        )

    for _, r in rel_df[(rel_df["From Cardinality"] == "Many") & (rel_df["To Cardinality"] == "Many")].iterrows():
        add(
            "Many-to-many relationship", f'{r["From Table"]} ↔ {r["To Table"]}', "Medium",
            "Both sides are Many, so the engine can't use this join directly — it resolves the "
            "filter by building a large temporary table at query time, which gets expensive as "
            "either table grows.",
            "Where possible, replace it with a bridge (dimension) table related One-to-Many to "
            "both sides. If a true many-to-many is unavoidable, keep both tables' row counts "
            "small or resolve the relationship inside the measure with TREATAS() instead of "
            "relying on the model relationship for every query.",
        )

    calc_df = model["calc_columns"]
    for _, r in calc_df.iterrows():
        expr = str(r.get("ColumnExpression") or "").upper()
        if expr and not any(fn in expr for fn in _ROW_CONTEXT_FUNCTIONS):
            add(
                "Calculated column could move upstream", f'{r["TableName"]}[{r["ColumnName"]}]', "Low",
                "No row-context/relationship functions detected — this may be cheaper to compute in Power Query/SQL.",
                "Recreate the same logic as a Power Query step (Add Column) or in the source SQL "
                "view, then delete the calculated column — moving it upstream computes it once at "
                "refresh instead of once per row at query time.",
            )

    columns_df = _user_facing_columns(model["columns"])
    cardinality_col = _name_column(columns_df, *_VPA_CARDINALITY_CANDIDATES)
    if cardinality_col and {"TableName", "ColumnName"}.issubset(columns_df.columns):
        key_pairs = set()
        for _, r in rel_df.iterrows():
            key_pairs.add((r["From Table"], r["From Column"]))
            key_pairs.add((r["To Table"], r["To Column"]))
        card_lookup = columns_df.dropna(subset=["TableName", "ColumnName"]).set_index(["TableName", "ColumnName"])[cardinality_col]
        for table, col in key_pairs:
            if (table, col) not in card_lookup.index:
                continue
            val = card_lookup.loc[(table, col)]
            if isinstance(val, pd.Series):
                val = val.iloc[0]
            val = pd.to_numeric(val, errors="coerce")
            if pd.notna(val) and val > 1_000_000:
                add(
                    "High-cardinality relationship key", f"{table}[{col}]", "High",
                    f"Cardinality ~{int(val):,} — large key columns increase model size and slow joins.",
                    "Replace the natural key with a smaller surrogate integer key generated at "
                    "load time, or split it (e.g. date/time apart) so each relationship side has "
                    "fewer distinct values — see Audit ➜ Compression Advisor for the same column.",
                )
    else:
        add(
            "High-cardinality relationship key", "(model-wide)", "Info",
            "Skipped — this .vpax export doesn't include a VertiPaq Cardinality field.",
            "Re-export with a tool that captures VertiPaq stats (DAX Studio's Advanced ➜ Export "
            "Metadata) to get this check.",
        )

    tables_df = model["tables"]
    table_name_col = _name_column(tables_df, "TableName", "Name")
    if "Description" in tables_df.columns and table_name_col:
        for _, r in tables_df.iterrows():
            if not str(r.get("Description") or "").strip():
                add("Missing description", str(r[table_name_col]), "Low", "Table has no description.",
                    "Tabular Editor: select the table ➜ Properties pane ➜ Description. In Power "
                    "BI Desktop: Model view ➜ select the table ➜ Properties ➜ Description.")
    if "Description" in columns_df.columns and {"TableName", "ColumnName"}.issubset(columns_df.columns):
        for _, r in columns_df.iterrows():
            if not str(r.get("Description") or "").strip():
                add("Missing description", f'{r["TableName"]}[{r["ColumnName"]}]', "Low", "Column has no description.",
                    "Tabular Editor: select the column ➜ Properties pane ➜ Description. In Power "
                    "BI Desktop: Model view ➜ select the column ➜ Properties ➜ Description.")
    meas_df = model["measures"]
    if "Description" in meas_df.columns and "MeasureName" in meas_df.columns:
        for _, r in meas_df.iterrows():
            if not str(r.get("Description") or "").strip():
                add("Missing description", str(r["MeasureName"]), "Low", "Measure has no description.",
                    "Tabular Editor: select the measure ➜ Properties pane ➜ Description. In Power "
                    "BI Desktop: Model view ➜ select the measure ➜ Properties ➜ Description. "
                    "Govern ➜ Fix Script (C#) can generate placeholder stubs for all of these.")

    if {"FormatString", "DataType"}.issubset(meas_df.columns):
        for dtype, grp in meas_df.dropna(subset=["DataType"]).groupby("DataType"):
            fmt_values = grp["FormatString"].dropna().astype(str)
            fmt_values = fmt_values[fmt_values.str.strip() != ""]
            distinct = sorted(fmt_values.unique())
            if len(distinct) > 1:
                shown = ", ".join(distinct[:5]) + (" …" if len(distinct) > 5 else "")
                add(
                    "Inconsistent format strings", str(dtype), "Low",
                    f"{len(distinct)} different format strings used for {dtype} measures: {shown}",
                    "Pick one FormatString for this data type and apply it to every measure of "
                    "that type — in Tabular Editor, multi-select the measures and set Format "
                    "String once; in Power BI Desktop, select each measure under Measure Tools "
                    "➜ Formatting.",
                )

    return _ensure_columns(
        pd.DataFrame(rows), ["Rule", "Object", "Severity", "Message", "How to Fix"]
    )


_CASING_PATTERNS = [
    ("snake_case", re.compile(r"^[a-z][a-z0-9]*(_[a-z0-9]+)+$")),
    ("PascalCase", re.compile(r"^[A-Z][a-zA-Z0-9]*$")),
    ("camelCase", re.compile(r"^[a-z][a-zA-Z0-9]*$")),
    ("Title Case With Spaces", re.compile(r"^[A-Z][a-z0-9]*(\s[A-Za-z0-9]+)*$")),
]


def _detect_casing(name: str) -> str:
    for label, pattern in _CASING_PATTERNS:
        if pattern.fullmatch(name):
            return label
    return "Other/mixed"


def _split_words(name: str) -> List[str]:
    """Break a name into words regardless of its current casing convention."""
    s = re.sub(r"[_\-]+", " ", name)
    s = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", s)  # camelCase/PascalCase boundary
    s = re.sub(r"(?<=[A-Za-z])(?=[0-9])", " ", s)
    return [w for w in s.split() if w]


def _convert_casing(name: str, convention: str) -> Optional[str]:
    """Render `name` in the given convention, e.g. 'Sale Amt' -> 'sale_amt'."""
    words = _split_words(name)
    if not words:
        return None
    if convention == "snake_case":
        return "_".join(w.lower() for w in words)
    if convention == "PascalCase":
        return "".join(w.capitalize() for w in words)
    if convention == "camelCase":
        return words[0].lower() + "".join(w.capitalize() for w in words[1:])
    if convention == "Title Case With Spaces":
        return " ".join(w.capitalize() for w in words)
    return None


def lint_naming(model: Dict[str, Any]) -> pd.DataFrame:
    """Flag inconsistent casing across measures, tables, and columns.

    Pure string heuristics, no DAX parsing. A convention is only flagged as
    inconsistent within its own group (e.g. a table's own columns) since
    different tables in the same model often use different, internally
    consistent conventions. VertiPaq's internal RowNumber pseudo-columns are
    excluded - their auto-generated GUID names aren't a naming decision.
    """
    rows: List[Dict[str, str]] = []

    def scan(obj_type: str, names_by_group: Dict[str, List[str]]) -> None:
        for group, names in names_by_group.items():
            if len(names) < 2:
                continue
            conventions = [_detect_casing(n) for n in names]
            dominant = pd.Series(conventions).value_counts().index[0]
            for name, conv in zip(names, conventions):
                if conv == dominant:
                    continue
                suggested = _convert_casing(name, dominant)
                if suggested and suggested != name:
                    suggestion = (
                        f"Most {obj_type.lower()}s on {group} use {dominant} — "
                        f"consider renaming to `{suggested}`."
                    )
                else:
                    suggestion = f"Most {obj_type.lower()}s on {group} use {dominant} — consider matching it."
                rows.append({
                    "Object Type": obj_type, "Table": group, "Name": name,
                    # Naming is presentation, never correctness - a mismatch is
                    # worth fixing but can't produce a wrong number, so it sits
                    # at Low in the shared severity vocabulary. Measures are
                    # nudged to Medium because report authors see those names.
                    "Severity": "Medium" if obj_type == "Measure" else "Low",
                    "Detected Convention": conv, "Suggestion": suggestion,
                })

    tables_df = model["tables"]
    table_name_col = _name_column(tables_df, "TableName", "Name")
    if table_name_col:
        names = tables_df[table_name_col].dropna().astype(str).tolist()
        scan("Table", {"(model)": names})

    meas_df = model["measures"]
    if {"TableName", "MeasureName"}.issubset(meas_df.columns):
        groups: Dict[str, List[str]] = {}
        for _, r in meas_df.dropna(subset=["MeasureName"]).iterrows():
            groups.setdefault(str(r["TableName"]), []).append(str(r["MeasureName"]))
        scan("Measure", groups)

    columns_df = _user_facing_columns(model["columns"])
    if {"TableName", "ColumnName"}.issubset(columns_df.columns):
        groups = {}
        for _, r in columns_df.dropna(subset=["ColumnName"]).iterrows():
            groups.setdefault(str(r["TableName"]), []).append(str(r["ColumnName"]))
        scan("Column", groups)

    return sort_by_severity(_ensure_columns(
        pd.DataFrame(rows),
        ["Object Type", "Table", "Name", "Severity", "Detected Convention", "Suggestion"],
    ))


# ==========================================================================
# Severity taxonomy
# ==========================================================================
# One vocabulary for every check in the app, so "High" means the same thing
# in Model Health, Fabric Readiness and the Compression Advisor, and the
# scorecard can roll them all up without special-casing each module.
SEVERITY_ORDER: Dict[str, int] = {"High": 0, "Medium": 1, "Low": 2, "Info": 3}
# Penalty applied to the 100-point model score per finding, per severity.
SEVERITY_ICON: Dict[str, str] = {"High": "🔴", "Medium": "🟠", "Low": "⚪", "Info": "🔵", "Clean": "🟢"}


def sort_by_severity(df: pd.DataFrame, column: str = "Severity") -> pd.DataFrame:
    """Highest-severity rows first, so the important findings are above the fold."""
    if df.empty or column not in df.columns:
        return df
    out = df.copy()
    out["_sev"] = out[column].map(lambda s: SEVERITY_ORDER.get(str(s), 9))
    out = out.sort_values("_sev", kind="stable").drop(columns=["_sev"])
    return out.reset_index(drop=True)


def severity_counts(df: pd.DataFrame, column: str = "Severity") -> Dict[str, int]:
    if df.empty or column not in df.columns:
        return {}
    return {str(k): int(v) for k, v in df[column].value_counts().items()}


# ==========================================================================
# VertiPaq compression & encoding advisor
# ==========================================================================

def _estimated_distinct(bits: Any) -> Optional[int]:
    """Rough distinct-value count from a segment's dictionary bit width.

    VertiPaq allocates just enough bits per value to address the column's
    dictionary, so 2^BitsCount is an upper bound on cardinality. It is an
    estimate, not a COUNTROWS(DISTINCT()) - a .vpax carries no row data - but
    it is the only cardinality signal in the export and is accurate enough to
    separate a 12-value status flag from a 20-million-value transaction ID.
    """
    try:
        b = int(bits)
    except (TypeError, ValueError):
        return None
    if b <= 0 or b > 40:
        return None
    return 2 ** b


def build_encoding_advice(model: Dict[str, Any]) -> pd.DataFrame:
    """Per-column storage advice from VertiPaq encoding and segment stats.

    The wins this looks for, in the order they usually pay off:
      * a big HASH-encoded numeric column - VALUE encoding is cheaper and is
        chosen automatically once the column is a clean integer type;
      * a big column whose segments are NOSPLIT (i.e. not run-length encoded)
        but which has few distinct values - sorting the source by that column
        during load lets RLE collapse long runs;
      * a dictionary that dwarfs the data itself, which almost always means a
        high-cardinality text or decimal column that could be split or rounded;
      * IsAvailableInMDX left on for hidden columns, which builds an attribute
        hierarchy nobody queries.
    """
    cols = _user_facing_columns(model["columns"])
    schema = ["Table", "Column", "Encoding", "Compression", "Total Size (MB)",
              "Dictionary %", "Max Distinct (est.)", "Severity", "Why", "How to Fix"]
    if cols.empty or "TotalSize" not in cols.columns:
        return _ensure_columns(pd.DataFrame(), schema)

    # Row counts bound the estimate: a column can't have more distinct values
    # than its table has rows, and 2^BitsCount alone routinely overshoots that
    # by orders of magnitude.
    rows_by_table: Dict[str, float] = {}
    tdf = model["tables"]
    if not tdf.empty and {"TableName", "RowsCount"}.issubset(tdf.columns):
        for _, t in tdf.iterrows():
            n = pd.to_numeric(pd.Series([t.get("RowsCount")]), errors="coerce").iloc[0]
            if pd.notna(n):
                rows_by_table[str(t["TableName"])] = float(n)

    # Segment-level compression + bit width, collapsed to one row per column.
    seg = model.get("segments")
    comp_by_col: Dict[Tuple[str, str], str] = {}
    bits_by_col: Dict[Tuple[str, str], Any] = {}
    if isinstance(seg, pd.DataFrame) and not seg.empty and {"TableName", "ColumnName"}.issubset(seg.columns):
        for _, s in seg.iterrows():
            key = (str(s.get("TableName")), str(s.get("ColumnName")))
            if "CompressionType" in seg.columns and s.get("CompressionType"):
                comp_by_col.setdefault(key, str(s.get("CompressionType")))
            if "BitsCount" in seg.columns and pd.notna(s.get("BitsCount")):
                bits_by_col[key] = max(bits_by_col.get(key, 0) or 0, s.get("BitsCount"))

    total_model_size = pd.to_numeric(cols["TotalSize"], errors="coerce").fillna(0).sum()
    # Only bother advising on columns big enough to matter: 0.5% of the model
    # or 1 MB, whichever is smaller. Micro-optimising a 4 KB column is noise.
    threshold = min(total_model_size * 0.005, 1_000_000) if total_model_size else 0

    rows: List[Dict[str, Any]] = []
    for _, c in cols.iterrows():
        table, col = str(c.get("TableName")), str(c.get("ColumnName"))
        total = pd.to_numeric(pd.Series([c.get("TotalSize")]), errors="coerce").fillna(0).iloc[0]
        if total < threshold:
            continue
        dict_size = pd.to_numeric(pd.Series([c.get("DictionarySize")]), errors="coerce").fillna(0).iloc[0]
        dict_pct = (dict_size / total * 100) if total else 0.0
        encoding = str(c.get("Encoding") or "").upper()
        dtype = str(c.get("DataType") or "")
        compression = comp_by_col.get((table, col), "")
        est_distinct = _estimated_distinct(bits_by_col.get((table, col)))
        row_count = rows_by_table.get(table)
        if est_distinct and row_count:
            est_distinct = int(min(est_distinct, row_count))
        hidden = bool(c.get("IsHidden"))
        in_mdx = bool(c.get("IsAvailableInMDX"))

        whys: List[str] = []
        hows: List[str] = []
        severities: List[str] = []

        def flag(sev: str, why: str, how: str) -> None:
            severities.append(sev)
            whys.append(why)
            hows.append(how)

        # "Big" here means big relative to this model, so the advisor stays
        # useful on a 20 MB model and doesn't drown you on a 20 GB one.
        very_big = total > threshold * 4

        numeric = dtype.lower() in ("int64", "double", "decimal", "currency")
        if encoding == "HASH" and numeric:
            flag(
                "High" if dict_pct > 40 else "Medium",
                "Numeric column stored with HASH encoding — it carries a dictionary it doesn't need.",
                "Tabular Editor: select the column ➜ set `EncodingHint = Value` (or clean the "
                "source so the column is a true integer type) — this drops the dictionary entirely.",
            )

        if compression.upper() == "NOSPLIT" and est_distinct and est_distinct <= 1024:
            flag(
                "High" if very_big else "Medium",
                f"At most ~{est_distinct:,} distinct values but segments are NOSPLIT, so "
                "run-length encoding isn't kicking in.",
                "Add an ORDER BY on this column to the source query (or `Table.Sort` in Power "
                "Query M) — sorting by the lowest-cardinality column first lets RLE collapse the "
                "long repeated runs.",
            )

        if dict_pct > 60:
            flag(
                "High" if very_big else "Medium",
                f"The dictionary is {dict_pct:.0f}% of this column's footprint — classic "
                "high-cardinality text or high-precision decimal.",
                "Split the column (e.g. date apart from time, prefix apart from suffix) or round "
                "the decimal to the precision actually reported on, either upstream in Power "
                "Query/SQL or as a replacement calculated column.",
            )

        if hidden and in_mdx:
            flag(
                "Medium",
                "Hidden but still `IsAvailableInMDX = true`, so VertiPaq builds an attribute "
                "hierarchy no report can use.",
                "Tabular Editor: select the column ➜ set `IsAvailableInMdx = false` — or generate "
                "this automatically from Govern ➜ Fix Script (C#).",
            )

        if not whys:
            continue
        severity = min(severities, key=lambda s: SEVERITY_ORDER.get(s, 9))
        rows.append({
            "Table": table, "Column": col,
            "Encoding": encoding or "—",
            "Compression": compression or "—",
            "Total Size (MB)": round(_bytes_to_mb(total), 3),
            "Dictionary %": round(dict_pct, 1),
            "Max Distinct (est.)": est_distinct if est_distinct else "—",
            "Severity": severity,
            "Why": " ".join(whys),
            "How to Fix": " ".join(hows),
        })

    df = _ensure_columns(pd.DataFrame(rows), schema)
    if df.empty:
        return df
    return sort_by_severity(df.sort_values("Total Size (MB)", ascending=False))


# ==========================================================================
# Microsoft Fabric / Direct Lake readiness
# ==========================================================================

def build_fabric_readiness(model: Dict[str, Any]) -> pd.DataFrame:
    """Check what would block this model from running in Direct Lake mode.

    Direct Lake reads Delta/Parquet straight out of OneLake, which means the
    features that need an import-engine transform step aren't available. When
    one of them is present the query falls back to DirectQuery, and the
    performance advantage disappears silently. These are the blockers that
    actually appear in .vpax metadata - it can't see everything (e.g. whether
    the Lakehouse tables are V-Order optimised), so the tab says so explicitly.
    """
    schema = ["Check", "Object", "Severity", "Finding", "Fix"]
    rows: List[Dict[str, Any]] = []
    bim_tables = model.get("bim_tables") or []

    calc_tables, calc_cols, m_tables = [], [], []
    for t in bim_tables:
        name = t.get("name")
        if not name:
            continue
        for p in t.get("partitions") or []:
            if not isinstance(p, dict):
                continue
            src = p.get("source") or {}
            stype = str((src or {}).get("type") or "").lower()
            if stype == "calculated":
                calc_tables.append(name)
            elif stype == "m":
                m_tables.append(name)
        for c in t.get("columns") or []:
            if isinstance(c, dict) and str(c.get("type") or "").lower() == "calculated":
                calc_cols.append(f"{name}[{c.get('name')}]")

    if calc_tables:
        rows.append({
            "Check": "Calculated tables", "Object": ", ".join(sorted(set(calc_tables))[:12]),
            "Severity": "High",
            "Finding": f"{len(set(calc_tables))} calculated table(s). Direct Lake has no DAX "
                       "engine at load time, so these are not supported.",
            "Fix": "Materialise them upstream as Lakehouse/Warehouse tables (a notebook, "
                   "dataflow, or SQL view) and import them as regular Delta tables.",
        })
    if calc_cols:
        rows.append({
            "Check": "Calculated columns", "Object": ", ".join(sorted(set(calc_cols))[:12]),
            "Severity": "High",
            "Finding": f"{len(set(calc_cols))} calculated column(s). Not supported in Direct "
                       "Lake — the whole table falls back to DirectQuery.",
            "Fix": "Compute them in the Delta table itself (Spark/SQL) so they arrive as "
                   "physical columns.",
        })
    if m_tables:
        rows.append({
            "Check": "Power Query (M) transforms", "Object": f"{len(set(m_tables))} table(s)",
            "Severity": "Medium",
            "Finding": "Tables load through M. Direct Lake tables must point at a Delta table "
                       "with no transform step, so any non-trivial M here has to move upstream.",
            "Fix": "Push the transform into the Lakehouse (dataflow Gen2 or a notebook) and "
                   "leave the semantic model as a thin passthrough.",
        })

    # Auto date/time tables are import-only and quietly bloat the model.
    auto_date = [t.get("name") for t in bim_tables
                 if str(t.get("name") or "").startswith(("LocalDateTable_", "DateTableTemplate_"))]
    if auto_date:
        rows.append({
            "Check": "Auto date/time tables", "Object": f"{len(auto_date)} hidden table(s)",
            "Severity": "High",
            "Finding": "Power BI's automatic date/time is on. It generates one hidden date "
                       "table per date column, is unsupported in Direct Lake, and inflates "
                       "model size in import mode too.",
            "Fix": "File ➜ Options ➜ Current File ➜ Data Load ➜ untick *Auto date/time*, "
                   "then use one shared, marked date table.",
        })

    # Memory guardrails: F-SKUs cap the model, and high-cardinality columns
    # plus attribute hierarchies are what push a model over the line.
    cols = _user_facing_columns(model["columns"])
    if not cols.empty and "TotalSize" in cols.columns:
        total_bytes = pd.to_numeric(cols["TotalSize"], errors="coerce").fillna(0).sum()
        total_gb = total_bytes / (1024 ** 3)
        if total_gb > 0:
            rows.append({
                "Check": "F-SKU memory guardrail", "Object": "Whole model",
                "Severity": "High" if total_gb > 25 else ("Medium" if total_gb > 3 else "Info"),
                "Finding": f"Column footprint is roughly {total_gb:.2f} GB. Direct Lake keeps "
                           "the columns it touches resident in memory, and each F-SKU has a "
                           "hard per-model limit (F64 ≈ 25 GB, F2 ≈ 3 GB).",
                "Fix": "Drop unused columns, reduce decimal precision, and check the "
                       "Compression Advisor before sizing the capacity.",
            })

        if "IsAvailableInMDX" in cols.columns and "IsHidden" in cols.columns:
            wasted = cols[(cols["IsHidden"] == True) & (cols["IsAvailableInMDX"] == True)]  # noqa: E712
            if not wasted.empty:
                rows.append({
                    "Check": "Unneeded attribute hierarchies",
                    "Object": f"{len(wasted)} hidden column(s)",
                    "Severity": "Medium",
                    "Finding": "Hidden columns still have `IsAvailableInMDX = true`, so an "
                               "attribute hierarchy is built and kept in memory for columns "
                               "no report can browse.",
                    "Fix": "Set IsAvailableInMDX to false on hidden columns — the Fix Script "
                           "tab generates this for you.",
                })

    # String-typed, high-cardinality relationship keys bloat the VertiPaq
    # dictionary and are a common trigger for breaching F-SKU guardrails -
    # Fabric's own guidance is to replace them with integer surrogate keys.
    cardinality_col = _name_column(cols, *_VPA_CARDINALITY_CANDIDATES)
    if not cols.empty and cardinality_col and "DataType" in cols.columns and {"TableName", "ColumnName"}.issubset(cols.columns):
        key_pairs = set()
        for _, r in model["relationships"].iterrows():
            key_pairs.add((r["From Table"], r["From Column"]))
            key_pairs.add((r["To Table"], r["To Column"]))
        lookup = cols.dropna(subset=["TableName", "ColumnName"]).set_index(["TableName", "ColumnName"])
        string_keys = []
        for table, col in key_pairs:
            if (table, col) not in lookup.index:
                continue
            row = lookup.loc[(table, col)]
            if isinstance(row, pd.DataFrame):
                row = row.iloc[0]
            card = pd.to_numeric(row.get(cardinality_col), errors="coerce")
            if str(row.get("DataType")) == "String" and pd.notna(card) and card > 100_000:
                string_keys.append(f"{table}[{col}] (~{int(card):,})")
        if string_keys:
            rows.append({
                "Check": "String-typed surrogate key", "Object": ", ".join(sorted(string_keys)[:8]),
                "Severity": "Medium",
                "Finding": f"{len(string_keys)} relationship key(s) are high-cardinality text. "
                           "Text keys build a much larger dictionary than an equivalent integer "
                           "key and are a common reason a model breaches its F-SKU guardrail.",
                "Fix": "Generate a surrogate BIGINT key upstream (identity column or hashed-to-"
                       "int in the ETL) and relate on that instead of the natural text key.",
            })

    # Direct Lake can't do bi-di relationships against a fallback-free model
    # reliably, and they're a correctness risk regardless.
    rels = model["relationships"]
    if not rels.empty and "Cross Filter Direction" in rels.columns:
        bidi = rels[rels["Cross Filter Direction"] == "Both"]
        if not bidi.empty:
            rows.append({
                "Check": "Bi-directional relationships",
                "Object": ", ".join(f"{r['From Table']} ↔ {r['To Table']}" for _, r in bidi.head(8).iterrows()),
                "Severity": "Medium",
                "Finding": f"{len(bidi)} bi-directional relationship(s). These are permitted "
                           "but are a common source of ambiguity and slow DAX in any mode.",
                "Fix": "Replace with single-direction filtering plus CROSSFILTER() only in "
                       "the specific measures that need it.",
            })

    if not rows:
        rows.append({
            "Check": "Direct Lake blockers", "Object": "Whole model", "Severity": "Info",
            "Finding": "No calculated tables/columns, auto date tables, or M-transform "
                       "blockers found in this export.",
            "Fix": "Still verify the source Delta tables are V-Order optimised — a .vpax "
                   "carries no information about the storage layer.",
        })
    return sort_by_severity(_ensure_columns(pd.DataFrame(rows), schema))


# ==========================================================================
# Display folder / taxonomy
# ==========================================================================

def build_taxonomy(model: Dict[str, Any]) -> Tuple[Dict[str, Any], pd.DataFrame]:
    """Group measures and columns by DisplayFolder, and flag taxonomy gaps.

    Returns (tree, issues). The tree is {table: {folder_path: [items]}} where
    a folder path of "" means the object sits loose at the root of the table.
    Issues call out the two things that actually confuse report authors: a
    measure with no display folder in a table that otherwise uses them, and a
    foreign-key column left visible in the field list.
    """
    tree: Dict[str, Dict[str, List[Dict[str, str]]]] = {}
    issues: List[Dict[str, Any]] = []

    def add(table: str, folder: str, name: str, kind: str, hidden: bool) -> None:
        tree.setdefault(table, {}).setdefault(folder or "", []).append(
            {"name": name, "kind": kind, "hidden": hidden}
        )

    meas = model["measures"]
    if {"TableName", "MeasureName"}.issubset(meas.columns):
        for _, m in meas.dropna(subset=["MeasureName"]).iterrows():
            add(str(m["TableName"]), str(m.get("DisplayFolder") or ""),
                str(m["MeasureName"]), "measure", False)

    cols = _user_facing_columns(model["columns"])
    if {"TableName", "ColumnName"}.issubset(cols.columns):
        for _, c in cols.dropna(subset=["ColumnName"]).iterrows():
            add(str(c["TableName"]), str(c.get("DisplayFolder") or ""),
                str(c["ColumnName"]), "column", bool(c.get("IsHidden")))

    # Orphaned measures: only a problem where the table clearly *has* a
    # folder taxonomy, otherwise "no folders anywhere" is a valid choice.
    for table, folders in tree.items():
        measures_in_folders = sum(
            1 for f, items in folders.items() if f for i in items if i["kind"] == "measure"
        )
        loose = [i["name"] for i in folders.get("", []) if i["kind"] == "measure"]
        if measures_in_folders and loose:
            issues.append({
                "Issue": "Measure outside the folder taxonomy", "Table": table,
                "Objects": ", ".join(sorted(loose)[:15]),
                "Severity": "Low",
                "Why it matters": f"{len(loose)} measure(s) sit at the root while "
                                  f"{measures_in_folders} are filed in folders — report "
                                  "authors will scroll past them.",
                "How to Fix": "Tabular Editor: multi-select these measures ➜ set Display Folder "
                              "to match the table's existing folders. In Power BI Desktop: "
                              "Model view ➜ select each measure ➜ Properties ➜ Display Folder.",
            })

    # Visible foreign keys: the single most common cause of a report author
    # dragging a key column onto a visual and getting a meaningless number.
    rels = model["relationships"]
    if not rels.empty and {"TableName", "ColumnName"}.issubset(cols.columns) and "IsHidden" in cols.columns:
        keys: Set[Tuple[str, str]] = set()
        for _, r in rels.iterrows():
            if r.get("From Table") and r.get("From Column"):
                keys.add((str(r["From Table"]), str(r["From Column"])))
            if r.get("To Table") and r.get("To Column"):
                keys.add((str(r["To Table"]), str(r["To Column"])))
        visible_keys: Dict[str, List[str]] = {}
        for _, c in cols.iterrows():
            key = (str(c.get("TableName")), str(c.get("ColumnName")))
            if key in keys and not bool(c.get("IsHidden")):
                visible_keys.setdefault(key[0], []).append(key[1])
        for table, names in visible_keys.items():
            issues.append({
                "Issue": "Relationship key visible in the field list", "Table": table,
                "Objects": ", ".join(sorted(names)),
                "Severity": "Medium",
                "Why it matters": "Join keys aren't meaningful to report authors and "
                                  "summing or grouping by one produces nonsense. Hide them.",
                "How to Fix": "Tabular Editor: select the column(s) ➜ set IsHidden = true. In "
                              "Power BI Desktop: right-click the column in the Fields pane ➜ "
                              "Hide in report view. Govern ➜ Fix Script (C#) generates this too.",
            })

    issues_df = _ensure_columns(
        pd.DataFrame(issues),
        ["Issue", "Table", "Objects", "Severity", "Why it matters", "How to Fix"],
    )
    return tree, sort_by_severity(issues_df)


# ==========================================================================
# RLS filter-propagation simulator
# ==========================================================================

def simulate_rls(model: Dict[str, Any]) -> pd.DataFrame:
    """Trace each role's filters through the relationship graph.

    RLS filters propagate exactly the way any other filter does: down the one
    side to the many side, and back up only where cross-filtering is set to
    both. So a role that secures a dimension protects every fact hanging off
    it, but a role that secures a *fact* protects nothing else unless a
    bi-directional relationship carries the filter back up - and a filter that
    has to travel up a single-direction relationship is trapped. That is the
    silent failure this simulates: the role looks configured, the data isn't
    actually secured.
    """
    schema = ["Role", "Secured Table", "Table", "Rows Filtered?", "Severity", "Path / Reason", "How to Fix"]
    roles = model["roles"]
    if roles.empty or "Role" not in roles.columns:
        return _ensure_columns(pd.DataFrame(), schema)

    rels = model["relationships"]
    # Adjacency with the direction filtering actually flows in.
    # one -> many always; many -> one only when cross-filter is Both.
    forward: Dict[str, List[Tuple[str, str]]] = {}

    def link(a: str, b: str, why: str) -> None:
        forward.setdefault(a, []).append((b, why))

    if not rels.empty:
        for _, r in rels.iterrows():
            ft, tt = str(r.get("From Table") or ""), str(r.get("To Table") or "")
            if not ft or not tt or ft == tt:
                continue
            if not bool(r.get("Active", True)):
                continue
            both = str(r.get("Cross Filter Direction")) == "Both"
            fcard = str(r.get("From Cardinality") or "Many").lower()
            # The "one" side filters the "many" side.
            if fcard.startswith("one"):
                link(ft, tt, "1→* relationship")
                if both:
                    link(tt, ft, "bi-directional relationship")
            else:
                link(tt, ft, "1→* relationship")
                if both:
                    link(ft, tt, "bi-directional relationship")

    all_tables = set(model["all_table_names"])
    rows: List[Dict[str, Any]] = []

    for role_name, grp in roles.groupby("Role"):
        secured = sorted({str(t) for t in grp["Table"].dropna() if str(t)})
        if not secured:
            rows.append({
                "Role": str(role_name), "Secured Table": "—", "Table": "—",
                "Rows Filtered?": "No", "Severity": "Medium",
                "Path / Reason": "This role has no table permissions at all — it grants "
                                 "unrestricted read access to the whole model.",
                "How to Fix": "Power BI Desktop: Modeling ➜ Manage roles ➜ select this role ➜ add "
                              "a table and a DAX filter expression — or delete the role if it's "
                              "unused.",
            })
            continue

        for start in secured:
            reached: Dict[str, str] = {start: "filter defined directly on this table"}
            queue = deque([start])
            while queue:
                cur = queue.popleft()
                for nxt, why in forward.get(cur, []):
                    if nxt in reached:
                        continue
                    reached[nxt] = f"{reached[cur]} → {cur} → {nxt} ({why})"
                    queue.append(nxt)

            for table in sorted(all_tables):
                protected = table in reached
                if protected:
                    sev = "Info"
                    reason = reached[table]
                    fix = "No action needed — this table is already secured."
                else:
                    # Is it unreachable because of direction, or not connected at all?
                    connected = any(
                        table in (str(r.get("From Table")), str(r.get("To Table")))
                        and start in (str(r.get("From Table")), str(r.get("To Table")))
                        for _, r in rels.iterrows()
                    ) if not rels.empty else False
                    if connected:
                        sev = "High"
                        reason = (f"Related to {start}, but the filter can't travel that way — "
                                  "it would have to go many→one across a single-direction "
                                  "relationship. Rows here are NOT secured.")
                        fix = (f"Power BI Desktop: Model view ➜ select the relationship between "
                               f"{start} and {table} ➜ tick *Apply security filter in both "
                               "directions* (bi-directional only for RLS, not for report "
                               "filtering) — or move the role's filter onto the dimension side "
                               "of the relationship instead of the fact side.")
                    else:
                        sev = "Low"
                        reason = f"No filter path from {start}; this table is unaffected by the role."
                        fix = ("Expected if this table is genuinely unrelated to what the role "
                               "secures — add a relationship first if it should be in scope.")
                rows.append({
                    "Role": str(role_name), "Secured Table": start, "Table": table,
                    "Rows Filtered?": "Yes" if protected else "No",
                    "Severity": sev, "Path / Reason": reason, "How to Fix": fix,
                })

    return _ensure_columns(pd.DataFrame(rows), schema)


def rls_exposure_summary(sim_df: pd.DataFrame, model: Dict[str, Any]) -> pd.DataFrame:
    """Per-role headline: how many tables each role actually secures.

    Fact tables are what matter — a role that secures every dimension but
    leaves the fact wide open protects nothing.
    """
    schema = ["Role", "Tables Secured", "Tables Exposed", "Trapped Filters", "Severity", "Verdict"]
    if sim_df.empty:
        return _ensure_columns(pd.DataFrame(), schema)
    rows = []
    for role, grp in sim_df.groupby("Role"):
        secured = int((grp["Rows Filtered?"] == "Yes").sum())
        trapped = int((grp["Severity"] == "High").sum())
        exposed = int((grp["Rows Filtered?"] == "No").sum())
        if trapped:
            sev, verdict = "High", (f"{trapped} related table(s) are left unsecured because the "
                                    "filter can't cross a single-direction relationship.")
        elif secured <= 1:
            sev, verdict = "Medium", ("Only the secured table itself is filtered — nothing "
                                      "propagates. Check this is intentional.")
        else:
            sev, verdict = "Info", f"Filters reach {secured} table(s) as expected."
        rows.append({"Role": role, "Tables Secured": secured, "Tables Exposed": exposed,
                     "Trapped Filters": trapped, "Severity": sev, "Verdict": verdict})
    return sort_by_severity(_ensure_columns(pd.DataFrame(rows), schema))


# ==========================================================================
# Model compare / metric drift
# ==========================================================================

def _normalise_dax(expr: Any) -> str:
    """Collapse formatting so only real logic changes register as drift.

    Whitespace, line breaks and comments are how one developer's copy of a
    measure differs from another's without the maths differing at all. Case is
    preserved deliberately - DAX is case-insensitive, but a renamed reference
    is worth seeing, and the diff view shows the original text anyway.
    """
    s = str(expr or "")

    # String literals are lifted out first: a space inside "North West" is
    # data, not formatting, and collapsing it would hide a real change.
    literals: List[str] = []

    def _stash(match: "re.Match[str]") -> str:
        literals.append(match.group(0))
        return f"\x00{len(literals) - 1}\x00"

    s = re.sub(r'"(?:[^"]|"")*"', _stash, s)
    s = re.sub(r"/\*.*?\*/", " ", s, flags=re.S)       # block comments
    s = re.sub(r"(--|//)[^\n]*", " ", s)                # line comments
    s = re.sub(r"\s+", " ", s)
    # Whitespace around operators and punctuation is pure formatting - one
    # developer's SUMX( Sales, ... ) is another's SUMX(Sales, ...).
    s = re.sub(r"\s*([(),\[\]{}=<>+\-*/&^:;])\s*", r"\1", s)
    s = s.strip()
    for i, lit in enumerate(literals):
        s = s.replace(f"\x00{i}\x00", lit)
    return s


def compare_models(base: Dict[str, Any], other: Dict[str, Any]) -> Dict[str, Any]:
    """Diff a working model against a certified baseline.

    The headline case is metric drift: a measure that exists in both models
    under the same name but whose DAX has been changed locally. That is how a
    "single source of truth" quietly stops being one — the name still matches
    the certified metric, so nobody notices the number no longer does.
    """
    def measure_map(m: Dict[str, Any]) -> Dict[str, Tuple[str, str]]:
        df = m["measures"]
        out: Dict[str, Tuple[str, str]] = {}
        if {"MeasureName", "MeasureExpression"}.issubset(df.columns):
            for _, r in df.dropna(subset=["MeasureName"]).iterrows():
                out[str(r["MeasureName"])] = (str(r.get("TableName") or ""),
                                              str(r.get("MeasureExpression") or ""))
        return out

    bm, om = measure_map(base), measure_map(other)
    drift_rows, added_rows, removed_rows = [], [], []

    for name in sorted(set(bm) & set(om)):
        b_table, b_expr = bm[name]
        o_table, o_expr = om[name]
        moved = b_table != o_table
        changed = _normalise_dax(b_expr) != _normalise_dax(o_expr)
        if not changed and not moved:
            continue
        if changed:
            sev, what = "High", "DAX changed"
            why = ("The same measure name now computes something different than the certified "
                   "baseline — anything trusting the name (reports, other measures, documentation) "
                   "is silently reading the wrong logic.")
            fix = ("Compare the DAX side by side below. If the change is intentional, update the "
                   "baseline .vpax to match; if it's not, revert this measure's expression to the "
                   "baseline's, or rename the local version so it stops shadowing the certified one.")
        else:
            sev, what = "Medium", "Moved to a different home table"
            why = "Same name and logic, but a different home table can break fully-qualified references."
            fix = ("Confirm the move was deliberate. If not, move the measure back to its "
                   "original table (Tabular Editor: drag the measure between tables in the "
                   "Explorer pane).")
        drift_rows.append({
            "Measure": name, "Change": what, "Severity": sev,
            "Baseline Table": b_table, "Compared Table": o_table,
            "Why It Matters": why, "How to Fix": fix,
            "Baseline DAX": b_expr, "Compared DAX": o_expr,
        })
    for name in sorted(set(om) - set(bm)):
        added_rows.append({
            "Measure": name, "Table": om[name][0], "Severity": "Medium", "DAX": om[name][1],
            "Note": "Exists only in the compared model — a local metric that "
                    "isn't part of the certified set.",
            "How to Fix": "If this is a genuinely new, approved metric, add it to the certified "
                          "baseline so it's tracked going forward. If it's a one-off, keep it — "
                          "just be aware it won't appear when comparing other models against the "
                          "same baseline.",
        })
    for name in sorted(set(bm) - set(om)):
        removed_rows.append({
            "Measure": name, "Table": bm[name][0], "Severity": "High", "DAX": bm[name][1],
            "Note": "Present in the baseline but missing here — anything "
                    "referencing it will break.",
            "How to Fix": "Check Investigate ➜ Impact Analysis on this measure name against the "
                          "baseline model to see what depended on it, then restore the measure "
                          "(or repoint those dependents at its replacement) before this model "
                          "replaces the baseline.",
        })

    def name_set(m: Dict[str, Any], kind: str) -> Set[str]:
        if kind == "tables":
            return set(m["all_table_names"])
        cols = _user_facing_columns(m["columns"])
        if not {"TableName", "ColumnName"}.issubset(cols.columns):
            return set()
        return {f"{r['TableName']}[{r['ColumnName']}]" for _, r in cols.dropna(
            subset=["TableName", "ColumnName"]).iterrows()}

    def rel_set(m: Dict[str, Any]) -> Set[str]:
        df = m["relationships"]
        if df.empty:
            return set()
        return {
            f"{r['From Table']}[{r['From Column']}] → {r['To Table']}[{r['To Column']}] "
            f"({r['Cross Filter Direction']})"
            for _, r in df.iterrows()
        }

    struct_rows = []
    for label, b_set, o_set in (
        ("Table", name_set(base, "tables"), name_set(other, "tables")),
        ("Column", name_set(base, "columns"), name_set(other, "columns")),
        ("Relationship", rel_set(base), rel_set(other)),
    ):
        for item in sorted(o_set - b_set):
            struct_rows.append({"Object Type": label, "Object": item, "Change": "Added",
                                "Severity": "Low"})
        for item in sorted(b_set - o_set):
            struct_rows.append({"Object Type": label, "Object": item, "Change": "Removed",
                                "Severity": "Medium"})

    return {
        "drift": _ensure_columns(pd.DataFrame(drift_rows), [
            "Measure", "Change", "Severity", "Baseline Table", "Compared Table",
            "Why It Matters", "How to Fix", "Baseline DAX", "Compared DAX"]),
        "added": _ensure_columns(pd.DataFrame(added_rows),
                                 ["Measure", "Table", "Severity", "DAX", "Note", "How to Fix"]),
        "removed": _ensure_columns(pd.DataFrame(removed_rows),
                                   ["Measure", "Table", "Severity", "DAX", "Note", "How to Fix"]),
        "structure": _ensure_columns(pd.DataFrame(struct_rows), ["Object Type", "Object", "Change", "Severity"]),
    }


# ==========================================================================
# Tabular Editor C# fix-script generation
# ==========================================================================

def _cs_string(value: Any) -> str:
    """C# verbatim-safe string literal."""
    return '"' + str(value).replace("\\", "\\\\").replace('"', '\\"') + '"'


def build_te_script(
    model: Dict[str, Any],
    naming_df: pd.DataFrame,
    include_renames: bool = True,
    include_mdx: bool = True,
    include_hide_keys: bool = True,
    include_formats: bool = True,
    include_descriptions: bool = True,
) -> str:
    """Emit a C# script for Tabular Editor's Advanced Scripting window.

    Everything the audit tabs flag as a mechanical fix — renames to match the
    dominant convention, IsAvailableInMdx on hidden columns, hiding join keys,
    a default format string on unformatted measures — is expressed here as TOM
    calls. Renames are commented out by default because a rename breaks every
    report visual that references the old name; the reviewer uncomments the
    ones they've checked.
    """
    lines: List[str] = [
        "// =====================================================================",
        "// Generated by VPAX Semantic Model Explorer",
        f"// Model: {model.get('model_name') or '(unnamed)'}",
        "//",
        "// HOW TO RUN",
        "//   1. Open the model in Tabular Editor (2 or 3).",
        "//   2. Advanced Scripting tab ➜ paste this script ➜ run (F5).",
        "//   3. Review the changes, then Save to the model.",
        "//",
        "// READ THIS FIRST",
        "//   Renames are commented out on purpose. Renaming a column or measure",
        "//   breaks every report visual, bookmark and RLS expression that refers",
        "//   to the old name. Uncomment only the ones you have checked in the",
        "//   Impact Analysis tab first.",
        "// =====================================================================",
        "",
    ]
    emitted = 0

    if include_renames and not naming_df.empty and {"Object Type", "Table", "Name"}.issubset(naming_df.columns):
        lines += ["// --- 1. Naming convention renames (COMMENTED OUT — review first) ---"]
        for _, r in naming_df.iterrows():
            suggestion = str(r.get("Suggestion") or "")
            m = re.search(r"renaming to `([^`]+)`", suggestion)
            if not m:
                continue
            new_name, obj_type = m.group(1), str(r["Object Type"])
            table, old = str(r["Table"]), str(r["Name"])
            if obj_type == "Measure":
                lines.append(f'// Model.Tables[{_cs_string(table)}].Measures[{_cs_string(old)}]'
                             f'.Name = {_cs_string(new_name)};')
            elif obj_type == "Column":
                lines.append(f'// Model.Tables[{_cs_string(table)}].Columns[{_cs_string(old)}]'
                             f'.Name = {_cs_string(new_name)};')
            elif obj_type == "Table":
                lines.append(f'// Model.Tables[{_cs_string(old)}].Name = {_cs_string(new_name)};')
            else:
                continue
            emitted += 1
        lines.append("")

    cols = _user_facing_columns(model["columns"])

    if include_mdx and not cols.empty and {"IsHidden", "IsAvailableInMDX"}.issubset(cols.columns):
        targets = cols[(cols["IsHidden"] == True) & (cols["IsAvailableInMDX"] == True)]  # noqa: E712
        if not targets.empty:
            lines += [
                "// --- 2. Drop attribute hierarchies on hidden columns ---",
                "// Hidden columns can't be browsed, so the attribute hierarchy VertiPaq",
                "// builds for them is pure memory overhead. Safe and reversible.",
            ]
            for _, c in targets.iterrows():
                lines.append(
                    f'Model.Tables[{_cs_string(c["TableName"])}]'
                    f'.Columns[{_cs_string(c["ColumnName"])}].IsAvailableInMdx = false;'
                )
                emitted += 1
            lines.append("")

    if include_hide_keys:
        rels = model["relationships"]
        key_targets: List[Tuple[str, str]] = []
        if not rels.empty and not cols.empty and "IsHidden" in cols.columns:
            keys: Set[Tuple[str, str]] = set()
            for _, r in rels.iterrows():
                if r.get("From Table") and r.get("From Column"):
                    keys.add((str(r["From Table"]), str(r["From Column"])))
                if r.get("To Table") and r.get("To Column"):
                    keys.add((str(r["To Table"]), str(r["To Column"])))
            for _, c in cols.iterrows():
                pair = (str(c.get("TableName")), str(c.get("ColumnName")))
                if pair in keys and not bool(c.get("IsHidden")):
                    key_targets.append(pair)
        if key_targets:
            lines += [
                "// --- 3. Hide relationship key columns ---",
                "// Join keys aren't meaningful to report authors; grouping by one",
                "// produces a misleading number. Hiding is non-breaking.",
            ]
            for table, col in key_targets:
                lines.append(
                    f'Model.Tables[{_cs_string(table)}].Columns[{_cs_string(col)}].IsHidden = true;'
                )
                emitted += 1
            lines.append("")

    if include_formats:
        meas = model["measures"]
        if not meas.empty and "MeasureName" in meas.columns:
            unformatted = meas[
                meas.get("FormatString", pd.Series([""] * len(meas), index=meas.index))
                .fillna("").astype(str).str.strip() == ""
            ] if "FormatString" in meas.columns else meas.iloc[0:0]
            if not unformatted.empty:
                lines += [
                    "// --- 4. Default format string on unformatted measures ---",
                    '// Adjust "#,0.00" to your standard before running.',
                ]
                for _, m in unformatted.iterrows():
                    lines.append(
                        f'Model.Tables[{_cs_string(m["TableName"])}]'
                        f'.Measures[{_cs_string(m["MeasureName"])}].FormatString = "#,0.00";'
                    )
                    emitted += 1
                lines.append("")

    if include_descriptions:
        meas = model["measures"]
        if not meas.empty and "MeasureName" in meas.columns and "Description" in meas.columns:
            undocumented = meas[meas["Description"].fillna("").astype(str).str.strip() == ""]
            if not undocumented.empty:
                lines += [
                    "// --- 5. Description placeholders ---",
                    "// Stubs only — replace the text with the real business definition.",
                ]
                for _, m in undocumented.head(200).iterrows():
                    lines.append(
                        f'// Model.Tables[{_cs_string(m["TableName"])}]'
                        f'.Measures[{_cs_string(m["MeasureName"])}].Description = '
                        f'"TODO: business definition";'
                    )
                    emitted += 1
                lines.append("")

    if not emitted:
        lines.append("// Nothing to fix — none of the selected checks found actionable items.")
    else:
        lines += [
            'Info("VPAX Explorer fix script finished. Review the changes, then Save.");',
        ]
    return "\n".join(lines)


# ==========================================================================
# Report-level usage: broken visuals and true "safe to delete"
# ==========================================================================

def _report_bindings(report: Optional[Dict[str, Any]]) -> pd.DataFrame:
    if not report:
        return _ensure_columns(pd.DataFrame(), BINDING_COLUMNS)
    df = report.get("bindings")
    if not isinstance(df, pd.DataFrame):
        return _ensure_columns(pd.DataFrame(), BINDING_COLUMNS)
    return df


def validate_report_bindings(bindings: pd.DataFrame, model: Dict[str, Any]) -> pd.DataFrame:
    """Find visuals bound to model objects that no longer exist.

    This is the other half of impact analysis. A .vpax alone tells you what
    the model contains; the .pbix tells you what the report *asks for*. When
    somebody renames or deletes a column, the model stays perfectly valid and
    the report breaks - Power BI only surfaces that when a user opens the
    page. Comparing the two catches it before they do.
    """
    schema = ["Page", "Visual", "Visual Type", "Kind", "Table", "Field", "Severity",
              "Problem", "How to Fix"]
    if bindings.empty:
        return _ensure_columns(pd.DataFrame(), schema)

    tables = set(model["all_table_names"])
    cols = _user_facing_columns(model["columns"])
    col_pairs: Set[Tuple[str, str]] = set()
    if {"TableName", "ColumnName"}.issubset(cols.columns):
        col_pairs = {(str(r["TableName"]), str(r["ColumnName"]))
                     for _, r in cols.dropna(subset=["TableName", "ColumnName"]).iterrows()}
    meas = model["measures"]
    measure_pairs: Set[Tuple[str, str]] = set()
    measure_any: Set[str] = set()
    if {"TableName", "MeasureName"}.issubset(meas.columns):
        for _, r in meas.dropna(subset=["MeasureName"]).iterrows():
            measure_pairs.add((str(r["TableName"]), str(r["MeasureName"])))
            measure_any.add(str(r["MeasureName"]))

    # Hierarchies are declared in the TOM, not in the VPA column list.
    hierarchy_levels: Set[Tuple[str, str]] = set()
    for t in model.get("bim_tables") or []:
        tname = str(t.get("name") or "")
        for h in t.get("hierarchies") or []:
            if not isinstance(h, dict):
                continue
            for lvl in h.get("levels") or []:
                if isinstance(lvl, dict) and lvl.get("name"):
                    hierarchy_levels.add((tname, str(lvl["name"])))

    rows: List[Dict[str, Any]] = []
    for _, b in bindings.iterrows():
        table, field, kind = str(b["Table"]), str(b["Field"]), str(b["Kind"])
        if not table:
            # An unresolvable alias usually means a visual-level calculation or
            # a literal, not a broken reference - reporting it would be noise.
            continue
        if table not in tables:
            rows.append({
                **b.to_dict(), "Severity": "High",
                "Problem": f"Table `{table}` doesn't exist in this model. The visual "
                           "will error or render blank.",
                "How to Fix": f"Open the report, go to page **{b['Page']}**, select the visual "
                              f"and rebind its fields to a table that still exists — or delete "
                              "the visual if `{table}` was removed on purpose.",
            })
            continue
        if kind == "Measure":
            if (table, field) in measure_pairs:
                continue
            if field in measure_any:
                # Measures can be moved between home tables without breaking
                # the report, so this is informational, not a break.
                continue
            if (table, field) in col_pairs:
                continue
            problem = f"No measure named `{field}` on `{table}`."
            fix = (f"The measure was likely renamed or deleted. Open page **{b['Page']}**, "
                   f"select the visual, and rebind it to the measure's new name — check "
                   "Investigate ➜ Measure Dependencies if you're not sure what replaced it.")
        elif kind == "Hierarchy level":
            if (table, field) in hierarchy_levels or (table, field) in col_pairs:
                continue
            problem = f"No hierarchy level `{field}` on `{table}`."
            fix = (f"Open page **{b['Page']}**, select the visual, and rebind it to a level that "
                   f"still exists on `{table}`'s hierarchy (or to the plain column).")
        else:
            if (table, field) in col_pairs or (table, field) in measure_pairs:
                continue
            problem = f"No column named `{field}` on `{table}`."
            fix = (f"The column was likely renamed or deleted. Open page **{b['Page']}**, select "
                   f"the visual, and rebind it to the correct field on `{table}`.")
        rows.append({**b.to_dict(), "Severity": "High", "Problem": problem, "How to Fix": fix})

    return sort_by_severity(_ensure_columns(pd.DataFrame(rows), schema))


def classify_column_disposition(
    model: Dict[str, Any], unused_df: pd.DataFrame, bindings: pd.DataFrame
) -> pd.DataFrame:
    """Combine the DAX scan with report usage into a real delete verdict.

    On its own the DAX scan can only say "nothing in the model references
    this", which is why the Unused Objects view is careful to call that a
    lead rather than a verdict. A column dropped straight onto a bar chart
    axis is invisible to it. With the .pbix in hand the two blind spots cancel
    out, and a column that neither DAX nor any visual touches can honestly be
    called safe to delete.
    """
    schema = ["Table", "Column", "In DAX", "In Report", "Verdict", "Severity", "Used On"]
    if unused_df.empty:
        return _ensure_columns(pd.DataFrame(), schema)

    have_report = not bindings.empty
    used_pages: Dict[Tuple[str, str], Set[str]] = {}
    if have_report:
        for _, b in bindings.iterrows():
            key = (str(b["Table"]), str(b["Field"]))
            used_pages.setdefault(key, set()).add(str(b["Page"]))
    # Field names are also matched table-blind, because a visual binding whose
    # alias couldn't be resolved still proves the *name* is in use somewhere.
    used_names = {f for (_, f) in used_pages}

    rows = []
    for _, r in unused_df.iterrows():
        table, col = str(r["Table"]), str(r["Column"])
        in_dax = str(r["Status"]) != "Likely unused"
        pages = used_pages.get((table, col), set())
        in_report = bool(pages) or (not pages and col in used_names)
        where = ", ".join(sorted(pages)[:6]) if pages else ("(name used, table unresolved)" if in_report else "")

        if not have_report:
            verdict, sev = ("Referenced in DAX" if in_dax else "Unknown — upload the .pbix"), "Info" if in_dax else "Low"
        elif in_dax and in_report:
            verdict, sev = "Keep — used in DAX and on report visuals", "Info"
        elif in_dax:
            verdict, sev = "Keep — used in DAX", "Info"
        elif in_report:
            verdict, sev = "Keep — used on report visuals only", "Info"
        else:
            verdict, sev = "Safe to delete — no DAX and no visual references it", "Medium"

        rows.append({
            "Table": table, "Column": col,
            "In DAX": "Yes" if in_dax else "No",
            "In Report": ("Yes" if in_report else "No") if have_report else "?",
            "Verdict": verdict, "Severity": sev, "Used On": where,
        })
    return _ensure_columns(pd.DataFrame(rows), schema)


# ==========================================================================
# Field parameters
# ==========================================================================

def find_field_parameters(
    model: Dict[str, Any], bindings: pd.DataFrame
) -> pd.DataFrame:
    """Detect field-parameter tables and whether the report actually uses them.

    A field parameter is a disconnected calculated table whose rows are
    NAMEOF() references to other fields, marked up with a `ParameterMetadata`
    extended property. They're easy to create, easy to forget, and because
    they're disconnected nothing else in the model ever points at them - so an
    abandoned one is invisible to every other check in this app. If no visual
    or slicer in the report binds to it, the whole table is dead weight.
    """
    schema = ["Table", "Columns", "Referenced Fields", "Used In Report",
              "Bound On", "Severity", "Verdict"]
    rows: List[Dict[str, Any]] = []
    have_report = not bindings.empty
    bound_tables = set(bindings["Table"].astype(str)) if have_report else set()
    pages_by_table: Dict[str, Set[str]] = {}
    if have_report:
        for _, b in bindings.iterrows():
            pages_by_table.setdefault(str(b["Table"]), set()).add(str(b["Page"]))

    # VPA carries the calculated-table expression; TOM carries the marker.
    expr_by_table: Dict[str, str] = {}
    tdf = model["tables"]
    if not tdf.empty and {"TableName", "TableExpression"}.issubset(tdf.columns):
        for _, t in tdf.iterrows():
            expr_by_table[str(t["TableName"])] = str(t.get("TableExpression") or "")

    for t in model.get("bim_tables") or []:
        name = str(t.get("name") or "")
        if not name:
            continue
        columns = [c for c in (t.get("columns") or []) if isinstance(c, dict)]
        has_marker = any(
            isinstance(ep, dict) and str(ep.get("name")) == "ParameterMetadata"
            for c in columns for ep in (c.get("extendedProperties") or [])
        )
        expr = expr_by_table.get(name, "")
        if not expr:
            for p in t.get("partitions") or []:
                src = (p or {}).get("source") if isinstance(p, dict) else None
                if isinstance(src, dict) and str(src.get("type") or "").lower() == "calculated":
                    expr = str(src.get("expression") or "")
        looks_like = "NAMEOF(" in expr.upper().replace(" ", "")
        if not has_marker and not looks_like:
            continue

        # A NAMEOF argument is always a field reference, and measure names
        # regularly contain brackets and parentheses ("% of Total (Selected)"),
        # so match the reference shape rather than "everything up to the next )".
        referenced = sorted(set(re.findall(
            r"NAMEOF\s*\(\s*('[^']+'\[[^\]]+\]|\"[^\"]+\"\[[^\]]+\]|[A-Za-z_][\w ]*\[[^\]]+\]|\[[^\]]+\])",
            expr, flags=re.I,
        )))
        col_names = [str(c.get("name")) for c in columns
                     if c.get("name") and str(c.get("type") or "").lower() != "rownumber"]
        pages = sorted(pages_by_table.get(name, set()))

        if not have_report:
            sev, verdict = "Info", ("Field parameter table. Upload the matching .pbix to check "
                                    "whether any slicer or visual actually uses it.")
        elif name in bound_tables:
            sev, verdict = "Info", f"In use on {len(pages)} page(s)."
        else:
            sev, verdict = "Medium", ("No visual or slicer in this report binds to it. The table, "
                                      "its columns and its hidden sort column can all be removed.")
        rows.append({
            "Table": name, "Columns": ", ".join(col_names),
            "Referenced Fields": ", ".join(referenced[:12]),
            "Used In Report": ("Yes" if name in bound_tables else "No") if have_report else "?",
            "Bound On": ", ".join(pages[:6]), "Severity": sev, "Verdict": verdict,
        })
    return sort_by_severity(_ensure_columns(pd.DataFrame(rows), schema))


# ==========================================================================
# Near-duplicate DAX
# ==========================================================================

_DAX_VAR_RE = re.compile(r"\bVAR\s+([A-Za-z_][A-Za-z0-9_]*)", re.I)


def _dax_fingerprint(expr: Any) -> str:
    """Normalise DAX down to its logic, ignoring cosmetic differences.

    Beyond the whitespace and comment stripping the drift comparison does,
    this also renames every VAR to a positional placeholder. Two developers
    solving the same problem will pick different variable names for the same
    intermediate value, and that difference says nothing about whether the
    measures compute the same number.
    """
    s = _normalise_dax(expr)
    names = []
    for m in _DAX_VAR_RE.finditer(s):
        if m.group(1) not in names:
            names.append(m.group(1))
    for i, name in enumerate(names):
        s = re.sub(rf"\b{re.escape(name)}\b", f"__V{i}__", s)
    return s.upper()   # DAX is case-insensitive


def _tokenise_dax(fingerprint: str) -> Set[str]:
    return set(re.findall(r"[A-Z_][A-Z0-9_]*|\[[^\]]+\]|\d+", fingerprint))


MAX_DUPLICATE_SCAN = 1200


def find_near_duplicate_measures(model: Dict[str, Any], threshold: int = 85) -> pd.DataFrame:
    """Measures that compute nearly the same thing under different names.

    Exact-duplicate detection misses the real-world case: someone copies a
    certified measure, adds `+ 0` to force a zero instead of a blank, renames
    a variable, and now the model carries two metrics that are the same number
    99% of the time and disagree the rest. Comparison is two-stage - a cheap
    token-overlap filter, then a character-level similarity ratio on the pairs
    that survive - so an O(n^2) comparison stays tractable on a big model.
    """
    schema = ["Measure A", "Table A", "Measure B", "Table B", "Similarity %",
              "Severity", "Verdict", "DAX A", "DAX B"]
    meas = model["measures"]
    if meas.empty or not {"MeasureName", "MeasureExpression"}.issubset(meas.columns):
        return _ensure_columns(pd.DataFrame(), schema)

    items: List[Tuple[str, str, str, str, Set[str]]] = []
    for _, r in meas.dropna(subset=["MeasureName"]).iterrows():
        expr = str(r.get("MeasureExpression") or "")
        fp = _dax_fingerprint(expr)
        if len(fp) < 8:      # trivial constants aren't interesting duplicates
            continue
        items.append((str(r["MeasureName"]), str(r.get("TableName") or ""), expr, fp, _tokenise_dax(fp)))
        if len(items) >= MAX_DUPLICATE_SCAN:
            break

    rows: List[Dict[str, Any]] = []
    for i in range(len(items)):
        name_a, table_a, expr_a, fp_a, tok_a = items[i]
        for j in range(i + 1, len(items)):
            name_b, table_b, expr_b, fp_b, tok_b = items[j]
            if not tok_a or not tok_b:
                continue
            # Cheap gate first: length ratio, then Jaccard on tokens.
            shorter, longer = sorted((len(fp_a), len(fp_b)))
            if longer and shorter / longer < threshold / 100 * 0.8:
                continue
            jaccard = len(tok_a & tok_b) / len(tok_a | tok_b)
            if jaccard * 100 < threshold * 0.75:
                continue
            ratio = difflib.SequenceMatcher(None, fp_a, fp_b).ratio() * 100
            if ratio < threshold:
                continue
            if ratio >= 99.9:
                sev, verdict = "High", ("Identical logic under two names. One of these should be "
                                        "deleted and its references repointed.")
            elif ratio >= 95:
                sev, verdict = "High", ("Near-identical — the difference is likely a blank-handling "
                                        "or formatting tweak. Confirm they're meant to differ.")
            else:
                sev, verdict = "Medium", ("Substantially similar. Check whether one can be expressed "
                                          "in terms of the other rather than duplicated.")
            rows.append({
                "Measure A": name_a, "Table A": table_a,
                "Measure B": name_b, "Table B": table_b,
                "Similarity %": round(ratio, 1), "Severity": sev, "Verdict": verdict,
                "DAX A": expr_a, "DAX B": expr_b,
            })

    df = _ensure_columns(pd.DataFrame(rows), schema)
    if df.empty:
        return df
    return sort_by_severity(df.sort_values("Similarity %", ascending=False)).reset_index(drop=True)


# ==========================================================================
# Source-system lineage from Power Query SQL
# ==========================================================================

try:                                    # optional - improves accuracy a lot
    import sqlglot as _sqlglot
    from sqlglot import exp as _sqlglot_exp
except Exception:                       # noqa: BLE001 - any import problem degrades gracefully
    _sqlglot = None
    _sqlglot_exp = None

SQL_DIALECTS = ["(auto-detect)", "tsql", "snowflake", "databricks", "bigquery",
                "postgres", "oracle", "mysql", "redshift", "hive", "spark"]

_SQL_TABLE_RE = re.compile(
    r"\b(?:FROM|JOIN)\s+([\[\]\"`\w.]+(?:\s*\.\s*[\[\]\"`\w]+){0,3})", re.I
)


def _clean_sql_identifier(raw: str) -> str:
    return raw.strip().strip("[]\"`").strip()


def _sql_sources(sql: str, dialect: Optional[str] = None) -> List[Tuple[str, str, str]]:
    """(database, schema, table) for every table a SELECT reads.

    Uses sqlglot when it's installed, which correctly ignores CTE names,
    subquery aliases and table-valued functions. Without it, falls back to a
    FROM/JOIN regex - less precise, but it keeps this feature working rather
    than disappearing when the optional dependency is missing.
    """
    sql = (sql or "").strip()
    if not sql:
        return []

    if _sqlglot is not None and _sqlglot_exp is not None:
        try:
            parsed = _sqlglot.parse(sql, read=dialect or None)
        except Exception:                # noqa: BLE001 - fall through to regex
            parsed = None
        if parsed:
            found: List[Tuple[str, str, str]] = []
            for statement in parsed:
                if statement is None:
                    continue
                # CTE names are not real source tables.
                cte_names = {
                    str(c.alias_or_name).lower()
                    for c in statement.find_all(_sqlglot_exp.CTE)
                }
                for tbl in statement.find_all(_sqlglot_exp.Table):
                    name = str(tbl.name or "")
                    if not name or name.lower() in cte_names:
                        continue
                    found.append((str(tbl.catalog or ""), str(tbl.db or ""), name))
            if found:
                return found

    out: List[Tuple[str, str, str]] = []
    for match in _SQL_TABLE_RE.finditer(sql):
        parts = [_clean_sql_identifier(p) for p in re.split(r"\s*\.\s*", match.group(1))]
        parts = [p for p in parts if p]
        if not parts or parts[-1].lower() in ("select", "("):
            continue
        db = parts[-3] if len(parts) >= 3 else ""
        schema = parts[-2] if len(parts) >= 2 else ""
        out.append((db, schema, parts[-1]))
    return out


def build_sql_lineage(model: Dict[str, Any], dialect: Optional[str] = None) -> pd.DataFrame:
    """Map every model table back to the warehouse objects it reads.

    This is the bridge between the BI side and the data-engineering side: when
    someone proposes dropping a warehouse table, this answers "what breaks?"
    without anyone opening Power BI Desktop.
    """
    schema = ["Model Table", "Partition", "Database", "Schema", "Source Table", "Full Source Name"]
    pq = model["power_query"]
    if pq.empty or "SQL" not in pq.columns:
        return _ensure_columns(pd.DataFrame(), schema)

    rows: List[Dict[str, str]] = []
    for _, p in pq.iterrows():
        sql = str(p.get("SQL") or "")
        for db, sch, tbl in _sql_sources(sql, dialect):
            full = ".".join([x for x in (db, sch, tbl) if x])
            rows.append({
                "Model Table": str(p.get("TableName") or ""),
                "Partition": str(p.get("PartitionName") or ""),
                "Database": db, "Schema": sch, "Source Table": tbl,
                "Full Source Name": full,
            })
    df = _ensure_columns(pd.DataFrame(rows), schema)
    return df.drop_duplicates().reset_index(drop=True)


def build_source_impact(lineage: pd.DataFrame) -> pd.DataFrame:
    """Reverse the lineage: which model tables break if a source is dropped."""
    schema = ["Full Source Name", "Model Tables Affected", "Count"]
    if lineage.empty:
        return _ensure_columns(pd.DataFrame(), schema)
    rows = []
    for source, grp in lineage.groupby("Full Source Name"):
        tables = sorted(set(grp["Model Table"].astype(str)))
        rows.append({"Full Source Name": source,
                     "Model Tables Affected": ", ".join(tables), "Count": len(tables)})
    return (_ensure_columns(pd.DataFrame(rows), schema)
            .sort_values("Count", ascending=False).reset_index(drop=True))


def build_lineage_dot(lineage: pd.DataFrame, max_nodes: int = 60) -> str:
    """Source-to-target diagram: warehouse objects on the left, model tables right."""
    sources = list(dict.fromkeys(lineage["Full Source Name"].astype(str)))[:max_nodes]
    targets = list(dict.fromkeys(lineage["Model Table"].astype(str)))[:max_nodes]
    lines = [
        "digraph lineage {",
        '  graph [rankdir=LR, splines=spline, bgcolor="transparent", nodesep=.28, ranksep=1.3];',
        '  node [shape=box, style="filled,rounded", fontname="Helvetica", fontsize=10];',
        '  edge [color="#94a3b8", arrowsize=.7];',
        "  { rank=same;",
    ]
    for s in sources:
        lines.append(f'    {_dot_id("src::" + s)} [label={_dot_id(s)}, fillcolor="#e8f1fb", '
                     f'color="#2d6ca8", fontcolor="#1f3a5f"];')
    lines.append("  }")
    lines.append("  { rank=same;")
    for t in targets:
        lines.append(f'    {_dot_id("tgt::" + t)} [label={_dot_id(t)}, fillcolor="#ffffff", '
                     f'color="#94a3b8", fontcolor="#1e293b"];')
    lines.append("  }")
    seen: Set[Tuple[str, str]] = set()
    for _, r in lineage.iterrows():
        s, t = str(r["Full Source Name"]), str(r["Model Table"])
        if s not in sources or t not in targets or (s, t) in seen:
            continue
        seen.add((s, t))
        lines.append(f'  {_dot_id("src::" + s)} -> {_dot_id("tgt::" + t)};')
    lines.append("}")
    return "\n".join(lines)


# ==========================================================================
# Model cleanup: delete script + cleaned Model.bim
# ==========================================================================

def build_cleanup_plan(
    disposition: pd.DataFrame,
    field_params: pd.DataFrame,
    include_columns: bool = True,
    include_field_params: bool = True,
) -> Dict[str, Any]:
    """What a cleanup would actually remove, as concrete object lists."""
    columns: List[Tuple[str, str]] = []
    if include_columns and not disposition.empty and "Verdict" in disposition.columns:
        safe = disposition[disposition["Verdict"].astype(str).str.startswith("Safe to delete")]
        columns = [(str(r["Table"]), str(r["Column"])) for _, r in safe.iterrows()]
    tables: List[str] = []
    if include_field_params and not field_params.empty and "Used In Report" in field_params.columns:
        dead = field_params[field_params["Used In Report"] == "No"]
        tables = [str(r["Table"]) for _, r in dead.iterrows()]
    # A column on a table that's being dropped entirely is redundant noise.
    columns = [(t, c) for t, c in columns if t not in set(tables)]
    return {"columns": columns, "tables": tables}


def build_cleanup_csharp(model: Dict[str, Any], plan: Dict[str, Any]) -> str:
    """Tabular Editor script that performs the deletions in the plan."""
    lines = [
        "// =====================================================================",
        "// Model cleanup — generated by VPAX Semantic Model Explorer",
        f"// Model: {model.get('model_name') or '(unnamed)'}",
        "//",
        "// These objects are referenced by no DAX in the model AND no visual in",
        "// the .pbix that was analysed alongside it.",
        "//",
        "// STILL CHECK, BEFORE YOU RUN THIS:",
        "//   * other reports built on this same semantic model — one .pbix",
        "//     cannot speak for all of them;",
        "//   * paginated reports, Excel Analyze-in-Excel and XMLA clients;",
        "//   * anything assembled dynamically at query time.",
        "// Run against a copy, and keep the original .pbix until you've verified.",
        "// =====================================================================",
        "",
    ]
    if plan["tables"]:
        lines.append("// --- Field parameter tables with no report binding ---")
        for t in plan["tables"]:
            lines.append(f"Model.Tables[{_cs_string(t)}].Delete();")
        lines.append("")
    if plan["columns"]:
        lines.append("// --- Columns unused in both DAX and visuals ---")
        for table, col in plan["columns"]:
            lines.append(f"Model.Tables[{_cs_string(table)}].Columns[{_cs_string(col)}].Delete();")
        lines.append("")
    if not plan["tables"] and not plan["columns"]:
        lines.append("// Nothing selected for removal.")
    else:
        total = len(plan["tables"]) + len(plan["columns"])
        lines.append(f'Info("Removed {total} object(s). Review, then Save to the model.");')
    return "\n".join(lines)


def build_cleanup_tmdl(plan: Dict[str, Any]) -> str:
    """The same plan expressed as TMDL, for a definition-file workflow."""
    lines = [
        "/// Model cleanup — generated by VPAX Semantic Model Explorer.",
        "/// TMDL describes what a model *should* contain, so a deletion is",
        "/// applied by removing these blocks from the .tmdl files under",
        "/// definition/tables/ and re-deploying — there is no delete verb.",
        "/// Each entry below gives the file and the block to remove.",
        "",
    ]
    if plan["tables"]:
        lines.append("// Whole tables — delete the file outright:")
        for t in plan["tables"]:
            lines.append(f"//   definition/tables/{t}.tmdl")
        lines.append("")
    if plan["columns"]:
        lines.append("// Columns — remove these blocks from their table files:")
        by_table: Dict[str, List[str]] = {}
        for table, col in plan["columns"]:
            by_table.setdefault(table, []).append(col)
        for table in sorted(by_table):
            lines.append(f"\n// --- definition/tables/{table}.tmdl ---")
            for col in sorted(by_table[table]):
                lines.append(f"\tcolumn '{col}'")
    if not plan["tables"] and not plan["columns"]:
        lines.append("// Nothing selected for removal.")
    return "\n".join(lines)


def build_cleaned_bim(model: Dict[str, Any], plan: Dict[str, Any]) -> bytes:
    """A Model.bim with the planned objects physically stripped out.

    Relationships that pointed at a removed column go too — leaving them
    behind produces a definition that won't deploy.
    """
    import copy

    bim = copy.deepcopy(model.get("bim") or {})
    m = bim.get("model")
    if not isinstance(m, dict):
        raise ValueError("This .vpax has no usable Model.bim to rewrite.")

    drop_tables = set(plan["tables"])
    drop_cols: Dict[str, Set[str]] = {}
    for table, col in plan["columns"]:
        drop_cols.setdefault(table, set()).add(col)

    tables = [t for t in (m.get("tables") or [])
              if not (isinstance(t, dict) and str(t.get("name")) in drop_tables)]
    for t in tables:
        if not isinstance(t, dict):
            continue
        gone = drop_cols.get(str(t.get("name")), set())
        if gone:
            t["columns"] = [c for c in (t.get("columns") or [])
                            if not (isinstance(c, dict) and str(c.get("name")) in gone)]
    m["tables"] = tables

    def rel_ok(r: Any) -> bool:
        if not isinstance(r, dict):
            return False
        for tkey, ckey in (("fromTable", "fromColumn"), ("toTable", "toColumn")):
            table = str(r.get(tkey) or "")
            if table in drop_tables:
                return False
            if str(r.get(ckey) or "") in drop_cols.get(table, set()):
                return False
        return True

    if isinstance(m.get("relationships"), list):
        m["relationships"] = [r for r in m["relationships"] if rel_ok(r)]

    return json.dumps(bim, indent=2, default=str).encode("utf-8")


# ==========================================================================
# Databricks Lakeview (.lvdash.json) export
# ==========================================================================
# Everything here was reverse-engineered against three real exports of the
# same "Supplier Dashboard" - one from Power BI (.pbix/.vpax) and two
# generations of the matching Databricks Lakeview export - not against
# published Databricks documentation, since Lakeview's dashboard JSON schema
# isn't publicly documented the way PBIR's is. Treat the output as a strong
# first draft, not a guaranteed-correct import: verify against a real
# Lakeview import before trusting it in production, especially the
# relationship direction/cardinality (see note on `_lakeview_cardinality`)
# and the layout grid mapping (both are informed guesses, not confirmed
# Databricks semantics).

def _lakeview_id(*parts: str, length: int = 8) -> str:
    """A short, deterministic hex id in the style of Lakeview's own ids.

    Deterministic (hashed from the object's own identity, e.g. table name)
    rather than random, so re-running the generator on an unchanged model
    produces byte-identical output - a random id would make every export a
    full diff even when nothing actually changed.
    """
    import hashlib
    h = hashlib.md5("::".join(parts).encode("utf-8")).hexdigest()
    return h[:length]


def _lakeview_slug(name: Any) -> str:
    """snake_case identifier, matching the style Lakeview's own measure/
    dataset `name` fields use (e.g. `active_specifications`)."""
    s = re.sub(r"[^0-9A-Za-z]+", "_", str(name or "")).strip("_").lower()
    return s or "field"


# --- DAX -> SQL, deliberately narrow ------------------------------------------
# This is not a DAX parser. It recognises exactly the shapes this app's own
# analysis has already shown are common in real models - COALESCE-wrapped
# CALCULATE with simple equality filters around one aggregate - and refuses
# anything else outright rather than guessing. A wrong-but-plausible-looking
# SQL translation is worse than an honest "translate this by hand," because
# it fails silently in a stakeholder's dashboard instead of at review time.

_DAX_AGG_FUNCS: Dict[str, Tuple[str, bool]] = {
    # DAX name -> (SQL function, needs DISTINCT)
    "DISTINCTCOUNTNOBLANK": ("COUNT", True),
    "DISTINCTCOUNT": ("COUNT", True),
    "COUNTROWS": ("COUNT", False),
    "SUM": ("SUM", False),
    "AVERAGE": ("AVG", False),
    "MIN": ("MIN", False),
    "MAX": ("MAX", False),
    "COUNT": ("COUNT", False),
}


def _dax_strip_comments(expr: Any) -> str:
    s = str(expr or "")
    s = re.sub(r"/\*.*?\*/", " ", s, flags=re.S)
    s = re.sub(r"(--|//)[^\n]*", " ", s)
    return s.strip()


def _dax_matching_paren(s: str, open_idx: int) -> int:
    depth, i, in_str = 0, open_idx, False
    while i < len(s):
        c = s[i]
        if in_str:
            if c == '"':
                in_str = False
        elif c == '"':
            in_str = True
        elif c == "(":
            depth += 1
        elif c == ")":
            depth -= 1
            if depth == 0:
                return i
        i += 1
    return -1


def _dax_match_call(s: str, func_name: str) -> Optional[str]:
    """If `s` is exactly `FUNC( ... )` (nothing before or after), its inner text."""
    s = s.strip()
    m = re.match(rf"^{re.escape(func_name)}\s*\(", s, re.I)
    if not m:
        return None
    open_idx = m.end() - 1
    close_idx = _dax_matching_paren(s, open_idx)
    if close_idx == -1 or close_idx != len(s) - 1:
        return None
    return s[open_idx + 1:close_idx]


def _dax_split_args(s: str) -> List[str]:
    """Top-level comma-separated arguments, respecting parens/brackets/strings."""
    parts, depth, buf, in_str = [], 0, [], False
    for c in s:
        if in_str:
            buf.append(c)
            if c == '"':
                in_str = False
            continue
        if c == '"':
            in_str = True
            buf.append(c)
            continue
        if c in "([":
            depth += 1
        elif c in ")]":
            depth -= 1
        if c == "," and depth == 0:
            parts.append("".join(buf))
            buf = []
        else:
            buf.append(c)
    parts.append("".join(buf))
    return [p.strip() for p in parts if p.strip()]


def _dax_column_ref(s: str) -> Optional[str]:
    """Just the column name out of `'Table'[Col]` or `Table[Col]` or `[Col]`."""
    m = re.match(r"^'?[^'\[\]]*'?\[([^\]]+)\]$", s.strip())
    return m.group(1).strip() if m else None


def _dax_filter_to_sql(s: str) -> Optional[str]:
    """A single CALCULATE filter arg -> a SQL boolean, or None if not a plain
    `<column> = "literal"` / `<column> <> "literal"` comparison (optionally
    wrapped in KEEPFILTERS(...))."""
    s = s.strip()
    inner = _dax_match_call(s, "KEEPFILTERS")
    if inner is not None:
        s = inner.strip()
    m = re.match(r'^(.*?)\s*(=|<>)\s*"((?:[^"]|"")*)"\s*$', s, re.S)
    if not m:
        return None
    colref, op, literal = m.groups()
    col = _dax_column_ref(colref)
    if col is None:
        return None
    return f"`{col}` {op} '{literal.replace(chr(39), chr(39)*2)}'"


def translate_dax_measure(expr: Any) -> Tuple[Optional[str], str]:
    """Best-effort DAX -> SQL for one measure. Returns (sql_or_None, note).

    Supported shape: optional `COALESCE(<agg>, <default>)` wrapping an
    optional `CALCULATE(<aggregate>, <filter>, <filter>, ...)`, where each
    filter is a plain equality/inequality against a string literal, and the
    aggregate is one of the functions in `_DAX_AGG_FUNCS` applied to a single
    column (or, for COUNTROWS, a bare table with no nested filtering - a
    filtered CALCULATETABLE/FILTER argument is refused rather than silently
    dropped, since dropping it would understate the count).
    """
    s = _dax_strip_comments(expr)
    if not s:
        return None, "Empty or fully commented-out expression."

    default_sql = None
    coalesce_inner = _dax_match_call(s, "COALESCE")
    if coalesce_inner is not None:
        args = _dax_split_args(coalesce_inner)
        if len(args) == 2:
            s, default_sql = args[0], args[1]
        elif args:
            s = args[0]

    filters_sql: List[str] = []
    agg_expr = s
    calc_inner = _dax_match_call(s, "CALCULATE")
    if calc_inner is not None:
        args = _dax_split_args(calc_inner)
        if not args:
            return None, "CALCULATE() with no arguments."
        agg_expr = args[0]
        for f in args[1:]:
            fsql = _dax_filter_to_sql(f)
            if fsql is None:
                return None, f"CALCULATE filter isn't a plain equality: `{f[:70]}`"
            filters_sql.append(fsql)

    match = None
    for dax_name, (sql_func, needs_distinct) in _DAX_AGG_FUNCS.items():
        arg = _dax_match_call(agg_expr, dax_name)
        if arg is not None:
            match = (dax_name, sql_func, needs_distinct, arg)
            break
    if match is None:
        return None, f"Unsupported aggregate or expression shape: `{agg_expr[:70]}`"
    dax_name, sql_func, needs_distinct, arg = match
    where = " AND ".join(filters_sql) if filters_sql else None

    if dax_name == "COUNTROWS":
        if not re.match(r"^'?[^'()]+'?$", arg.strip()):
            return None, "COUNTROWS wraps a filtered table expression, not a bare table."
        sql = f"COUNT(CASE WHEN {where} THEN 1 END)" if where else "COUNT(*)"
    else:
        col = _dax_column_ref(arg)
        if col is None:
            return None, f"Aggregate argument isn't a plain column reference: `{arg[:70]}`"
        distinct = "DISTINCT " if needs_distinct else ""
        sql = (f"{sql_func}({distinct}CASE WHEN {where} THEN `{col}` END)" if where
               else f"{sql_func}({distinct}`{col}`)")

    if default_sql is not None:
        sql = f"COALESCE({sql}, {default_sql})"
    return sql, "Translated."


# --- Table -> dataset -----------------------------------------------------

def _lakeview_measure_home_table(
    table: str, expr: str, measure_group_tables: Set[str], all_table_names: List[str]
) -> str:
    """Where a measure's SQL should live.

    A PBIX "measure group" table (a hidden table that exists only to hold
    per-page measures, has no columns of its own, and therefore has no real
    counterpart in the warehouse) can't host a Lakeview measure - Lakeview
    measures belong to a dataset backed by a real table. Retarget to the
    first real table the measure's own DAX references; if it references
    none (a constant, or something this app can't resolve), the caller drops
    it and the conversion report explains why.
    """
    if table not in measure_group_tables:
        return table
    for candidate in all_table_names:
        if candidate == table or candidate in measure_group_tables:
            continue
        if find_referenced_tables(expr, [candidate]):
            return candidate
    return ""


def build_lakeview_datasets(
    model: Dict[str, Any], catalog: str, schema: str, table_name_map: Dict[str, str]
) -> Tuple[List[Dict[str, Any]], Dict[str, str], Dict[Tuple[str, str], str], pd.DataFrame]:
    """One flat, single-table dataset per model table (plus the model's own
    relationships, built separately) - not a pre-joined tree.

    That choice is deliberate, not a simplification for its own sake: the
    live experiment run against this exact model showed that once a
    `relationshipGraphs` block exists, Lakeview resolves joins across
    datasets at query time, and a dataset's own nested `joins[]` tree is only
    needed to *expose a joined table's columns as if they belonged to it* -
    which this generator never needs, because every widget binding is
    resolved back to the real table that owns the field.

    Returns (datasets, {table_name: dataset_id}, {(table, measure_name): sql_name},
    conversion_report).
    """
    report_rows: List[Dict[str, Any]] = []
    datasets: List[Dict[str, Any]] = []
    table_to_dataset: Dict[str, str] = {}
    measure_sql_names: Dict[Tuple[str, str], str] = {}

    bim_tables = model.get("bim_tables") or []
    measure_group_tables = {t["name"] for t in bim_tables if _is_measure_group(t) and t.get("name")}
    all_tables = model["all_table_names"]
    cols_df = _user_facing_columns(model["columns"])
    meas_df = model["measures"]

    for table in all_tables:
        if table in measure_group_tables:
            continue  # no real backing table - its measures are retargeted below
        ds_id = _lakeview_id("dataset", table)
        table_to_dataset[table] = ds_id
        db_name = table_name_map.get(table, _lakeview_slug(table))

        dimensions = []
        if {"TableName", "ColumnName"}.issubset(cols_df.columns):
            for _, c in cols_df.loc[cols_df["TableName"] == table].dropna(subset=["ColumnName"]).iterrows():
                col = str(c["ColumnName"])
                dimensions.append({
                    "name": col, "expr": f"source.{col}",
                    "displayName": col,
                    "comment": str(c.get("Description") or ""),
                })

        measures = [{"name": "count", "expr": "COUNT(*)",
                     "comment": "Total row count.", "displayName": "Count"}]
        if {"TableName", "MeasureName", "MeasureExpression"}.issubset(meas_df.columns):
            for _, m in meas_df.dropna(subset=["MeasureName"]).iterrows():
                home = _lakeview_measure_home_table(
                    str(m["TableName"]), str(m.get("MeasureExpression") or ""),
                    measure_group_tables, all_tables,
                )
                if home != table:
                    continue
                sql, note = translate_dax_measure(m.get("MeasureExpression"))
                sql_name = _lakeview_slug(m["MeasureName"])
                report_rows.append({
                    "Object Type": "Measure", "Table": table, "Name": str(m["MeasureName"]),
                    "Status": "Translated" if sql else "Not translated",
                    "Severity": "Info" if sql else "Medium", "Detail": note,
                })
                if sql:
                    measures.append({
                        "name": sql_name, "expr": sql,
                        "comment": (str(m.get("Description") or "").strip()
                                   or f"Translated from DAX measure `{m['MeasureName']}` — verify against the source report."),
                        "displayName": str(m["MeasureName"]),
                    })
                    measure_sql_names[(str(m["TableName"]), str(m["MeasureName"]))] = sql_name

        datasets.append({
            "name": ds_id, "displayName": table,
            "config": {
                "version": "1.1",
                "source": f"{catalog}.{schema}.{db_name}",
                "dimensions": dimensions,
                "measures": measures,
            },
        })

    # Measures whose home table couldn't be resolved at all (no real table
    # referenced anywhere in their DAX) - report, don't silently drop.
    if {"TableName", "MeasureName", "MeasureExpression"}.issubset(meas_df.columns):
        for _, m in meas_df.dropna(subset=["MeasureName"]).iterrows():
            table = str(m["TableName"])
            if table not in measure_group_tables:
                continue
            home = _lakeview_measure_home_table(
                table, str(m.get("MeasureExpression") or ""), measure_group_tables, all_tables,
            )
            if not home:
                report_rows.append({
                    "Object Type": "Measure", "Table": table, "Name": str(m["MeasureName"]),
                    "Status": "Not translated", "Severity": "Medium",
                    "Detail": "This measure lives on a report-only measure-group table and its "
                              "DAX doesn't clearly reference a real model table — assign it to a "
                              "dataset by hand.",
                })

    report_df = sort_by_severity(_ensure_columns(
        pd.DataFrame(report_rows), ["Object Type", "Table", "Name", "Status", "Severity", "Detail"]
    ))
    return datasets, table_to_dataset, measure_sql_names, report_df


def _lakeview_cardinality(from_card: str, to_card: str) -> str:
    """PBIX one/many per side -> Lakeview's CARDINALITY_* enum.

    Caveat, stated plainly: the one real example available (the second
    Lakeview export inspected in this session) records `from`/`to` in the
    direction the join tree happened to be authored in, which did not match
    the intuitive "fact is many, dimension is one" reading of the equivalent
    PBIX relationship. Whether Databricks treats `from`/`to` as directional
    in a way that affects query results, or purely as a label, is unconfirmed
    - verify the generated direction against a real Lakeview import rather
    than trusting this mapping blindly.
    """
    f, t = from_card.lower(), to_card.lower()
    if f.startswith("many") and t.startswith("one"):
        return "CARDINALITY_MANY_TO_ONE"
    if f.startswith("one") and t.startswith("many"):
        return "CARDINALITY_ONE_TO_MANY"
    return "CARDINALITY_ONE_TO_ONE"


def build_lakeview_relationship_graph(
    model: Dict[str, Any], table_to_dataset: Dict[str, str]
) -> Tuple[List[Dict[str, Any]], pd.DataFrame]:
    rels = model["relationships"]
    report_rows: List[Dict[str, Any]] = []
    if rels.empty or not table_to_dataset:
        return [], _ensure_columns(pd.DataFrame(), ["Object Type", "Table", "Name", "Status", "Severity", "Detail"])

    included_tables = sorted(table_to_dataset)
    sources = [{"name": t, "datasetName": table_to_dataset[t]} for t in included_tables]
    relationships = []
    for _, r in rels.iterrows():
        ft, tt = str(r.get("From Table") or ""), str(r.get("To Table") or "")
        fc, tc = str(r.get("From Column") or ""), str(r.get("To Column") or "")
        if not ft or not tt or ft == tt:
            continue
        if ft not in table_to_dataset or tt not in table_to_dataset:
            report_rows.append({
                "Object Type": "Relationship", "Table": f"{ft} → {tt}", "Name": f"{fc} = {tc}",
                "Status": "Not migrated", "Severity": "Medium",
                "Detail": "One endpoint's table has no generated dataset (its home table is a "
                          "measure-group table, or it wasn't in the model's table list).",
            })
            continue
        relationships.append({
            "name": _lakeview_id("rel", ft, fc, tt, tc, length=20),
            "from": ft, "to": tt,
            "on": f"`{ft}`.`{fc}` = `{tt}`.`{tc}`",
            "cardinality": _lakeview_cardinality(
                str(r.get("From Cardinality") or "Many"), str(r.get("To Cardinality") or "One")
            ),
        })
        if str(r.get("Cross Filter Direction")) == "Both":
            report_rows.append({
                "Object Type": "Relationship", "Table": f"{ft} ↔ {tt}", "Name": f"{fc} = {tc}",
                "Status": "Migrated — verify direction", "Severity": "Low",
                "Detail": "Bi-directional in the PBIX model. Lakeview's relationship graph has "
                          "no cross-filter-direction equivalent — any widget that depended on "
                          "the filter travelling from the 'many' side back up to the 'one' side "
                          "needs manual verification.",
            })

    graph = [{"sources": sources, "relationships": relationships}] if relationships else []
    return graph, sort_by_severity(_ensure_columns(
        pd.DataFrame(report_rows), ["Object Type", "Table", "Name", "Status", "Severity", "Detail"]
    ))


# --- Report layer: PBIR pages/visuals -> Lakeview pages/widgets -----------

_LAKEVIEW_WIDGET_MAP: Dict[str, str] = {
    "card": "counter", "multiRowCard": "counter",
    "barChart": "bar", "clusteredColumnChart": "bar", "columnChart": "bar",
    "clusteredBarChart": "bar", "lineChart": "line", "lineClusteredColumnComboChart": "line",
    "pieChart": "pie", "donutChart": "pie",
    "tableEx": "table", "pivotTable": "table",
    "slicer": "filter-multi-select",
}
_LAKEVIEW_UNSUPPORTED_NOTE = (
    "No Lakeview widget type corresponds to this visual — it will need to be "
    "rebuilt by hand, or dropped, in the Databricks dashboard."
)


def _pbir_page_dirs(z: zipfile.ZipFile) -> Dict[str, str]:
    """{pbix page displayName: internal page-folder id} for PBIR-format reports."""
    norm = {n.replace("\\", "/"): n for n in z.namelist()}
    index_key = next((k for k in norm if k.endswith("Report/definition/pages/pages.json")), None)
    if index_key is None:
        return {}
    prefix = index_key.rsplit("/", 1)[0] + "/"
    out = {}
    for key, orig in norm.items():
        if key.startswith(prefix) and key.endswith("/page.json"):
            page_id = key[len(prefix):].split("/", 1)[0]
            try:
                page = _read_json_member(z, orig)
            except (json.JSONDecodeError, UnicodeDecodeError):
                continue
            out[str(page.get("displayName") or page.get("name") or page_id)] = page_id
    return out


def _pbir_page_visuals(z: zipfile.ZipFile, page_id: str) -> Tuple[int, int, List[Dict[str, Any]]]:
    """(page width, page height, [visual dict, ...]) for one PBIR page."""
    norm = {n.replace("\\", "/"): n for n in z.namelist()}
    page_key = next(
        (orig for key, orig in norm.items()
         if key.endswith(f"Report/definition/pages/{page_id}/page.json")), None
    )
    if page_key is None:
        return 1280, 720, []
    page = _read_json_member(z, page_key)
    width, height = int(page.get("width") or 1280), int(page.get("height") or 720)

    prefix = None
    for key in norm:
        if key.endswith(f"Report/definition/pages/{page_id}/page.json"):
            prefix = key.rsplit("/", 1)[0] + "/visuals/"
            break
    visuals = []
    if prefix:
        for key, orig in norm.items():
            if key.startswith(prefix) and key.endswith("/visual.json"):
                try:
                    visuals.append(_read_json_member(z, orig))
                except (json.JSONDecodeError, UnicodeDecodeError):
                    continue
    return width, height, visuals


def _visual_title(visual_def: Dict[str, Any], fallback: str) -> str:
    """A human-readable title for a visual that (usually) never explicitly set one.

    PBIX visuals are frequently untitled - Power BI falls back to auto-
    generated text at render time, which isn't stored in the file. This
    checks for an explicit title override first, then borrows the first
    field's own display name as a reasonable stand-in, and only falls back
    to the visual's internal id if neither exists.
    """
    visual = visual_def.get("visual") or {}
    try:
        lit = visual["visualContainerObjects"]["title"][0]["properties"]["text"]["expr"]["Literal"]["Value"]
        if isinstance(lit, str) and lit.strip("'\""):
            return lit.strip("'\"")
    except (KeyError, IndexError, TypeError):
        pass

    def find_display(node: Any) -> Optional[str]:
        if isinstance(node, dict):
            for key in ("displayName", "nativeQueryRef"):
                v = node.get(key)
                if isinstance(v, str) and v.strip():
                    return v
            for val in node.values():
                found = find_display(val)
                if found:
                    return found
        elif isinstance(node, list):
            for item in node:
                found = find_display(item)
                if found:
                    return found
        return None

    return find_display(visual.get("query") or {}) or fallback


def _visual_text(visual_def: Dict[str, Any]) -> Optional[str]:
    """The literal text out of a textbox visual, or None if this isn't one."""
    v = visual_def.get("visual") or {}
    if str(v.get("visualType")) != "textbox":
        return None
    paragraphs = ((v.get("objects") or {}).get("general") or [{}])[0].get("properties", {}).get("paragraphs", [])
    lines = []
    for p in paragraphs:
        lines.append("".join(str(r.get("value") or "") for r in p.get("textRuns") or []))
    return "\n".join(lines)


def _grid_position(x: float, y: float, w: float, h: float, page_w: int, row_px: float) -> Dict[str, int]:
    """Absolute PBIX pixels -> Lakeview's 12-column relative grid.

    A linear bin, not pixel-perfect: PBIX layouts are freeform (arbitrary
    overlap, z-order) and Lakeview's grid is stacked and non-overlapping, so
    some manual nudging after import should be expected regardless of how
    this function is tuned.
    """
    col_w = max(page_w, 1) / 12
    gx = max(0, min(11, round(x / col_w)))
    gw = max(1, min(12 - gx, round(w / col_w)))
    gy = max(0, round(y / row_px))
    gh = max(1, round(h / row_px))
    return {"x": gx, "y": gy, "width": gw, "height": gh}


def build_lakeview_pages(
    model: Dict[str, Any], pbix_bytes: bytes, chosen_pages: List[str],
    table_to_dataset: Dict[str, str], measure_sql_names: Dict[Tuple[str, str], str],
    row_px: float = 40.0,
) -> Tuple[List[Dict[str, Any]], pd.DataFrame]:
    report_rows: List[Dict[str, Any]] = []
    pages_out: List[Dict[str, Any]] = []

    with zipfile.ZipFile(io.BytesIO(pbix_bytes)) as z:
        page_dirs = _pbir_page_dirs(z)
        if not page_dirs:
            report_rows.append({
                "Object Type": "Page", "Table": "", "Name": "(whole report)",
                "Status": "Not migrated", "Severity": "Medium",
                "Detail": "This .pbix uses the classic 'Report/Layout' format. Page/widget "
                          "generation currently only supports the newer PBIR format (Save As, "
                          "or File ➜ Options ➜ Preview features ➜ 'Power BI Project' in Desktop). "
                          "Datasets and relationships above are unaffected.",
            })
            return [], _ensure_columns(pd.DataFrame(report_rows),
                                       ["Object Type", "Table", "Name", "Status", "Severity", "Detail"])

        for page_name in chosen_pages:
            page_id = page_dirs.get(page_name)
            if page_id is None:
                continue
            page_w, page_h, visuals = _pbir_page_visuals(z, page_id)
            layout = []
            for v in visuals:
                name = str(v.get("name") or "")
                pos = v.get("position") or {}
                x, y = float(pos.get("x", 0)), float(pos.get("y", 0))
                w, h = float(pos.get("width", 100)), float(pos.get("height", 100))
                grid = _grid_position(x, y, w, h, page_w, row_px)

                visual = v.get("visual") or {}
                v_type = visual.get("visualType")

                text = _visual_text(v)
                if text is not None:
                    layout.append({
                        "widget": {"name": _lakeview_slug(text.splitlines()[0] if text else name) or name,
                                  "multilineTextboxSpec": {"lines": [line + "\n" for line in text.splitlines()] or [""]}},
                        "position": grid,
                    })
                    continue

                if "visualGroup" in v:
                    continue  # a layout container only, not a widget - nothing to migrate

                widget_type = _LAKEVIEW_WIDGET_MAP.get(str(v_type))
                if widget_type is None:
                    report_rows.append({
                        "Object Type": "Visual", "Table": page_name, "Name": f"{name} ({v_type})",
                        "Status": "Not migrated", "Severity": "Medium",
                        "Detail": _LAKEVIEW_UNSUPPORTED_NOTE,
                    })
                    continue

                bindings = []
                seen: Set[Tuple[str, str, str]] = set()
                for table, field, kind in _collect_field_refs(v):
                    if not field or table not in table_to_dataset:
                        continue
                    key = (table, field, kind)
                    if key in seen:
                        continue
                    seen.add(key)
                    if kind == "Measure":
                        sql_name = measure_sql_names.get((table, field))
                        if sql_name is None:
                            continue  # untranslated measure - already reported by the dataset builder
                        bindings.append((table, field, "Measure", sql_name))
                    else:
                        bindings.append((table, field, "Column", field))

                title = _visual_title(v, name)

                if not bindings:
                    report_rows.append({
                        "Object Type": "Visual", "Table": page_name, "Name": f"{name} ({v_type})",
                        "Status": "Not migrated", "Severity": "Medium",
                        "Detail": "None of this visual's fields could be resolved to a generated "
                                  "dataset (every one was either an untranslated measure or on a "
                                  "table outside this generator's scope).",
                    })
                    continue

                by_table: Dict[str, List[Tuple[str, str, str, str]]] = {}
                for b in bindings:
                    by_table.setdefault(b[0], []).append(b)
                queries = []
                field_query: Dict[str, str] = {}
                for i, (table, items) in enumerate(by_table.items()):
                    qname = "main_query" if len(by_table) == 1 else f"{_lakeview_slug(table)}_query"
                    fields = []
                    for _, disp_name, kind, sql_name in items:
                        expr = f"MEASURE(`{sql_name}`)" if kind == "Measure" else f"`{sql_name}`"
                        field_name = f"measure({sql_name})" if kind == "Measure" else sql_name
                        fields.append({"name": field_name, "expression": expr})
                        field_query[field_name] = qname
                    queries.append({"name": qname, "query": {
                        "datasetName": table_to_dataset[table], "fields": fields, "disaggregated": False,
                    }})

                all_field_names = list(field_query)
                measure_fields = [f for f in all_field_names if f.startswith("measure(")]
                column_fields = [f for f in all_field_names if not f.startswith("measure(")]

                if widget_type == "counter":
                    value_field = (measure_fields or column_fields or [None])[0]
                    if value_field is None:
                        continue
                    spec = {"version": 2, "frame": {"showTitle": True, "title": title},
                           "widgetType": "counter",
                           "encodings": {"value": {"fieldName": value_field}},
                           "data": {"queryName": field_query[value_field]}}
                elif widget_type in ("bar", "line", "pie"):
                    x_field = (measure_fields or all_field_names or [None])[0]
                    y_field = (column_fields or [f for f in all_field_names if f != x_field] or [None])[0]
                    if x_field is None or y_field is None:
                        continue
                    spec = {"version": 3, "frame": {"showTitle": True, "title": title},
                           "widgetType": widget_type,
                           "encodings": {
                               "x": {"fieldName": x_field, "scale": {"type": "quantitative"}},
                               "y": {"fieldName": y_field, "scale": {"type": "categorical"}},
                           },
                           "data": {"queryName": field_query[x_field]}}
                    if len(all_field_names) > 2:
                        report_rows.append({
                            "Object Type": "Visual", "Table": page_name, "Name": name,
                            "Status": "Migrated — verify encodings", "Severity": "Low",
                            "Detail": f"{len(all_field_names)} fields bound on this visual; only "
                                      "the first measure/column pair was mapped to x/y — check the "
                                      "remaining fields.",
                        })
                elif widget_type == "table":
                    spec = {"version": 2, "frame": {"showTitle": True, "title": title},
                           "widgetType": "table",
                           "encodings": {"columns": [
                               {"fieldName": f, "useForSearch": False, "displayName": f}
                               for f in all_field_names
                           ]},
                           "data": {"queryName": queries[0]["name"] if len(queries) == 1 else queries[0]["name"]}}
                    if len(queries) > 1:
                        report_rows.append({
                            "Object Type": "Visual", "Table": page_name, "Name": name,
                            "Status": "Migrated — verify columns", "Severity": "Low",
                            "Detail": "This table binds fields from more than one dataset — only "
                                      "the first dataset's query is wired as the table's data "
                                      "source; the rest need manual reconciliation.",
                        })
                else:  # filter-multi-select
                    field_name = all_field_names[0]
                    qname = field_query[field_name]
                    queries[0]["query"]["fields"].append({
                        "name": f"{field_name}_associativity",
                        "expression": "COUNT_IF(`associative_filter_predicate_group`)",
                    })
                    spec = {"version": 2, "frame": {"showTitle": True, "title": title},
                           "widgetType": widget_type,
                           "encodings": {"fields": [{"fieldName": field_name, "queryName": qname}]}}

                layout.append({
                    "widget": {"name": _lakeview_id("widget", page_name, name, length=20),
                              "queries": queries, "spec": spec},
                    "position": grid,
                })

            if layout:
                pages_out.append({
                    "name": _lakeview_id("page", page_name), "displayName": page_name,
                    "pageType": "PAGE_TYPE_CANVAS", "layoutVersion": "GRID_V1", "layout": layout,
                })

    return pages_out, sort_by_severity(_ensure_columns(
        pd.DataFrame(report_rows), ["Object Type", "Table", "Name", "Status", "Severity", "Detail"]
    ))


def build_lakeview_theme(report: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    default = {
        "canvasBackgroundColor": {"light": "#F5F5F5", "dark": "#1F272D"},
        "widgetBackgroundColor": {"light": "#FFFFFF", "dark": "#2C3640"},
        "fontColor": {"light": "#424242", "dark": "#E0E0E0"},
        "selectionColor": {"light": "#2196F3", "dark": "#2196F3"},
        "visualizationColors": ["#2196F3", "#4CAF50", "#FF9800", "#F44336", "#9C27B0",
                                "#00BCD4", "#FFEB3B", "#795548", "#607D8B", "#E91E63"],
        "fontFamily": "Inter", "widgetPadding": 12, "widgetCornerRadius": 8, "widgetShadow": 8,
    }
    if not report:
        return default
    theme = (report.get("theme") or {})
    if not theme.get("found"):
        return default
    tj = theme.get("json") or {}
    data_colors = tj.get("dataColors") or []
    out = dict(default)
    if tj.get("background"):
        out["canvasBackgroundColor"] = {"light": tj["background"], "dark": default["canvasBackgroundColor"]["dark"]}
    if tj.get("foreground"):
        out["fontColor"] = {"light": tj["foreground"], "dark": default["fontColor"]["dark"]}
    if data_colors:
        out["visualizationColors"] = (data_colors + default["visualizationColors"])[:10]
    return out


def build_lakeview_dashboard(
    model: Dict[str, Any], catalog: str, schema: str, table_name_map: Dict[str, str],
    pbix_bytes: Optional[bytes], chosen_pages: List[str], report: Optional[Dict[str, Any]],
    include_relationships: bool = True,
) -> Tuple[Dict[str, Any], pd.DataFrame]:
    datasets, table_to_dataset, measure_sql_names, ds_report = build_lakeview_datasets(
        model, catalog, schema, table_name_map
    )
    reports = [ds_report]

    dashboard: Dict[str, Any] = {"datasets": datasets}

    if include_relationships:
        graph, rel_report = build_lakeview_relationship_graph(model, table_to_dataset)
        if graph:
            dashboard["relationshipGraphs"] = graph
        reports.append(rel_report)

    pages: List[Dict[str, Any]] = []
    if pbix_bytes and chosen_pages:
        pages, page_report = build_lakeview_pages(
            model, pbix_bytes, chosen_pages, table_to_dataset, measure_sql_names
        )
        reports.append(page_report)
    dashboard["pages"] = pages or [{
        "name": _lakeview_id("page", "overview"), "displayName": "Overview",
        "pageType": "PAGE_TYPE_CANVAS", "layoutVersion": "GRID_V1", "layout": [],
    }]

    dashboard["uiSettings"] = {"theme": build_lakeview_theme(report), "applyModeEnabled": False}

    report_df = pd.concat(reports, ignore_index=True) if reports else pd.DataFrame()
    report_df = sort_by_severity(_ensure_columns(
        report_df, ["Object Type", "Table", "Name", "Status", "Severity", "Detail"]
    ))
    return dashboard, report_df


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


def build_formulas_sheet(model: Dict[str, Any]) -> pd.DataFrame:
    """Every measure and calculated column in one list, each row tagged with
    its Type - so "what DAX exists in this model" is answered by one sheet
    instead of cross-referencing two. Measures are listed first, then
    calculated columns, matching how most modelers think about a model
    (measures are the primary calculation layer; calculated columns are the
    exception).
    """
    meas_df, calc_df = model["measures"], model["calc_columns"]
    rows: List[Dict[str, str]] = []

    for _, r in meas_df.iterrows():
        rows.append({
            "Type": "Measure",
            "Table": r.get("TableName", ""),
            "Name": r.get("MeasureName", ""),
            "Expression": r.get("MeasureExpression", ""),
            "DataType": r.get("DataType", ""),
            "FormatString": r.get("FormatString", ""),
            "Description": r.get("Description", ""),
        })
    for _, r in calc_df.iterrows():
        rows.append({
            "Type": "Calculated Column",
            "Table": r.get("TableName", ""),
            "Name": r.get("ColumnName", ""),
            "Expression": r.get("ColumnExpression", ""),
            "DataType": r.get("DataType", ""),
            "FormatString": "",
            "Description": "",
        })

    return _ensure_columns(
        pd.DataFrame(rows),
        ["Type", "Table", "Name", "Expression", "DataType", "FormatString", "Description"],
    )


def build_data_dictionary_excel(
    model: Dict[str, Any],
    health_df: pd.DataFrame,
    naming_df: pd.DataFrame,
    unused_df: pd.DataFrame,
    extra_sheets: Optional[List[Tuple[str, pd.DataFrame]]] = None,
) -> bytes:
    """One workbook, one sheet per topic - the whole model plus this app's
    health/naming/unused-object findings, in the format people actually want
    to filter, pivot, and share: Excel, not a Word doc.

    Diagrams aren't embedded - the app's ER and measure-dependency diagrams
    are rendered entirely client-side (viz.js in the browser), and there's no
    server-side render path to produce an image from here.
    """
    if EXCEL_ENGINE is None:
        raise RuntimeError("No Excel writer installed (pip install openpyxl)")

    sheets: List[Tuple[str, pd.DataFrame]] = [
        ("Tables", model["tables"]),
        ("Columns", model["columns"]),
        ("Measures & Calc Columns", build_formulas_sheet(model)),
        ("Relationships", model["relationships"]),
        ("Power Query (SQL)", model["power_query"]),
        ("Roles (RLS)", model["roles"]),
        ("Perspectives", model["perspectives"]),
        ("Date Tables", model["date_tables"]),
        ("Model Health", health_df),
        ("Naming Conventions", naming_df),
        ("Unused Columns", unused_df),
    ]
    sheets.extend(extra_sheets or [])

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine=EXCEL_ENGINE) as writer:
        used_names: Set[str] = set()
        for name, df in sheets:
            safe_name = re.sub(r"[:\\/?*\[\]]", "-", name)[:31] or "Sheet1"
            base, i = safe_name, 2
            while safe_name in used_names:
                suffix = f" ({i})"
                safe_name = base[: 31 - len(suffix)] + suffix
                i += 1
            used_names.add(safe_name)
            out_df = df if not df.empty else pd.DataFrame([{"Info": "None found."}])
            _safe(out_df).to_excel(writer, index=False, sheet_name=safe_name)
    return buf.getvalue()


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


def show_table(
    df: pd.DataFrame, name: str, height: int = 380, key: str = "", row_height: int = 68
) -> None:
    """Render a dataframe with CSV + XLSX download buttons.

    `row_height` turns on word-wrap: Streamlit's dataframe grid only wraps
    text within a row when given an explicit pixel height taller than one
    line - left at its default (auto), long text is silently truncated with
    no ellipsis. There's no per-row auto-height in this grid, so every row
    in a table shares one height; callers with especially long text columns
    (recommendations, DAX, multi-sentence explanations) should pass a taller
    value so their longest cells don't get clipped.
    """
    safe = _safe(df)
    st.dataframe(safe, width="stretch", height=height, hide_index=True, row_height=row_height)

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
# AI DAX Assistant (bring-your-own-key LLM regeneration)
# ==========================================================================
# Nothing here runs without the user explicitly supplying their own API key
# for this browser session. The key, the DAX, and the schema summary sent to
# the provider are never written to disk - the key lives only in
# st.session_state (server-side, in-memory, cleared when the session ends).
# Regeneration is advisory only: this app has no way to write back into a
# .vpax or Model.bim, so the result is something to review and copy into
# Tabular Editor or Power BI Desktop yourself - never applied automatically.

# Model lists are curated, not fetched live - verified against each
# provider's own docs. First entry in "models" is the pre-selected default
# (picked for low cost / good-enough quality, matching this app's use case:
# short, well-defined DAX/insight generations, not long agentic sessions).
LLM_PROVIDERS: Dict[str, Dict[str, object]] = {
    "OpenAI (ChatGPT)": {
        "models": ["gpt-5.6-luna", "gpt-5.6-terra", "gpt-5.6", "gpt-4o-mini", "gpt-4o"],
        "key_placeholder": "sk-…",
    },
    "Anthropic (Claude)": {
        "models": ["claude-sonnet-5", "claude-opus-5", "claude-haiku-4-5", "claude-fable-5"],
        "key_placeholder": "sk-ant-…",
    },
    "DeepSeek": {
        "models": ["deepseek-v4-flash", "deepseek-v4-pro"],
        "key_placeholder": "sk-…",
    },
    "Llama (via Groq)": {
        "models": ["llama-3.3-70b-versatile", "llama-3.1-8b-instant"],
        "key_placeholder": "gsk_…",
    },
    "Google (Gemini)": {
        "models": ["gemini-3.5-flash-lite", "gemini-3.6-flash", "gemini-3.1-pro-preview"],
        "key_placeholder": "AIza…",
    },
}


class LLMError(Exception):
    """Any problem calling an LLM provider - always caught and shown via
    st.error, never left to crash the page."""


def _compact_schema_summary(model: Dict[str, Any], max_tables: int = 40, max_cols_per_table: int = 25) -> str:
    """A bounded, plain-text schema summary for an LLM prompt.

    A full schema on a large model can run to hundreds of columns - capping
    both dimensions keeps prompt size (and the cost/latency of the call)
    predictable without silently truncating mid-table.
    """
    cols_by_table: Dict[str, List[str]] = {}
    cdf = _user_facing_columns(model["columns"])
    if {"TableName", "ColumnName"}.issubset(cdf.columns):
        for _, r in cdf.dropna(subset=["ColumnName"]).iterrows():
            cols_by_table.setdefault(str(r["TableName"]), []).append(str(r["ColumnName"]))

    all_tables = model["all_table_names"]
    lines = []
    for t in all_tables[:max_tables]:
        cols = cols_by_table.get(t, [])
        shown = cols[:max_cols_per_table]
        extra = f", … +{len(cols) - len(shown)} more" if len(cols) > len(shown) else ""
        lines.append(f"- {t}: {', '.join(shown)}{extra}" if shown else f"- {t}")
    if len(all_tables) > max_tables:
        lines.append(f"… +{len(all_tables) - max_tables} more tables")

    rel_lines = []
    rel_df = model["relationships"]
    for _, r in rel_df.head(60).iterrows():
        rel_lines.append(
            f"{r['From Table']}[{r['From Column']}] -> {r['To Table']}[{r['To Column']}] "
            f"({r['From Cardinality']}-to-{r['To Cardinality']}, "
            f"{'bi-directional' if r['Cross Filter Direction'] == 'Both' else 'single-direction'}"
            f"{'' if r['Active'] else ', inactive'})"
        )
    other_measures = sorted(model["measure_names"])[:80]

    return (
        "Tables and columns:\n" + "\n".join(lines)
        + "\n\nRelationships:\n" + ("\n".join(rel_lines) if rel_lines else "(none)")
        + "\n\nOther measure names available to reference:\n"
        + (", ".join(other_measures) if other_measures else "(none)")
    )


def build_dax_regeneration_prompt(
    object_kind: str, table: str, name: str, expression: str, model: Dict[str, Any]
) -> str:
    schema = _compact_schema_summary(model)
    return f"""You are a senior Power BI / DAX consultant reviewing one {object_kind} in an existing semantic model.

{object_kind.capitalize()} name: {name}
Home table: {table}
Current DAX:
{expression}

Model schema (for context - do not invent tables/columns/measures that aren't listed):
{schema}

Rewrite this {object_kind}'s DAX to follow current DAX best practices (e.g. fully-qualified column
references, un-qualified measure references, DIVIDE() instead of "/", explicit CALCULATE filters
rather than relying on implicit row context where it helps clarity, avoiding unnecessary iterators).
Specifically check for and fix these well-known engine-level anti-patterns where they apply:
- CALCULATE (or any measure reference, which triggers an implicit CALCULATE) inside a row iterator
  like SUMX/AVERAGEX/FILTER over a large table forces a context transition on every row - replace
  with a native scalar expression over the iterator's own row context where the logic allows it
  (e.g. SUMX(Sales, Sales[Qty] * Sales[Price]) instead of SUMX(Sales, CALCULATE(SUM(Sales[Qty])) * Sales[Price])).
- A filter argument inside CALCULATE that should narrow an *existing* filter rather than replace it
  (CALCULATE implicitly does ALL() on the filtered column) needs KEEPFILTERS to intersect instead of
  override.
- A sub-expression referenced more than once anywhere in the measure (across IF/SWITCH branches or
  repeated elsewhere) should be hoisted into a VAR so it's evaluated once and short-circuits, not
  recomputed per reference.
- Row-by-row text matching (SEARCH, CONTAINSSTRING) or complex conditional logic inside a
  high-cardinality iterator forces expensive Storage-Engine-to-Formula-Engine callbacks
  (CallbackDataID) - flag this in notes even if you can't fully eliminate it, since it often means
  the logic belongs upstream in the data model instead.
Keep the result logically equivalent unless you believe the original has an actual bug - call that
out explicitly in your notes if so.

Also consider whether the *data model* itself (not just this expression) could be improved to make
this {object_kind} simpler, faster, or more reliable - e.g. a missing relationship, a column that
should be added upstream, a role this {object_kind} plays that suggests a modelling change. Only
suggest changes grounded in the schema above - never invent objects that aren't listed.

Respond with ONLY a single JSON object (no markdown fences, no commentary outside the JSON), with
exactly these keys:
{{
  "revised_dax": "the rewritten DAX expression, or the original if no change is warranted",
  "notes": "a short explanation of what changed and why, or why nothing changed",
  "model_suggestions": "any data-modelling changes worth considering, or an empty string if none"
}}"""


def build_insights_prompt(model: Dict[str, Any], score: int, grade: str, cards: List[Dict[str, Any]]) -> str:
    """Synthesize the Scorecard's already-computed per-section findings into
    one prompt - the LLM prioritizes and explains, it never re-derives the
    findings themselves, so it can't invent a problem the rule engine didn't
    actually find."""
    schema = _compact_schema_summary(model, max_tables=25, max_cols_per_table=10)
    lines = []
    for c in cards:
        bits = ", ".join(f"{s}: {c['counts'][s]}" for s in ("High", "Medium", "Low") if c["counts"].get(s))
        lines.append(f"- {c['page']} ({c['title']}) — worst: {c['worst']}" + (f" [{bits}]" if bits else " — clean"))
    findings = "\n".join(lines) if lines else "(no sections evaluated)"

    return f"""You are a senior Power BI / DAX consultant giving a busy stakeholder a fast, honest read on
a semantic model's health, based only on the automated checks below - not a full manual review.

Overall score: {score}/100 ({grade})

Findings by section (severity = worst finding in that section, counts in brackets):
{findings}

Model schema (abbreviated, for context only):
{schema}

Write a short, prioritized summary in markdown with exactly these three headings:
### Bottom line
One or two plain-language sentences. No hedging, no filler.

### Fix first
The highest-impact items to address (at most 5), each one line, naming the exact section to open
(e.g. "Audit -> Model Health") and why it matters for this specific model.

### Can wait
Anything flagged but genuinely low-risk, so the reader knows what to skip for now.

Be concrete and specific to what's actually listed above. Never invent a finding, section, table, or
column that isn't listed above."""


def _http_post_json(url: str, headers: Dict[str, str], payload: Dict[str, Any], timeout: int = 60) -> Dict[str, Any]:
    """Minimal JSON-over-HTTPS POST via the standard library - deliberately
    not an extra SDK dependency for a bring-your-own-key, opt-in feature."""
    import urllib.error
    import urllib.request

    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        url, data=data, headers={**headers, "Content-Type": "application/json"}, method="POST"
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="ignore")
        try:
            detail = json.loads(body).get("error", {}).get("message", body)
        except (json.JSONDecodeError, AttributeError):
            detail = body
        if exc.code == 401:
            raise LLMError(
                "That API key was rejected (401 Unauthorized). Double-check you copied it correctly."
            ) from exc
        if exc.code == 429:
            raise LLMError(
                "Rate limited (429) - wait a moment and try again, or check your plan's usage limits."
            ) from exc
        raise LLMError(f"API error {exc.code}: {detail}") from exc
    except urllib.error.URLError as exc:
        raise LLMError(f"Network error reaching the API: {exc.reason}") from exc
    except TimeoutError as exc:
        raise LLMError("The request timed out - the model may be under heavy load; try again.") from exc


def call_openai_compatible(base_url: str, api_key: str, model_id: str, prompt: str, provider_label: str) -> str:
    """Shared dispatch for any provider that speaks the OpenAI Chat
    Completions wire format - OpenAI itself, plus DeepSeek and Groq (Llama),
    which both implement the same API shape."""
    result = _http_post_json(
        base_url,
        {"Authorization": f"Bearer {api_key}"},
        {
            "model": model_id,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.2,
            "response_format": {"type": "json_object"},
        },
    )
    try:
        return result["choices"][0]["message"]["content"]
    except (KeyError, IndexError) as exc:
        raise LLMError(f"Unexpected response shape from {provider_label}: {result}") from exc


def call_openai(api_key: str, model_id: str, prompt: str) -> str:
    return call_openai_compatible("https://api.openai.com/v1/chat/completions", api_key, model_id, prompt, "OpenAI")


def call_deepseek(api_key: str, model_id: str, prompt: str) -> str:
    return call_openai_compatible("https://api.deepseek.com/chat/completions", api_key, model_id, prompt, "DeepSeek")


def call_llama_groq(api_key: str, model_id: str, prompt: str) -> str:
    return call_openai_compatible(
        "https://api.groq.com/openai/v1/chat/completions", api_key, model_id, prompt, "Groq"
    )


def call_anthropic(api_key: str, model_id: str, prompt: str) -> str:
    result = _http_post_json(
        "https://api.anthropic.com/v1/messages",
        {"x-api-key": api_key, "anthropic-version": "2023-06-01"},
        {"model": model_id, "max_tokens": 2000, "messages": [{"role": "user", "content": prompt}]},
    )
    try:
        return "".join(block.get("text", "") for block in result["content"] if block.get("type") == "text")
    except (KeyError, TypeError) as exc:
        raise LLMError(f"Unexpected response shape from Anthropic: {result}") from exc


def call_gemini(api_key: str, model_id: str, prompt: str) -> str:
    import urllib.parse

    url = (
        f"https://generativelanguage.googleapis.com/v1beta/models/"
        f"{urllib.parse.quote(model_id)}:generateContent?key={urllib.parse.quote(api_key)}"
    )
    result = _http_post_json(
        url, {},
        {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": 0.2, "responseMimeType": "application/json"},
        },
    )
    try:
        return result["candidates"][0]["content"]["parts"][0]["text"]
    except (KeyError, IndexError) as exc:
        raise LLMError(f"Unexpected response shape from Gemini: {result}") from exc


def _llm_ready() -> bool:
    return bool(st.session_state.get("llm_api_key", "").strip())


_LLM_DISPATCH: Dict[str, Any] = {
    "OpenAI (ChatGPT)": call_openai,
    "Anthropic (Claude)": call_anthropic,
    "DeepSeek": call_deepseek,
    "Llama (via Groq)": call_llama_groq,
    "Google (Gemini)": call_gemini,
}


def call_llm(prompt: str) -> str:
    """Dispatch to whichever provider/key/model is configured in the
    sidebar's AI assistant section - the one place every AI feature reads
    its credentials from, so a key entered once works everywhere."""
    provider = st.session_state.get("llm_provider") or next(iter(LLM_PROVIDERS))
    api_key = st.session_state.get("llm_api_key", "").strip()
    if not api_key:
        raise LLMError("No API key configured — set one in the sidebar's AI assistant section.")
    model_id = st.session_state.get("llm_model_id_active", "").strip() or LLM_PROVIDERS[provider]["models"][0]
    fn = _LLM_DISPATCH.get(provider, call_openai)
    return fn(api_key, model_id, prompt)


def _parse_llm_json(raw: str) -> Dict[str, Any]:
    """Parse an LLM reply that's supposed to be one JSON object.

    Includes a defensive fallback for when the model wraps its JSON in a
    markdown code fence or adds stray commentary around it - common enough
    behaviour, even when explicitly asked for JSON only, that silently
    failing on it would make every JSON-shaped AI feature unreliable.
    """
    text = raw.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.I).strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", text, flags=re.S)
        if not match:
            raise LLMError(f"The model didn't return parseable JSON. Raw reply:\n\n{raw[:2000]}")
        try:
            return json.loads(match.group(0))
        except json.JSONDecodeError as exc:
            raise LLMError(f"The model's JSON couldn't be parsed. Raw reply:\n\n{raw[:2000]}") from exc


def regenerate_dax_with_llm(
    object_kind: str, table: str, name: str, expression: str, model: Dict[str, Any],
) -> Dict[str, str]:
    """Call the configured LLM and parse its JSON reply."""
    prompt = build_dax_regeneration_prompt(object_kind, table, name, expression, model)
    parsed = _parse_llm_json(call_llm(prompt))
    return {
        "revised_dax": str(parsed.get("revised_dax") or "").strip(),
        "notes": str(parsed.get("notes") or "").strip(),
        "model_suggestions": str(parsed.get("model_suggestions") or "").strip(),
    }


def _regenerate_one(object_kind: str, table: str, name: str, expr: str, model: Dict[str, Any]) -> Dict[str, Any]:
    """regenerate_dax_with_llm, but never raises - errors are folded into the
    result shape so a batch run can keep going past one bad call."""
    try:
        result = regenerate_dax_with_llm(object_kind, table, name, expr, model)
        return {**result, "error": None}
    except LLMError as exc:
        return {"revised_dax": "", "notes": "", "model_suggestions": "", "error": str(exc)}
    except Exception as exc:  # noqa: BLE001
        return {"revised_dax": "", "notes": "", "model_suggestions": "", "error": f"Unexpected error: {exc}"}


def build_grouped_fix_prompt(
    group: str, findings: pd.DataFrame, model: Dict[str, Any], object_col: str, message_col: str,
) -> str:
    """One group's worth of findings at a time, not one call per finding -
    findings sharing a group (a rule, a check, a table...) usually share a
    root cause, so batching by group keeps the call count proportional to
    distinct problem types instead of the total finding count (which can be
    hundreds on a real model)."""
    schema = _compact_schema_summary(model, max_tables=20, max_cols_per_table=8)
    capped = findings.head(30)
    rows = "\n".join(
        f"- Object: {r[object_col]}"
        + (f" | Severity: {r['Severity']}" if "Severity" in findings.columns else "")
        + f" | Detail: {r[message_col]}"
        for _, r in capped.iterrows()
    )
    extra = f"\n… +{len(findings) - len(capped)} more object(s) in this same group (same fix pattern likely applies)" \
        if len(findings) > len(capped) else ""

    return f"""You are a senior Power BI / DAX consultant writing specific, actionable remediation guidance
for a batch of automated findings that all share one group.

Group: {group}

Findings in this group:
{rows}{extra}

Model schema (abbreviated, for context only):
{schema}

For EACH object listed above, write a specific fix tailored to that exact object - reference its
real name, not generic advice like "review this column". If several objects share an identical
mechanical fix, it's fine for the wording to be similar, but still address every object listed by
name.

Respond with ONLY a single JSON object (no markdown fences, no commentary outside the JSON):
{{
  "fixes": {{"<exact Object value from above>": "specific fix for this object", ...}},
  "summary": "one or two sentences on the overall pattern and priority for this group"
}}
Never invent an object that isn't listed above."""


def get_ai_grouped_fixes(
    group: str, findings: pd.DataFrame, model: Dict[str, Any], object_col: str, message_col: str,
) -> Dict[str, Any]:
    prompt = build_grouped_fix_prompt(group, findings, model, object_col, message_col)
    parsed = _parse_llm_json(call_llm(prompt))
    fixes = parsed.get("fixes")
    return {
        "fixes": fixes if isinstance(fixes, dict) else {},
        "summary": str(parsed.get("summary") or "").strip(),
    }


def render_ai_grouped_fixes(
    findings_df: pd.DataFrame, model: Dict[str, Any], key_prefix: str,
    group_col: str, object_col: str, message_col: str, group_label: str,
    out_col: str = "AI How to Fix",
) -> pd.DataFrame:
    """AI-enhanced fix guidance for any findings table with a groupable
    category column (Model Health's Rule, Fabric Readiness's Check,
    Compression Advisor's Table...) - one LLM call per distinct group, not
    per finding, so it stays cheap even when there are hundreds of findings.
    Returns findings_df with `out_col` merged in once a batch has run, for
    the caller to hand to show_table.
    """
    batch_key = f"{key_prefix}_batch_results"
    with st.expander(f"AI-enhanced fix guidance — specific advice per finding, grouped by {group_label}",
                      icon=":material/smart_toy:", expanded=False):
        st.caption(
            "The fix guidance above is the same general text for every finding in a group. This "
            "sends each group's actual findings (the real object names and details) to the "
            f"provider configured in the sidebar and asks for advice specific to those exact "
            f"objects — one call per {group_label}, not per finding."
        )
        if not _llm_ready():
            st.info(
                "Set an OpenAI or Claude API key in the sidebar's **AI assistant** section to "
                "use this.", icon=":material/key:",
            )
            return findings_df
        if group_col not in findings_df.columns or findings_df.empty:
            st.info("No findings to explain.")
            return findings_df

        groups = sorted(findings_df[group_col].dropna().unique().tolist())
        cache: Dict[str, Any] = st.session_state.get(batch_key, {})
        done, total = len(cache), len(groups)
        provider = st.session_state.get("llm_provider") or next(iter(LLM_PROVIDERS))

        m1, m2 = st.columns(2)
        m1.metric(f"{group_label.capitalize()}s processed", f"{done} / {total}")
        m2.metric("Errors", sum(1 for r in cache.values() if r.get("error")))

        confirm = st.checkbox(
            f"I understand this makes {total - done} API call(s) to {provider}",
            key=f"{key_prefix}_confirm", value=False,
        ) if done < total else True

        c1, c2 = st.columns([3, 1])
        go_label = f"✨ Get AI fixes for all {total} {group_label}s" if done == 0 else f"✨ Get fixes for remaining {total - done}"
        go = c1.button(go_label, key=f"{key_prefix}_go", width="stretch", disabled=(done >= total) or not confirm)
        if cache and c2.button("Clear", key=f"{key_prefix}_clear", width="stretch"):
            st.session_state[batch_key] = {}
            st.rerun()

        if go:
            remaining = [g for g in groups if g not in cache]
            progress = st.progress(0.0)
            status = st.empty()
            for i, g in enumerate(remaining):
                status.caption(f"Asking {provider} about “{g}”… ({i + 1}/{len(remaining)})")
                group_findings = findings_df[findings_df[group_col] == g]
                try:
                    cache[g] = {**get_ai_grouped_fixes(str(g), group_findings, model, object_col, message_col), "error": None}
                except LLMError as exc:
                    cache[g] = {"fixes": {}, "summary": "", "error": str(exc)}
                except Exception as exc:  # noqa: BLE001
                    cache[g] = {"fixes": {}, "summary": "", "error": f"Unexpected error: {exc}"}
                st.session_state[batch_key] = dict(cache)
                progress.progress((i + 1) / max(len(remaining), 1))
            status.empty()
            progress.empty()
            st.rerun()

        for g, result in cache.items():
            if result.get("error"):
                st.error(f"**{g}**: {result['error']}")
            elif result.get("summary"):
                st.markdown(f"**{g}**")
                st.write(result["summary"])

    if not cache:
        return findings_df

    def _fix_for(row: pd.Series) -> str:
        r = cache.get(row[group_col])
        if not r or r.get("error"):
            return ""
        return r.get("fixes", {}).get(str(row[object_col]), "")

    out = findings_df.copy()
    out[out_col] = out.apply(_fix_for, axis=1)
    return out


def build_duplicate_judgment_prompt(row: pd.Series, model: Dict[str, Any]) -> str:
    schema = _compact_schema_summary(model, max_tables=15, max_cols_per_table=6)
    return f"""You are a senior Power BI / DAX consultant helping consolidate a pair of near-duplicate measures.

Measure A: {row['Measure A']} (table: {row['Table A']})
{row['DAX A']}

Measure B: {row['Measure B']} (table: {row['Table B']})
{row['DAX B']}

Similarity: {row['Similarity %']}%

Model schema (abbreviated, for context only):
{schema}

Decide which measure should be KEPT as the single source of truth, and which should be retired and
repointed to it. Consider which table is the more natural home, which name is clearer, whether one
is a strict subset/superset of the other's logic, and whether the DAX difference looks like a
genuine intentional business rule or accidental copy-paste drift that should be fixed, not kept.

Respond with ONLY a single JSON object (no markdown fences, no commentary outside the JSON):
{{
  "keep": "A" or "B",
  "reasoning": "one or two sentences, referencing the actual DAX difference if there is one",
  "is_intentional_difference": true or false
}}"""


def get_duplicate_judgment(row: pd.Series, model: Dict[str, Any]) -> Dict[str, Any]:
    prompt = build_duplicate_judgment_prompt(row, model)
    parsed = _parse_llm_json(call_llm(prompt))
    keep = str(parsed.get("keep") or "").strip().upper()
    return {
        "keep": keep if keep in ("A", "B") else "",
        "reasoning": str(parsed.get("reasoning") or "").strip(),
        "is_intentional": bool(parsed.get("is_intentional_difference")),
    }


def render_ai_duplicate_judgments(dupes_df: pd.DataFrame, model: Dict[str, Any], key_prefix: str) -> pd.DataFrame:
    """One LLM call per near-duplicate pair - unlike the grouped-fix helper,
    every pair is genuinely distinct (different DAX, different context), so
    there's no shared root cause to batch by."""
    batch_key = f"{key_prefix}_batch_results"
    with st.expander("AI consolidation advice — which measure to keep, per pair",
                      icon=":material/smart_toy:", expanded=False):
        st.caption(
            "For each pair, sends both DAX expressions to the provider configured in the "
            "sidebar and asks which one to keep as the single source of truth, and whether the "
            "difference looks like a genuine business rule or accidental drift."
        )
        if not _llm_ready():
            st.info(
                "Set an OpenAI or Claude API key in the sidebar's **AI assistant** section to "
                "use this.", icon=":material/key:",
            )
            return dupes_df
        if dupes_df.empty:
            return dupes_df

        pairs = dupes_df.apply(lambda r: f"{r['Measure A']} ||| {r['Measure B']}", axis=1).tolist()
        cache: Dict[str, Any] = st.session_state.get(batch_key, {})
        done, total = len(cache), len(pairs)
        provider = st.session_state.get("llm_provider") or next(iter(LLM_PROVIDERS))

        m1, m2 = st.columns(2)
        m1.metric("Pairs judged", f"{done} / {total}")
        m2.metric("Errors", sum(1 for r in cache.values() if r.get("error")))

        confirm = st.checkbox(
            f"I understand this makes {total - done} API call(s) to {provider}",
            key=f"{key_prefix}_confirm", value=False,
        ) if done < total else True
        c1, c2 = st.columns([3, 1])
        go_label = f"✨ Judge all {total} pairs" if done == 0 else f"✨ Judge remaining {total - done}"
        go = c1.button(go_label, key=f"{key_prefix}_go", width="stretch", disabled=(done >= total) or not confirm)
        if cache and c2.button("Clear", key=f"{key_prefix}_clear", width="stretch"):
            st.session_state[batch_key] = {}
            st.rerun()

        if go:
            remaining_idx = [i for i, k in enumerate(pairs) if k not in cache]
            progress = st.progress(0.0)
            status = st.empty()
            for n, i in enumerate(remaining_idx):
                row = dupes_df.iloc[i]
                k = pairs[i]
                status.caption(
                    f"Asking {provider} about “{row['Measure A']}” vs “{row['Measure B']}”… "
                    f"({n + 1}/{len(remaining_idx)})"
                )
                try:
                    cache[k] = {**get_duplicate_judgment(row, model), "error": None}
                except LLMError as exc:
                    cache[k] = {"keep": "", "reasoning": "", "is_intentional": False, "error": str(exc)}
                except Exception as exc:  # noqa: BLE001
                    cache[k] = {"keep": "", "reasoning": "", "is_intentional": False, "error": f"Unexpected error: {exc}"}
                st.session_state[batch_key] = dict(cache)
                progress.progress((n + 1) / max(len(remaining_idx), 1))
            status.empty()
            progress.empty()
            st.rerun()

    if not cache:
        return dupes_df

    out = dupes_df.copy()

    def _recommend(row: pd.Series) -> str:
        k = f"{row['Measure A']} ||| {row['Measure B']}"
        r = cache.get(k)
        if not r or r.get("error"):
            return ""
        if r["keep"] == "A":
            return f"Keep {row['Measure A']} ({row['Table A']})"
        if r["keep"] == "B":
            return f"Keep {row['Measure B']} ({row['Table B']})"
        return ""

    def _reasoning(row: pd.Series) -> str:
        k = f"{row['Measure A']} ||| {row['Measure B']}"
        r = cache.get(k)
        if not r or r.get("error"):
            return ""
        tag = "" if r.get("is_intentional") else " (looks accidental)"
        return f"{r['reasoning']}{tag}"

    out["AI Recommendation"] = out.apply(_recommend, axis=1)
    out["AI Reasoning"] = out.apply(_reasoning, axis=1)
    return out


def build_unused_risk_prompt(table: str, cols: pd.DataFrame, model: Dict[str, Any]) -> str:
    schema = _compact_schema_summary(model, max_tables=20, max_cols_per_table=8)
    capped = cols.head(30)
    names = "\n".join(f"- {r['Column']}" for _, r in capped.iterrows())
    extra = f"\n… +{len(cols) - len(capped)} more column(s) on this table" if len(cols) > len(capped) else ""

    return f"""You are a senior Power BI consultant reviewing columns a static DAX-reference scan found no
measure, calculated column, or relationship pointing at — candidates for deletion, but the scan
cannot see Power BI report visuals or row-level-security filter expressions, so some of these may
still be in active use.

Table: {table}

Columns flagged as likely unused on this table:
{names}{extra}

Model schema (abbreviated, for context only):
{schema}

For EACH column listed above, give a confidence judgment based only on its name and the schema
context — not invented facts. Use "Likely safe" for columns whose name suggests nothing else would
plausibly need them (internal keys, deprecated-looking names, clear duplicates of another column).
Use "Verify first" for columns whose name suggests a real risk a DAX-only scan can't see — anything
that looks like it could be a report-visual axis/slicer field, an RLS/security-related column, a
display or tooltip label, or a natural key someone might filter on directly.

Respond with ONLY a single JSON object (no markdown fences, no commentary outside the JSON):
{{
  "assessments": {{"<exact column name from above>": {{"confidence": "Likely safe" or "Verify first", "reason": "one short sentence"}}, ...}}
}}
Never invent a column that isn't listed above."""


def get_ai_unused_risk(table: str, cols: pd.DataFrame, model: Dict[str, Any]) -> Dict[str, Any]:
    prompt = build_unused_risk_prompt(table, cols, model)
    parsed = _parse_llm_json(call_llm(prompt))
    assessments = parsed.get("assessments")
    return {"assessments": assessments if isinstance(assessments, dict) else {}}


def render_ai_unused_risk(unused_view: pd.DataFrame, model: Dict[str, Any], key_prefix: str) -> pd.DataFrame:
    """AI confidence judgment on the 'likely unused' list, grouped by table
    (one call per table, not per column - this model's real list ran to 400+
    columns, so per-column would be prohibitively expensive)."""
    batch_key = f"{key_prefix}_batch_results"
    with st.expander("AI risk assessment — which of these are actually safe to delete",
                      icon=":material/smart_toy:", expanded=False):
        st.caption(
            "This scan is DAX-only — it can't see report visuals or RLS filters, so 'likely "
            "unused' isn't a guarantee. This sends each table's flagged columns to the provider "
            "configured in the sidebar and asks it to flag which look genuinely safe to drop "
            "versus which are worth double-checking first, based on naming patterns alone."
        )
        if not _llm_ready():
            st.info(
                "Set an OpenAI or Claude API key in the sidebar's **AI assistant** section to "
                "use this.", icon=":material/key:",
            )
            return unused_view
        if unused_view.empty:
            return unused_view

        tables = sorted(unused_view["Table"].dropna().unique().tolist())
        cache: Dict[str, Any] = st.session_state.get(batch_key, {})
        done, total = len(cache), len(tables)
        provider = st.session_state.get("llm_provider") or next(iter(LLM_PROVIDERS))

        m1, m2 = st.columns(2)
        m1.metric("Tables processed", f"{done} / {total}")
        m2.metric("Errors", sum(1 for r in cache.values() if r.get("error")))

        confirm = st.checkbox(
            f"I understand this makes {total - done} API call(s) to {provider}",
            key=f"{key_prefix}_confirm", value=False,
        ) if done < total else True
        c1, c2 = st.columns([3, 1])
        go_label = f"✨ Assess all {total} tables" if done == 0 else f"✨ Assess remaining {total - done}"
        go = c1.button(go_label, key=f"{key_prefix}_go", width="stretch", disabled=(done >= total) or not confirm)
        if cache and c2.button("Clear", key=f"{key_prefix}_clear", width="stretch"):
            st.session_state[batch_key] = {}
            st.rerun()

        if go:
            remaining = [t for t in tables if t not in cache]
            progress = st.progress(0.0)
            status = st.empty()
            for i, table in enumerate(remaining):
                status.caption(f"Asking {provider} about “{table}”… ({i + 1}/{len(remaining)})")
                table_cols = unused_view[unused_view["Table"] == table]
                try:
                    cache[table] = {**get_ai_unused_risk(table, table_cols, model), "error": None}
                except LLMError as exc:
                    cache[table] = {"assessments": {}, "error": str(exc)}
                except Exception as exc:  # noqa: BLE001
                    cache[table] = {"assessments": {}, "error": f"Unexpected error: {exc}"}
                st.session_state[batch_key] = dict(cache)
                progress.progress((i + 1) / max(len(remaining), 1))
            status.empty()
            progress.empty()
            st.rerun()

    if not cache:
        return unused_view

    def _confidence(row: pd.Series) -> str:
        r = cache.get(row["Table"])
        if not r or r.get("error"):
            return ""
        a = r.get("assessments", {}).get(str(row["Column"]))
        return a.get("confidence", "") if isinstance(a, dict) else ""

    def _reason(row: pd.Series) -> str:
        r = cache.get(row["Table"])
        if not r or r.get("error"):
            return ""
        a = r.get("assessments", {}).get(str(row["Column"]))
        return a.get("reason", "") if isinstance(a, dict) else ""

    out = unused_view.copy()
    out["AI Confidence"] = out.apply(_confidence, axis=1)
    out["AI Reason"] = out.apply(_reason, axis=1)
    return out


def build_description_draft_prompt(
    table: str, measures: pd.DataFrame, cols: pd.DataFrame, model: Dict[str, Any],
) -> str:
    schema = _compact_schema_summary(model, max_tables=15, max_cols_per_table=6)
    meas_lines = "\n".join(
        f"- Measure: {r['MeasureName']} = {r['MeasureExpression']}" for _, r in measures.head(20).iterrows()
    )
    col_lines = "\n".join(
        f"- Column: {r['ColumnName']} ({r.get('DataType', '')})" for _, r in cols.head(20).iterrows()
    )
    return f"""You are a senior BI analyst writing a data dictionary. Draft short, plain-English business
descriptions for the undocumented objects below on table "{table}" - the kind a new analyst or a
business stakeholder (not a DAX expert) would find useful. Base descriptions only on the object's
name, its DAX (for measures), and the schema context - never invent business meaning that isn't
evidenced by the name or logic.

Measures with no description:
{meas_lines or '(none)'}

Columns with no description:
{col_lines or '(none)'}

Model schema (abbreviated, for context only):
{schema}

Respond with ONLY a single JSON object (no markdown fences, no commentary outside the JSON):
{{
  "measures": {{"<exact measure name from above>": "one-sentence description", ...}},
  "columns": {{"<exact column name from above>": "one-sentence description", ...}}
}}
Only include objects listed above."""


def get_ai_description_drafts(
    table: str, measures: pd.DataFrame, cols: pd.DataFrame, model: Dict[str, Any],
) -> Dict[str, Any]:
    prompt = build_description_draft_prompt(table, measures, cols, model)
    parsed = _parse_llm_json(call_llm(prompt))
    return {
        "measures": parsed.get("measures") if isinstance(parsed.get("measures"), dict) else {},
        "columns": parsed.get("columns") if isinstance(parsed.get("columns"), dict) else {},
    }


def render_ai_descriptions(model: Dict[str, Any], key_prefix: str) -> pd.DataFrame:
    """AI-drafted plain-English descriptions for measures/columns that have
    none, grouped by table (one call per table with undocumented objects).
    Returns a standalone {Table, Object Type, Object Name, AI-Drafted
    Description} frame - a review sheet, not a write into the live model, so
    nothing else in the app is affected by what the AI drafts here."""
    meas_df = model["measures"]
    cols_df = _user_facing_columns(model["columns"])
    undoc_meas = meas_df[meas_df.get("Description", "").astype(str).str.strip() == ""] if "Description" in meas_df.columns else meas_df.iloc[0:0]
    undoc_cols = cols_df[cols_df.get("Description", "").astype(str).str.strip() == ""] if "Description" in cols_df.columns else cols_df.iloc[0:0]

    with st.expander("AI-drafted descriptions — for measures/columns with none",
                      icon=":material/smart_toy:", expanded=False):
        n_meas, n_cols = len(undoc_meas), len(undoc_cols)
        st.caption(
            f"{n_meas} measure(s) and {n_cols} column(s) in this model have no description. This "
            "drafts plain-English descriptions from each object's name and DAX (for measures) — "
            "review before publishing, since it can't know business context beyond what's "
            "evidenced in the model itself. One call per table with undocumented objects."
        )
        if not _llm_ready():
            st.info(
                "Set an OpenAI or Claude API key in the sidebar's **AI assistant** section to "
                "use this.", icon=":material/key:",
            )
            return pd.DataFrame(columns=["Table", "Object Type", "Object Name", "AI-Drafted Description"])
        if n_meas == 0 and n_cols == 0:
            st.success("Every measure and column already has a description.", icon="✅")
            return pd.DataFrame(columns=["Table", "Object Type", "Object Name", "AI-Drafted Description"])

        tables = sorted(set(undoc_meas["TableName"].dropna().tolist()) | set(undoc_cols["TableName"].dropna().tolist()))
        batch_key = f"{key_prefix}_batch_results"
        cache: Dict[str, Any] = st.session_state.get(batch_key, {})
        done, total = len(cache), len(tables)
        provider = st.session_state.get("llm_provider") or next(iter(LLM_PROVIDERS))

        m1, m2 = st.columns(2)
        m1.metric("Tables drafted", f"{done} / {total}")
        m2.metric("Errors", sum(1 for r in cache.values() if r.get("error")))

        confirm = st.checkbox(
            f"I understand this makes {total - done} API call(s) to {provider}",
            key=f"{key_prefix}_confirm", value=False,
        ) if done < total else True
        c1, c2 = st.columns([3, 1])
        go_label = f"✨ Draft descriptions for all {total} tables" if done == 0 else f"✨ Draft remaining {total - done}"
        go = c1.button(go_label, key=f"{key_prefix}_go", width="stretch", disabled=(done >= total) or not confirm)
        if cache and c2.button("Clear", key=f"{key_prefix}_clear", width="stretch"):
            st.session_state[batch_key] = {}
            st.rerun()

        if go:
            remaining = [t for t in tables if t not in cache]
            progress = st.progress(0.0)
            status = st.empty()
            for i, table in enumerate(remaining):
                status.caption(f"Asking {provider} about “{table}”… ({i + 1}/{len(remaining)})")
                t_meas = undoc_meas[undoc_meas["TableName"] == table]
                t_cols = undoc_cols[undoc_cols["TableName"] == table]
                try:
                    cache[table] = {**get_ai_description_drafts(table, t_meas, t_cols, model), "error": None}
                except LLMError as exc:
                    cache[table] = {"measures": {}, "columns": {}, "error": str(exc)}
                except Exception as exc:  # noqa: BLE001
                    cache[table] = {"measures": {}, "columns": {}, "error": f"Unexpected error: {exc}"}
                st.session_state[batch_key] = dict(cache)
                progress.progress((i + 1) / max(len(remaining), 1))
            status.empty()
            progress.empty()
            st.rerun()

        if cache:
            rows = []
            for table, result in cache.items():
                if result.get("error"):
                    continue
                for name, desc in result.get("measures", {}).items():
                    rows.append({"Table": table, "Object Type": "Measure", "Object Name": name,
                                 "AI-Drafted Description": desc})
                for name, desc in result.get("columns", {}).items():
                    rows.append({"Table": table, "Object Type": "Column", "Object Name": name,
                                 "AI-Drafted Description": desc})
            out = pd.DataFrame(rows, columns=["Table", "Object Type", "Object Name", "AI-Drafted Description"])
            if not out.empty:
                show_table(out, "AI-Drafted Descriptions", height=320, key=f"{key_prefix}_preview", row_height=90)
            return out

    return pd.DataFrame(columns=["Table", "Object Type", "Object Name", "AI-Drafted Description"])


def _fmt_ai_context_rows(df: pd.DataFrame, cols: List[str], cap: int = 20) -> str:
    """A bounded, plain-text row dump for an LLM prompt - shared by the
    single-call 'narrative' AI features (changelog, likely-cause) that need
    to show a table's content without the batch-by-group machinery."""
    if df.empty:
        return "(none)"
    present = [c for c in cols if c in df.columns]
    lines = [" | ".join(f"{c}: {r[c]}" for c in present) for _, r in df.head(cap).iterrows()]
    extra = f"\n… +{len(df) - cap} more" if len(df) > cap else ""
    return "- " + "\n- ".join(lines) + extra


def build_compare_changelog_prompt(
    drift: pd.DataFrame, added: pd.DataFrame, removed: pd.DataFrame, structure: pd.DataFrame,
    b_name: str, model: Dict[str, Any],
) -> str:
    return f"""You are a senior Power BI consultant writing a plain-English changelog for a stakeholder who
needs to know what changed between a certified baseline model and a newer version, and whether they
should be worried.

Baseline: {b_name}

Drifted measures (same name, different DAX):
{_fmt_ai_context_rows(drift, ["Measure", "Change", "Severity", "Baseline Table", "Compared Table"])}

Measures missing from the newer model:
{_fmt_ai_context_rows(removed, ["Measure", "Table", "Severity"])}

Extra measures only in the newer model:
{_fmt_ai_context_rows(added, ["Measure", "Table", "Severity"])}

Structural differences:
{_fmt_ai_context_rows(structure, list(structure.columns)[:5])}

Write a short markdown summary with exactly these three headings:
### What changed
Plain-language summary, grouped by theme if there's a pattern — not a re-listing of every row.

### Risk
Which changes are most likely to break downstream reports or produce wrong numbers, and why.

### Recommended next step
One or two concrete actions.

Be specific to what's actually listed above. Never invent a measure or change that isn't listed."""


def render_ai_compare_changelog(
    drift: pd.DataFrame, added: pd.DataFrame, removed: pd.DataFrame, structure: pd.DataFrame,
    b_name: str, model: Dict[str, Any], key_prefix: str,
) -> None:
    """A single LLM call summarizing the whole diff - unlike the audit tabs,
    there's one coherent story to tell here, not many independent findings
    to batch through."""
    with st.expander("AI changelog — what changed and why it matters",
                      icon=":material/smart_toy:", expanded=True):
        if not _llm_ready():
            st.info(
                "Set an OpenAI or Claude API key in the sidebar's **AI assistant** section to "
                "use this.", icon=":material/key:",
            )
            return
        provider = st.session_state.get("llm_provider") or next(iter(LLM_PROVIDERS))
        cache_key = f"{key_prefix}_result"
        if st.button("✨ Generate AI changelog", key=f"{key_prefix}_go", width="stretch"):
            with st.spinner(f"Asking {provider} to summarize the diff…"):
                try:
                    prompt = build_compare_changelog_prompt(drift, added, removed, structure, b_name, model)
                    st.session_state[cache_key] = call_llm(prompt)
                except LLMError as exc:
                    st.error(str(exc))
                except Exception as exc:  # noqa: BLE001
                    st.error(f"Unexpected error calling {provider}: {exc}")
        text = st.session_state.get(cache_key)
        if text:
            st.markdown(text)


def _cs_verbatim(value: str) -> str:
    """C# verbatim string literal (@"...") - handles multi-line DAX cleanly;
    the only character that needs escaping is an embedded double-quote,
    doubled rather than backslash-escaped."""
    return '@"' + str(value).replace('"', '""') + '"'


def build_dax_apply_csharp(items: List[Dict[str, str]]) -> str:
    """items: [{"object_kind": "measure"|"calculated column", "table", "name", "dax"}, ...]"""
    lines: List[str] = [
        "// =====================================================================",
        "// Generated by VPAX Semantic Model Explorer - AI DAX Assistant",
        "// Applies ONLY the AI-suggested rewrites accepted below.",
        "//",
        "// HOW TO RUN",
        "//   1. Open the model in Tabular Editor (2 or 3).",
        "//   2. Advanced Scripting (C#) tab - paste this script, click Run (F5).",
        "//   3. Review the diff Tabular Editor shows before saving.",
        "//   4. Save changes back to the model (Ctrl+S, or File > Save to XMLA/PBIP).",
        "//",
        "// This DAX was AI-generated and has not been validated against your live",
        "// data - review every expression below before running.",
        "// =====================================================================",
        "",
    ]
    for item in items:
        collection = "Measures" if item["object_kind"] == "measure" else "Columns"
        lines.append(
            f'Model.Tables[{_cs_string(item["table"])}].{collection}[{_cs_string(item["name"])}]'
            f'.Expression = {_cs_verbatim(item["dax"])};'
        )
    lines.append("")
    lines.append(f"// {len(items)} object(s) updated.")
    return "\n".join(lines)


def render_apply_dax_script(
    df: pd.DataFrame, name_col: str, object_kind: str, ai_key_prefix: str, apply_key_prefix: str,
) -> None:
    """Lets the user pick which AI-suggested DAX rewrites to accept, then
    download a Tabular Editor C# script that applies just those.

    This app has no safe way to write DAX directly into a .pbix - the
    compiled semantic model inside is a proprietary Analysis Services binary,
    not JSON or text, and only the real Tabular Object Model engine (what
    Tabular Editor runs on) can edit it correctly. Patching the bytes from
    here risks silently producing a file Power BI Desktop can't open. This
    is the safe equivalent: same one-click feel, but the actual write
    happens inside a tool built to do it.
    """
    results: Dict[str, Any] = st.session_state.get(f"{ai_key_prefix}_batch_results", {})
    candidates = {
        name: r for name, r in results.items()
        if not r.get("error") and str(r.get("revised_dax") or "").strip()
    }
    if not candidates:
        return

    with st.expander("Apply accepted suggestions — download a Tabular Editor script",
                      icon=":material/rocket_launch:", expanded=False):
        st.caption(
            "This app can't safely write DAX directly into a .pbix — its compiled semantic "
            "model is a proprietary binary that only Tabular Editor's engine can edit correctly; "
            "patching it here risks corrupting the file. Pick which AI suggestions below you "
            "accept, and download a ready-to-run script that applies exactly those changes in "
            "Tabular Editor instead."
        )
        options = list(candidates.keys())
        picked = st.multiselect(
            f"{object_kind.capitalize()}s to apply", options, default=options,
            key=f"{apply_key_prefix}_pick",
        )
        if not picked:
            st.info("Select at least one accepted suggestion to generate a script.")
            return

        items = []
        for name in picked:
            row = df[df[name_col].astype(str) == name]
            if row.empty:
                continue
            table = str(row.iloc[0].get("TableName") or "")
            items.append({
                "object_kind": object_kind, "table": table, "name": name,
                "dax": candidates[name]["revised_dax"],
            })
        script = build_dax_apply_csharp(items)
        preview = script if len(script) <= 2000 else script[:2000] + "\n// … (truncated preview, full script downloads below)"
        st.code(preview, language="csharp")
        st.download_button(
            f"⬇ Download apply script ({len(items)} change{'s' if len(items) != 1 else ''})",
            script.encode("utf-8"),
            file_name=f"apply_ai_dax_{_slug(object_kind)}.csx", mime="text/plain",
            key=f"{apply_key_prefix}_dl", width="stretch",
        )


def build_source_description_prompt(table: str, sql: str, model: Dict[str, Any]) -> str:
    schema = _compact_schema_summary(model, max_tables=10, max_cols_per_table=6)
    return f"""You are a data analyst documenting where a semantic model's tables come from. Given the
SQL that loads table "{table}", write a short, plain-English description of what this data source
represents in business terms — what real-world entity or event it captures — based only on the SQL
(table/column names, joins, filters) and the schema context. Never invent business meaning that
isn't evidenced by the SQL itself.

SQL:
{sql}

Model schema (abbreviated, for context only):
{schema}

Respond with ONLY a single JSON object (no markdown fences, no commentary outside the JSON):
{{"description": "two or three sentences"}}"""


def get_ai_source_description(table: str, sql: str, model: Dict[str, Any]) -> str:
    parsed = _parse_llm_json(call_llm(build_source_description_prompt(table, sql, model)))
    return str(parsed.get("description") or "").strip()


def render_ai_source_descriptions(pq_df: pd.DataFrame, model: Dict[str, Any], key_prefix: str) -> pd.DataFrame:
    """One LLM call per table with an embedded SQL query, describing what
    that source represents in business terms - the kind of context a .vpax/
    .pbix carries nothing about, since it only knows the query, not why it
    exists."""
    batch_key = f"{key_prefix}_batch_results"
    sql_len = pq_df["SQL"].fillna("").astype(str).str.len() if "SQL" in pq_df.columns else pd.Series(dtype=int)
    with_sql = pq_df[sql_len > 0]

    with st.expander("AI-generated data source descriptions",
                      icon=":material/smart_toy:", expanded=False):
        st.caption(
            "For each table whose Power Query step embeds a SQL query, sends that SQL to the "
            "provider configured in the sidebar and asks for a plain-English description of what "
            "the source represents — useful for onboarding or a data catalog, since a .vpax/.pbix "
            "carries the query itself but nothing about why it exists."
        )
        if not _llm_ready():
            st.info(
                "Set an OpenAI or Claude API key in the sidebar's **AI assistant** section to "
                "use this.", icon=":material/key:",
            )
            return pq_df
        if with_sql.empty:
            st.info("No table in this model has an embedded SQL query to describe.")
            return pq_df

        tables = sorted(with_sql["TableName"].dropna().unique().tolist())
        cache: Dict[str, Any] = st.session_state.get(batch_key, {})
        done, total = len(cache), len(tables)
        provider = st.session_state.get("llm_provider") or next(iter(LLM_PROVIDERS))

        m1, m2 = st.columns(2)
        m1.metric("Tables described", f"{done} / {total}")
        m2.metric("Errors", sum(1 for r in cache.values() if r.get("error")))

        confirm = st.checkbox(
            f"I understand this makes {total - done} API call(s) to {provider}",
            key=f"{key_prefix}_confirm", value=False,
        ) if done < total else True
        c1, c2 = st.columns([3, 1])
        go_label = f"✨ Describe all {total} sources" if done == 0 else f"✨ Describe remaining {total - done}"
        go = c1.button(go_label, key=f"{key_prefix}_go", width="stretch", disabled=(done >= total) or not confirm)
        if cache and c2.button("Clear", key=f"{key_prefix}_clear", width="stretch"):
            st.session_state[batch_key] = {}
            st.rerun()

        if go:
            remaining = [t for t in tables if t not in cache]
            progress = st.progress(0.0)
            status = st.empty()
            for i, table in enumerate(remaining):
                status.caption(f"Asking {provider} about “{table}”… ({i + 1}/{len(remaining)})")
                sql = str(with_sql[with_sql["TableName"] == table]["SQL"].iloc[0])
                try:
                    cache[table] = {"description": get_ai_source_description(table, sql, model), "error": None}
                except LLMError as exc:
                    cache[table] = {"description": "", "error": str(exc)}
                except Exception as exc:  # noqa: BLE001
                    cache[table] = {"description": "", "error": f"Unexpected error: {exc}"}
                st.session_state[batch_key] = dict(cache)
                progress.progress((i + 1) / max(len(remaining), 1))
            status.empty()
            progress.empty()
            st.rerun()

    if not cache:
        return pq_df

    def _desc(table: str) -> str:
        r = cache.get(table)
        return "" if not r or r.get("error") else r.get("description", "")

    out = pq_df.copy()
    out["AI Source Description"] = out["TableName"].map(_desc)
    return out


def build_page_summary_prompt(page_name: str, report: Dict[str, Any], model: Dict[str, Any]) -> str:
    bindings = report["bindings"]
    page_bindings = bindings[bindings["Page"] == page_name] if not bindings.empty else bindings

    visuals: List[str] = []
    if not page_bindings.empty and "Visual" in page_bindings.columns:
        for visual_id, grp in page_bindings.groupby("Visual"):
            vtype = str(grp["Visual Type"].iloc[0]) if "Visual Type" in grp.columns and not grp.empty else ""
            fields = sorted({f"{r['Table']}[{r['Field']}]" for _, r in grp.iterrows() if r.get("Field")})
            if fields:
                visuals.append(f"- {vtype or 'visual'}: {', '.join(fields[:10])}")
    visuals_text = "\n".join(visuals[:25]) if visuals else "(no visual bindings read for this page)"

    tables = tables_used_by_page(page_name, report, model)
    page_fields = set(page_bindings["Field"].dropna().astype(str)) if not page_bindings.empty else set()
    measures_df = model["measures"]
    used_measures = measures_df[measures_df["MeasureName"].isin(page_fields)] if not measures_df.empty else measures_df
    meas_text = "\n".join(
        f"- {r['MeasureName']} = {r['MeasureExpression']}" for _, r in used_measures.head(20).iterrows()
    ) or "(no measures identified on this page)"

    pq_df = model["power_query"]
    sql_bits = []
    if not pq_df.empty and "SQL" in pq_df.columns:
        for t in sorted(tables)[:10]:
            rows = pq_df[(pq_df["TableName"] == t) & (pq_df["SQL"].fillna("").astype(str).str.len() > 0)]
            if not rows.empty:
                sql_bits.append(f"-- {t}\n{str(rows.iloc[0]['SQL'])[:400]}")
    sql_text = "\n\n".join(sql_bits) if sql_bits else "(no source SQL available for these tables)"

    return f"""You are a BI analyst writing a plain-English summary of a Power BI report page for a
business audience — what it shows and what business questions it helps answer. Base this only on
the visuals, measures, and source data listed below; never invent one that isn't listed.

Page: {page_name}

Visuals and the fields they show:
{visuals_text}

Measures used on this page:
{meas_text}

Tables involved: {", ".join(sorted(tables)) or "(none identified)"}

Source SQL for those tables (abbreviated):
{sql_text}

Write a summary in markdown with exactly these two headings:
### What this page shows
One paragraph, plain business language — not a re-listing of every visual.

### Key metrics
A short bullet list of the most important measures on this page (at most 5) and what each one
means in one sentence."""


def render_ai_page_summary(page_name: str, report: Dict[str, Any], model: Dict[str, Any], key_prefix: str) -> None:
    """A single LLM call summarizing one report page - the visuals, the DAX
    behind them, and the source SQL behind that, all of which a .pbix+.vpax
    pair together make available but nothing in the file itself narrates."""
    with st.expander(f"AI summary — what “{page_name}” shows",
                      icon=":material/smart_toy:", expanded=True):
        if not _llm_ready():
            st.info(
                "Set an OpenAI or Claude API key in the sidebar's **AI assistant** section to "
                "use this.", icon=":material/key:",
            )
            return
        provider = st.session_state.get("llm_provider") or next(iter(LLM_PROVIDERS))
        cache_key = f"{key_prefix}_{_slug(page_name)}"
        if st.button("✨ Generate page summary", key=f"{cache_key}_go", width="stretch"):
            with st.spinner(f"Asking {provider} to summarize “{page_name}”…"):
                try:
                    st.session_state[cache_key] = call_llm(build_page_summary_prompt(page_name, report, model))
                except LLMError as exc:
                    st.error(str(exc))
                except Exception as exc:  # noqa: BLE001
                    st.error(f"Unexpected error calling {provider}: {exc}")
        text = st.session_state.get(cache_key)
        if text:
            st.markdown(text)


def render_ai_dax_assistant(
    object_kind: str, df: pd.DataFrame, name_col: str, expr_col: str, model: Dict[str, Any], key_prefix: str
) -> pd.DataFrame:
    """AI DAX panel shared by the Measures and Calculated Columns pages -
    either regenerate one object with a full before/after view, or regenerate
    every object and get the result back as extra columns.

    Reads its provider/key/model from the sidebar's AI assistant section -
    set once there, used everywhere. Returns `df` with AI columns merged in
    when a batch run has produced results, so the caller can hand that
    straight to show_table (and therefore its CSV/Excel export) instead of
    the original frame.
    """
    batch_key = f"{key_prefix}_batch_results"
    with st.expander(f"AI DAX Assistant — regenerate {object_kind} DAX with your configured LLM",
                      icon=":material/smart_toy:", expanded=True):
        st.caption(
            f"Sends each {object_kind}'s DAX, plus a compact summary of this model's tables, "
            "columns and relationships, to the provider configured in the sidebar, and asks it "
            "to rewrite the DAX following best practice and suggest any data-modelling changes "
            "worth considering. **Nothing is applied automatically** — this app has no way to "
            "write back into a .vpax, so review the result and copy it into Tabular Editor or "
            "Power BI Desktop yourself."
        )
        if not _llm_ready():
            st.info(
                "Set an OpenAI or Claude API key in the sidebar's **AI assistant** section to "
                "use this.", icon=":material/key:",
            )
            return df

        options = df[name_col].dropna().astype(str).tolist()
        if not options:
            st.info(f"No {object_kind}s to regenerate.")
            return df

        provider = st.session_state.get("llm_provider") or next(iter(LLM_PROVIDERS))
        tab_all, tab_one = st.tabs([f"Regenerate all ({len(options)})", "Regenerate one"])

        with tab_all:
            st.caption(
                f"One API call per {object_kind} — **{len(options)} calls total**. This can take "
                "several minutes and uses real API quota on your key. Progress is saved as it "
                "goes, so it's safe to stop partway and resume later; already-done objects are "
                "skipped on the next run."
            )
            results: Dict[str, Dict[str, Any]] = st.session_state.get(batch_key, {})
            done, total = len(results), len(options)
            n_err = sum(1 for r in results.values() if r.get("error"))
            m1, m2 = st.columns(2)
            m1.metric("Processed", f"{done} / {total}")
            m2.metric("Errors", n_err)

            confirm = st.checkbox(
                f"I understand this makes {total - done} API call(s) to {provider}",
                key=f"{key_prefix}_batch_confirm", value=False,
            ) if done < total else True

            c1, c2 = st.columns([3, 1])
            go_label = f"✨ Regenerate all {total} with AI" if done == 0 else f"✨ Regenerate remaining {total - done}"
            go = c1.button(
                go_label, key=f"{key_prefix}_batch_go", width="stretch",
                disabled=(done >= total) or not confirm,
            )
            if results and c2.button("Clear", key=f"{key_prefix}_batch_clear", width="stretch"):
                st.session_state[batch_key] = {}
                st.rerun()

            if go:
                remaining = [o for o in options if o not in results]
                progress = st.progress(0.0)
                status = st.empty()
                for i, name in enumerate(remaining):
                    status.caption(f"Asking {provider} to review “{name}”… ({i + 1}/{len(remaining)})")
                    row = df[df[name_col].astype(str) == name].iloc[0]
                    table = str(row.get("TableName") or "")
                    expr = str(row.get(expr_col) or "")
                    results[name] = _regenerate_one(object_kind, table, name, expr, model)
                    st.session_state[batch_key] = dict(results)
                    progress.progress((i + 1) / max(len(remaining), 1))
                status.empty()
                progress.empty()
                st.rerun()

        with tab_one:
            picked = st.selectbox(f"Choose a {object_kind}", options, key=f"{key_prefix}_pick")
            row = df[df[name_col].astype(str) == picked].iloc[0]
            table = str(row.get("TableName") or "")
            expr = str(row.get(expr_col) or "")
            # Included in the cache key so a picked object whose DAX changed
            # (e.g. after re-uploading a different .vpax) doesn't show a
            # stale suggestion generated against the old expression.
            cache_key = f"{key_prefix}_result_{picked}_{hash(expr) & 0xffffffff}"

            if st.button("✨ Regenerate with AI", key=f"{key_prefix}_go", width="stretch"):
                with st.spinner(f"Asking {provider} to review {picked}…"):
                    result = _regenerate_one(object_kind, table, picked, expr, model)
                    if result["error"]:
                        st.error(result["error"])
                    else:
                        st.session_state[cache_key] = result

            result = st.session_state.get(cache_key)
            if result:
                d1, d2 = st.columns(2)
                with d1:
                    st.markdown("**Current**")
                    st.code(expr or "(empty)", language="dax")
                with d2:
                    st.markdown("**AI-suggested**")
                    st.code(result["revised_dax"] or "(no change suggested)", language="dax")
                if result["notes"]:
                    st.markdown("**Notes**")
                    st.write(result["notes"])
                if result["model_suggestions"]:
                    st.markdown("**Model-change suggestions**")
                    st.info(result["model_suggestions"], icon=":material/lightbulb:")

    batch_results = st.session_state.get(batch_key, {})
    if not batch_results:
        return df

    def _cell(name: str, key: str) -> str:
        r = batch_results.get(name)
        if r is None:
            return ""
        if r.get("error"):
            return f"ERROR: {r['error']}" if key == "revised_dax" else ""
        if key == "revised_dax":
            return r["revised_dax"] or "(no change suggested)"
        return r.get(key) or ""

    out = df.copy()
    names = out[name_col].astype(str)
    out["AI-Suggested DAX"] = names.map(lambda n: _cell(n, "revised_dax"))
    out["AI Notes"] = names.map(lambda n: _cell(n, "notes"))
    out["AI Model Suggestions"] = names.map(lambda n: _cell(n, "model_suggestions"))
    return out


# ==========================================================================
# UI
# ==========================================================================

inject_css()

st.markdown(
    """
    <div class="app-hero">
      <h1>🧩 VPAX Semantic Model Explorer</h1>
      <p>Explore, audit and govern the semantic model behind a Power BI report — ER
         diagrams, lineage, DAX quality, VertiPaq compression, RLS propagation, metric
         drift and Fabric readiness — straight from a .vpax export.</p>
      <span class="badge">Made by Sourin</span>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown('<div class="sb-title">⚙️ Data source</div>', unsafe_allow_html=True)
    with st.expander("Upload .vpax file", expanded=True, icon=":material/upload_file:"):
        uploaded = st.file_uploader("Choose a .vpax file", type=["vpax"], label_visibility="collapsed")
        st.caption(
            "Export from **DAX Studio** (Advanced ➜ Export Metadata) or "
            "**Tabular Editor** (File ➜ Export ➜ Metadata)."
        )
        if st.button(
            "Try a sample model", width="stretch", key="btn_sample",
            icon=":material/science:",
            help="Loads a small built-in star schema so you can see every section "
                 "without exporting anything first.",
        ):
            st.session_state["use_sample"] = True
    with st.expander("Add report pages (.pbix or .pbit) — optional", icon=":material/dashboard:"):
        pbix_file = st.file_uploader(
            "Choose a .pbix or .pbit file", type=["pbix", "pbit"], label_visibility="collapsed"
        )
        st.caption(
            "A .vpax holds the **semantic model only** and carries no report pages. Add a "
            "**.pbix** (a real report) or a **.pbit** (a template — same report layout, no "
            "data) to list the real pages and see the model per page."
        )
    with st.expander("Compare against a baseline (.vpax) — optional", icon=":material/compare_arrows:"):
        baseline_file = st.file_uploader(
            "Choose the certified/baseline .vpax", type=["vpax"],
            label_visibility="collapsed", key="baseline_upload",
        )
        st.caption(
            "Upload the **certified** model here and your working model above. "
            "The *Model Compare* section then shows which measures have drifted."
        )

    st.markdown('<div class="sb-title" style="margin-top:1rem">AI assistant</div>', unsafe_allow_html=True)
    with st.expander("Connect your own OpenAI or Claude key — optional", icon=":material/smart_toy:"):
        st.caption(
            "Set this once and it powers every AI feature in the app — the Scorecard's AI "
            "Insights and the DAX Assistant on Measures/Calculated Columns. Nothing is sent "
            "anywhere until you click a **Generate**/**Regenerate** button. Your key is kept "
            "only in this session's server-side memory and is never written to disk."
        )
        llm_provider = st.selectbox("Provider", list(LLM_PROVIDERS), key="llm_provider")
        st.text_input(
            "API key", type="password", key="llm_api_key",
            placeholder=LLM_PROVIDERS[llm_provider].get("key_placeholder", "sk-…"),
        )
        # Keyed per-provider so switching providers doesn't leave the other
        # provider's model choice sitting in the box; mirrored into a
        # provider-agnostic key so downstream code has one stable place to
        # read the active model id from regardless of which provider is set.
        _model_options = LLM_PROVIDERS[llm_provider]["models"]
        _custom_choice = "✏️ Custom (type a model ID)…"
        llm_model_choice = st.selectbox(
            "Model", _model_options + [_custom_choice],
            key=f"llm_model_choice_{llm_provider}",
            help="Curated from each provider's current model lineup. Pick "
                 "Custom to type any other model ID (e.g. a newer release).",
        )
        if llm_model_choice == _custom_choice:
            llm_model_id = st.text_input(
                "Custom model ID", key=f"llm_model_id_custom_{llm_provider}",
                placeholder=_model_options[0],
            ).strip() or _model_options[0]
        else:
            llm_model_id = llm_model_choice
        st.session_state["llm_model_id_active"] = llm_model_id

# A real upload always wins over the sample, so the sample never sticks
# around confusingly once the user brings their own file.
if uploaded is not None:
    st.session_state["use_sample"] = False
    source_bytes: Optional[bytes] = uploaded.getvalue()
    is_sample = False
elif st.session_state.get("use_sample"):
    source_bytes, is_sample = build_sample_vpax_bytes(), True
else:
    source_bytes, is_sample = None, False

if source_bytes is None:
    st.info(
        "👈 Upload a **.vpax** file from the sidebar to get started — or hit "
        "**Try a sample model** to explore with a built-in one."
    )
    st.markdown(f'<div class="app-footer">{AUTHOR}</div>', unsafe_allow_html=True)
    st.stop()

try:
    with st.spinner("Parsing model metadata…"):
        model = load_model(source_bytes)
except Exception as exc:  # noqa: BLE001
    st.error(f"Could not parse this file as a vpax model: {exc}")
    st.stop()

if is_sample:
    st.info(
        "Showing the **built-in sample model** — a four-table star schema with a few "
        "deliberate flaws so the audit sections have something real to report. "
        "Upload your own .vpax in the sidebar to replace it.",
        icon="🧪",
    )

baseline_model: Optional[Dict[str, Any]] = None
if baseline_file is not None:
    try:
        baseline_model = load_model(baseline_file.getvalue())
    except Exception as exc:  # noqa: BLE001
        st.warning(f"Could not read the baseline .vpax: {exc}", icon="⚠️")

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
    model should degrade that tab only - the other tabs still work, and the
    error is reported in place with enough detail to act on.
    """
    try:
        yield
    except Exception as exc:  # noqa: BLE001
        st.error(f"Something went wrong rendering **{tab_name}**: `{type(exc).__name__}: {exc}`")
        with st.expander("Technical details", icon=":material/code:"):
            st.code(traceback.format_exc(), language="text")
        st.caption(
            "The other tabs are unaffected. If this looks like a bug, the details above "
            "identify exactly where it happened."
        )


# ==========================================================================
# One scan, reused everywhere
# ==========================================================================
# Every audit section, the scorecard, the sidebar badges and the Excel export
# read the same finding sets. Computing them once - cached on the uploaded
# bytes, so a rerun costs nothing - keeps a big model responsive and
# guarantees the scorecard can never disagree with the tab it links to.

@st.cache_data(show_spinner=False)
def compute_all_findings(file_bytes: bytes, duplicate_threshold: int = 85) -> Dict[str, Any]:
    m = load_model(file_bytes)
    graph = build_measure_graph(m)
    taxonomy_tree, taxonomy_issues = build_taxonomy(m)
    rls_sim = simulate_rls(m)
    return {
        "health": run_health_checks(m),
        "naming": lint_naming(m),
        "unused": find_unused_columns(m),
        "encoding": build_encoding_advice(m),
        "fabric": build_fabric_readiness(m),
        "taxonomy_tree": taxonomy_tree,
        "taxonomy_issues": taxonomy_issues,
        "rls_sim": rls_sim,
        "rls_summary": rls_exposure_summary(rls_sim, m),
        "measure_graph": graph,
        "cycles": find_cycles(graph),
        "size": build_model_size_summary(m),
        "duplicates": find_near_duplicate_measures(m, duplicate_threshold),
    }


@st.cache_data(show_spinner=False)
def compute_report_findings(
    file_bytes: bytes, pbix_bytes: Optional[bytes], unused_df: pd.DataFrame
) -> Dict[str, Any]:
    """Checks that need the model and the report together.

    Kept separate from the model-only scan so uploading a .pbix doesn't
    invalidate the (much more expensive) model cache, and so every one of
    these degrades to a clear "upload the .pbix" state rather than an error
    when the report isn't there.
    """
    m = load_model(file_bytes)
    rep = load_report_pages(pbix_bytes) if pbix_bytes else None
    bindings = _report_bindings(rep)
    return {
        "bindings": bindings,
        "broken": validate_report_bindings(bindings, m),
        "disposition": classify_column_disposition(m, unused_df, bindings),
        "field_params": find_field_parameters(m, bindings),
    }


with st.status("Scanning the model…", expanded=False) as _scan_status:
    findings = compute_all_findings(source_bytes)
    report_findings = compute_report_findings(
        source_bytes, pbix_file.getvalue() if pbix_file is not None else None,
        findings["unused"],
    )
    _scan_status.update(
        label=f"Scanned {len(model['all_table_names'])} tables, "
              f"{len(model['measures'])} measures"
              + (f", {len(report_findings['bindings'])} report field bindings"
                 if not report_findings["bindings"].empty else ""),
        state="complete",
    )

health_df = findings["health"]
naming_df = findings["naming"]
unused_df = findings["unused"]
encoding_df = findings["encoding"]
fabric_df = findings["fabric"]
taxonomy_tree = findings["taxonomy_tree"]
taxonomy_issues = findings["taxonomy_issues"]
rls_sim_df = findings["rls_sim"]
rls_summary_df = findings["rls_summary"]
duplicates_df = findings["duplicates"]
bindings_df = report_findings["bindings"]
broken_df = report_findings["broken"]
disposition_df = report_findings["disposition"]
field_params_df = report_findings["field_params"]
has_report = not bindings_df.empty


# ==========================================================================
# Navigation
# ==========================================================================
# Grouped by *what you came here to do*, not by object type. Eighteen flat
# tabs meant a junior developer had to already know which tab answered their
# question; four verbs plus a scorecard means they don't.
NAV_GROUPS: Dict[str, List[str]] = {
    "🧭 Overview": ["Model Scorecard"],
    "🔍 Explore": [
        "Dashboard Screens", "Semantic Model by Screen", "Tables", "Columns / Schema",
        "Measures", "Calculated Columns", "Relationships", "Power Query (SQL)",
        "Display Folders",
    ],
    "🩺 Audit": [
        "Model Health", "Naming Conventions", "Unused Objects", "Report Usage",
        "Duplicate Measures", "Field Parameters", "Date Table Check",
        "Model Size", "Compression Advisor", "Fabric Readiness",
    ],
    "🔬 Investigate": [
        "Impact Analysis", "Measure Dependencies", "Source Lineage", "Model Compare",
    ],
    "🛡️ Govern": [
        "Security & Perspectives", "RLS Simulator", "Fix Script (C#)", "Model Cleanup",
        "Data Dictionary Export", "Theme",
    ],
    "🧱 Migrate": ["Databricks Lakeview Export", "TWBX → Power BI"],
}

# Which finding set backs each page, so the scorecard and the sidebar badges
# both derive from one declaration instead of two hand-maintained lists.
PAGE_FINDINGS: Dict[str, Tuple[str, pd.DataFrame, str]] = {
    "Model Health": ("Best-practice violations", health_df, "Severity"),
    "Naming Conventions": ("Objects off the dominant convention", naming_df, "Severity"),
    "Unused Objects": ("Columns nothing references", unused_df, "Severity"),
    "Compression Advisor": ("Columns wasting memory", encoding_df, "Severity"),
    "Fabric Readiness": ("Direct Lake blockers", fabric_df, "Severity"),
    "Display Folders": ("Taxonomy gaps", taxonomy_issues, "Severity"),
    "RLS Simulator": ("Roles with unsecured tables", rls_summary_df, "Severity"),
    "Report Usage": ("Broken visuals & deletable columns",
                     pd.concat([broken_df[["Severity"]], disposition_df[["Severity"]]],
                               ignore_index=True) if has_report else pd.DataFrame(),
                     "Severity"),
    "Duplicate Measures": ("Near-identical DAX", duplicates_df, "Severity"),
    "Field Parameters": ("Abandoned parameter tables", field_params_df, "Severity"),
}


def _page_severity_counts(page: str) -> Dict[str, int]:
    entry = PAGE_FINDINGS.get(page)
    if not entry:
        return {}
    _, df, col = entry
    counts = severity_counts(df, col)
    counts.pop("Info", None)  # Info is context, not a finding
    return counts


def _group_severity_counts(group: str) -> Dict[str, int]:
    total: Dict[str, int] = {}
    for page in NAV_GROUPS[group]:
        for sev, n in _page_severity_counts(page).items():
            total[sev] = total.get(sev, 0) + n
    return total


def _goto(group: str, page: str) -> None:
    """Jump the nav somewhere else — used by the scorecard's issue cards."""
    st.session_state["nav_group"] = group
    st.session_state["nav_page"] = page


def _page_group(page: str) -> str:
    for g, pages in NAV_GROUPS.items():
        if page in pages:
            return g
    return "🧭 Overview"


def _badge(counts: Dict[str, int]) -> str:
    if counts.get("High"):
        return f" 🔴 {counts['High']}"
    if counts.get("Medium"):
        return f" 🟠 {counts['Medium']}"
    if counts.get("Low"):
        return f" ⚪ {counts['Low']}"
    return " 🟢"


with st.sidebar:
    st.markdown('<div class="sb-title">🔎 Find anything</div>', unsafe_allow_html=True)
    search_term = st.text_input(
        "search", placeholder="table, column or measure…",
        label_visibility="collapsed", key="global_search",
    )

# Sections live as a two-tier segmented control (group, then page within it)
# right below the header, not in the sidebar - the sidebar is reserved for
# data-source uploads and the global search. st.segmented_control returns
# None on first render unless a `default` is given, but Streamlit also warns
# if `default` is passed once the key already has a session_state entry - so
# `default` is only supplied on the very first render for that key. On every
# later render (including a scorecard "Open ->" jump, which pre-sets
# st.session_state[key] before the widget runs) the existing value wins,
# same as st.radio before it.
group_options = list(NAV_GROUPS)
nav_group = st.segmented_control(
    "Sections", group_options,
    format_func=lambda g: g + ("" if g == "🧭 Overview" else _badge(_group_severity_counts(g))),
    key="nav_group", label_visibility="collapsed", required=True,
    default=group_options[0] if "nav_group" not in st.session_state else None,
    width="stretch",
)
if nav_group not in NAV_GROUPS:
    # Guards against a stale/mismatched widget-state value surviving a
    # hot-reload (e.g. `required=True` doesn't always stop segmented_control
    # from resolving to None when the session's cached selection can't be
    # reconciled with the current run) - fall back to the first group rather
    # than crashing the whole page on a KeyError.
    nav_group = group_options[0]

_pages = NAV_GROUPS[nav_group]
if st.session_state.get("nav_page") not in _pages:
    st.session_state["nav_page"] = _pages[0]

if len(_pages) > 1:
    nav_page = st.segmented_control(
        "Pages", _pages, key="nav_page", label_visibility="collapsed",
        format_func=lambda p: p + (_badge(_page_severity_counts(p)) if p in PAGE_FINDINGS else ""),
        required=True,
        default=_pages[0] if "nav_page" not in st.session_state else None,
        width="stretch",
    )
    if nav_page not in _pages:
        nav_page = _pages[0]
else:
    nav_page = _pages[0]

st.caption("🔴 high severity · 🟠 medium · ⚪ low · 🟢 clear")

# --- Global search ------------------------------------------------------------
if search_term and search_term.strip():
    with tab_guard("Search"):
        term = search_term.strip().lower()
        hits: List[Dict[str, str]] = []
        for t in model["all_table_names"]:
            if term in t.lower():
                hits.append({"Type": "Table", "Table": t, "Name": t, "Detail": ""})
        _cols = _user_facing_columns(model["columns"])
        if {"TableName", "ColumnName"}.issubset(_cols.columns):
            for _, c in _cols.dropna(subset=["ColumnName"]).iterrows():
                if term in str(c["ColumnName"]).lower():
                    hits.append({"Type": "Column", "Table": str(c["TableName"]),
                                 "Name": str(c["ColumnName"]),
                                 "Detail": str(c.get("DataType") or "")})
        _meas = model["measures"]
        if {"TableName", "MeasureName"}.issubset(_meas.columns):
            for _, m in _meas.dropna(subset=["MeasureName"]).iterrows():
                name, expr = str(m["MeasureName"]), str(m.get("MeasureExpression") or "")
                if term in name.lower() or term in expr.lower():
                    where = "name" if term in name.lower() else "DAX"
                    hits.append({"Type": "Measure", "Table": str(m["TableName"]),
                                 "Name": name, "Detail": f"matched in {where}"})
        st.markdown(f"### 🔎 {len(hits)} match(es) for “{html.escape(search_term.strip())}”")
        if not hits:
            st.info("Nothing in this model matches that. Search covers table, column and "
                    "measure names, plus the text of every measure's DAX.")
        else:
            show_table(pd.DataFrame(hits), f"search_{_slug(search_term)}", height=320, key="search_hits")
            st.caption("Use **Investigate ➜ Impact Analysis** to see what references any of these "
                       "before you rename or delete it.")
        st.space("large")


# --- Model Scorecard ----------------------------------------------------------
if nav_page == "Model Scorecard":
    with tab_guard("Model Scorecard"):
        st.subheader("Model scorecard")
        st.caption(
            "Every check in this app, rolled into one view. The score starts at 100 and "
            "loses points per finding weighted by severity — it's a relative health "
            "signal for triage, not an official Microsoft metric."
        )

        # Each section can lose at most an equal share of the 100 points, and
        # within a section it's the *worst* severity that dominates, not the
        # raw count. Straight count-weighting doesn't work here: nearly every
        # real model has hundreds of missing descriptions, which would peg
        # every model at the same rock-bottom score and hide the one genuine
        # high-severity problem that actually needs attention.
        SECTION_CAP = 100.0 / max(len(PAGE_FINDINGS), 1)
        BASE_BY_WORST = {"High": 0.60, "Medium": 0.35, "Low": 0.12, "Clean": 0.0}

        penalty = 0.0
        cards: List[Dict[str, Any]] = []
        for page, (title, df, col) in PAGE_FINDINGS.items():
            counts = _page_severity_counts(page)
            worst = "Clean"
            for s in ("High", "Medium", "Low"):
                if counts.get(s):
                    worst = s
                    break
            share = BASE_BY_WORST[worst]
            if worst != "Clean":
                # How widespread the worst class is scales the remaining
                # headroom, so 1 high-severity finding and 40 don't score alike.
                spread = min(1.0, counts.get(worst, 0) / 10.0)
                share += (1.0 - share) * spread * 0.7
            penalty += SECTION_CAP * share
            cards.append({"page": page, "title": title, "counts": counts, "worst": worst})

        score = max(0, min(100, round(100 - penalty)))
        grade, colour = (
            ("Healthy", "#16a34a") if score >= 85 else
            ("Needs attention", "#d97706") if score >= 60 else
            ("At risk", "#dc2626")
        )
        total_high = sum(c["counts"].get("High", 0) for c in cards)
        total_med = sum(c["counts"].get("Medium", 0) for c in cards)
        total_low = sum(c["counts"].get("Low", 0) for c in cards)

        head = st.columns([1, 3])
        with head[0]:
            st.markdown(
                f'<div class="score-ring" style="background:{colour}">'
                f'<span class="num">{score}</span><span class="lbl">score</span></div>',
                unsafe_allow_html=True,
            )
            st.markdown(
                f"<div style='text-align:center;font-weight:600;color:{colour};margin-top:.4rem'>"
                f"{grade}</div>", unsafe_allow_html=True,
            )
        with head[1]:
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("🔴 High", total_high, help="Likely wrong, insecure, or a real performance risk.")
            k2.metric("🟠 Medium", total_med, help="A design choice worth double-checking.")
            k3.metric("⚪ Low", total_low, help="Hygiene and documentation gaps.")
            k4.metric("Tables", len(model["all_table_names"]))
            if total_high:
                st.error(f"**{total_high} high-severity finding(s)** — start there.", icon="🔴")
            elif total_med:
                st.warning(f"No high-severity findings. {total_med} medium item(s) to review.", icon="🟠")
            else:
                st.success("No high or medium findings. This model is in good shape.", icon="✅")

        with st.expander("AI insights — a prioritized summary of everything below",
                          icon=":material/smart_toy:", expanded=True):
            if not _llm_ready():
                st.caption(
                    "Set an OpenAI or Claude API key in the sidebar's **AI assistant** section "
                    "to turn the raw findings below into a short, prioritized, plain-language "
                    "summary — what to fix first and what can wait."
                )
            else:
                findings_fingerprint = "|".join(f"{c['page']}:{c['worst']}:{sum(c['counts'].values())}" for c in cards)
                insights_key = f"scorecard_insights_{score}_{hash(findings_fingerprint) & 0xffffffff}"
                provider = st.session_state.get("llm_provider") or next(iter(LLM_PROVIDERS))
                if st.button("✨ Generate AI insights", key="scorecard_insights_go"):
                    with st.spinner(f"Asking {provider} to review this model…"):
                        try:
                            prompt = build_insights_prompt(model, score, grade, cards)
                            st.session_state[insights_key] = call_llm(prompt)
                        except LLMError as exc:
                            st.error(str(exc))
                        except Exception as exc:  # noqa: BLE001
                            st.error(f"Unexpected error calling {provider}: {exc}")
                insight_text = st.session_state.get(insights_key)
                if insight_text:
                    st.markdown(insight_text)

        st.markdown("#### Where the findings are")
        cards.sort(key=lambda c: (SEVERITY_ORDER.get(c["worst"], 9),
                                  -sum(c["counts"].values())))
        for i, c in enumerate(cards):
            row = st.columns([6, 1])
            bits = " · ".join(f"{SEVERITY_ICON[s]} {c['counts'][s]} {s.lower()}"
                              for s in ("High", "Medium", "Low") if c["counts"].get(s))
            desc = bits or "Nothing flagged."
            row[0].markdown(
                f'<div class="issue-card sev-{c["worst"]}">'
                f'<div class="title">{html.escape(c["page"])}'
                f'<span class="badge-pill {c["worst"]}">{c["worst"]}</span></div>'
                f'<div class="desc">{html.escape(c["title"])} — {desc}</div></div>',
                unsafe_allow_html=True,
            )
            row[1].button(
                "Open →", key=f"goto_{i}", width="stretch",
                on_click=_goto, args=(_page_group(c["page"]), c["page"]),
            )

        with st.expander("Guided review checklist — first pass on an unfamiliar model", icon=":material/checklist:"):
            st.markdown(
                "Work top to bottom. Each step names the section that answers it.\n\n"
                "1. **Does it load and what's in it?** — *Explore ➜ Tables* and *Measures*.\n"
                "2. **Is the shape a star schema?** — *Explore ➜ Relationships*; facts should "
                "sit in the middle with dimensions on the one side.\n"
                "3. **Is there one marked date table?** — *Audit ➜ Date Table Check*. Time "
                "intelligence silently misbehaves without it.\n"
                "4. **Anything obviously wrong?** — *Audit ➜ Model Health*, high severity first.\n"
                "5. **Where is the memory going?** — *Audit ➜ Model Size*, then "
                "*Compression Advisor* for the specific fix.\n"
                "6. **Is security real?** — *Govern ➜ RLS Simulator*, not just the role list.\n"
                "7. **Can it move to Fabric?** — *Audit ➜ Fabric Readiness*.\n"
                "8. **Before you change anything** — *Investigate ➜ Impact Analysis* on the "
                "object you're about to touch.\n"
                "9. **Hand it over** — *Govern ➜ Data Dictionary Export*, and "
                "*Fix Script (C#)* for the mechanical cleanups."
            )

# --- Dashboard Screens ----------------------------------------------------
if nav_page == "Dashboard Screens":
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
            page_names = sorted(view["Screen / Page Name"].dropna().astype(str).tolist())
            if page_names:
                picked_page = st.selectbox("Summarize a page", page_names, key="page_summary_pick")
                render_ai_page_summary(picked_page, report, model, key_prefix="ai_page_summary")
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
if nav_page == "Semantic Model by Screen":
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
                                   f"{screen_table} measures", height=330, key=f"meas_{idx}", row_height=110)

# --- Tables ---------------------------------------------------------------
if nav_page == "Tables":
    with tab_guard('Tables'):
        st.subheader("Tables in the model")
        st.caption(
            "Every table the .vpax export knows about, with whatever metadata came with it "
            "(row count, description, storage mode). Use this as a starting inventory before "
            "diving into a specific table's columns or relationships."
        )
        show_table(model["tables"], "Tables", height=420, key="tables", row_height=90)

# --- Columns / Schema -----------------------------------------------------
if nav_page == "Columns / Schema":
    with tab_guard('Columns / Schema'):
        st.subheader("Columns per table")
        st.caption(
            "Every column across every table, including hidden ones and VertiPaq's internal "
            "RowNumber columns (unlike the audit tabs, which filter those out as noise). Filter "
            "to one or a few tables below, or use **Explore ➜ Impact Analysis** once you've "
            "found a specific column you're thinking about changing."
        )
        schema_df = model["columns"]
        if schema_df.empty:
            st.info("No column metadata found.")
        elif "TableName" not in schema_df.columns:
            # Filtering needs the column; without it just show what we have.
            show_table(schema_df.reset_index(drop=True), "Columns", height=460, key="columns", row_height=90)
        else:
            # key=str: a model can mix numeric and text table names, and
            # sorting those against each other raises TypeError.
            options = sorted(schema_df["TableName"].dropna().unique().tolist(), key=str)
            chosen = st.multiselect("Filter by table", options=options, key="schema_filter")
            view = schema_df[schema_df["TableName"].isin(chosen)] if chosen else schema_df
            show_table(view.reset_index(drop=True), "Columns", height=460, key="columns", row_height=90)

# --- Measures -------------------------------------------------------------
if nav_page == "Measures":
    with tab_guard('Measures'):
        st.subheader("DAX measures")
        meas_df = model["measures"]
        if meas_df.empty:
            st.info("No measures found in this model.")
        else:
            st.caption(
                "Every measure's DAX, as stored in the model. Use the AI DAX Assistant above to "
                "get a best-practice rewrite and modelling suggestions — for one measure, or for "
                "all of them at once as extra, downloadable columns on the table below."
            )
            meas_view_df = render_ai_dax_assistant(
                "measure", meas_df, "MeasureName", "MeasureExpression", model, key_prefix="ai_meas"
            )
            render_apply_dax_script(
                meas_df, "MeasureName", "measure", ai_key_prefix="ai_meas", apply_key_prefix="apply_meas"
            )
            search = st.text_input("Search measure name or expression…", key="measure_search")
            view = meas_view_df
            if search:
                view = meas_view_df[
                    meas_view_df["MeasureName"].str.contains(search, case=False, na=False)
                    | meas_view_df["MeasureExpression"].str.contains(search, case=False, na=False)
                ]
            show_table(view.reset_index(drop=True), "Measures", height=460, key="measures", row_height=130)

# --- Calculated Columns ---------------------------------------------------
if nav_page == "Calculated Columns":
    with tab_guard('Calculated Columns'):
        st.subheader("Calculated columns")
        st.caption(
            "Columns computed by a row-context DAX expression instead of loaded from the "
            "source, stored on disk like any other column. Worth a second look if there are "
            "many of these — **Audit ➜ Model Health** flags ones that look cheap enough to move "
            "upstream into Power Query/SQL instead, which computes them once at refresh instead "
            "of storing a value per row."
        )
        cols_df = model["calc_columns"]
        if cols_df.empty:
            st.info("No calculated columns found in this model.")
        else:
            cols_view_df = render_ai_dax_assistant(
                "calculated column", cols_df, "ColumnName", "ColumnExpression", model, key_prefix="ai_calc"
            )
            render_apply_dax_script(
                cols_df, "ColumnName", "calculated column", ai_key_prefix="ai_calc", apply_key_prefix="apply_calc"
            )
            show_table(cols_view_df.reset_index(drop=True), "Calculated Columns", height=460, key="calccols", row_height=130)

# --- Relationships --------------------------------------------------------
if nav_page == "Relationships":
    with tab_guard('Relationships'):
        st.subheader("Relationships")
        st.caption(
            "Every join the model defines, with cardinality and filter direction. **Bi-"
            "directional** relationships are worth a second look — see **Audit ➜ Model Health** "
            "for why. **Inactive** relationships exist but don't filter unless a measure "
            "explicitly activates them with USERELATIONSHIP()."
        )
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
if nav_page == "Power Query (SQL)":
    with tab_guard('Power Query (SQL)'):
        st.subheader("Source SQL behind each table")
        pq_df = model["power_query"]
        if pq_df.empty:
            st.info("No Power Query expressions found in this model.")
        else:
            sql_len = pq_df["SQL"].fillna("").astype(str).str.len()
            with_sql = pq_df[sql_len > 0]
            st.caption(f"{len(with_sql)} of {len(pq_df)} partitions embed a SQL query.")
            pq_view_df = render_ai_source_descriptions(pq_df, model, key_prefix="ai_source_desc")
            source_desc_cache: Dict[str, Any] = st.session_state.get("ai_source_desc_batch_results", {})
            choice = st.selectbox(
                "View the full SQL for a table",
                ["(all — summary table)"]
                + sorted(with_sql["TableName"].dropna().unique().tolist(), key=str),
            )
            if choice.startswith("(all"):
                show_table(pq_view_df.reset_index(drop=True), "Power Query SQL", height=460, key="pq", row_height=140)
            else:
                matches = with_sql.loc[with_sql["TableName"] == choice, "SQL"]
                sql = str(matches.iloc[0]) if not matches.empty else ""
                st.code(sql or "-- no SQL found for this table", language="sql")
                desc_result = source_desc_cache.get(choice)
                if desc_result and not desc_result.get("error") and desc_result.get("description"):
                    st.info(desc_result["description"], icon=":material/smart_toy:")
                if sql:
                    sources = _sql_sources(sql)
                    if sources:
                        names = sorted({".".join(x for x in s if x) for s in sources})
                        st.caption(
                            "**Reads from:** " + ", ".join(f"`{n}`" for n in names)
                            + " — see **Investigate ➜ Source Lineage** for dialect selection, "
                              "blast-radius analysis, and a diagram."
                        )
                st.download_button(
                    "⬇ Download SQL", sql.encode("utf-8"),
                    file_name=f"{_slug(choice, 'query')}.sql",
                    mime="text/plain", key="dl_sql", disabled=not sql,
                )

# --- Model Size (VertiPaq) -------------------------------------------------
if nav_page == "Model Size":
    with tab_guard('Model Size'):
        st.subheader("Model size & VertiPaq statistics")
        st.caption(
            "**What this is:** Power BI/Analysis Services loads your model into an in-memory "
            "column store called VertiPaq, and this .vpax export carries that engine's own "
            "size/cardinality stats per table and column. **Why it matters:** if your model "
            "feels slow or your .pbix is large, the tables/columns below are exactly where the "
            "memory is going — the usual fixes are removing unused columns (see the Unused "
            "Objects tab), reducing cardinality on high-distinct-value columns, or disabling "
            "columns entirely in favour of a measure. All sizes are shown in MB (compressed "
            "in-memory size, not file size)."
        )
        size_summary = findings["size"]
        if not size_summary.get("available"):
            st.info(
                "This .vpax export doesn't include VertiPaq size/cardinality statistics. "
                "Re-export with a tool that captures them (e.g. DAX Studio's Advanced ➜ Export "
                "Metadata, or Tabular Editor's Vertipaq Analyzer)."
            )
        else:
            if "total_model_size_mb" in size_summary:
                st.metric("Total model size", f'{size_summary["total_model_size_mb"]:,.1f} MB')
            if "top_tables" in size_summary:
                st.markdown("**Top tables by size (MB)**")
                show_table(size_summary["top_tables"], "Top Tables By Size", height=320, key="vpa_tables")
            if "top_columns" in size_summary:
                st.markdown("**Top columns by size (MB)** — biggest opportunities to shrink the model")
                show_table(size_summary["top_columns"], "Top Columns By Size", height=320, key="vpa_columns")
            if "top_cardinality" in size_summary:
                st.markdown("**Highest-cardinality columns** (count of distinct values)")
                show_table(size_summary["top_cardinality"], "Top Cardinality Columns", height=320, key="vpa_cardinality")
            if "dict_vs_data" in size_summary:
                st.markdown(
                    "**Dictionary vs. data size (MB)** — a column's Dictionary Size holds its "
                    "distinct values; Data Size holds the per-row encoded pointers into that "
                    "dictionary. A dictionary much larger than the data size usually means the "
                    "column has too many distinct values for how it's being used."
                )
                show_table(size_summary["dict_vs_data"], "Dictionary Vs Data Size", height=320, key="vpa_dict")
            if "encoding_breakdown" in size_summary:
                st.markdown("**Encoding breakdown** — how many columns use each VertiPaq encoding")
                show_table(size_summary["encoding_breakdown"], "Encoding Breakdown", height=220, key="vpa_encoding")

# --- Unused Objects ---------------------------------------------------------
if nav_page == "Unused Objects":
    with tab_guard('Unused Objects'):
        st.subheader("Columns not referenced by any measure, calculated column, or relationship")
        st.caption(
            "A static DAX reference scan — it can't see Power BI report visuals (a .vpax carries "
            "none) or RLS filter expressions. Treat 'Likely unused' as a starting point for "
            "review, not a guarantee it's safe to delete."
        )
        if unused_df.empty:
            st.info("No column metadata available to check.")
        else:
            likely_unused = int((unused_df["Status"] == "Likely unused").sum())
            c1, c2 = st.columns(2)
            c1.metric("⚪ Likely unused columns", likely_unused)
            c2.metric("Total columns", len(unused_df))
            only_unused = st.checkbox("Hide referenced columns", value=True, key="unused_only")
            view = unused_df[unused_df["Status"] == "Likely unused"] if only_unused else unused_df
            if view.empty:
                st.success("Every column is referenced somewhere in the model.", icon="✅")
            else:
                view = render_ai_unused_risk(view, model, key_prefix="ai_unused")
                show_table(view, "Unused Columns", height=460, key="unused")
                if has_report:
                    rescued = int((disposition_df["Verdict"] == "Keep — used on report visuals only").sum())
                    safe_n = int(disposition_df["Verdict"].astype(str).str.startswith("Safe to delete").sum())
                    st.success(
                        f"The .pbix is loaded, so this list has been checked against the report: "
                        f"**{rescued}** of these are used on a visual after all, and **{safe_n}** "
                        "are genuinely safe to delete. See **Audit ➜ Report Usage** for the verdict "
                        "per column.", icon="🖥️",
                    )
                else:
                    st.caption(
                        "Before removing any of these, confirm in **Investigate ➜ Impact Analysis** "
                        "— and upload the matching **.pbix** to turn \"likely unused\" into a real "
                        "verdict. A column used only on a visual axis looks completely unused to a "
                        "DAX-only scan."
                    )

# --- Impact Analysis ---------------------------------------------------------
if nav_page == "Impact Analysis":
    with tab_guard('Impact Analysis'):
        st.subheader("What references this table or column?")
        st.caption(
            "Pick a table or column to see every measure, calculated column, and relationship "
            "that touches it before you rename or delete it. Column lookups count only explicit "
            "`'Table'[Column]`-qualified DAX references, since the same column name often exists "
            "on more than one table (e.g. a join key shared by a fact and a dimension) — that's "
            "why you pick the table first."
        )
        target_kind = st.segmented_control(
            "Look up a", ["Table", "Column"], key="impact_kind", required=True,
            default="Table" if "impact_kind" not in st.session_state else None,
        )

        result = None
        label = ""
        if target_kind == "Table":
            options = model["all_table_names"]
            if not options:
                st.info("No tables found in this model.")
            else:
                target_name = st.selectbox("Choose a table", options, key="impact_target_table")
                result = impact_of("table", target_name, model)
                label = target_name
        else:
            table_options = model["all_table_names"]
            if not table_options:
                st.info("No tables found in this model.")
            else:
                picked_table = st.selectbox("Table", table_options, key="impact_col_table")
                cols_df = _user_facing_columns(model["columns"])
                col_options = []
                if {"TableName", "ColumnName"}.issubset(cols_df.columns):
                    col_options = sorted(
                        cols_df.loc[cols_df["TableName"] == picked_table, "ColumnName"].dropna().unique().tolist(),
                        key=str,
                    )
                if not col_options:
                    st.info(f"No columns found on **{picked_table}**.")
                else:
                    target_name = st.selectbox("Column", col_options, key="impact_col_col")
                    result = impact_of("column", target_name, model, table_name=picked_table)
                    label = f"{picked_table}[{target_name}]"

        if result is not None:
            st.markdown(f"**Measures referencing `{label}`** ({len(result['measures'])})")
            show_table(result["measures"].reset_index(drop=True), f"{label} measures", height=260, key="impact_measures", row_height=110)
            st.markdown(f"**Calculated columns referencing `{label}`** ({len(result['calc_columns'])})")
            show_table(result["calc_columns"].reset_index(drop=True), f"{label} calc columns", height=220, key="impact_calc", row_height=110)
            st.markdown(f"**Relationships involving `{label}`** ({len(result['relationships'])})")
            show_table(result["relationships"].reset_index(drop=True), f"{label} relationships", height=220, key="impact_rels")
            if target_kind == "Table" and not result["related_tables"].empty:
                st.markdown("**Directly related tables**")
                show_table(result["related_tables"], f"{label} related tables", height=180, key="impact_related")

# --- Measure Dependencies ----------------------------------------------------
if nav_page == "Measure Dependencies":
    with tab_guard('Measure Dependencies'):
        st.subheader("Which measures call other measures")
        st.caption(
            "\"Depends On\" is what a measure's own DAX calls; \"Used By\" is every measure that "
            "calls it. Both are read straight from each measure's expression - no report visuals "
            "involved. Blank cells just mean that measure neither calls nor is called by another "
            "measure in this model."
        )
        graph = findings["measure_graph"]
        if not graph:
            st.info("No measures found in this model.")
        else:
            cycles = findings["cycles"]
            if cycles:
                st.warning(
                    f"⚠️ {len(cycles)} circular measure reference(s) found — these can never "
                    "fully evaluate: " + "; ".join(" → ".join(c) for c in cycles[:5])
                )
            dep_table = build_measure_dependency_table(graph)
            show_table(dep_table, "Measure Dependencies", height=460, key="measure_deps", row_height=90)

            if any(graph.values()):
                with st.expander("Show as a diagram", icon=":material/hub:"):
                    participating = sorted(
                        {n for n, calls in graph.items() if calls} | {c for calls in graph.values() for c in calls}
                    )
                    focus_options = ["(whole graph)"] + participating
                    choice = st.selectbox("Focus on a measure (optional)", focus_options, key="measure_graph_focus")
                    focus = None if choice.startswith("(whole") else choice
                    dot = build_measure_dependency_dot(graph, focus=focus)
                    static_diagram_panel(dot, engine="dot", filename="measure_dependencies")

# --- Model Health -------------------------------------------------------------
if nav_page == "Model Health":
    with tab_guard('Model Health'):
        st.subheader("Model health & best-practice checks")
        st.caption(
            "A checklist of common Power BI modeling issues, ranked by how much they matter:\n"
            "- **High** — likely wrong or a real performance risk: circular measure references "
            "(can never evaluate), duplicate measure names, or a relationship key with millions "
            "of distinct values.\n"
            "- **Medium** — a design choice worth double-checking, not necessarily wrong: "
            "bi-directional or many-to-many relationships can cause ambiguous, double-counted, "
            "or slow results.\n"
            "- **Low** — hygiene/documentation gaps that don't affect correctness: missing "
            "descriptions, inconsistent format strings, or a calculated column that might be "
            "cheaper to compute upstream in Power Query."
        )
        if health_df.empty:
            st.success("No health-check findings for this model.", icon="✅")
        else:
            health_view_df = render_ai_grouped_fixes(
                health_df, model, key_prefix="ai_health",
                group_col="Rule", object_col="Object", message_col="Message", group_label="rule",
            )
            counts = health_view_df["Severity"].value_counts()
            c1, c2, c3 = st.columns(3)
            c1.metric("🔴 High", int(counts.get("High", 0)), help="Likely wrong or a real performance risk — worth fixing.")
            c2.metric("🟠 Medium", int(counts.get("Medium", 0)), help="A design choice worth double-checking.")
            c3.metric("⚪ Low", int(counts.get("Low", 0)), help="Hygiene/documentation gaps — doesn't affect correctness.")
            only_high = st.checkbox(
                "Show high severity only", value=bool(counts.get("High", 0) > 12),
                key="health_only_high",
            )
            view = health_view_df[health_view_df["Severity"] == "High"] if only_high else health_view_df
            show_table(sort_by_severity(view), "Model Health", height=460, key="health", row_height=120)
            st.caption(
                "Mechanical fixes for several of these — hiding join keys, dropping unused "
                "attribute hierarchies, adding format strings — are generated as a runnable "
                "script under **Govern ➜ Fix Script (C#)**."
            )

# --- Naming Conventions -------------------------------------------------------
if nav_page == "Naming Conventions":
    with tab_guard('Naming Conventions'):
        st.subheader("Naming convention consistency")
        if naming_df.empty:
            st.success("Every object matches the dominant convention in its group.", icon="✅")
        else:
            st.caption(
                f"{len(naming_df)} object(s) don't match the dominant naming convention in "
                "their group. Measure names are rated Medium because report authors see them; "
                "column and table names Low."
            )
            kinds = ["(all)"] + sorted(naming_df["Object Type"].dropna().unique().tolist())
            pick_kind = st.selectbox("Object type", kinds, key="naming_kind")
            view = naming_df if pick_kind == "(all)" else naming_df[naming_df["Object Type"] == pick_kind]
            show_table(sort_by_severity(view), "Naming Conventions", height=460, key="naming", row_height=90)
            st.caption(
                "**Govern ➜ Fix Script (C#)** emits these as ready-to-run renames — commented "
                "out, because a rename breaks every visual that referenced the old name. "
                "Check **Investigate ➜ Impact Analysis** first."
            )

# --- Date Table Check ---------------------------------------------------------
if nav_page == "Date Table Check":
    with tab_guard('Date Table Check'):
        st.subheader("Date table / time-intelligence check")
        st.info(
            "A .vpax carries model **metadata only**, not row data — this check can confirm a "
            "table is *marked* as a date table and has a plausible date column, but it "
            "**cannot** verify the dates are actually contiguous (no gaps) without the real data."
        )
        date_df = model["date_tables"]
        if date_df.empty:
            st.warning("No table is marked `dataCategory: Time`, and no table has a date-typed column.")
        else:
            show_table(date_df, "Date Tables", height=320, key="date_tables")

# --- Security & Perspectives ---------------------------------------------------
if nav_page == "Security & Perspectives":
    with tab_guard('Security & Perspectives'):
        st.subheader("Row-level security roles & perspectives")
        st.caption(
            "The raw role/perspective definitions from the model, as declared — not whether "
            "they actually work. A role's filter expression only *looks* correct here; "
            "**Govern ➜ RLS Simulator** traces whether it actually reaches the tables it's "
            "meant to secure. Perspectives are curated subsets of the model shown to specific "
            "tools/audiences — they don't restrict data access."
        )
        sub = st.tabs(["Roles (RLS)", "Perspectives"])
        with sub[0]:
            roles_df = model["roles"]
            if roles_df.empty:
                st.info("No RLS roles defined in this model.")
            else:
                show_table(roles_df, "Roles", height=380, key="roles", row_height=110)
        with sub[1]:
            persp_df = model["perspectives"]
            if persp_df.empty:
                st.info("No perspectives defined in this model.")
            else:
                show_table(persp_df, "Perspectives", height=380, key="perspectives")

# --- Data Dictionary Export -----------------------------------------------------
if nav_page == "Data Dictionary Export":
    with tab_guard('Data Dictionary Export'):
        st.subheader("Auto-generated data dictionary")
        st.caption(
            "One Excel workbook, one sheet per topic: tables, columns, measures, calculated "
            "columns, relationships, Power Query SQL, security/perspectives, and this app's "
            "health/naming/unused-object findings. Diagrams aren't embedded — the ER and "
            "measure-dependency diagrams render client-side only; download them separately from "
            "their own tabs."
        )
        ai_descriptions_df = render_ai_descriptions(model, key_prefix="ai_desc")

        if EXCEL_ENGINE is None:
            st.button(
                "⬇ Download Data Dictionary (.xlsx)", disabled=True, width="stretch",
                help="Excel export needs a writer library — run: pip install openpyxl",
            )
        else:
            # Reused from the single centralised scan, so the workbook can
            # never disagree with what the audit sections are showing.
            extra_sheets = [
                ("Compression Advice", encoding_df),
                ("Fabric Readiness", fabric_df),
                ("Taxonomy Issues", taxonomy_issues),
                ("RLS Exposure", rls_summary_df),
                ("Near-Duplicate Measures", duplicates_df.drop(columns=["DAX A", "DAX B"],
                                                               errors="ignore")),
                ("Field Parameters", field_params_df),
                ("Broken Visuals", broken_df),
                ("Column Disposition", disposition_df),
                ("Source Lineage", build_sql_lineage(model)),
            ]
            if not ai_descriptions_df.empty:
                extra_sheets.append(("AI-Drafted Descriptions", ai_descriptions_df))
            excel_bytes = build_data_dictionary_excel(
                model, health_df, naming_df, unused_df, extra_sheets=extra_sheets,
            )
            st.download_button(
                "⬇ Download Data Dictionary (.xlsx)", excel_bytes,
                file_name=f"{_slug(model.get('model_name') or 'model')}_data_dictionary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_data_dictionary", width="stretch",
            )

# --- Theme -------------------------------------------------------------------
if nav_page == "Theme":
    with tab_guard('Theme'):
        st.subheader("Report theme")
        st.caption(
            "A .vpax has no styling metadata at all — it's the semantic model only. This reads "
            "the actual Power BI theme file out of the .pbix (colors, fonts, and sizes for "
            "titles/headers/labels/callouts) and lets you download it as-is, so it can be "
            "re-imported into another report via **View ➜ Themes ➜ Browse for themes**."
        )

        uploaded_theme_file = st.file_uploader(
            "Upload a theme.json instead — optional", type=["json"], key="theme_upload",
            help="Use this if the theme read from your .pbix above looks wrong or came back "
                 "empty, or to preview/re-download a theme from a different, known-good report "
                 "without re-uploading that whole .pbix.",
        )

        theme: Optional[Dict[str, Any]] = None
        if uploaded_theme_file is not None:
            raw = uploaded_theme_file.getvalue()
            try:
                tj = json.loads(raw.decode("utf-8-sig"))
                if not isinstance(tj, dict):
                    raise ValueError("Top-level JSON must be an object.")
                theme = {
                    "found": True, "is_custom": True, "raw_bytes": raw, "json": tj,
                    "name": tj.get("name") or _slug(uploaded_theme_file.name.rsplit(".", 1)[0], "theme"),
                }
                st.success(f"Using the uploaded theme file instead of the .pbix's own theme.", icon="✅")
            except (json.JSONDecodeError, UnicodeDecodeError, ValueError) as exc:
                st.error(f"That file isn't valid theme JSON: {exc}")
        elif report is not None:
            theme = report.get("theme") or {"found": False}

        if theme is None:
            st.info(
                "👈 Upload the matching **.pbix** in the sidebar (*Add report pages*) to read its "
                "theme, or upload a **theme.json** file directly above."
            )
        elif not theme.get("found"):
            name = theme.get("name")
            if name:
                st.warning(
                    f"This report references a theme named **{name}**, but its JSON file "
                    "couldn't be located inside the .pbix. Try uploading the theme.json directly "
                    "above if you have it."
                )
            else:
                st.info(
                    "No explicit theme override found — this report uses Power BI's built-in "
                    "default theme, which isn't stored as a file inside the .pbix."
                )
        else:
            tj = theme["json"]
            kind = "custom theme uploaded into this report" if theme["is_custom"] else "built-in Power BI theme"
            st.markdown(f"**{theme['name']}** — {kind}")

            data_colors = tj.get("dataColors") or []
            if data_colors:
                swatches = "".join(
                    f'<span style="display:inline-block;width:22px;height:22px;'
                    f'background:{html.escape(str(c))};border:1px solid #cbd5e1;'
                    f'border-radius:4px;margin:2px" title="{html.escape(str(c))}"></span>'
                    for c in data_colors
                )
                st.markdown(swatches, unsafe_allow_html=True)

            info_cols = st.columns(3)
            info_cols[0].metric("Background", tj.get("background", "—"))
            info_cols[1].metric("Foreground (text)", tj.get("foreground", "—"))
            info_cols[2].metric("Table/visual accent", tj.get("tableAccent", "—"))

            text_classes = tj.get("textClasses") or {}
            if text_classes:
                st.markdown("**Text styles**")
                for role in ("header", "title", "label", "callout"):
                    tc = text_classes.get(role)
                    if isinstance(tc, dict):
                        st.caption(
                            f"**{role.capitalize()}** — {tc.get('fontFace', '—')}, "
                            f"{tc.get('fontSize', '—')}pt, {tc.get('color', '—')}"
                        )

            st.download_button(
                "⬇ Download theme (.json)", theme["raw_bytes"],
                file_name=f"{_slug(theme['name'], 'theme')}.json",
                mime="application/json", key="dl_theme", width="stretch",
            )

# --- Display Folders ----------------------------------------------------------
if nav_page == "Display Folders":
    with tab_guard("Display Folders"):
        st.subheader("Display folder taxonomy")
        st.caption(
            "How the model looks in the **Fields** pane of Power BI Desktop. This is the "
            "only part of a semantic model a business user ever navigates directly, so a "
            "sloppy taxonomy costs more support time than a slow measure does."
        )
        if not taxonomy_issues.empty:
            counts = severity_counts(taxonomy_issues)
            st.warning(
                f"{len(taxonomy_issues)} taxonomy issue(s) — "
                + ", ".join(f"{n} {s.lower()}" for s, n in counts.items()),
                icon="🗂️",
            )
            show_table(taxonomy_issues, "Taxonomy Issues", height=240, key="tax_issues", row_height=110)
        else:
            st.success("No orphaned measures or visible join keys found.", icon="✅")

        show_hidden = st.checkbox(
            "Show hidden objects", value=False, key="tax_hidden",
            help="Hidden columns don't appear in the Fields pane — off by default so this "
                 "view matches what a report author actually sees.",
        )
        st.markdown("#### Folder tree")
        if not taxonomy_tree:
            st.info("No measures or columns to organise.")
        for table in sorted(taxonomy_tree):
            folders = taxonomy_tree[table]
            visible_total = sum(
                1 for items in folders.values() for i in items if show_hidden or not i["hidden"]
            )
            if not visible_total:
                continue
            with st.expander(f"📁 {table}  ·  {visible_total} object(s)"):
                for folder in sorted(folders, key=lambda f: (f == "", f.lower())):
                    items = [i for i in folders[folder] if show_hidden or not i["hidden"]]
                    if not items:
                        continue
                    if folder:
                        st.markdown(f"**📂 {html.escape(folder)}**")
                    else:
                        st.markdown("**(no display folder)**")
                    lines = []
                    for i in sorted(items, key=lambda x: (x["kind"], x["name"].lower())):
                        icon = "📐" if i["kind"] == "measure" else "🧱"
                        suffix = " *(hidden)*" if i["hidden"] else ""
                        lines.append(f"{icon} {i['name']}{suffix}")
                    st.markdown(
                        "<div style='columns:3;font-size:.85rem;color:#475569'>"
                        + "".join(f"<div>{html.escape(l)}</div>" for l in lines)
                        + "</div>",
                        unsafe_allow_html=True,
                    )

# --- Compression Advisor ------------------------------------------------------
if nav_page == "Compression Advisor":
    with tab_guard("Compression Advisor"):
        st.subheader("VertiPaq compression & encoding advisor")
        st.caption(
            "Model Size tells you *where* the memory went; this tells you *what to do about "
            "it*. Every row is a column big enough to matter, with the specific change that "
            "would shrink it. Distinct counts are estimated from VertiPaq's dictionary bit "
            "width — a .vpax carries no row data, so treat them as an order of magnitude."
        )
        if encoding_df.empty:
            st.success(
                "No storage problems found on the columns large enough to be worth tuning — "
                "or this export has no VertiPaq statistics (Tabular Editor exports often "
                "omit them; use DAX Studio ➜ Export Metadata for the full picture).",
                icon="✅",
            )
        else:
            counts = severity_counts(encoding_df)
            c1, c2, c3 = st.columns(3)
            c1.metric("🔴 High", counts.get("High", 0), help="Large, clearly fixable waste.")
            c2.metric("🟠 Medium", counts.get("Medium", 0))
            c3.metric("Columns advised", len(encoding_df))
            encoding_ai_df = encoding_df.copy()
            encoding_ai_df["Object"] = encoding_ai_df["Table"] + "[" + encoding_ai_df["Column"] + "]"
            encoding_view_df = render_ai_grouped_fixes(
                encoding_ai_df, model, key_prefix="ai_compression",
                group_col="Table", object_col="Object", message_col="Why", group_label="table",
            ).drop(columns=["Object"])
            show_table(encoding_view_df, "Compression Advice", height=440, key="encoding", row_height=130)
            with st.expander("How to act on these", icon=":material/build:"):
                st.markdown(
                    "- **HASH on a numeric column** — set `EncodingHint = Value` in Tabular "
                    "Editor, or fix the source so the column is a clean integer. Removes the "
                    "dictionary outright.\n"
                    "- **NOSPLIT with few distinct values** — add an `ORDER BY` on that column "
                    "to the source query (or `Table.Sort` in M) so run-length encoding can "
                    "collapse the runs. Sort by the *lowest*-cardinality column first.\n"
                    "- **Dictionary over 60%** — split the column (date from time, prefix from "
                    "suffix) or round decimals to reporting precision. High-cardinality text "
                    "is almost always the single biggest line in a large model.\n"
                    "- **Hidden but available in MDX** — the *Fix Script (C#)* section "
                    "generates the `IsAvailableInMdx = false` calls for you."
                )

# --- Fabric Readiness ---------------------------------------------------------
if nav_page == "Fabric Readiness":
    with tab_guard("Fabric Readiness"):
        st.subheader("Microsoft Fabric / Direct Lake readiness")
        st.caption(
            "Direct Lake reads Parquet straight out of OneLake, so anything that needs the "
            "import engine to transform data at load time isn't supported — and when one is "
            "present, queries silently fall back to DirectQuery and the speed advantage "
            "disappears. This flags the blockers that a .vpax can actually see."
        )
        counts = severity_counts(fabric_df)
        c1, c2, c3 = st.columns(3)
        c1.metric("🔴 Blockers", counts.get("High", 0), help="Would prevent or break Direct Lake.")
        c2.metric("🟠 Warnings", counts.get("Medium", 0))
        c3.metric("Checks run", len(fabric_df))
        if counts.get("High"):
            st.error(
                "This model would **not** run cleanly in Direct Lake today. Work the high-"
                "severity rows first — calculated tables and columns have to move upstream.",
                icon="🔴",
            )
        elif counts.get("Medium"):
            st.warning("No hard blockers, but some items need attention before migrating.", icon="🟠")
        else:
            st.success("No Direct Lake blockers detected in this export.", icon="✅")
        fabric_view_df = fabric_df
        if not fabric_df.empty:
            fabric_view_df = render_ai_grouped_fixes(
                fabric_df, model, key_prefix="ai_fabric",
                group_col="Check", object_col="Object", message_col="Finding", group_label="check",
            )
        show_table(fabric_view_df, "Fabric Readiness", height=420, key="fabric", row_height=130)
        st.info(
            "**What this can't see:** whether the source Delta tables are V-Order optimised, "
            "the size and count of Parquet row groups, or the capacity SKU you're targeting. "
            "Confirm those in the Fabric portal before committing to a migration.",
            icon="ℹ️",
        )

# --- Model Compare ------------------------------------------------------------
if nav_page == "Model Compare":
    with tab_guard("Model Compare"):
        st.subheader("Model compare & metric drift")
        st.caption(
            "The failure mode this catches: someone connects to the certified model, adds "
            "local measures in a composite model, and one of them reuses a certified "
            "measure's name with different DAX. The name still matches the single source of "
            "truth — the number no longer does, and nothing in Power BI warns anyone."
        )
        if baseline_model is None:
            st.info(
                "👈 Upload the **certified/baseline .vpax** under *Compare against a baseline* "
                "in the sidebar. The model you already loaded is treated as the working copy.",
                icon="🆚",
            )
        else:
            diff = compare_models(baseline_model, model)
            drift, added, removed = diff["drift"], diff["added"], diff["removed"]
            b_name = baseline_model.get("model_name") or "baseline"
            st.markdown(
                f"Comparing **{html.escape(str(model.get('model_name') or 'this model'))}** "
                f"against baseline **{html.escape(str(b_name))}**."
            )
            k = st.columns(4)
            k[0].metric("🔴 Drifted measures", int((drift["Change"] == "DAX changed").sum()) if not drift.empty else 0,
                        help="Same name, different DAX — the single source of truth is broken.")
            k[1].metric("Missing", len(removed), help="In the baseline, absent here.")
            k[2].metric("Extra", len(added), help="Local measures not in the certified set.")
            k[3].metric("Structure changes", len(diff["structure"]))

            render_ai_compare_changelog(
                drift, added, removed, diff["structure"], b_name, model, key_prefix="ai_compare",
            )

            if drift.empty and added.empty and removed.empty:
                st.success("Every measure matches the baseline, name and DAX.", icon="✅")

            if not drift.empty:
                st.markdown("#### 🔴 Metric drift")
                show_table(
                    drift[["Measure", "Change", "Severity", "Baseline Table", "Compared Table",
                           "Why It Matters", "How to Fix"]],
                    "Metric Drift", height=260, key="drift", row_height=120,
                )
                pick = st.selectbox("Inspect the DAX for", drift["Measure"].tolist(), key="drift_pick")
                row = drift[drift["Measure"] == pick].iloc[0]
                d1, d2 = st.columns(2)
                with d1:
                    st.markdown(f"**Baseline** — `{row['Baseline Table']}`")
                    st.code(row["Baseline DAX"] or "(empty)", language="dax")
                with d2:
                    st.markdown(f"**This model** — `{row['Compared Table']}`")
                    st.code(row["Compared DAX"] or "(empty)", language="dax")
                st.caption(
                    "Comments and whitespace are ignored when deciding whether DAX changed, "
                    "so only real logic differences appear here."
                )

            if not removed.empty:
                st.markdown("#### Missing from this model")
                show_table(removed[["Measure", "Table", "Severity", "Note", "How to Fix"]],
                           "Missing Measures", height=220, key="cmp_removed", row_height=100)
            if not added.empty:
                st.markdown("#### Only in this model")
                show_table(added[["Measure", "Table", "Severity", "Note", "How to Fix"]],
                           "Extra Measures", height=220, key="cmp_added", row_height=100)
            if not diff["structure"].empty:
                with st.expander(f"Structural differences ({len(diff['structure'])})"):
                    show_table(diff["structure"], "Structure Diff", height=320, key="cmp_struct")

# --- RLS Simulator ------------------------------------------------------------
if nav_page == "RLS Simulator":
    with tab_guard("RLS Simulator"):
        st.subheader("Row-level security propagation simulator")
        st.caption(
            "A role's filter travels the same path any filter does: down the one side to the "
            "many side, and back up only across a bi-directional relationship. So securing a "
            "dimension protects the facts hanging off it — but securing a fact protects "
            "nothing else, and a filter that needs to travel many→one across a single-"
            "direction relationship is trapped. That's the silent failure: the role looks "
            "configured and the data isn't secured."
        )
        if rls_summary_df.empty:
            st.info(
                "No RLS roles are defined in this model — every user who can open the report "
                "sees every row. That may be correct; it's worth confirming it's deliberate."
            )
        else:
            worst = severity_counts(rls_summary_df)
            if worst.get("High"):
                st.error(
                    f"{worst['High']} role(s) leave related tables unsecured because the "
                    "filter can't cross a single-direction relationship.", icon="🔴",
                )
            st.markdown("#### Per-role verdict")
            show_table(rls_summary_df, "RLS Summary", height=240, key="rls_summary", row_height=90)

            rls_sim_view_df = render_ai_grouped_fixes(
                rls_sim_df, model, key_prefix="ai_rls",
                group_col="Role", object_col="Table", message_col="Path / Reason", group_label="role",
                out_col="AI How to Fix (role-specific)",
            )

            roles_list = sorted(rls_sim_view_df["Role"].dropna().unique().tolist())
            picked_role = st.selectbox("Trace filter propagation for", roles_list, key="rls_role")
            trace = rls_sim_view_df[rls_sim_view_df["Role"] == picked_role]
            only_problems = st.checkbox(
                "Show only unsecured related tables", value=True, key="rls_only_problems",
                help="Tables with no relationship path to the secured table are expected to be "
                     "unfiltered — hiding them keeps the real gaps visible.",
            )
            view = trace[trace["Severity"] == "High"] if only_problems else trace
            if view.empty:
                st.success(
                    "No trapped filters for this role — every related table is reached.",
                    icon="✅",
                )
            else:
                show_table(sort_by_severity(view), f"RLS {picked_role}", height=380, key="rls_trace", row_height=120)
            with st.expander("How to fix a trapped filter", icon=":material/build:"):
                st.markdown(
                    "In order of preference:\n\n"
                    "1. **Secure the dimension, not the fact.** Put the filter on the table "
                    "that sits on the *one* side; it then flows down to every fact.\n"
                    "2. **Use a bridge table** for many-to-many security, filtered on the "
                    "user's identity via `USERPRINCIPALNAME()`.\n"
                    "3. **Turn on bi-directional cross-filtering for the security role only** "
                    "(the *Apply security filter in both directions* checkbox) rather than "
                    "making the relationship bi-directional for everyone.\n\n"
                    "Note that `USERELATIONSHIP` cannot be used in a measure that queries a "
                    "table with RLS applied — if a role secures a table your time-intelligence "
                    "measures reactivate relationships on, those measures will error for that role."
                )
            st.info(
                "This simulates the *filter path only*. It can't evaluate whether a role's DAX "
                "expression returns the right rows — that needs **View as role** in Power BI "
                "Desktop against real data.", icon="ℹ️",
            )

# --- Fix Script (C#) ----------------------------------------------------------
if nav_page == "Fix Script (C#)":
    with tab_guard("Fix Script (C#)"):
        st.subheader("Tabular Editor fix script")
        st.caption(
            "Turns the mechanical findings from the audit sections into a runnable C# script "
            "for Tabular Editor's **Advanced Scripting** window. Everything reversible is "
            "emitted live; renames are commented out, because renaming an object breaks every "
            "visual, bookmark and RLS expression that referenced the old name."
        )
        opts = st.columns(2)
        with opts[0]:
            inc_mdx = st.checkbox("Drop attribute hierarchies on hidden columns", True, key="fs_mdx")
            inc_keys = st.checkbox("Hide relationship key columns", True, key="fs_keys")
            inc_fmt = st.checkbox("Set a default format string on unformatted measures", False, key="fs_fmt")
        with opts[1]:
            inc_rename = st.checkbox("Naming-convention renames (commented out)", True, key="fs_rename")
            inc_desc = st.checkbox("Description placeholders (commented out)", False, key="fs_desc")

        script = build_te_script(
            model, naming_df,
            include_renames=inc_rename, include_mdx=inc_mdx, include_hide_keys=inc_keys,
            include_formats=inc_fmt, include_descriptions=inc_desc,
        )
        st.code(script, language="csharp")
        st.download_button(
            "⬇ Download fix script (.csx)", script.encode("utf-8"),
            file_name=f"{_slug(model.get('model_name') or 'model')}_fixes.csx",
            mime="text/plain", key="dl_te_script", width="stretch",
        )
        st.warning(
            "Run this against a **copy** first, and check *Investigate ➜ Impact Analysis* on "
            "anything you're about to rename. Tabular Editor writes changes straight to the "
            "model — there's no undo once saved.", icon="⚠️",
        )

# --- Report Usage -------------------------------------------------------------
if nav_page == "Report Usage":
    with tab_guard("Report Usage"):
        st.subheader("Broken visuals & true delete safety")
        st.caption(
            "The .vpax says what the model contains; the .pbix says what the report asks "
            "for. Comparing them catches both directions — visuals bound to objects that "
            "no longer exist, and columns that genuinely nothing uses."
        )
        if not has_report:
            st.info(
                "👈 Upload the matching **.pbix** in the sidebar (*Add report pages*). "
                "Without it, the Unused Objects scan can only see DAX — a column dropped "
                "straight onto a chart axis is invisible to it, which is why it says "
                "\"likely unused\" rather than \"safe to delete\".",
                icon="🖥️",
            )
        else:
            k = st.columns(4)
            safe_n = int(disposition_df["Verdict"].astype(str).str.startswith("Safe to delete").sum())
            visual_only = int((disposition_df["Verdict"] == "Keep — used on report visuals only").sum())
            k[0].metric("🔴 Broken bindings", len(broken_df),
                        help="Visuals pointing at a table, column or measure that isn't in the model.")
            k[1].metric("🟠 Safe to delete", safe_n,
                        help="Referenced by no DAX and no visual in this report.")
            k[2].metric("Rescued from 'unused'", visual_only,
                        help="Flagged unused by the DAX scan, but actually used on a visual.")
            k[3].metric("Field bindings read", len(bindings_df))

            st.markdown("#### 🔴 Broken visuals")
            if broken_df.empty:
                st.success("Every field every visual asks for exists in the model.", icon="✅")
            else:
                st.error(
                    f"{len(broken_df)} binding(s) reference objects this model doesn't have. "
                    "These visuals will error or render blank when someone opens the page.",
                    icon="🔴",
                )
                broken_view_df = render_ai_grouped_fixes(
                    broken_df, model, key_prefix="ai_broken",
                    group_col="Table", object_col="Field", message_col="Problem", group_label="table",
                    out_col="AI Likely Cause",
                )
                show_table(broken_view_df, "Broken Visuals", height=320, key="broken_visuals", row_height=110)

            st.markdown("#### Delete safety per column")
            hide_kept = st.checkbox("Show only columns safe to delete", value=True, key="disp_only_safe")
            view = (disposition_df[disposition_df["Verdict"].astype(str).str.startswith("Safe to delete")]
                    if hide_kept else disposition_df)
            if view.empty:
                st.success("Nothing is safe to delete — every column earns its place.", icon="✅")
            else:
                show_table(view, "Column Disposition", height=400, key="disposition", row_height=90)
            st.warning(
                "**One .pbix speaks for one report.** If other reports, paginated reports, "
                "Excel workbooks or XMLA clients connect to this same semantic model, they "
                "can use columns this analysis sees as unused. Verify before deleting.",
                icon="⚠️",
            )

            with st.expander(f"All {len(bindings_df)} field bindings read from the report"):
                show_table(bindings_df, "Report Bindings", height=360, key="bindings")

# --- Duplicate Measures -------------------------------------------------------
if nav_page == "Duplicate Measures":
    with tab_guard("Duplicate Measures"):
        st.subheader("Near-duplicate measures")
        st.caption(
            "Exact-string matching misses how redundancy actually happens: someone copies a "
            "certified measure, adds `+ 0` to force a zero instead of a blank, renames a "
            "variable, and now two metrics disagree in exactly the cases nobody tests. "
            "Comparison strips comments, whitespace, casing and VAR names first, so only "
            "genuine logic differences count."
        )
        thr = st.slider(
            "Similarity threshold (%)", min_value=70, max_value=100, value=85, step=5,
            key="dup_threshold",
            help="Lower catches more candidates but adds false positives. 95%+ is usually "
                 "a copy-paste; 80–90% often means one measure could be built from the other.",
        )
        dupes = findings["duplicates"] if thr == 85 else find_near_duplicate_measures(model, thr)
        if dupes.empty:
            st.success(
                f"No measure pairs are {thr}% or more alike. Nothing to consolidate.", icon="✅",
            )
        else:
            counts = severity_counts(dupes)
            c1, c2, c3 = st.columns(3)
            c1.metric("🔴 Very close (95%+)", counts.get("High", 0))
            c2.metric("🟠 Similar", counts.get("Medium", 0))
            c3.metric("Pairs found", len(dupes))
            dupes_view = render_ai_duplicate_judgments(dupes, model, key_prefix="ai_dupes")
            show_cols = ["Measure A", "Table A", "Measure B", "Table B", "Similarity %", "Severity", "Verdict"]
            show_cols += [c for c in ("AI Recommendation", "AI Reasoning") if c in dupes_view.columns]
            show_table(
                dupes_view[show_cols],
                "Near-Duplicate Measures", height=340, key="dupes", row_height=100,
            )
            labels = [f"{r['Measure A']}  ↔  {r['Measure B']}  ({r['Similarity %']}%)"
                      for _, r in dupes.iterrows()]
            pick = st.selectbox("Compare the DAX side by side", labels, key="dup_pick")
            row = dupes.iloc[labels.index(pick)]
            d1, d2 = st.columns(2)
            with d1:
                st.markdown(f"**{row['Measure A']}** — `{row['Table A']}`")
                st.code(row["DAX A"] or "(empty)", language="dax")
            with d2:
                st.markdown(f"**{row['Measure B']}** — `{row['Table B']}`")
                st.code(row["DAX B"] or "(empty)", language="dax")
            st.caption(
                "Before consolidating, run both names through **Investigate ➜ Impact Analysis** "
                "— the one that looks redundant may be the one every report actually uses."
            )

# --- Field Parameters ---------------------------------------------------------
if nav_page == "Field Parameters":
    with tab_guard("Field Parameters"):
        st.subheader("Field parameter tables")
        st.caption(
            "Field parameters are disconnected calculated tables whose rows are `NAMEOF()` "
            "references to other fields. Because nothing in the model points at them, an "
            "abandoned one is invisible to every other check here — no relationship, no "
            "measure and no calculated column will ever mention it."
        )
        if field_params_df.empty:
            st.info("No field parameter tables found in this model.")
        else:
            counts = severity_counts(field_params_df)
            c1, c2 = st.columns(2)
            c1.metric("Field parameter tables", len(field_params_df))
            c2.metric("🟠 Unused in this report", counts.get("Medium", 0))
            if not has_report:
                st.info(
                    "👈 Upload the matching **.pbix** to check which of these a slicer or "
                    "visual actually binds to. Without it, they can only be listed.",
                    icon="🖥️",
                )
            show_table(field_params_df, "Field Parameters", height=360, key="field_params", row_height=110)
            st.caption(
                "Removing an unused one takes the table, its visible label column, its hidden "
                "field column and its sort-order column with it — **Govern ➜ Model Cleanup** "
                "generates that deletion."
            )

# --- Source Lineage -----------------------------------------------------------
if nav_page == "Source Lineage":
    with tab_guard("Source Lineage"):
        st.subheader("Source-system lineage")
        st.caption(
            "Parses the `SELECT` statements behind each table's Power Query step to work out "
            "which warehouse objects the model actually reads. The point is to answer the "
            "data engineer's question — *if I deprecate this table, what breaks?* — without "
            "anyone opening Power BI Desktop."
        )
        if _sqlglot is None:
            st.warning(
                "`sqlglot` isn't installed, so this falls back to a FROM/JOIN regex — it "
                "will mistake CTE and subquery aliases for real tables. Run "
                "`pip install sqlglot` for accurate parsing.",
                icon="⚠️",
            )
        dialect_label = st.selectbox(
            "SQL dialect", SQL_DIALECTS, key="lineage_dialect",
            help="Auto-detect works for most standard SQL. Pick the specific dialect if your "
                 "source uses vendor syntax (T-SQL square brackets, Snowflake semi-structured "
                 "access, and so on).",
        )
        dialect = None if dialect_label.startswith("(") else dialect_label
        lineage = build_sql_lineage(model, dialect)

        if lineage.empty:
            st.info(
                "No SQL source tables could be resolved. This model's tables may load from "
                "a non-SQL source (files, OData, a dataflow), use native M transforms with "
                "no SELECT, or connect via DirectQuery navigation rather than a query."
            )
        else:
            impact = build_source_impact(lineage)
            k = st.columns(3)
            k[0].metric("Source tables", lineage["Full Source Name"].nunique())
            k[1].metric("Model tables mapped", lineage["Model Table"].nunique())
            k[2].metric("Distinct schemas", lineage["Schema"].replace("", pd.NA).nunique())

            st.markdown("#### If a source table is dropped, these break")
            show_table(impact, "Source Impact", height=280, key="source_impact", row_height=90)

            st.markdown("#### Source ➜ model mapping")
            show_table(lineage, "Source Lineage", height=320, key="lineage")

            with st.expander("Show as a diagram", expanded=False, icon=":material/hub:"):
                static_diagram_panel(build_lineage_dot(lineage), engine="dot", filename="source_lineage")

# --- Model Cleanup ------------------------------------------------------------
if nav_page == "Model Cleanup":
    with tab_guard("Model Cleanup"):
        st.subheader("Clean model export")
        st.caption(
            "Turns the delete-safe findings into something you can actually run, instead of "
            "a list somebody works through by hand. Only objects that no DAX **and** no "
            "visual references are ever included."
        )
        if not has_report:
            st.error(
                "This section needs the matching **.pbix**. Deleting columns on the strength "
                "of a DAX-only scan is how reports get broken — a column used on a chart axis "
                "and nowhere else looks completely unused from the model side.",
                icon="🛑",
            )
        else:
            opt = st.columns(2)
            with opt[0]:
                inc_cols = st.checkbox(
                    "Columns unused in DAX and visuals", True, key="cl_cols",
                )
            with opt[1]:
                inc_fp = st.checkbox(
                    "Field parameter tables with no report binding", True, key="cl_fp",
                )
            plan = build_cleanup_plan(disposition_df, field_params_df, inc_cols, inc_fp)
            n_cols, n_tables = len(plan["columns"]), len(plan["tables"])

            m1, m2, m3 = st.columns(3)
            m1.metric("Columns to remove", n_cols)
            m2.metric("Tables to remove", n_tables)
            saved = 0.0
            cols_meta = _user_facing_columns(model["columns"])
            if n_cols and "TotalSize" in cols_meta.columns:
                targets = set(plan["columns"])
                for _, c in cols_meta.iterrows():
                    if (str(c.get("TableName")), str(c.get("ColumnName"))) in targets:
                        saved += pd.to_numeric(pd.Series([c.get("TotalSize")]),
                                               errors="coerce").fillna(0).iloc[0]
            m3.metric("Memory reclaimed", f"{_bytes_to_mb(saved):.1f} MB" if saved else "—")

            if not n_cols and not n_tables:
                st.success(
                    "Nothing qualifies for removal — every object is referenced by DAX or by "
                    "a visual in this report.", icon="✅",
                )
            else:
                with st.expander(f"Review what would be removed ({n_cols + n_tables} object(s))",
                                 expanded=True):
                    if plan["tables"]:
                        st.markdown("**Tables**")
                        st.write(plan["tables"])
                    if plan["columns"]:
                        st.markdown("**Columns**")
                        show_table(
                            pd.DataFrame(plan["columns"], columns=["Table", "Column"]),
                            "Cleanup Columns", height=260, key="cleanup_cols",
                        )

                fmt = st.segmented_control(
                    "Output format", ["Tabular Editor C#", "TMDL notes", "Cleaned Model.bim"],
                    key="cleanup_fmt", required=True,
                    default="Tabular Editor C#" if "cleanup_fmt" not in st.session_state else None,
                    help="C# runs the deletions directly. TMDL notes tell you which blocks to "
                         "remove from a source-controlled definition. Model.bim is the "
                         "rewritten definition itself.",
                )
                slug = _slug(model.get("model_name") or "model")
                if fmt == "Tabular Editor C#":
                    script = build_cleanup_csharp(model, plan)
                    st.code(script, language="csharp")
                    st.download_button(
                        "⬇ Download cleanup script (.csx)", script.encode("utf-8"),
                        file_name=f"{slug}_cleanup.csx", mime="text/plain",
                        key="dl_cleanup_cs", width="stretch",
                    )
                elif fmt == "TMDL notes":
                    tmdl = build_cleanup_tmdl(plan)
                    st.code(tmdl, language="text")
                    st.download_button(
                        "⬇ Download TMDL notes (.tmdl)", tmdl.encode("utf-8"),
                        file_name=f"{slug}_cleanup.tmdl", mime="text/plain",
                        key="dl_cleanup_tmdl", width="stretch",
                    )
                else:
                    try:
                        cleaned = build_cleaned_bim(model, plan)
                        st.success(
                            f"Rewritten definition is {len(cleaned) / 1024:.0f} KB. "
                            "Relationships that pointed at a removed column have been dropped "
                            "too, so the definition still deploys.", icon="✅",
                        )
                        st.download_button(
                            "⬇ Download cleaned Model.bim", cleaned,
                            file_name=f"{slug}_cleaned_Model.bim", mime="application/json",
                            key="dl_cleaned_bim", width="stretch",
                        )
                        st.caption(
                            "Deploy with Tabular Editor (File ➜ Open ➜ From File, then Model ➜ "
                            "Deploy) or via XMLA. It carries model metadata only — the report "
                            "layer is untouched."
                        )
                    except Exception as exc:  # noqa: BLE001
                        st.error(f"Could not rewrite the model definition: {exc}")

                st.warning(
                    "**Before you run this:** other reports on the same semantic model, "
                    "paginated reports, Analyze-in-Excel and XMLA clients are all invisible "
                    "to a single .pbix. Run against a copy and keep a backup.",
                    icon="⚠️",
                )

# --- Databricks Lakeview Export ------------------------------------------------
if nav_page == "Databricks Lakeview Export":
    with tab_guard("Databricks Lakeview Export"):
        st.subheader("Generate a Databricks Lakeview dashboard (.lvdash.json)")
        st.caption(
            "Reverse-engineered against three real exports of the same dashboard from both "
            "platforms — not against published Databricks documentation, since Lakeview's "
            "schema isn't publicly documented the way Power BI's PBIR format is. Treat the "
            "output as a strong first draft: a table per model table, a relationship graph "
            "from the model's relationships, and (with a .pbix) widgets rebuilt from each "
            "page's visuals. Verify the result against a real Lakeview import — especially "
            "relationship direction and the layout grid, both flagged below as informed "
            "guesses, not confirmed Databricks behaviour."
        )

        lineage_guess = build_sql_lineage(model)
        default_catalog, default_schema = "workspace", "default"
        if not lineage_guess.empty:
            db_mode = lineage_guess["Database"].mode()
            sch_mode = lineage_guess["Schema"].mode()
            if len(db_mode) and db_mode.iloc[0]:
                default_catalog = db_mode.iloc[0]
            if len(sch_mode) and sch_mode.iloc[0]:
                default_schema = sch_mode.iloc[0]

        c1, c2 = st.columns(2)
        catalog = c1.text_input(
            "Unity Catalog catalog", value=default_catalog, key="lv_catalog",
            help="Guessed from the Source Lineage scan where possible.",
        )
        schema = c2.text_input("Schema", value=default_schema, key="lv_schema")

        with st.expander("Table name mapping", expanded=False, icon=":material/swap_horiz:"):
            st.caption(
                "Every model table becomes one Lakeview dataset, sourced from "
                "`catalog.schema.<name below>`. Edit any row where the Databricks table "
                "name won't exactly match the Power BI table name."
            )
            name_map_df = pd.DataFrame({
                "Power BI Table": model["all_table_names"],
                "Databricks Table Name": [_lakeview_slug(t) for t in model["all_table_names"]],
            })
            edited = st.data_editor(
                name_map_df, hide_index=True, width="stretch", key="lv_name_map",
                disabled=["Power BI Table"],
            )
            table_name_map = dict(zip(edited["Power BI Table"], edited["Databricks Table Name"]))

        include_rel = st.checkbox(
            "Include a relationshipGraphs block", value=True, key="lv_include_rel",
            help="Built from this model's relationships. Leave off if you'd rather define "
                 "joins by hand in Databricks.",
        )

        chosen_pages: List[str] = []
        pbix_bytes_for_export = pbix_file.getvalue() if pbix_file is not None else None
        if pbix_bytes_for_export is None:
            st.info(
                "👈 Upload the matching **.pbix** in the sidebar to also generate pages and "
                "widgets. Without it, this produces datasets and relationships only.",
                icon="🖥️",
            )
        else:
            all_pbix_pages = model["screens"]["Screen / Page Name"].tolist() if not model["screens"].empty else []
            report_pages = report["pages"]["Screen / Page Name"].tolist() if report and not report["pages"].empty else []
            page_options = report_pages or all_pbix_pages
            if not page_options:
                st.warning("Couldn't read any page names from this .pbix.", icon="⚠️")
            else:
                st.caption(
                    "Pick the pages worth migrating. Most real reports carry tooltip, "
                    "definition, and helper pages alongside the ones people actually look at — "
                    "there's no reliable way to tell those apart automatically, so nothing is "
                    "pre-selected."
                )
                chosen_pages = st.multiselect(
                    "Pages to include", page_options, key="lv_pages",
                    help="Layout generation currently supports the newer PBIR-format .pbix "
                         "only (Power BI Desktop's 'Power BI project' file format).",
                )

        if st.button("Generate Lakeview JSON", key="lv_generate", width="stretch", icon=":material/settings:"):
            with st.spinner("Translating measures, relationships, and pages…"):
                dashboard, conv_report = build_lakeview_dashboard(
                    model, catalog.strip() or "workspace", schema.strip() or "default",
                    table_name_map, pbix_bytes_for_export, chosen_pages, report,
                    include_relationships=include_rel,
                )
            st.session_state["lv_dashboard"] = dashboard
            st.session_state["lv_report"] = conv_report

        dashboard = st.session_state.get("lv_dashboard")
        conv_report = st.session_state.get("lv_report")
        if dashboard is not None:
            n_datasets = len(dashboard.get("datasets", []))
            n_measures = sum(len(d["config"].get("measures", [])) for d in dashboard.get("datasets", []))
            n_rels = sum(len(g.get("relationships", [])) for g in dashboard.get("relationshipGraphs", []))
            n_widgets = sum(len(p.get("layout", [])) for p in dashboard.get("pages", []))

            k = st.columns(4)
            k[0].metric("Datasets", n_datasets)
            k[1].metric("Measures included", n_measures)
            k[2].metric("Relationships", n_rels)
            k[3].metric("Widgets generated", n_widgets)

            if conv_report is not None and not conv_report.empty:
                counts = severity_counts(conv_report)
                translated = int((conv_report["Status"] == "Translated").sum())
                not_translated = int((conv_report["Status"] == "Not translated").sum())
                if not_translated:
                    st.warning(
                        f"{not_translated} item(s) need manual work in Databricks — see the "
                        "conversion report below. Everything else generated cleanly.",
                        icon="🟠",
                    )
                else:
                    st.success("Every measure, relationship, and visual translated cleanly.", icon="✅")
                st.markdown("#### Conversion report")
                st.caption(
                    "One row per object this generator had an opinion about. `Info`/`Translated` "
                    "rows are informational; everything else needs a look before you trust the "
                    "output."
                )
                show_table(conv_report, "Lakeview Conversion Report", height=360, key="lv_report_table")

            payload = json.dumps(dashboard, indent=2).encode("utf-8")
            with st.expander("Preview generated JSON", expanded=False, icon=":material/data_object:"):
                st.code(json.dumps(dashboard, indent=2)[:20000], language="json")
                if len(payload) > 20000:
                    st.caption(f"Preview truncated — full file is {len(payload) / 1024:.0f} KB.")
            st.download_button(
                "⬇ Download .lvdash.json", payload,
                file_name=f"{_slug(model.get('model_name') or 'dashboard')}.lvdash.json",
                mime="application/json", key="dl_lvdash", width="stretch",
            )
            st.info(
                "Import via Databricks: Dashboards ➜ Create ➜ Import from file, or the "
                "Lakeview REST API. Confirm the catalog/schema/table names above match what's "
                "actually registered in Unity Catalog before importing.",
                icon="ℹ️",
            )

if nav_page == "TWBX → Power BI":
    with tab_guard("TWBX → Power BI"):
        _twbx_lineage_guess = build_sql_lineage(model)
        _twbx_default_catalog, _twbx_default_schema = "workspace", "default"
        if not _twbx_lineage_guess.empty:
            _db_mode = _twbx_lineage_guess["Database"].mode()
            _sch_mode = _twbx_lineage_guess["Schema"].mode()
            if len(_db_mode) and _db_mode.iloc[0]:
                _twbx_default_catalog = _db_mode.iloc[0]
            if len(_sch_mode) and _sch_mode.iloc[0]:
                _twbx_default_schema = _sch_mode.iloc[0]
        render_twbx_conversion_page(
            default_catalog=_twbx_default_catalog, default_schema=_twbx_default_schema,
            call_llm=call_llm if _llm_ready() else None, llm_ready=_llm_ready(),
        )

st.markdown(
    f'<div class="app-footer">🧩 VPAX Semantic Model Explorer &nbsp;·&nbsp; <b>{AUTHOR}</b></div>',
    unsafe_allow_html=True,
)
