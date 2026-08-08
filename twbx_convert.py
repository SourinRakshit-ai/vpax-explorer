"""twbx_convert.py — Tableau (.twbx) -> Power BI Project converter, Tier 1.

Standalone module for the Vpax Explorer Streamlit app. Converts a Tableau
workbook's structural metadata (schema, relationships, calculated fields,
worksheets, dashboards, RLS group filters) into a TMDL + PBIR bundle the
user opens as a Power BI Project in Desktop and saves as .pbix from there —
there is no supported way to write a .pbix file directly (same constraint
already documented on the Databricks Lakeview export page).

Scope is Tier 1 only. Nested LOD, table calculations, custom-SQL
datasources, range/date parameters, non-filter dashboard actions, and
unsupported mark classes are refused with a specific reason in the
conversion report — never silently approximated. See the accompanying build
spec (Accelerator Work/TWBX_to_PBIX_Build_Instructions.md) for the full
scope rationale.

This module has no dependency on app.py — app.py imports two entry points
from here and passes in whatever cross-cutting helpers it needs (like an
LLM caller), rather than this module importing back from app.py, which
would create an import cycle (app.py imports this module at its own
top level).
"""

from __future__ import annotations

import hashlib
import html
import io
import json
import re
import zipfile
from typing import Any, Callable, Dict, List, Optional, Set, Tuple
import xml.etree.ElementTree as ET

import pandas as pd
import streamlit as st


# ==========================================================================
# Small shared helpers
# ==========================================================================
# Reimplemented here (not imported from app.py) rather than shared, per the
# build spec's "don't touch existing app.py functions, additions only"
# constraint — importing app.py from here (while app.py imports this module)
# would be a circular import, and app.py's own top-level code runs real
# Streamlit calls (st.set_page_config) that can't safely execute twice.

def det_id(*parts: str, length: int = 8) -> str:
    """A short, deterministic hex id — same pattern as app.py's
    `_lakeview_id`: hashed from the object's own identity so re-running the
    converter on an unchanged workbook produces byte-identical output."""
    h = hashlib.md5("::".join(parts).encode("utf-8")).hexdigest()
    return h[:length]


def slug(name: Any, fallback: str = "item") -> str:
    s = re.sub(r"\W+", "_", str(name)).strip("_")
    return s[:80] or fallback


def pbi_name(name: Any, fallback: str = "Item") -> str:
    """A displayable TMDL/PBIR object name — collapse whitespace, strip
    control characters. TMDL quotes names containing spaces at emission
    time (see tmdl_ident), so spaces themselves are fine to keep."""
    s = re.sub(r"[\r\n\t]", " ", str(name or "")).strip()
    s = re.sub(r"\s+", " ", s)
    return s or fallback


# ==========================================================================
# Stage 1 — parse_twbx(): structural extraction
# ==========================================================================
# Real XML parser only (xml.etree.ElementTree), never regex, per the build
# spec's constraint #4 — attribute quoting/escaping/nesting edge cases need
# a real parser. ElementTree already resolves standard XML entity escapes
# (&quot; &apos; &amp; &lt; &gt; and numeric refs like &#13;&#10;) as part of
# parsing every attribute value and text node, confirmed directly against
# the real sample files (a raw `&quot;WALMART L4&quot;` in the file's bytes
# comes back from `.get()` as `"WALMART L4"` with real quote characters
# already). html.unescape() is still applied once more before tokenizing
# formulas, per the build spec, as cheap defense-in-depth — it's a no-op on
# already-clean text.

_ISMEMBEROF_RE = re.compile(r"ISMEMBEROF\(\s*'((?:[^'\\]|\\.)*)'\s*\)")


def _local_tag(tag: str) -> str:
    """Strip an XML namespace prefix like '{...}group' down to 'group'."""
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _unescape(s: Optional[str]) -> str:
    return html.unescape(s) if s else ""


def _bracket_strip(s: Optional[str]) -> str:
    """`[Field Name]` -> `Field Name`. A no-op if there are no brackets."""
    s = (s or "").strip()
    if s.startswith("[") and s.endswith("]"):
        return s[1:-1]
    return s


def _parse_relation_tree(rel: Optional[ET.Element]) -> Optional[Dict[str, Any]]:
    """One `<relation>` node (table / join / text) -> a small recursive
    dict. Join trees can nest (a join's own children can themselves be
    joins), so this recurses rather than assuming exactly two leaf children."""
    if rel is None:
        return None
    rtype = rel.get("type")
    if rtype == "table":
        return {
            "kind": "table",
            "name": _bracket_strip(rel.get("name")),
            "physical_table": _bracket_strip(rel.get("table")),
        }
    if rtype == "text":
        return {"kind": "text", "name": _bracket_strip(rel.get("name")), "sql": (rel.text or "").strip()}
    if rtype == "join":
        clauses = []
        for clause in rel.findall("clause"):
            expr = clause.find("expression")
            # A join clause's expression is itself a binary-op tree of
            # <expression op='='><expression op='[a].[b]'/><expression op='[c].[d]'/></expression>.
            # Two-level flatten covers every real join clause in the corpus
            # (deep dive found no nested-boolean join predicates); anything
            # deeper is reported as an unresolved clause, not guessed at.
            left = right = None
            if expr is not None:
                kids = list(expr)
                if len(kids) == 2:
                    left, right = kids[0].get("op"), kids[1].get("op")
            clauses.append({"op": expr.get("op") if expr is not None else None, "left": left, "right": right})
        children = [c for c in rel.findall("relation")]
        return {
            "kind": "join",
            "join_type": rel.get("join"),
            "clauses": clauses,
            "children": [_parse_relation_tree(c) for c in children],
        }
    return {"kind": "unknown", "raw_type": rtype}


def _leaf_tables(tree: Optional[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Every `table`-kind leaf in a (possibly join-nested) relation tree."""
    if tree is None:
        return []
    if tree["kind"] == "table":
        return [tree]
    if tree["kind"] == "join":
        out = []
        for c in tree["children"]:
            out.extend(_leaf_tables(c))
        return out
    return []


def _parse_metadata_records(ds: ET.Element) -> Dict[str, Dict[str, Any]]:
    """{local column name (bracket-stripped) -> {remote_name, local_type, parent_name}}.

    This is Tableau's own resolved schema cache — preferred over
    <relation><columns> per the build spec, since it carries the resolved
    local type rather than the raw remote guess.
    """
    out: Dict[str, Dict[str, Any]] = {}
    for mr in ds.findall(".//metadata-records/metadata-record"):
        if mr.get("class") != "column":
            continue
        local_name = _bracket_strip(mr.findtext("local-name"))
        if not local_name:
            continue
        out[local_name] = {
            "remote_name": mr.findtext("remote-name") or "",
            "local_type": mr.findtext("local-type") or "",
            "parent_name": _bracket_strip(mr.findtext("parent-name")),
        }
    return out


def _parse_group_rls(ds: ET.Element) -> List[Dict[str, Any]]:
    """RLS candidates from this datasource's <group> elements.

    Real-file correction vs. the build spec's first-pass description:
    ISMEMBEROF() filters do NOT live inside a worksheet <filter
    class='categorical'> block. They live inside a datasource-level
    <group name='[...]'> element's nested <groupfilter> tree — verified by
    walking the actual parent chain in the sample corpus, not assumed from
    the spec. One <group> commonly wraps 20-30 ISMEMBEROF blocks (one per
    AD/SSO security group), each paired with the dimension member values
    that group may see, via nested <groupfilter function='member'
    level='[field]' member='"value"'/> children (or `function='empty-level'`
    for a group with zero allowed members on this field).

    A `<groupfilter expression='false' function='filter'>` sibling — a
    permanently-dead branch left behind by Tableau's group-editor UI — is
    skipped; `expression='false'` never evaluates true.
    """
    candidates: List[Dict[str, Any]] = []

    def walk(node: ET.Element, group_name: str) -> None:
        expr = node.get("expression")
        if expr and expr != "false":
            m = _ISMEMBEROF_RE.search(expr)
            if m:
                ad_group = m.group(1).replace("\\.", ".").replace("\\\\", "\\")
                members: List[Tuple[str, str]] = []
                level_field = None
                for mf in node.iter():
                    if _local_tag(mf.tag) != "groupfilter" or mf is node:
                        continue
                    fn = mf.get("function")
                    if fn == "member":
                        lvl = _bracket_strip(mf.get("level"))
                        val = mf.get("member") or ""
                        val = val[1:-1] if val.startswith('"') and val.endswith('"') else val
                        members.append((lvl, val))
                        level_field = level_field or lvl
                    elif fn == "empty-level":
                        level_field = level_field or _bracket_strip(mf.get("member"))
                candidates.append({
                    "group_name": group_name,
                    "ad_group": ad_group,
                    "level_field": level_field or "",
                    "allowed_values": sorted({v for _, v in members}),
                })
                return  # don't also recurse into an already-matched ISMEMBEROF block
        for child in node:
            if _local_tag(child.tag) == "groupfilter":
                walk(child, group_name)

    for grp in ds.findall(".//group"):
        name = _bracket_strip(grp.get("name"))
        if not name:
            continue
        for gf in grp:
            if _local_tag(gf.tag) == "groupfilter":
                walk(gf, name)
    return candidates


def _parse_datasource(ds: ET.Element) -> Dict[str, Any]:
    name = ds.get("name") or ""
    caption = ds.get("caption") or name
    conn = ds.find("connection")
    conn_class = conn.get("class") if conn is not None else None
    top_relation = conn.find("relation") if conn is not None else None
    relation_tree = _parse_relation_tree(top_relation)
    has_custom_sql = any(
        n.get("type") == "text" for n in (conn.iter("relation") if conn is not None else [])
    )

    schema = _parse_metadata_records(ds)

    logical: Dict[str, Dict[str, Any]] = {}
    for col in ds.findall("column"):
        col_name = _bracket_strip(col.get("name"))
        if not col_name:
            continue
        calc = col.find("calculation")
        formula = None
        if calc is not None and calc.get("class") == "tableau":
            formula = _unescape(calc.get("formula"))
        logical[col_name] = {
            "caption": col.get("caption") or col_name,
            "role": col.get("role") or "",
            "datatype": col.get("datatype") or "",
            "default_agg": col.get("aggregation") or "",
            "formula": formula,
            "hidden": (col.get("hidden") or "false").lower() == "true",
        }

    return {
        "name": name,
        "caption": caption,
        "connection_class": conn_class,
        "has_custom_sql": has_custom_sql,
        "relation": relation_tree,
        "schema": schema,
        "logical_columns": logical,
        "rls_candidates": _parse_group_rls(ds),
    }


def _parse_datasource_relationships(root: ET.Element) -> List[Dict[str, Any]]:
    out = []
    for dsr in root.findall(".//datasource-relationships/datasource-relationship"):
        mappings = []
        for cm in dsr.findall("column-mapping"):
            for m in cm.findall("map"):
                mappings.append((m.get("key") or "", m.get("value") or ""))
        out.append({
            "source_ds": dsr.get("source") or "",
            "target_ds": dsr.get("target") or "",
            "mappings": mappings,
        })
    return out


def _parse_colref(ref: str) -> Tuple[Optional[str], Optional[str]]:
    """`[datasource].[usr:Calculation_XXX:qk]` or `[ds].[none:field_nm:nk]`
    -> (datasource_name, bare_field_id). The bare field id is what matches a
    datasource's logical <column name='[...]'> (after stripping the
    `namespace:`/`:role` decoration) — verified directly against the real
    files: a shelf/encoding reference to `usr:Calculation_197...:qk`
    resolves to the calc field `<column name='[Calculation_197...]'>` in the
    same datasource.

    Returns (None, None) if `ref` isn't the two-bracket-group shape at all
    (e.g. a parameter reference, which has no datasource prefix) — callers
    treat that as "can't resolve," never guess.
    """
    m = re.match(r"^\[([^\]]+)\]\.\[([^\]]+)\]$", (ref or "").strip())
    if not m:
        return None, None
    ds_name, field_raw = m.group(1), m.group(2)
    parts = field_raw.split(":")
    field_id = parts[1] if len(parts) == 3 else field_raw
    return ds_name, field_id


def _parse_worksheet(ws: ET.Element) -> Dict[str, Any]:
    name = ws.get("name") or ""
    table = ws.find("table")
    encodings: List[Dict[str, str]] = []
    mark_class = None
    if table is not None:
        panes = table.findall("panes/pane")
        for p in panes:
            mark = p.find("mark")
            if mark is not None and mark.get("class"):
                mark_class = mark_class or mark.get("class")
            enc = p.find("encodings")
            if enc is not None:
                for child in enc:
                    channel = _local_tag(child.tag)
                    col = child.get("column")
                    if col:
                        ds_name, field_id = _parse_colref(col)
                        encodings.append({"channel": channel, "ds": ds_name or "", "field": field_id or ""})
        rows_text = table.findtext("rows") or ""
        cols_text = table.findtext("cols") or ""
    else:
        rows_text, cols_text = "", ""

    shelf_refs: List[Dict[str, str]] = []
    for shelf_name, text in (("rows", rows_text), ("cols", cols_text)):
        for ref in re.findall(r"\[[^\[\]]+\]\.\[[^\[\]]+\]", text):
            ds_name, field_id = _parse_colref(ref)
            if ds_name:
                shelf_refs.append({"shelf": shelf_name, "ds": ds_name, "field": field_id or ""})

    filters: List[Dict[str, Any]] = []
    view = table.find("view") if table is not None else None
    if view is not None:
        for filt in view.findall("filter"):
            col = filt.get("column")
            ds_name, field_id = (_parse_colref(col) if col else (None, None))
            filters.append({"class": filt.get("class"), "ds": ds_name or "", "field": field_id or ""})

    return {
        "name": name,
        "mark_class": mark_class,
        "encodings": encodings,
        "shelf_refs": shelf_refs,
        "filters": filters,
    }


def _parse_zone(zone: ET.Element) -> Dict[str, Any]:
    def to_int(v: Optional[str]) -> int:
        try:
            return int(float(v)) if v is not None else 0
        except ValueError:
            return 0

    node = {
        "type": zone.get("type-v2"),
        "x": to_int(zone.get("x")), "y": to_int(zone.get("y")),
        "w": to_int(zone.get("w")), "h": to_int(zone.get("h")),
        "name": zone.get("name"),
        "param": zone.get("param"),
        "children": [_parse_zone(c) for c in zone.findall("zone")],
    }
    return node


def _parse_dashboard(dash: ET.Element) -> Dict[str, Any]:
    size = dash.find("size")
    zones_el = dash.find("zones")
    return {
        "name": dash.get("name") or "",
        "maxwidth": int(size.get("maxwidth", 1300)) if size is not None else 1300,
        "maxheight": int(size.get("maxheight", 800)) if size is not None else 800,
        "sizing_mode": size.get("sizing-mode") if size is not None else "fixed",
        "zones": _parse_zone(zones_el) if zones_el is not None else None,
    }


def _parse_parameters(root: ET.Element) -> Dict[str, Dict[str, Any]]:
    out: Dict[str, Dict[str, Any]] = {}
    top_ds = root.find("datasources")
    if top_ds is None:
        return out
    params_ds = next((d for d in top_ds if d.get("name") == "Parameters"), None)
    if params_ds is None:
        return out
    for col in params_ds.findall("column"):
        name = _bracket_strip(col.get("name"))
        if not name:
            continue
        members = [m.get("value") for m in col.findall("members/member") if m.get("value")]
        aliases = {a.get("key"): a.get("value") for a in col.findall("aliases/alias") if a.get("key")}
        out[name] = {
            "caption": col.get("caption") or name,
            "domain_type": col.get("param-domain-type") or "",
            "current_value": col.get("value"),
            "members": members,
            "aliases": aliases,
        }
    return out


def _parse_actions(root: ET.Element) -> List[Dict[str, Any]]:
    out = []
    for action in root.findall(".//actions/action"):
        cmd = action.find(".//command")
        out.append({
            "name": action.get("caption") or "",
            "command": cmd.get("command") if cmd is not None else None,
        })
    return out


def parse_twbx(file_bytes: bytes) -> Dict[str, Any]:
    """Stage 1: raw structural extraction from a .twbx's bytes.

    Ignores Data/Extracts/*.hyper and Image/* entirely — this direction
    never needs the extract itself (schema is already in plain-text XML,
    and the working assumption is the underlying source stays the same).
    """
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        twb_names = [n for n in z.namelist() if n.lower().endswith(".twb")]
        if not twb_names:
            raise ValueError("No .twb file found inside this .twbx archive.")
        twb_bytes = z.read(twb_names[0])
        workbook_name = twb_names[0].rsplit("/", 1)[-1].rsplit(".", 1)[0]

    root = ET.fromstring(twb_bytes)

    top_ds_container = root.find("datasources")
    datasources: Dict[str, Dict[str, Any]] = {}
    if top_ds_container is not None:
        for ds in top_ds_container:
            if _local_tag(ds.tag) != "datasource" or ds.get("name") == "Parameters":
                continue
            parsed = _parse_datasource(ds)
            datasources[parsed["name"]] = parsed

    worksheets: Dict[str, Dict[str, Any]] = {}
    for ws in root.findall(".//worksheets/worksheet"):
        parsed = _parse_worksheet(ws)
        worksheets[parsed["name"]] = parsed

    dashboards: Dict[str, Dict[str, Any]] = {}
    for dash in root.findall(".//dashboards/dashboard"):
        parsed = _parse_dashboard(dash)
        dashboards[parsed["name"]] = parsed

    return {
        "workbook_name": workbook_name,
        "datasources": datasources,
        "datasource_relationships": _parse_datasource_relationships(root),
        "worksheets": worksheets,
        "dashboards": dashboards,
        "parameters": _parse_parameters(root),
        "actions": _parse_actions(root),
    }


# ==========================================================================
# Formula translation: Tableau formula language -> DAX
# ==========================================================================
# String-level, paren/brace-aware recursive descent — mirrors the coding
# pattern of app.py's translate_dax_measure / _dax_match_call / _dax_split_args
# / _dax_matching_paren (read, not called — different grammar), including its
# hard rule: refuse (return None + a specific reason) rather than emit a
# partial/best-guess translation whenever a construct isn't confidently
# handled. A wrong-but-plausible DAX measure is worse than an honest
# "translate this by hand," because it fails silently in a stakeholder's
# report instead of at review time.
#
# Real corpus grounding (not just the build spec's worked examples): sampled
# ~500 real formulas across all 5 files before writing this. Confirmed here,
# beyond what the spec's own examples show:
#   - Function names are case-insensitive (`sum(`, `Sum(`, `SUM(` all appear).
#   - `//` line comments appear inside formulas, including mid-CASE.
#   - Aggregates commonly wrap a whole IF/CASE expression, not just a bare
#     column (`SUM(if [x]='LW' then [y] end)`) — a translator limited to
#     literal `SUM([x])` pattern-matching would refuse a large fraction of
#     real "simple"-tier formulas, so aggregate arguments are translated
#     recursively, not pattern-matched.
#   - `{FIXED : expr}` (zero dimensions) is valid — an LOD over the whole
#     table.
#   - Parameters are referenced as `[Parameters].[Param Name]`, a distinct
#     shape from a plain `[Field]` or datasource-qualified `[ds].[field]`
#     shelf/encoding reference.

_TABLE_CALC_RE = re.compile(
    r"\b(WINDOW_\w+|RUNNING_\w+|RANK\w*|INDEX|LOOKUP|TOTAL|FIRST|LAST|PREVIOUS_VALUE|SIZE)\s*\(",
    re.I,
)

# Tableau function name -> (DAX function name, arg-translation mode).
# mode 'agg': one arg, translated as a scalar expression, wrapped as an
#   aggregate over the home table (SUMX-style) UNLESS the single argument is
#   itself a bare field reference, in which case a plain SUM/AVERAGE/etc.
#   over the column is emitted (much more idiomatic DAX, and what a human
#   author would write).
_TF_AGGREGATES: Dict[str, str] = {
    "SUM": "SUM", "AVG": "AVERAGE", "COUNT": "COUNT", "COUNTD": "DISTINCTCOUNT",
    "MIN": "MIN", "MAX": "MAX", "MEDIAN": "MEDIAN", "STDEV": "STDEV.S", "VAR": "VAR.S",
}
# Scalar (non-aggregate) function name -> DAX equivalent, all confidently
# 1:1 mappings only — anything with different null/type semantics is left
# out deliberately rather than mapped approximately.
_TF_SCALAR_FUNCS: Dict[str, str] = {
    "ISNULL": "ISBLANK", "TRIM": "TRIM", "UPPER": "UPPER", "LOWER": "LOWER",
    "LEN": "LEN", "LEFT": "LEFT", "RIGHT": "RIGHT", "MID": "MID",
    "ABS": "ABS", "ROUND": "ROUND", "CEILING": "CEILING", "FLOOR": "FLOOR",
    "SQRT": "SQRT", "EXP": "EXP", "LN": "LN", "LOG": "LOG",
    "TODAY": "TODAY", "NOW": "NOW", "YEAR": "YEAR", "MONTH": "MONTH", "DAY": "DAY",
    "REPLACE": "SUBSTITUTE",
    "STARTSWITH": "", "ENDSWITH": "",  # handled specially (arg order differs) — see _tf_call
}


def _tf_strip_comments(s: str) -> str:
    return re.sub(r"//[^\n]*", "", s)


def _tf_matching_close(s: str, open_idx: int, open_ch: str, close_ch: str) -> int:
    depth, i, in_str, str_ch = 0, open_idx, False, ""
    while i < len(s):
        c = s[i]
        if in_str:
            if c == str_ch:
                in_str = False
        elif c in "'\"":
            in_str, str_ch = True, c
        elif c == open_ch:
            depth += 1
        elif c == close_ch:
            depth -= 1
            if depth == 0:
                return i
        i += 1
    return -1


def _tf_split_top(s: str, sep: str = ",") -> List[str]:
    """Top-level `sep`-separated parts, respecting (), {}, [], and quotes."""
    parts, depth, buf, in_str, str_ch = [], 0, [], False, ""
    for c in s:
        if in_str:
            buf.append(c)
            if c == str_ch:
                in_str = False
            continue
        if c in "'\"":
            in_str, str_ch = True, c
            buf.append(c)
            continue
        if c in "([{":
            depth += 1
        elif c in ")]}":
            depth -= 1
        if c == sep and depth == 0:
            parts.append("".join(buf))
            buf = []
        else:
            buf.append(c)
    parts.append("".join(buf))
    return [p.strip() for p in parts]


def _tf_unwrap_parens(s: str) -> str:
    s = s.strip()
    while s.startswith("(") and _tf_matching_close(s, 0, "(", ")") == len(s) - 1:
        inner = s[1:-1].strip()
        if not inner:
            break
        s = inner
    return s


_FIELD_RE = re.compile(r"^\[([^\[\]]+)\](?:\.\[([^\[\]]+)\])?$")
_STR_LIT_RE = re.compile(r"^'((?:[^']|'')*)'$|^\"((?:[^\"]|\"\")*)\"$", re.S)
_NUM_LIT_RE = re.compile(r"^-?(\d+\.?\d*|\.\d+)$")


class _Refused(Exception):
    def __init__(self, reason: str):
        self.reason = reason


class TfContext:
    """What the translator needs to resolve field references and know its
    own limits — built once per datasource by build_semantic_model()."""

    def __init__(
        self, table_name: str, field_map: Dict[str, Dict[str, str]],
        param_table: str = "Parameters", group_names: Optional[Set[str]] = None,
    ):
        self.table_name = table_name
        # field_id -> {"kind": "column"|"measure", "name": dax_display_name}
        self.field_map = field_map
        self.param_table = param_table
        # Tableau <group> names (bracket-stripped) — referencing one of
        # these in a formula is the "boolean-calc-field RLS pattern" the
        # build spec calls out as syntactically ambiguous and out of scope;
        # recognizing it here doesn't translate it, just makes the refusal
        # reason specific instead of a generic "can't resolve this field".
        self.group_names = group_names or set()


def _tf_field_dax(field_id: str, ds_prefix: Optional[str], ctx: TfContext) -> str:
    if ds_prefix == "Parameters":
        return f"SELECTEDVALUE('{ctx.param_table}'[{pbi_name(field_id)}])"
    entry = ctx.field_map.get(field_id)
    if entry is None:
        if field_id in ctx.group_names:
            raise _Refused(
                f"References Tableau group `[{field_id}]` — a security/grouping construct "
                "surfaced separately under RLS role candidates, not translated as a formula."
            )
        raise _Refused(f"Unresolvable field reference: `[{field_id}]`")
    if entry["kind"] == "measure":
        return f"[{entry['name']}]"
    return f"'{ctx.table_name}'[{entry['name']}]"


def _tf_translate(expr: str, ctx: TfContext) -> Tuple[str, bool]:
    """(dax_fragment, is_bare_field). Raises _Refused on anything not
    confidently handled — callers never see a partial translation."""
    expr = _tf_unwrap_parens(_tf_strip_comments(expr).strip())
    if not expr:
        raise _Refused("Empty expression.")

    if _TABLE_CALC_RE.search(expr):
        raise _Refused(
            "Table calculations depend on visual-level addressing/partitioning, out of scope for this build."
        )

    if expr.startswith("{") and _tf_matching_close(expr, 0, "{", "}") == len(expr) - 1:
        return _tf_translate_lod(expr, ctx), False

    m = _STR_LIT_RE.match(expr)
    if m:
        inner = (m.group(1) or m.group(2) or "").replace("''", "'").replace('""', '"')
        return '"' + inner.replace('"', '""') + '"', False

    if _NUM_LIT_RE.match(expr):
        return expr, False

    if expr.upper() in ("NULL", "TRUE", "FALSE"):
        return {"NULL": "BLANK()", "TRUE": "TRUE()", "FALSE": "FALSE()"}[expr.upper()], False

    fm = _FIELD_RE.match(expr)
    if fm:
        ds_prefix, field_id = fm.group(1), fm.group(2)
        if field_id is None:
            # bare [Field] (as seen inside a <calculation formula='...'>,
            # which never carries a datasource prefix — that form only
            # appears in shelf/encoding column='' attributes, Section 5).
            return _tf_field_dax(ds_prefix, None, ctx), True
        return _tf_field_dax(field_id, ds_prefix, ctx), True

    up = expr.upper()
    if up.startswith("IF ") or up == "IF" or up.startswith("IF("):
        return _tf_translate_if(expr, ctx), False
    if up.startswith("CASE"):
        return _tf_translate_case(expr, ctx), False

    call_m = re.match(r"^([A-Za-z_][A-Za-z0-9_]*)\s*\(", expr)
    if call_m and _tf_matching_close(expr, call_m.end() - 1, "(", ")") == len(expr) - 1:
        return _tf_translate_call(call_m.group(1), expr[call_m.end():-1], ctx), False

    return _tf_translate_binop(expr, ctx), False


def _tf_translate_binop(expr: str, ctx: TfContext) -> str:
    """Lowest-to-highest precedence: OR, AND, comparison, + -, * /, unary."""
    for tf_op, dax_op in (("OR", "||"), ("AND", "&&")):
        idx = _tf_find_top_op(expr, rf"\b{tf_op}\b", regex=True)
        if idx is not None:
            left, right = expr[:idx[0]], expr[idx[1]:]
            l_dax, _ = _tf_translate(left, ctx)
            r_dax, _ = _tf_translate(right, ctx)
            return f"({l_dax} {dax_op} {r_dax})"

    idx = _tf_find_top_op(expr, r"\bIN\s*\(", regex=True)
    if idx is not None and expr.rstrip().endswith(")"):
        left = expr[:idx[0]]
        paren_start = idx[1] - 1
        if left.strip() and _tf_matching_close(expr, paren_start, "(", ")") == len(expr.rstrip()) - 1:
            l_dax, _ = _tf_translate(left, ctx)
            items = _tf_split_top(expr[paren_start + 1:len(expr.rstrip()) - 1])
            item_dax = []
            for it in items:
                d, _ = _tf_translate(it, ctx)
                item_dax.append(d)
            return f"({l_dax} IN {{{', '.join(item_dax)}}})"

    idx = _tf_find_top_op(expr, r"<>|<=|>=|=|<|>")
    if idx is not None:
        op = expr[idx[0]:idx[1]]
        dax_op = "<>" if op == "<>" else op
        left, right = expr[:idx[0]], expr[idx[1]:]
        l_dax, _ = _tf_translate(left, ctx)
        r_dax, _ = _tf_translate(right, ctx)
        return f"({l_dax} {dax_op} {r_dax})"

    idx = _tf_find_top_op(expr, r"[+\-]")
    if idx is not None:
        left, right = expr[:idx[0]], expr[idx[1]:]
        if left.strip():  # not a unary +/- at the start
            l_dax, _ = _tf_translate(left, ctx)
            r_dax, _ = _tf_translate(right, ctx)
            return f"({l_dax} {expr[idx[0]:idx[1]]} {r_dax})"

    idx = _tf_find_top_op(expr, r"[*/%]")
    if idx is not None:
        left, right = expr[:idx[0]], expr[idx[1]:]
        if left.strip():
            l_dax, _ = _tf_translate(left, ctx)
            r_dax, _ = _tf_translate(right, ctx)
            op = "MOD" if expr[idx[0]:idx[1]] == "%" else expr[idx[0]:idx[1]]
            return f"({l_dax} {op} {r_dax})" if op != "MOD" else f"MOD({l_dax}, {r_dax})"

    if expr.upper().startswith("NOT "):
        inner, _ = _tf_translate(expr[4:], ctx)
        return f"NOT({inner})"
    if expr.startswith("-"):
        inner, _ = _tf_translate(expr[1:], ctx)
        return f"(-{inner})"

    raise _Refused(f"Unparseable expression shape: `{expr[:70]}`")


def _tf_find_top_op(expr: str, pattern: str, regex: bool = False) -> Optional[Tuple[int, int]]:
    """The LAST top-level (paren/brace/bracket/quote-depth-0) match of
    `pattern`, scanned right-to-left so left-associative chains
    (a - b - c) split at the rightmost operator first, matching normal
    left-to-right evaluation order once recursion unwinds."""
    depth, in_str, str_ch = 0, False, ""
    spans: List[Tuple[int, int]] = []
    rx = re.compile(pattern, re.I)
    i = 0
    while i < len(expr):
        c = expr[i]
        if in_str:
            if c == str_ch:
                in_str = False
            i += 1
            continue
        if c in "'\"":
            in_str, str_ch = True, c
            i += 1
            continue
        if c in "([{":
            depth += 1
        elif c in ")]}":
            depth -= 1
        elif depth == 0:
            m = rx.match(expr, i)
            if m and m.end() > m.start():
                spans.append((m.start(), m.end()))
                i = m.end()
                continue
        i += 1
    return spans[-1] if spans else None


def _tf_translate_call(func: str, args_str: str, ctx: TfContext) -> str:
    fn = func.upper()
    args = _tf_split_top(args_str) if args_str.strip() else []

    if fn in ("MAX", "MIN") and len(args) >= 2:
        # Tableau's MAX/MIN(a, b, ...) two-or-more-argument form is a plain
        # scalar comparison, not an aggregation — DAX's MAX()/MIN() support
        # the identical multi-argument scalar overload directly.
        dax_args = [_tf_translate(a, ctx)[0] for a in args]
        return f"{fn}({', '.join(dax_args)})"

    if fn in _TF_AGGREGATES:
        if len(args) != 1:
            raise _Refused(f"{fn}() with other than one argument.")
        dax_fn = _TF_AGGREGATES[fn]
        inner_dax, is_bare = _tf_translate(args[0], ctx)
        if is_bare:
            return f"{dax_fn}({inner_dax})"
        agg_x = {"SUM": "SUMX", "AVERAGE": "AVERAGEX", "MIN": "MINX", "MAX": "MAXX",
                 "MEDIAN": "MEDIANX", "COUNT": "COUNTX", "DISTINCTCOUNT": "COUNTX"}.get(dax_fn)
        if agg_x is None:
            raise _Refused(f"{fn}() over a non-column expression has no confident DAX iterator equivalent.")
        return f"{agg_x}('{ctx.table_name}', {inner_dax})"

    if fn == "ZN":
        if len(args) != 1:
            raise _Refused("ZN() with other than one argument.")
        inner, _ = _tf_translate(args[0], ctx)
        return f"COALESCE({inner}, 0)"
    if fn == "IFNULL":
        if len(args) != 2:
            raise _Refused("IFNULL() with other than two arguments.")
        a, _ = _tf_translate(args[0], ctx)
        b, _ = _tf_translate(args[1], ctx)
        return f"COALESCE({a}, {b})"
    if fn == "CONTAINS":
        if len(args) != 2:
            raise _Refused("CONTAINS() with other than two arguments.")
        a, _ = _tf_translate(args[0], ctx)
        b, _ = _tf_translate(args[1], ctx)
        return f"CONTAINSSTRING({a}, {b})"
    if fn in ("STARTSWITH", "ENDSWITH"):
        raise _Refused(f"{fn}() has no direct DAX equivalent with matching null semantics.")
    if fn == "IIF":
        if len(args) not in (2, 3):
            raise _Refused("IIF() with an unexpected argument count.")
        cond, _ = _tf_translate(args[0], ctx)
        then_, _ = _tf_translate(args[1], ctx)
        else_ = _tf_translate(args[2], ctx)[0] if len(args) == 3 else "BLANK()"
        return f"IF({cond}, {then_}, {else_})"
    if fn == "DATEDIFF":
        if len(args) != 3:
            raise _Refused("DATEDIFF() with other than three arguments.")
        unit, _ = _tf_translate(args[0], ctx)
        d1, _ = _tf_translate(args[1], ctx)
        d2, _ = _tf_translate(args[2], ctx)
        return f"DATEDIFF({d1}, {d2}, {unit})"
    if fn == "STR":
        if len(args) != 1:
            raise _Refused("STR() with other than one argument.")
        inner, _ = _tf_translate(args[0], ctx)
        return f'FORMAT({inner}, "General Number")'

    if fn in _TF_SCALAR_FUNCS and _TF_SCALAR_FUNCS[fn]:
        dax_args = []
        for a in args:
            d, _ = _tf_translate(a, ctx)
            dax_args.append(d)
        return f"{_TF_SCALAR_FUNCS[fn]}({', '.join(dax_args)})"

    raise _Refused(f"Unsupported or unmapped function: `{fn}(...)`")


_IF_CASE_END_KW_RE = re.compile(r"\b(IF|CASE|END)\b", re.I)


def _tf_find_matching_end(body: str) -> int:
    """Index of the END that closes the ONE outer IF/CASE block this body
    is already inside (depth starts at 1) — not just the first literal
    'END' anywhere, which breaks the moment a branch contains its own
    nested IF/CASE. Very common in the real corpus: `SUM(IF x THEN y END)`
    as one WHEN-branch's expression inside an outer CASE, or an IF nested
    directly inside another IF's THEN branch. Field-bracket contents are
    skipped so a (hypothetical) field name containing the word END/IF/CASE
    can't miscount; string literals are skipped for the same reason."""
    depth, i, in_str, str_ch, bracket_depth = 1, 0, False, "", 0
    while i < len(body):
        c = body[i]
        if in_str:
            if c == str_ch:
                in_str = False
            i += 1
            continue
        if c in "'\"":
            in_str, str_ch = True, c
            i += 1
            continue
        if c == "[":
            bracket_depth += 1
        elif c == "]":
            bracket_depth -= 1
        if bracket_depth == 0:
            m = _IF_CASE_END_KW_RE.match(body, i)
            if m:
                kw = m.group(1).upper()
                if kw in ("IF", "CASE"):
                    depth += 1
                elif kw == "END":
                    depth -= 1
                    if depth == 0:
                        return m.start()
                i = m.end()
                continue
        i += 1
    return -1


def _tf_translate_if(expr: str, ctx: TfContext) -> str:
    body = expr[2:] if expr.upper().startswith("IF") else expr
    end_idx = _tf_find_matching_end(body)
    if end_idx == -1:
        raise _Refused("IF...END without a matching END.")
    body = body[:end_idx]

    # Split into IF-cond/THEN-expr, (ELSEIF-cond/THEN-expr)*, ELSE-expr? —
    # scanned at bracket/paren/brace/quote depth 0 only.
    keywords = _tf_find_all_top_keywords(body, ("THEN", "ELSEIF", "ELSE"))
    if not keywords or keywords[0][0] != "THEN":
        raise _Refused("IF without a top-level THEN.")

    segments: List[Tuple[str, str]] = []  # (keyword, text-until-next-keyword)
    positions = keywords + [("END", len(body))]
    cond_text = body[:positions[0][1] - len("THEN")] if positions[0][0] == "THEN" else body
    prev_kw, prev_end = "IF", 0
    for kw, end in positions:
        kw_len = {"THEN": 4, "ELSEIF": 6, "ELSE": 4, "END": 0}[kw]
        start_of_kw = end - kw_len
        segments.append((prev_kw, body[prev_end:start_of_kw]))
        prev_kw, prev_end = kw, end
    segments.append((prev_kw, body[prev_end:]))

    branches: List[Tuple[Optional[str], str]] = []  # (cond_or_None_for_else, expr_text)
    i = 0
    cur_cond = None
    while i < len(segments):
        kw, text = segments[i]
        if kw in ("IF", "ELSEIF"):
            cur_cond = text
            i += 1
            if i < len(segments) and segments[i][0] == "THEN":
                branches.append((cur_cond, segments[i][1]))
                i += 1
            else:
                raise _Refused("IF/ELSEIF without a following THEN.")
        elif kw == "ELSE":
            branches.append((None, text))
            i += 1
        else:
            i += 1

    dax = "BLANK()"
    for cond, then_text in reversed(branches):
        if cond is None:
            dax = _tf_translate(then_text, ctx)[0]
        else:
            cond_dax, _ = _tf_translate(cond, ctx)
            then_dax, _ = _tf_translate(then_text, ctx)
            dax = f"IF({cond_dax}, {then_dax}, {dax})"
    return dax


def _tf_find_all_top_keywords(s: str, keywords: Tuple[str, ...]) -> List[Tuple[str, int]]:
    """Top-level occurrences of `keywords` — top-level meaning both outside
    any (), [], {} nesting AND outside any nested IF...END / CASE...END
    block. The second condition matters as much as the first: a branch's
    THEN-expression can itself be a whole IF or CASE statement with no
    wrapping parens at all (`IF a THEN CASE b WHEN 1 THEN c WHEN 2 THEN d
    END END` — the inner CASE's own WHEN/THEN keywords must never be
    mistaken for the outer IF's branch separators, verified against a real
    formula in the corpus with exactly this shape)."""
    pattern = re.compile(r"\b(" + "|".join(keywords) + r")\b", re.I)
    block_kw = re.compile(r"\b(IF|CASE|END)\b", re.I)
    depth, block_depth, in_str, str_ch = 0, 0, False, ""
    out = []
    i = 0
    while i < len(s):
        c = s[i]
        if in_str:
            if c == str_ch:
                in_str = False
            i += 1
            continue
        if c in "'\"":
            in_str, str_ch = True, c
            i += 1
            continue
        if c in "([{":
            depth += 1
        elif c in ")]}":
            depth -= 1
        elif depth == 0:
            bm = block_kw.match(s, i)
            if bm:
                bkw = bm.group(1).upper()
                if bkw in ("IF", "CASE"):
                    block_depth += 1
                    i = bm.end()
                    continue
                if bkw == "END":
                    block_depth -= 1
                    i = bm.end()
                    continue
            if block_depth == 0:
                m = pattern.match(s, i)
                if m:
                    out.append((m.group(1).upper(), m.end()))
                    i = m.end()
                    continue
        i += 1
    return out


def _tf_translate_case(expr: str, ctx: TfContext) -> str:
    body = expr[4:] if expr.upper().startswith("CASE") else expr
    end_idx = _tf_find_matching_end(body)
    if end_idx == -1:
        raise _Refused("CASE...END without a matching END.")
    body = body[:end_idx]

    kws = _tf_find_all_top_keywords(body, ("WHEN", "THEN", "ELSE"))
    if not kws or kws[0][0] != "WHEN":
        raise _Refused("CASE without a top-level WHEN.")
    switch_expr = body[:kws[0][1] - len("WHEN")]
    switch_dax, _ = _tf_translate(switch_expr, ctx)

    pairs: List[str] = []
    else_dax = None
    i, prev_kw, prev_end = 0, "WHEN", kws[0][1]
    positions = kws[1:] + [("END", len(body))]
    for kw, end in positions:
        kw_len = {"WHEN": 4, "THEN": 4, "ELSE": 4, "END": 0}[kw]
        seg = body[prev_end:end - kw_len]
        if prev_kw == "WHEN":
            val_dax, _ = _tf_translate(seg, ctx)
            pairs.append(val_dax)
        elif prev_kw == "THEN":
            res_dax, _ = _tf_translate(seg, ctx)
            pairs.append(res_dax)
        elif prev_kw == "ELSE":
            else_dax, _ = _tf_translate(seg, ctx)
        prev_kw, prev_end = kw, end

    args = [switch_dax] + pairs
    if else_dax is not None:
        args.append(else_dax)
    return f"SWITCH({', '.join(args)})"


_LOD_HEAD_RE = re.compile(r"^\{\s*(FIXED|INCLUDE|EXCLUDE)\b(.*?):(.*)\}$", re.I | re.S)


def _tf_translate_lod(expr: str, ctx: TfContext) -> str:
    if _tf_matching_close(expr, 0, "{", "}") != len(expr) - 1:
        raise _Refused("Malformed LOD block (unbalanced braces).")
    m = _LOD_HEAD_RE.match(expr)
    if not m:
        raise _Refused("Not a recognized {FIXED|INCLUDE|EXCLUDE ...: ...} LOD shape.")
    kind, dims_text, agg_text = m.group(1).upper(), m.group(2), m.group(3)

    # Nested LOD anywhere inside this one's aggregation body is refused —
    # per scope, only single-level LOD is Tier 1.
    if re.search(r"\{\s*(FIXED|INCLUDE|EXCLUDE)\b", agg_text, re.I):
        raise _Refused("Nested LOD expressions are out of scope for this build.")

    dim_fields = [d.strip() for d in _tf_split_top(dims_text, ",") if d.strip()]
    dax_dims = []
    for d in dim_fields:
        dax, is_bare = _tf_translate(d, ctx)
        if not is_bare:
            raise _Refused("LOD dimension isn't a plain field reference.")
        dax_dims.append(dax)

    agg_dax, _ = _tf_translate(agg_text, ctx)

    if kind == "FIXED":
        if not dax_dims:
            return f"CALCULATE({agg_dax}, ALL('{ctx.table_name}'))"
        return f"CALCULATE({agg_dax}, ALLEXCEPT('{ctx.table_name}', {', '.join(dax_dims)}))"
    if kind == "EXCLUDE":
        if not dax_dims:
            raise _Refused("EXCLUDE with no dimensions.")
        return f"CALCULATE({agg_dax}, REMOVEFILTERS({', '.join(dax_dims)}))"
    # INCLUDE: outer aggregator should match the visual's own default
    # aggregation of the field, which a model-layer translator can't see —
    # default to SUMX and flag for verification, per the build spec.
    if not dax_dims:
        raise _Refused("INCLUDE with no dimensions.")
    if len(dax_dims) == 1:
        return f"SUMX(VALUES({dax_dims[0]}), CALCULATE({agg_dax}))"
    return f"SUMX(SUMMARIZE('{ctx.table_name}', {', '.join(dax_dims)}), CALCULATE({agg_dax}))"


def _tf_dependency_order(logical_columns: Dict[str, Dict[str, Any]]) -> List[str]:
    """Calc field ids in dependency order (a field referencing another calc
    field is translated after it) — real corpus example: a FIXED LOD
    summing two other calc fields inside its aggregation body. Same
    ordering problem the Lakeview exporter's `_lakeview_measure_home_table`
    solves for DAX measure dependencies, reimplemented here for Tableau's
    own formula grammar (not reused — different reference syntax).

    Kahn's algorithm; a genuine circular reference (rare, and arguably a
    modeling error in the source workbook) falls back to leaving the
    cyclic fields in their original order rather than raising — the
    translator will refuse them individually if their dependency truly
    isn't resolvable yet, which is a safer failure mode than crashing the
    whole conversion.
    """
    calc_ids = [fid for fid, c in logical_columns.items() if c.get("formula")]
    calc_id_set = set(calc_ids)
    deps: Dict[str, Set[str]] = {}
    for fid in calc_ids:
        formula = logical_columns[fid]["formula"] or ""
        refs = {m for m in re.findall(r"\[([^\[\]]+)\]", formula) if m in calc_id_set and m != fid}
        deps[fid] = refs

    in_degree = {fid: 0 for fid in calc_ids}
    dependents: Dict[str, List[str]] = {fid: [] for fid in calc_ids}
    for fid, refs in deps.items():
        for r in refs:
            dependents[r].append(fid)
            in_degree[fid] += 1

    queue = [fid for fid in calc_ids if in_degree[fid] == 0]
    ordered: List[str] = []
    while queue:
        fid = queue.pop(0)
        ordered.append(fid)
        for dep in dependents[fid]:
            in_degree[dep] -= 1
            if in_degree[dep] == 0:
                queue.append(dep)

    if len(ordered) != len(calc_ids):
        remaining = [fid for fid in calc_ids if fid not in ordered]
        ordered.extend(remaining)  # cyclic remainder, original order
    return ordered


def _pick_relationship_key(mappings: List[Tuple[str, str]]) -> Optional[Tuple[Tuple[str, str], Tuple[str, str], bool]]:
    """One representative (left, right) column-mapping pair out of what can
    be dozens per datasource-relationship edge (real corpus finding: most
    of a <datasource-relationship>'s <map> entries are Tableau's own
    auto-generated date-hierarchy rollups — hr:/dy:/wk:/mn:/qr:/yr:-prefixed
    variants of the same underlying timestamp, not distinct business join
    keys). Prefers a `none:`-prefixed, `:nk`-suffixed key (Tableau's own
    convention for "no time-grain transform, natural key") when one exists.

    Returns (left_(ds,field), right_(ds,field), was_ambiguous) or None if
    no mapping parses at all.
    """
    parsed = []
    for left_raw, right_raw in mappings:
        l_ds, l_field = _parse_colref(left_raw)
        r_ds, r_field = _parse_colref(right_raw)
        if l_ds and l_field and r_ds and r_field:
            score = 0
            if ":none:" in f":{left_raw}:" or left_raw.count(":") >= 2 and left_raw.split(":")[-2] == "none":
                score += 1
            if left_raw.rstrip("]").endswith(":nk"):
                score += 1
            parsed.append((score, (l_ds, l_field), (r_ds, r_field)))
    if not parsed:
        return None
    parsed.sort(key=lambda t: -t[0])
    best = parsed[0]
    ambiguous = len({p[1:] for p in parsed}) > 1
    return best[1], best[2], ambiguous


def build_semantic_model(
    twbx: Dict[str, Any],
    table_name_overrides: Optional[Dict[str, str]] = None,
    llm_fallback: Optional[Callable[[str, str, str], Optional[Tuple[str, str]]]] = None,
) -> Dict[str, Any]:
    """Stage 2: tables, relationships, measures/columns, RLS role candidates.

    `llm_fallback(formula, table_name, deep_context)`, if given, is called
    ONLY for formulas the deterministic translator refuses — never for ones
    it already handled — and must return either (dax, note) or None (if it
    also declines). Results from it are tagged tier 'llm_assisted' /
    Status 'Migrated — verify (AI-translated)', never silently trusted
    the way a deterministic 'simple' translation is.
    """
    table_name_overrides = table_name_overrides or {}
    tables: Dict[str, Dict[str, Any]] = {}
    ds_to_table: Dict[str, str] = {}
    measures: List[Dict[str, Any]] = []
    calc_columns: List[Dict[str, Any]] = []
    report_rows: List[Dict[str, Any]] = []
    rls_roles: List[Dict[str, Any]] = []

    def add_report(obj_type: str, table: str, name: str, status: str, severity: str, detail: str) -> None:
        report_rows.append({
            "Object Type": obj_type, "Table": table, "Name": name,
            "Status": status, "Severity": severity, "Detail": detail,
        })

    # --- Tables + field maps -------------------------------------------
    for ds_name, ds in twbx["datasources"].items():
        if ds.get("has_custom_sql"):
            add_report("Table", ds["caption"], ds["caption"], "Not migrated", "Medium",
                       "Custom SQL datasources are out of scope for this build.")
            continue
        leaves = _leaf_tables(ds.get("relation"))
        if not leaves:
            add_report("Table", ds["caption"], ds["caption"], "Not migrated", "Medium",
                       "No physical table relation could be resolved for this datasource.")
            continue
        home_leaf = leaves[0]
        pbi_table_name = table_name_overrides.get(ds_name) or pbi_name(ds["caption"], ds_name)
        ds_to_table[ds_name] = pbi_table_name

        columns = []
        seen_cols: Set[str] = set()
        for field_id, sch in ds["schema"].items():
            if field_id in seen_cols or not field_id:
                continue
            seen_cols.add(field_id)
            columns.append({
                "name": pbi_name(field_id), "source_column": sch["remote_name"] or field_id,
                "datatype": sch["local_type"] or "string",
            })
        tables[pbi_table_name] = {
            "source_ds": ds_name, "physical_table": home_leaf.get("physical_table") or home_leaf.get("name"),
            "columns": columns,
        }
        if len(leaves) > 1:
            add_report(
                "Table", pbi_table_name, pbi_table_name, "Migrated — verify context", "Medium",
                f"This datasource's physical layer joins {len(leaves)} tables; only the first "
                "leaf's schema is represented here (physical-join splitting is a known Tier-1 "
                "limitation — this corpus has no join example to validate against).",
            )

    # --- Field maps per datasource (for the formula translator) --------
    field_maps: Dict[str, Dict[str, Dict[str, str]]] = {}
    for ds_name, ds in twbx["datasources"].items():
        if ds_name not in ds_to_table:
            continue
        fmap: Dict[str, Dict[str, str]] = {}
        for field_id, col in ds["logical_columns"].items():
            fmap[field_id] = {
                "kind": "measure" if col["formula"] else "column",
                "name": pbi_name(col["caption"] or field_id),
            }
        for field_id, sch in ds["schema"].items():
            if field_id not in fmap:
                fmap[field_id] = {"kind": "column", "name": pbi_name(field_id)}
        field_maps[ds_name] = fmap

    # --- Measures / calculated columns, dependency-ordered -------------
    tier_status = {
        "simple": "Translated", "lod": "Migrated — verify context",
        "llm_assisted": "Migrated — verify (AI-translated)",
    }
    for ds_name, ds in twbx["datasources"].items():
        if ds_name not in ds_to_table:
            continue
        table_name = ds_to_table[ds_name]
        group_names = {c["group_name"] for c in ds.get("rls_candidates", [])}
        ctx = TfContext(table_name, field_maps[ds_name], group_names=group_names)
        order = _tf_dependency_order(ds["logical_columns"])
        for field_id in order:
            col = ds["logical_columns"][field_id]
            formula = col["formula"]
            display_name = pbi_name(col["caption"] or field_id)
            dax, tier, note = translate_tableau_formula(formula, ctx)
            if dax is None and llm_fallback is not None:
                ai_result = llm_fallback(formula, table_name, note)
                if ai_result is not None:
                    dax, ai_note = ai_result
                    tier, note = "llm_assisted", ai_note

            is_measure = bool((col.get("role") or "").lower() == "measure" or _uses_aggregation(formula))
            target_list = measures if is_measure else calc_columns
            if dax is not None:
                target_list.append({
                    "table": table_name, "name": display_name, "dax": dax,
                    "source_formula": formula,
                })
                add_report(
                    "Measure" if is_measure else "Calculated column", table_name, display_name,
                    tier_status.get(tier, "Translated"), "Info" if tier == "simple" else "Medium",
                    note,
                )
            else:
                add_report(
                    "Measure" if is_measure else "Calculated column", table_name, display_name,
                    "Not translated", "Medium", note,
                )

    # --- Relationships ---------------------------------------------------
    relationships: List[Dict[str, Any]] = []
    for dsr in twbx["datasource_relationships"]:
        src_table = ds_to_table.get(dsr["source_ds"])
        tgt_table = ds_to_table.get(dsr["target_ds"])
        if not src_table or not tgt_table:
            continue
        picked = _pick_relationship_key(dsr["mappings"])
        if picked is None:
            add_report(
                "Relationship", src_table, f"{src_table} -> {tgt_table}", "Not translated", "Medium",
                "No resolvable column-mapping pair found for this datasource relationship.",
            )
            continue
        (l_ds, l_field), (r_ds, r_field), ambiguous = picked
        l_name = field_maps.get(l_ds, {}).get(l_field, {}).get("name", pbi_name(l_field))
        r_name = field_maps.get(r_ds, {}).get(r_field, {}).get("name", pbi_name(r_field))
        rel_id = det_id("rel", src_table, l_name, tgt_table, r_name)
        relationships.append({
            "id": rel_id, "from_table": src_table, "from_column": l_name,
            "to_table": tgt_table, "to_column": r_name,
        })
        add_report(
            "Relationship", src_table, f"{src_table}[{l_name}] -> {tgt_table}[{r_name}]",
            "Migrated — verify context", "Medium",
            "Cardinality defaults to single-direction, one-to-many — Tableau's XML doesn't "
            "reliably expose key uniqueness, so this must be verified in Power BI Desktop."
            + (" Multiple candidate key pairs were found on this edge; the most likely business "
               "key was picked heuristically — verify it's the right one." if ambiguous else ""),
        )

    # --- Physical joins within a single datasource (never observed in the
    # validation corpus, per Section 1.1 — implemented per spec, untested
    # against a real example) ------------------------------------------
    for ds_name, ds in twbx["datasources"].items():
        rel_tree = ds.get("relation")
        if rel_tree and rel_tree.get("kind") == "join":
            add_report(
                "Relationship", ds["caption"], ds["caption"], "Migrated — verify context", "Medium",
                "This datasource's physical layer is a join tree. Only the first leaf table is "
                "represented as a Power BI table (see the Table row above) — the join itself "
                "was not translated into a separate relationship, since no example of this shape "
                "exists in the files this build was validated against. Recreate this join "
                "manually as a Power BI relationship.",
            )

    # --- RLS role candidates (surfaced for confirmation, never auto-applied) --
    # One Power BI role commonly needs a `tablePermission` line per table it
    # secures — grouped by AD group (not by group+table+column), so the UI
    # shows one candidate per real security group (23 in the validation
    # corpus's largest file) rather than one row per datasource that
    # happens to reference it (339 raw ISMEMBEROF blocks in that same
    # file — the flat count the build spec explicitly warns would make the
    # UI unusable if surfaced un-deduplicated).
    rls_by_group: Dict[str, Dict[str, Any]] = {}
    for ds_name, ds in twbx["datasources"].items():
        table_name = ds_to_table.get(ds_name)
        if not table_name:
            continue
        for cand in ds.get("rls_candidates", []):
            field_entry = field_maps.get(ds_name, {}).get(cand["level_field"])
            column_name = field_entry["name"] if field_entry else pbi_name(cand["level_field"])
            group = rls_by_group.setdefault(cand["ad_group"], {"ad_group": cand["ad_group"], "permissions": []})
            perm_key = (table_name, column_name)
            existing = next((p for p in group["permissions"] if (p["table"], p["column"]) == perm_key), None)
            if existing is None:
                group["permissions"].append({
                    "table": table_name, "column": column_name,
                    "allowed_values": list(cand["allowed_values"]),
                })
            else:
                existing["allowed_values"] = sorted(set(existing["allowed_values"]) | set(cand["allowed_values"]))
    deduped_rls = list(rls_by_group.values())

    # --- Out-of-scope items get an explicit report row, never a silent drop --
    for ds_name, ds in twbx["datasources"].items():
        if ds_name in ds_to_table:
            continue
    for param_name, p in twbx["parameters"].items():
        if p["domain_type"] != "list":
            add_report("Parameter", "(Parameters)", p["caption"], "Not migrated", "Low",
                       "Only list-domain parameters are supported.")
    for action in twbx["actions"]:
        if action.get("command") and action["command"] != "tsc:tsl-filter":
            add_report("Action", "(workbook)", action.get("name") or action["command"], "Not migrated", "Low",
                       "Only plain filter actions are supported.")

    report_df = pd.DataFrame(report_rows, columns=["Object Type", "Table", "Name", "Status", "Severity", "Detail"])
    return {
        "tables": tables,
        "relationships": relationships,
        "measures": measures,
        "calculated_columns": calc_columns,
        "rls_roles": deduped_rls,
        "field_maps": field_maps,
        "ds_to_table": ds_to_table,
        "report_df": report_df,
    }


# ==========================================================================
# Stage 3 — build_report_model(): mark class -> visual type, zone -> layout
# ==========================================================================
# Mark class + shelf shape -> Power BI visualType, per the build spec's
# Section 10 lookup table (itself adopted from the architecture doc's own
# worked table, Section 6). Circle is scatterChart by default; Square is
# matrix by default — the spec's own "or treemap if color-intensity-driven"
# refinement for Square, and the heatmap-vs-scatter disambiguation for
# Circle, would need per-visual color-encoding-type inspection this build
# doesn't attempt — refuse to the safer, more common default rather than
# guess at the fancier one.
_MARK_TO_VISUAL: Dict[str, str] = {
    "Bar": "clusteredColumnChart",
    "Line": "lineChart",
    "Square": "matrix",
    "Circle": "scatterChart",
    "Text": "tableEx",
    "Automatic": "tableEx",
    "Pie": "pieChart",
}
_MARK_UNSUPPORTED_NOTE = {
    "Shape": "Custom shape marks have no confident Power BI visual equivalent.",
    "Density": "Density marks are Tier 3 (out of scope) — no PBIR equivalent.",
    "Polygon": "Polygon marks are Tier 3 (out of scope) — no PBIR equivalent.",
}


def _visual_type_for_worksheet(ws: Dict[str, Any]) -> Tuple[Optional[str], str]:
    """(visualType_or_None, reason_if_none)."""
    mark = ws.get("mark_class")
    all_fields = ws.get("encodings", []) + ws.get("shelf_refs", [])
    n_fields = len({(f.get("ds"), f.get("field")) for f in all_fields if f.get("field")})

    if n_fields == 0:
        return None, "No resolvable field references on this worksheet."
    if mark is None and n_fields <= 2:
        # No mark class parsed and very few fields — most consistent with a
        # worksheet used purely as a filter control (a slicer never carries
        # a mark). Treated as a best-effort default, not a confident
        # detection — flagged 'Migrated — verify context', not 'Translated'.
        return "slicer", "verify"
    if mark in _MARK_UNSUPPORTED_NOTE:
        return None, _MARK_UNSUPPORTED_NOTE[mark]
    if mark in _MARK_TO_VISUAL:
        return _MARK_TO_VISUAL[mark], "ok"
    if mark is None:
        return "tableEx", "verify"
    return None, f"Unsupported mark class: {mark}"


def _resolve_field(ref: Dict[str, str], field_maps: Dict[str, Dict[str, Dict[str, str]]],
                    ds_to_table: Dict[str, str]) -> Optional[Dict[str, Any]]:
    ds, field_id = ref.get("ds"), ref.get("field")
    if not ds or not field_id:
        return None
    table = ds_to_table.get(ds)
    fmap = field_maps.get(ds, {})
    entry = fmap.get(field_id)
    if table is None or entry is None:
        return None
    return {"table": table, "name": entry["name"], "kind": entry["kind"]}


def _queryState_projection(field: Dict[str, Any]) -> Dict[str, Any]:
    if field["kind"] == "measure":
        return {"field": {"Measure": {"Expression": {"SourceRef": {"Entity": field["table"]}},
                                       "Property": field["name"]}}}
    return {"field": {"Column": {"Expression": {"SourceRef": {"Entity": field["table"]}},
                                  "Property": field["name"]}}}


def build_report_model(
    twbx: Dict[str, Any], semantic: Dict[str, Any],
) -> Dict[str, Any]:
    """Stage 3: one PBIR page per dashboard, one visual per worksheet zone
    that resolves to a known mark type and at least one resolvable field."""
    field_maps = semantic["field_maps"]
    ds_to_table = semantic["ds_to_table"]
    pages: List[Dict[str, Any]] = []
    report_rows: List[Dict[str, Any]] = []

    def add_report(obj_type: str, page: str, name: str, status: str, severity: str, detail: str) -> None:
        report_rows.append({
            "Object Type": obj_type, "Table": page, "Name": name,
            "Status": status, "Severity": severity, "Detail": detail,
        })

    for dash_name, dash in twbx["dashboards"].items():
        canvas_w, canvas_h = dash["maxwidth"], dash["maxheight"]
        visuals: List[Dict[str, Any]] = []

        def walk(zone: Optional[Dict[str, Any]]) -> None:
            if zone is None:
                return
            ws_name = zone.get("name")
            if ws_name and ws_name in twbx["worksheets"] and not zone.get("children"):
                ws = twbx["worksheets"][ws_name]
                x = zone["x"] / 100000 * canvas_w
                y = zone["y"] / 100000 * canvas_h
                w = zone["w"] / 100000 * canvas_w
                h = zone["h"] / 100000 * canvas_h

                vtype, reason = _visual_type_for_worksheet(ws)
                if vtype is None:
                    add_report("Visual", dash_name, ws_name, "Not migrated", "Medium", reason)
                    return

                fields = []
                for ref in ws.get("encodings", []) + ws.get("shelf_refs", []):
                    resolved = _resolve_field(ref, field_maps, ds_to_table)
                    if resolved is not None:
                        fields.append((ref.get("channel") or ref.get("shelf") or "Values", resolved))
                if not fields:
                    add_report("Visual", dash_name, ws_name, "Not migrated", "Medium",
                               "No field reference on this worksheet resolved to a known table/column.")
                    return

                query_state: Dict[str, Any] = {}
                bucket = "Values" if vtype == "tableEx" else "Y"
                seen_props: Set[Tuple[str, str, str]] = set()
                for channel, f in fields:
                    key = (channel, f["table"], f["name"])
                    if key in seen_props:
                        continue
                    seen_props.add(key)
                    role = {"color": "Series", "rows": "Category", "tooltip": None, "detail": None}.get(channel, bucket)
                    if role is None:
                        continue
                    query_state.setdefault(role, {"projections": []})["projections"].append(
                        _queryState_projection(f)
                    )

                visual_id = det_id("visual", dash_name, ws_name)
                visuals.append({
                    "name": visual_id, "worksheet": ws_name,
                    "position": {"x": round(x, 1), "y": round(y, 1), "width": round(w, 1),
                                 "height": round(h, 1), "z": 1000.0 + len(visuals)},
                    "visual": {"visualType": vtype, "queryState": query_state},
                })
                status = "Translated" if reason == "ok" else "Migrated — verify context"
                detail = ("Visual type and field wells inferred from the mark class and shelf "
                          "fields." if reason == "ok" else
                          "Visual type inferred heuristically (no explicit mark class, or too few "
                          "fields to be confident) — verify against the original worksheet.")
                add_report("Visual", dash_name, ws_name, status, "Info" if reason == "ok" else "Medium", detail)
                return

            for child in zone.get("children", []):
                walk(child)
            # A leaf zone with no worksheet name and no children is a
            # layout container, image, or web-object placeholder.
            if not zone.get("children") and not ws_name:
                ztype = zone.get("type") or ""
                param = zone.get("param") or ""
                if param.startswith("Image/") or ztype == "bitmap":
                    add_report("Visual", dash_name, param or "(image)", "Not migrated", "Low",
                               "Layout element present (image) — not converted, needs manual placement.")

        walk(dash.get("zones"))
        pages.append({
            "name": dash_name, "width": canvas_w, "height": canvas_h, "visuals": visuals,
        })

    report_df = pd.DataFrame(report_rows, columns=["Object Type", "Table", "Name", "Status", "Severity", "Detail"])
    return {"pages": pages, "report_df": report_df}


# ==========================================================================
# Stage 4 — emit_tmdl() / emit_pbir(): target file contents
# ==========================================================================

_TMDL_TYPE_MAP: Dict[str, str] = {
    "string": "string", "integer": "int64", "real": "double", "boolean": "boolean",
    "date": "dateTime", "datetime": "dateTime", "table": "string",
}


def _tmdl_ident(name: str) -> str:
    """A TMDL object name — quoted with single quotes whenever it contains
    a space or any character TMDL treats as syntactically significant;
    bare otherwise. Matches the spec's own worked examples (`'Sales $'`
    quoted, `SalesFactTable`/`terr_nm` bare)."""
    if re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", name):
        return name
    return "'" + name.replace("'", "''") + "'"


def _tmdl_string(s: str) -> str:
    return '"' + str(s).replace('"', '""') + '"'


def emit_tmdl(semantic: Dict[str, Any], catalog: str, schema: str) -> Dict[str, str]:
    """Stage 4a: {relative_path -> TMDL text}, one file per table plus
    relationships.tmdl and (if any roles were confirmed) roles.tmdl."""
    files: Dict[str, str] = {}
    measures_by_table: Dict[str, List[Dict[str, Any]]] = {}
    for m in semantic["measures"]:
        measures_by_table.setdefault(m["table"], []).append(m)
    calc_cols_by_table: Dict[str, List[Dict[str, Any]]] = {}
    for c in semantic["calculated_columns"]:
        calc_cols_by_table.setdefault(c["table"], []).append(c)

    for table_name, table in semantic["tables"].items():
        lines = [f"table {_tmdl_ident(table_name)}", ""]
        for col in table["columns"]:
            dtype = _TMDL_TYPE_MAP.get((col["datatype"] or "").lower(), "string")
            lines.append(f"\tcolumn {_tmdl_ident(col['name'])}")
            lines.append(f"\t\tdataType: {dtype}")
            lines.append(f"\t\tsourceColumn: {col['source_column']}")
            lines.append("")
        for cc in calc_cols_by_table.get(table_name, []):
            lines.append(f"\tcolumn {_tmdl_ident(cc['name'])} = {cc['dax']}")
            lines.append("")
        for meas in measures_by_table.get(table_name, []):
            lines.append(f"\tmeasure {_tmdl_ident(meas['name'])} = {meas['dax']}")
            lines.append("")
        source_table = table["physical_table"] or table_name
        lines.extend([
            f"\tpartition {_tmdl_ident(table_name + '-Partition')} = m",
            "\t\tmode: import",
            "\t\tsource =",
            "\t\t\tlet",
            f"\t\t\t\tSource = {catalog}.{schema}.{source_table}  // placeholder — confirm the real M query for your source",
            "\t\t\tin",
            "\t\t\t\tSource",
            "",
        ])
        files[f"tables/{slug(table_name, 'table')}.tmdl"] = "\n".join(lines)

    if semantic["relationships"]:
        rel_lines = []
        for rel in semantic["relationships"]:
            rel_lines.append(f"relationship {rel['id']}")
            rel_lines.append(f"\tfromColumn: {_tmdl_ident(rel['from_table'])}[{rel['from_column']}]")
            rel_lines.append(f"\ttoColumn: {_tmdl_ident(rel['to_table'])}[{rel['to_column']}]")
            rel_lines.append("\t# TODO verify cardinality — Tableau XML does not expose key uniqueness")
            rel_lines.append("\tcrossFilteringBehavior: OneDirection")
            rel_lines.append("")
        files["relationships.tmdl"] = "\n".join(rel_lines)

    files["database.tmdl"] = "database\n\tcompatibilityLevel: 1567\n"
    return files


def emit_roles_tmdl(confirmed_roles: List[Dict[str, Any]]) -> str:
    """Stage 4a (roles are emitted separately from emit_tmdl since the
    build spec requires them opt-in — only user-confirmed groups from the
    RLS candidates review, never auto-applied)."""
    lines = []
    for role in confirmed_roles:
        lines.append(f"role {_tmdl_ident(role['ad_group'])}")
        lines.append("\tmodelPermission: read")
        lines.append("")
        for perm in role["permissions"]:
            values = " , ".join(_tmdl_string(v) for v in perm["allowed_values"])
            filter_expr = f"'{perm['table']}'[{perm['column']}] IN {{{values}}}"
            lines.append(f"\ttablePermission {_tmdl_ident(perm['table'])} = {_tmdl_string(filter_expr)}")
        lines.append("")
    return "\n".join(lines)


def emit_pbir(report: Dict[str, Any]) -> Dict[str, Any]:
    """Stage 4b: {relative_path -> dict}, JSON-serialized at packaging time.
    Mirrors the exact field shapes app.py's own PBIR *reader* already
    trusts (`_pbir_page_visuals`: position.x/y/z/width/height,
    visual.visualType, visual.query.queryState / SourceRef.Entity/Property)
    — same schema, write direction."""
    files: Dict[str, Any] = {}
    page_index = []
    for page in report["pages"]:
        page_id = slug(page["name"], "page")
        page_index.append({"displayName": page["name"], "id": page_id})
        files[f"pages/{page_id}/page.json"] = {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/1.4.0/schema.json",
            "name": page_id, "displayName": page["name"],
            "width": page["width"], "height": page["height"],
        }
        for v in page["visuals"]:
            files[f"pages/{page_id}/visuals/{v['name']}/visual.json"] = {
                "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/1.5.0/schema.json",
                "name": v["name"],
                "position": v["position"],
                "visual": v["visual"],
            }
    files["pages/pages.json"] = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json",
        "pageOrder": [p["id"] for p in page_index],
        "activePageName": page_index[0]["id"] if page_index else "",
    }
    return files


# ==========================================================================
# Stage 5 — package_pbip(): zip the whole project
# ==========================================================================
# Beyond the TMDL/PBIR content files themselves, a real .pbip project needs
# a handful of small pointer/metadata files (.pbip, definition.pbir,
# definition.pbism, .platform) that tell Power BI Desktop how the pieces
# connect. These are built from documented PBIP structure, not from
# unzipping a real .pbip this build was validated against (none was
# available in this environment) — this is the one part of the whole
# pipeline that genuinely needs the manual "open it in real Power BI
# Desktop" acceptance step called out in the build spec's test plan; the
# TMDL/PBIR content itself is validated against the real .twbx corpus, but
# these wrapper files are not.

def package_pbip(
    report_name: str, tmdl_files: Dict[str, str], pbir_files: Dict[str, Any],
) -> bytes:
    safe_name = re.sub(r'[\\/:*?"<>|]', "_", report_name).strip() or "Converted Report"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(f"{safe_name}.pbip", json.dumps({
            "version": "1.0",
            "artifacts": [{"report": {"path": f"{safe_name}.Report"}}],
            "settings": {"enableAutoRecovery": True},
        }, indent=2))

        model_root = f"{safe_name}.SemanticModel"
        z.writestr(f"{model_root}/definition.pbism", json.dumps({"version": "4.2"}, indent=2))
        z.writestr(f"{model_root}/.platform", json.dumps({
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
            "metadata": {"type": "SemanticModel", "displayName": safe_name},
            "config": {"version": "2.0", "logicalId": det_id("model", safe_name, length=32)},
        }, indent=2))
        for rel_path, content in tmdl_files.items():
            z.writestr(f"{model_root}/definition/{rel_path}", content)

        report_root = f"{safe_name}.Report"
        z.writestr(f"{report_root}/definition.pbir", json.dumps({
            "version": "4.0",
            "datasetReference": {"byPath": {"path": f"../{model_root}"}},
        }, indent=2))
        z.writestr(f"{report_root}/.platform", json.dumps({
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
            "metadata": {"type": "Report", "displayName": safe_name},
            "config": {"version": "2.0", "logicalId": det_id("report", safe_name, length=32)},
        }, indent=2))
        for rel_path, content in pbir_files.items():
            z.writestr(f"{report_root}/definition/{rel_path}", json.dumps(content, indent=2))

    return buf.getvalue()


def _uses_aggregation(formula: Optional[str]) -> bool:
    if not formula:
        return False
    return bool(re.search(
        r"\b(SUM|AVG|COUNT|COUNTD|MIN|MAX|MEDIAN|STDEV|VAR|ATTR)\s*\(|\{\s*(FIXED|INCLUDE|EXCLUDE)\b",
        formula, re.I,
    ))


def translate_tableau_formula(
    formula: str, ctx: TfContext,
) -> Tuple[Optional[str], str, str]:
    """(dax_or_None, tier, note). tier is 'simple', 'lod', or 'refuse'.

    Tier is decided by whether an LOD block appears anywhere in the
    formula (checked textually before translation, since a refused LOD
    still deserves an 'lod'-flavoured note rather than a generic one) —
    the recursive translator itself doesn't track tier per sub-expression,
    it either fully succeeds or raises _Refused with a specific reason.
    """
    has_lod = bool(re.search(r"\{\s*(FIXED|INCLUDE|EXCLUDE)\b", formula, re.I))
    try:
        dax, _ = _tf_translate(formula, ctx)
        return dax, ("lod" if has_lod else "simple"), "Translated."
    except _Refused as exc:
        return None, "refuse", exc.reason


# ==========================================================================
# Optional AI fallback (Tier 2.5) — routes ONLY deterministic refusals
# ==========================================================================
# Reuses whatever LLM provider/key the user already configured in app.py's
# sidebar — app.py passes its own `call_llm(prompt) -> str` in as a plain
# callable rather than this module importing app.py's LLM machinery
# directly, which would create an import cycle (app.py imports this module
# at its own top level).

def _parse_llm_json_local(raw: str) -> Dict[str, Any]:
    text = (raw or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.I).strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", text, flags=re.S)
        if not m:
            return {}
        try:
            return json.loads(m.group(0))
        except json.JSONDecodeError:
            return {}


def _compact_schema_text(semantic: Dict[str, Any], max_tables: int = 15, max_cols: int = 15) -> str:
    lines = []
    tables = list(semantic["tables"].items())
    for tname, t in tables[:max_tables]:
        cols = ", ".join(c["name"] for c in t["columns"][:max_cols])
        lines.append(f"- {tname}: {cols}")
    if len(tables) > max_tables:
        lines.append(f"… +{len(tables) - max_tables} more tables")
    return "\n".join(lines) or "(no tables)"


def make_llm_formula_fallback(
    call_llm: Callable[[str], str], schema_text: str,
) -> Callable[[str, str, str], Optional[Tuple[str, str]]]:
    """Wraps app.py's injected call_llm into the (formula, table,
    refusal_reason) -> (dax, note) | None shape build_semantic_model's
    optional llm_fallback expects.

    Only ever invoked on formulas the deterministic translator already
    refused — this is the deep dive's own "Tier 2.5": a middle ground
    between confident translation and outright refusal, worth having but
    never trusted the way a deterministic 'Translated' result is. An LLM
    can produce syntactically perfect DAX that computes the wrong number;
    every result from here is tagged 'Migrated — verify (AI-translated)'
    by build_semantic_model, never 'Translated'.
    """

    def fallback(formula: str, table_name: str, refusal_reason: str) -> Optional[Tuple[str, str]]:
        prompt = f"""You are a senior Power BI / DAX consultant. A deterministic rule-based translator
could not convert the following Tableau calculated-field formula to DAX, for this specific reason:
{refusal_reason}

Tableau formula:
{formula}

Target table: {table_name}

Model schema (abbreviated, for context only):
{schema_text}

If you can confidently translate this to a single valid DAX expression, do so — reference columns
as 'Table'[Column] and other measures as [Measure Name], using only tables/columns that appear in
the schema above. If you are NOT confident the translation is correct, return null for "dax" rather
than guessing — a wrong-but-plausible DAX expression is worse than admitting you can't do it.

Respond with ONLY a single JSON object (no markdown fences, no commentary):
{{"dax": "the DAX expression, or null if not confident", "note": "one sentence explaining the translation, or why you declined"}}"""
        try:
            raw = call_llm(prompt)
        except Exception:  # noqa: BLE001
            return None
        parsed = _parse_llm_json_local(raw)
        dax = parsed.get("dax")
        if not dax or not isinstance(dax, str):
            return None
        note = str(parsed.get("note") or "AI-translated — review before use.")
        return dax, note

    return fallback


# ==========================================================================
# UI — render_twbx_conversion_page()
# ==========================================================================
# Structurally cloned from the "Databricks Lakeview Export" page in app.py
# (upload -> configure -> generate -> report -> download), per the build
# spec — same shape, new content. This module has no `show_table` to reuse
# (that lives in app.py and importing it back would create the same import
# cycle noted above), so a small local table+CSV-download helper stands in.

def _show_df_local(df: pd.DataFrame, label: str, key: str, height: int = 360) -> None:
    st.dataframe(df, width="stretch", height=height, hide_index=True)
    st.download_button(
        f"⬇ Download {label} (.csv)", df.to_csv(index=False).encode("utf-8"),
        file_name=f"{slug(label)}.csv", mime="text/csv", key=f"dl_{key}", width="stretch",
    )


def render_twbx_conversion_page(
    default_catalog: str = "workspace",
    default_schema: str = "default",
    call_llm: Optional[Callable[[str], str]] = None,
    llm_ready: bool = False,
) -> None:
    st.subheader("Convert a Tableau workbook (.twbx) to a Power BI Project")
    st.caption(
        "**Tier 1 only** — schema, relationships, single-level LOD and simple calculated fields, "
        "list-domain parameters, mark-to-visual mapping, dashboard layout, and `ISMEMBEROF()`-"
        "based RLS candidates. Nested LOD, table calculations, custom-SQL datasources, and a "
        "handful of other constructs are explicitly refused with a reason in the conversion "
        "report below — never silently approximated. The output is a **Power BI Project (.pbip)** "
        "folder to open in Power BI Desktop and save as `.pbix` from there — there's no supported "
        "way to write a `.pbix` file directly, the same constraint the Databricks Lakeview export "
        "works under."
    )

    uploaded = st.file_uploader("Choose a .twbx file", type=["twbx"], key="twbx_upload")
    if uploaded is None:
        st.info("👈 Upload a **.twbx** (Tableau packaged workbook) to get started.", icon="🗂️")
        return

    twbx_bytes = uploaded.getvalue()
    try:
        twbx = parse_twbx(twbx_bytes)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not parse this file as a .twbx workbook: {exc}")
        return

    k = st.columns(4)
    k[0].metric("Datasources", len(twbx["datasources"]))
    k[1].metric("Worksheets", len(twbx["worksheets"]))
    k[2].metric("Dashboards", len(twbx["dashboards"]))
    k[3].metric("Calculated fields", sum(
        1 for ds in twbx["datasources"].values() for c in ds["logical_columns"].values() if c["formula"]
    ))

    c1, c2 = st.columns(2)
    catalog = c1.text_input(
        "Unity Catalog / warehouse catalog", value=default_catalog, key="twbx_catalog",
        help="Used only as a placeholder in each table's M partition source — you'll need to "
             "confirm the real connection details in Power BI Desktop regardless.",
    )
    schema_name = c2.text_input("Schema", value=default_schema, key="twbx_schema")

    with st.expander("Table name mapping", expanded=False, icon=":material/swap_horiz:"):
        st.caption(
            "Every Tableau datasource becomes one Power BI table, named from its caption below "
            "(never flattened into merged tables — see the note at the bottom of this page). "
            "Edit any row where you'd rather use a different name."
        )
        ds_names = list(twbx["datasources"].keys())
        name_map_df = pd.DataFrame({
            "Tableau Datasource": [twbx["datasources"][n]["caption"] for n in ds_names],
            "Power BI Table Name": [pbi_name(twbx["datasources"][n]["caption"], n) for n in ds_names],
        })
        edited = st.data_editor(
            name_map_df, hide_index=True, width="stretch", key="twbx_name_map",
            disabled=["Tableau Datasource"],
        )
        table_name_map = dict(zip(ds_names, edited["Power BI Table Name"]))

    # A fast, deterministic-only pass so RLS candidates can be reviewed
    # before the (potentially AI-assisted, slower) final Generate pass —
    # never auto-applied, per the build spec's Section 8.
    preview_semantic = build_semantic_model(twbx, table_name_map)
    confirmed_roles: List[Dict[str, Any]] = []
    if preview_semantic["rls_roles"]:
        with st.expander(
            f"RLS role candidates ({len(preview_semantic['rls_roles'])}) — review before including",
            expanded=True, icon=":material/shield_lock:",
        ):
            st.caption(
                "Detected from `ISMEMBEROF()` group-membership filters — each one names an AD/SSO "
                "security group and the dimension values it may see. Nothing here is applied "
                "automatically: only groups checked below are written into `roles.tmdl`."
            )
            for i, role in enumerate(preview_semantic["rls_roles"]):
                checked = st.checkbox(
                    f"**{role['ad_group']}** — secures {len(role['permissions'])} table(s)",
                    value=False, key=f"twbx_rls_{i}",
                )
                if checked:
                    confirmed_roles.append(role)

    use_ai_fallback = False
    if call_llm is not None:
        use_ai_fallback = st.checkbox(
            "Use AI to attempt formulas the deterministic translator can't confidently handle",
            value=False, key="twbx_ai_fallback", disabled=not llm_ready,
            help=("Only tried on formulas the deterministic translator already refused — nested "
                  "LOD, table calculations, and custom SQL are always refused regardless of this "
                  "setting, since they're out of scope by design, not by translation failure. "
                  "AI results are always tagged 'verify' in the report, never trusted the way a "
                  "deterministic translation is.") if llm_ready else
                 "Set an API key in the sidebar's AI assistant section to enable this.",
        )

    if st.button("⚙️ Generate PBIP Bundle", key="twbx_generate", width="stretch", icon=":material/settings:"):
        with st.spinner("Translating formulas, relationships, and layout…"):
            fallback_fn = None
            if use_ai_fallback and call_llm is not None:
                schema_text = _compact_schema_text(preview_semantic)
                fallback_fn = make_llm_formula_fallback(call_llm, schema_text)
            semantic = build_semantic_model(twbx, table_name_map, llm_fallback=fallback_fn)
            report = build_report_model(twbx, semantic)
            tmdl_files = emit_tmdl(semantic, catalog.strip() or "workspace", schema_name.strip() or "default")
            if confirmed_roles:
                tmdl_files["roles.tmdl"] = emit_roles_tmdl(confirmed_roles)
            pbir_files = emit_pbir(report)
            zip_bytes = package_pbip(twbx["workbook_name"], tmdl_files, pbir_files)
            combined_report = pd.concat(
                [semantic["report_df"], report["report_df"]], ignore_index=True,
            )
        st.session_state["twbx_result"] = {
            "semantic": semantic, "report": report, "tmdl_files": tmdl_files,
            "pbir_files": pbir_files, "zip_bytes": zip_bytes, "combined_report": combined_report,
            "confirmed_roles": confirmed_roles,
        }

    result = st.session_state.get("twbx_result")
    if result is None:
        return

    semantic, report = result["semantic"], result["report"]
    n_tables = len(semantic["tables"])
    n_rel = len(semantic["relationships"])
    n_meas = len(semantic["measures"])
    n_cols = len(semantic["calculated_columns"])
    n_refused = int((result["combined_report"]["Status"] == "Not translated").sum())
    n_pages = len(report["pages"])
    n_visuals = sum(len(p["visuals"]) for p in report["pages"])
    n_unsupported_visuals = int(
        (result["combined_report"][result["combined_report"]["Object Type"] == "Visual"]["Status"] == "Not migrated").sum()
    )

    m = st.columns(4)
    m[0].metric("Tables", n_tables)
    m[1].metric("Relationships", n_rel)
    m[2].metric("Measures translated", n_meas, help=f"{n_cols} calculated column(s) also translated.")
    m[3].metric("Formulas refused", n_refused)
    m2 = st.columns(3)
    m2[0].metric("Pages", n_pages)
    m2[1].metric("Visuals generated", n_visuals)
    m2[2].metric("Visuals unsupported", n_unsupported_visuals)

    if n_refused or n_unsupported_visuals:
        st.warning(
            f"{n_refused} formula(s) and {n_unsupported_visuals} visual(s) need manual work in "
            "Power BI Desktop — see the conversion report below for the specific reason on each.",
            icon="🟠",
        )
    else:
        st.success("Every calculated field and visual translated cleanly.", icon="✅")

    st.markdown("#### Conversion report")
    st.caption(
        "One row per calculated field, table, relationship, and visual this converter had an "
        "opinion about — `Info`-severity `Translated` rows are informational; everything else is "
        "worth a look before you trust the output."
    )
    _show_df_local(result["combined_report"], "TWBX Conversion Report", "twbx_report", height=380)

    with st.expander("Preview a generated file", expanded=False, icon=":material/data_object:"):
        preview_options = ["(none)"] + sorted(result["tmdl_files"].keys()) + sorted(
            f"Report/{k}" for k in result["pbir_files"].keys()
        )
        choice = st.selectbox("File", preview_options, key="twbx_preview_pick")
        if choice != "(none)":
            if choice.startswith("Report/"):
                content = json.dumps(result["pbir_files"][choice[len("Report/"):]], indent=2)
                st.code(content[:8000], language="json")
            else:
                st.code(result["tmdl_files"][choice][:8000], language="text")

    st.download_button(
        "⬇ Download PBIP Bundle (.zip)", result["zip_bytes"],
        file_name=f"{slug(twbx['workbook_name'], 'converted')}_pbip.zip",
        mime="application/zip", key="dl_twbx_pbip", width="stretch",
    )
    st.info(
        "**This is a Power BI Project folder, not a .pbix.** Unzip it, then in Power BI Desktop: "
        "File ➜ Open ➜ browse to the extracted `.Report` folder (or double-click the `.pbip` "
        "file if you have the Power BI Project file association installed). Desktop will load "
        "the semantic model and report together — confirm the M query source in each table's "
        "partition (the catalog/schema above is a placeholder), verify relationship cardinality, "
        "and review every row in the conversion report before publishing. This has not been "
        "opened in a real Power BI Desktop install as part of this build — treat it as a strong "
        "first draft, not a validated file.",
        icon="ℹ️",
    )
