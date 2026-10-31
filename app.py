import streamlit as st
import json
import io
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

st.set_page_config(page_title="JSON Comparison Tool (Excel)", layout="wide")

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

def load_config():
    with open("config.json") as f:
        return json.load(f)

CFG = load_config()
COL = CFG["columns"]
EXCLUDE_FIELDS = set(CFG.get("exclude_fields", []))
UNWRAP_KEYS    = CFG.get("unwrap_keys", [])

EXCEL_CELL_LIMIT = 32767

# ---------------------------------------------------------------------------
# JSON diff engine (same as the DB version)
# ---------------------------------------------------------------------------

def _fmt(val):
    if val is None:
        return ""
    if isinstance(val, (dict, list)):
        return json.dumps(val)
    if isinstance(val, bool):
        return "true" if val else "false"
    return str(val)

def _is_empty(obj):
    if isinstance(obj, (dict, list)):
        return len(obj) == 0
    return False

def strip_excluded(obj, exclude=EXCLUDE_FIELDS):
    """Recursively remove excluded fields and any empty dicts/lists they leave behind."""
    if isinstance(obj, dict):
        result = {}
        for k, v in obj.items():
            if k in exclude:
                continue
            cleaned = strip_excluded(v, exclude)
            if not _is_empty(cleaned):
                result[k] = cleaned
        return result
    elif isinstance(obj, list):
        cleaned = [strip_excluded(item, exclude) for item in obj]
        return [item for item in cleaned if not _is_empty(item)]
    return obj

def _normalize(val):
    """Normalize scalar values for comparison:
    - Collapse all internal whitespace (spaces, newlines, tabs) to single space
    - Strip leading/trailing whitespace
    - Treat boolean-string equivalents as equal (false == "false")
    - Case-insensitive string comparison
    """
    if isinstance(val, bool):
        return str(val).lower()
    if isinstance(val, str):
        v = " ".join(val.split()).lower()
        if v in ("true", "false"):
            return v
        return v
    return val

def deep_diff(a, b, path=""):
    diffs = []
    # Normalize scalars for comparison but keep originals for display
    if not isinstance(a, (dict, list)) and not isinstance(b, (dict, list)):
        if _normalize(a) != _normalize(b):
            diffs.append({"path": path or "root", "old": _fmt(a), "new": _fmt(b), "type": "changed"})
        return diffs
    if type(a) != type(b):
        diffs.append({"path": path or "root", "old": _fmt(a), "new": _fmt(b), "type": "changed"})
        return diffs
    if isinstance(a, dict):
        for key in set(a) | set(b):
            p = f"{path}.{key}" if path else key
            if key not in a:
                diffs.append({"path": p, "old": "", "new": _fmt(b[key]), "type": "added"})
            elif key not in b:
                diffs.append({"path": p, "old": _fmt(a[key]), "new": "", "type": "removed"})
            else:
                diffs.extend(deep_diff(a[key], b[key], p))
    elif isinstance(a, list):
        for i in range(max(len(a), len(b))):
            p = f"{path}[{i}]"
            if i >= len(a):
                diffs.append({"path": p, "old": "", "new": _fmt(b[i]), "type": "added"})
            elif i >= len(b):
                diffs.append({"path": p, "old": _fmt(a[i]), "new": "", "type": "removed"})
            else:
                diffs.extend(deep_diff(a[i], b[i], p))
    else:
        if _normalize(a) != _normalize(b):
            diffs.append({"path": path or "root", "old": _fmt(a), "new": _fmt(b), "type": "changed"})
    return diffs

def parse_json(val):
    """Returns (parsed_value, error_message). Empty cells return (None, None)."""
    if val is None:
        return None, None
    if isinstance(val, str):
        if not val.strip():
            return None, None
        try:
            return json.loads(val), None
        except Exception as e:
            return val, f"Invalid JSON: {e}"
    return val, None

# ---------------------------------------------------------------------------
# Excel input
# ---------------------------------------------------------------------------

def read_input_excel(file):
    """Read the uploaded workbook (first sheet). Header matching is case-insensitive."""
    wb = openpyxl.load_workbook(file, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = ws.iter_rows(values_only=True)

    header = next(rows, None)
    if not header:
        raise ValueError("The Excel file is empty.")
    col_index = {str(h).strip().lower(): i for i, h in enumerate(header) if h is not None}

    primary_id = COL["primary_id"]
    required = [primary_id] + [c for s in COL["json_sets"] for c in (s["left"], s["right"])]
    missing = [c for c in required if c.lower() not in col_index]
    if missing:
        raise ValueError(
            f"Missing column(s): {', '.join(missing)}. "
            f"Found: {', '.join(str(h) for h in header if h is not None)}"
        )

    records = []
    for row_num, row in enumerate(rows, start=2):
        if row is None or all(v is None or str(v).strip() == "" for v in row):
            continue
        row_map = {c: (row[col_index[c.lower()]] if col_index[c.lower()] < len(row) else None) for c in required}
        row_map["_row"] = row_num
        records.append(row_map)
    wb.close()
    return records

# ---------------------------------------------------------------------------
# Compare
# ---------------------------------------------------------------------------

def run_comparison(records):
    json_sets  = COL["json_sets"]
    primary_id = COL["primary_id"]

    results = []
    for row_map in records:
        ref = row_map.get(primary_id)
        record = {
            "id":   "" if ref is None else str(ref),
            "row":  row_map["_row"],
            "sets": [],
        }
        for s in json_sets:
            left,  left_err  = parse_json(row_map.get(s["left"]))
            right, right_err = parse_json(row_map.get(s["right"]))
            errors = [f"{s.get('left_label', s['left'])}: {left_err}"] if left_err else []
            if right_err:
                errors.append(f"{s.get('right_label', s['right'])}: {right_err}")

            if errors:
                diffs, status = [], "error"
            elif left is None and right is None:
                diffs, status = [], "empty"
            elif left is None or right is None:
                diffs  = [{"path": "root", "old": _fmt(left), "new": _fmt(right),
                           "type": "added" if left is None else "removed"}]
                status = "diff"
            else:
                for uw in UNWRAP_KEYS:
                    if isinstance(left, dict) and uw in left:
                        left = left[uw]
                    if isinstance(right, dict) and uw in right:
                        right = right[uw]
                diffs  = deep_diff(strip_excluded(left), strip_excluded(right))
                status = "diff" if diffs else "same"

            record["sets"].append({
                "label":       s["label"],
                "left_col":    s["left"],
                "right_col":   s["right"],
                "left_label":  s.get("left_label", s["left"]),
                "right_label": s.get("right_label", s["right"]),
                "diffs":       diffs,
                "has_diff":    len(diffs) > 0,
                "status":      status,
                "errors":      errors,
            })
        results.append(record)

    return results

# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

def _xl(val):
    """Make a value safe for an Excel cell (no control chars, within the cell size limit)."""
    s = ILLEGAL_CHARACTERS_RE.sub("", str(val or ""))
    if len(s) > EXCEL_CELL_LIMIT:
        s = s[:EXCEL_CELL_LIMIT - 20] + " ...[truncated]"
    return s

def build_excel(results):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "JSON Diff"

    hdr_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    hdr_font = Font(color="FFFFFF", bold=True)
    fills    = {
        "added":   PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
        "removed": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
        "changed": PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
        "error":   PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"),
    }

    headers = [COL["primary_id"], "Source Row", "Comparison Set", "Left Column", "Right Column",
               "Path", "Difference Type", "Old Value", "New Value"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = hdr_fill
        cell.font = hdr_font

    for rec in results:
        base = [_xl(rec["id"]), rec["row"]]
        for s in rec["sets"]:
            cols = [s["label"], s["left_label"], s["right_label"]]
            if s["status"] == "error":
                ws.append(base + cols + ["(invalid JSON)", "error", _xl("; ".join(s["errors"])), ""])
                for cell in ws[ws.max_row]:
                    cell.fill = fills["error"]
                continue
            if s["status"] == "empty":
                ws.append(base + cols + ["(both empty)", "", "", ""])
                continue
            if not s["diffs"]:
                ws.append(base + cols + ["(no differences)", "", "", ""])
            for d in s["diffs"]:
                ws.append(base + cols + [_xl(d["path"]), d["type"], _xl(d["old"]), _xl(d["new"])])
                if d["type"] in fills:
                    for cell in ws[ws.max_row]:
                        cell.fill = fills[d["type"]]

    for col_cells in ws.columns:
        width = max(len(str(c.value or "")) for c in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(width + 4, 60)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

def build_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Input"
    cols = [COL["primary_id"]] + [c for s in COL["json_sets"] for c in (s["left"], s["right"])]
    ws.append(cols)
    ws.append(["REF-001", '{"name": "A", "qty": 1}', '{"name": "A", "qty": 2}',
               '{"status": "open"}', '{"status": "open"}'])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ---------------------------------------------------------------------------
# UI helpers
# ---------------------------------------------------------------------------

DIFF_COLORS = {"added": "#c6efce", "removed": "#ffc7ce", "changed": "#ffeb9c"}
DIFF_TEXT   = {"added": "#276221", "removed": "#8b0000",  "changed": "#7d5800"}

def _esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def diff_badge(dtype):
    bg = DIFF_COLORS.get(dtype, "#eee")
    fg = DIFF_TEXT.get(dtype, "#000")
    return f'<span style="background:{bg};color:{fg};padding:1px 7px;border-radius:3px;font-size:.75rem;font-weight:700">{dtype.upper()}</span>'

def render_diffs(diffs):
    for d in diffs:
        bg = DIFF_COLORS.get(d["type"], "#fafafa")
        border = {"added": "#4caf50", "removed": "#f44336", "changed": "#ff9800"}.get(d["type"], "#ccc")
        old_html = f'<span style="background:#ffd7d7;color:#000;padding:1px 6px;border-radius:3px;font-family:monospace;font-size:.8rem">{_esc(d["old"])}</span>' if d["old"] else ""
        new_html = f'<span style="background:#d7ffd7;color:#000;padding:1px 6px;border-radius:3px;font-family:monospace;font-size:.8rem">{_esc(d["new"])}</span>' if d["new"] else ""
        arrow    = " &rarr; " if old_html and new_html else ""
        st.markdown(
            f'<div style="background:{bg};color:#000;border-left:3px solid {border};padding:6px 10px;'
            f'border-radius:4px;margin-bottom:4px">'
            f'{diff_badge(d["type"])} &nbsp;<code style="font-size:.8rem">{_esc(d["path"])}</code><br/>'
            f'<div style="margin-top:4px">{old_html}{arrow}{new_html}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

# ---------------------------------------------------------------------------
# Main app
# ---------------------------------------------------------------------------

st.title("JSON Comparison Tool (Excel)")

_input_cols = [COL["primary_id"]] + [c for s in COL["json_sets"] for c in (s["left"], s["right"])]

# ── Sidebar ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Input Excel")
    uploaded = st.file_uploader("Upload .xlsx file", type=["xlsx", "xlsm"])
    st.caption("Expected columns: " + ", ".join(f"`{c}`" for c in _input_cols))
    st.download_button(
        "Download sample template",
        data=build_template(),
        file_name="json_compare_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.divider()
    st.header("Comparison Sets")
    for s in COL["json_sets"]:
        st.markdown(f"**{s['label']}**: `{s['left']}` vs `{s['right']}`")

    run_btn = st.button("Run Comparison", type="primary", use_container_width=True)

# ── Run comparison ─────────────────────────────────────────────────────────
if run_btn:
    if uploaded is None:
        st.error("Please upload an Excel file.")
    else:
        with st.spinner("Comparing..."):
            try:
                records = read_input_excel(uploaded)
                st.session_state["results"]   = run_comparison(records)
                st.session_state["excel_buf"] = build_excel(st.session_state["results"])
                st.session_state["source"]    = uploaded.name
            except Exception as e:
                st.error(f"Comparison failed: {e}")
                st.session_state.pop("results", None)

# ── Results ────────────────────────────────────────────────────────────────
if "results" in st.session_state:
    results = st.session_state["results"]

    total      = len(results)
    diff_count = sum(1 for r in results if any(s["has_diff"] for s in r["sets"]))
    err_count  = sum(1 for r in results if any(s["status"] == "error" for s in r["sets"]))

    st.caption(f"Source: {st.session_state.get('source', '')}")

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total Records", total)
    m2.metric("Records with Differences", diff_count)
    m3.metric("Identical Records", total - diff_count - err_count)
    m4.metric("Invalid JSON", err_count)
    with m5:
        fname = f"json_diff_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        m5.download_button(
            label="Export to Excel",
            data=st.session_state["excel_buf"],
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # Filter
    c1, c2 = st.columns([3, 1])
    q = c1.text_input("Filter by reference ID, path, or value", placeholder="Search...")
    only_diffs = c2.checkbox("Only show differences", value=False)
    if only_diffs:
        results = [r for r in results if any(s["has_diff"] or s["status"] == "error" for s in r["sets"])]
    if q:
        ql = q.lower()
        results = [
            r for r in results
            if ql in r["id"].lower()
            or any(
                ql in d["path"].lower() or ql in d["old"].lower() or ql in d["new"].lower()
                for s in r["sets"] for d in s["diffs"]
            )
        ]
    if q or only_diffs:
        st.caption(f"{len(results)} matching records")

    st.divider()

    for rec in results:
        has_any_diff = any(s["has_diff"] for s in rec["sets"])
        has_error    = any(s["status"] == "error" for s in rec["sets"])
        status_icon  = "⚠" if has_any_diff else ("✖" if has_error else "✓")

        with st.expander(
            f"{status_icon} **{COL['primary_id']}: {rec['id']}** &nbsp;|&nbsp; row {rec['row']}",
            expanded=has_any_diff or has_error,
        ):
            for s in rec["sets"]:
                st.markdown(
                    f"**{s['label']}** &nbsp; `{s['left_label']}` vs `{s['right_label']}`",
                    unsafe_allow_html=True,
                )
                if s["status"] == "error":
                    for err in s["errors"]:
                        st.error(err)
                elif s["status"] == "empty":
                    st.info("Both columns empty")
                elif not s["has_diff"]:
                    st.success("No differences")
                else:
                    render_diffs(s["diffs"])
                st.markdown("---")
