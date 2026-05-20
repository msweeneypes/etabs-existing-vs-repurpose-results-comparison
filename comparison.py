"""
Comparison engine for ETABS Design Output Tables.

Parses Design Forces and Steel Frame Design Summary sheets from two Excel
exports and produces a member-level force and D/C ratio comparison table.

Performance: parsed DataFrames and comparison results are cached in
module-level dicts keyed on MD5 hash of file content, so re-running with
different thresholds or display filters is near-instant after the first parse.
"""
import hashlib
import io
import re
from collections import defaultdict

import pandas as pd
from openpyxl import load_workbook

from storage_cache import get_cached, set_cached

FORCE_COMPONENTS = ['P', 'V2', 'V3', 'T', 'M2', 'M3']

DESIGN_FORCES_SHEETS = {
    'Columns': 'Design Forces - Columns',
    'Beams':   'Design Forces - Beams',
    'Braces':  'Design Forces - Braces',
}

# Member-specific label column name in each design forces sheet
LABEL_COL = {
    'Columns': 'Column',
    'Beams':   'Beam',
    'Braces':  'Brace',
}

SUMMARY_SHEET_PREFIX = 'Stl Frm Sum'

# Design type string as it appears in the Design Summary sheet
DESIGN_TYPE = {
    'Columns': 'Column',
    'Beams':   'Beam',
    'Braces':  'Brace',
}

# Lateral load token: E or W followed by one or more letters (e.g. EQ, EQB, WA, WG).
# Lookbehind excludes letters only so "0.5WG" and "1EQB" (digit-adjacent) still match,
# while "Dead", "Lr", "SB" etc. don't start with E/W and are unaffected.
LATERAL_RE = re.compile(r'(?<![A-Za-z])([EW][A-Za-z]+)(?![A-Za-z])')


# Caching is handled by storage_cache (two-level: in-process + Viktor Storage).
# Prefixes: 'etabs_parse' for parsed workbook data, 'etabs_cmp' for comparison results.


# ---------------------------------------------------------------------------
# Low-level helpers
# ---------------------------------------------------------------------------

def _file_hash(file_obj) -> str:
    """Return MD5 hex digest of the file content."""
    return hashlib.md5(file_obj.getvalue_binary()).hexdigest()


def _stream_etabs_rows(ws):
    """
    Return (col_map, rows_iter) for an ETABS-formatted worksheet.

    ETABS layout: row 0 = title, row 1 = headers, row 2 = units, row 3+ = data.
    col_map maps column header name → 0-based column index.
    Raises StopIteration if the sheet has fewer than 3 rows.
    """
    rows_iter = ws.rows
    next(rows_iter)                      # row 0: table title
    header_cells = next(rows_iter)       # row 1: column headers
    next(rows_iter)                      # row 2: units — skip
    col_map = {
        cell.value: i
        for i, cell in enumerate(header_cells)
        if cell.value is not None
    }
    return col_map, rows_iter


# ---------------------------------------------------------------------------
# Workbook-level parsers (take an already-open read_only workbook)
# ---------------------------------------------------------------------------

def _parse_forces_from_wb(wb, member_type: str) -> pd.DataFrame:
    """
    Stream one Design Forces worksheet from an open workbook.

    Online max-abs aggregation: stores one float per (Story, Label, component).
    Peak memory is O(unique members), not O(rows), regardless of file size.
    """
    sheet_name    = DESIGN_FORCES_SHEETS[member_type]
    label_col_name = LABEL_COL[member_type]
    empty = pd.DataFrame(columns=['Story', 'Label'] + FORCE_COMPONENTS)

    if sheet_name not in wb.sheetnames:
        return empty

    try:
        col_map, rows_iter = _stream_etabs_rows(wb[sheet_name])
    except StopIteration:
        return empty

    story_i = col_map.get('Story')
    label_i = col_map.get(label_col_name)
    if story_i is None or label_i is None:
        return empty

    force_col_indices = {c: col_map.get(c) for c in FORCE_COMPONENTS}
    accum: dict = {}

    for row in rows_iter:
        vals = [cell.value for cell in row]
        n    = len(vals)
        story = vals[story_i] if story_i < n else None
        label = vals[label_i] if label_i < n else None
        if story is None or label is None:
            continue
        key = (story, label)
        if key not in accum:
            accum[key] = dict.fromkeys(FORCE_COMPONENTS, 0.0)
        for c, ci in force_col_indices.items():
            if ci is not None and ci < n and vals[ci] is not None:
                try:
                    v = float(vals[ci])
                    if abs(v) > abs(accum[key][c]):
                        accum[key][c] = v
                except (TypeError, ValueError):
                    pass

    if not accum:
        return empty
    return pd.DataFrame(
        [{'Story': s, 'Label': l, **forces} for (s, l), forces in accum.items()]
    )


def _parse_summary_from_wb(wb) -> pd.DataFrame:
    """
    Stream the Steel Frame Design Summary worksheet from an open workbook.

    Returns one row per member with design section and D/C ratio data.
    """
    empty = pd.DataFrame(columns=[
        'Story', 'Label', 'MemberType', 'DesignSection',
        'PMMCombo', 'PMMRatio', 'VMajCombo', 'VMajRatio',
    ])

    summary_sheet = next(
        (s for s in wb.sheetnames if s.startswith(SUMMARY_SHEET_PREFIX)), None
    )
    if summary_sheet is None:
        return empty

    try:
        col_map, rows_iter = _stream_etabs_rows(wb[summary_sheet])
    except StopIteration:
        return empty

    name_remap = {
        'Design Type':    'MemberType',
        'Design Section': 'DesignSection',
        'PMM Combo':      'PMMCombo',
        'PMM Ratio':      'PMMRatio',
        'V Major Combo':  'VMajCombo',
        'V Major Ratio':  'VMajRatio',
    }
    idx = {name_remap.get(k, k): i for k, i in col_map.items()}

    story_i = idx.get('Story')
    label_i = idx.get('Label')
    if story_i is None or label_i is None:
        return empty

    extra_fields = ('MemberType', 'DesignSection', 'PMMCombo', 'PMMRatio', 'VMajCombo', 'VMajRatio')
    rows_out = []

    for row in rows_iter:
        vals = [cell.value for cell in row]
        n    = len(vals)
        story = vals[story_i] if story_i < n else None
        label = vals[label_i] if label_i < n else None
        if story is None or label is None:
            continue
        row_dict = {'Story': story, 'Label': label}
        for field in extra_fields:
            fi = idx.get(field)
            row_dict[field] = vals[fi] if fi is not None and fi < n else None
        rows_out.append(row_dict)

    if not rows_out:
        return empty

    df = pd.DataFrame(rows_out)
    df = df[df['Label'].notna()].copy()

    for col in ('PMMCombo', 'VMajCombo'):
        if col in df.columns:
            df[col] = (
                df[col].astype(str)
                .str.replace(r'\s*\([CT]\)\s*$', '', regex=True)
                .str.strip()
            )
    for col in ('PMMRatio', 'VMajRatio'):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    return df


# ---------------------------------------------------------------------------
# Public parser wrappers (open their own workbook — used for standalone testing)
# ---------------------------------------------------------------------------

def parse_design_forces(file_source, member_type: str) -> pd.DataFrame:
    """Parse one Design Forces sheet. _get_parsed_file uses _parse_forces_from_wb instead."""
    try:
        if isinstance(file_source, (bytes, bytearray)):
            fh = io.BytesIO(file_source)
        else:
            with file_source.open_binary() as raw:
                fh = io.BytesIO(raw.read())
        wb = load_workbook(fh, read_only=True, data_only=True)
        try:
            return _parse_forces_from_wb(wb, member_type)
        finally:
            wb.close()
    except Exception:
        return pd.DataFrame(columns=['Story', 'Label'] + FORCE_COMPONENTS)


def parse_design_summary(file_source) -> pd.DataFrame:
    """Parse the Design Summary sheet. _get_parsed_file uses _parse_summary_from_wb instead."""
    try:
        if isinstance(file_source, (bytes, bytearray)):
            fh = io.BytesIO(file_source)
        else:
            with file_source.open_binary() as raw:
                fh = io.BytesIO(raw.read())
        wb = load_workbook(fh, read_only=True, data_only=True)
        try:
            return _parse_summary_from_wb(wb)
        finally:
            wb.close()
    except Exception:
        return pd.DataFrame()


# ---------------------------------------------------------------------------
# Cache helpers
# ---------------------------------------------------------------------------

def _get_parsed_file(file_obj) -> tuple:
    """
    Return (file_hash, parsed_dict) for a file, computing and caching if absent.

    Opens the workbook exactly once and reads all four sheets in a single pass.
    File bytes are released before parsing begins to minimise peak memory usage.

    parsed_dict = {'summary': df, 'forces': {'Columns': df, 'Beams': df, 'Braces': df}}
    """
    file_bytes = file_obj.getvalue_binary()
    h = hashlib.md5(file_bytes).hexdigest()

    cached = get_cached('etabs_parse', h)
    if cached is not None:
        return h, cached

    fh = io.BytesIO(file_bytes)
    del file_bytes
    try:
        wb = load_workbook(fh, read_only=True, data_only=True)
        try:
            result = {
                'summary': _parse_summary_from_wb(wb),
                'forces': {
                    mt: _parse_forces_from_wb(wb, mt)
                    for mt in ('Columns', 'Beams', 'Braces')
                },
                'sheet_names': list(wb.sheetnames),
            }
        finally:
            wb.close()
    except Exception:
        result = {
            'summary': pd.DataFrame(),
            'forces': {
                mt: pd.DataFrame(columns=['Story', 'Label'] + FORCE_COMPONENTS)
                for mt in ('Columns', 'Beams', 'Braces')
            },
            'sheet_names': [],
        }

    set_cached('etabs_parse', h, result)
    return h, result


# ---------------------------------------------------------------------------
# Classification
# ---------------------------------------------------------------------------

def classify_combo(combo_name: str) -> str:
    """
    Return 'lateral' if the combo name contains a seismic (E...) or wind (W...)
    load token; otherwise 'gravity'.

    ETABS internal design combo codes (DStlSXX) are treated as 'gravity'
    because the underlying load pattern is opaque.
    """
    if not combo_name or combo_name in ('nan', 'None', ''):
        return 'gravity'
    if re.match(r'^DStl', combo_name, re.IGNORECASE):
        return 'gravity'
    return 'lateral' if LATERAL_RE.search(combo_name) else 'gravity'


# ---------------------------------------------------------------------------
# % change computation
# ---------------------------------------------------------------------------

def _pct_change(exist_val: float, new_val: float):
    """
    Compute percent change in magnitude: (|new| - |exist|) / |exist| * 100.

    Returns (pct, sign_reversal) where pct is a float, 'INF', or 0.0.
    sign_reversal is True when both values are non-zero and have opposite signs.
    """
    near_zero = 1e-4
    e_abs = abs(exist_val)
    n_abs = abs(new_val)

    sign_rev = (e_abs >= near_zero and n_abs >= near_zero and exist_val * new_val < 0)

    if e_abs < near_zero and n_abs < near_zero:
        return 0.0, False
    if e_abs < near_zero:
        return 'INF', sign_rev

    pct = (n_abs - e_abs) / e_abs * 100.0
    return round(pct, 1), sign_rev


def _fmt(val, decimals=3):
    """Format a numeric value for display, or pass through non-numeric strings."""
    if isinstance(val, float) and not (val != val):  # not NaN
        return round(val, decimals)
    return val


# ---------------------------------------------------------------------------
# Fail reason helper
# ---------------------------------------------------------------------------

def _compute_fail_reason(row: dict, threshold: float, load_type: str,
                          any_sign_rev: bool, is_fail: bool) -> str:
    """
    Build a human-readable string describing the primary cause of a FAIL.

    Returns '' for passing or non-comparable rows.
    """
    if not is_fail:
        return ''

    # Priority 1: any force component jumped from ~zero (INF change)
    for c in FORCE_COMPONENTS:
        if row.get(f'{c}_Pct') == 'INF':
            return f'{c} INF ({load_type})'
    if row.get('PMM_Pct') == 'INF':
        return f'PMM INF ({load_type})'

    # Priority 2: largest % overage across forces and PMM
    worst_label = None
    worst_pct = 0.0
    for c in FORCE_COMPONENTS:
        pct = row.get(f'{c}_Pct')
        if isinstance(pct, float) and pct > threshold and pct > worst_pct:
            worst_pct = pct
            worst_label = c
    pmm_pct = row.get('PMM_Pct')
    if isinstance(pmm_pct, float) and pmm_pct > threshold and pmm_pct > worst_pct:
        worst_pct = pmm_pct
        worst_label = 'PMM'

    if worst_label is not None:
        sign_tag = ' [sign rev]' if any_sign_rev else ''
        return f'{worst_label} +{worst_pct:.1f}% > {threshold:.0f}% {load_type}{sign_tag}'

    return 'See detail'


# ---------------------------------------------------------------------------
# Internal comparison (no file I/O — works from pre-parsed dicts)
# ---------------------------------------------------------------------------

def _run_comparison_internal(
    parsed_exist: dict,
    parsed_new: dict,
    member_type_filter: str,
    gravity_threshold: float,
    lateral_threshold: float,
) -> list:
    """
    Core comparison logic. Operates on pre-parsed DataFrames.
    Returns the full unfiltered result list.
    """
    types_to_process = (
        ['Columns', 'Beams', 'Braces']
        if member_type_filter == 'All'
        else [member_type_filter]
    )

    sum_exist = parsed_exist['summary']
    sum_new   = parsed_new['summary']

    all_results = []

    for mtype in types_to_process:
        design_type = DESIGN_TYPE[mtype]

        ds_exist = (
            sum_exist[sum_exist['MemberType'] == design_type].copy()
            if 'MemberType' in sum_exist.columns else pd.DataFrame()
        )
        ds_new = (
            sum_new[sum_new['MemberType'] == design_type].copy()
            if 'MemberType' in sum_new.columns else pd.DataFrame()
        )

        forces_exist = parsed_exist['forces'][mtype]
        forces_new   = parsed_new['forces'][mtype]

        def _force_member_set(df):
            if df.empty or 'Story' not in df.columns:
                return set()
            return set(zip(df['Story'], df['Label']))

        exist_members = _force_member_set(forces_exist)
        new_members   = _force_member_set(forces_new)
        all_members   = exist_members | new_members

        def _build_lookup(ds):
            out = {}
            if ds.empty:
                return out
            for _, r in ds.iterrows():
                out[(r['Story'], r['Label'])] = r
            return out

        exist_lookup = _build_lookup(ds_exist)
        new_lookup   = _build_lookup(ds_new)

        def _build_force_lookup(df):
            out = {}
            if df.empty:
                return out
            for _, r in df.iterrows():
                out[(r['Story'], r['Label'])] = r
            return out

        fe_lookup = _build_force_lookup(forces_exist)
        fn_lookup = _build_force_lookup(forces_new)

        for story, label in sorted(all_members, key=lambda x: (str(x[0]), str(x[1]))):
            in_exist = (story, label) in exist_members
            in_new   = (story, label) in new_members

            es = exist_lookup.get((story, label))   # summary row — may be None
            ns = new_lookup.get((story, label))     # summary row — may be None
            ef = fe_lookup.get((story, label))
            nf = fn_lookup.get((story, label))

            row = {
                'Story':      story,
                'Label':      label,
                'MemberType': design_type,
            }

            # ---- REMOVED ----
            if not in_new:
                row.update({
                    'DesignSection_Exist': es.get('DesignSection', '') if es is not None else '',
                    'DesignSection_New':   'N/A',
                    'GovCombo_Exist': es.get('PMMCombo', '') if es is not None else '',
                    'GovCombo_New':   'N/A',
                    'LoadType': 'N/A',
                    'PMM_Exist': _fmt(es.get('PMMRatio')) if es is not None else 'N/A',
                    'PMM_New': 'N/A', 'PMM_Pct': 'N/A',
                    'VMaj_Exist': _fmt(es.get('VMajRatio')) if es is not None else 'N/A',
                    'VMaj_New': 'N/A', 'VMaj_Pct': 'N/A',
                    'SignReversal': '',
                    'FailReason': '',
                    'Pass': 'REMOVED',
                })
                for c in FORCE_COMPONENTS:
                    row[f'{c}_Exist'] = _fmt(float(ef[c]), 2) if ef is not None else 'N/A'
                    row[f'{c}_New']   = 'N/A'
                    row[f'{c}_Pct']   = 'N/A'
                all_results.append(row)
                continue

            # ---- ADDED ----
            if not in_exist:
                row.update({
                    'DesignSection_Exist': 'N/A',
                    'DesignSection_New':   ns.get('DesignSection', '') if ns is not None else '',
                    'GovCombo_Exist': 'N/A',
                    'GovCombo_New': ns.get('PMMCombo', '') if ns is not None else '',
                    'LoadType': classify_combo(ns.get('PMMCombo', '') if ns is not None else ''),
                    'PMM_Exist': 'N/A',
                    'PMM_New': _fmt(ns.get('PMMRatio')) if ns is not None else 'N/A',
                    'PMM_Pct': 'N/A',
                    'VMaj_Exist': 'N/A',
                    'VMaj_New': _fmt(ns.get('VMajRatio')) if ns is not None else 'N/A',
                    'VMaj_Pct': 'N/A',
                    'SignReversal': '',
                    'FailReason': '',
                    'Pass': 'ADDED',
                })
                for c in FORCE_COMPONENTS:
                    row[f'{c}_Exist'] = 'N/A'
                    row[f'{c}_New']   = _fmt(float(nf[c]), 2) if nf is not None else 'N/A'
                    row[f'{c}_Pct']   = 'N/A'
                all_results.append(row)
                continue

            # ---- MATCHED ----
            e_sec       = es.get('DesignSection', '') if es is not None else ''
            n_sec       = ns.get('DesignSection', '') if ns is not None else ''
            e_pmm_combo = es.get('PMMCombo', '')      if es is not None else ''
            n_pmm_combo = ns.get('PMMCombo', '')      if ns is not None else ''
            e_pmm  = float(es['PMMRatio'])  if es is not None and pd.notna(es.get('PMMRatio'))  else None
            n_pmm  = float(ns['PMMRatio'])  if ns is not None and pd.notna(ns.get('PMMRatio'))  else None
            e_vmaj = float(es['VMajRatio']) if es is not None and pd.notna(es.get('VMajRatio')) else None
            n_vmaj = float(ns['VMajRatio']) if ns is not None and pd.notna(ns.get('VMajRatio')) else None

            load_type = classify_combo(n_pmm_combo)
            threshold = lateral_threshold if load_type == 'lateral' else gravity_threshold

            row['DesignSection_Exist'] = e_sec
            row['DesignSection_New']   = n_sec
            row['GovCombo_Exist']      = e_pmm_combo
            row['GovCombo_New']        = n_pmm_combo
            row['LoadType']            = load_type

            # D/C ratios
            if e_pmm is not None and n_pmm is not None:
                pmm_pct, _ = _pct_change(e_pmm, n_pmm)
            else:
                pmm_pct = 'N/A'
            if e_vmaj is not None and n_vmaj is not None:
                vmaj_pct, _ = _pct_change(e_vmaj, n_vmaj)
            else:
                vmaj_pct = 'N/A'

            row['PMM_Exist']  = _fmt(e_pmm)  if e_pmm  is not None else 'N/A'
            row['PMM_New']    = _fmt(n_pmm)  if n_pmm  is not None else 'N/A'
            row['PMM_Pct']    = pmm_pct
            row['VMaj_Exist'] = _fmt(e_vmaj) if e_vmaj is not None else 'N/A'
            row['VMaj_New']   = _fmt(n_vmaj) if n_vmaj is not None else 'N/A'
            row['VMaj_Pct']   = vmaj_pct

            any_sign_rev = False
            is_fail = False

            for c in FORCE_COMPONENTS:
                e_val = _fmt(float(ef[c]), 2) if ef is not None else 'No Data'
                n_val = _fmt(float(nf[c]), 2) if nf is not None else 'No Data'

                row[f'{c}_Exist'] = e_val
                row[f'{c}_New']   = n_val

                if isinstance(e_val, (int, float)) and isinstance(n_val, (int, float)):
                    pct, sign_rev = _pct_change(float(e_val), float(n_val))
                    row[f'{c}_Pct'] = pct
                    if sign_rev:
                        any_sign_rev = True
                    if pct == 'INF' or (isinstance(pct, float) and pct > threshold):
                        is_fail = True
                else:
                    row[f'{c}_Pct'] = 'N/A'

            # D/C threshold check
            if pmm_pct == 'INF' or (isinstance(pmm_pct, float) and pmm_pct > threshold):
                is_fail = True

            row['SignReversal'] = 'YES' if any_sign_rev else ''
            row['FailReason']   = _compute_fail_reason(row, threshold, load_type, any_sign_rev, is_fail)
            row['Pass']         = 'FLAG' if is_fail else 'PASS'

            all_results.append(row)

    return all_results


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def run_comparison(
    existing_file,
    modified_file,
    member_type_filter: str,
    gravity_threshold: float,
    lateral_threshold: float,
    show_failures_only: bool,
) -> list:
    """
    Compare ETABS design output between two models.

    Returns a list of row dicts — one per matched (or unmatched) member.
    Results are cached by file hash + comparison parameters; only file parsing
    (on first upload) is slow.

    Keys per row: Story, Label, MemberType, DesignSection_Exist/New,
    GovCombo_Exist/New, LoadType, P/V2/V3/T/M2/M3_Exist/New/Pct,
    PMM_Exist/New/Pct, VMaj_Exist/New/Pct, SignReversal, FailReason, Pass
    """
    exist_hash, parsed_exist = _get_parsed_file(existing_file)
    new_hash,   parsed_new   = _get_parsed_file(modified_file)
    cache_key  = (exist_hash, new_hash, member_type_filter,
                  gravity_threshold, lateral_threshold)

    all_results = get_cached('etabs_cmp', cache_key)
    if all_results is None:
        all_results = _run_comparison_internal(
            parsed_exist, parsed_new,
            member_type_filter, gravity_threshold, lateral_threshold,
        )
        set_cached('etabs_cmp', cache_key, all_results)

    if show_failures_only:
        return [r for r in all_results if r.get('Pass') == 'FLAG']
    return list(all_results)


def build_summary(results: list) -> list:
    """
    Group results by MemberType and Story; count PASS/FAIL/ADDED/REMOVED.
    Returns list of row dicts: Story, MemberType, Total, PASS, FAIL, ADDED, REMOVED.
    """
    counts = defaultdict(lambda: {'Total': 0, 'PASS': 0, 'WARN': 0, 'FLAG': 0, 'ADDED': 0, 'REMOVED': 0})

    for r in results:
        key = (r.get('Story', ''), r.get('MemberType', ''))
        counts[key]['Total'] += 1
        status = r.get('Pass', '')
        if status in ('PASS', 'WARN', 'FLAG', 'ADDED', 'REMOVED'):
            counts[key][status] += 1

    summary = []
    for (story, mtype), c in sorted(counts.items(), key=lambda x: (str(x[0][0]), str(x[0][1]))):
        summary.append({
            'Story':      story,
            'MemberType': mtype,
            'Total':      c['Total'],
            'PASS':       c['PASS'],
            'WARN':       c['WARN'],
            'FLAG':       c['FLAG'],
            'ADDED':      c['ADDED'],
            'REMOVED':    c['REMOVED'],
        })
    return summary


def results_to_csv(results: list) -> str:
    """Convert results list-of-dicts to a CSV string."""
    if not results:
        return 'No results\n'
    return pd.DataFrame(results).to_csv(index=False)
