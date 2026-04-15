"""
Comparison engine for ETABS Design Output Tables.

Parses Design Forces and Steel Frame Design Summary sheets from two Excel
exports and produces a member-level force and D/C ratio comparison table.
"""
import re

import pandas as pd

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

SUMMARY_SHEET = 'Stl Frm Sum - AISC 360-16'

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


# ---------------------------------------------------------------------------
# Low-level helpers
# ---------------------------------------------------------------------------

def _read_etabs_sheet(file_obj, sheet_name: str) -> pd.DataFrame:
    """
    Read one ETABS-exported sheet.

    ETABS layout: row 0 = table title, row 1 = column headers,
    row 2 = units row, row 3+ = data.
    Returns an empty DataFrame on any error.
    """
    try:
        with file_obj.open_binary() as f:
            df = pd.read_excel(
                f,
                sheet_name=sheet_name,
                header=1,      # row 1 = headers
                skiprows=[2],  # row 2 = units — drop
                engine='openpyxl',
            )
        return df
    except Exception:
        return pd.DataFrame()


def _max_abs_val(series: pd.Series) -> float:
    """Return the value with the largest absolute magnitude, preserving sign."""
    s = series.dropna()
    if s.empty:
        return 0.0
    return float(s.iloc[s.abs().argmax()])


# ---------------------------------------------------------------------------
# Parsers
# ---------------------------------------------------------------------------

def parse_design_forces(file_obj, member_type: str) -> pd.DataFrame:
    """
    Parse one Design Forces sheet and aggregate to one row per member.

    For each member (Story + Label) the force value stored is the one with
    the largest absolute magnitude across all design combos and stations.
    Returns empty DataFrame if the sheet is missing.
    """
    sheet = DESIGN_FORCES_SHEETS[member_type]
    label_col = LABEL_COL[member_type]

    df = _read_etabs_sheet(file_obj, sheet)
    empty = pd.DataFrame(columns=['Story', 'Label'] + FORCE_COMPONENTS)

    if df.empty or label_col not in df.columns:
        return empty

    df = df.rename(columns={label_col: 'Label'})
    df = df[df['Label'].notna()].copy()

    if df.empty:
        return empty

    for col in FORCE_COMPONENTS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
        else:
            df[col] = 0.0

    result = (
        df.groupby(['Story', 'Label'], sort=False)[FORCE_COMPONENTS]
        .agg(_max_abs_val)
        .reset_index()
    )
    return result


def parse_design_summary(file_obj) -> pd.DataFrame:
    """
    Parse the Steel Frame Design Summary sheet.

    Returns one row per member with design section and D/C ratio data.
    """
    df = _read_etabs_sheet(file_obj, SUMMARY_SHEET)
    empty = pd.DataFrame(columns=[
        'Story', 'Label', 'MemberType', 'DesignSection',
        'PMMCombo', 'PMMRatio', 'VMajCombo', 'VMajRatio',
    ])

    if df.empty:
        return empty

    rename = {
        'Design Type':    'MemberType',
        'Design Section': 'DesignSection',
        'PMM Combo':      'PMMCombo',
        'PMM Ratio':      'PMMRatio',
        'V Major Combo':  'VMajCombo',
        'V Major Ratio':  'VMajRatio',
    }
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})

    needed = ['Story', 'Label', 'MemberType', 'DesignSection',
              'PMMCombo', 'PMMRatio', 'VMajCombo', 'VMajRatio']
    df = df[[c for c in needed if c in df.columns]].copy()
    df = df[df['Label'].notna()].copy()

    # Strip the "(C)" / "(T)" compression/tension suffix from combo names
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
# Main comparison
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
    Keys: Story, Label, MemberType, DesignSection_Exist, DesignSection_New,
          GovCombo_Exist, GovCombo_New, LoadType,
          P_Exist … M3_New … M3_Pct,
          PMM_Exist, PMM_New, PMM_Pct,
          VMaj_Exist, VMaj_New, VMaj_Pct,
          SignReversal, Pass
    """
    types_to_process = (
        ['Columns', 'Beams', 'Braces']
        if member_type_filter == 'All'
        else [member_type_filter]
    )

    # Parse design summaries once (covers all member types)
    sum_exist = parse_design_summary(existing_file)
    sum_new   = parse_design_summary(modified_file)

    all_results = []

    for mtype in types_to_process:
        design_type = DESIGN_TYPE[mtype]

        # Filter summary to this member type
        ds_exist = sum_exist[sum_exist['MemberType'] == design_type].copy() if 'MemberType' in sum_exist.columns else pd.DataFrame()
        ds_new   = sum_new[sum_new['MemberType'] == design_type].copy()   if 'MemberType' in sum_new.columns   else pd.DataFrame()

        # Parse forces
        forces_exist = parse_design_forces(existing_file, mtype)
        forces_new   = parse_design_forces(modified_file, mtype)

        # Build member sets from design summary (authoritative for membership)
        def _member_set(ds):
            if ds.empty or 'Story' not in ds.columns:
                return set()
            return set(zip(ds['Story'], ds['Label']))

        exist_members = _member_set(ds_exist)
        new_members   = _member_set(ds_new)
        all_members   = exist_members | new_members

        # Build lookup dicts: (story, label) → summary row
        def _build_lookup(ds):
            out = {}
            if ds.empty:
                return out
            for _, r in ds.iterrows():
                out[(r['Story'], r['Label'])] = r
            return out

        exist_lookup = _build_lookup(ds_exist)
        new_lookup   = _build_lookup(ds_new)

        # Build force lookup: (story, label) → forces row
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

            es = exist_lookup.get((story, label))
            ns = new_lookup.get((story, label))
            ef = fe_lookup.get((story, label))
            nf = fn_lookup.get((story, label))

            row = {
                'Story':      story,
                'Label':      label,
                'MemberType': design_type,
            }

            # ---- REMOVED: exists only in existing model ----
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
                    'Pass': 'REMOVED',
                })
                for c in FORCE_COMPONENTS:
                    row[f'{c}_Exist'] = _fmt(float(ef[c]), 2) if ef is not None else 'N/A'
                    row[f'{c}_New']   = 'N/A'
                    row[f'{c}_Pct']   = 'N/A'
                all_results.append(row)
                continue

            # ---- ADDED: exists only in new model ----
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
                    'Pass': 'ADDED',
                })
                for c in FORCE_COMPONENTS:
                    row[f'{c}_Exist'] = 'N/A'
                    row[f'{c}_New']   = _fmt(float(nf[c]), 2) if nf is not None else 'N/A'
                    row[f'{c}_Pct']   = 'N/A'
                all_results.append(row)
                continue

            # ---- MATCHED: in both models ----
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

            # Forces — flag if no design forces data
            if ef is None:
                row['_no_exist_forces'] = True
            if nf is None:
                row['_no_new_forces'] = True

            any_sign_rev = False
            is_fail = False

            for c in FORCE_COMPONENTS:
                if ef is None:
                    e_val = 'No Data'
                else:
                    e_val = _fmt(float(ef[c]), 2)

                if nf is None:
                    n_val = 'No Data'
                else:
                    n_val = _fmt(float(nf[c]), 2)

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

            # Also check D/C increase
            if pmm_pct == 'INF' or (isinstance(pmm_pct, float) and pmm_pct > threshold):
                is_fail = True

            row['SignReversal'] = 'YES' if any_sign_rev else ''
            row['Pass'] = 'FAIL' if is_fail else 'PASS'

            # Clean up internal flags
            row.pop('_no_exist_forces', None)
            row.pop('_no_new_forces', None)

            if show_failures_only and not is_fail:
                continue

            all_results.append(row)

    return all_results


def build_summary(results: list) -> list:
    """
    Group results by MemberType and Story; count PASS/FAIL/ADDED/REMOVED.
    Returns list of row dicts with keys: Story, MemberType, Total, PASS, FAIL, ADDED, REMOVED.
    """
    from collections import defaultdict
    counts = defaultdict(lambda: {'Total': 0, 'PASS': 0, 'FAIL': 0, 'ADDED': 0, 'REMOVED': 0})

    for r in results:
        key = (r.get('Story', ''), r.get('MemberType', ''))
        counts[key]['Total'] += 1
        status = r.get('Pass', '')
        if status in ('PASS', 'FAIL', 'ADDED', 'REMOVED'):
            counts[key][status] += 1

    summary = []
    for (story, mtype), c in sorted(counts.items(), key=lambda x: (str(x[0][0]), str(x[0][1]))):
        summary.append({
            'Story':      story,
            'MemberType': mtype,
            'Total':      c['Total'],
            'PASS':       c['PASS'],
            'FAIL':       c['FAIL'],
            'ADDED':      c['ADDED'],
            'REMOVED':    c['REMOVED'],
        })
    return summary


def results_to_csv(results: list) -> str:
    """Convert results list-of-dicts to a CSV string."""
    if not results:
        return 'No results\n'
    return pd.DataFrame(results).to_csv(index=False)
