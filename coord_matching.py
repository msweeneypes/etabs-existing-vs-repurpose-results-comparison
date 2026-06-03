"""
Coordinate-based label remapping for ETABS models.

Parses 'Objects and Elements - Joints' and 'Objects and Elements - Frames'
sheets from two workbooks, builds a {new_label: existing_label} dict keyed
on 3D endpoint geometry, and exposes a helper to apply that map to parsed
force DataFrames before the standard (Story, Label) comparison runs.

This module has no imports from comparison.py or analysis_comparison.py
to avoid circular dependencies; _stream_etabs_rows is re-implemented inline.
"""
import copy

JOINTS_SHEET = 'Objects and Elements - Joints'
FRAMES_SHEET = 'Objects and Elements - Frames'
COORD_TOLERANCE = 0.1  # inches


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _stream_rows(ws):
    """
    Yield (col_map, rows_iter) for an ETABS-formatted worksheet.

    ETABS layout: row 0 = title, row 1 = headers, row 2 = units, row 3+ = data.
    Returns None, None if the sheet has fewer than 3 rows.
    """
    rows_iter = ws.rows
    try:
        next(rows_iter)               # row 0: table title
        header_cells = next(rows_iter)  # row 1: column headers
        next(rows_iter)               # row 2: units — skip
    except StopIteration:
        return None, None
    col_map = {
        cell.value: i
        for i, cell in enumerate(header_cells)
        if cell.value is not None
    }
    return col_map, rows_iter


def _snap(v, tol):
    """Snap float to tolerance grid for use as a dict key."""
    return round(float(v) / tol)


def _geom_key(x1, y1, z1, x2, y2, z2, tol):
    """
    Canonical geometry key for a frame element defined by two 3D endpoints.
    Sorted so (JtI, JtJ) and (JtJ, JtI) produce the same key.
    """
    a = (_snap(x1, tol), _snap(y1, tol), _snap(z1, tol))
    b = (_snap(x2, tol), _snap(y2, tol), _snap(z2, tol))
    return (min(a, b), max(a, b))


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def parse_coord_index(wb, tolerance: float = COORD_TOLERANCE) -> dict:
    """
    Parse joints + frames sheets from an open workbook.

    Returns {object_label: geom_key} for all frame elements,
    where geom_key is a canonical sorted endpoint tuple.
    Returns {} if either sheet is absent or malformed.
    """
    if JOINTS_SHEET not in wb.sheetnames or FRAMES_SHEET not in wb.sheetnames:
        return {}

    # --- parse joints: element_name → (X, Y, Z) ---
    col_map, rows_iter = _stream_rows(wb[JOINTS_SHEET])
    if col_map is None:
        return {}

    ename_i = col_map.get('Element Name')
    x_i     = col_map.get('Global X')
    y_i     = col_map.get('Global Y')
    z_i     = col_map.get('Global Z')
    if any(v is None for v in (ename_i, x_i, y_i, z_i)):
        return {}

    joint_coords = {}
    for row in rows_iter:
        vals = [c.value for c in row]
        n = len(vals)
        try:
            ename = vals[ename_i] if ename_i < n else None
            x     = float(vals[x_i]) if x_i < n and vals[x_i] is not None else None
            y     = float(vals[y_i]) if y_i < n and vals[y_i] is not None else None
            z     = float(vals[z_i]) if z_i < n and vals[z_i] is not None else None
        except (TypeError, ValueError):
            continue
        if ename is not None and x is not None and y is not None and z is not None:
            joint_coords[ename] = (x, y, z)

    if not joint_coords:
        return {}

    # --- parse frames: object_label → geom_key ---
    col_map, rows_iter = _stream_rows(wb[FRAMES_SHEET])
    if col_map is None:
        return {}

    label_i = col_map.get('Object Label')
    jti_i   = col_map.get('Elm JtI')
    jtj_i   = col_map.get('Elm JtJ')
    if any(v is None for v in (label_i, jti_i, jtj_i)):
        return {}

    result = {}
    for row in rows_iter:
        vals = [c.value for c in row]
        n = len(vals)
        label = vals[label_i] if label_i < n else None
        jti   = vals[jti_i]   if jti_i   < n else None
        jtj   = vals[jtj_i]   if jtj_i   < n else None
        if label is None or jti is None or jtj is None:
            continue
        # Element Name from joints may be int or str; normalise to match
        jti_key = int(jti) if isinstance(jti, float) and jti == int(jti) else jti
        jtj_key = int(jtj) if isinstance(jtj, float) and jtj == int(jtj) else jtj
        ci = joint_coords.get(jti_key) or joint_coords.get(str(jti_key))
        cj = joint_coords.get(jtj_key) or joint_coords.get(str(jtj_key))
        if ci is None or cj is None:
            continue
        result[label] = _geom_key(ci[0], ci[1], ci[2], cj[0], cj[1], cj[2], tolerance)

    return result


def build_label_map(
    wb_exist,
    wb_new,
    tolerance: float = COORD_TOLERANCE,
) -> tuple:
    """
    Build a {new_label: existing_label} remapping dict.

    Returns (label_map, warnings).
    - label_map is non-empty only when both workbooks have the required sheets
      and at least one geometric match is found.
    - warnings contains human-readable strings for missing sheets or unmatched
      elements; an empty warnings list means everything resolved cleanly.
    """
    warnings = []

    exist_index = parse_coord_index(wb_exist, tolerance)
    new_index   = parse_coord_index(wb_new,   tolerance)

    if not exist_index:
        warnings.append(
            'Coordinate matching: existing model is missing Joints/Frames sheets '
            'or they could not be parsed. Falling back to label matching.'
        )
        return {}, warnings

    if not new_index:
        warnings.append(
            'Coordinate matching: modified model is missing Joints/Frames sheets '
            'or they could not be parsed. Falling back to label matching.'
        )
        return {}, warnings

    # Invert existing index: geom_key → existing_label
    geom_to_exist = {v: k for k, v in exist_index.items()}

    label_map = {}
    unmatched = []
    for new_label, geom_key in new_index.items():
        exist_label = geom_to_exist.get(geom_key)
        if exist_label is not None:
            label_map[new_label] = exist_label
        else:
            unmatched.append(str(new_label))

    if unmatched:
        sample = ', '.join(unmatched[:5])
        suffix = f' (and {len(unmatched) - 5} more)' if len(unmatched) > 5 else ''
        warnings.append(
            f'Coordinate matching: {len(unmatched)} element(s) in the modified model '
            f'could not be matched by geometry and will be treated as ADDED: '
            f'{sample}{suffix}.'
        )

    return label_map, warnings


def apply_label_map(parsed_dict: dict, label_map: dict) -> dict:
    """
    Return a shallow copy of parsed_dict with Label values remapped per label_map.

    parsed_dict structure (design mode):
        {'summary': df, 'forces': {'Columns': df, 'Beams': df, 'Braces': df}, ...}

    Both 'summary' and 'forces' DataFrames contain a Label column used for
    (Story, Label) matching — both are remapped. Rows whose label is not in
    label_map are left unchanged (they fall through to ADDED/REMOVED in the
    comparison engine).

    A copy is returned — the original cached parsed_dict is never mutated.
    """
    if not label_map:
        return parsed_dict

    result = dict(parsed_dict)

    # Remap forces DataFrames
    forces_orig = parsed_dict.get('forces', {})
    forces_new = {}
    for mt, df in forces_orig.items():
        if df is None or df.empty:
            forces_new[mt] = df
            continue
        df2 = df.copy()
        df2['Label'] = df2['Label'].map(lambda lbl: label_map.get(lbl, lbl))
        forces_new[mt] = df2
    result['forces'] = forces_new

    # Remap summary DataFrame (design mode only; analysis mode has no summary)
    summary_orig = parsed_dict.get('summary')
    if summary_orig is not None and not summary_orig.empty and 'Label' in summary_orig.columns:
        summary_new = summary_orig.copy()
        summary_new['Label'] = summary_new['Label'].map(
            lambda lbl: label_map.get(lbl, lbl)
        )
        result['summary'] = summary_new

    return result
