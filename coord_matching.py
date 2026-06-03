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


def _norm_joint_key(v):
    """
    Normalise a joint Element Name to a consistent type for dict lookup.

    ETABS can export joint IDs as int, float (e.g. 101.0), or numeric string
    ('101').  We canonicalise all whole-number values to int so that
    joint_coords built in one pass is always looked up with matching key types.
    """
    if isinstance(v, int):
        return v
    if isinstance(v, float) and v == int(v):
        return int(v)
    if isinstance(v, str):
        try:
            f = float(v)
            if f == int(f):
                return int(f)
        except (ValueError, TypeError):
            pass
    return v


def _col(col_map, *candidates):
    """Return the first index found among candidate column names, or None."""
    for name in candidates:
        if name in col_map:
            return col_map[name]
    return None


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

    # ETABS exports joint element IDs as 'Unique Name' in modern versions,
    # 'Element Name' in older ones.
    ename_i = _col(col_map, 'Unique Name', 'Element Name', 'Object Name')
    x_i     = _col(col_map, 'Global X', 'GlobalX', 'X')
    y_i     = _col(col_map, 'Global Y', 'GlobalY', 'Y')
    z_i     = _col(col_map, 'Global Z', 'GlobalZ', 'Z')
    if any(v is None for v in (ename_i, x_i, y_i, z_i)):
        return {}

    joint_coords = {}
    for row in rows_iter:
        vals = [c.value for c in row]
        n = len(vals)
        try:
            ename = _norm_joint_key(vals[ename_i] if ename_i < n else None)
            x     = float(vals[x_i]) if x_i < n and vals[x_i] is not None else None
            y     = float(vals[y_i]) if y_i < n and vals[y_i] is not None else None
            z     = float(vals[z_i]) if z_i < n and vals[z_i] is not None else None
        except (TypeError, ValueError):
            continue
        if ename is not None and x is not None and y is not None and z is not None:
            joint_coords[ename] = (x, y, z)

    if not joint_coords:
        return {}

    # --- parse frames: collect all element endpoints per object label ---
    col_map, rows_iter = _stream_rows(wb[FRAMES_SHEET])
    if col_map is None:
        return {}

    label_i = _col(col_map, 'Object Label', 'Label', 'Frame')
    jti_i   = _col(col_map, 'Elm JtI', 'Elm Jt I', 'ElemJtI', 'JtI', 'StartJoint', 'I Joint')
    jtj_i   = _col(col_map, 'Elm JtJ', 'Elm Jt J', 'ElemJtJ', 'JtJ', 'EndJoint',   'J Joint')
    if any(v is None for v in (label_i, jti_i, jtj_i)):
        return {}

    # Each ETABS frame object is divided into ≥1 finite elements.
    # We need the OBJECT endpoints (I-joint and J-joint of the whole member),
    # not the per-element joints.  The terminal joints are those that appear
    # exactly once across all elements of a given frame label — they are the
    # object endpoints.  Internal (intermediate) joints appear exactly twice.
    frame_elem_joints: dict = {}   # label → [(jti_key, jtj_key), ...]
    for row in rows_iter:
        vals = [c.value for c in row]
        n    = len(vals)
        label_raw = vals[label_i] if label_i < n else None
        jti       = vals[jti_i]   if jti_i   < n else None
        jtj       = vals[jtj_i]   if jtj_i   < n else None
        if label_raw is None or jti is None or jtj is None:
            continue
        label = str(label_raw).strip()
        jti_key = _norm_joint_key(jti)
        jtj_key = _norm_joint_key(jtj)
        if label not in frame_elem_joints:
            frame_elem_joints[label] = []
        frame_elem_joints[label].append((jti_key, jtj_key))

    result = {}
    for label, elem_joints in frame_elem_joints.items():
        if len(elem_joints) == 1:
            # Single-element frame — use directly
            jti_key, jtj_key = elem_joints[0]
        else:
            # Multi-element frame — find terminal joints (appear exactly once)
            cnt: dict = {}
            for jti, jtj in elem_joints:
                cnt[jti] = cnt.get(jti, 0) + 1
                cnt[jtj] = cnt.get(jtj, 0) + 1
            terminal = [j for j, c in cnt.items() if c == 1]
            if len(terminal) != 2:
                continue  # ring or degenerate — skip
            jti_key, jtj_key = terminal[0], terminal[1]

        ci = joint_coords.get(jti_key)
        cj = joint_coords.get(jtj_key)
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


def build_label_map_from_indices(exist_index: dict, new_index: dict) -> tuple:
    """
    Build a {new_label: existing_label} map from two pre-parsed coord indices.

    Equivalent to build_label_map but works from already-extracted dicts so no
    workbook I/O is required.  Returns (label_map, unmatched_labels).
    """
    if not exist_index or not new_index:
        return {}, []

    geom_to_exist = {v: k for k, v in exist_index.items()}

    label_map: dict = {}
    unmatched: list = []
    for new_label, geom_key in new_index.items():
        exist_label = geom_to_exist.get(geom_key)
        if exist_label is not None:
            label_map[new_label] = exist_label
        else:
            unmatched.append(str(new_label))

    return label_map, unmatched


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
        str_map = {str(k).strip(): v for k, v in label_map.items()}
        df2['Label'] = df2['Label'].map(
            lambda lbl: str_map.get(str(lbl).strip(), lbl)
        )
        forces_new[mt] = df2
    result['forces'] = forces_new

    # Remap summary DataFrame (design mode only; analysis mode has no summary)
    summary_orig = parsed_dict.get('summary')
    if summary_orig is not None and not summary_orig.empty and 'Label' in summary_orig.columns:
        summary_new = summary_orig.copy()
        str_map = {str(k).strip(): v for k, v in label_map.items()}
        summary_new['Label'] = summary_new['Label'].map(
            lambda lbl: str_map.get(str(lbl).strip(), lbl)
        )
        result['summary'] = summary_new

    return result
