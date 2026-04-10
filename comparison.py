import io
import re

import pandas as pd

FORCE_COMPONENTS = ['P', 'V2', 'V3', 'T', 'M2', 'M3']

MEMBER_SHEETS = {
    'Columns': 'Element Forces - Columns',
    'Beams':   'Element Forces - Beams',
    'Braces':  'Element Forces - Braces',
}

# Member-specific label column name in each sheet
MEMBER_LABEL_COL = {
    'Columns': 'Column',
    'Beams':   'Beam',
    'Braces':  'Brace',
}

# Matches lateral load tokens like "1.0SA", "0.7WA", "Nx", "EQ1" in combo names.
# Tokens may be preceded by a coefficient, operator, or start-of-string.
LATERAL_RE = re.compile(
    r'(?:^|[\s+\-*/,()])(?:\d+\.?\d*)?'
    r'(W[A-Z]|S[AB][0-9]?|EQ[A-Z0-9]*|SDX|SDY|SX|SY|Nx|Ny|Nz)',
    re.IGNORECASE,
)

MATCH_COLS = ['Story', 'Frame', 'Station', 'OutputCase']
EMPTY_COLS = MATCH_COLS + ['CaseType'] + FORCE_COMPONENTS


def parse_force_sheet(file_obj, member_type: str) -> pd.DataFrame:
    """Read one force sheet and return a normalised DataFrame.

    Returns an empty DataFrame (with correct columns) if the sheet is missing.
    """
    sheet_name = MEMBER_SHEETS[member_type]
    label_col = MEMBER_LABEL_COL[member_type]

    try:
        with file_obj.open() as f:
            df = pd.read_excel(
                f,
                sheet_name=sheet_name,
                header=1,      # row index 1 (0-based) is the column header row
                skiprows=[2],  # row index 2 is the units row — drop it
                engine='openpyxl',
            )
    except Exception:
        return pd.DataFrame(columns=EMPTY_COLS)

    # Rename member-specific label to 'Frame', normalise other column names
    rename = {
        label_col:    'Frame',
        'Output Case': 'OutputCase',
        'Case Type':   'CaseType',
    }
    df = df.rename(columns=rename)

    # Drop rows with no Frame value (trailing blank rows in the sheet)
    df = df[df['Frame'].notna()].copy()

    if df.empty:
        return pd.DataFrame(columns=EMPTY_COLS)

    # Coerce force columns to numeric (in case of stray text)
    for col in FORCE_COMPONENTS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

    # Round Station to 4 decimal places for stable matching
    df['Station'] = df['Station'].round(4)

    # Deduplicate: sub-element rows (e.g. "133-1", "133-2") share the same key.
    # Keep the first occurrence for each (Story, Frame, Station, OutputCase).
    df = df.drop_duplicates(subset=MATCH_COLS, keep='first')

    return df[EMPTY_COLS]


def parse_combo_names(file_obj) -> list:
    """Return sorted unique load combination names from the file.

    Returns an empty list if the sheet is missing or unreadable.
    """
    try:
        with file_obj.open() as f:
            df = pd.read_excel(
                f,
                sheet_name='Load Combination Definitions',
                header=1,
                skiprows=[2],
                usecols=['Name'],
                engine='openpyxl',
            )
        return sorted(df['Name'].dropna().unique().tolist())
    except Exception:
        return []


def classify_combo(combo_name: str) -> str:
    """Return 'lateral' if combo_name contains a lateral load token, else 'gravity'."""
    return 'lateral' if LATERAL_RE.search(combo_name) else 'gravity'


def run_comparison(
    existing_file,
    modified_file,
    selected_combos: list,
    member_type_filter: str,
    gravity_threshold: float,
    lateral_threshold: float,
    show_failures_only: bool,
) -> list:
    """Core comparison engine.

    Returns a list of row dicts with keys:
    MemberType, Story, Frame, Station, OutputCase, Component,
    ExistingValue, ModifiedValue, PctChange, Pass
    """
    types_to_process = (
        ['Columns', 'Beams', 'Braces']
        if member_type_filter == 'All'
        else [member_type_filter]
    )

    # Pre-classify each combo once
    combo_class = {c: classify_combo(c) for c in selected_combos}
    selected_set = set(selected_combos)

    all_results = []

    for mtype in types_to_process:
        df_exist = parse_force_sheet(existing_file, mtype)
        df_mod   = parse_force_sheet(modified_file, mtype)

        # Filter to only selected combos
        df_exist = df_exist[df_exist['OutputCase'].isin(selected_set)]
        df_mod   = df_mod[df_mod['OutputCase'].isin(selected_set)]

        if df_mod.empty:
            continue

        # Left-merge: modified is left; existing is right.
        # Rows in modified with no match in existing → NaN for _exist columns.
        merged = df_mod.merge(
            df_exist[MATCH_COLS + FORCE_COMPONENTS],
            on=MATCH_COLS,
            how='left',
            suffixes=('_mod', '_exist'),
        )

        # Melt force components into rows
        mod_cols   = [f'{c}_mod'   for c in FORCE_COMPONENTS]
        exist_cols = [f'{c}_exist' for c in FORCE_COMPONENTS]

        id_cols = MATCH_COLS + ['CaseType']

        mod_melted = merged[id_cols + mod_cols].melt(
            id_vars=id_cols, var_name='Component', value_name='ModifiedValue'
        )
        mod_melted['Component'] = mod_melted['Component'].str.replace('_mod', '', regex=False)

        exist_melted = merged[id_cols + exist_cols].melt(
            id_vars=id_cols, var_name='Component', value_name='ExistingValue'
        )
        exist_melted['Component'] = exist_melted['Component'].str.replace('_exist', '', regex=False)

        rows = mod_melted.copy()
        rows['ExistingValue'] = exist_melted['ExistingValue'].values

        # Compute % change
        mod_vals   = rows['ModifiedValue'].astype(float)
        exist_vals = rows['ExistingValue']  # may be NaN for new members

        near_zero = 1e-6

        # Determine thresholds per row based on combo classification
        thresholds = rows['OutputCase'].map(
            lambda c: lateral_threshold if combo_class.get(c) == 'lateral' else gravity_threshold
        )

        # Build pct_change and pass columns
        pct_changes = []
        pass_flags  = []

        for i in range(len(rows)):
            m = float(mod_vals.iloc[i])
            e_raw = exist_vals.iloc[i]
            thresh = float(thresholds.iloc[i])

            if pd.isna(e_raw):
                # New member — no match in existing model
                pct_changes.append('N/A')
                pass_flags.append('FAIL')
                continue

            e = float(e_raw)

            if abs(e) < near_zero and abs(m) < near_zero:
                pct_changes.append(0.0)
                pass_flags.append('PASS')
            elif abs(e) < near_zero:
                pct_changes.append('INF')
                pass_flags.append('FAIL')
            else:
                pct = (m - e) / abs(e) * 100.0
                pct_changes.append(round(pct, 2))
                pass_flags.append('PASS' if abs(pct) <= thresh else 'FAIL')

        rows['PctChange'] = pct_changes
        rows['Pass']      = pass_flags

        # Round display values
        rows['ExistingValue'] = rows['ExistingValue'].apply(
            lambda v: round(float(v), 4) if pd.notna(v) else 'N/A'
        )
        rows['ModifiedValue'] = rows['ModifiedValue'].apply(lambda v: round(float(v), 4))
        rows['MemberType'] = mtype.rstrip('s')  # 'Column', 'Beam', 'Brace'

        if show_failures_only:
            rows = rows[rows['Pass'] == 'FAIL']

        output_cols = [
            'MemberType', 'Story', 'Frame', 'Station',
            'OutputCase', 'Component',
            'ExistingValue', 'ModifiedValue', 'PctChange', 'Pass',
        ]
        all_results.extend(rows[output_cols].to_dict('records'))

    return all_results


def results_to_csv(results: list) -> str:
    """Convert results list-of-dicts to a CSV string."""
    if not results:
        cols = [
            'MemberType', 'Story', 'Frame', 'Station',
            'OutputCase', 'Component',
            'ExistingValue', 'ModifiedValue', 'PctChange', 'Pass',
        ]
        return ','.join(cols) + '\n'
    return pd.DataFrame(results).to_csv(index=False)
