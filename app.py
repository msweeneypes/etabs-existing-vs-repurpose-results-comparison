import io

import pandas as pd
import viktor as vkt

from comparison import parse_combo_names, run_comparison, results_to_csv

COLUMN_HEADERS = [
    'MemberType', 'Story', 'Frame', 'Station',
    'OutputCase', 'Component',
    'ExistingValue', 'ModifiedValue', 'PctChange', 'Pass',
]


# ---------------------------------------------------------------------------
# Dynamic options callback — populates MultiSelectField from uploaded file
# ---------------------------------------------------------------------------

def get_combo_options(params, **kwargs):
    if not params.step1.existing_file:
        return []
    try:
        return parse_combo_names(params.step1.existing_file.file)
    except Exception:
        return []


# ---------------------------------------------------------------------------
# Step 1 validation — both files must be uploaded before proceeding
# ---------------------------------------------------------------------------

def validate_step1(params, **kwargs):
    violations = []
    if not params.step1.existing_file:
        violations.append(
            vkt.InputViolation(
                'Please upload the existing model file',
                fields=['step1.existing_file'],
            )
        )
    if not params.step1.modified_file:
        violations.append(
            vkt.InputViolation(
                'Please upload the modified model file',
                fields=['step1.modified_file'],
            )
        )
    if violations:
        raise vkt.UserError(
            'Both ETABS export files must be uploaded before proceeding.',
            input_violations=violations,
        )


# ---------------------------------------------------------------------------
# Parametrization
# ---------------------------------------------------------------------------

class Parametrization(vkt.Parametrization):

    # -- Step 1: File Upload -------------------------------------------------
    step1 = vkt.Step('Upload Files', on_next=validate_step1)

    step1.intro = vkt.Text("""
## Upload ETABS Exports

Upload two ETABS Excel exports (.xlsx) — one for the **existing** model and one
for the **modified** (repurposed) model.

Each file should contain sheets named:
- *Element Forces - Columns*
- *Element Forces - Beams*
- *Element Forces - Braces*
- *Load Combination Definitions*
""")

    step1.existing_file = vkt.FileField(
        'Existing Model (.xlsx)',
        file_types=['.xlsx'],
        description='ETABS export for the pre-modification model',
    )
    step1.modified_file = vkt.FileField(
        'Modified Model (.xlsx)',
        file_types=['.xlsx'],
        description='ETABS export for the modified / repurposed model',
    )

    # -- Step 2: Configure & Compare -----------------------------------------
    step2 = vkt.Step('Configure & Compare', views=['results_table'])

    step2.section_combos = vkt.Section('Load Combinations')
    step2.section_combos.load_combos = vkt.MultiSelectField(
        'Filter Load Combinations (optional)',
        options=get_combo_options,
        description='Leave empty to compare ALL combinations. Select specific ones to narrow the scope.',
    )

    step2.section_options = vkt.Section('Comparison Options')
    step2.section_options.member_type = vkt.OptionField(
        'Member Type',
        options=['All', 'Columns', 'Beams', 'Braces'],
        default='All',
    )
    step2.section_options.gravity_threshold = vkt.NumberField(
        'Gravity Load Threshold (%)',
        default=5,
        min=0,
        description='Maximum allowable % increase for gravity load combinations (IBC 3403)',
    )
    step2.section_options.lateral_threshold = vkt.NumberField(
        'Lateral Load Threshold (%)',
        default=10,
        min=0,
        description='Maximum allowable % increase for lateral load combinations (IBC 3403)',
    )
    step2.section_options.display_filter = vkt.OptionField(
        'Display Filter',
        options=['All Results', 'Failures Only'],
        default='Failures Only',
    )

    step2.section_export = vkt.Section('Export')
    step2.section_export.download_btn = vkt.DownloadButton(
        'Export Results to CSV',
        method='download_csv',
        longpoll=True,
    )


# ---------------------------------------------------------------------------
# Controller
# ---------------------------------------------------------------------------

class Controller(vkt.Controller):
    parametrization = Parametrization

    @vkt.TableView('Comparison Results', duration_guess=15)
    def results_table(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')

        results = self._run(params)

        if not results:
            # --- Temporary diagnostics: show what's in the file ---
            diag = ['No results returned. Diagnostic info:']
            for label, ff in [('EXISTING', params.step1.existing_file),
                               ('MODIFIED', params.step1.modified_file)]:
                if not ff:
                    diag.append(f'{label}: not uploaded')
                    continue
                try:
                    with ff.file.open() as fh:
                        file_bytes = fh.read()
                    xl = pd.ExcelFile(io.BytesIO(file_bytes), engine='openpyxl')
                    sheet_names = xl.sheet_names
                    safe_names = [s.encode('ascii', errors='replace').decode('ascii')
                                  for s in sheet_names]
                    diag.append(f'{label} sheets: {safe_names}')
                    for sheet in sheet_names:
                        raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet,
                                            header=None, nrows=4, engine='openpyxl')
                        safe_cols = [str(c).encode('ascii', errors='replace').decode('ascii')
                                     for c in raw.iloc[1].tolist()]
                        safe_sheet = sheet.encode('ascii', errors='replace').decode('ascii')
                        diag.append(f'  "{safe_sheet}" row1: {safe_cols}')
                except Exception as e:
                    diag.append(f'{label} error: {e}')
            from comparison import parse_combo_names
            combos = parse_combo_names(params.step1.existing_file.file)
            diag.append(f'Combos discovered ({len(combos)}): {combos[:5]}')
            raise vkt.UserError('\n'.join(diag))
            # --- End diagnostics ---

        data = []
        for row in results:
            is_pass = (row['Pass'] == 'PASS')
            pass_cell = vkt.TableCell(
                row['Pass'],
                background_color=vkt.Color(34, 139, 34) if is_pass else vkt.Color(178, 34, 34),
                text_color=vkt.Color(255, 255, 255),
            )
            data.append([
                row['MemberType'],
                row['Story'],
                row['Frame'],
                row['Station'],
                row['OutputCase'],
                row['Component'],
                row['ExistingValue'],
                row['ModifiedValue'],
                row['PctChange'],
                pass_cell,
            ])

        return vkt.TableResult(data, column_headers=COLUMN_HEADERS)

    def download_csv(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')

        results = self._run(params)
        csv_string = results_to_csv(results)
        return vkt.DownloadResult(csv_string, 'etabs_comparison_results.csv')

    # -------------------------------------------------------------------------
    # Shared helper — extracts params and calls run_comparison
    # -------------------------------------------------------------------------

    def _run(self, params):
        p = params.step2.section_options
        selected = list(params.step2.section_combos.load_combos or [])
        return run_comparison(
            existing_file=params.step1.existing_file.file,
            modified_file=params.step1.modified_file.file,
            selected_combos=selected if selected else None,
            member_type_filter=p.member_type or 'All',
            gravity_threshold=float(p.gravity_threshold or 5),
            lateral_threshold=float(p.lateral_threshold or 10),
            show_failures_only=(p.display_filter == 'Failures Only'),
        )
