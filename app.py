import viktor as vkt

from comparison import build_summary, results_to_csv, run_comparison

COLUMN_HEADERS = [
    'Story', 'Label', 'Type',
    'Section (Exist)', 'Section (New)',
    'Gov Combo (Exist)', 'Gov Combo (New)', 'Load Type',
    'P (Exist)', 'P (New)', 'P (%)',
    'V2 (Exist)', 'V2 (New)', 'V2 (%)',
    'V3 (Exist)', 'V3 (New)', 'V3 (%)',
    'T (Exist)', 'T (New)', 'T (%)',
    'M2 (Exist)', 'M2 (New)', 'M2 (%)',
    'M3 (Exist)', 'M3 (New)', 'M3 (%)',
    'PMM (Exist)', 'PMM (New)', 'PMM (%)',
    'V Maj (Exist)', 'V Maj (New)', 'V Maj (%)',
    'Sign Rev.', 'Result',
]

SUMMARY_HEADERS = ['Story', 'Type', 'Total', 'PASS', 'FAIL', 'ADDED', 'REMOVED']

_RESULT_COLORS = {
    'PASS':    vkt.Color(34, 139, 34),   # green
    'FAIL':    vkt.Color(178, 34, 34),   # red
    'ADDED':   vkt.Color(30, 100, 200),  # blue
    'REMOVED': vkt.Color(180, 100, 0),   # orange
}


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

def validate_step1(params, **kwargs):
    violations = []
    if not params.step1.existing_file:
        violations.append(vkt.InputViolation(
            'Please upload the existing model file',
            fields=['step1.existing_file'],
        ))
    if not params.step1.modified_file:
        violations.append(vkt.InputViolation(
            'Please upload the modified model file',
            fields=['step1.modified_file'],
        ))
    if violations:
        raise vkt.UserError(
            'Both ETABS design output files must be uploaded before proceeding.',
            input_violations=violations,
        )


# ---------------------------------------------------------------------------
# Parametrization
# ---------------------------------------------------------------------------

class Parametrization(vkt.Parametrization):

    step1 = vkt.Step('Upload Files', on_next=validate_step1)

    step1.intro = vkt.Text("""
## Upload ETABS Design Output Exports

Upload two ETABS Excel exports (.xlsx) — one for the **existing** model and
one for the **modified** model.

Each file must contain:
- *Design Forces - Columns / Beams / Braces* — governing forces per member
- *Stl Frm Sum - AISC 360-16* — D/C ratios and governing combo names

Forces are compared using the **maximum absolute value across all design
combos and stations** for each force component. IBC 3403 thresholds are
applied based on whether the governing combo is gravity or lateral.
""")

    step1.existing_file = vkt.FileField(
        'Existing Model (.xlsx)',
        file_types=['.xlsx'],
        description='ETABS design output for the pre-modification model',
    )
    step1.modified_file = vkt.FileField(
        'Modified Model (.xlsx)',
        file_types=['.xlsx'],
        description='ETABS design output for the modified / repurposed model',
    )

    step2 = vkt.Step('Configure & Compare', views=['results_table', 'summary_table'])

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
        description='Maximum allowable % increase for gravity combos (IBC 3403)',
    )
    step2.section_options.lateral_threshold = vkt.NumberField(
        'Lateral Load Threshold (%)',
        default=10,
        min=0,
        description='Maximum allowable % increase for lateral combos (IBC 3403)',
    )
    step2.section_options.display_filter = vkt.OptionField(
        'Display Filter',
        options=['All Results', 'Failures Only'],
        default='Failures Only',
    )

    step2.section_export = vkt.Section('Export')
    step2.section_export.download_btn = vkt.DownloadButton(
        'Export Full Results to CSV',
        method='download_csv',
        longpoll=True,
    )


# ---------------------------------------------------------------------------
# Controller
# ---------------------------------------------------------------------------

class Controller(vkt.Controller):
    parametrization = Parametrization

    @vkt.TableView('Comparison Results', duration_guess=30)
    def results_table(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')

        results = self._run(params)

        if not results:
            raise vkt.UserError(
                'No results to display. If "Failures Only" is selected, '
                'all members may be passing. Try switching to "All Results".'
            )

        data = []
        for r in results:
            pass_val = r.get('Pass', '')
            color = _RESULT_COLORS.get(pass_val, vkt.Color(128, 128, 128))
            pass_cell = vkt.TableCell(
                pass_val,
                background_color=color,
                text_color=vkt.Color(255, 255, 255),
            )
            data.append([
                r.get('Story', ''),
                r.get('Label', ''),
                r.get('MemberType', ''),
                r.get('DesignSection_Exist', ''),
                r.get('DesignSection_New', ''),
                r.get('GovCombo_Exist', ''),
                r.get('GovCombo_New', ''),
                r.get('LoadType', ''),
                r.get('P_Exist', ''), r.get('P_New', ''), r.get('P_Pct', ''),
                r.get('V2_Exist', ''), r.get('V2_New', ''), r.get('V2_Pct', ''),
                r.get('V3_Exist', ''), r.get('V3_New', ''), r.get('V3_Pct', ''),
                r.get('T_Exist', ''), r.get('T_New', ''), r.get('T_Pct', ''),
                r.get('M2_Exist', ''), r.get('M2_New', ''), r.get('M2_Pct', ''),
                r.get('M3_Exist', ''), r.get('M3_New', ''), r.get('M3_Pct', ''),
                r.get('PMM_Exist', ''), r.get('PMM_New', ''), r.get('PMM_Pct', ''),
                r.get('VMaj_Exist', ''), r.get('VMaj_New', ''), r.get('VMaj_Pct', ''),
                r.get('SignReversal', ''),
                pass_cell,
            ])

        return vkt.TableResult(data, column_headers=COLUMN_HEADERS)

    @vkt.TableView('Summary by Story', duration_guess=30)
    def summary_table(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')

        # Summary always runs over all results (no failures-only filter)
        p = params.step2.section_options
        all_results = run_comparison(
            existing_file=params.step1.existing_file.file,
            modified_file=params.step1.modified_file.file,
            member_type_filter=p.member_type or 'All',
            gravity_threshold=float(p.gravity_threshold or 5),
            lateral_threshold=float(p.lateral_threshold or 10),
            show_failures_only=False,
        )

        summary = build_summary(all_results)

        if not summary:
            raise vkt.UserError('No data to summarise.')

        data = []
        for s in summary:
            fail_count = s['FAIL']
            fail_cell = vkt.TableCell(
                str(fail_count),
                background_color=(
                    vkt.Color(178, 34, 34) if fail_count > 0
                    else vkt.Color(34, 139, 34)
                ),
                text_color=vkt.Color(255, 255, 255),
            )
            data.append([
                s['Story'],
                s['MemberType'],
                s['Total'],
                s['PASS'],
                fail_cell,
                s['ADDED'],
                s['REMOVED'],
            ])

        return vkt.TableResult(data, column_headers=SUMMARY_HEADERS)

    def download_csv(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')

        p = params.step2.section_options
        results = run_comparison(
            existing_file=params.step1.existing_file.file,
            modified_file=params.step1.modified_file.file,
            member_type_filter=p.member_type or 'All',
            gravity_threshold=float(p.gravity_threshold or 5),
            lateral_threshold=float(p.lateral_threshold or 10),
            show_failures_only=False,  # always export everything
        )
        csv_string = results_to_csv(results)
        return vkt.DownloadResult(csv_string, 'etabs_comparison_results.csv')

    # -------------------------------------------------------------------------
    # Shared helper
    # -------------------------------------------------------------------------

    def _run(self, params):
        p = params.step2.section_options
        return run_comparison(
            existing_file=params.step1.existing_file.file,
            modified_file=params.step1.modified_file.file,
            member_type_filter=p.member_type or 'All',
            gravity_threshold=float(p.gravity_threshold or 5),
            lateral_threshold=float(p.lateral_threshold or 10),
            show_failures_only=(p.display_filter == 'Failures Only'),
        )
