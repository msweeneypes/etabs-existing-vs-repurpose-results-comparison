import re

import viktor as vkt

from comparison import build_summary, results_to_csv, run_comparison, FORCE_COMPONENTS

# ---------------------------------------------------------------------------
# Column headers
# ---------------------------------------------------------------------------

OVERVIEW_HEADERS = [
    'Story', 'Label', 'Type',
    'Section (Exist)', 'Section (New)',
    'Load Type',
    'PMM (Exist)', 'PMM (New)', 'PMM (%)',
    'V Major (%)',
    'Worst Force (%)',
    'Fail Reason',
    'Result',
]

DETAIL_HEADERS = [
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
    'Sign Rev.', 'Fail Reason', 'Result',
]

SUMMARY_HEADERS = ['Story', 'Type', 'Total', 'PASS', 'FAIL', 'ADDED', 'REMOVED']

# ---------------------------------------------------------------------------
# Colors
# ---------------------------------------------------------------------------

_RESULT_COLORS = {
    'PASS':    vkt.Color(34, 139, 34),
    'FAIL':    vkt.Color(178, 34, 34),
    'ADDED':   vkt.Color(30, 100, 200),
    'REMOVED': vkt.Color(180, 100, 0),
}

_CHART_COLORS = {
    'FAIL':    '#B22222',
    'PASS':    '#228B22',
    'ADDED':   '#1E64C8',
    'REMOVED': '#B46400',
}


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------

def _result_cell(pass_val: str) -> vkt.TableCell:
    color = _RESULT_COLORS.get(pass_val, vkt.Color(128, 128, 128))
    return vkt.TableCell(pass_val, background_color=color,
                         text_color=vkt.Color(255, 255, 255))


def _fmt_pct(val) -> str:
    """Format a percent change value for display with explicit sign."""
    if isinstance(val, float):
        sign = '+' if val >= 0 else ''
        return f'{sign}{val:.1f}%'
    return str(val)


def _worst_force_pct(row: dict) -> str:
    """Return the largest positive force % change across all components."""
    vals = [
        row.get(f'{c}_Pct') for c in FORCE_COMPONENTS
        if isinstance(row.get(f'{c}_Pct'), float)
    ]
    if not vals:
        return 'N/A'
    worst = max(vals)
    sign = '+' if worst >= 0 else ''
    return f'{sign}{worst:.1f}%'


def _story_sort_key(name: str) -> tuple:
    """Sort stories by trailing digits (lowest first), then alphabetical."""
    m = re.search(r'(\d+)', str(name))
    return (int(m.group(1)) if m else 0, str(name))


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

    step2 = vkt.Step('Configure & Compare', views=[
        'results_table',
        'results_detail',
        'results_chart',
        'summary_table',
        'key_metrics',
    ])

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

    # -- Overview table (primary view) ---------------------------------------

    @vkt.TableView('Overview', duration_guess=30)
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
            data.append([
                r.get('Story', ''),
                r.get('Label', ''),
                r.get('MemberType', ''),
                r.get('DesignSection_Exist', ''),
                r.get('DesignSection_New', ''),
                r.get('LoadType', ''),
                r.get('PMM_Exist', ''),
                r.get('PMM_New', ''),
                _fmt_pct(r.get('PMM_Pct', '')),
                _fmt_pct(r.get('VMaj_Pct', '')),
                _worst_force_pct(r),
                r.get('FailReason', ''),
                _result_cell(r.get('Pass', '')),
            ])

        return vkt.TableResult(data, column_headers=OVERVIEW_HEADERS)

    # -- Full detail table ---------------------------------------------------

    @vkt.TableView('Full Detail', duration_guess=30)
    def results_detail(self, params, **kwargs):
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
            data.append([
                r.get('Story', ''),
                r.get('Label', ''),
                r.get('MemberType', ''),
                r.get('DesignSection_Exist', ''),
                r.get('DesignSection_New', ''),
                r.get('GovCombo_Exist', ''),
                r.get('GovCombo_New', ''),
                r.get('LoadType', ''),
                r.get('P_Exist', ''), r.get('P_New', ''), _fmt_pct(r.get('P_Pct', '')),
                r.get('V2_Exist', ''), r.get('V2_New', ''), _fmt_pct(r.get('V2_Pct', '')),
                r.get('V3_Exist', ''), r.get('V3_New', ''), _fmt_pct(r.get('V3_Pct', '')),
                r.get('T_Exist', ''), r.get('T_New', ''), _fmt_pct(r.get('T_Pct', '')),
                r.get('M2_Exist', ''), r.get('M2_New', ''), _fmt_pct(r.get('M2_Pct', '')),
                r.get('M3_Exist', ''), r.get('M3_New', ''), _fmt_pct(r.get('M3_Pct', '')),
                r.get('PMM_Exist', ''), r.get('PMM_New', ''), _fmt_pct(r.get('PMM_Pct', '')),
                r.get('VMaj_Exist', ''), r.get('VMaj_New', ''), _fmt_pct(r.get('VMaj_Pct', '')),
                r.get('SignReversal', ''),
                r.get('FailReason', ''),
                _result_cell(r.get('Pass', '')),
            ])

        return vkt.TableResult(data, column_headers=DETAIL_HEADERS)

    # -- Plotly chart --------------------------------------------------------

    @vkt.PlotlyView('Results Chart', duration_guess=30)
    def results_chart(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')

        all_results = self._run_all(params)
        if not all_results:
            raise vkt.UserError('No data to chart.')

        import pandas as pd
        from plotly.subplots import make_subplots
        import plotly.graph_objects as go

        df = pd.DataFrame(all_results)[['Story', 'MemberType', 'Pass']]

        member_types = sorted(df['MemberType'].unique())
        stories = sorted(df['Story'].unique(), key=_story_sort_key, reverse=True)
        n_types = len(member_types)

        grouped = (
            df.groupby(['Story', 'MemberType', 'Pass'])
            .size()
            .reset_index(name='Count')
        )

        fig = make_subplots(
            rows=1, cols=n_types,
            subplot_titles=member_types,
            shared_yaxes=True,
            horizontal_spacing=0.04,
        )

        status_order = ['FAIL', 'PASS', 'ADDED', 'REMOVED']
        for col_idx, mtype in enumerate(member_types, start=1):
            sub = grouped[grouped['MemberType'] == mtype]
            for status in status_order:
                s = sub[sub['Pass'] == status]
                counts_by_story = dict(zip(s['Story'], s['Count']))
                x_vals = [counts_by_story.get(st, 0) for st in stories]
                fig.add_trace(
                    go.Bar(
                        name=status,
                        x=x_vals,
                        y=stories,
                        orientation='h',
                        marker_color=_CHART_COLORS[status],
                        showlegend=(col_idx == 1),
                        legendgroup=status,
                        hovertemplate='%{y}: %{x}<extra>' + status + '</extra>',
                    ),
                    row=1, col=col_idx,
                )

        fig.update_layout(
            barmode='stack',
            title='Members by Story and Status (highest story at top)',
            height=max(420, 28 * len(stories) + 120),
            legend=dict(orientation='h', yanchor='bottom', y=1.05,
                        xanchor='right', x=1),
            margin=dict(l=10, r=20, t=80, b=40),
        )

        return vkt.PlotlyResult(fig)

    # -- Summary table -------------------------------------------------------

    @vkt.TableView('Summary by Story', duration_guess=30)
    def summary_table(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')

        summary = build_summary(self._run_all(params))
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
                s['Story'], s['MemberType'], s['Total'],
                s['PASS'], fail_cell, s['ADDED'], s['REMOVED'],
            ])

        return vkt.TableResult(data, column_headers=SUMMARY_HEADERS)

    # -- Key metrics DataView ------------------------------------------------

    @vkt.DataView('Key Metrics', duration_guess=30)
    def key_metrics(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')

        all_results = self._run_all(params)
        if not all_results:
            raise vkt.UserError('No data.')

        from collections import Counter
        total    = len(all_results)
        failures = [r for r in all_results if r.get('Pass') == 'FAIL']
        added    = [r for r in all_results if r.get('Pass') == 'ADDED']
        removed  = [r for r in all_results if r.get('Pass') == 'REMOVED']
        n_fail   = len(failures)
        fail_rate = round(n_fail / total * 100, 1) if total > 0 else 0.0

        # Worst PMM change among matched members
        worst_pmm_pct = None
        worst_pmm_label = ''
        for r in all_results:
            pct = r.get('PMM_Pct')
            if isinstance(pct, float) and (worst_pmm_pct is None or pct > worst_pmm_pct):
                worst_pmm_pct = pct
                worst_pmm_label = f"{r.get('Story')} / {r.get('Label')}"

        # Most affected story
        story_fail_counts = Counter(r.get('Story') for r in failures)
        if story_fail_counts:
            top_story, top_count = story_fail_counts.most_common(1)[0]
        else:
            top_story, top_count = 'None', 0

        fail_status = vkt.DataStatus.ERROR if n_fail > 0 else vkt.DataStatus.SUCCESS
        rate_status = (
            vkt.DataStatus.ERROR   if fail_rate > 10 else
            vkt.DataStatus.WARNING if fail_rate > 0  else
            vkt.DataStatus.SUCCESS
        )
        pmm_status = (
            vkt.DataStatus.ERROR if (worst_pmm_pct or 0) > 10 else
            vkt.DataStatus.WARNING if (worst_pmm_pct or 0) > 0 else
            vkt.DataStatus.INFO
        )

        data = vkt.DataGroup(
            vkt.DataItem('Total Members Compared', total),
            vkt.DataItem(
                'Failures',
                n_fail,
                status=fail_status,
                status_message=(
                    'Members exceeding IBC 3403 thresholds'
                    if n_fail > 0 else 'All members within threshold'
                ),
            ),
            vkt.DataItem(
                'Failure Rate',
                fail_rate,
                suffix='%',
                number_of_decimals=1,
                status=rate_status,
            ),
            vkt.DataItem('Added Members', len(added)),
            vkt.DataItem('Removed Members', len(removed)),
            vkt.DataItem(
                'Worst PMM Change',
                worst_pmm_pct if worst_pmm_pct is not None else 0.0,
                suffix='%',
                number_of_decimals=1,
                status=pmm_status,
                explanation_label=worst_pmm_label,
            ),
            vkt.DataItem(
                'Most Affected Story',
                top_story,
                explanation_label=f'{top_count} failures',
            ),
        )
        return vkt.DataResult(data)

    # -- CSV export ----------------------------------------------------------

    def download_csv(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')
        results = self._run_all(params)
        return vkt.DownloadResult(results_to_csv(results), 'etabs_comparison_results.csv')

    # -- Shared helpers ------------------------------------------------------

    def _run(self, params):
        """Filtered results (respects display_filter param)."""
        p = params.step2.section_options
        return run_comparison(
            existing_file=params.step1.existing_file.file,
            modified_file=params.step1.modified_file.file,
            member_type_filter=p.member_type or 'All',
            gravity_threshold=float(p.gravity_threshold or 5),
            lateral_threshold=float(p.lateral_threshold or 10),
            show_failures_only=(p.display_filter == 'Failures Only'),
        )

    def _run_all(self, params):
        """Unfiltered results — used by chart, summary, key metrics, and CSV."""
        p = params.step2.section_options
        return run_comparison(
            existing_file=params.step1.existing_file.file,
            modified_file=params.step1.modified_file.file,
            member_type_filter=p.member_type or 'All',
            gravity_threshold=float(p.gravity_threshold or 5),
            lateral_threshold=float(p.lateral_threshold or 10),
            show_failures_only=False,
        )
