import re
from typing import Optional

import viktor as vkt
from openai import OpenAI

_llm_client = OpenAI(
    base_url=vkt.ViktorOpenAI.get_base_url(version="v1"),
    api_key=vkt.ViktorOpenAI.get_api_key(),
)

from comparison import (
    build_summary, results_to_csv, run_comparison, FORCE_COMPONENTS,
    DESIGN_FORCES_SHEETS, _get_parsed_file,
)
from analysis_comparison import (
    run_analysis_comparison, ANALYSIS_FORCES_SHEETS, _get_parsed_analysis_file,
)

# ---------------------------------------------------------------------------
# Column headers
# ---------------------------------------------------------------------------

OVERVIEW_HEADERS = [
    'Story', 'Label', 'Type',
    'Load Type',
    'Section (Exist)', 'Section (New)',
    'PMM (Exist)', 'PMM (New)', 'PMM (%)',
    'V Major D/C',
    'Worst Force (%)',
    'Flag Reason',
    'Result',
]

DETAIL_HEADERS = [
    'Story', 'Label', 'Type',
    'Section', 'Gov Combo', 'Load Type',
    'P (Exist)', 'P (New)', 'P (%)',
    'V2 (Exist)', 'V2 (New)', 'V2 (%)',
    'V3 (Exist)', 'V3 (New)', 'V3 (%)',
    'M2 (Exist)', 'M2 (New)', 'M2 (%)',
    'M3 (Exist)', 'M3 (New)', 'M3 (%)',
    'Net Demand', 'Flag Reason', 'Result',
]

SUMMARY_HEADERS = ['Story', 'Type', 'Total', 'PASS', 'WARN', 'FLAG', 'ADDED', 'REMOVED']

ANALYSIS_OVERVIEW_HEADERS = [
    'Story', 'Label', 'Type',
    'P (New)', 'P (%)', 'V2 (New)', 'V2 (%)', 'M2 (New)', 'M2 (%)', 'M3 (New)', 'M3 (%)',
    'Worst (%)', 'Flag Reason', 'Result',
]

ANALYSIS_DETAIL_HEADERS = [
    'Story', 'Label', 'Type',
    'P (Exist)', 'P (New)', 'P (%)',
    'V2 (Exist)', 'V2 (New)', 'V2 (%)',
    'M2 (Exist)', 'M2 (New)', 'M2 (%)',
    'M3 (Exist)', 'M3 (New)', 'M3 (%)',
    'Worst (%)', 'Flag Reason', 'Result',
]

# ---------------------------------------------------------------------------
# Colors
# ---------------------------------------------------------------------------

_RESULT_COLORS = {
    'PASS':    vkt.Color(34, 139, 34),
    'FLAG':    vkt.Color(178, 34, 34),
    'WARN':    vkt.Color(180, 130, 0),
    'ADDED':   vkt.Color(30, 100, 200),
    'REMOVED': vkt.Color(180, 100, 0),
}

_CHART_COLORS = {
    'FLAG':    '#B22222',
    'WARN':    '#B48200',
    'PASS':    '#228B22',
    'ADDED':   '#1E64C8',
    'REMOVED': '#B46400',
}

# Member type helpers for typed tabs
_TYPE_SINGULAR = {'Columns': 'Column', 'Beams': 'Beam', 'Braces': 'Brace'}
_GOVERNING_FORCES_BY_TYPE = {
    'Columns': ['P', 'M3', 'M2'],
    'Beams':   ['M3', 'V2'],
    'Braces':  ['P'],
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


def _net_demand(r: dict) -> str:
    """UP / DOWN / MIXED based on direction of primary force changes (excludes INF)."""
    if r.get('Pass') in ('ADDED', 'REMOVED'):
        return 'N/A'
    pcts = [
        r.get('P_Pct'), r.get('V2_Pct'), r.get('V3_Pct'),
        r.get('M2_Pct'), r.get('M3_Pct'), r.get('PMM_Pct'),
    ]
    numeric = [p for p in pcts if isinstance(p, float)]
    if not numeric:
        return 'N/A'
    if all(p <= 0 for p in numeric):
        return 'DOWN'
    if all(p > 0 for p in numeric):
        return 'UP'
    return 'MIXED'


_CAPACITY_WARN_THRESHOLD = 0.95  # PMM below this + FLAG → WARN (member still under capacity)


def _postprocess_results(results: list) -> list:
    """Add NetDemand; downgrade FLAG → WARN for two conditions:
    1. All numeric forces decreased (INF-only failure on previously-zero component).
    2. Threshold exceeded but modified PMM D/C ratio is still < 0.95 (below capacity).
    """
    out = []
    for r in results:
        r = dict(r)
        nd = _net_demand(r)
        r['NetDemand'] = nd
        if r.get('Pass') == 'FLAG':
            pmm_new = r.get('PMM_New')
            forces_down = nd == 'DOWN'
            below_capacity = isinstance(pmm_new, float) and pmm_new < _CAPACITY_WARN_THRESHOLD
            if forces_down or below_capacity:
                r['Pass'] = 'WARN'
        out.append(r)
    return out


def _recompute_typed_fail_reason(
    r: dict, gov_forces: list, force_thresholds: dict,
    grav_thresh: float, lat_thresh: float,
    warn_thresh: float, fail_thresh: float,
    is_analysis: bool, status: str,
) -> dict:
    """Return a copy of r with FailReason restricted to gov_forces above their abs threshold.

    Prevents low-magnitude forces (e.g. M2 = 0.3 kip-ft) or non-governing forces
    (e.g. M2 for beams) from appearing as the primary flag reason in the typed tabs.
    """
    if is_analysis:
        pct_threshold = fail_thresh if status == 'FLAG' else warn_thresh
        load_type = None
        sign_tag = ''
    else:
        load_type = r.get('LoadType', 'gravity')
        pct_threshold = lat_thresh if load_type == 'lateral' else grav_thresh
        sign_tag = ' [sign rev]' if r.get('SignReversal') == 'YES' else ''

    def _max_abs(f):
        return max(
            (abs(r[k]) for k in (f'{f}_Exist', f'{f}_New') if isinstance(r.get(k), (int, float))),
            default=0.0,
        )

    # Priority 1: INF change on a gov force above its abs threshold
    for f in gov_forces:
        if r.get(f'{f}_Pct') == 'INF' and _max_abs(f) >= force_thresholds.get(f, 0.0):
            new_reason = (
                f'{f} INF > {pct_threshold:.0f}%' if is_analysis
                else f'{f} INF ({load_type}){sign_tag}'
            )
            return {**r, 'FailReason': new_reason}

    # Priority 2: worst % exceeding pct_threshold on a gov force above its abs threshold
    worst_label = None
    worst_pct_val = 0.0
    for f in gov_forces:
        pct = r.get(f'{f}_Pct')
        if not isinstance(pct, float) or pct <= pct_threshold:
            continue
        if _max_abs(f) >= force_thresholds.get(f, 0.0) and pct > worst_pct_val:
            worst_pct_val = pct
            worst_label = f

    if worst_label is not None:
        new_reason = (
            f'{worst_label} +{worst_pct_val:.1f}% > {pct_threshold:.0f}%' if is_analysis
            else f'{worst_label} +{worst_pct_val:.1f}% > {pct_threshold:.0f}% {load_type}{sign_tag}'
        )
        return {**r, 'FailReason': new_reason}

    # No qualifying force found among gov_forces above abs threshold
    return {**r, 'FailReason': ''}


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
## Upload ETABS Output Exports

Upload two ETABS Excel exports (.xlsx) — one for the **existing** model and
one for the **modified** (repurposed) model. Select the result type below to
match the tables you exported.

**Design Results** — post-design-run tables. Export: *Design Forces - Columns*,
*Design Forces - Beams*, *Design Forces - Braces*, and *Steel Frame Design
Summary - AISC 360-16* to a single workbook. Compares PMM/shear D/C ratios
and force envelopes; classifies combos as gravity or lateral for IBC 3403.

**Analysis Results** — pre-design element force tables. Export: *Element
Forces - Columns*, *Element Forces - Beams*, *Element Forces - Braces* to a
single workbook. Compares max-absolute force envelopes (P, M3/M2, V2) against
a single adjustable threshold.

The first run will take ~1–2 minutes to parse; subsequent tab switches are
near-instant.
""")

    step1.mode = vkt.OptionField(
        'Result Type',
        options=['Design Results', 'Analysis Results'],
        default='Design Results',
        description='Match this to the ETABS tables you exported',
    )

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
        'braces_table',
        'columns_table',
        'beams_table',
        'key_metrics',
        'results_table',
        'results_detail',
        'results_chart',
        'summary_table',
        'beam_detail_view',
        'member_detail_view',
    ])

    step2.section_about = vkt.Section('How This Works')

    step2.section_about.intro_method = vkt.Text("""
## How This Comparison Works

This tool compares structural demand on every steel member between the **existing** and **modified** models to check compliance with **IBC Section 3403**, which governs force increases in members of existing buildings being modified or repurposed.

**Member matching** is performed by Story + Label (the ETABS frame label). Members present in only one model are shown as ADDED or REMOVED — no threshold check is applied to them.

For each matched member the tool finds the **worst-case force across all ETABS design load combinations and stations**, then computes the percent change as (abs(new) − abs(exist)) / abs(exist) × 100. Using absolute values ensures a sign reversal (tension flipping to compression) does not artificially inflate the percentage. Sign reversals are flagged separately in the Sign Rev. column.

**INF** in a % change column means the existing model force was effectively zero (< 0.01) while the new model has a non-zero force — the ratio is mathematically infinite. This most often indicates a newly loaded member. INF alone does not necessarily require a response; evaluate the absolute magnitude.

**Net Demand** (UP / DOWN / MIXED) summarizes whether the governing forces increased, decreased, or split direction across the member. DOWN members are hidden from the typed tabs by default since they represent reduced demand.
""")

    step2.section_about.intro_columns = vkt.Text("""
### What the columns mean

**P** - Axial force (kip). Positive = tension, negative = compression.

**V2 / V3** - Shear in the local 2- and 3-axis directions (kip).

**T** - Torsion (kip-ft).

**M2 / M3** - Minor- and major-axis bending moments (kip-ft).

**PMM** - P-M-M interaction ratio. This is the AISC 360-16 Equation H1-1 combined demand/capacity (D/C) ratio for axial force plus biaxial bending acting together. A value of 1.0 means the member is exactly at capacity. Values above 1.0 mean the member is overstressed. This is the primary design check for columns and braces.

**V Major** - Major-axis shear D/C ratio from the ETABS Steel Frame Design Summary.

**Worst Force %** - The largest percent increase among P, V2, V3, T, M2, M3 for that member.

**Flag Reason** - Plain-English description of what tripped the FLAG threshold, e.g. "M3 +23.5% > 5% gravity". FLAG does not imply a design failure — use engineering judgment.
""")

    step2.section_about.intro_thresholds = vkt.Text("""
### Status and Load Classification (IBC 3403)

A member is **FLAG**ged if any governing force or D/C ratio increases by more than the threshold for its load type:

- **Gravity combos** (dead, live, snow, roof live) — default 5% increase triggers a flag
- **Lateral combos** (wind, seismic) — default 10% increase triggers a flag

Load type is determined from the governing combo name in the modified model. Combos with wind tokens (WA, WB, WG) or seismic tokens (EQ, EQB) are lateral; all others are gravity. Thresholds can be adjusted in Comparison Options below.

**FLAG** — increase exceeded the threshold. Does not imply a design failure; use engineering judgment.

**WARN** — increase exceeded the threshold, but demand decreased overall or the member is well below capacity.

**ADDED / REMOVED** — member exists only in one model; no threshold check applied.

Start with the **Braces**, **Columns**, or **Beams** tab for focused review. Each shows only the relevant forces for that member class and filters out demand decreases and low-magnitude noise. Use **Overview** or **Full Detail** for the full unfiltered picture.
""")

    step2.section_options = vkt.Section('Comparison Options')

    step2.section_options.lbl_design = vkt.Text('**Design Mode — IBC 3403 Thresholds**')
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

    step2.section_options.lbl_analysis = vkt.Text('**Analysis Mode — Flag Thresholds**')
    step2.section_options.warn_threshold = vkt.NumberField(
        'Warn Threshold (%)',
        default=5,
        min=0,
        description='Force increase above this triggers WARN',
    )
    step2.section_options.fail_threshold = vkt.NumberField(
        'Flag Threshold (%)',
        default=10,
        min=0,
        description='Force increase above this triggers FLAG',
    )

    step2.section_options.lbl_abs = vkt.Text('**Absolute Force Minimums — Typed Tabs**')
    step2.section_options.thresh_P = vkt.NumberField(
        'Min |P| to show (kips)',
        default=5.0,
        min=0,
        description='Members excluded from Braces/Columns tabs if |P| stays below this in both models.',
    )
    step2.section_options.thresh_M3 = vkt.NumberField(
        'Min |M3| to show (kip-ft)',
        default=20.0,
        min=0,
        description='Members excluded from Beams/Columns tabs if |M3| stays below this in both models.',
    )
    step2.section_options.thresh_M2 = vkt.NumberField(
        'Min |M2| to show (kip-ft)',
        default=10.0,
        min=0,
        description='Members excluded from Columns tabs if |M2| stays below this in both models.',
    )
    step2.section_options.thresh_V2 = vkt.NumberField(
        'Min |V2| to show (kips)',
        default=5.0,
        min=0,
        description='Members excluded from Beams tabs if |V2| stays below this in both models.',
    )
    step2.section_options.thresh_V3 = vkt.NumberField(
        'Min |V3| to show (kips)',
        default=0.0,
        min=0,
        description='Members excluded from Columns tabs if |V3| stays below this. Set to 0 to disable.',
    )

    step2.section_options.lbl_display = vkt.Text('**Display Options**')
    step2.section_options.hide_decreases = vkt.BooleanField(
        'Hide members where demand decreased',
        default=True,
        description=(
            'When checked, members where all governing forces decreased are excluded '
            'from the Braces / Columns / Beams tabs.'
        ),
    )
    step2.section_options.show_added_removed = vkt.BooleanField(
        'Show added / removed members',
        default=True,
        description='When checked, members that exist only in one model appear in the typed tabs.',
    )
    step2.section_options.display_filter = vkt.OptionField(
        'Display Filter — Overview / Full Detail tabs',
        options=['All Results', 'Failures Only'],
        default='Failures Only',
    )

    step2.section_export = vkt.Section('Export')
    step2.section_export.project_name = vkt.TextField(
        'Project Name (optional)',
        description='Included in the CSV filename, e.g. "BuildingA" → etabs_BuildingA_columns.csv',
    )
    step2.section_export.member_type_export = vkt.OptionField(
        'Export Member Type',
        options=['All', 'Columns', 'Beams', 'Braces'],
        default='All',
        description='Filter CSV export to one member class for manual strike-through review',
    )
    step2.section_export.download_btn = vkt.DownloadButton(
        'Export Results to CSV',
        method='download_csv',
        longpoll=True,
    )

    step2.section_detail = vkt.Section('Member Detail')
    step2.section_detail.detail_intro = vkt.Text(
        'Enter a member label to view a full detail card for that member. '
        'The label must match exactly as it appears in the comparison results (e.g. "C12", "B5").'
    )
    step2.section_detail.detail_label = vkt.TextField(
        'Member Label',
        description='Label to look up (e.g. "C12"). Case-insensitive.',
    )
    step2.section_detail.detail_story = vkt.TextField(
        'Story (optional)',
        description='Specify story if the label appears on multiple stories.',
    )

    step2.section_ai = vkt.Section('AI Assistant')
    step2.section_ai.chat = vkt.Chat(
        'Ask the Structural Assistant',
        method='call_llm',
        first_message=(
            'I can see the comparison results for this model. '
            'Ask me about which members failed, why they failed, '
            'or what the most critical issues are.'
        ),
    )


# ---------------------------------------------------------------------------
# Diagnostics
# ---------------------------------------------------------------------------

def _no_results_message(params) -> str:
    """Return a descriptive error message explaining why no results were produced."""
    _, pe = _get_parsed_file(params.step1.existing_file.file)
    _, pn = _get_parsed_file(params.step1.modified_file.file)

    e_sheets = pe.get('sheet_names', [])
    n_sheets = pn.get('sheet_names', [])

    expected = list(DESIGN_FORCES_SHEETS.values())
    e_missing = [s for s in expected if s not in e_sheets]
    n_missing = [s for s in expected if s not in n_sheets]

    if e_missing or n_missing:
        lines = [f'No members found. Expected sheets: {expected}']
        if e_missing:
            lines.append(f'Existing file missing: {e_missing}  (found: {e_sheets})')
        if n_missing:
            lines.append(f'Modified file missing: {n_missing}  (found: {n_sheets})')
        return '\n'.join(lines)

    e_counts = {mt: len(pe['forces'][mt]) for mt in DESIGN_FORCES_SHEETS}
    n_counts = {mt: len(pn['forces'][mt]) for mt in DESIGN_FORCES_SHEETS}
    if all(v == 0 for v in e_counts.values()) or all(v == 0 for v in n_counts.values()):
        return (
            f'Forces sheets parsed but returned 0 rows.\n'
            f'Existing: {e_counts}\nModified: {n_counts}'
        )

    return (
        'No results to display. If "Failures Only" is selected, '
        'all members may be passing. Try switching to "All Results".'
    )


def _no_results_message_analysis(params) -> str:
    _, pe = _get_parsed_analysis_file(params.step1.existing_file.file)
    _, pn = _get_parsed_analysis_file(params.step1.modified_file.file)

    e_sheets = pe.get('sheet_names', [])
    n_sheets = pn.get('sheet_names', [])

    expected = list(ANALYSIS_FORCES_SHEETS.values())
    e_missing = [s for s in expected if s not in e_sheets]
    n_missing = [s for s in expected if s not in n_sheets]

    if e_missing or n_missing:
        lines = [f'No members found. Expected sheets: {expected}']
        if e_missing:
            lines.append(f'Existing file missing: {e_missing}  (found: {e_sheets})')
        if n_missing:
            lines.append(f'Modified file missing: {n_missing}  (found: {n_sheets})')
        return '\n'.join(lines)

    return (
        'No results to display. If "Failures Only" is selected, '
        'all members may be passing. Try switching to "All Results".'
    )


# ---------------------------------------------------------------------------
# Beam detail HTML renderer
# ---------------------------------------------------------------------------

def _numeric_pct(val) -> float:
    """Convert a pct value (float, 'INF', or other) to a float for comparison."""
    if val == 'INF':
        return 999.0
    if isinstance(val, float):
        return val
    return 0.0


def _render_beam_detail_html(beam: Optional[dict], gravity_thresh: float, lateral_thresh: float) -> str:
    if beam is None:
        return (
            '<p style="font-family:sans-serif;padding:24px;color:#555;font-size:15px;">'
            'No beam results to display. Ensure beams are included in the member type filter '
            'and both files are uploaded.</p>'
        )

    load_type = beam.get('LoadType', 'gravity')
    threshold = lateral_thresh if load_type == 'lateral' else gravity_thresh
    pass_val  = beam.get('Pass', '')
    label     = beam.get('Label', '')
    story     = beam.get('Story', '')
    sect_e    = beam.get('DesignSection_Exist', 'N/A')
    sect_n    = beam.get('DesignSection_New', 'N/A')
    combo_e   = beam.get('GovCombo_Exist', 'N/A')
    combo_n   = beam.get('GovCombo_New', 'N/A')
    fail_rsn  = beam.get('FailReason', '')
    sign_rev  = beam.get('SignReversal', '')

    banner_bg = '#B22222' if pass_val == 'FLAG' else '#228B22'

    if pass_val == 'FLAG':
        result_phrase = '<span style="font-weight:bold">FLAGGED</span>'
        narrative = (
            f'Beam <strong>{label}</strong> on Story <strong>{story}</strong> {result_phrase} — '
            f'{fail_rsn}. Review required; FLAG does not imply a design failure.'
        )
    else:
        result_phrase = '<span style="font-weight:bold">PASSED</span>'
        narrative = (
            f'Beam <strong>{label}</strong> on Story <strong>{story}</strong> {result_phrase} '
            f'all threshold checks ({load_type} threshold: {threshold:.1f}%).'
        )
    if sect_e != sect_n:
        narrative += f' Section changed from <strong>{sect_e}</strong> to <strong>{sect_n}</strong>.'
    else:
        narrative += f' Section unchanged: <strong>{sect_e}</strong>.'
    if sign_rev == 'YES':
        narrative += ' <span style="color:#B46400;font-weight:bold">&#9888; Sign reversal detected on at least one force component.</span>'

    TH = 'padding:8px 12px;text-align:left;font-size:13px;border-bottom:2px solid #ccc;white-space:nowrap'
    TD = 'padding:7px 12px;font-size:13px;border-bottom:1px solid #e5e5e5'

    def _fmt(val) -> str:
        if isinstance(val, float):
            return f'{val:.3f}'
        return str(val) if val not in (None, '') else 'N/A'

    def _fmt_change(pct_val) -> str:
        if pct_val == 'INF':
            return 'INF'
        if isinstance(pct_val, float):
            sign = '+' if pct_val >= 0 else ''
            return f'{sign}{pct_val:.1f}%'
        return 'N/A'

    def _force_row(name: str, exist_val, new_val, pct_val, thresh: float) -> str:
        is_fail = (pct_val == 'INF') or (isinstance(pct_val, float) and pct_val > thresh)
        row_bg  = '#fff0f0' if is_fail else 'transparent'
        status  = 'FLAG' if is_fail else 'PASS'
        sc      = '#B22222' if is_fail else '#228B22'
        return (
            f'<tr style="background:{row_bg}">'
            f'<td style="{TD};font-weight:bold">{name}</td>'
            f'<td style="{TD}">{_fmt(exist_val)}</td>'
            f'<td style="{TD}">{_fmt(new_val)}</td>'
            f'<td style="{TD}">{_fmt_change(pct_val)}</td>'
            f'<td style="{TD}">{thresh:.1f}%</td>'
            f'<td style="{TD};color:{sc};font-weight:bold">{status}</td>'
            f'</tr>'
        )

    force_rows = ''.join(
        _force_row(
            c,
            beam.get(f'{c}_Exist'), beam.get(f'{c}_New'), beam.get(f'{c}_Pct'), threshold
        )
        for c in FORCE_COMPONENTS
    )

    def _dc_row(name: str, combo_exist: str, combo_new: str, exist_val, new_val, pct_val, thresh: float) -> str:
        is_fail = (pct_val == 'INF') or (isinstance(pct_val, float) and pct_val > thresh)
        row_bg  = '#fff0f0' if is_fail else 'transparent'
        status  = 'FLAG' if is_fail else 'PASS'
        sc      = '#B22222' if is_fail else '#228B22'
        ce = combo_exist if combo_exist and combo_exist != 'N/A' else '—'
        cn = combo_new   if combo_new   and combo_new   != 'N/A' else '—'
        return (
            f'<tr style="background:{row_bg}">'
            f'<td style="{TD};font-weight:bold">{name}</td>'
            f'<td style="{TD};font-family:monospace;font-size:12px">{ce}</td>'
            f'<td style="{TD};font-family:monospace;font-size:12px">{cn}</td>'
            f'<td style="{TD}">{_fmt(exist_val)}</td>'
            f'<td style="{TD}">{_fmt(new_val)}</td>'
            f'<td style="{TD}">{_fmt_change(pct_val)}</td>'
            f'<td style="{TD}">{thresh:.1f}%</td>'
            f'<td style="{TD};color:{sc};font-weight:bold">{status}</td>'
            f'</tr>'
        )

    dc_rows = (
        _dc_row('PMM', combo_e, combo_n,
                beam.get('PMM_Exist'), beam.get('PMM_New'), beam.get('PMM_Pct'), threshold) +
        _dc_row('V Major', combo_e, combo_n,
                beam.get('VMaj_Exist'), beam.get('VMaj_New'), beam.get('VMaj_Pct'), threshold)
    )

    card_style = 'background:#f7f7f7;padding:12px 16px;border-radius:5px;border:1px solid #e0e0e0'
    label_style = 'color:#777;font-size:11px;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:4px'
    value_style = 'font-size:15px;font-weight:500'

    sect_arrow = f'{sect_e} &rarr; {sect_n}' if sect_e != sect_n else sect_e

    html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;padding:24px;max-width:960px;color:#222;margin:0 auto">

<div style="background:{banner_bg};color:#fff;padding:12px 20px;border-radius:6px;margin-bottom:18px;font-size:17px;font-weight:bold">
  {pass_val} &mdash; Beam {label}, Story {story}
  <span style="font-weight:normal;font-size:13px;margin-left:16px;opacity:0.9">Worst beam auto-selected</span>
</div>

<p style="font-size:14px;line-height:1.65;margin-bottom:20px;padding:12px 16px;background:#fafafa;border-left:4px solid {banner_bg};border-radius:0 4px 4px 0">
  {narrative}
</p>

<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:24px">
  <div style="{card_style}">
    <div style="{label_style}">Design Section</div>
    <div style="{value_style}">{sect_arrow}</div>
  </div>
  <div style="{card_style}">
    <div style="{label_style}">Load Type / Threshold</div>
    <div style="{value_style}">{load_type.capitalize()} &mdash; {threshold:.1f}%</div>
  </div>
  <div style="{card_style}">
    <div style="{label_style}">Governing Combo (Existing)</div>
    <div style="font-size:13px;font-family:monospace">{combo_e}</div>
  </div>
  <div style="{card_style}">
    <div style="{label_style}">Governing Combo (Modified)</div>
    <div style="font-size:13px;font-family:monospace">{combo_n}</div>
  </div>
</div>

<h3 style="font-size:14px;text-transform:uppercase;letter-spacing:0.05em;color:#555;margin-bottom:8px">Force Components</h3>
<table style="width:100%;border-collapse:collapse;margin-bottom:28px">
  <thead>
    <tr style="background:#f0f0f0">
      <th style="{TH}">Force</th>
      <th style="{TH}">Existing</th>
      <th style="{TH}">Modified</th>
      <th style="{TH}">Change</th>
      <th style="{TH}">Threshold</th>
      <th style="{TH}">Status</th>
    </tr>
  </thead>
  <tbody>{force_rows}</tbody>
</table>

<h3 style="font-size:14px;text-transform:uppercase;letter-spacing:0.05em;color:#555;margin-bottom:8px">Demand / Capacity Ratios</h3>
<table style="width:100%;border-collapse:collapse">
  <thead>
    <tr style="background:#f0f0f0">
      <th style="{TH}">Check</th>
      <th style="{TH}">Combo (Existing)</th>
      <th style="{TH}">Combo (Modified)</th>
      <th style="{TH}">Existing</th>
      <th style="{TH}">Modified</th>
      <th style="{TH}">Change</th>
      <th style="{TH}">Threshold</th>
      <th style="{TH}">Status</th>
    </tr>
  </thead>
  <tbody>{dc_rows}</tbody>
</table>

</body></html>"""
    return html


# ---------------------------------------------------------------------------
# Generic member detail HTML renderer
# ---------------------------------------------------------------------------

_BANNER_COLORS = {
    'FLAG':    '#B22222',
    'WARN':    '#B48200',
    'PASS':    '#228B22',
    'ADDED':   '#1E64C8',
    'REMOVED': '#B46400',
}


def _render_member_detail_html(
    member: Optional[dict],
    gravity_thresh: float,
    lateral_thresh: float,
    mode: str,
) -> str:
    if member is None:
        return (
            '<p style="font-family:sans-serif;padding:24px;color:#555;font-size:15px;">'
            'Enter a member label in <strong>Member Detail</strong> above to view its '
            'detail card. The label must match exactly as shown in the results tables '
            '(e.g. "C12", "B5"). Optionally enter a story to disambiguate.</p>'
        )

    is_analysis = (mode == 'Analysis Results')
    load_type   = member.get('LoadType', 'gravity') if not is_analysis else None
    threshold   = (lateral_thresh if load_type == 'lateral' else gravity_thresh) if not is_analysis else gravity_thresh
    pass_val    = member.get('Pass', '')
    label       = member.get('Label', '')
    story       = member.get('Story', '')
    mtype       = member.get('MemberType', '')
    sect_e      = member.get('DesignSection_Exist', 'N/A') or 'N/A'
    sect_n      = member.get('DesignSection_New',   'N/A') or 'N/A'
    combo_e     = member.get('GovCombo_Exist', '') or 'N/A'
    combo_n     = member.get('GovCombo_New',   '') or 'N/A'
    fail_rsn    = member.get('FailReason', '')
    sign_rev    = member.get('SignReversal', '')
    net_demand  = member.get('NetDemand', '')
    banner_bg   = _BANNER_COLORS.get(pass_val, '#888888')

    # Narrative
    if pass_val in ('FLAG', 'WARN'):
        narrative = (
            f'{mtype} <strong>{label}</strong>, Story <strong>{story}</strong> — '
            f'<strong>{pass_val}</strong>. {fail_rsn}.'
        )
        if pass_val == 'WARN':
            narrative += ' Threshold exceeded but demand is within acceptable range — review advised.'
    elif pass_val == 'PASS':
        narrative = (
            f'{mtype} <strong>{label}</strong>, Story <strong>{story}</strong> — '
            f'<strong>PASS</strong>. All force components within threshold.'
        )
    else:
        narrative = (
            f'{mtype} <strong>{label}</strong>, Story <strong>{story}</strong> — '
            f'<strong>{pass_val}</strong>.'
        )

    if not is_analysis and sect_e != sect_n and sect_e != 'N/A' and sect_n != 'N/A':
        narrative += f' Section changed: <strong>{sect_e}</strong> &rarr; <strong>{sect_n}</strong>.'
    elif not is_analysis and sect_e != 'N/A':
        narrative += f' Section: <strong>{sect_e}</strong>.'
    if sign_rev == 'YES':
        narrative += ' <span style="color:#B46400;font-weight:bold">&#9888; Sign reversal on at least one force component.</span>'

    TH = 'padding:8px 12px;text-align:left;font-size:13px;border-bottom:2px solid #ccc;white-space:nowrap'
    TD = 'padding:7px 12px;font-size:13px;border-bottom:1px solid #e5e5e5'

    def _fmt(val) -> str:
        if isinstance(val, float):
            return f'{val:.3f}'
        return str(val) if val not in (None, '') else 'N/A'

    def _fmt_change(pct_val) -> str:
        if pct_val == 'INF':
            return 'INF'
        if isinstance(pct_val, float):
            sign = '+' if pct_val >= 0 else ''
            return f'{sign}{pct_val:.1f}%'
        return str(pct_val) if pct_val not in (None, '') else 'N/A'

    def _force_row(name: str, exist_val, new_val, pct_val, thresh: float) -> str:
        no_data = pct_val in ('N/A', None, '')
        is_fail = not no_data and ((pct_val == 'INF') or (isinstance(pct_val, float) and pct_val > thresh))
        row_bg  = '#fff0f0' if is_fail else 'transparent'
        status  = 'FLAG' if is_fail else ('PASS' if not no_data else '—')
        sc      = '#B22222' if is_fail else ('#228B22' if status == 'PASS' else '#aaa')
        return (
            f'<tr style="background:{row_bg}">'
            f'<td style="{TD};font-weight:bold">{name}</td>'
            f'<td style="{TD}">{_fmt(exist_val)}</td>'
            f'<td style="{TD}">{_fmt(new_val)}</td>'
            f'<td style="{TD}">{_fmt_change(pct_val)}</td>'
            f'<td style="{TD}">{thresh:.1f}%</td>'
            f'<td style="{TD};color:{sc};font-weight:bold">{status}</td>'
            f'</tr>'
        )

    # Show all FORCE_COMPONENTS; in analysis mode skip slots where both sides are N/A
    forces_to_show = [
        c for c in FORCE_COMPONENTS
        if not is_analysis or not (
            member.get(f'{c}_Exist') in ('N/A', None, '')
            and member.get(f'{c}_New') in ('N/A', None, '')
        )
    ]
    force_rows = ''.join(
        _force_row(c, member.get(f'{c}_Exist'), member.get(f'{c}_New'), member.get(f'{c}_Pct'), threshold)
        for c in forces_to_show
    )

    # D/C section — design mode only
    dc_section = ''
    if not is_analysis and pass_val not in ('ADDED', 'REMOVED'):
        def _dc_row(name, exist_val, new_val, pct_val, thresh) -> str:
            no_data = pct_val in ('N/A', None, '')
            is_fail = not no_data and ((pct_val == 'INF') or (isinstance(pct_val, float) and pct_val > thresh))
            row_bg  = '#fff0f0' if is_fail else 'transparent'
            status  = 'FLAG' if is_fail else ('PASS' if not no_data else '—')
            sc      = '#B22222' if is_fail else ('#228B22' if status == 'PASS' else '#aaa')
            return (
                f'<tr style="background:{row_bg}">'
                f'<td style="{TD};font-weight:bold">{name}</td>'
                f'<td style="{TD}">{_fmt(exist_val)}</td>'
                f'<td style="{TD}">{_fmt(new_val)}</td>'
                f'<td style="{TD}">{_fmt_change(pct_val)}</td>'
                f'<td style="{TD}">{thresh:.1f}%</td>'
                f'<td style="{TD};color:{sc};font-weight:bold">{status}</td>'
                f'</tr>'
            )
        dc_rows_html = (
            _dc_row('PMM',     member.get('PMM_Exist'),  member.get('PMM_New'),  member.get('PMM_Pct'),  threshold) +
            _dc_row('V Major', member.get('VMaj_Exist'), member.get('VMaj_New'), member.get('VMaj_Pct'), threshold)
        )
        dc_section = f"""
<h3 style="font-size:14px;text-transform:uppercase;letter-spacing:0.05em;color:#555;margin:24px 0 8px">Demand / Capacity Ratios</h3>
<table style="width:100%;border-collapse:collapse">
  <thead>
    <tr style="background:#f0f0f0">
      <th style="{TH}">Check</th><th style="{TH}">Existing</th><th style="{TH}">Modified</th>
      <th style="{TH}">Change</th><th style="{TH}">Threshold</th><th style="{TH}">Status</th>
    </tr>
  </thead>
  <tbody>{dc_rows_html}</tbody>
</table>"""

    # Summary cards
    card_s = 'background:#f7f7f7;padding:12px 16px;border-radius:5px;border:1px solid #e0e0e0'
    lbl_s  = 'color:#777;font-size:11px;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:4px'
    val_s  = 'font-size:15px;font-weight:500'

    cards_html = f'<div style="{card_s}"><div style="{lbl_s}">Member Type</div><div style="{val_s}">{mtype}</div></div>'
    cards_html += f'<div style="{card_s}"><div style="{lbl_s}">Story</div><div style="{val_s}">{story}</div></div>'

    if not is_analysis:
        sect_arrow = f'{sect_e} &rarr; {sect_n}' if sect_e != sect_n else sect_e
        cards_html += f'<div style="{card_s}"><div style="{lbl_s}">Design Section</div><div style="{val_s}">{sect_arrow}</div></div>'
        cards_html += (
            f'<div style="{card_s}"><div style="{lbl_s}">Load Type / Threshold</div>'
            f'<div style="{val_s}">{(load_type or "").capitalize()} &mdash; {threshold:.1f}%</div></div>'
        )

    cards_html += (
        f'<div style="{card_s}"><div style="{lbl_s}">Governing Combo (Existing)</div>'
        f'<div style="font-size:13px;font-family:monospace">{combo_e}</div></div>'
        f'<div style="{card_s}"><div style="{lbl_s}">Governing Combo (Modified)</div>'
        f'<div style="font-size:13px;font-family:monospace">{combo_n}</div></div>'
    )

    if net_demand and net_demand != 'N/A':
        nd_color = '#228B22' if net_demand == 'DOWN' else ('#B22222' if net_demand == 'UP' else '#B48200')
        cards_html += (
            f'<div style="{card_s}"><div style="{lbl_s}">Net Demand</div>'
            f'<div style="{val_s};color:{nd_color}">{net_demand}</div></div>'
        )

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;padding:24px;max-width:960px;color:#222;margin:0 auto">

<div style="background:{banner_bg};color:#fff;padding:12px 20px;border-radius:6px;margin-bottom:18px;font-size:17px;font-weight:bold">
  {pass_val} &mdash; {mtype} {label}, Story {story}
</div>

<p style="font-size:14px;line-height:1.65;margin-bottom:20px;padding:12px 16px;background:#fafafa;border-left:4px solid {banner_bg};border-radius:0 4px 4px 0">
  {narrative}
</p>

<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:24px">
  {cards_html}
</div>

<h3 style="font-size:14px;text-transform:uppercase;letter-spacing:0.05em;color:#555;margin-bottom:8px">Force Components</h3>
<table style="width:100%;border-collapse:collapse;margin-bottom:4px">
  <thead>
    <tr style="background:#f0f0f0">
      <th style="{TH}">Force</th><th style="{TH}">Existing</th><th style="{TH}">Modified</th>
      <th style="{TH}">Change</th><th style="{TH}">Threshold</th><th style="{TH}">Status</th>
    </tr>
  </thead>
  <tbody>{force_rows}</tbody>
</table>
{dc_section}

</body></html>"""


# ---------------------------------------------------------------------------
# Controller
# ---------------------------------------------------------------------------

class Controller(vkt.Controller):
    parametrization = Parametrization

    # -- Overview table (primary view) ---------------------------------------

    @vkt.TableView('Overview', duration_guess=2)
    def results_table(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')

        if (params.step1.mode or 'Design Results') == 'Analysis Results':
            results = self._run_analysis(params)
            if not results:
                raise vkt.UserError(_no_results_message_analysis(params))
            data = []
            for r in results:
                data.append([
                    r.get('Story', ''),
                    r.get('Label', ''),
                    r.get('MemberType', ''),
                    r.get('P_New', ''),
                    _fmt_pct(r.get('P_Pct', '')),
                    r.get('V2_New', ''),
                    _fmt_pct(r.get('V2_Pct', '')),
                    r.get('M2_New', ''),
                    _fmt_pct(r.get('M2_Pct', '')),
                    r.get('M3_New', ''),
                    _fmt_pct(r.get('M3_Pct', '')),
                    _fmt_pct(r.get('WorstPct', '')),
                    r.get('FailReason', ''),
                    _result_cell(r.get('Pass', '')),
                ])
            return vkt.TableResult(data, column_headers=ANALYSIS_OVERVIEW_HEADERS)

        results = self._run(params)
        if not results:
            raise vkt.UserError(_no_results_message(params))

        data = []
        for r in results:
            data.append([
                r.get('Story', ''),
                r.get('Label', ''),
                r.get('MemberType', ''),
                r.get('LoadType', ''),
                r.get('DesignSection_Exist', ''),
                r.get('DesignSection_New', ''),
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

    @vkt.TableView('Full Detail', duration_guess=2)
    def results_detail(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')

        if (params.step1.mode or 'Design Results') == 'Analysis Results':
            results = self._run_analysis(params)
            if not results:
                raise vkt.UserError(_no_results_message_analysis(params))
            data = []
            for r in results:
                data.append([
                    r.get('Story', ''),
                    r.get('Label', ''),
                    r.get('MemberType', ''),
                    r.get('P_Exist', ''), r.get('P_New', ''), _fmt_pct(r.get('P_Pct', '')),
                    r.get('V2_Exist', ''), r.get('V2_New', ''), _fmt_pct(r.get('V2_Pct', '')),
                    r.get('M2_Exist', ''), r.get('M2_New', ''), _fmt_pct(r.get('M2_Pct', '')),
                    r.get('M3_Exist', ''), r.get('M3_New', ''), _fmt_pct(r.get('M3_Pct', '')),
                    _fmt_pct(r.get('WorstPct', '')),
                    r.get('FailReason', ''),
                    _result_cell(r.get('Pass', '')),
                ])
            return vkt.TableResult(data, column_headers=ANALYSIS_DETAIL_HEADERS)

        results = self._run(params)
        if not results:
            raise vkt.UserError(_no_results_message(params))

        data = []
        for r in results:
            section = r.get('DesignSection_New') or r.get('DesignSection_Exist', '')
            data.append([
                r.get('Story', ''),
                r.get('Label', ''),
                r.get('MemberType', ''),
                section,
                r.get('GovCombo_New', ''),
                r.get('LoadType', ''),
                r.get('P_Exist', ''), r.get('P_New', ''), _fmt_pct(r.get('P_Pct', '')),
                r.get('V2_Exist', ''), r.get('V2_New', ''), _fmt_pct(r.get('V2_Pct', '')),
                r.get('V3_Exist', ''), r.get('V3_New', ''), _fmt_pct(r.get('V3_Pct', '')),
                r.get('M2_Exist', ''), r.get('M2_New', ''), _fmt_pct(r.get('M2_Pct', '')),
                r.get('M3_Exist', ''), r.get('M3_New', ''), _fmt_pct(r.get('M3_Pct', '')),
                r.get('NetDemand', ''),
                r.get('FailReason', ''),
                _result_cell(r.get('Pass', '')),
            ])

        return vkt.TableResult(data, column_headers=DETAIL_HEADERS)

    # -- Plotly chart --------------------------------------------------------

    @vkt.PlotlyView('Results Chart', duration_guess=2)
    def results_chart(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')

        mode = params.step1.mode or 'Design Results'
        all_results = self._run_analysis_all(params) if mode == 'Analysis Results' else self._run_all(params)
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

        status_order = ['FLAG', 'WARN', 'PASS', 'ADDED', 'REMOVED']
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
            title={
                'text': 'Members by Story and Status (highest story at top)'
                        '<br><sup style="font-size:11px;color:#777">All members shown — independent of display filter</sup>',
                'x': 0.5, 'xanchor': 'center',
            },
            height=max(420, 28 * len(stories) + 120),
            legend=dict(orientation='h', yanchor='bottom', y=1.05,
                        xanchor='right', x=1),
            margin=dict(l=10, r=20, t=80, b=40),
        )

        return vkt.PlotlyResult(fig)

    # -- Summary table -------------------------------------------------------

    @vkt.TableView('By Story', duration_guess=2)
    def summary_table(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')

        mode = params.step1.mode or 'Design Results'
        all_results = self._run_analysis_all(params) if mode == 'Analysis Results' else self._run_all(params)
        summary = build_summary(all_results)
        if not summary:
            raise vkt.UserError('No data to summarise.')

        data = []
        for s in summary:
            fail_count = s['FLAG']
            warn_count = s['WARN']
            fail_cell = vkt.TableCell(
                str(fail_count),
                background_color=(
                    vkt.Color(178, 34, 34) if fail_count > 0
                    else vkt.Color(34, 139, 34)
                ),
                text_color=vkt.Color(255, 255, 255),
            )
            warn_cell = vkt.TableCell(
                str(warn_count),
                background_color=(
                    vkt.Color(180, 130, 0) if warn_count > 0
                    else vkt.Color(34, 139, 34)
                ),
                text_color=vkt.Color(255, 255, 255),
            )
            data.append([
                s['Story'], s['MemberType'], s['Total'],
                s['PASS'], warn_cell, fail_cell, s['ADDED'], s['REMOVED'],
            ])

        return vkt.TableResult(data, column_headers=SUMMARY_HEADERS)

    # -- Key metrics DataView ------------------------------------------------

    @vkt.DataView('Summary', duration_guess=2)
    def key_metrics(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')

        mode = params.step1.mode or 'Design Results'
        all_results = self._run_analysis_all(params) if mode == 'Analysis Results' else self._run_all(params)
        if not all_results:
            raise vkt.UserError('No data.')

        from collections import Counter
        total    = len(all_results)
        failures = [r for r in all_results if r.get('Pass') == 'FLAG']
        warnings = [r for r in all_results if r.get('Pass') == 'WARN']
        added    = [r for r in all_results if r.get('Pass') == 'ADDED']
        removed  = [r for r in all_results if r.get('Pass') == 'REMOVED']
        n_fail   = len(failures)
        n_warn   = len(warnings)
        fail_rate = round(n_fail / total * 100, 1) if total > 0 else 0.0

        story_fail_counts = Counter(r.get('Story') for r in failures)
        top_story, top_count = (
            story_fail_counts.most_common(1)[0] if story_fail_counts else ('None', 0)
        )

        fail_status = vkt.DataStatus.WARNING if n_fail > 0 else vkt.DataStatus.SUCCESS
        rate_status = (
            vkt.DataStatus.ERROR   if fail_rate > 10 else
            vkt.DataStatus.WARNING if fail_rate > 0  else
            vkt.DataStatus.SUCCESS
        )

        if mode == 'Analysis Results':
            p = params.step2.section_options
            fail_thresh = float(p.fail_threshold or 10)
            worst_force_pct  = None
            worst_force_label = ''
            for r in all_results:
                pct = r.get('WorstPct')
                if isinstance(pct, float) and (worst_force_pct is None or pct > worst_force_pct):
                    worst_force_pct  = pct
                    worst_force_label = f"{r.get('Story')} / {r.get('Label')}"
            worst_status = (
                vkt.DataStatus.ERROR   if (worst_force_pct or 0) > fail_thresh else
                vkt.DataStatus.WARNING if (worst_force_pct or 0) > 0           else
                vkt.DataStatus.INFO
            )
            data = vkt.DataGroup(
                vkt.DataItem('Total Members Compared', total),
                vkt.DataItem('Flagged', n_fail, status=fail_status,
                             status_message='Members exceeding flag threshold — review required' if n_fail > 0 else 'All members within threshold'),
                vkt.DataItem('Failure Rate', fail_rate, suffix='%', number_of_decimals=1, status=rate_status),
                vkt.DataItem('Warnings', n_warn,
                             status=vkt.DataStatus.WARNING if n_warn > 0 else vkt.DataStatus.SUCCESS,
                             status_message='Threshold exceeded but forces are small or demand decreased — review, likely acceptable' if n_warn > 0 else 'No warnings'),
                vkt.DataItem('Added Members', len(added)),
                vkt.DataItem('Removed Members', len(removed)),
                vkt.DataItem('Worst Force Change', worst_force_pct if worst_force_pct is not None else 0.0,
                             suffix='%', number_of_decimals=1, status=worst_status,
                             explanation_label=worst_force_label),
                vkt.DataItem('Most Affected Story', top_story, explanation_label=f'{top_count} failures'),
            )
        else:
            p = params.step2.section_options
            grav_thresh = float(p.gravity_threshold or 5)
            lat_thresh  = float(p.lateral_threshold or 10)
            pmm_flag_thresh = max(grav_thresh, lat_thresh)

            worst_pmm_pct = None
            worst_pmm_label = ''
            for r in all_results:
                pct = r.get('PMM_Pct')
                if isinstance(pct, float) and (worst_pmm_pct is None or pct > worst_pmm_pct):
                    worst_pmm_pct  = pct
                    worst_pmm_label = f"{r.get('Story')} / {r.get('Label')}"
            pmm_status = (
                vkt.DataStatus.ERROR   if (worst_pmm_pct or 0) > pmm_flag_thresh else
                vkt.DataStatus.WARNING if (worst_pmm_pct or 0) > 0               else
                vkt.DataStatus.INFO
            )

            grav_flags = [r for r in failures if r.get('LoadType') == 'gravity']
            lat_flags  = [r for r in failures if r.get('LoadType') == 'lateral']

            data = vkt.DataGroup(
                vkt.DataItem('Total Members Compared', total),
                vkt.DataItem('Flagged', n_fail, status=fail_status,
                             status_message='Members exceeding IBC 3403 thresholds — use engineering judgment' if n_fail > 0 else 'All members within threshold'),
                vkt.DataItem('  Gravity Flags', len(grav_flags),
                             status=vkt.DataStatus.WARNING if grav_flags else vkt.DataStatus.SUCCESS,
                             explanation_label=f'>{grav_thresh:.0f}% gravity combos'),
                vkt.DataItem('  Lateral Flags', len(lat_flags),
                             status=vkt.DataStatus.WARNING if lat_flags else vkt.DataStatus.SUCCESS,
                             explanation_label=f'>{lat_thresh:.0f}% lateral combos'),
                vkt.DataItem('Warnings', n_warn,
                             status=vkt.DataStatus.WARNING if n_warn > 0 else vkt.DataStatus.SUCCESS,
                             status_message='Threshold exceeded but demand reduced overall or member below capacity — review, likely acceptable' if n_warn > 0 else 'No warnings'),
                vkt.DataItem('Added Members', len(added)),
                vkt.DataItem('Removed Members', len(removed)),
                vkt.DataItem('Failure Rate', fail_rate, suffix='%', number_of_decimals=1, status=rate_status),
                vkt.DataItem('Worst PMM Change', worst_pmm_pct if worst_pmm_pct is not None else 0.0,
                             suffix='%', number_of_decimals=1, status=pmm_status,
                             explanation_label=worst_pmm_label),
                vkt.DataItem('Most Affected Story', top_story, explanation_label=f'{top_count} flags'),
            )
        return vkt.DataResult(data)

    # -- Beam detail HTML view -----------------------------------------------

    @vkt.WebView('Worst Beam (Design)', duration_guess=2)
    def beam_detail_view(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')

        if (params.step1.mode or 'Design Results') == 'Analysis Results':
            html = (
                '<p style="font-family:sans-serif;padding:24px;color:#555;font-size:15px;">'
                'Beam detail view is only available in Design mode.'
                '</p>'
            )
            return vkt.WebResult(html=html)

        p = params.step2.section_options
        gravity_thresh  = float(p.gravity_threshold or 5)
        lateral_thresh  = float(p.lateral_threshold or 10)

        results = self._run_all(params)
        beam    = self._find_worst_beam(results)
        html    = _render_beam_detail_html(beam, gravity_thresh, lateral_thresh)
        return vkt.WebResult(html=html)

    # -- Member detail HTML view -----------------------------------------------

    @vkt.WebView('Member Detail', duration_guess=2)
    def member_detail_view(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')

        mode           = params.step1.mode or 'Design Results'
        p              = params.step2.section_options
        d              = params.step2.section_detail
        gravity_thresh = float(p.gravity_threshold or 5)
        lateral_thresh = float(p.lateral_threshold or 10)

        search_label = (d.detail_label or '').strip()
        search_story = (d.detail_story or '').strip().lower()

        all_results = self._run_analysis_all(params) if mode == 'Analysis Results' else self._run_all(params)

        member = None
        if search_label:
            candidates = [
                r for r in all_results
                if str(r.get('Label', '')).strip().lower() == search_label.lower()
            ]
            if search_story:
                candidates = [r for r in candidates if str(r.get('Story', '')).strip().lower() == search_story] or candidates
            if candidates:
                # Prefer flagged members if multiple match; otherwise take first
                member = next((r for r in candidates if r.get('Pass') == 'FLAG'), candidates[0])

        html = _render_member_detail_html(member, gravity_thresh, lateral_thresh, mode)
        return vkt.WebResult(html=html)

    # -- Typed member-class tabs -----------------------------------------------

    @vkt.TableView('Braces', duration_guess=2)
    def braces_table(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')
        mode    = params.step1.mode or 'Design Results'
        results = self._run_typed(params, 'Braces')
        if not results:
            raise vkt.UserError(
                'No braces flagged for review. All may be passing, below the absolute force '
                'threshold, or demand decreased for all. Switch to "All Results" or lower '
                'the absolute threshold to see the full brace list.'
            )
        if mode == 'Analysis Results':
            headers = [
                'Story', 'Label', 'Combo (Exist)', 'Combo (New)',
                'P (Exist)', 'P (New)', 'P (%)',
                'Worst (%)', 'Flag Reason', 'Result',
            ]
            data = []
            for r in results:
                data.append([
                    r.get('Story', ''), r.get('Label', ''),
                    r.get('GovCombo_Exist', ''), r.get('GovCombo_New', ''),
                    r.get('P_Exist', ''), r.get('P_New', ''),
                    _fmt_pct(r.get('P_Pct', '')),
                    _fmt_pct(r.get('WorstPct', '')),
                    r.get('FailReason', ''), _result_cell(r.get('Pass', '')),
                ])
        else:
            headers = [
                'Story', 'Label', 'Section (Exist)', 'Section (New)',
                'Combo (Exist)', 'Combo (New)', 'Load Type',
                'P (Exist)', 'P (New)', 'P (%)',
                'PMM (Exist)', 'PMM (New)', 'PMM (%)',
                'Flag Reason', 'Result',
            ]
            data = []
            for r in results:
                data.append([
                    r.get('Story', ''), r.get('Label', ''),
                    r.get('DesignSection_Exist', ''), r.get('DesignSection_New', ''),
                    r.get('GovCombo_Exist', ''), r.get('GovCombo_New', ''),
                    r.get('LoadType', ''),
                    r.get('P_Exist', ''), r.get('P_New', ''),
                    _fmt_pct(r.get('P_Pct', '')),
                    r.get('PMM_Exist', ''), r.get('PMM_New', ''),
                    _fmt_pct(r.get('PMM_Pct', '')),
                    r.get('FailReason', ''), _result_cell(r.get('Pass', '')),
                ])
        return vkt.TableResult(data, column_headers=headers)

    @vkt.TableView('Columns', duration_guess=2)
    def columns_table(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')
        mode    = params.step1.mode or 'Design Results'
        results = self._run_typed(params, 'Columns')
        if not results:
            raise vkt.UserError(
                'No columns flagged for review. All may be passing, below the absolute force '
                'threshold, or demand decreased for all. Switch to "All Results" or lower '
                'the absolute threshold to see the full column list.'
            )
        if mode == 'Analysis Results':
            headers = [
                'Story', 'Label', 'Combo (Exist)', 'Combo (New)',
                'P (Exist)', 'P (New)', 'P (%)',
                'M3 (Exist)', 'M3 (New)', 'M3 (%)',
                'M2 (Exist)', 'M2 (New)', 'M2 (%)',
                'Worst (%)', 'Flag Reason', 'Result',
            ]
            data = []
            for r in results:
                data.append([
                    r.get('Story', ''), r.get('Label', ''),
                    r.get('GovCombo_Exist', ''), r.get('GovCombo_New', ''),
                    r.get('P_Exist', ''), r.get('P_New', ''),
                    _fmt_pct(r.get('P_Pct', '')),
                    r.get('M3_Exist', ''), r.get('M3_New', ''),
                    _fmt_pct(r.get('M3_Pct', '')),
                    r.get('M2_Exist', ''), r.get('M2_New', ''),
                    _fmt_pct(r.get('M2_Pct', '')),
                    _fmt_pct(r.get('WorstPct', '')),
                    r.get('FailReason', ''), _result_cell(r.get('Pass', '')),
                ])
        else:
            headers = [
                'Story', 'Label', 'Section (Exist)', 'Section (New)',
                'Combo (Exist)', 'Combo (New)', 'Load Type',
                'P (Exist)', 'P (New)', 'P (%)',
                'M3 (Exist)', 'M3 (New)', 'M3 (%)',
                'M2 (Exist)', 'M2 (New)', 'M2 (%)',
                'PMM (Exist)', 'PMM (New)', 'PMM (%)',
                'Flag Reason', 'Result',
            ]
            data = []
            for r in results:
                data.append([
                    r.get('Story', ''), r.get('Label', ''),
                    r.get('DesignSection_Exist', ''), r.get('DesignSection_New', ''),
                    r.get('GovCombo_Exist', ''), r.get('GovCombo_New', ''),
                    r.get('LoadType', ''),
                    r.get('P_Exist', ''), r.get('P_New', ''),
                    _fmt_pct(r.get('P_Pct', '')),
                    r.get('M3_Exist', ''), r.get('M3_New', ''),
                    _fmt_pct(r.get('M3_Pct', '')),
                    r.get('M2_Exist', ''), r.get('M2_New', ''),
                    _fmt_pct(r.get('M2_Pct', '')),
                    r.get('PMM_Exist', ''), r.get('PMM_New', ''),
                    _fmt_pct(r.get('PMM_Pct', '')),
                    r.get('FailReason', ''), _result_cell(r.get('Pass', '')),
                ])
        return vkt.TableResult(data, column_headers=headers)

    @vkt.TableView('Beams', duration_guess=2)
    def beams_table(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')
        mode    = params.step1.mode or 'Design Results'
        results = self._run_typed(params, 'Beams')
        if not results:
            raise vkt.UserError(
                'No beams flagged for review. All may be passing, below the absolute force '
                'threshold, or demand decreased for all. Switch to "All Results" or lower '
                'the absolute threshold to see the full beam list.'
            )
        if mode == 'Analysis Results':
            headers = [
                'Story', 'Label', 'Combo (Exist)', 'Combo (New)',
                'M3 (Exist)', 'M3 (New)', 'M3 (%)',
                'V2 (Exist)', 'V2 (New)', 'V2 (%)',
                'Worst (%)', 'Flag Reason', 'Result',
            ]
            data = []
            for r in results:
                data.append([
                    r.get('Story', ''), r.get('Label', ''),
                    r.get('GovCombo_Exist', ''), r.get('GovCombo_New', ''),
                    r.get('M3_Exist', ''), r.get('M3_New', ''),
                    _fmt_pct(r.get('M3_Pct', '')),
                    r.get('V2_Exist', ''), r.get('V2_New', ''),
                    _fmt_pct(r.get('V2_Pct', '')),
                    _fmt_pct(r.get('WorstPct', '')),
                    r.get('FailReason', ''), _result_cell(r.get('Pass', '')),
                ])
        else:
            headers = [
                'Story', 'Label', 'Section (Exist)', 'Section (New)',
                'Combo (Exist)', 'Combo (New)', 'Load Type',
                'M3 (Exist)', 'M3 (New)', 'M3 (%)',
                'V2 (Exist)', 'V2 (New)', 'V2 (%)',
                'PMM (Exist)', 'PMM (New)', 'PMM (%)',
                'Flag Reason', 'Result',
            ]
            data = []
            for r in results:
                data.append([
                    r.get('Story', ''), r.get('Label', ''),
                    r.get('DesignSection_Exist', ''), r.get('DesignSection_New', ''),
                    r.get('GovCombo_Exist', ''), r.get('GovCombo_New', ''),
                    r.get('LoadType', ''),
                    r.get('M3_Exist', ''), r.get('M3_New', ''),
                    _fmt_pct(r.get('M3_Pct', '')),
                    r.get('V2_Exist', ''), r.get('V2_New', ''),
                    _fmt_pct(r.get('V2_Pct', '')),
                    r.get('PMM_Exist', ''), r.get('PMM_New', ''),
                    _fmt_pct(r.get('PMM_Pct', '')),
                    r.get('FailReason', ''), _result_cell(r.get('Pass', '')),
                ])
        return vkt.TableResult(data, column_headers=headers)

    def _find_worst_beam(self, results: list) -> Optional[dict]:
        beams = [r for r in results if r.get('MemberType') == 'Beam'
                 and r.get('Pass') not in ('ADDED', 'REMOVED')]
        fails = [r for r in beams if r.get('Pass') == 'FLAG']
        pool  = fails if fails else beams
        if not pool:
            return None

        def score(r):
            force_max = max(
                (_numeric_pct(r.get(f'{c}_Pct', 0)) for c in FORCE_COMPONENTS),
                default=0.0,
            )
            return max(force_max, _numeric_pct(r.get('PMM_Pct', 0)))

        return max(pool, key=score)

    # -- CSV export ----------------------------------------------------------

    def download_csv(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')
        mode = params.step1.mode or 'Design Results'
        results = self._run_analysis_all(params) if mode == 'Analysis Results' else self._run_all(params)
        export_type = params.step2.section_export.member_type_export or 'All'
        if export_type != 'All':
            singular = _TYPE_SINGULAR[export_type]
            results = [r for r in results if r.get('MemberType') == singular]
        project = (params.step2.section_export.project_name or '').strip()
        project_slug = re.sub(r'[^\w-]', '_', project) if project else ''
        name_parts = ['etabs']
        if project_slug:
            name_parts.append(project_slug)
        if export_type != 'All':
            name_parts.append(export_type.lower())
        filename = '_'.join(name_parts) + '.csv'
        return vkt.DownloadResult(results_to_csv(results), filename)

    # -- LLM chat ------------------------------------------------------------

    def call_llm(self, params, **kwargs):
        import openai
        conversation = params.step2.section_ai.chat
        if not conversation:
            return None

        mode = params.step1.mode or 'Design Results'
        p    = params.step2.section_options

        def _severity(r):
            pct = r.get('WorstPct')
            if pct == 'INF':
                return 999.0
            return float(pct) if isinstance(pct, float) else -1.0

        if mode == 'Analysis Results':
            all_results = self._run_analysis_all(params)
            flags       = sorted(
                [r for r in all_results if r.get('Pass') == 'FLAG'],
                key=_severity, reverse=True,
            )
            warns       = [r for r in all_results if r.get('Pass') == 'WARN']
            added       = [r for r in all_results if r.get('Pass') == 'ADDED']
            removed     = [r for r in all_results if r.get('Pass') == 'REMOVED']
            top60       = (flags + warns)[:60]
            fail_threshold = float(p.fail_threshold or 10)
            warn_threshold = float(p.warn_threshold or 5)
            failure_lines = []
            for r in top60:
                status = r.get('Pass', '')
                failure_lines.append(
                    f"- [{status}] {r.get('MemberType')} {r.get('Label')} (Story {r.get('Story')}): "
                    f"{r.get('FailReason', 'N/A')} | worst: {r.get('WorstPct', 'N/A')}% | "
                    f"P: {r.get('P_Exist', 'N/A')} → {r.get('P_New', 'N/A')} | "
                    f"M3: {r.get('M3_Exist', 'N/A')} → {r.get('M3_New', 'N/A')} | "
                    f"combo (new): {r.get('GovCombo_New', 'N/A')}"
                )
            failure_block = "\n".join(failure_lines) if failure_lines else "None — all members within threshold."
            system_prompt = (
                f"You are a structural engineering assistant reviewing an ETABS element force comparison.\n\n"
                f"Mode: Analysis Results. Thresholds: WARN >{warn_threshold}%, FLAG >{fail_threshold}%.\n"
                f"Total members: {len(all_results)}. Flags: {len(flags)}. Warnings: {len(warns)}. "
                f"Added: {len(added)}. Removed: {len(removed)}.\n\n"
                f"WARN means the threshold was exceeded but forces are small or demand decreased overall — "
                f"review advised but likely acceptable without remediation.\n\n"
                f"Flagged / warned members (sorted by severity):\n{failure_block}\n\n"
                f"Each line shows: status, member type, label, story, fail reason, worst % change, "
                f"P and M3 existing→modified, governing combo in modified model.\n\n"
                f"Answer concisely in engineering terms. Reference specific members and stories. "
                f"If asked about a member not listed, note it is passing."
            )
        else:
            gravity_thresh = float(p.gravity_threshold or 5)
            lateral_thresh = float(p.lateral_threshold or 10)
            all_results    = self._run_all(params)
            flags          = sorted(
                [r for r in all_results if r.get('Pass') == 'FLAG'],
                key=_severity, reverse=True,
            )
            warns   = [r for r in all_results if r.get('Pass') == 'WARN']
            added   = [r for r in all_results if r.get('Pass') == 'ADDED']
            removed = [r for r in all_results if r.get('Pass') == 'REMOVED']
            top60   = (flags + warns)[:60]
            failure_lines = []
            for r in top60:
                sect_change = (
                    f"{r.get('DesignSection_Exist')} → {r.get('DesignSection_New')}"
                    if r.get('DesignSection_Exist') != r.get('DesignSection_New')
                    else r.get('DesignSection_Exist', 'N/A')
                )
                sign_rev = r.get('SignReversal', '')
                sign_note = f' | sign rev: {sign_rev}' if sign_rev else ''
                status = r.get('Pass', '')
                failure_lines.append(
                    f"- [{status}] {r.get('MemberType')} {r.get('Label')} (Story {r.get('Story')}): "
                    f"{r.get('FailReason', 'N/A')} | net demand: {r.get('NetDemand', 'N/A')} | "
                    f"load type: {r.get('LoadType', 'N/A')} | section: {sect_change} | "
                    f"PMM: {r.get('PMM_Exist', 'N/A')} → {r.get('PMM_New', 'N/A')}{sign_note}"
                )
            failure_block = "\n".join(failure_lines) if failure_lines else "None — all members are passing."
            system_prompt = (
                f"You are a structural engineering assistant reviewing an ETABS model comparison "
                f"for IBC Section 3403 compliance. The tool compares steel member demands between "
                f"an existing building and a proposed modified model.\n\n"
                f"Mode: Design Results. Thresholds: gravity {gravity_thresh}%, lateral {lateral_thresh}%.\n"
                f"Total members: {len(all_results)}. Flags: {len(flags)}. Warnings: {len(warns)}. "
                f"Added: {len(added)}. Removed: {len(removed)}.\n\n"
                f"WARN means the FLAG threshold was exceeded, but the modified model PMM D/C ratio is "
                f"still below 0.95 or all forces decreased — threshold tripped but member is not overstressed.\n\n"
                f"Flagged / warned members (sorted by severity):\n{failure_block}\n\n"
                f"Each line: status, member type, label, story, fail reason, net demand direction, "
                f"load type, section change, PMM ratio change, sign reversal if any.\n\n"
                f"Answer concisely in engineering terms. Reference specific members and stories. "
                f"If asked about a member not listed, note it is passing."
            )

        messages = [
            {"role": "system", "content": system_prompt},
            *conversation.get_messages(),
        ]

        try:
            stream = _llm_client.chat.completions.create(
                model="openai.gpt-oss-120b",
                messages=messages,
                stream=True,
            )
            text_stream = (
                chunk.choices[0].delta.content
                for chunk in stream
                if chunk.choices[0].delta.content is not None
            )
            return vkt.ChatResult(conversation, text_stream)
        except openai.RateLimitError:
            raise vkt.UserError('AI rate limit reached — wait a moment and try again.')

    # -- Shared helpers ------------------------------------------------------

    def _run(self, params):
        """Filtered results (respects display_filter param)."""
        p = params.step2.section_options
        results = run_comparison(
            existing_file=params.step1.existing_file.file,
            modified_file=params.step1.modified_file.file,
            member_type_filter='All',
            gravity_threshold=float(p.gravity_threshold or 5),
            lateral_threshold=float(p.lateral_threshold or 10),
            show_failures_only=(p.display_filter == 'Failures Only'),
        )
        return _postprocess_results(results)

    def _run_all(self, params):
        """Unfiltered results — used by chart, summary, key metrics, CSV, and typed tabs."""
        p = params.step2.section_options
        results = run_comparison(
            existing_file=params.step1.existing_file.file,
            modified_file=params.step1.modified_file.file,
            member_type_filter='All',
            gravity_threshold=float(p.gravity_threshold or 5),
            lateral_threshold=float(p.lateral_threshold or 10),
            show_failures_only=False,
        )
        return _postprocess_results(results)

    def _run_analysis(self, params):
        """Filtered analysis results."""
        p = params.step2.section_options
        return run_analysis_comparison(
            existing_file=params.step1.existing_file.file,
            modified_file=params.step1.modified_file.file,
            member_type_filter='All',
            warn_threshold=float(p.warn_threshold or 5),
            fail_threshold=float(p.fail_threshold or 10),
            show_failures_only=(p.display_filter == 'Failures Only'),
        )

    def _run_analysis_all(self, params):
        """Unfiltered analysis results — used by chart, summary, key metrics, CSV, and typed tabs."""
        p = params.step2.section_options
        return run_analysis_comparison(
            existing_file=params.step1.existing_file.file,
            modified_file=params.step1.modified_file.file,
            member_type_filter='All',
            warn_threshold=float(p.warn_threshold or 5),
            fail_threshold=float(p.fail_threshold or 10),
            show_failures_only=False,
        )

    def _run_typed(self, params, member_type: str) -> list:
        """Filtered + focused results for one member type (typed review tabs).

        Applies: member type filter, absolute force threshold, demand direction
        filter, and display_filter (Failures Only / All Results).
        """
        p = params.step2.section_options
        mode = params.step1.mode or 'Design Results'
        is_analysis = (mode == 'Analysis Results')
        force_thresholds = {
            'P':  float(p.thresh_P  if p.thresh_P  is not None else 5.0),
            'M3': float(p.thresh_M3 if p.thresh_M3 is not None else 20.0),
            'M2': float(p.thresh_M2 if p.thresh_M2 is not None else 10.0),
            'V2': float(p.thresh_V2 if p.thresh_V2 is not None else 5.0),
            'V3': float(p.thresh_V3 if p.thresh_V3 is not None else 0.0),
        }
        grav_thresh  = float(p.gravity_threshold or 5)
        lat_thresh   = float(p.lateral_threshold or 10)
        warn_thresh  = float(p.warn_threshold or 5)
        fail_thresh  = float(p.fail_threshold or 10)
        hide_decreases     = p.hide_decreases if p.hide_decreases is not None else True
        show_added_removed = p.show_added_removed if p.show_added_removed is not None else True
        display_filter     = p.display_filter or 'Failures Only'

        all_results = (
            self._run_analysis_all(params)
            if mode == 'Analysis Results'
            else self._run_all(params)
        )

        singular   = _TYPE_SINGULAR[member_type]
        gov_forces = _GOVERNING_FORCES_BY_TYPE[member_type]

        out = []
        for r in all_results:
            if r.get('MemberType') != singular:
                continue

            status = r.get('Pass', '')

            if status in ('ADDED', 'REMOVED') and not show_added_removed:
                continue

            # Per-force absolute thresholds.
            # For FLAG/WARN: at least one *flagging* governing force (exceeds the % threshold
            #   or is INF) must also be above the abs limit. Forces that trip the % threshold
            #   but are below the abs minimum are excluded — the member is filtered out
            #   entirely rather than shown with a misleading or empty reason.
            # For PASS: at least one governing force (either model) must be above its limit.
            if status not in ('ADDED', 'REMOVED'):
                if status in ('FLAG', 'WARN'):
                    if is_analysis:
                        member_pct_thresh = fail_thresh if status == 'FLAG' else warn_thresh
                    else:
                        _lt = r.get('LoadType', 'gravity')
                        member_pct_thresh = lat_thresh if _lt == 'lateral' else grav_thresh
                    any_meaningful = False
                    for f in gov_forces:
                        pct = r.get(f'{f}_Pct')
                        # Must be flagging (INF or exceeds % threshold), not just increasing
                        is_flagging = pct == 'INF' or (isinstance(pct, float) and pct > member_pct_thresh)
                        if not is_flagging:
                            continue
                        limit = force_thresholds.get(f, 0.0)
                        max_f = max(
                            (abs(r.get(k)) for k in (f'{f}_Exist', f'{f}_New')
                             if isinstance(r.get(k), (int, float))),
                            default=0.0,
                        )
                        if max_f >= limit:
                            any_meaningful = True
                            break
                    if not any_meaningful:
                        continue
                else:  # PASS
                    any_above = any(
                        max((abs(r.get(k)) for k in (f'{f}_Exist', f'{f}_New')
                             if isinstance(r.get(k), (int, float))), default=0.0)
                        >= force_thresholds.get(f, 0.0)
                        for f in gov_forces
                    )
                    if not any_above:
                        continue

            # Demand direction filter
            if hide_decreases and status not in ('ADDED', 'REMOVED'):
                nd = r.get('NetDemand') or _net_demand(r)
                if nd == 'DOWN':
                    continue

            # Display filter
            if display_filter == 'Failures Only' and status == 'PASS':
                continue

            # Recompute FailReason for FLAG/WARN to only reference gov_forces above abs threshold
            if status in ('FLAG', 'WARN'):
                r = _recompute_typed_fail_reason(
                    r, gov_forces, force_thresholds,
                    grav_thresh, lat_thresh, warn_thresh, fail_thresh,
                    is_analysis, status,
                )

            out.append(r)

        return out
