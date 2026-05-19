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
    'Section', 'Gov Combo', 'Load Type',
    'P (Exist)', 'P (New)', 'P (%)',
    'V2 (Exist)', 'V2 (New)', 'V2 (%)',
    'V3 (Exist)', 'V3 (New)', 'V3 (%)',
    'M2 (Exist)', 'M2 (New)', 'M2 (%)',
    'M3 (Exist)', 'M3 (New)', 'M3 (%)',
    'Net Demand', 'Fail Reason', 'Result',
]

SUMMARY_HEADERS = ['Story', 'Type', 'Total', 'PASS', 'WARN', 'FAIL', 'ADDED', 'REMOVED']

ANALYSIS_OVERVIEW_HEADERS = [
    'Story', 'Label', 'Type',
    'P (%)', 'V2 (%)', 'M2 (%)', 'M3 (%)',
    'Worst (%)', 'Fail Reason', 'Result',
]

ANALYSIS_DETAIL_HEADERS = [
    'Story', 'Label', 'Type',
    'P (Exist)', 'P (New)', 'P (%)',
    'V2 (Exist)', 'V2 (New)', 'V2 (%)',
    'M2 (Exist)', 'M2 (New)', 'M2 (%)',
    'M3 (Exist)', 'M3 (New)', 'M3 (%)',
    'Worst (%)', 'Fail Reason', 'Result',
]

# ---------------------------------------------------------------------------
# Colors
# ---------------------------------------------------------------------------

_RESULT_COLORS = {
    'PASS':    vkt.Color(34, 139, 34),
    'FAIL':    vkt.Color(178, 34, 34),
    'WARN':    vkt.Color(180, 130, 0),
    'ADDED':   vkt.Color(30, 100, 200),
    'REMOVED': vkt.Color(180, 100, 0),
}

_CHART_COLORS = {
    'FAIL':    '#B22222',
    'WARN':    '#B48200',
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


_CAPACITY_WARN_THRESHOLD = 0.95  # PMM below this + FAIL → WARN (member still under capacity)


def _postprocess_results(results: list) -> list:
    """Add NetDemand; downgrade FAIL → WARN for two conditions:
    1. All numeric forces decreased (INF-only failure on previously-zero component).
    2. Threshold exceeded but modified PMM D/C ratio is still < 0.95 (below capacity).
    """
    out = []
    for r in results:
        r = dict(r)
        nd = _net_demand(r)
        r['NetDemand'] = nd
        if r.get('Pass') == 'FAIL':
            pmm_new = r.get('PMM_New')
            forces_down = nd == 'DOWN'
            below_capacity = isinstance(pmm_new, float) and pmm_new < _CAPACITY_WARN_THRESHOLD
            if forces_down or below_capacity:
                r['Pass'] = 'WARN'
        out.append(r)
    return out


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
        'results_table',
        'results_detail',
        'results_chart',
        'summary_table',
        'key_metrics',
        'beam_detail_view',
    ])

    step2.section_about = vkt.Section('How This Works')

    step2.section_about.intro_method = vkt.Text("""
## How This Comparison Works

This tool compares structural demand on every steel member between the **existing** and **modified** models to check compliance with **IBC Section 3403**, which governs force increases in members of existing buildings being modified or repurposed.

For each matched member the tool finds the **worst-case force across all ETABS design load combinations and stations**, then computes the percent change in magnitude as (abs(new) - abs(exist)) / abs(exist) x 100. Using absolute values ensures a sign reversal (tension flipping to compression) does not artificially inflate the percentage. Sign reversals are flagged separately in the Sign Rev. column.
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

**Fail Reason** - Plain-English description of what caused the FAIL flag, e.g. "M3 +23.5% > 5% gravity".
""")

    step2.section_about.intro_thresholds = vkt.Text("""
### Pass / Fail and Load Classification (IBC 3403)

A member is flagged **FAIL** if any force component or D/C ratio increases by more than the threshold for its governing load type:

- **Gravity combos** (dead, live, snow, roof live) - default 5% increase triggers a flag
- **Lateral combos** (wind, seismic) - default 10% increase triggers a flag

Load type is determined from the governing combo name in the **modified** model. Combos with wind tokens (WA, WB, WG) or seismic tokens (EQ, EQB) in the name are classified as lateral; all others are gravity. Thresholds can be adjusted in the Comparison Options below.

**ADDED** means the member exists in the modified model but not the existing model. No threshold check is applied.

**REMOVED** means the member exists in the existing model but not the modified model.

Start with the **Overview** tab to spot failures quickly. Use **Full Detail** to see all six force components. The **Results Chart** shows the breakdown by story at a glance.
""")

    step2.section_options = vkt.Section('Comparison Options')
    step2.section_options.member_type = vkt.OptionField(
        'Member Type',
        options=['All', 'Columns', 'Beams', 'Braces'],
        default='All',
    )
    step2.section_options.gravity_threshold = vkt.NumberField(
        'Gravity Load Threshold — Design mode (%)',
        default=5,
        min=0,
        description='Maximum allowable % increase for gravity combos (IBC 3403)',
    )
    step2.section_options.lateral_threshold = vkt.NumberField(
        'Lateral Load Threshold — Design mode (%)',
        default=10,
        min=0,
        description='Maximum allowable % increase for lateral combos (IBC 3403)',
    )
    step2.section_options.warn_threshold = vkt.NumberField(
        'Warn Threshold — Analysis mode (%)',
        default=5,
        min=0,
        description='Force increase above this triggers WARN',
    )
    step2.section_options.fail_threshold = vkt.NumberField(
        'Fail Threshold — Analysis mode (%)',
        default=10,
        min=0,
        description='Force increase above this triggers FAIL',
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

    step2.section_ai = vkt.Section('AI Assistant')
    step2.section_ai.chat = vkt.Chat(
        'Ask about failures',
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

    banner_bg = '#B22222' if pass_val == 'FAIL' else '#228B22'

    if pass_val == 'FAIL':
        result_phrase = '<span style="font-weight:bold">FAILED</span>'
        narrative = (
            f'Beam <strong>{label}</strong> on Story <strong>{story}</strong> {result_phrase} — '
            f'{fail_rsn}.'
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
        status  = 'FAIL' if is_fail else 'PASS'
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
        status  = 'FAIL' if is_fail else 'PASS'
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
# Controller
# ---------------------------------------------------------------------------

class Controller(vkt.Controller):
    parametrization = Parametrization

    # -- Overview table (primary view) ---------------------------------------

    @vkt.TableView('Overview', duration_guess=30)
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
                    _fmt_pct(r.get('P_Pct', '')),
                    _fmt_pct(r.get('V2_Pct', '')),
                    _fmt_pct(r.get('M2_Pct', '')),
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

    @vkt.PlotlyView('Results Chart', duration_guess=30)
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

        mode = params.step1.mode or 'Design Results'
        all_results = self._run_analysis_all(params) if mode == 'Analysis Results' else self._run_all(params)
        summary = build_summary(all_results)
        if not summary:
            raise vkt.UserError('No data to summarise.')

        data = []
        for s in summary:
            fail_count = s['FAIL']
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

    @vkt.DataView('Key Metrics', duration_guess=30)
    def key_metrics(self, params, **kwargs):
        if not params.step1.existing_file or not params.step1.modified_file:
            raise vkt.UserError('Please upload both model files in Step 1.')

        mode = params.step1.mode or 'Design Results'
        all_results = self._run_analysis_all(params) if mode == 'Analysis Results' else self._run_all(params)
        if not all_results:
            raise vkt.UserError('No data.')

        from collections import Counter
        total    = len(all_results)
        failures = [r for r in all_results if r.get('Pass') == 'FAIL']
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

        fail_status = vkt.DataStatus.ERROR if n_fail > 0 else vkt.DataStatus.SUCCESS
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
                vkt.DataItem('Failures', n_fail, status=fail_status,
                             status_message='Members exceeding fail threshold' if n_fail > 0 else 'All members within threshold'),
                vkt.DataItem('Failure Rate', fail_rate, suffix='%', number_of_decimals=1, status=rate_status),
                vkt.DataItem('Warnings', n_warn,
                             status=vkt.DataStatus.WARNING if n_warn > 0 else vkt.DataStatus.SUCCESS),
                vkt.DataItem('Added Members', len(added)),
                vkt.DataItem('Removed Members', len(removed)),
                vkt.DataItem('Worst Force Change', worst_force_pct if worst_force_pct is not None else 0.0,
                             suffix='%', number_of_decimals=1, status=worst_status,
                             explanation_label=worst_force_label),
                vkt.DataItem('Most Affected Story', top_story, explanation_label=f'{top_count} failures'),
            )
        else:
            worst_pmm_pct = None
            worst_pmm_label = ''
            for r in all_results:
                pct = r.get('PMM_Pct')
                if isinstance(pct, float) and (worst_pmm_pct is None or pct > worst_pmm_pct):
                    worst_pmm_pct  = pct
                    worst_pmm_label = f"{r.get('Story')} / {r.get('Label')}"
            pmm_status = (
                vkt.DataStatus.ERROR   if (worst_pmm_pct or 0) > 10 else
                vkt.DataStatus.WARNING if (worst_pmm_pct or 0) > 0  else
                vkt.DataStatus.INFO
            )
            data = vkt.DataGroup(
                vkt.DataItem('Total Members Compared', total),
                vkt.DataItem('Failures', n_fail, status=fail_status,
                             status_message='Members exceeding IBC 3403 thresholds' if n_fail > 0 else 'All members within threshold'),
                vkt.DataItem('Failure Rate', fail_rate, suffix='%', number_of_decimals=1, status=rate_status),
                vkt.DataItem('Warnings', n_warn,
                             status=vkt.DataStatus.WARNING if n_warn > 0 else vkt.DataStatus.SUCCESS,
                             status_message='Failed only on INF while primary forces decreased — review but likely acceptable' if n_warn > 0 else 'No warnings'),
                vkt.DataItem('Added Members', len(added)),
                vkt.DataItem('Removed Members', len(removed)),
                vkt.DataItem('Worst PMM Change', worst_pmm_pct if worst_pmm_pct is not None else 0.0,
                             suffix='%', number_of_decimals=1, status=pmm_status,
                             explanation_label=worst_pmm_label),
                vkt.DataItem('Most Affected Story', top_story, explanation_label=f'{top_count} failures'),
            )
        return vkt.DataResult(data)

    # -- Beam detail HTML view -----------------------------------------------

    @vkt.WebView('Beam Detail', duration_guess=30)
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

    def _find_worst_beam(self, results: list) -> Optional[dict]:
        beams = [r for r in results if r.get('MemberType') == 'Beam'
                 and r.get('Pass') not in ('ADDED', 'REMOVED')]
        fails = [r for r in beams if r.get('Pass') == 'FAIL']
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
        return vkt.DownloadResult(results_to_csv(results), 'etabs_comparison_results.csv')

    # -- LLM chat ------------------------------------------------------------

    def call_llm(self, params, **kwargs):
        import openai
        conversation = params.step2.section_ai.chat
        if not conversation:
            return None

        p = params.step2.section_options
        gravity_thresh  = float(p.gravity_threshold or 5)
        lateral_thresh  = float(p.lateral_threshold or 10)
        all_results     = self._run_all(params)
        failures        = [r for r in all_results if r.get('Pass') in ('FAIL', 'WARN')]

        failure_lines = []
        for r in failures[:60]:
            sect_change = (
                f"{r.get('DesignSection_Exist')} → {r.get('DesignSection_New')}"
                if r.get('DesignSection_Exist') != r.get('DesignSection_New')
                else r.get('DesignSection_Exist', 'N/A')
            )
            status = r.get('Pass', 'FAIL')
            failure_lines.append(
                f"- [{status}] {r['MemberType']} {r['Label']} (Story {r['Story']}): "
                f"{r.get('FailReason', 'N/A')} | net demand: {r.get('NetDemand', 'N/A')} | "
                f"load type: {r.get('LoadType', 'N/A')} | section: {sect_change} | "
                f"PMM: {r.get('PMM_Exist', 'N/A')} → {r.get('PMM_New', 'N/A')}"
            )

        failure_block = "\n".join(failure_lines) if failure_lines else "None — all members are passing."

        system_prompt = (
            f"You are a structural engineering assistant reviewing an ETABS model comparison "
            f"for IBC Section 3403 compliance. The tool compares steel member demands between "
            f"an existing building and a proposed modified model.\n\n"
            f"Thresholds in use: gravity combos {gravity_thresh}%, lateral combos {lateral_thresh}%.\n"
            f"Total members compared: {len(all_results)}. Failures: {len(failures)}.\n\n"
            f"Failing members:\n{failure_block}\n\n"
            f"Each failure line shows: member type, label, story, fail reason (which force/ratio "
            f"exceeded the threshold and by how much), load type, section change, and PMM ratio change.\n\n"
            f"Answer concisely in engineering terms. Reference specific members and stories. "
            f"If asked about a member not in the failure list, note that it is passing."
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
            raise vkt.UserError("LLM rate limit reached — please wait a moment and try again.")

    # -- Shared helpers ------------------------------------------------------

    def _run(self, params):
        """Filtered results (respects display_filter param)."""
        p = params.step2.section_options
        results = run_comparison(
            existing_file=params.step1.existing_file.file,
            modified_file=params.step1.modified_file.file,
            member_type_filter=p.member_type or 'All',
            gravity_threshold=float(p.gravity_threshold or 5),
            lateral_threshold=float(p.lateral_threshold or 10),
            show_failures_only=(p.display_filter == 'Failures Only'),
        )
        return _postprocess_results(results)

    def _run_all(self, params):
        """Unfiltered results — used by chart, summary, key metrics, and CSV."""
        p = params.step2.section_options
        results = run_comparison(
            existing_file=params.step1.existing_file.file,
            modified_file=params.step1.modified_file.file,
            member_type_filter=p.member_type or 'All',
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
            member_type_filter=p.member_type or 'All',
            warn_threshold=float(p.warn_threshold or 5),
            fail_threshold=float(p.fail_threshold or 10),
            show_failures_only=(p.display_filter == 'Failures Only'),
        )

    def _run_analysis_all(self, params):
        """Unfiltered analysis results — used by chart, summary, key metrics, and CSV."""
        p = params.step2.section_options
        return run_analysis_comparison(
            existing_file=params.step1.existing_file.file,
            modified_file=params.step1.modified_file.file,
            member_type_filter=p.member_type or 'All',
            warn_threshold=float(p.warn_threshold or 5),
            fail_threshold=float(p.fail_threshold or 10),
            show_failures_only=False,
        )
