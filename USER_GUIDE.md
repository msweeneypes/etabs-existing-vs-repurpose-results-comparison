# ETABS Model Comparison Tool — User Guide

Compares structural demand between an existing and modified ETABS model to
support IBC Section 3403 compliance review for repurposed or modified
buildings.

---

## 1. Prepare Your ETABS Exports

Both files must be exported from ETABS as a single `.xlsx` workbook. Match the
export to the result type you intend to use.

### Design Results mode

Run Steel Frame Design, then export to one workbook:

| ETABS Table | Sheet Name in Export |
|---|---|
| Design Forces - Columns | Design Forces - Columns |
| Design Forces - Beams | Design Forces - Beams |
| Design Forces - Braces | Design Forces - Braces |
| Steel Frame Design Summary - AISC 360-16 | Steel Frame Design Summary |

Use **File → Export → Tables to Excel** in ETABS. All four sheets must be in
the same workbook. Repeat for both models.

### Analysis Results mode

Run analysis (no design run required), then export to one workbook:

| ETABS Table | Sheet Name in Export |
|---|---|
| Element Forces - Columns | Element Forces - Columns |
| Element Forces - Beams | Element Forces - Beams |
| Element Forces - Braces | Element Forces - Braces |

---

## 2. Step 1 — Upload Files

1. Select **Result Type**: **Design Results** (post-design D/C ratios and force
   envelopes) or **Analysis Results** (element force envelopes only).
2. Upload the existing model workbook to **Existing Model**.
3. Upload the modified model workbook to **Modified Model**.
4. Click **Next**.

The first run parses both files and caches the result — typically 1–2 minutes
for large models. All subsequent tab switches and threshold adjustments are
near-instant.

---

## 3. Step 2 — Configure & Compare

### Comparison Options

**Design mode thresholds (IBC 3403)**

| Field | Default | Meaning |
|---|---|---|
| Gravity Load Threshold | 5% | Max allowable increase for gravity combos |
| Lateral Load Threshold | 10% | Max allowable increase for lateral combos |

Load type (gravity vs. lateral) is determined from the governing combo name.
Combos containing `WA`, `WB`, `WG` (wind) or `EQ`, `EQB` (seismic) tokens are
classified as lateral; all others are gravity.

**Analysis mode thresholds**

| Field | Default | Meaning |
|---|---|---|
| Warn Threshold | 5% | Force increase above this triggers WARN |
| Flag Threshold | 10% | Force increase above this triggers FLAG |

**Absolute force minimums — Typed Tabs**

These filter out low-magnitude noise in the Braces, Columns, and Beams tabs.
A member is hidden from the typed tab if its force stays below both the existing
and modified values for the relevant component.

| Field | Default | Applies To |
|---|---|---|
| Min \|P\| (kips) | 5 | Braces, Columns |
| Min \|M3\| (kip-ft) | 20 | Beams, Columns |
| Min \|M2\| (kip-ft) | 10 | Columns |
| Min \|V2\| (kips) | 5 | Beams |
| Min \|V3\| (kips) | 0 (disabled) | Columns |

Raise these values to reduce noise; lower them to see light members.

**Display options**

- **Hide members where demand decreased** — hides members from the typed tabs
  where all governing forces decreased. Enabled by default; the Overview tab
  always shows the full picture.
- **Show added/removed members** — includes members that exist only in one
  model.
- **Display Filter** — controls the Overview and Full Detail tabs: *All
  Results* or *Failures Only* (FLAG and WARN only).

---

## 4. Result Tabs

### Overview

One row per member across all types. Controlled by the **Display Filter**
option — defaults to *Failures Only* (FLAG and WARN); switch to *All Results*
to see every member.

**Design mode columns:**

| Column | Description |
|---|---|
| Story / Label / Type | Member identification |
| Load Type | gravity or lateral |
| Section (Exist/New) | Design section from ETABS |
| PMM (Exist/New/%) | AISC 360-16 H1-1 interaction ratio and percent change |
| V Major D/C | Percent change in major-axis shear D/C ratio |
| Worst Force % | Largest percent increase among P, V2, V3, T, M2, M3 |
| Flag Reason | Plain-language description of what triggered the flag |
| Result | PASS / FLAG / WARN / ADDED / REMOVED |

**Analysis mode columns:**

| Column | Description |
|---|---|
| Story / Label / Type | Member identification |
| P/V2/M2/M3 (New) | Modified model force value |
| P/V2/M2/M3 (%) | Percent change from existing |
| Worst % | Largest percent increase among governing forces |
| Flag Reason | What triggered the flag |
| Result | PASS / FLAG / WARN / ADDED / REMOVED |

### Braces / Columns / Beams

Focused view for each member type. Shows only governing forces for that class,
filtered by absolute thresholds and demand direction. Use these for structured
review rather than the full table.

Governing forces per member type:

| Type | Governing Forces Checked |
|---|---|
| Columns | P, M3, M2 |
| Beams | M3, V2 |
| Braces | P |

The Flag Reason in these tabs is restricted to the governing forces above their
absolute threshold, so a negligible M2 value will not appear as the reason on a
beam or column.

### Member Detail

HTML card view for a single member. Auto-selects the worst flagged member when
no label is entered. To look up a specific member:

1. Expand **Member Detail** in the sidebar.
2. Enter the exact ETABS label (e.g. `C12`, `B5`). Case-insensitive.
3. Optionally enter a story if the label appears on multiple stories.

The card shows: banner status, section and combo information, a force-by-force
table with threshold highlighting, and D/C ratios (Design mode).

### Key Metrics

High-level summary: total members compared, FLAG/WARN/PASS counts, worst member
by type, and top-5 flagged members.

### By Story

Member counts broken down by story and type. Use to identify floors with
concentrated demand increases.

### Full Detail

All members, all force components (P, V2, V3, T, M2, M3), unfiltered. Useful
for exporting context or investigating specific members. Controlled by the
**Display Filter** option.

---

## 5. Status Codes

| Status | Meaning |
|---|---|
| **PASS** | All governing forces within threshold |
| **FLAG** | Threshold exceeded; review required |
| **WARN** | Threshold exceeded, but demand decreased overall or member is well below capacity (PMM < 0.95) |
| **ADDED** | Member exists only in the modified model |
| **REMOVED** | Member exists only in the existing model |

FLAG does not imply a design failure. It flags members where the percent
increase in demand crosses the IBC 3403 threshold and requires engineering
judgment.

**INF** in a percent change column means the existing model force was
effectively zero (< 0.01) while the new model is non-zero. Evaluate the
absolute magnitude — INF alone may not require action.

**Sign Reversal** is flagged separately when a force changed direction (e.g.,
tension to compression). The percent change is computed on absolute values, so
the sign reversal will not inflate the percentage but is noted for review.

---

## 6. Export

In **Export** (sidebar), optionally enter a project name and select a member
type, then click **Export Results to CSV**. The CSV contains all columns from
the Full Detail tab and is suitable for attaching to a calculation package.

Filename format: `etabs_[ProjectName]_[MemberType]_YYYY-MM-DD.csv`

---

## 7. AI Assistant

The **AI Assistant** chat has access to the full comparison results. You can
ask:

- Which members failed and why?
- What is the most critical issue?
- Summarize the flagged columns by story.
- Are any members showing sign reversals?

The assistant is context-aware for the current run; it does not retain
information between sessions.

---

## 8. Member Matching

Members are matched by **Story + ETABS Label**. If the label changes between
models (e.g., a remap after story insertion), the member will appear as REMOVED
in one and ADDED in the other rather than matched. Verify that label assignments
are consistent between the two ETABS models before running the comparison.

---

## 9. Troubleshooting

**"No members found"** — The expected sheet names were not found in the
workbook. Confirm you exported the correct tables and that the sheet names match
exactly. ETABS sometimes appends version strings to table names.

**Typed tab shows no results** — All members may be passing, below the
absolute force minimums, or showing decreased demand with "Hide decreases"
enabled. Lower the absolute thresholds or switch the Display Filter to *All
Results* to verify.

**Label not found in Member Detail** — Labels are matched exactly as they
appear in the ETABS export. Check for leading/trailing spaces or abbreviation
differences between models.

**Slow first load** — Parsing large workbooks takes 1–2 minutes. Results are
cached after the first run; subsequent threshold changes and tab switches are
fast.
