# ITSM Report Automation Agent — CLAUDE.md

# Version: 3.0 | Based on real Excel template analysis + POC review

# Last updated: April 2026

---

## FIRST THINGS FIRST — READ BEFORE TOUCHING ANYTHING

### Step 1: Understand the existing code

The POC is ALREADY WORKING and generating PPTX files.
Do NOT rewrite working code. Do NOT change structure unless broken.

Read every file in this order before making ANY changes:

1. config.yaml — understand current config structure
2. main.py — understand entry points
3. agent/orchestrator.py — understand pipeline flow
4. tools/t2_excel_reader.py — understand how Excel is currently read
5. tools/t3_kpi_parser.py — understand KPI evaluation
6. tools/t4_ai_engine.py — understand AI prompt structure
7. tools/t5_chart_generator.py — understand current chart logic
8. tools/t6_pptx_builder.py — understand current slide layout
9. workflows/w1_ingest.py through w5_slide_builder.py
10. tests/ — understand what is already tested

### Step 2: Report findings before doing anything

After reading all files, tell the user:

- What is working correctly
- What is missing or incomplete
- What files (if any) are unused or redundant
- Ask for confirmation before removing ANY file

### Step 3: Ask questions for anything unclear

If ANY requirement below contradicts the existing code,
or if you are unsure how to extend without breaking,
ASK THE USER before changing anything.

### Step 4: Do NOT clean up **pycache** or .pytest_cache

These are auto-generated. Leave them alone.

---

## PROJECT OVERVIEW

MUFG ITSM teams (MIM, Problem, Change, CMDB, Monitoring, NOC,
ITOR, APAC SD, EMEA SD, ServiceNow) record weekly performance
data in a shared Excel workbook.

The agent automatically:

1. Detects Excel file update (watchdog or --run-once)
2. Reads all team sheets from the Excel file
3. Reads user-provided summary from "Weekly Comments/Achievements" column
4. Sends data + context to AI (Ollama) for analysis
5. Generates charts from raw data (NOT from AI output)
6. Builds one professional slide per team in a PowerPoint file
7. Saves PPTX to ./output/ folder

---

## CRITICAL FACTS ABOUT THE REAL EXCEL FILE

These are facts discovered from the actual Scorecard_Template.xlsx.
The code MUST handle all of these correctly.

### Fact 1: Sheet names in the real workbook

Actual sheets present:
Sheet1 ← skip (empty/unused)
Cover Page ← skip (KPI reference, description only)
ITSM_Score_Card ← special handling (cross-team scorecard)
MIM (2) ← SKIP — duplicate/variant
MIM1 ← SKIP — duplicate/variant
MIM ← PROCESS this one
Problem ← PROCESS
Change ← PROCESS
CMDB ← PROCESS
Monitoring ← PROCESS
NOC ← PROCESS
ITOR ← PROCESS
APAC SD ← PROCESS
APAC SD1 ← SKIP — duplicate/variant
EMEA SD1 ← SKIP — duplicate/variant
EMEA SD ← PROCESS
ServiceNow ← PROCESS

Duplicate detection rule:
If sheet name ends with a number (1, 2, 3...) AND a sheet
with the same base name exists → SKIP the numbered variant.
Example: "MIM1" skipped because "MIM" exists.
Example: "MIM (2)" skipped because "MIM" exists.
Exception: If only numbered variant exists and no base → process it.

### Fact 2: Each sheet has TWO halves — only read LEFT half

Every team sheet has:
LEFT HALF (columns A onward): Weekly data — THIS IS WHAT TO READ
RIGHT HALF (starts around column M-R): Quarterly/monthly aggregates
with SUMIF/AVERAGEIF formulas

How to detect the split:
Right half starts where column headers REPEAT after a gap of
blank columns OR where "Month" / "Mothly No." header appears again.

Implementation:
After reading row 1 (headers), find the FIRST blank column
after the first few data columns. Everything before that blank
column = left half (weekly data). Read ONLY those columns.

Alternative (more reliable):
Read all headers. Find index of first repeated header.
Use only columns up to that index.

### Fact 3: KPI values are decimals (0.0 to 1.0), not percentages

Examples from real data:
0.6666666666666666 means 66.67%
1 means 100%
0.0061 means 0.61%
0.84 means 84%

Rule: IF column is classified as "percent" type AND values are <= 1.0
THEN multiply by 100 for display.
IF values are already > 1.0 (e.g. 66.67) THEN use as-is.

### Fact 4: "Weekly Comments/Achievements" = user-provided summary

This column already exists in EVERY team sheet.
Column name pattern: "Weekly Comments/\nAchievements" or
"Weekly Comments/Achievements"
This is where team leads write their weekly narrative.
It contains detailed text, bullet points, achievements, reasons
for KPI results.

This column is the USER SUMMARY. The AI should reference it.
Do NOT add a new "Summary" column. Use this existing column.

### Fact 5: Some cells contain Excel formulas, not values

openpyxl with data_only=True returns computed values, not formulas.
BUT: if the file was not opened/saved in Excel after formula entry,
openpyxl may return None for formula cells.

Handling:
If data_only=True returns None for a cell that should have a value:

- Try reading with data_only=False and evaluate manually if simple
- OR mark cell as "formula_unresolved" and skip it
- Log a warning but do NOT crash

### Fact 6: Percentages in column headers indicate KPI thresholds

Examples from real data:
"Change Success Ratio (>99%)"
"Change Causing MI vs Other MI's (<10%)"
"Failed Change Ratio (<1%)"
"Urgent Change rate (<5%)"

These embedded thresholds are the target values.
T3 (KPI Parser) must extract these using regex.
Pattern: r'\(([<>≥≤]=?)\s*(\d+\.?\d*)\s\*(%?)\)'

### Fact 7: Some columns contain mixed/text data

Example from NOC: FCR % column contains "(479/1736) 28%"
This is a STRING, not a float.
Classification: "text" type. Exclude from numeric charts.
Still display in KPI table if it's a KPI column.

### Fact 8: ITSM_Score_Card is a special cross-team sheet

This sheet consolidates ALL teams in one wide table.
Row 3 onward = one row per team.
Columns span MIM KPIs, Problem KPIs, Change KPIs, etc.
This sheet requires SPECIAL HANDLING — it is not a simple team sheet.
For now: skip it in the main pipeline but do not delete it.
Future: use it as the source for a consolidated summary slide.

---

## NEW FEATURES TO ADD (do not change existing working code)

The existing POC works. These are the ADDITIONS needed.

### Addition 1: Chart Customisation via Report_Config Sheet

The user wants to control charts from inside the Excel file.
No YAML editing, no code access. They are non-technical users.

#### Add a "Report_Config" sheet to the Excel workbook

This sheet is the user-facing configuration interface.

Structure:
Row 1: Headers (dark blue fill, white text, bold)
Row 2+: One row per team

Columns:
A: Team Name (text — must match sheet name exactly)
B: Chart 1 Type (dropdown)
C: Chart 1 Columns (text: "auto" or comma-separated column names)
D: Chart 2 Type (dropdown or blank)
E: Chart 2 Columns (text or blank)
F: Chart 3 Type (dropdown or blank)
G: Slide Layout (dropdown)
H: Summary Mode (dropdown: "ai_write" | "use_excel" | "ai_refine")
I: Include Insights (dropdown: "yes" | "no")
J: Skip This Team (dropdown: "yes" | "no")
K: Priority (dropdown: "high" | "normal")

Excel Data Validation dropdowns to create:
Chart Type (columns B, D, F):
"auto,rag_scorecard,grouped_bar,pie,stacked_bar_line,
multi_panel_bar,bar_line_combo,line,bar_dotted_line,
simple_bar,dashboard_grid,none"

"none" means: replace chart area with a data table instead.
"auto" means: system auto-detects best chart for this data.

Slide Layout (column G):
"standard,dashboard_grid,minimal,full_chart"

Summary Mode (column H):
"ai_write" — AI generates summary from scratch (default)
"use_excel" — Use "Weekly Comments/Achievements" as-is, no AI summary
"ai_refine" — AI reads Excel summary and improves/expands it

Include Insights (column I):
"yes,no"

Skip This Team (column J):
"yes,no"

Priority (column K):
"high,normal"

Pre-fill defaults for known teams:
MIM | auto | auto | - | - | - | standard | ai_write | yes | no | normal
Problem | auto | auto | - | - | - | standard | ai_write | yes | no | normal
Change | dashboard_gr | auto | - | - | - | dashboard_grid| ai_write | yes | no | normal
CMDB | auto | auto | - | - | - | standard | ai_write | yes | no | normal
NOC | auto | auto | - | - | - | standard | ai_write | yes | no | normal
Monitoring| auto | auto | - | - | - | standard | ai_write | yes | no | normal
ITOR | auto | auto | - | - | - | standard | ai_write | yes | no | normal
APAC SD | auto | auto | - | - | - | standard | ai_write | yes | no | normal
EMEA SD | auto | auto | - | - | - | standard | ai_write | yes | no | normal
ServiceNow| auto | auto | - | - | - | standard | ai_write | yes | no | normal

#### Config priority order (implement exactly):

1. Report_Config sheet in Excel (highest — user's live choice)
2. chart_config.yaml (developer defaults)
3. Auto-detection in T5 (last resort)

#### Where to add this code:

Add read_report_config() method to T2 (ExcelReader).
Call it in W2 (parse workflow) before processing team sheets.
Pass the config dict to orchestrator for use in W3, W4, W5.

#### Method signature:

def read_report_config(self, file_path: str) -> dict:
"""
Returns: {
"MIM": {
"charts": [{"type": "grouped_bar", "columns": "auto"}, ...],
"layout": "standard",
"summary_mode": "ai_write", # or "use_excel" or "ai_refine"
"include_insights": True,
"skip": False,
"priority": "normal"
},
...
}
Returns {} if Report_Config sheet not found (fall through to defaults).
"""

---

### Addition 2: User Summary Handling (Summary Mode)

The "Weekly Comments/Achievements" column in each sheet contains
the team lead's narrative. This is rich text with achievements,
reasons for KPI results, bullet points.

The Report_Config "Summary Mode" column controls what happens:

#### Mode: "use_excel" (user wrote summary, AI not needed for summary)

- Extract text from "Weekly Comments/Achievements" column (last non-empty row)
- Use this text directly as the slide summary
- Skip AI summary generation entirely for this team
- AI still generates kpi_evaluation, key_achievements, insights
  (unless include_insights = "no")
- Slide summary box shows: the Excel text (truncated to 300 chars if too long)
- Add indicator on slide: small italic "(Team-provided summary)"

#### Mode: "ai_refine" (user wrote summary, AI should improve it)

- Extract text from "Weekly Comments/Achievements" column
- Send to AI with instruction: "Refine and improve this summary for
  an executive presentation. Keep all facts. Make it professional.
  Max 3 sentences."
- AI returns refined version
- Slide shows the AI-refined version

#### Mode: "ai_write" (default — AI generates from scratch)

- Normal behaviour: AI generates summary from data
- Check if "Weekly Comments/Achievements" is non-empty
- If it IS non-empty: pass it to AI as CONTEXT (not as the summary)
  Add to prompt: "Team context provided: {excel_summary}"
  This gives AI richer information to write a better summary.

#### Implementation in T4 (AI Engine):

Add summary_mode parameter to generate_analysis():

def generate_analysis(self, sheet_name, headers, rows,
kpi_dict, excel_summary="",
summary_mode="ai_write") -> dict:

Adjust prompt based on mode:
"use_excel":
System prompt: generate only kpi_evaluation + insights + achievements
Skip summary generation
Return: {"summary": excel_summary, "kpi_evaluation": [...], ...}

"ai_refine":
User prompt includes: "Refine this summary: {excel_summary}"
AI returns improved version in "summary" field

"ai_write" with non-empty excel_summary:
User prompt includes: "Team context: {excel_summary}"
AI writes fresh summary using this as reference

---

### Addition 3: Chart System — 10 Chart Types + "none" Option

The current T5 likely generates basic charts.
This addition creates a proper chart renderer system.

#### New folder structure to add:

chart_renderers/
**init**.py
base_renderer.py ← abstract base class
c1_rag_scorecard.py ← KPI table with RAG dots
c2_grouped_bar.py ← grouped bar (multiple series)
c3_pie.py ← pie chart
c4_stacked_bar_line.py ← stacked bars + overlay lines
c5_multi_panel.py ← multiple subplots (one per BU)
c6_bar_line_combo.py ← bars + line on same chart
c7_line.py ← pure line chart
c8_bar_dotted_line.py ← bar + dotted line combo
c9_simple_bar.py ← single metric bar chart
c10_dashboard_grid.py ← multiple charts in grid layout
auto_selector.py ← decides which chart type fits the data

#### IMPORTANT: Do NOT break existing T5

If T5 currently works, preserve its interface.
Add chart_renderers as a NEW module that T5 calls internally.
T5 becomes the orchestrator that:

1. Reads chart config for this sheet
2. Calls the appropriate renderer
3. Returns PNG bytes (or None if "none" / no data)

#### Base renderer contract:

class BaseRenderer:
def render(self, df, chart_spec, history=None) -> bytes | None:
"""Returns PNG bytes or None. NEVER raises an exception."""
try:
return self.\_render_impl(df, chart_spec, history)
except Exception as e:
logger.warning(f"{self.**class**.**name**} failed: {e}")
return None # Slide will show data table instead

def \_render_impl(self, df, chart_spec, history) -> bytes:
raise NotImplementedError

#### Auto-selector logic (auto_selector.py):

def auto_select(df, column_types) -> str:
"""
Returns chart type string based on data shape.
Called when user specifies "auto" or no config.
"""
numeric_cols = [c for c,t in column_types.items()
if t in ("numeric", "percent")]
percent_cols = [c for c,t in column_types.items()
if t == "percent"]
row_count = len(df.dropna(how='all'))

if row_count == 0:
return "none"
if row_count == 1 and len(percent_cols) >= 3:
return "rag_scorecard"
if row_count == 1 and len(numeric_cols) >= 2:
return "grouped_bar"
if row_count >= 2 and len(numeric_cols) >= 2:
return "grouped_bar" # or stacked if many series
if row_count >= 2 and len(numeric_cols) == 1:
return "simple_bar"
if len(percent_cols) == 2:
return "pie"
return "simple_bar" # fallback

#### "none" chart type behaviour:

When chart type = "none" OR chart returns None:
DO NOT leave chart area blank.
Instead: render a data TABLE in the chart area.
Table shows: the raw data from the sheet (last N rows).
Table styling: dark blue header, alternating row colors.
This gives "no chart" users a clean data view instead.

#### Data table fallback (add to T6 PPTXBuilder):

def \_add_data_table(self, slide, df, zone_coords):
"""
Renders a python-pptx native table from DataFrame.
Used when chart = None or chart_type = "none".
Shows last 4 rows maximum.
Header row: dark blue background, white text.
Alternating data rows: white / light gray.
Numeric cells: right-aligned.
Text cells: left-aligned.
"""

---

### Addition 4: Chart Renderer Specifications

Implement each renderer to match the reference images provided.

#### C1: RAG Scorecard (c1_rag_scorecard.py)

Reference: Image 1 top section (KPI table with colored dots)

What it renders:
Table where each row = one KPI
Columns: KPI Name | Target | Actual | Status (colored dot)
Last column: green/amber/red circle only, no text
Header: dark blue background, white text

Input from chart_spec:
columns: list of KPI column names (or "auto" = all percent cols)
rag_thresholds: {green: 100, amber: 70} (percent values)

Implementation:
Use matplotlib NOT python-pptx table (for consistent PNG output)
Draw a matplotlib figure with no axes
Use ax.table() for the grid
For status dots: draw colored circles using ax.plot()
figsize: (9, max(2, len(kpis) \* 0.6))

RAG evaluation:
actual_pct = value \* 100 if value <= 1.0 else value
if actual_pct >= green_threshold: color = "#00B050" (green)
elif actual_pct >= amber_threshold: color = "#FFC000" (amber)
else: color = "#FF0000" (red)

Also render "Overall RAG" as a single large dot:
Overall = worst individual KPI RAG status

#### C2: Grouped Bar (c2_grouped_bar.py)

Reference: Image 1 bottom-left (Priority Split — 4 week rolling)

What it renders:
Multiple colored bar groups per x-axis category
Each x = time period (week/sprint), each bar = one series
Value labels above each bar

Implementation:
bar_width = 0.35
n_series = len(series)
for i, series_spec in enumerate(series):
offset = (i - n_series/2 + 0.5) \* bar_width
ax.bar(x + offset, values, width=bar_width, color=color, label=label)
ax.bar_label() for value labels
Legend: bottom center
Remove top and right spines

MUFG colors for default series (when user doesn't specify):
Series 1: "#1F3864" (dark blue)
Series 2: "#C55A11" (orange-red) ← matches Image 1
Series 3: "#2E75B6" (mid blue)
Series 4: "#0F6E56" (teal)

#### C3: Pie Chart (c3_pie.py)

Reference: Image 1 bottom-right (Root Cause split)

What it renders:
Pie chart with percentage + count labels outside
Label format: "Category\nCount, XX%"
No shadow, start angle 90

Implementation:
wedges, texts, autotexts = ax.pie(
values, labels=None, autopct='%1.0f%%',
startangle=90, colors=colors
)
Add custom labels outside:
for i, (wedge, label, count) in enumerate(zip(wedges, labels, counts)):
angle = (wedge.theta2 + wedge.theta1) / 2
x = 1.3 _ cos(radians(angle))
y = 1.3 _ sin(radians(angle))
ax.text(x, y, f"{label}\n{count}, {pct:.0f}%",
ha='center', va='center', fontsize=9)

Colors: ["#F0C040", "#2E75B6", "#404040", "#C00000", "#0F6E56"]

Data source for pie:
If a text column has few unique values (<=8): use value_counts()
If user specifies columns: sum those columns
Example: "Change: 1", "Application: 1" from two separate rows

#### C4: Stacked Bar + Line Combo (c4_stacked_bar_line.py)

Reference: Image 2 (Total Problems By BU — Rolling 12 Monthly)

What it renders:
Multiple series stacked on top of each other (bars)
Two overlay lines (e.g. Ageing trend, Closed trend)
Total label above each stacked bar
Value labels inside each segment

Implementation:
bottom = np.zeros(len(df))
for each stacked series:
bars = ax.bar(x, values, bottom=bottom, color=color, label=label)
for bar, val in zip(bars, values):
if val > 0:
ax.text(bar.get_x() + bar.get_width()/2,
bar.get_y() + bar.get_height()/2,
str(int(val)), ha='center', va='center',
fontsize=8, color='white')
bottom += values

Total labels above bars:
for i, total in enumerate(totals):
ax.text(i, total + 1, str(int(total)), ha='center',
va='bottom', fontsize=9, fontweight='bold')

Lines (secondary or same axis):
for line_spec in line_series:
ax2.plot(x, values, color=color, marker='o',
linewidth=2, markersize=5, label=label)
for i, val in enumerate(values):
ax2.annotate(str(val), (x[i], val),
textcoords="offset points", xytext=(0,6),
ha='center', fontsize=8)

Legend: combine ax and ax2 legends at bottom

#### C5: Multi-Panel Bar (c5_multi_panel.py)

Reference: Image 3 (By BU charts — 3 side-by-side panels)

What it renders:
One subplot per unique value in a panel_column (e.g. Business Unit)
Each panel: grouped bars for that BU across time periods
Shared title above all panels
Individual panel titles

Implementation:
panels = df[panel_column].unique()[:max_panels]
fig, axes = plt.subplots(1, len(panels),
figsize=(5 \* len(panels), 3.5))
if len(panels) == 1: axes = [axes]

for ax, panel_val in zip(axes, panels):
panel_df = df[df[panel_column] == panel_val]
draw grouped bars on ax
ax.set_title(str(panel_val), fontsize=10, fontweight='bold')
ax.set_ylabel("Change Request Count" if first else "")

plt.suptitle(title, fontsize=11, fontweight='bold', y=1.02)
plt.tight_layout()

#### C6: Bar + Line Combo (c6_bar_line_combo.py)

Reference: Image 4 top-left (Total vs Successful Changes)

What it renders:
Dark bars for total/main metric
Red line for comparison metric (Successful)
Value labels on bars and line points

Colors matching Image 4:
Bars: "#404040" (dark gray)
Line: "#CC2222" (red)

#### C7: Line Chart (c7_line.py)

Reference: Image 4 top-middle and top-right

Simple line with markers and value labels at each point.
Grid: horizontal only, light gray.

#### C8: Bar + Dotted Line Combo (c8_bar_dotted_line.py)

Reference: Image 5 left (Problem Tickets Closed vs Opened)

Orange bars = closed metric
Dotted blue line = open metric
Both on same axis

Line: linestyle=':', linewidth=2.5, marker='s'

#### C9: Simple Bar (c9_simple_bar.py)

Reference: Image 5 right (Ageing Backlog), Image 4 bottom panels

Single metric, one bar per time period.
bar_color defaults to group-specific color.
Value labels above bars.

#### C10: Dashboard Grid (c10_dashboard_grid.py)

Reference: Image 4 (full Change dashboard — 6 charts in 2x3 grid)

Renders ALL configured charts into ONE PNG using gridspec.
Used when slide layout = "dashboard_grid".

Implementation:
fig = plt.figure(figsize=(14, 8))
gs = gridspec.GridSpec(rows, cols, figure=fig,
hspace=0.45, wspace=0.35)

for i, chart_spec in enumerate(charts):
row_idx = i // cols
col_idx = i % cols
ax = fig.add_subplot(gs[row_idx, col_idx])
renderer = get_renderer(chart_spec["type"])
renderer.render_in_axis(ax, df, chart_spec)

Returns single PNG bytes.
T6 embeds this as one full-width image on the slide.

NOTE: Each renderer must have BOTH:
render(df, chart_spec) -> bytes (standalone PNG)
render_in_axis(ax, df, chart_spec) (for grid embedding)

---

### Addition 5: Existing T5 Integration

DO NOT rewrite T5 if it works.
EXTEND it to use chart_renderers:

Add to T5:

- Import all renderers from chart_renderers/
- Add read_chart_config() method that reads Report_Config from Excel
- Add generate_for_sheet() method that:
  1. Gets chart config from Report_Config sheet
  2. Falls back to chart_config.yaml
  3. Falls back to auto_selector
  4. Calls appropriate renderer
  5. Returns list of {chart_id, png_bytes, position, title}
  6. Returns None for chart_type="none"
     (T6 will use data table instead)

---

### Addition 6: T6 Slide Builder — Zone-Based Layout

If T6 already places charts on slides, preserve that logic.
Add the following enhancements:

#### Enhancement A: Data table fallback

When charts list is empty OR all charts returned None:
Call \_add_data_table(slide, df, zone_coords)
This shows the raw data table in the chart zone.
Never leave the chart zone blank.

#### Enhancement B: Multiple chart placement

When a team has 2-3 charts:
Layout: "standard" with positions bottom_left + bottom_right
If 3 charts: reduce sizes, place in thirds

#### Enhancement C: Dashboard grid layout

When Report_Config says layout = "dashboard_grid":
C10 DashboardGridRenderer produces ONE PNG
Embed that single PNG in the full chart zone
Summary area is smaller (1 line only)

#### Enhancement D: Summary mode rendering

When summary_mode = "use_excel":
Display summary with small italic note: "(Team-provided)"
Show summary in normal summary box
No changes to layout

---

## EXCEL READING CHANGES

### Changes needed in T2 (ExcelReader):

#### Change A: Split detection (read only left half)

Current behavior: reads all columns
Required behavior: detect and read only the weekly data (left half)

def \_detect_weekly_columns(self, ws) -> list:
"""
Returns list of column indices belonging to the LEFT (weekly) half.

Strategy:

1. Read row 1 (headers)
2. Find the first occurrence of a repeated header OR
   Find a blank column after first 3+ data columns
3. Return column indices BEFORE that point

Example MIM sheet:
Headers: Month, Monday of week start, Week, Total MI,
100% Adherence..., ..., Weekly Comments/Achievements,
[blank], [blank], Quarter, Month, ...
Split point: first blank column after column 15
Return indices: 0 to 14 (15 columns)
"""

#### Change B: Weekly Comments extraction

Add method to extract user summary:

def get_user_summary(self, df) -> str:
"""
Looks for column named "Weekly Comments/\nAchievements"
or any column containing "Comments" AND "Achievements".
Returns the non-empty text from the LAST data row.
Returns "" if column not found or all empty.
"""
comment_col = None
for col in df.columns:
if "comment" in str(col).lower() and "achievement" in str(col).lower():
comment_col = col
break

if comment_col is None:
return ""

series = df[comment_col].dropna()
if len(series) == 0:
return ""

return str(series.iloc[-1]).strip()

#### Change C: Decimal to percent conversion

Add to \_classify_column_type():

def \_classify_column_type(self, col_name, series):
if "%" in str(col_name) or any(
kw in str(col_name).lower()
for kw in ["ratio", "rate", "accuracy", "compliance",
"adherence", "percent", "fcr", "sla"]
): # Check if values are in 0-1 range (decimal) or 0-100 range
numeric = pd.to_numeric(series.dropna(), errors='coerce').dropna()
if len(numeric) > 0 and numeric.max() <= 1.0:
return "percent_decimal" # needs \* 100 for display
return "percent"

# ... rest of existing logic

Then in display: if type == "percent_decimal": value \* 100

#### Change D: Handle duplicate sheet names

def \_deduplicate_sheets(self, sheet_names) -> list:
"""
Returns list of sheets to process, removing duplicates.
Rule: if "MIM1" and "MIM" both exist, keep "MIM".
Rule: if "MIM (2)" and "MIM" both exist, keep "MIM".
Strips trailing digits and spaces/brackets for comparison.
"""

#### Change E: Read Report_Config sheet

def read_report_config(self, file_path) -> dict:
"""See Addition 1 for full spec."""

---

## AI PROMPT — UPDATED STRUCTURE

Update T4's prompt to include:

1. Summary mode handling (use_excel / ai_refine / ai_write)
2. Excel summary as context when available
3. Note about decimal-to-percent conversion already done

System prompt (update if different from current):
"You are a senior IT service management analyst at a financial
institution (MUFG). You analyze weekly operational data for
ITSM teams and produce executive-ready reports.
Respond ONLY with valid JSON. No markdown. No explanation.
No text outside the JSON object."

User prompt structure:
Group: {sheet_name}
Reporting period: Week {week_no} / {date_range}

KPI Data (values already converted to percentages):
{table of: kpi*name | actual*% | target\_% | status}

{if excel_summary is not empty and mode == "ai_write":}
Team context (use as reference for your summary):
{excel_summary[:500]}

{if mode == "ai_refine":}
Refine this team-provided summary for executive presentation.
Keep all facts. Max 3 sentences. Professional tone.
Original: {excel_summary[:500]}

Respond with ONLY:
{
"summary": "...",
"kpi_evaluation": [
{"kpi": "name", "value": "actual_display",
"threshold": "target_display", "status": "PASS/FAIL"}
],
"key_achievements": ["...", "..."],
"insights": ["...", "..."],
"overall_rag": "GREEN" | "AMBER" | "RED"
}

{if mode == "use_excel":}
Generate ONLY kpi_evaluation, key_achievements, insights, overall_rag.
Set "summary" = "" (will be replaced with team-provided text).

---

## WHAT TO SKIP / IGNORE

Sheets to always skip in the pipeline:
Sheet1
Cover Page
ITSM_Score_Card (special — skip for now, do not delete)
Any sheet name ending in digit where non-digit version exists
Any sheet name containing "(2)", "(3)" etc. where base exists

Columns to always exclude from charts:
Any column where > 30% of values are strings (text type)
"Month", "Week No.", "Week", "Monday of week start", "Quarter"
"Monthly Comments/Achievements", "Weekly Comments/Achievements"

---

## FILES TO CHECK FOR REMOVAL

After reading all files:

- Check if any of the workflow files (w1-w5) are empty stubs
- Check if tests still pass: pytest tests/ -v
- Check if templates/slide_template.pptx is used by T6
  If not used: ask user before removing
- Check output/ITSM_Report_sheet_state.json — is this used?
  If yes: preserve and understand its purpose
  If no: ask user

DO NOT remove without asking:

- Any file that existing tests import
- Any file that main.py or orchestrator.py imports
- Any sample data or template file

---

## CONFIG.YAML — WHAT TO ADD

Add these new sections to the EXISTING config.yaml.
Do NOT remove existing keys.

New keys to add:

excel:

# existing keys preserved...

sheets_to_skip: - "Sheet1" - "Cover Page" - "ITSM_Score_Card" - "ITSM Score Card"
duplicate_detection: true # NEW: skip numbered variants
report_config_sheet: "Report_Config" # NEW: config sheet name

summary: # NEW section
default_mode: "ai_write" # "ai_write" | "use_excel" | "ai_refine"
excel_summary_max_chars: 500 # truncate before sending to AI

charts: # NEW section
default_type: "auto"
fallback_to_table: true # show data table when no chart
table_max_rows: 4 # max rows in fallback table
mufg_colors: - "#1F3864" # dark blue - "#C55A11" # orange - "#2E75B6" # mid blue - "#0F6E56" # teal - "#C00000" # red - "#404040" # dark gray - "#F0C040" # amber/yellow

---

## SAMPLE DATA UPDATE

The existing sample_itsm.xlsx should be updated to match
the real Excel template structure.

But ONLY update if tests still pass with new structure.
Do NOT modify the real Scorecard_Template.xlsx — that is
the reference template, not the working data file.

Updates needed in sample_data/sample_itsm.xlsx:

1. Add "Report_Config" sheet with defaults for all teams
2. Ensure "Weekly Comments/Achievements" column exists in each team sheet
3. Ensure KPI values are in 0-1 decimal range (not 0-100)
4. Add 3-4 rows of data per team (for trend charts)
5. Keep existing structure that the POC already reads correctly

If updating sample data breaks existing tests, fix the tests
to handle the new structure rather than reverting the data.

---

## BUILD ORDER FOR NEW FEATURES

Add features in this exact order.
After each step, run: pytest tests/ -v
If tests fail, fix before proceeding.

Step 1: Add read_report_config() to T2

- Test: create mock Excel with Report_Config sheet
- Verify config dict returned correctly
- Verify fallback works when sheet absent

Step 2: Add \_detect_weekly_columns() to T2

- Test: verify MIM sheet only returns left-half columns
- Verify quarterly/monthly columns excluded

Step 3: Add get_user_summary() to T2

- Test: verify "Weekly Comments/Achievements" text extracted
- Verify empty string when column absent

Step 4: Add decimal-to-percent detection to T2/T3

- Test: 0.6666 → "66.7%" in display
- Test: 100.0 → "100%" in display (already percent)

Step 5: Add summary_mode handling to T4

- Test with mode="use_excel": no AI summary generation
- Test with mode="ai_refine": prompt includes refine instruction
- Test with mode="ai_write" + excel_summary: context included

Step 6: Create chart_renderers/ folder

- Start with base_renderer.py and auto_selector.py
- Then c9_simple_bar.py (simplest)
- Then c7_line.py
- Then c2_grouped_bar.py
- Then c3_pie.py
- Then c1_rag_scorecard.py
- Then c6_bar_line_combo.py
- Then c8_bar_dotted_line.py
- Then c4_stacked_bar_line.py
- Then c5_multi_panel.py
- Finally c10_dashboard_grid.py

Step 7: Update T5 to use chart_renderers

- Preserve existing T5 interface
- Add generate_for_sheet() that reads Report_Config
- Test: python main.py --run-once generates charts

Step 8: Update T6 for data table fallback + multi-chart zones

- Test: slide with no charts shows data table

Step 9: Update sample_data/sample_itsm.xlsx

- Add Report_Config sheet
- Verify tests still pass

Step 10: Full integration test

- python main.py --run-once
- Open PPTX — verify each slide has correct layout
- Verify MIM slide shows charts based on Report_Config
- Verify summary modes work correctly

---

## SUCCESS CRITERIA

System is working with new features when ALL of these are true:

1. pytest tests/ -v — ALL tests pass
2. python main.py --run-once — completes without errors
3. Output PPTX exists and opens in PowerPoint
4. MIM slide reflects chart type from Report_Config sheet
5. When Report_Config has chart_type="none": data table appears
6. When Summary Mode="use_excel": Excel text appears on slide
7. When Summary Mode="ai_refine": refined version appears
8. Decimal KPI values (0.6666) display as "66.7%" on slides
9. Duplicate sheets (MIM1, APAC SD1) are skipped silently
10. "Weekly Comments/Achievements" text reaches AI as context
    when Summary Mode="ai_write"
11. Changing Report_Config in Excel changes next report output
12. No working code from the existing POC is broken

---

## QUESTIONS TO ASK BEFORE STARTING

1. Run: pytest tests/ -v
   Report which tests pass and which fail currently.

2. Open output/ITSM_Report_20260408.pptx (the latest one).
   Describe what each slide currently looks like.
   This tells us the current baseline to preserve and extend.

3. Does tools/t5_chart_generator.py currently use matplotlib?
   What chart types does it generate today?

4. Does tools/t6_pptx_builder.py have a zone/position system,
   or does it use fixed pixel coordinates?
   This determines how much of it to preserve vs extend.

5. Is output/ITSM_Report_sheet_state.json actively used?
   What is it tracking? (Likely used for change detection.)

6. Does the current code already handle the decimal-to-percent
   conversion for KPI values?
   If yes: do not add it again.

---

## FINAL RULE

The existing POC works. Every change you make must preserve
that working state. If a change breaks something, revert it
and ask the user before trying again.

When in doubt: ask. Do not assume.
