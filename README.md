# AI Usage Report

Generate consolidated AI Dev Journal Excel reports from multiple staff journal files, then add chart sheets such as:

- `🧭 SDLC Summary` — task counts, efficiency % per SDLC Stage, and a stage-axis chart with tasks inside each SDLC Stage
- `📈 Efficiency Charts` — EST vs Actual and time saved
- `⭐ Rating Charts` — user satisfaction by tool
- `🏷️ Error Charts` / `🏷️ Prompt Error Data` — prompt error classification
- Professional Excel polish — themed tab colors, hidden gridlines, freeze panes, filters, number formats, and color-scale highlights

## Requirements

Use Python 3.10+.

Install the Python packages used by the scripts:

```bash
python -m pip install openpyxl requests python-dotenv matplotlib numpy
```

If you want AI lesson inference and SDLC/error classification, create a `.env` file:

```bash
OPENAI_BASE_URL=https://your-openai-compatible-endpoint
OPENAI_API_KEY=your-api-key
```

You can skip all AI calls with `--no-ai`.

## Input Files

Place staff journal workbooks in `data/`.

Recommended filename format:

```text
data/journal_<staff-name>.xlsx
```

Example:

```text
data/journal_binh.xlsx
data/journal_giangdn.xlsx
```

Each workbook should follow `template.xlsx` and include the journal sheet:

```text
📝 Nhật Ký
```

## Main Command: Report + Excel Charts

Generate the final workbook with report sheets and chart sheets in one command:

```bash
python report_with_charts.py data/*.xlsx -o out/report_with_charts.xlsx --model gpt-5.4-mini
```

Without AI calls:

```bash
python report_with_charts.py data/*.xlsx -o out/report_with_charts.xlsx --no-ai
```

Keep the intermediate report workbook:

```bash
python report_with_charts.py data/*.xlsx \
  -o out/report_with_charts.xlsx \
  --report-output out/report.xlsx \
  --model gpt-5.4-mini
```

Use different models for the report step and chart classification step:

```bash
python report_with_charts.py data/*.xlsx \
  -o out/report_with_charts.xlsx \
  --report-model gpt-5.4-mini \
  --chart-model gpt-5.4-mini
```

## Step-by-Step Commands

### 1. Generate consolidated report only

```bash
python report.py data/*.xlsx -o out/report.xlsx --model gpt-5.4-mini
```

Without AI:

```bash
python report.py data/*.xlsx -o out/report.xlsx --no-ai
```

### 2. Add Excel chart sheets to an existing report

```bash
python plot_charts.py out/report.xlsx -o out/report_with_charts.xlsx --model gpt-5.4-mini
```

Without chart AI classification:

```bash
python plot_charts.py out/report.xlsx -o out/report_with_charts.xlsx --no-ai
```

### 3. Generate PNG/PDF charts

This is the legacy visual chart export path:

```bash
python charts.py out/report.xlsx -o charts_output --model gpt-5.4-mini
```

Without AI:

```bash
python charts.py out/report.xlsx -o charts_output --no-ai
```

Output includes:

```text
charts_output/ai_journal_charts.pdf
charts_output/*.png
```

The PNG/PDF export also includes:

```text
charts_output/01_sdlc_tasks_by_stage.png  # includes task count + efficiency % by SDLC Stage
charts_output/02_staff_ai_effectiveness.png
```

In the PDF, `01_sdlc_tasks_by_stage.png` is generated as the first slide/page, and `02_staff_ai_effectiveness.png` is generated as the next slide/page.

For the best SDLC task grouping in `charts.py`, run it on the enriched workbook from `report_with_charts.py` or `plot_charts.py`; otherwise `charts.py` will classify SDLC stages itself when AI is enabled.

## Useful Help Commands

```bash
python report.py --help
python plot_charts.py --help
python report_with_charts.py --help
python charts.py --help
```

## Cache Files

The scripts reuse AI results through cache files:

```text
.ai_journal_cache.json
.ai_chart_cache.json
```

If you want to force fresh AI inference/classification, remove the cache files:

```bash
rm -f .ai_journal_cache.json .ai_chart_cache.json
```

## Common Workflow

```bash
# 1. Put journal files in data/
ls data/*.xlsx

# 2. Generate final report workbook with charts
python report_with_charts.py data/*.xlsx -o out/report_with_charts.xlsx --model gpt-5.4-mini

# 3. Open the workbook
open out/report_with_charts.xlsx
```
