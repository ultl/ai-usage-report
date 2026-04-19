#!/usr/bin/env python3
"""
AI Dev Journal — Executive Text Report Generator

Reads the consolidated report .xlsx produced by ai_journal.py, extracts all
key data, sends it to an LLM to generate a professional executive summary,
and outputs a Markdown report with embedded chart references.

Usage:
    python generate_report.py report.xlsx --charts-dir charts_output
    python generate_report.py report.xlsx --charts-dir charts_output --model gpt-5.4-mini
    python generate_report.py report.xlsx --charts-dir charts_output -o executive_report.md
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from pathlib import Path
from typing import Any

import requests
from openpyxl import load_workbook

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")


def _call_openai(model: str, prompt: str, timeout: int = 600) -> str:
    base = (OPENAI_BASE_URL or "").rstrip("/")
    is_azure = "cognitiveservices.azure.com" in base or "openai.azure.com" in base

    if is_azure:
        host = re.sub(r"/openai(/v1)?$", "", base).rstrip("/")
        url = f"{host}/openai/deployments/{model}/chat/completions?api-version=2024-12-01-preview"
        headers = {"api-key": OPENAI_API_KEY or ""}
    else:
        url = f"{base}/chat/completions"
        headers = {}
        if OPENAI_API_KEY:
            headers["Authorization"] = f"Bearer {OPENAI_API_KEY}"

    body: dict[str, Any] = {
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.3,
        "max_completion_tokens": 8192,
    }
    if not is_azure:
        body["model"] = model

    r = requests.post(url, headers=headers, json=body, timeout=timeout)
    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"]


def _cell_str(ws, r: int, c: int) -> str:
    v = ws.cell(r, c).value
    return str(v).strip() if v is not None else ""


def _cell_num(ws, r: int, c: int) -> float | None:
    v = ws.cell(r, c).value
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        m = re.search(r"-?\d+(?:\.\d+)?", v.replace(",", "."))
        if m:
            return float(m.group(0))
    return None


def _find_section_row(ws, label: str) -> int | None:
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").strip().startswith(label):
            return r
    return None


def extract_data(xlsx_path: Path) -> dict[str, Any]:
    """Extract all report data into a structured dict for the LLM prompt."""
    wb = load_workbook(xlsx_path, data_only=True)
    data: dict[str, Any] = {}

    # === Dashboard ===
    if "📊 Dashboard" in wb.sheetnames:
        ws = wb["📊 Dashboard"]

        # KPIs
        kpis = {}
        for r in range(5, 20):
            metric = _cell_str(ws, r, 1)
            if not metric:
                continue
            assessed = _cell_str(ws, r, 2)
            self_rep = _cell_str(ws, r, 3)
            deviation = _cell_str(ws, r, 4)
            if assessed:
                kpis[metric] = {"assessed": assessed, "self_reported": self_rep, "deviation": deviation}
        data["kpis"] = kpis

        # Self-Report Accuracy
        acc_row = _find_section_row(ws, "SELF-REPORT ACCURACY")
        if acc_row:
            accuracy = []
            hr = acc_row + 1
            for r in range(hr + 1, hr + 20):
                staff = _cell_str(ws, r, 1)
                if not staff or staff.upper().startswith("TOTAL"):
                    break
                accuracy.append({
                    "staff": staff,
                    "assessed_saved": _cell_num(ws, r, 8),
                    "self_rep_saved": _cell_num(ws, r, 9),
                    "delta_saved": _cell_num(ws, r, 10),
                    "assessed_pct": _cell_num(ws, r, 11),
                    "self_rep_pct": _cell_num(ws, r, 12),
                    "accuracy": _cell_str(ws, r, 14),
                })
            data["self_report_accuracy"] = accuracy

        # By Staff breakdown
        staff_row = _find_section_row(ws, "BY STAFF")
        if staff_row:
            by_staff = []
            hr = staff_row + 1
            for r in range(hr + 1, hr + 20):
                name = _cell_str(ws, r, 1)
                if not name or name.upper().startswith("TOTAL"):
                    break
                by_staff.append({
                    "staff": name,
                    "sessions": _cell_num(ws, r, 2),
                    "assessed_saved": _cell_num(ws, r, 5),
                    "assessed_pct": _cell_num(ws, r, 6),
                    "self_rep_saved": _cell_num(ws, r, 9),
                    "self_rep_pct": _cell_num(ws, r, 10),
                    "avg_rating": _cell_num(ws, r, 12),
                    "primary_tool": _cell_str(ws, r, 14),
                })
            data["by_staff"] = by_staff

        # By Tool
        tool_row = _find_section_row(ws, "BY AI TOOL")
        if tool_row:
            by_tool = []
            hr = tool_row + 1
            for r in range(hr + 1, hr + 15):
                tool = _cell_str(ws, r, 1)
                if not tool or tool.upper().startswith("TOTAL"):
                    break
                by_tool.append({
                    "tool": tool,
                    "sessions": _cell_num(ws, r, 2),
                    "assessed_saved": _cell_num(ws, r, 5),
                    "assessed_pct": _cell_num(ws, r, 6),
                    "self_rep_saved": _cell_num(ws, r, 9),
                    "self_rep_pct": _cell_num(ws, r, 10),
                    "avg_rating": _cell_num(ws, r, 12),
                })
            data["by_tool"] = by_tool

        # By Category
        cat_row = _find_section_row(ws, "BY CATEGORY")
        if cat_row:
            by_cat = []
            hr = cat_row + 1
            for r in range(hr + 1, hr + 20):
                cat = _cell_str(ws, r, 1)
                if not cat or cat.upper().startswith("TOTAL"):
                    break
                by_cat.append({
                    "category": cat,
                    "sessions": _cell_num(ws, r, 2),
                    "assessed_saved": _cell_num(ws, r, 5),
                    "assessed_pct": _cell_num(ws, r, 6),
                    "self_rep_saved": _cell_num(ws, r, 9),
                    "self_rep_pct": _cell_num(ws, r, 10),
                })
            data["by_category"] = by_cat

        # Rating
        rating_row = _find_section_row(ws, "RATING DISTRIBUTION")
        if rating_row:
            ratings = []
            hr = rating_row + 1
            for r in range(hr + 1, hr + 15):
                tool = _cell_str(ws, r, 1)
                if not tool:
                    break
                ratings.append({
                    "tool": tool,
                    "avg_rating": _cell_num(ws, r, 8),
                    "total_rated": _cell_num(ws, r, 7),
                })
            data["ratings"] = ratings

    # === SDLC ===
    if "🧭 SDLC Summary" in wb.sheetnames:
        ws = wb["🧭 SDLC Summary"]
        sdlc = []
        for r in range(6, 18):
            stage = _cell_str(ws, r, 1)
            count = _cell_num(ws, r, 3)
            if stage and count and count > 0:
                sdlc.append({
                    "stage": stage,
                    "task_count": int(count),
                    "assessed_pct": _cell_num(ws, r, 5),
                    "self_rep_pct": _cell_num(ws, r, 6),
                })
        data["sdlc"] = sdlc

    # === Error Charts ===
    if "🏷️ Error Charts" in wb.sheetnames:
        ws = wb["🏷️ Error Charts"]
        errors = []
        for r in range(6, 20):
            label = _cell_str(ws, r, 1)
            count = _cell_num(ws, r, 2)
            desc = _cell_str(ws, r, 3)
            if label and count:
                errors.append({"label": label, "count": int(count), "description": desc})
        data["errors"] = errors

    return data


REPORT_PROMPT = """<role>
You are a senior management consultant writing an executive report for C-level leadership.
You analyze AI adoption data and write clear, data-driven insights with specific numbers.
You explain WHAT the numbers mean in plain business language, not just list them.
</role>

<context>
This report covers an AI Dev Journal pilot program where {n_staff} staff members used AI tools
across {n_sessions} development tasks. The report compares two perspectives:

- **Assessed (AI)**: An independent AI estimation of how long each task should take, both
  with and without AI tools. This is generated BLINDLY — the AI never sees the user's own
  numbers. It estimates based solely on the task description, category, tool used, and the
  staff member's profile (role, experience, tech stack). This serves as the OBJECTIVE baseline.
- **Self-Reported**: The staff member's own estimation of hours. This is SUBJECTIVE and
  reveals how the user PERCEIVES the difficulty of their work and the value of AI.

The GAP between these two perspectives reveals important behavioral insights:

1. **EST gap (Self-Rep EST < Assessed EST)** = The user thinks the task is EASIER than it
   objectively is. They underestimate how long the work would take without AI, which means
   they don't fully appreciate the complexity AI is handling for them.

2. **Saved gap (Self-Rep Saved < Assessed Saved)** = The user under-reports how much AI
   helped them. They may take AI assistance for granted, or not realize how much longer
   the task would have taken manually.

3. **Efficiency gap (Self-Rep % ≠ Assessed %)** = Shows whether the user over- or under-
   estimates their own productivity gain. A user claiming 79% efficiency when AI assesses
   52% is inflating their perceived gain. A user claiming 41% when AI assesses 54% is
   being too modest (or not recognizing AI's contribution).

4. **Accuracy = Self-Rep Saved / Assessed Saved × 100** = How closely the user's perception
   matches objective reality. 100% = perfect alignment. Below 50% = the user is missing
   more than half of AI's actual contribution.

Key formulas:
- EST = Estimated hours without AI (manual work)
- Actual = Hours spent with AI assistance
- Saved = EST − Actual
- Efficiency % = Saved / EST × 100
- Δ = Self-Reported − Assessed (negative = user under-reports)
</context>

<data>
{data_json}
</data>

<available_figures>
The following chart images are available to reference in the report:
- 01_sdlc_tasks_by_stage.png — SDLC stage breakdown with Assessed vs Self-Reported efficiency
- 02_staff_ai_effectiveness.png — Staff effectiveness comparison
- 03_kpi_summary.png — Executive KPI summary dashboard
- 04_est_actual_tool.png — EST vs Actual hours by AI tool
- 05_est_actual_category.png — EST vs Actual hours by category
- 06_rating_distribution.png — User satisfaction rating distribution
- 07_top_errors.png — Top prompt engineering errors
- 08_error_heatmap.png — Error distribution heatmap by staff
- 09_user_vs_ai_comparison.png — Assessed vs Self-Reported comparison per staff
</available_figures>

<instructions>
Write a professional executive report in Markdown format with the following structure:

# AI Dev Journal — Sprint 0 Executive Report

## 1. Executive Summary
- 3-4 bullet points with the most impactful findings
- Lead with the headline number (total assessed hours saved)
- Mention the self-reporting accuracy gap

## 2. Overall Impact Assessment
- Compare Assessed vs Self-Reported KPIs in a table
- Highlight the systematic under-reporting pattern
- EXPLAIN what the EST gap means: "Users estimated X hours for manual work, but AI
  assessed Y hours — meaning users think their tasks are Z% easier than they objectively
  are. This underestimation means users don't fully appreciate the complexity AI handles."
- EXPLAIN what the Saved gap means: "Users reported saving X hours, but AI assessed Y
  hours of actual savings — users are not recognizing Z hours of AI-delivered value."
- Reference figure: 03_kpi_summary.png

## 3. Self-Report Accuracy Analysis
- Table showing each staff member's accuracy
- For EACH person, explain what their accuracy means in plain language. Examples:
  - 94% accuracy: "Tester's perception closely matches reality — they understand how
    much AI helps them"
  - 33% accuracy: "Naduc11 perceives only a third of AI's actual contribution — they
    think tasks are simpler than they are and don't recognize how much AI accelerates
    their work"
- Identify the most and least accurate reporters
- Flag any anomalies (e.g., someone claiming higher efficiency % than AI assessed —
  meaning they think they're more productive than they objectively are)
- Reference figure: 09_user_vs_ai_comparison.png

## 4. AI Tool Effectiveness
- Which tools save the most hours (use Assessed numbers as the objective baseline)
- User satisfaction ratings
- Reference figures: 04_est_actual_tool.png, 06_rating_distribution.png

## 5. Category Analysis
- Which work categories benefit most from AI
- Where the biggest gaps between Assessed and Self-Reported exist
- EXPLAIN what the gaps mean: e.g., "In Database Design, users reported saving only 4.1h
  but AI assessed 11.5h — users drastically underestimate AI's contribution to complex
  data modeling tasks, possibly because they attribute the quality of AI output to their
  own domain knowledge rather than the tool."
- Reference figure: 05_est_actual_category.png

## 6. SDLC Stage Distribution
- How tasks distribute across development lifecycle stages
- Assessed vs Self-Reported efficiency per stage
- EXPLAIN anomalies: e.g., if a stage shows Self-Reported % much higher than Assessed %,
  it means users overestimate their productivity in that stage. If lower, they undervalue
  AI's help in that stage.
- Reference figure: 01_sdlc_tasks_by_stage.png

## 7. Prompt Engineering Quality
- Top errors staff make when prompting AI
- Recommendations for improvement
- Reference figures: 07_top_errors.png, 08_error_heatmap.png

## 8. Recommendations
- 4-5 specific, actionable recommendations based on the data
- Each recommendation should cite the specific data point that supports it

## 9. Methodology Note
- Brief explanation of how Assessed vs Self-Reported works
- Note that Assessed estimates are blind (AI doesn't see user's numbers)

Rules:
- Use SPECIFIC numbers from the data, not vague language
- Always prefer Assessed (AI) numbers as the objective truth
- When referencing a figure, use the format: ![Description](figure_filename.png)
- Use tables in Markdown format for comparisons
- Write in professional, concise English suitable for CEO/CTO audience
- Keep total length to 1000-1500 words (excluding tables)
- Flag any data anomalies or concerns
- CRITICAL: For every gap/deviation, EXPLAIN what it means in plain business language.
  Don't just say "Δ = -25.4h". Say "Naduc11 thinks their tasks are simpler than they
  are — they estimated 30.5h of manual work but AI assessed 70.5h, meaning they
  underestimate task complexity by 57%. As a result, they only recognize 33% of AI's
  actual contribution to their productivity."
- When accuracy is low, explain the behavioral insight: the user either (a) thinks tasks
  are easier than they are, (b) takes AI help for granted, or (c) doesn't realize how
  long the work would take without AI.
- When a user reports HIGHER efficiency % than AI assessed, explain they are overestimating
  their own productivity gain.
</instructions>"""


def generate_report(xlsx_path: Path, charts_dir: Path, model: str, output: Path) -> int:
    print(f"📊  Reading report: {xlsx_path}")
    data = extract_data(xlsx_path)

    if not data:
        print("No data extracted from the report.", file=sys.stderr)
        return 1

    n_sessions = 0
    n_staff = 0
    kpis = data.get("kpis", {})
    if "Total Sessions" in kpis:
        n_sessions = kpis["Total Sessions"].get("assessed", "0")
    if "Staff Count" in kpis:
        n_staff = kpis["Staff Count"].get("assessed", "0")

    data_json = json.dumps(data, ensure_ascii=False, indent=2, default=str)

    prompt = REPORT_PROMPT.format(
        n_staff=n_staff,
        n_sessions=n_sessions,
        data_json=data_json,
    )

    print(f"🤖  Generating executive report with {model}...")
    try:
        report_md = _call_openai(model, prompt)
    except Exception as e:
        print(f"Failed to generate report: {e}", file=sys.stderr)
        return 1

    # Clean up any markdown code fences the LLM might wrap around the output
    report_md = re.sub(r"^```(?:markdown)?\s*", "", report_md.strip())
    report_md = re.sub(r"\s*```$", "", report_md.strip())

    # Resolve figure paths relative to charts_dir
    def _resolve_figure(m: re.Match) -> str:
        alt = m.group(1)
        filename = m.group(2)
        fig_path = charts_dir / filename
        if fig_path.exists():
            return f"![{alt}]({fig_path})"
        return m.group(0)

    report_md = re.sub(r"!\[([^\]]*)\]\(([^)]+\.png)\)", _resolve_figure, report_md)

    output.write_text(report_md, encoding="utf-8")
    print(f"✔  Executive report saved to: {output}")
    print(f"   Word count: ~{len(report_md.split())}")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Generate an executive text report from AI Dev Journal data using LLM.")
    ap.add_argument("report", type=Path, help="Input report .xlsx (output of ai_journal.py)")
    ap.add_argument("--charts-dir", type=Path, default=Path("charts_output"),
                    help="Directory containing chart PNGs (default: charts_output)")
    ap.add_argument("--model", default="gpt-5.4-mini",
                    help="Model for report generation (default: gpt-5.4-mini)")
    ap.add_argument("-o", "--output", type=Path, default=Path("executive_report.md"),
                    help="Output Markdown file (default: executive_report.md)")
    args = ap.parse_args()

    if not args.report.exists():
        print(f"Report not found: {args.report}", file=sys.stderr)
        return 1

    return generate_report(args.report, args.charts_dir, args.model, args.output)


if __name__ == "__main__":
    sys.exit(main())
