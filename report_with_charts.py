#!/usr/bin/env python3
"""
AI Dev Journal - One-command report + charts generator.

This wrapper runs:
  1. report.py       -> consolidated report workbook
  2. plot_charts.py  -> enriched workbook with chart/SDLC sheets

So you can produce the final workbook with charts using one command.

Usage:
    python report_with_charts.py data/*.xlsx -o out/report_with_charts.xlsx
    python report_with_charts.py data/*.xlsx -o out/report_with_charts.xlsx --model gpt-5.4-mini
    python report_with_charts.py data/*.xlsx -o out/report_with_charts.xlsx --no-ai
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parent
REPORT_SCRIPT = ROOT / "report.py"
CHART_SCRIPT = ROOT / "plot_charts.py"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate the AI Dev Journal report and chart workbook in one command.",
    )
    parser.add_argument("files", nargs="+", type=Path, help="Input AI Dev Journal .xlsx files")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("ai_journal_report_with_charts.xlsx"),
        help="Final enriched workbook with charts (default: ai_journal_report_with_charts.xlsx)",
    )
    parser.add_argument(
        "--report-output",
        type=Path,
        help="Optional path to keep the intermediate report.py workbook; otherwise a temp file is used and deleted.",
    )
    parser.add_argument(
        "--model",
        default="qwen2.5:7b",
        help="OpenAI-compatible model used by both report.py and plot_charts.py unless overridden.",
    )
    parser.add_argument("--report-model", help="Model override for report.py AI lesson inference")
    parser.add_argument("--chart-model", help="Model override for plot_charts.py error/SDLC classification")
    parser.add_argument("--no-ai", action="store_true", help="Skip AI in both report.py and plot_charts.py")
    parser.add_argument("--no-report-ai", action="store_true", help="Skip only report.py AI lesson inference")
    parser.add_argument("--no-chart-ai", action="store_true", help="Skip only plot_charts.py error/SDLC classification")
    parser.add_argument(
        "--chart-cache",
        type=Path,
        default=Path(".ai_chart_cache.json"),
        help="Cache path passed to plot_charts.py (default: .ai_chart_cache.json)",
    )
    parser.add_argument(
        "--chart-batch-size",
        type=int,
        default=20,
        help="Batch size passed to plot_charts.py (default: 20)",
    )
    parser.add_argument(
        "--chart-timeout",
        type=int,
        default=300,
        help="AI request timeout passed to plot_charts.py in seconds (default: 300)",
    )
    return parser.parse_args(argv)


def _resolve_inputs(paths: list[Path]) -> list[Path]:
    resolved = [path.expanduser().resolve() for path in paths]
    missing = [str(path) for path in resolved if not path.exists()]
    if missing:
        raise FileNotFoundError("Input workbook(s) not found: " + ", ".join(missing))
    return resolved


def _run(cmd: list[str]) -> None:
    print("\n$ " + " ".join(cmd), flush=True)
    subprocess.run(cmd, cwd=ROOT, check=True)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    try:
        input_files = _resolve_inputs(args.files)
    except FileNotFoundError as exc:
        print(exc, file=sys.stderr)
        return 1

    output_path = args.output.expanduser().resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    temp_report_path: Path | None = None
    if args.report_output:
        report_output = args.report_output.expanduser().resolve()
        report_output.parent.mkdir(parents=True, exist_ok=True)
    else:
        with tempfile.NamedTemporaryFile(prefix="ai_journal_report_", suffix=".xlsx", delete=False) as temp_file:
            temp_report_path = Path(temp_file.name)
        report_output = temp_report_path

    report_model = args.report_model or args.model
    chart_model = args.chart_model or args.model

    report_cmd = [
        sys.executable,
        str(REPORT_SCRIPT),
        *[str(path) for path in input_files],
        "-o",
        str(report_output),
        "--model",
        report_model,
    ]
    if args.no_ai or args.no_report_ai:
        report_cmd.append("--no-ai")

    chart_cmd = [
        sys.executable,
        str(CHART_SCRIPT),
        str(report_output),
        "-o",
        str(output_path),
        "--model",
        chart_model,
        "--cache",
        str(args.chart_cache.expanduser().resolve()),
        "--batch-size",
        str(args.chart_batch_size),
        "--timeout",
        str(args.chart_timeout),
    ]
    if args.no_ai or args.no_chart_ai:
        chart_cmd.append("--no-ai")

    try:
        print("📘 Step 1/2: Building consolidated report workbook...", flush=True)
        _run(report_cmd)
        print("\n📊 Step 2/2: Adding chart and SDLC sheets...", flush=True)
        _run(chart_cmd)
    except subprocess.CalledProcessError as exc:
        print(f"\nCommand failed with exit code {exc.returncode}", file=sys.stderr)
        return exc.returncode
    finally:
        if temp_report_path and temp_report_path.exists():
            temp_report_path.unlink()

    print(f"\n✔  Final workbook with charts saved to: {output_path}", flush=True)
    if args.report_output:
        print(f"✔  Intermediate report workbook kept at: {report_output}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
