from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .core import ExcelDiffError, compare


def _text(report, limit: int) -> str:
    lines = [f"Excel Diff: {report.left} -> {report.right}"]
    if report.identical:
        return "\n".join(lines + ["Result: identical"])
    lines.append("Result: differences found")
    if report.added_sheets:
        lines.append("Added sheets: " + ", ".join(report.added_sheets))
    if report.removed_sheets:
        lines.append("Removed sheets: " + ", ".join(report.removed_sheets))
    shown = report.changes[:limit]
    for c in shown:
        lines.append(f"{c.sheet}!{c.cell}: {c.before!r} -> {c.after!r}")
    if len(report.changes) > len(shown):
        lines.append(f"... {len(report.changes) - len(shown)} more cell change(s)")
    lines.append(f"Changed cells: {len(report.changes)}")
    return "\n".join(lines)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="excel-diff", description="Compare XLSX/XLSM workbooks or CSV files.")
    p.add_argument("left")
    p.add_argument("right")
    p.add_argument("--format", choices=("text", "json"), default="text", dest="output_format")
    p.add_argument("--output", type=Path, help="Write report to a file instead of stdout")
    p.add_argument("--formulas", action="store_true", help="Compare formula expressions instead of cached results")
    p.add_argument("--ignore-case", action="store_true", help="Ignore letter case in string cells")
    p.add_argument("--limit", type=int, default=200, help="Maximum changed cells shown in text output")
    p.add_argument("--version", action="version", version="excel-diff 1.0.0 — Radwan Abdulhadi Ahmed / @rad03i2")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.limit < 0:
        print("error: --limit must be >= 0", file=sys.stderr)
        return 2
    try:
        report = compare(args.left, args.right, formulas=args.formulas, ignore_case=args.ignore_case)
    except ExcelDiffError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    rendered = report.to_json() if args.output_format == "json" else _text(report, args.limit)
    if args.output:
        try:
            args.output.write_text(rendered + "\n", encoding="utf-8")
        except OSError as exc:
            print(f"error: cannot write report: {exc}", file=sys.stderr)
            return 2
    else:
        print(rendered)
    return 0 if report.identical else 1


if __name__ == "__main__":
    raise SystemExit(main())
