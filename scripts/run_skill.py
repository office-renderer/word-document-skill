#!/usr/bin/env python3
"""Validate first, then run only targeted Word formatting repairs."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FORMATTER = ROOT / "scripts" / "format_docx.py"
VALIDATOR = ROOT / "scripts" / "validate_docx.py"
DEFAULT_RULES = ROOT / "config" / "rules.json"


def run(args: list[str]) -> int:
    return subprocess.run(args).returncode


def load_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def validate(document: Path, rules: Path, report: Path, strict: bool = False):
    cmd = [
        sys.executable, str(VALIDATOR), str(document),
        "--rules", str(rules), "--report", str(report),
    ]
    if strict:
        cmd.append("--strict")
    code = run(cmd)
    return code, load_json(report)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Validate a DOCX first, then apply only targeted repairs."
    )
    parser.add_argument("document", type=Path)
    parser.add_argument("--rules", type=Path, default=DEFAULT_RULES)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--strict", action="store_true")
    args = parser.parse_args()

    rules = load_json(args.rules)
    max_repairs = int(rules["workflow"]["max_auto_repairs"])
    report = args.report or args.document.with_name(
        args.document.stem + ".validation.json"
    )

    # Quick validation comes first. A clean document is never rewritten.
    code, result = validate(args.document, args.rules, report, strict=False)
    repairs = 0

    while code != 0 or not result.get("ok"):
        if repairs >= max_repairs or not result.get("auto_fixable"):
            codes = sorted({
                item.get("code")
                for item in result.get("issues", [])
                if item.get("code")
            })
            print("STOP: automatic repair limit reached or no auto-fixable issues remain.")
            if codes:
                print("Remaining validator codes:", ", ".join(codes))
            print("Report:", report)
            return 1

        repairs += 1
        print(f"Targeted repair cycle {repairs}/{max_repairs}")
        if run([
            sys.executable, str(FORMATTER), str(args.document),
            "--rules", str(args.rules),
            "--fix", str(report),
            "--in-place",
        ]) != 0:
            return 2

        code, result = validate(args.document, args.rules, report, strict=False)

    if args.strict:
        code, result = validate(args.document, args.rules, report, strict=True)
        if code != 0 or not result.get("ok"):
            codes = sorted({
                item.get("code")
                for item in result.get("issues", [])
                if item.get("code")
            })
            print("FAIL [strict]")
            if codes:
                print("Strict validator codes:", ", ".join(codes))
            print("Report:", report)
            return 1

    if repairs:
        print(f"PASS after {repairs} targeted repair cycle(s)")
    else:
        print("PASS without rewriting the document")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
