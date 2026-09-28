#!/usr/bin/env python3
"""Run the complete Word formatting loop with a bounded repair count."""

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
    proc = subprocess.run(args)
    return proc.returncode


def load_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def main() -> int:
    parser = argparse.ArgumentParser(description="Format, validate, and auto-repair a DOCX.")
    parser.add_argument("document", type=Path)
    parser.add_argument("--rules", type=Path, default=DEFAULT_RULES)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--strict", action="store_true")
    args = parser.parse_args()

    rules = load_json(args.rules)
    max_repairs = int(rules["workflow"]["max_auto_repairs"])
    report = args.report or args.document.with_name(args.document.stem + ".validation.json")

    if run([
        sys.executable, str(FORMATTER), str(args.document),
        "--rules", str(args.rules), "--in-place",
    ]) != 0:
        return 2

    repairs = 0
    while True:
        cmd = [
            sys.executable, str(VALIDATOR), str(args.document),
            "--rules", str(args.rules), "--report", str(report),
        ]
        if args.strict:
            cmd.append("--strict")

        code = run(cmd)
        result = load_json(report)

        if code == 0 and result.get("ok"):
            print(f"PASS after {repairs} repair cycle(s)")
            return 0

        if repairs >= max_repairs or not result.get("auto_fixable"):
            codes = sorted({item.get("code") for item in result.get("issues", []) if item.get("code")})
            print("STOP: automatic repair limit reached or no auto-fixable issues remain.")
            if codes:
                print("Remaining validator codes:", ", ".join(codes))
            print("Report:", report)
            return 1

        repairs += 1
        print(f"Auto-repair cycle {repairs}/{max_repairs}")
        if run([
            sys.executable, str(FORMATTER), str(args.document),
            "--rules", str(args.rules), "--fix", str(report), "--in-place",
        ]) != 0:
            return 2


if __name__ == "__main__":
    raise SystemExit(main())
