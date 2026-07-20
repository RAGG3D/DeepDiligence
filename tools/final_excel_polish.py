#!/usr/bin/env python3
"""Run Excel-native final formatting that is unsafe to emulate by style IDs."""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import time
from pathlib import Path


def _default_path(ticker: str) -> Path:
    return Path(f"/mnt/c/Users/yzsun/Desktop/DD/{ticker}/DCF {ticker}.xlsx")


def _to_windows_path(path: Path) -> str:
    text = str(path)
    match = re.match(r"^/mnt/([a-zA-Z])/(.*)$", text)
    if match:
        drive = match.group(1).upper()
        rest = match.group(2).replace("/", "\\")
        return f"{drive}:\\{rest}"
    return text


def main() -> None:
    parser = argparse.ArgumentParser(description="Run Excel COM final polish")
    parser.add_argument("--ticker")
    parser.add_argument("--path")
    parser.add_argument("--no-backup", action="store_true")
    args = parser.parse_args()

    path = Path(args.path) if args.path else _default_path(args.ticker)
    if not path.exists():
        raise FileNotFoundError(path)
    if not args.no_backup:
        backup = path.with_name(f"{path.stem}_pre_final_polish_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
        shutil.copy2(path, backup)
        print(f"Backup: {backup}")

    script = Path(__file__).with_name("final_excel_polish.ps1")
    subprocess.run(
        [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script),
            "-Path",
            _to_windows_path(path),
        ],
        check=True,
        timeout=600,
    )


if __name__ == "__main__":
    main()
