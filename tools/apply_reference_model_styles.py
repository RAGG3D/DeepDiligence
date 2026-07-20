#!/usr/bin/env python3
"""Apply the approved MOLN visual system to every delivered model tab.

The Excel-native worker copies formats only.  This wrapper preserves an active
Catalyst run by temporarily restoring the neutral fill/font appearance and
reapplying the same run after the style rebuild without creating another
database record.
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

REPO = Path(__file__).resolve().parents[1]
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

from tools import catalyst_workflow as lifecycle  # noqa: E402

DEFAULT_REFERENCE = Path("/mnt/c/Users/yzsun/Desktop/DD/MOLN/DCF MOLN.xlsx")
DEFAULT_PIPELINE_REFERENCE = Path(
    "/mnt/c/Users/yzsun/Desktop/DD/MOLN/"
    "DCF MOLN completed_pre_calcstate_20260709_035047.xlsx"
)


def windows_path(path: Path) -> str:
    result = subprocess.run(
        ["wslpath", "-w", str(path)], check=True, capture_output=True, text=True
    )
    return result.stdout.strip()


def restore_active_fills(path: Path, state: dict) -> None:
    original = state.get("original_styles") or {}
    backup = Path(state.get("backup") or "")
    if not original or not backup.exists():
        return
    parts, order = lifecycle._read_parts(path)
    sheet_path = lifecycle._sheet_path(parts, "Catalyst")
    sheet = ET.fromstring(parts[sheet_path])
    lifecycle._restore_original_fills(parts, sheet, original, backup)
    parts[sheet_path] = lifecycle._excel_xml(sheet)
    lifecycle._write_parts(path, parts, order)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--ticker", required=True)
    parser.add_argument("--path")
    parser.add_argument("--reference", default=str(DEFAULT_REFERENCE))
    parser.add_argument(
        "--pipeline-reference",
        default=str(DEFAULT_PIPELINE_REFERENCE),
        help="Approved Pipeline-only style source; all other tabs use --reference",
    )
    parser.add_argument("--no-backup", action="store_true")
    parser.add_argument(
        "--pipeline-only",
        action="store_true",
        help="Apply only the locked Pipeline style source",
    )
    parser.add_argument(
        "--scenarios-only",
        action="store_true",
        help="Apply only semantic Scenarios row styles from --reference",
    )
    args = parser.parse_args()
    if args.pipeline_only and args.scenarios_only:
        raise SystemExit("--pipeline-only and --scenarios-only are mutually exclusive")

    ticker = args.ticker.upper()
    path = Path(args.path) if args.path else Path(
        f"/mnt/c/Users/yzsun/Desktop/DD/{ticker}/DCF {ticker}.xlsx"
    )
    reference = Path(args.reference)
    pipeline_reference = Path(args.pipeline_reference)
    if args.pipeline_only:
        required_paths = (path, pipeline_reference)
    elif args.scenarios_only:
        required_paths = (path, reference)
    else:
        required_paths = (path, reference, pipeline_reference)
    for required in required_paths:
        if not required.exists():
            raise FileNotFoundError(required)
    if not args.no_backup:
        backup = path.with_name(
            f"{path.stem}_pre_reference_styles_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        shutil.copy2(path, backup)
        print(f"Backup: {backup}")

    state_path = REPO / "artifacts" / ticker / f"{ticker}_catalyst_active_state.json"
    state = json.loads(state_path.read_text(encoding="utf-8")) if state_path.exists() else None
    if state and not (args.pipeline_only or args.scenarios_only):
        restore_active_fills(path, state)

    command = [
            "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
            "-File", windows_path(REPO / "tools" / "apply_moln_reference_styles.ps1"),
            "-Path", windows_path(path),
            "-ReferencePath", windows_path(reference),
            "-PipelineReferencePath", windows_path(pipeline_reference),
        ]
    if args.pipeline_only:
        command.append("-PipelineOnly")
    if args.scenarios_only:
        command.append("-ScenariosOnly")
    subprocess.run(
        command,
        check=True,
        timeout=1200,
    )

    if state and not (args.pipeline_only or args.scenarios_only):
        manifest_path = REPO / "artifacts" / ticker / f"{ticker}_catalyst_manifest.json"
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        neutral_backup = path.with_name(
            f"{path.stem}_pre_active_reference_styles_"
            f"{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        )
        shutil.copy2(path, neutral_backup)
        state["original_styles"] = lifecycle._patch_run(
            path, manifest, state.get("research") or {}
        )
        state["backup"] = str(neutral_backup)
        state["workbook"] = str(path)
        state["reference_style_workbook"] = str(reference)
        state["pipeline_reference_style_workbook"] = str(pipeline_reference)
        state_path.write_text(
            json.dumps(state, indent=2, ensure_ascii=False), encoding="utf-8"
        )
        print(f"Reapplied active Catalyst run {state.get('run_id')} after style normalization")

    if args.pipeline_only:
        print(f"Pipeline reference styles applied: {pipeline_reference} -> {path}")
    elif args.scenarios_only:
        print(f"Scenarios reference styles applied: {reference} -> {path}")
    else:
        print(
            f"Reference styles applied: {reference} -> {path}; "
            f"Pipeline only: {pipeline_reference}"
        )


if __name__ == "__main__":
    main()
