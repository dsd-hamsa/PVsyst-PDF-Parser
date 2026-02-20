#!/usr/bin/env python3
"""Helpers for using the PVsyst parser as a library.

This module intentionally avoids writing any output files. It exposes a small
function you can import from other scripts.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict

from pvsyst_parser import PVsystParser


def _parse_pdf_no_outputs(pdf_path: str, *, interactive: bool = False) -> PVsystParser:
    """Run PVsystParser parse steps without writing any files."""

    p = Path(pdf_path)
    if not p.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    parser = PVsystParser()

    blocks = parser.extract_text_blocks(str(p))
    parser.sections = parser.identify_sections(blocks)
    parser.section_contents = parser.extract_section_contents(blocks, parser.sections)
    parser.total_inverters_from_power_section = parser._parse_total_inverter_power()

    parser.extract_equipment_info(blocks)
    parser.orientations = parser.extract_orientations(blocks)

    if (
        "Array Losses" in parser.section_contents
        and parser.section_contents["Array Losses"]
    ):
        try:
            parser.array_losses = parser.parse_array_losses_section(
                parser.section_contents["Array Losses"][0]
            )
        except Exception as exc:  # noqa: BLE001
            # Match pvsyst_parser.py behavior: warn and keep going.
            print(f"  Warning: failed to parse array losses: {exc}")

    parser.arrays = parser.parse_arrays_from_text(blocks, interactive=interactive)
    parser.inverter_types = parser._collect_inverter_types()
    parser.calculate_monthly_production(blocks)

    # Ensure summaries are built.
    parser.to_dict()
    return parser


def parse_pdf_to_dict(pdf_path: str, *, interactive: bool = False) -> Dict[str, Any]:
    """Parse a PVsyst PDF and return the structured dict without writing files."""

    parser = _parse_pdf_no_outputs(pdf_path, interactive=interactive)
    return parser.to_dict()


def parse_pdf_to_powertrack_patch_dict(
    pdf_path: str, *, interactive: bool = False
) -> Dict[str, Any]:
    """Parse PVsyst PDF and return PowerTrack patch dict keyed as PV0/PV1/...."""

    parser = _parse_pdf_no_outputs(pdf_path, interactive=interactive)
    return parser.to_powertrack_patches_by_inverter(omit_nulls=True)


def _main() -> None:
    import argparse

    ap = argparse.ArgumentParser(
        description="Parse PVsyst PDF -> JSON to stdout (no files)"
    )
    ap.add_argument("pdf_file", help="Path to PVsyst PDF")
    ap.add_argument(
        "--powertrack-patch",
        action="store_true",
        help="Output PowerTrack patch JSON (keys PV0, PV1, ...) instead of full parser JSON",
    )
    ap.add_argument(
        "--interactive",
        action="store_true",
        help="Prompt to override array inverter/MPPT parsing",
    )
    args = ap.parse_args()

    if args.powertrack_patch:
        data = parse_pdf_to_powertrack_patch_dict(
            args.pdf_file, interactive=args.interactive
        )
    else:
        data = parse_pdf_to_dict(args.pdf_file, interactive=args.interactive)

    print(json.dumps(data, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    _main()
