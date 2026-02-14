#!/usr/bin/env python3
"""PVsyst PDF Parser V3

This version merges:
  robust equipment parsing, inverter-type propagation,
  array-loss parsing, and per-MPPT string/module/DC-kWp allocation
  and stronger inverter-range parsing.

Key V3 behavior choices:
- Preserves MPPT numbers from PVsyst headers when present.
- Assigns MPPT numbers only when missing/unknown (mppt=None).
- Table extraction is disabled for speed.

Outputs:
- Writes a structured JSON file and a human-readable text report.
- `to_dict()` returns the same structure as the JSON output.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import defaultdict
from pathlib import Path
from re import Match
from typing import Any, Dict, List, Optional, Tuple

import pdfplumber


# -----------------------------------------------------------------------------
# Helper functions
# -----------------------------------------------------------------------------


def clean_power_to_kw_or_w(power_str: str) -> Optional[float]:
    """Extract numeric portion and interpret MW/kW when present.

    Returns:
      - kW for strings containing 'kW' or 'MW' (MW converted to kW)
      - W for strings without kW/MW
    """
    if not power_str:
        return None
    s = power_str.strip().lower()
    m = re.search(r"([0-9]*\.?[0-9]+)", s)
    if not m:
        return None
    value = float(m.group(1))
    if "mw" in s:
        return value * 1000.0
    if "kw" in s:
        return value
    return value


class PVsystParser:
    """Comprehensive parser for PVsyst PDF reports."""

    def __init__(self) -> None:
        self.sections: Dict[str, Any] = {}
        self.section_contents: Dict[str, List[str]] = {}

        self.arrays: Dict[str, Dict[str, Any]] = {}
        self.expanded_arrays: List[Dict[str, Any]] = []

        self.module_info: Dict[str, Any] = {}
        self.inverter_info: Dict[str, Any] = {}

        self.orientations: Dict[str, Dict[str, Any]] = {}

        self.system_monthly_production: Dict[str, float] = {}
        self.system_monthly_globhor: Dict[str, float] = {}
        self.monthly_production: Dict[str, Dict[str, float]] = {}

        self.inverter_capacities: Dict[str, float] = {}
        self.associations: Dict[str, Any] = {}
        self.inverter_summary: Dict[str, Any] = {}

        self.array_losses: Dict[str, Any] = {}
        self.inverter_types: List[Dict[str, Any]] = []

    # -------------------------------------------------------------------------
    # Text extraction
    # -------------------------------------------------------------------------

    def extract_text_blocks(self, pdf_path: str) -> Dict[int, Dict[str, Any]]:
        """Extract text blocks and key-value pairs from PDF."""
        blocks: Dict[int, Dict[str, Any]] = {}

        print("  Extracting text with pdfplumber...")
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                txt = page.extract_text() or ""
                lines = [ln for ln in txt.splitlines() if ln.strip()]

                kv_pairs: List[Dict[str, str]] = []
                others: List[str] = []
                for ln in lines:
                    if ":" in ln and not ln.strip().startswith(":"):
                        k, v = ln.split(":", 1)
                        if k.strip():
                            kv_pairs.append({"key": k.strip(), "value": v.strip()})
                            continue
                    others.append(ln.strip())

                blocks[i] = {"kv": kv_pairs, "text_lines": others, "full_text": txt}

        return blocks

    # -------------------------------------------------------------------------
    # Section identification
    # -------------------------------------------------------------------------

    def identify_sections(
        self, blocks: Dict[int, Dict[str, Any]]
    ) -> Dict[str, Dict[str, Any]]:
        """Identify high-level sections in the document."""
        print("  Identifying sections...")

        all_text = "\n".join(
            blocks[p].get("full_text") or "" for p in sorted(blocks.keys())
        )

        # Union of patterns from both versions.
        section_patterns = {
            "Project Summary": r"Project summary|System summary|Results summary",
            "PV Array Characteristics": r"PV Array Characteristics|Array Characteristics|PV Modules|Module Configuration",
            "System Losses": r"System losses|Loss diagram",
            "Array Losses": r"Array losses",
            "Horizon Definition": r"Horizon definition",
            "Near Shading": r"Near shading|Iso-shadings diagram",
            "Main Results": r"Main results",
            "Predefined Graphs": r"Predef\.? graphs",
            "P50-P90 Evaluation": r"P50.*P90 evaluation",
        }

        sections: Dict[str, Dict[str, Any]] = {}
        for section_name, pattern in section_patterns.items():
            matches = list(re.finditer(pattern, all_text, re.IGNORECASE))
            if matches:
                sections[section_name] = {
                    "start_positions": [m.start() for m in matches],
                    "matches": [m.group() for m in matches],
                }

        return sections

    def extract_section_contents(
        self, blocks: Dict[int, Dict[str, Any]], sections: Dict[str, Dict[str, Any]]
    ) -> Dict[str, List[str]]:
        """Extract full text content for each identified section."""
        all_text = "\n".join(
            blocks[p].get("full_text") or "" for p in sorted(blocks.keys())
        )

        all_starts: List[Tuple[int, str]] = []
        for sec_name, sec_data in sections.items():
            for pos in sec_data.get("start_positions", []):
                all_starts.append((int(pos), sec_name))
        all_starts.sort(key=lambda x: x[0])

        section_contents: Dict[str, List[str]] = {}
        for i, (pos, sec_name) in enumerate(all_starts):
            start = pos
            end = all_starts[i + 1][0] if i + 1 < len(all_starts) else len(all_text)
            content = all_text[start:end].strip()
            section_contents.setdefault(sec_name, []).append(content)

        return section_contents

    # -------------------------------------------------------------------------
    # Equipment parsing
    # -------------------------------------------------------------------------

    @staticmethod
    def _two_column_values(
        line: str, label: str
    ) -> Tuple[Optional[str], Optional[str]]:
        """Parse PVsyst two-column rows with repeated labels or wide spacing."""
        if not line or not label:
            return (None, None)

        pat_two = re.compile(
            rf"{re.escape(label)}\s+(.+?)\s+{re.escape(label)}\s+(.+)$",
            re.IGNORECASE,
        )
        m = pat_two.search(line)
        if m:
            return (m.group(1).strip() or None, m.group(2).strip() or None)

        pat_one = re.compile(rf"{re.escape(label)}\s+(.+)$", re.IGNORECASE)
        m = pat_one.search(line)
        if not m:
            return (None, None)

        remainder = m.group(1).strip()
        if not remainder:
            return (None, None)

        parts = re.split(r"\s{2,}", remainder)
        if len(parts) >= 2:
            return (parts[0].strip() or None, parts[1].strip() or None)

        return (remainder, None)

    @staticmethod
    def _second_column_value(line: str, label: str) -> Optional[str]:
        left, right = PVsystParser._two_column_values(line, label)
        return right or left

    def clean_nom_power(self, power_str: str) -> Optional[float]:
        """Parse nominal power strings; returns kW if 'kW/MW' else W."""
        if not power_str:
            return None

        s = power_str.strip().lower()
        m = re.search(r"([0-9]*\.?[0-9]+)", s)
        if not m:
            return None

        value = float(m.group(1))
        if "mw" in s:
            return value * 1000.0
        if "kw" in s:
            return value
        return value

    def extract_equipment_info(
        self, blocks: Dict[int, Dict[str, Any]]
    ) -> Tuple[Dict[str, Any], Dict[str, Any]]:
        """Extract global PV module + inverter equipment info."""
        module_info: Dict[str, Any] = {}
        inverter_info: Dict[str, Any] = {}

        all_text = "\n".join(
            blocks[p].get("full_text") or "" for p in sorted(blocks.keys())
        )
        m = re.search(
            r"\bPV\s+module\b(.{0,2200})", all_text, re.IGNORECASE | re.DOTALL
        )
        if not m:
            self.module_info = module_info
            self.inverter_info = inverter_info
            return module_info, inverter_info

        block = "PV module\n" + m.group(1)
        lines = [ln.strip() for ln in block.splitlines() if ln.strip()]

        manu_line = next(
            (ln for ln in lines if re.search(r"\bManufacturer\b", ln, re.IGNORECASE)),
            None,
        )
        if manu_line:
            left, right = self._two_column_values(manu_line, "Manufacturer")
            if left:
                module_info["manufacturer"] = left
            if right:
                inverter_info["manufacturer"] = right

        model_line = next(
            (ln for ln in lines if re.search(r"\bModel\b", ln, re.IGNORECASE)), None
        )
        if model_line:
            left, right = self._two_column_values(model_line, "Model")
            if left:
                module_info["model"] = left
            if right:
                inverter_info["model"] = right

        power_line = next(
            (
                ln
                for ln in lines
                if re.search(r"Unit\s+Nom\.?\s*Power", ln, re.IGNORECASE)
            ),
            None,
        )
        if power_line:
            left, right = self._two_column_values(power_line, "Unit Nom. Power")
            if left is None and right is None:
                left, right = self._two_column_values(power_line, "Unit Nom Power")

            if left:
                module_info["unit_nom_power_raw"] = left
                numeric = clean_power_to_kw_or_w(left)
                if numeric is not None:
                    lower_left = left.lower()
                    if "mw" in lower_left:
                        module_info["unit_nom_power_w"] = int(
                            round(numeric * 1_000_000)
                        )
                    elif "kw" in lower_left:
                        module_info["unit_nom_power_w"] = int(round(numeric * 1_000))
                    else:
                        module_info["unit_nom_power_w"] = int(round(numeric))

            if right:
                inverter_info["unit_nom_power_raw"] = right
                numeric = clean_power_to_kw_or_w(right)
                if numeric is not None:
                    inverter_info["unit_nom_power_kw"] = numeric

        self.module_info = module_info
        self.inverter_info = inverter_info
        return module_info, inverter_info

    # -------------------------------------------------------------------------
    # Orientation
    # -------------------------------------------------------------------------

    @staticmethod
    def pvsyst_azimuth_to_compass(az_pvsyst: float) -> float:
        return (180.0 + az_pvsyst) % 360.0

    def extract_orientations(
        self, blocks: Dict[int, Dict[str, Any]]
    ) -> Dict[str, Dict[str, Any]]:
        """Extract Orientation #n entries and associate tilt/azimuth."""
        print("  Extracting orientations...")

        all_text = "\n".join(
            blocks[p].get("full_text") or "" for p in sorted(blocks.keys())
        )

        orientations: Dict[str, Dict[str, Any]] = {}

        ori_matches = list(
            re.finditer(r"Orientation\s*#?\s*(\d+)", all_text, re.IGNORECASE)
        )
        tilt_matches = list(
            re.finditer(
                r"Tilt\s*[/]?\s*Azimuth\s*([-\d.]+)\s*[/]\s*([-\d.]+)°?",
                all_text,
                re.IGNORECASE,
            )
        )

        for ori_m in ori_matches:
            ori_id = ori_m.group(1)
            ori_pos = ori_m.start()

            closest_tilt: Optional[Match[str]] = None
            min_dist = float("inf")
            for tilt_m in tilt_matches:
                dist = abs(tilt_m.start() - ori_pos)
                if dist < min_dist:
                    min_dist = dist
                    closest_tilt = tilt_m

            if closest_tilt:
                tilt = float(closest_tilt.group(1))
                az_pv = float(closest_tilt.group(2))
                az_compass = self.pvsyst_azimuth_to_compass(az_pv)
                orientations[ori_id] = {
                    "tilt": tilt,
                    "azimuth_pvsyst_deg": az_pv,
                    "azimuth_deg": az_compass,
                    "azimuth_compass_deg": az_compass,
                }

        # Local fallback for any missing
        for m in ori_matches:
            ori_id = m.group(1)
            if ori_id in orientations:
                continue
            window = all_text[m.start() : m.start() + 800]
            tilt_match = re.search(
                r"Tilt\s*[/]?\s*Azimuth\s*([-\d.]+)\s*[/]\s*([-\d.]+)°?",
                window,
                re.IGNORECASE,
            )
            if tilt_match:
                tilt = float(tilt_match.group(1))
                az_pv = float(tilt_match.group(2))
                az_compass = self.pvsyst_azimuth_to_compass(az_pv)
                orientations[ori_id] = {
                    "tilt": tilt,
                    "azimuth_pvsyst_deg": az_pv,
                    "azimuth_deg": az_compass,
                    "azimuth_compass_deg": az_compass,
                }

        print(f"    Found {len(orientations)} orientations")
        return orientations

    # -------------------------------------------------------------------------
    # Inverter/MPPT parsing helpers
    # -------------------------------------------------------------------------

    def parse_inverter_range(self, inv_text: str) -> List[str]:
        """Parse complex inverter notation into individual inverter names.

        Examples:
          - "INV01" -> ["INV01"]
          - "INV02-05" -> ["INV02", "INV03", "INV04", "INV05"]
          - "INV02-05, 7,8" -> ["INV02", "INV03", "INV04", "INV05", "INV07", "INV08"]
          - "INV R1-3" -> ["INVR01", "INVR02", "INVR03"]
        """
        inverters: List[str] = []

        inv_text = (inv_text or "").strip()
        parts = [p.strip() for p in inv_text.split(",") if p.strip()]

        for part in parts:
            if not part.upper().startswith("INV"):
                part = "INV " + part

            m_range = re.search(
                r"INV\s*([A-Za-z]*)(\d+)\s*-\s*([A-Za-z]*)(\d+)",
                part,
                re.IGNORECASE,
            )
            if m_range:
                p1, start, p2, end = (
                    m_range.group(1),
                    int(m_range.group(2)),
                    m_range.group(3),
                    int(m_range.group(4)),
                )
                p2 = p2 or p1
                for i in range(start, end + 1):
                    inverters.append(f"INV{p1}{i:02d}")
                continue

            m_single = re.search(r"INV\s*([A-Za-z]*)(\d+)", part, re.IGNORECASE)
            if m_single:
                prefix = m_single.group(1)
                num = int(m_single.group(2))
                inverters.append(f"INV{prefix}{num:02d}")
                continue

        return inverters

    def parse_mppt_range(self, mppt_text: str) -> List[str]:
        mppt_text = (mppt_text or "").strip()
        mppt_text = re.sub(r"^MPPT\s*", "", mppt_text, flags=re.IGNORECASE)
        parts = [p.strip() for p in mppt_text.split(",") if p.strip()]

        mppts: List[str] = []
        for part in parts:
            if "-" in part:
                m = re.search(r"(\d+)\s*-\s*(\d+)", part)
                if m:
                    start = int(m.group(1))
                    end = int(m.group(2))
                    for i in range(start, end + 1):
                        mppts.append(f"MPPT {i}")
            else:
                m = re.search(r"(\d+)", part)
                if m:
                    mppts.append(f"MPPT {int(m.group(1))}")

        return mppts

    # -------------------------------------------------------------------------
    # Array parsing
    # -------------------------------------------------------------------------

    def expand_array_notation(self, array_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        array_id = array_data.get("array_id")
        inverter_ids = array_data.get("inverter_ids") or []
        mppt_ids = array_data.get("mppt_ids")

        if not mppt_ids:
            mppt_count = array_data.get("mppt_count")
            if isinstance(mppt_count, int) and mppt_count > 0:
                mppt_ids = [f"MPPT {i}" for i in range(1, mppt_count + 1)]

        combos: List[Dict[str, Any]] = []
        if not inverter_ids:
            return combos

        original_notation = array_data.get("original_notation", "")

        if mppt_ids:
            for inv in inverter_ids:
                for mppt in mppt_ids:
                    combos.append(
                        {
                            "array_id": array_id,
                            "inverter": inv,
                            "mppt": mppt,
                            "original_notation": original_notation,
                        }
                    )
        else:
            for inv in inverter_ids:
                combos.append(
                    {
                        "array_id": array_id,
                        "inverter": inv,
                        "mppt": None,
                        "original_notation": original_notation,
                    }
                )

        return combos

    def _parse_pvsyst_inverter_type_block(self, text: str) -> Dict[str, Any]:
        """Parse a PVsyst equipment block between arrays; return inverter fields only."""
        if not text:
            return {}

        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        if not lines:
            return {}

        inv_idx: Optional[int] = None
        for i, ln in enumerate(lines):
            if re.fullmatch(r"Inverter", ln, flags=re.IGNORECASE):
                inv_idx = i
                break
        if inv_idx is None:
            for i, ln in enumerate(lines):
                if re.search(r"\bInverter\b", ln, flags=re.IGNORECASE):
                    inv_idx = i
                    break
        if inv_idx is None:
            return {}

        inv_lines = lines[inv_idx:]
        out: Dict[str, Any] = {}

        manu_line = next(
            (
                ln
                for ln in inv_lines
                if re.search(r"\bManufacturer\b", ln, re.IGNORECASE)
            ),
            None,
        )
        if manu_line:
            v = self._second_column_value(manu_line, "Manufacturer")
            if v:
                out["inverter_manufacturer"] = v

        model_line = next(
            (ln for ln in inv_lines if re.search(r"\bModel\b", ln, re.IGNORECASE)), None
        )
        if model_line:
            v = self._second_column_value(model_line, "Model")
            if v:
                out["inverter_model"] = v

        power_line = next(
            (
                ln
                for ln in inv_lines
                if re.search(r"Unit\s+Nom\.?\s*Power", ln, re.IGNORECASE)
            ),
            None,
        )
        if power_line:
            v = self._second_column_value(power_line, r"Unit\s+Nom\.?\s*Power")
            if v:
                out["inverter_unit_nom_power_raw"] = v
                kw = self.clean_nom_power(v)
                if kw is not None:
                    out["inverter_unit_nom_power_kw"] = kw

        return out

    def _parse_array_block(self, section_text: str, array_id: str) -> Dict[str, Any]:
        array_data: Dict[str, Any] = {
            "array_id": array_id,
            "original_block_text": section_text,
            "original_notation": f"Array #{array_id}",
        }

        header_line = section_text.splitlines()[0] if section_text.splitlines() else ""

        inverter_ids: List[str] = []

        # Prefer INV ... MPPT header notation (better for complex ranges)
        m_inv_mppt = re.search(r"INV\s+(.+?)\s+MPPT", header_line, re.IGNORECASE)
        if m_inv_mppt:
            inv_spec = m_inv_mppt.group(1).strip()
            inverter_ids = self.parse_inverter_range(f"INV {inv_spec}")

        # Fallback: find first INV token in the header
        if not inverter_ids:
            m_inv_simple = re.search(
                r"INV\s*([A-Za-z]*)(\d+)", header_line, re.IGNORECASE
            )
            if m_inv_simple:
                prefix = m_inv_simple.group(1)
                num = int(m_inv_simple.group(2))
                inverter_ids = [f"INV{prefix}{num:02d}"]

        if inverter_ids:
            array_data["inverter_ids"] = inverter_ids
            array_data["inverter_id"] = inverter_ids[0]

        # MPPT IDs from header, if present
        m_mppt_header = re.search(
            r"MPPT[#\s]*([0-9,\-\s]+)", header_line, re.IGNORECASE
        )
        if m_mppt_header:
            mppt_ids = self.parse_mppt_range(m_mppt_header.group(1))
            if mppt_ids:
                array_data["mppt_ids"] = mppt_ids

        # MPPT info from PVsyst format: "Number of inverters X * MPPT Y% Z unit"
        # PVsyst uses this line to describe how many inverter units / MPPT inputs are used.
        # In reports where the header expands to multiple inverters (e.g. INV01-03), the
        # first number is commonly the total MPPT endpoints across all listed inverters.
        m_mppt = re.search(
            r"Number of inverters\s*(\d+)\s*\*\s*MPPT\s*([\d.]+)%\s*([\d.]+)\s*unit",
            section_text,
            re.IGNORECASE,
        )
        if m_mppt:
            total_mppts = int(m_mppt.group(1))
            num_invs = len(inverter_ids) if inverter_ids else 1
            mppt_per_inv = max(1, total_mppts // max(1, num_invs))
            array_data["mppt_total_endpoints"] = total_mppts
            array_data["mppt_count"] = mppt_per_inv
            array_data["mppt_share_percent"] = float(m_mppt.group(2))
            array_data["inverter_unit_fraction"] = float(m_mppt.group(3))

        # Orientation #n inside the block
        m_ori = re.search(r"Orientation\s*#?\s*(\d+)", section_text, re.IGNORECASE)
        if m_ori:
            array_data["orientation_id"] = int(m_ori.group(1))

        # Number of PV modules
        m_mods = re.search(
            r"Number of PV modules\s*(\d+)units?", section_text, re.IGNORECASE
        )
        if m_mods:
            array_data["number_of_modules"] = int(m_mods.group(1))

        unit_wp = self.module_info.get("unit_nom_power_w")
        if isinstance(unit_wp, int) and "number_of_modules" in array_data:
            nominal_kwp_from_module = unit_wp * array_data["number_of_modules"] / 1000.0
            array_data["nominal_stc_kwp_from_module"] = round(
                nominal_kwp_from_module, 3
            )

        m_stc = re.search(
            r"Nominal\s*\(STC\)\s*([\d.]+)kWp", section_text, re.IGNORECASE
        )
        if m_stc:
            array_data["nominal_stc_kwp"] = float(m_stc.group(1))

        # Modules configuration
        m_cfg = re.search(
            r"Modules\s*(\d+)\s*string[s]?\s*x\s*(\d+)", section_text, re.IGNORECASE
        )
        if m_cfg:
            strings = int(m_cfg.group(1))
            series = int(m_cfg.group(2))
            array_data["strings"] = strings
            array_data["modules_in_series"] = series
            array_data["modules_config_text"] = f"Modules {strings} string x {series}"

        # Tilt/Azimuth
        m_tilt_az = re.search(
            r"Tilt/Azimuth\s*([-\d.]+)\s*/\s*([-\d.]+)\s*°", section_text, re.IGNORECASE
        )
        if m_tilt_az:
            tilt = float(m_tilt_az.group(1))
            az_pv = float(m_tilt_az.group(2))
            az_compass = self.pvsyst_azimuth_to_compass(az_pv)
            array_data["tilt"] = tilt
            array_data["azimuth_pvsyst_deg"] = az_pv
            array_data["azimuth_deg"] = az_compass
            array_data["azimuth_compass_deg"] = az_compass

        # U mpp / I mpp
        m_umpp = re.search(r"U mpp\s*([\d.]+)V", section_text, re.IGNORECASE)
        if m_umpp:
            array_data["u_mpp_v"] = float(m_umpp.group(1))
        m_impp = re.search(r"I mpp\s*([\d.]+)A", section_text, re.IGNORECASE)
        if m_impp:
            array_data["i_mpp_a"] = float(m_impp.group(1))

        # Inverter details embedded in this array block (seen in some PVsyst exports)
        m_eq = re.search(
            r"\nPV\s*module\b.*", section_text, flags=re.IGNORECASE | re.DOTALL
        )
        if m_eq:
            array_data.update(
                self._parse_pvsyst_inverter_type_block(section_text[m_eq.start() :])
            )

        return array_data

    def _interactive_array_config(
        self,
        header_line: str,
        auto_inverter_ids: List[str],
        auto_mppt_ids: Optional[List[str]],
        strings: int,
    ) -> Tuple[Optional[List[str]], Optional[List[str]]]:
        """Prompt user to override inverter/MPPT parsing for an array."""
        print(f"\nArray header: {header_line}")
        print(f"Auto-parsed inverters: {auto_inverter_ids}")
        print(f"Auto-parsed MPPTs: {auto_mppt_ids or 'None'}")

        multi_inv = (
            input("Does this array apply to multiple inverters? (Y/N): ")
            .strip()
            .lower()
        )
        user_inverter_ids: Optional[List[str]] = None
        if multi_inv == "y":
            inv_input = input("Enter inverter IDs (comma-separated): ").strip()
            if inv_input:
                inv_list = [inv.strip() for inv in inv_input.split(",") if inv.strip()]
                valid = all(re.match(r"^[A-Za-z0-9\-_.]+$", inv) for inv in inv_list)
                if valid:
                    user_inverter_ids = inv_list
                else:
                    print("Invalid inverter names. Falling back to auto-parsed.")

        spec_mppt = (
            input("Does this array apply to specific MPPTs? (Y/N): ").strip().lower()
        )
        user_mppt_ids: Optional[List[str]] = None
        if spec_mppt == "y":
            mppt_input = input(
                "Enter MPPT numbers (comma-separated, e.g., 1,2,3): "
            ).strip()
            if mppt_input:
                try:
                    mppt_nums = [
                        int(num.strip()) for num in mppt_input.split(",") if num.strip()
                    ]
                    user_mppt_ids = [f"MPPT {num}" for num in mppt_nums]
                except ValueError:
                    print("Invalid MPPT numbers. Falling back to auto-parsed.")

        if user_mppt_ids and strings > 0:
            n = len(user_mppt_ids)
            base = strings // n
            remainder = strings % n
            print(
                f"Distribution: {strings} strings across {n} MPPTs -> {base} base, {remainder} extra"
            )

        return user_inverter_ids, user_mppt_ids

    def _assign_missing_mppt_labels(self) -> None:
        """Assign MPPT labels only for combinations where mppt is None."""
        by_inverter: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
        for combo in self.expanded_arrays:
            by_inverter[combo["inverter"]].append(combo)

        mppt_num_re = re.compile(r"^MPPT\s*(\d+)$", re.IGNORECASE)

        for inv, combos in by_inverter.items():
            used: set[int] = set()
            missing: List[Dict[str, Any]] = []

            for c in combos:
                mppt = c.get("mppt")
                if mppt is None:
                    missing.append(c)
                    continue
                m = mppt_num_re.match(str(mppt).strip())
                if m:
                    used.add(int(m.group(1)))

            if not missing:
                continue

            # Stable ordering of missing MPPTs
            def sort_key(c: Dict[str, Any]) -> Tuple[int, str]:
                try:
                    aid = int(c.get("array_id") or 0)
                except ValueError:
                    aid = 0
                return (aid, c.get("original_notation") or "")

            missing.sort(key=sort_key)

            next_num = 1
            for c in missing:
                while next_num in used:
                    next_num += 1
                c["mppt"] = f"MPPT {next_num}"
                used.add(next_num)
                next_num += 1

    def _infer_mppt_topology(self) -> Optional[Dict[str, Any]]:
        """Infer MPPT topology from inverter manufacturer/model.

        Returns a dict with:
          - mppt_per_inverter
          - strings_per_mppt_max
          - source
        """
        manufacturer = str(self.inverter_info.get("manufacturer") or "").lower()
        model = str(self.inverter_info.get("model") or "").lower()

        if "sma" in manufacturer and "core" in model:
            return {
                "mppt_per_inverter": 6,
                "strings_per_mppt_max": 2,
                "source": "SMA Core1 heuristic",
            }

        if (
            ("chint" in manufacturer)
            or ("cps" in manufacturer)
            or ("cps" in model)
            or ("chint" in model)
        ):
            return {
                "mppt_per_inverter": 3,
                "strings_per_mppt_max": 6,
                "source": "CPS/CHINT heuristic",
            }

        return None

    @staticmethod
    def _sort_inv_ids(inv_ids: List[str]) -> List[str]:
        def key(inv: str) -> Tuple[int, str]:
            m = re.match(r"^INV\D*(\d+)$", inv, re.IGNORECASE)
            if m:
                return (int(m.group(1)), inv)
            return (10**9, inv)

        return sorted(inv_ids, key=key)

    @staticmethod
    def _sort_mppt_ids(mppt_ids: List[str]) -> List[str]:
        def key(mppt: str) -> Tuple[int, str]:
            m = re.match(r"^MPPT\s*(\d+)$", mppt, re.IGNORECASE)
            if m:
                return (int(m.group(1)), mppt)
            return (10**9, mppt)

        return sorted(mppt_ids, key=key)

    def _allocate_strings_single_config(
        self,
        inverter_ids: List[str],
        mppt_ids: List[str],
        total_strings: int,
        strings_per_mppt_max: int,
    ) -> Dict[Tuple[str, str], int]:
        """Allocate strings in inverter order, filling 1 per MPPT then round-robin.

        For each inverter: add 1 string to each MPPT (in order), repeat until either
        strings are exhausted or the inverter's MPPTs reach strings_per_mppt_max.
        Then move to the next inverter.
        """
        alloc: Dict[Tuple[str, str], int] = {
            (inv, mppt): 0 for inv in inverter_ids for mppt in mppt_ids
        }
        remaining = int(total_strings)

        for inv in inverter_ids:
            while remaining > 0:
                progressed = False
                for mppt in mppt_ids:
                    if remaining <= 0:
                        break
                    key = (inv, mppt)
                    if alloc[key] < strings_per_mppt_max:
                        alloc[key] += 1
                        remaining -= 1
                        progressed = True
                if not progressed:
                    break  # this inverter saturated

        if remaining > 0:
            print(
                f"  Warning: {remaining} strings could not be allocated within inferred MPPT limits; "
                "distributing beyond per-MPPT max"
            )
            # Best-effort: distribute remaining across all endpoints round-robin without cap.
            all_endpoints = [(inv, mppt) for inv in inverter_ids for mppt in mppt_ids]
            idx = 0
            while remaining > 0 and all_endpoints:
                inv, mppt = all_endpoints[idx % len(all_endpoints)]
                alloc[(inv, mppt)] += 1
                remaining -= 1
                idx += 1

        return alloc

    def _parse_single_configuration(self, text: str) -> Optional[Dict[str, Any]]:
        """Fallback for sites with a single configuration and no 'Array #' blocks."""
        if not text:
            return None

        if not re.search(r"PV Array Characteristics", text, re.IGNORECASE):
            return None

        m_mods = re.search(
            r"Number of PV modules\s*(\d+)\s*units?", text, re.IGNORECASE
        )
        if not m_mods:
            m_mods = re.search(
                r"Nb\.\s*of\s*modules\s*(\d+)\s*units?", text, re.IGNORECASE
            )
        if not m_mods:
            return None

        m_inv = re.search(r"Number of inverters\s*(\d+)\s*units?", text, re.IGNORECASE)
        if not m_inv:
            m_inv = re.search(
                r"Nb\.\s*of\s*units\s*(\d+)\s*units?", text, re.IGNORECASE
            )
        if not m_inv:
            return None

        # Accept both "string(s)" and "Strings", tolerate "17In series".
        m_cfg = re.search(
            r"Modules\s*(\d+)\s*(?:string[s]?|Strings)\s*x\s*(\d+)\s*In\s*series",
            text,
            re.IGNORECASE,
        )
        if not m_cfg:
            return None

        strings = int(m_cfg.group(1))
        series = int(m_cfg.group(2))
        number_of_modules = int(m_mods.group(1))
        inverter_units_reported = int(m_inv.group(1))

        topology = self._infer_mppt_topology() or {
            "mppt_per_inverter": 1,
            "strings_per_mppt_max": max(1, strings),
            "source": "default",
        }
        mppt_per_inv = int(topology["mppt_per_inverter"])
        strings_per_mppt_max = int(topology["strings_per_mppt_max"])

        # If PVsyst doesn't provide explicit per-inverter stringing, prefer the
        # minimum inverter count that can host the total strings within MPPT limits.
        strings_per_inverter_max = max(1, mppt_per_inv * strings_per_mppt_max)
        inverter_units_required = (
            strings + strings_per_inverter_max - 1
        ) // strings_per_inverter_max

        inverter_units_used = inverter_units_reported
        if inverter_units_reported > inverter_units_required:
            inverter_units_used = inverter_units_required

        inverter_ids = [f"INV{i:02d}" for i in range(1, inverter_units_used + 1)]
        mppt_ids = [f"MPPT {i}" for i in range(1, mppt_per_inv + 1)]

        array_data: Dict[str, Any] = {
            "array_id": "1",
            "original_block_text": "PV Array Characteristics (single configuration)",
            "original_notation": "Single configuration",
            "strings": strings,
            "modules_in_series": series,
            "number_of_modules": number_of_modules,
            "inverter_ids": inverter_ids,
            "mppt_ids": mppt_ids,
            "inferred_single_config": True,
            "inferred_mppt_per_inverter": mppt_per_inv,
            "inferred_strings_per_mppt_max": strings_per_mppt_max,
            "inferred_topology_source": topology.get("source"),
            "inferred_inverters_reported": inverter_units_reported,
            "inferred_inverters_required": inverter_units_required,
            "inferred_inverters_used": inverter_units_used,
        }

        # Tilt/Azimuth if present
        m_tilt_az = re.search(
            r"Tilt/Azimuth\s*([-\d.]+)\s*/\s*([-\d.]+)\s*°", text, re.IGNORECASE
        )
        if m_tilt_az:
            tilt = float(m_tilt_az.group(1))
            az_pv = float(m_tilt_az.group(2))
            az_compass = self.pvsyst_azimuth_to_compass(az_pv)
            array_data["tilt"] = tilt
            array_data["azimuth_pvsyst_deg"] = az_pv
            array_data["azimuth_deg"] = az_compass
            array_data["azimuth_compass_deg"] = az_compass

        # If only one orientation exists, bind it
        if self.orientations and len(self.orientations) == 1:
            ori_id_str = next(iter(self.orientations.keys()))
            try:
                array_data["orientation_id"] = int(ori_id_str)
            except ValueError:
                array_data["orientation_id"] = ori_id_str

        unit_wp = self.module_info.get("unit_nom_power_w")
        if isinstance(unit_wp, int):
            array_data["nominal_stc_kwp_from_module"] = round(
                unit_wp * number_of_modules / 1000.0, 3
            )

        return array_data

    def parse_arrays_from_text(
        self, blocks: Dict[int, Dict[str, Any]], interactive: bool = False
    ) -> Dict[str, Dict[str, Any]]:
        print("  Parsing array data (generic)...")

        pages_with_arrays: List[int] = []
        for page_num, page_data in blocks.items():
            full = page_data.get("full_text", "") or ""
            if (
                re.search(r"PV Array Characteristics", full, re.IGNORECASE)
                or re.search(r"Array\s*#?\s*\d+", full, re.IGNORECASE)
                or re.search(r"Array Characteristics", full, re.IGNORECASE)
                or re.search(r"PV Modules", full, re.IGNORECASE)
                or re.search(r"Module Configuration", full, re.IGNORECASE)
            ):
                pages_with_arrays.append(int(page_num))

        if not pages_with_arrays:
            print("    No PV Array Characteristics / Array pages found.")
            return {}

        pages_with_arrays.sort()
        start_page = pages_with_arrays[0]
        end_page = pages_with_arrays[-1]

        combined_text = "\n".join(
            blocks[p].get("full_text") or "" for p in range(start_page, end_page + 1)
        )

        array_pattern = re.compile(
            r"(Array\s*#?\s*(\d+).*?)(?=Array\s*#?\s*\d+|AC wiring losses|Page \d+/\d+|$)",
            re.DOTALL | re.IGNORECASE,
        )

        arrays: Dict[str, Dict[str, Any]] = {}
        seen_ids: set[str] = set()
        pending_inverter_type: Dict[str, Any] = {}

        for match in array_pattern.finditer(combined_text):
            block_text = match.group(1)
            array_id = match.group(2)

            if array_id in seen_ids:
                continue

            if not re.search(
                r"Modules\s+\d+\s+(?:string|Strings)", block_text, re.IGNORECASE
            ):
                continue

            trailing_equipment = None
            m_eq = re.search(
                r"\nPV\s*module\b.*", block_text, flags=re.IGNORECASE | re.DOTALL
            )
            if m_eq:
                trailing_equipment = block_text[m_eq.start() :]
                block_text = block_text[: m_eq.start()].rstrip()

            array_data = self._parse_array_block(block_text, array_id)

            if pending_inverter_type and array_data.get("inverter_id"):
                if not array_data.get("inverter_model") and not array_data.get(
                    "inverter_manufacturer"
                ):
                    array_data.update(pending_inverter_type)

            if interactive:
                header_line = array_data.get("original_block_text", "").splitlines()[0]
                user_inv, user_mppt = self._interactive_array_config(
                    header_line,
                    array_data.get("inverter_ids", []),
                    array_data.get("mppt_ids"),
                    int(array_data.get("strings") or 0),
                )
                if user_inv is not None:
                    array_data["inverter_ids"] = user_inv
                if user_mppt is not None:
                    array_data["mppt_ids"] = user_mppt

            arrays[array_id] = array_data
            seen_ids.add(array_id)

            if trailing_equipment:
                parsed_type = self._parse_pvsyst_inverter_type_block(trailing_equipment)
                if parsed_type:
                    pending_inverter_type = parsed_type

        # Fallback: PVsyst reports with a single configuration and no Array # blocks
        if not arrays:
            single = self._parse_single_configuration(combined_text)
            if single:
                arrays[single["array_id"]] = single

        # Expand combinations
        self.expanded_arrays = []
        for arr in arrays.values():
            arr["expanded_combinations"] = self.expand_array_notation(arr)
            self.expanded_arrays.extend(arr["expanded_combinations"])

        # If we have any combos but some MPPTs are unknown, assign only missing
        if self.expanded_arrays:
            self._assign_missing_mppt_labels()

        # Back-fill orientation when only one orientation exists
        if self.orientations and len(self.orientations) == 1:
            ori_id_str, ori_data = next(iter(self.orientations.items()))
            try:
                ori_id_int: Any = int(ori_id_str)
            except ValueError:
                ori_id_int = ori_id_str

            for arr in arrays.values():
                if "orientation_id" not in arr:
                    arr["orientation_id"] = ori_id_int
                    if "tilt" in ori_data:
                        arr["tilt"] = ori_data["tilt"]
                    if "azimuth_pvsyst_deg" in ori_data:
                        arr["azimuth_pvsyst_deg"] = ori_data["azimuth_pvsyst_deg"]
                    if "azimuth_compass_deg" in ori_data:
                        arr["azimuth_deg"] = ori_data["azimuth_compass_deg"]
                        arr["azimuth_compass_deg"] = ori_data["azimuth_compass_deg"]

        return arrays

    # -------------------------------------------------------------------------
    # Array losses parsing (unchanged from v2)
    # -------------------------------------------------------------------------

    def parse_array_losses_section(self, content: str) -> Dict[str, Any]:
        parsed: Dict[str, Any] = {}
        lines = content.splitlines()

        sections: Dict[str, List[str]] = {"array_losses": lines}
        current_section: Optional[str] = None
        current_lines: List[str] = []

        for line in lines:
            line = line.strip()
            if not line:
                continue

            if re.search(r"Array Soiling Losses", line, re.IGNORECASE):
                if current_section:
                    sections[current_section] = current_lines
                current_section = "soiling_losses"
                current_lines = [line]
            elif re.search(r"Thermal Loss factor", line, re.IGNORECASE):
                if current_section:
                    sections[current_section] = current_lines
                current_section = "thermal_losses"
                current_lines = [line]
            elif re.search(r"Module mismatch losses", line, re.IGNORECASE):
                if current_section:
                    sections[current_section] = current_lines
                current_section = "module_mismatch_losses"
                current_lines = [line]
            elif re.search(r"IAM loss factor", line, re.IGNORECASE):
                if current_section:
                    sections[current_section] = current_lines
                current_section = "iam_losses"
                current_lines = [line]
            elif re.search(r"AC wiring losses", line, re.IGNORECASE):
                if current_section:
                    sections[current_section] = current_lines
                current_section = "ac_wiring_losses"
                current_lines = [line]
            else:
                current_lines.append(line)

        if current_section:
            sections[current_section] = current_lines

        if "array_losses" in sections:
            parsed["dc_wiring_losses"] = self._parse_dc_wiring_losses(
                sections["array_losses"]
            )

        for sec, sec_lines in sections.items():
            if sec == "soiling_losses":
                parsed["soiling_losses"] = self._parse_soiling_losses(sec_lines)
            elif sec == "thermal_losses":
                parsed["thermal_losses"] = self._parse_thermal_losses(sec_lines)
            elif sec == "module_mismatch_losses":
                parsed["module_mismatch_losses"] = self._parse_mismatch_losses(
                    sec_lines
                )
            elif sec == "iam_losses":
                parsed["iam_losses"] = self._parse_iam_losses(sec_lines)
            elif sec == "ac_wiring_losses":
                parsed["ac_wiring_losses"] = self._parse_ac_wiring_losses(sec_lines)

        return parsed

    def _parse_soiling_losses(self, lines: List[str]) -> Dict[str, Any]:
        data: Dict[str, Any] = {}
        for line in lines:
            if "Average loss Fraction" in line:
                m = re.search(r"Average loss Fraction\s+([\d.]+)%", line)
                if m:
                    data["average_loss_fraction_percent"] = float(m.group(1))
            elif re.search(r"\d+\.\d+%", line):
                parts = line.split()
                months = [
                    "Jan",
                    "Feb",
                    "Mar",
                    "Apr",
                    "May",
                    "Jun",
                    "Jul",
                    "Aug",
                    "Sep",
                    "Oct",
                    "Nov",
                    "Dec",
                ]
                if len(parts) >= 12:
                    data["monthly_percentages"] = {
                        months[i]: float(parts[i].rstrip("%")) for i in range(12)
                    }
        return data

    def _parse_thermal_losses(self, lines: List[str]) -> Dict[str, Any]:
        data: Dict[str, Any] = {}
        for line in lines:
            if "Loss Fraction" in line and "Module temperature" not in line:
                m = re.search(r"Loss Fraction\s+(-?[\d.]+)%", line)
                if m:
                    data["loss_fraction_percent"] = float(m.group(1))
            elif "Uc (const)" in line:
                m = re.search(r"Uc \(const\)\s+([\d.]+)", line)
                if m:
                    data["uc_const_w_per_m2_k"] = float(m.group(1))
            elif "Uv (wind)" in line:
                m = re.search(r"Uv \(wind\)\s+([\d.]+)", line)
                if m:
                    data["uv_wind_w_per_m2_k_per_ms"] = float(m.group(1))
        return data

    def _parse_mismatch_losses(self, lines: List[str]) -> Dict[str, Any]:
        data: Dict[str, Any] = {}
        for line in lines:
            if "Loss Fraction" in line:
                m = re.search(r"Loss Fraction\s+([\d.]+)%", line)
                if m:
                    data["loss_fraction_percent"] = float(m.group(1))
        return data

    def _parse_iam_losses(self, lines: List[str]) -> Dict[str, Any]:
        data: Dict[str, Any] = {}
        for line in lines:
            line = line.strip()
            if not line:
                continue
            if "DC wiring losses" in line or "Array #" in line:
                break
            if "Incidence effect (IAM):" in line:
                m = re.search(r"Incidence effect \(IAM\):\s+(.+)", line)
                if m:
                    data["incidence_effect"] = m.group(1).strip()
            elif re.search(r"\d+\.\d+", line) and not any(
                c in line for c in ["°", "mΩ", "%"]
            ):
                parts = line.split()
                if all(p.replace(".", "").replace("-", "").isdigit() for p in parts):
                    factors = [float(p) for p in parts]
                    angles = [0, 20, 30, 40, 50, 60, 70, 80, 90]
                    data["iam_profile"] = dict(zip(angles, factors))
        return data

    def _parse_dc_wiring_losses(self, lines: List[str]) -> Dict[str, Any]:
        data: Dict[str, Any] = {"arrays": []}
        full_text = " ".join(lines)

        if "Global wiring resistance" in full_text:
            m = re.search(
                r"Global wiring resistance\s+([\d.]+)mΩ\s+Loss Fraction\s+([\d.]+)%",
                full_text,
            )
            if m:
                data["global_wiring_resistance_mohm"] = float(m.group(1))
                data["global_loss_fraction_percent"] = float(m.group(2))

        notations: List[Tuple[int, str]] = []
        for match in re.finditer(
            r"Array #(\d+)\s*-\s*(.+?)(?=Array #|\s*Global|$)", full_text
        ):
            notations.append((int(match.group(1)), match.group(2).strip()))

        res_list = re.findall(r"Global array res\.\s*([\d.]+)mΩ", full_text)
        loss_list = re.findall(r"Loss Fraction\s+([\d.]+)%", full_text)

        if (
            notations
            and len(res_list) >= len(notations)
            and len(loss_list) >= len(notations)
        ):
            for (array_id, notation), res, loss in zip(
                notations, res_list[: len(notations)], loss_list[: len(notations)]
            ):
                data["arrays"].append(
                    {
                        "array_id": array_id,
                        "notation": notation,
                        "global_array_resistance_mohm": float(res),
                        "loss_fraction_percent": float(loss),
                    }
                )

        return data

    def _parse_ac_wiring_losses(self, lines: List[str]) -> Dict[str, Any]:
        data: Dict[str, Any] = {}
        for line in lines:
            if "Loss Fraction" in line:
                m = re.search(r"Loss Fraction\s+([\d.]+)%", line)
                if m:
                    data["loss_fraction_percent"] = float(m.group(1))
            elif "Inverter voltage" in line:
                m = re.search(r"Inverter voltage\s+([\d.]+)Vac", line)
                if m:
                    data["inverter_voltage_vac"] = float(m.group(1))
            elif "Wire section" in line:
                m = re.search(r"Wire section\s+(.+)", line)
                if m:
                    data["wire_section"] = m.group(1).strip()
            elif "Wires length" in line:
                m = re.search(r"Wires length\s+([\d.]+)m", line)
                if m:
                    data["wires_length_m"] = float(m.group(1))
        return data

    def _write_array_losses_text(self, f, losses: Dict[str, Any]) -> None:
        for key, value in losses.items():
            f.write(f"{key.replace('_', ' ').title()}:\n")
            if isinstance(value, dict):
                for sub_key, sub_value in value.items():
                    f.write(f"  {sub_key.replace('_', ' ').title()}: {sub_value}\n")
            elif isinstance(value, list):
                for item in value:
                    if isinstance(item, dict):
                        for sub_key, sub_value in item.items():
                            f.write(
                                f"  {sub_key.replace('_', ' ').title()}: {sub_value}\n"
                            )
                        f.write("\n")
                    else:
                        f.write(f"  {item}\n")
            else:
                f.write(f"  {value}\n")
            f.write("\n")

    # -------------------------------------------------------------------------
    # Inverter types / production
    # -------------------------------------------------------------------------

    def _collect_inverter_types(self) -> List[Dict[str, Any]]:
        types: Dict[Tuple[str, str, float], Dict[str, Any]] = {}
        type_counter = 1

        for arr_data in self.arrays.values():
            man = arr_data.get("inverter_manufacturer")
            mod = arr_data.get("inverter_model")
            power = arr_data.get("inverter_unit_nom_power_kw")
            if man or mod or power is not None:
                key = (man or "", mod or "", float(power or 0))
                if key not in types:
                    type_id = f"inverter_{type_counter}"
                    types[key] = {
                        "id": type_id,
                        "manufacturer": man,
                        "model": mod,
                        "unit_nom_power_kw": power,
                    }
                    type_counter += 1
                arr_data["inverter_type_id"] = types[key]["id"]

        global_man = self.inverter_info.get("manufacturer")
        global_mod = self.inverter_info.get("model")
        global_power = self.inverter_info.get("unit_nom_power_kw")
        if global_man or global_mod or global_power is not None:
            key = (global_man or "", global_mod or "", float(global_power or 0))
            if key not in types:
                type_id = f"inverter_{type_counter}"
                types[key] = {
                    "id": type_id,
                    "manufacturer": global_man,
                    "model": global_mod,
                    "unit_nom_power_kw": global_power,
                }
            for arr_data in self.arrays.values():
                arr_data.setdefault("inverter_type_id", types[key]["id"])

        return list(types.values())

    @staticmethod
    def _format_kw(kw: Any) -> str:
        if kw is None:
            return "?"
        try:
            fkw = float(kw)
        except (TypeError, ValueError):
            return str(kw)
        if fkw.is_integer():
            return str(int(fkw))
        return str(fkw)

    def _inverter_display_name(self, inverter_id: str) -> str:
        """Format inverter display name as: Inv [n] - ([kW] kW) - [mfr] [model]."""
        type_by_id: Dict[str, Dict[str, Any]] = {
            str(t.get("id")): t
            for t in self.inverter_types
            if isinstance(t, dict) and t.get("id")
        }

        inverter_type_id: Optional[str] = None
        # Use first linked array's inverter_type_id (if present)
        for combo in self.expanded_arrays:
            if combo.get("inverter") != inverter_id:
                continue
            arr = self.arrays.get(str(combo.get("array_id")), {})
            tid = arr.get("inverter_type_id")
            if tid:
                inverter_type_id = str(tid)
                break

        type_data = type_by_id.get(inverter_type_id or "") if inverter_type_id else None

        manufacturer = None
        model = None
        unit_nom_power_kw = None

        if type_data:
            manufacturer = type_data.get("manufacturer")
            model = type_data.get("model")
            unit_nom_power_kw = type_data.get("unit_nom_power_kw")

        manufacturer = manufacturer or self.inverter_info.get("manufacturer")
        model = model or self.inverter_info.get("model")
        unit_nom_power_kw = unit_nom_power_kw or self.inverter_info.get(
            "unit_nom_power_kw"
        )

        if manufacturer is None and model is None and unit_nom_power_kw is None:
            return inverter_id

        kw_str = self._format_kw(unit_nom_power_kw)
        manu_model = f"{manufacturer or 'Unknown'} {model or ''}".strip()

        # Prefer "Inv NN" prefix for display (keep raw inverter_id elsewhere).
        label = inverter_id
        m = re.match(r"^INV([A-Za-z]*)(\d+)$", inverter_id, re.IGNORECASE)
        if m and not m.group(1):
            label = f"Inv {int(m.group(2)):02d}"

        return f"{label} - ({kw_str} kW) - {manu_model}"

    def extract_monthly_production(
        self, blocks: Dict[int, Dict[str, Any]]
    ) -> Dict[str, float]:
        print("  Extracting monthly production data...")

        self.system_monthly_globhor = {}
        monthly_data: Dict[str, float] = {}

        all_lines: List[str] = []
        for p in sorted(blocks.keys()):
            txt = blocks[p].get("full_text", "") or ""
            all_lines.extend(txt.splitlines())

        month_pattern = re.compile(
            r"^(January|February|March|April|May|June|July|August|September|October|November|December)\b"
        )

        for raw_line in all_lines:
            line = raw_line.strip()
            if not line:
                continue

            m = month_pattern.match(line)
            if not m:
                continue

            month = m.group(1)
            parts = line.split()
            if len(parts) < 8:
                continue

            if not re.match(r"[-\d.,]+$", parts[1]):
                continue

            def to_float(s: str) -> float:
                return float(s.replace(",", ""))

            try:
                globhor = to_float(parts[1])
                e_grid = to_float(parts[-2])
            except ValueError:
                continue

            self.system_monthly_globhor[month] = globhor
            monthly_data[month] = e_grid

        total_annual = sum(monthly_data.values())
        print(
            f"    Found {len(monthly_data)} months, total annual: {total_annual:,.0f} kWh"
        )

        self.system_monthly_production = monthly_data
        return monthly_data

    def extract_total_modules(self, blocks: Dict[int, Dict[str, Any]]) -> int:
        all_text = "\n".join(
            blocks[p].get("full_text") or "" for p in sorted(blocks.keys())
        )
        match = re.search(r"Nb\.\s*of\s*modules\s*(\d+)units?", all_text)
        if match:
            return int(match.group(1))

        return sum(int(a.get("number_of_modules") or 0) for a in self.arrays.values())

    def calculate_inverter_capacities_and_modules(
        self,
    ) -> Tuple[Dict[str, float], Dict[str, int]]:
        by_inverter: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
        for combo in self.expanded_arrays:
            by_inverter[combo["inverter"]].append(combo)

        array_usage_count: Dict[str, set[str]] = {}
        for inverter, combinations in by_inverter.items():
            for combo in combinations:
                array_id = str(combo["array_id"])
                array_usage_count.setdefault(array_id, set()).add(inverter)

        inverter_capacities: Dict[str, float] = {}
        inverter_modules: Dict[str, int] = {}

        print("  Calculating inverter capacities and module counts...")
        for inverter, combinations in by_inverter.items():
            total_capacity = 0.0
            total_modules = 0

            by_array: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
            for combo in combinations:
                by_array[str(combo["array_id"])].append(combo)

            for array_id, array_combos in by_array.items():
                if array_id not in self.arrays:
                    continue

                array_data = self.arrays[array_id]
                array_capacity = float(array_data.get("nominal_stc_kwp") or 0.0)
                array_modules = int(array_data.get("number_of_modules") or 0)

                num_inverters_using_array = len(array_usage_count.get(array_id, set()))
                mppts_per_inverter = len(array_combos)

                total_mppts = num_inverters_using_array * mppts_per_inverter
                if total_mppts <= 0:
                    continue

                capacity_per_mppt = array_capacity / total_mppts
                modules_per_mppt = array_modules / total_mppts

                total_capacity += capacity_per_mppt * mppts_per_inverter
                total_modules += int(modules_per_mppt * mppts_per_inverter)

            inverter_capacities[inverter] = round(total_capacity, 1)
            inverter_modules[inverter] = total_modules

        print(f"    Calculated capacities for {len(inverter_capacities)} inverters")
        return inverter_capacities, inverter_modules

    def calculate_monthly_production(
        self, blocks: Dict[int, Dict[str, Any]]
    ) -> Dict[str, Dict[str, float]]:
        monthly_data = self.extract_monthly_production(blocks)
        total_system_modules = self.extract_total_modules(blocks)
        inverter_capacities, inverter_modules = (
            self.calculate_inverter_capacities_and_modules()
        )
        self.inverter_capacities = inverter_capacities

        if not inverter_modules:
            print(
                "    No inverter/module mapping found; leaving only system_monthly_production"
            )
            self.monthly_production = {}
            return {}

        inverter_monthly: Dict[str, Dict[str, float]] = {}
        print("  Calculating monthly production allocation...")
        for inverter, module_count in inverter_modules.items():
            share = module_count / total_system_modules if total_system_modules else 0.0
            inverter_monthly[inverter] = {
                month: round(system_production * share, 0)
                for month, system_production in monthly_data.items()
            }

        self.monthly_production = inverter_monthly
        return inverter_monthly

    # -------------------------------------------------------------------------
    # Output
    # -------------------------------------------------------------------------

    def generate_text_report(self, output_path: str) -> None:
        print(f"  Generating text report: {output_path}")

        with open(output_path, "w", encoding="utf-8") as f:
            f.write("PVsyst PDF Analysis Report (V3)\n")
            f.write("=" * 60 + "\n\n")

            f.write("SUMMARY\n" + "-" * 20 + "\n")
            f.write(f"Total Arrays Found: {len(self.arrays)}\n")
            f.write(f"Total Expanded Combinations: {len(self.expanded_arrays)}\n")
            f.write(f"Total Inverters: {len(self.inverter_capacities)}\n")
            f.write(f"Sections Identified: {len(self.sections)}\n\n")

            if self.monthly_production:
                f.write("MONTHLY PRODUCTION SUMMARY\n" + "-" * 35 + "\n")
                for inverter in sorted(self.monthly_production.keys()):
                    display_name = self._inverter_display_name(inverter)
                    cap = float(self.inverter_capacities.get(inverter, 0.0) or 0.0)
                    annual = sum(self.monthly_production[inverter].values())
                    spec = (annual / cap) if cap > 0 else 0.0
                    f.write(
                        f"{display_name}: {cap:.1f} kWp, {annual:,.0f} kWh/year ({spec:.0f} kWh/kWp)\n"
                    )
                f.write("\n")

            if self.array_losses:
                f.write("ARRAY LOSSES\n" + "-" * 15 + "\n")
                self._write_array_losses_text(f, self.array_losses)

    def _build_output_data(self) -> Dict[str, Any]:
        # Ensure MPPT labels are present for all combinations.
        if self.expanded_arrays:
            self._assign_missing_mppt_labels()

        def _rename_array_id_to_config_id(obj: Any) -> Any:
            if isinstance(obj, dict):
                out: Dict[str, Any] = {}
                for k, v in obj.items():
                    key = "config_id" if k == "array_id" else k
                    out[key] = _rename_array_id_to_config_id(v)
                return out
            if isinstance(obj, list):
                return [_rename_array_id_to_config_id(x) for x in obj]
            return obj

        # Array configurations (drop internal fields)
        array_configurations: Dict[str, Any] = {}
        for array_id, array_data in self.arrays.items():
            array_configurations[array_id] = {
                k: v
                for k, v in array_data.items()
                if k
                not in [
                    "expanded_combinations",
                    "original_notation",
                    "inverter_manufacturer",
                    "inverter_model",
                    "inverter_unit_nom_power_raw",
                    "inverter_unit_nom_power_kw",
                ]
            }

        array_configurations = {
            k: _rename_array_id_to_config_id(v) for k, v in array_configurations.items()
        }

        # MPPT allocation mapping: distribute strings across unique (inv, mppt) endpoints for each array.
        mppt_allocation: Dict[Tuple[str, str, str], Dict[str, Any]] = {}

        combos_by_array: Dict[str, List[Tuple[str, str]]] = defaultdict(list)
        for combo in self.expanded_arrays:
            mppt = combo.get("mppt")
            if mppt is None:
                continue
            combos_by_array[str(combo["array_id"])].append(
                (combo["inverter"], str(mppt))
            )

        for arr_id, pairs in combos_by_array.items():
            unique_endpoints = sorted(set(pairs))
            n_endpoints = len(unique_endpoints)

            arr = self.arrays.get(arr_id, {})
            strings_val = arr.get("strings")
            series_val = arr.get("modules_in_series")
            strings = int(strings_val) if isinstance(strings_val, int) else 0
            series = int(series_val) if isinstance(series_val, int) else 0

            stc_kwp = arr.get("nominal_stc_kwp_from_module") or arr.get(
                "nominal_stc_kwp"
            )
            if not isinstance(stc_kwp, (int, float)):
                stc_kwp = None

            total_modules = strings * series

            # Special case: inferred single-configuration site (no Array # blocks).
            if arr.get("inferred_single_config"):
                strings_per_mppt_max = arr.get("inferred_strings_per_mppt_max")
                if isinstance(strings_per_mppt_max, int) and strings_per_mppt_max > 0:
                    inv_ids = self._sort_inv_ids(
                        sorted({inv for inv, _ in unique_endpoints})
                    )
                    mppt_ids = self._sort_mppt_ids(
                        sorted({mppt for _, mppt in unique_endpoints})
                    )

                    strings_alloc = self._allocate_strings_single_config(
                        inv_ids,
                        mppt_ids,
                        total_strings=strings,
                        strings_per_mppt_max=strings_per_mppt_max,
                    )

                    for inv, mppt in unique_endpoints:
                        strings_here = int(strings_alloc.get((inv, mppt), 0))
                        modules_here = strings_here * series

                        if stc_kwp and total_modules:
                            dc_here = round(
                                float(stc_kwp) * (modules_here / total_modules), 3
                            )
                        else:
                            dc_here = None

                        mppt_allocation[(inv, mppt, arr_id)] = {
                            "strings": strings_here,
                            "modules": modules_here,
                            "dc_kwp": dc_here,
                        }
                    continue

            # Default: distribute strings evenly across endpoints.
            if n_endpoints > 0:
                base = strings // n_endpoints
                remainder = strings % n_endpoints
            else:
                base = 0
                remainder = 0

            for idx, (inv, mppt) in enumerate(unique_endpoints):
                extra = 1 if idx < remainder else 0
                strings_here = base + extra
                modules_here = strings_here * series

                if stc_kwp and total_modules:
                    dc_here = round(float(stc_kwp) * (modules_here / total_modules), 3)
                else:
                    dc_here = None

                mppt_allocation[(inv, mppt, arr_id)] = {
                    "strings": strings_here,
                    "modules": modules_here,
                    "dc_kwp": dc_here,
                }

        # Associations (raw): inverter_id -> mppt -> config_id + allocation
        raw_associations: Dict[str, Dict[str, Dict[str, Any]]] = {}
        for combo in self.expanded_arrays:
            inv_id = combo["inverter"]
            mppt = combo.get("mppt")
            if mppt is None:
                continue
            mppt = str(mppt)
            config_id = str(combo["array_id"])

            raw_associations.setdefault(inv_id, {})
            alloc = mppt_allocation.get((inv_id, mppt, config_id), {})
            raw_associations[inv_id][mppt] = {"config_id": config_id, **alloc}

        # Keep raw inverter IDs as keys in JSON.
        associations: Dict[str, Dict[str, Dict[str, Any]]] = raw_associations
        self.associations = associations

        type_by_id: Dict[str, Dict[str, Any]] = {
            str(t.get("id")): t
            for t in self.inverter_types
            if isinstance(t, dict) and t.get("id") is not None
        }

        def inverter_type_for(inv_id: str) -> Optional[Dict[str, Any]]:
            inverter_type_id: Optional[str] = None
            for combo in self.expanded_arrays:
                if combo.get("inverter") != inv_id:
                    continue
                arr = self.arrays.get(str(combo.get("array_id")), {})
                tid = arr.get("inverter_type_id")
                if tid:
                    inverter_type_id = str(tid)
                    break
            if inverter_type_id and inverter_type_id in type_by_id:
                return type_by_id[inverter_type_id]
            return None

        # Inverter summary (inverter_id-keyed) with one-place monitoring config.
        inverter_summary: Dict[str, Any] = {}
        for inv_id in sorted(raw_associations.keys()):
            description = self._inverter_display_name(inv_id)
            inv_type = inverter_type_for(inv_id)

            cap = float(self.inverter_capacities.get(inv_id, 0.0) or 0.0)
            monthly = self.monthly_production.get(inv_id, {})
            annual = float(sum(monthly.values()))

            # Flatten MPPT associations into a combined list with expanded config values.
            combined: List[Dict[str, Any]] = []
            for mppt, assoc in sorted(raw_associations[inv_id].items()):
                config_id = str(assoc.get("config_id"))
                arr = self.arrays.get(config_id, {})

                strings_total = arr.get("strings")
                strings_on_mppt = assoc.get("strings")
                i_mpp_total = arr.get("i_mpp_a")

                i_mpp_mppt = i_mpp_total
                if (
                    isinstance(i_mpp_total, (int, float))
                    and isinstance(strings_total, int)
                    and strings_total > 0
                ):
                    i_mpp_per_string = i_mpp_total / strings_total
                    if isinstance(strings_on_mppt, int) and strings_on_mppt > 0:
                        i_mpp_mppt = round(i_mpp_per_string * strings_on_mppt, 3)
                    else:
                        i_mpp_mppt = round(i_mpp_per_string, 3)

                combined.append(
                    {
                        "mppt": mppt,
                        "config_id": config_id,
                        "strings": strings_on_mppt,
                        "modules": assoc.get("modules"),
                        "dc_kwp": assoc.get("dc_kwp"),
                        "tilt": arr.get("tilt"),
                        "azimuth": arr.get("azimuth_deg")
                        if arr.get("azimuth_deg") is not None
                        else arr.get("azimuth_compass_deg"),
                        "modules_in_series": arr.get("modules_in_series"),
                        "u_mpp_v": arr.get("u_mpp_v"),
                        "i_mpp_a": i_mpp_mppt,
                    }
                )

            inverter_summary[inv_id] = {
                "description": description,
                "pv_module": self.module_info,
                "inverter_type": inv_type,
                "capacity_kwp": cap,
                "annual_production_kwh": annual,
                "specific_production_kwh_per_kwp": round(annual / cap, 0)
                if cap > 0
                else 0,
                "monthly_production": monthly,
                "associations": raw_associations[inv_id],
                "combined_configuration": combined,
            }

        self.inverter_summary = inverter_summary

        total_capacity_kwp = (
            sum(self.inverter_capacities.values()) if self.inverter_capacities else 0.0
        )
        total_annual_kwh = (
            sum(self.system_monthly_production.values())
            if self.system_monthly_production
            else 0.0
        )

        return {
            "metadata": {
                "version": "v3",
                "total_arrays": len(self.arrays),
                "total_expanded_combinations": len(self.expanded_arrays),
                "total_inverters": len(associations),
                "total_system_capacity_kwp": total_capacity_kwp,
                "total_annual_production_kwh": total_annual_kwh,
            },
            "pv_module": self.module_info,
            "inverter": self.inverter_info,
            "inverter_types": self.inverter_types,
            "array_configurations": array_configurations,
            "associations": associations,
            "inverter_summary": inverter_summary,
            "system_monthly_production": self.system_monthly_production,
            "system_monthly_globhor": self.system_monthly_globhor,
            "orientations": self.orientations,
            "array_losses": _rename_array_id_to_config_id(self.array_losses),
        }

    def generate_json_output(self, output_path: str) -> None:
        print(f"  Generating JSON output: {output_path}")
        output_data = self._build_output_data()
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(output_data, f, indent=2, ensure_ascii=False)

    def to_dict(self) -> Dict[str, Any]:
        return self._build_output_data()

    # -------------------------------------------------------------------------
    # PowerTrack patch output
    # -------------------------------------------------------------------------

    @staticmethod
    def _to_int_or_none(v: Any) -> Optional[int]:
        if v is None:
            return None
        if isinstance(v, bool):
            return int(v)
        if isinstance(v, int):
            return v
        if isinstance(v, float):
            return int(round(v))
        s = str(v).strip()
        if not s:
            return None
        try:
            return int(round(float(s)))
        except ValueError:
            return None

    @staticmethod
    def _to_float_or_none(v: Any) -> Optional[float]:
        if v is None:
            return None
        if isinstance(v, bool):
            return float(int(v))
        if isinstance(v, (int, float)):
            return float(v)
        s = str(v).strip()
        if not s:
            return None
        try:
            return float(s)
        except ValueError:
            return None

    @staticmethod
    def _omit_none(d: Dict[str, Any]) -> Dict[str, Any]:
        return {k: v for k, v in d.items() if v is not None}

    @staticmethod
    def _month_full_to_powertrack_abbrev(monthly: Dict[str, Any]) -> Dict[str, int]:
        month_map = {
            "January": "jan",
            "February": "feb",
            "March": "mar",
            "April": "apr",
            "May": "may",
            "June": "jun",
            "July": "jul",
            "August": "aug",
            "September": "sep",
            "October": "oct",
            "November": "nov",
            "December": "dec",
        }

        out: Dict[str, int] = {abbr: 0 for abbr in month_map.values()}
        if not isinstance(monthly, dict):
            return out

        for full_name, abbr in month_map.items():
            f = PVsystParser._to_float_or_none(monthly.get(full_name))
            out[abbr] = int(round(f)) if f is not None else 0
        return out

    @staticmethod
    def _mppt_sort_key(row: Dict[str, Any]) -> Tuple[int, str]:
        mppt = str(row.get("mppt") or "")
        m = re.search(r"(\d+)", mppt)
        if m:
            return (int(m.group(1)), mppt)
        return (10**9, mppt)

    @staticmethod
    def _inv_id_to_powertrack_key(inv_id: str, *, used: set[int]) -> str:
        idx: Optional[int] = None
        m = re.search(r"(\d+)", str(inv_id))
        if m:
            try:
                parsed = int(m.group(1)) - 1
                if parsed >= 0:
                    idx = parsed
            except ValueError:
                idx = None

        if idx is None or idx in used:
            candidate = 0
            while candidate in used:
                candidate += 1
            idx = candidate

        used.add(idx)
        return f"PV{idx}"

    def to_powertrack_patches_by_inverter(
        self,
        *,
        omit_nulls: bool = True,
        include_optional_mpp: bool = True,
    ) -> Dict[str, Any]:
        patches: Dict[str, Any] = {}

        watts_per_panel = self._to_int_or_none(self.module_info.get("unit_nom_power_w"))
        inv_kw = self._to_float_or_none(self.inverter_info.get("unit_nom_power_kw"))

        degrade = None
        if isinstance(self.array_losses, dict):
            thermal = self.array_losses.get("thermal_losses")
            if isinstance(thermal, dict):
                degrade = self._to_float_or_none(thermal.get("loss_fraction_percent"))

        used: set[int] = set()
        inv_ids = self._sort_inv_ids(
            [str(k) for k in (self.inverter_summary or {}).keys()]
        )
        for inv_id in inv_ids:
            inv_summary = (self.inverter_summary or {}).get(inv_id)
            if not isinstance(inv_summary, dict):
                continue

            pv_key = self._inv_id_to_powertrack_key(inv_id, used=used)

            description = inv_summary.get("description")

            monthly = inv_summary.get("monthly_production")
            monthly_out = (
                self._month_full_to_powertrack_abbrev(monthly)
                if isinstance(monthly, dict)
                else None
            )

            inv_type_kw = None
            inv_type = inv_summary.get("inverter_type")
            if isinstance(inv_type, dict):
                inv_type_kw = self._to_float_or_none(inv_type.get("unit_nom_power_kw"))
            inverter_kw = inv_kw if inv_kw is not None else inv_type_kw

            rows = inv_summary.get("combined_configuration") or []
            if not isinstance(rows, list):
                rows = []

            mppt_entries: List[Dict[str, Any]] = []
            for row in sorted(
                [r for r in rows if isinstance(r, dict)], key=self._mppt_sort_key
            ):
                v_mpp = self._to_float_or_none(row.get("u_mpp_v"))
                a_mpp = self._to_float_or_none(row.get("i_mpp_a"))
                mpp_watts = (
                    (v_mpp * a_mpp)
                    if (v_mpp is not None and a_mpp is not None)
                    else None
                )

                entry: Dict[str, Any] = {
                    "numOfStrings": self._to_int_or_none(row.get("strings")),
                    "panelsPerString": self._to_int_or_none(
                        row.get("modules_in_series")
                    ),
                    "wattsPerPanel": watts_per_panel,
                    "inverterKw": inverter_kw,
                    "azimuth": self._to_float_or_none(row.get("azimuth")),
                    "tilt": self._to_float_or_none(row.get("tilt")),
                    "dcSize": self._to_float_or_none(row.get("dc_kwp")),
                }

                if include_optional_mpp:
                    entry["mppVoltage"] = v_mpp
                    entry["mppAmps"] = a_mpp
                    entry["mppWatts"] = self._to_float_or_none(mpp_watts)

                if omit_nulls:
                    entry = self._omit_none(entry)

                if entry:
                    mppt_entries.append(entry)

            pv_config: Dict[str, Any] = {"inverters": mppt_entries}
            if monthly_out is not None:
                pv_config["monthlyOutput"] = monthly_out
            if degrade is not None:
                pv_config["degrade"] = degrade

            patch: Dict[str, Any] = {"description": description, "pvConfig": pv_config}
            if omit_nulls:
                patch = self._omit_none(patch)

            patches[pv_key] = patch

        return patches

    # -------------------------------------------------------------------------
    # Top-level parse
    # -------------------------------------------------------------------------

    def parse_pdf(
        self,
        pdf_path: str,
        output_dir: Optional[str] = None,
        *,
        interactive: bool = False,
    ) -> Dict[str, Any]:
        if output_dir is None:
            output_dir = str(Path(pdf_path).parent)

        out_dir = Path(output_dir)
        out_dir.mkdir(exist_ok=True)

        pdf_name = Path(pdf_path).stem

        print(f"Parsing PVsyst PDF (V3): {pdf_path}")
        print(f"Output directory: {out_dir}")

        # Text extraction
        blocks = self.extract_text_blocks(pdf_path)

        self.sections = self.identify_sections(blocks)
        self.section_contents = self.extract_section_contents(blocks, self.sections)

        self.extract_equipment_info(blocks)
        self.orientations = self.extract_orientations(blocks)

        # Array losses (if present)
        if (
            "Array Losses" in self.section_contents
            and self.section_contents["Array Losses"]
        ):
            try:
                self.array_losses = self.parse_array_losses_section(
                    self.section_contents["Array Losses"][0]
                )
            except Exception as exc:  # noqa: BLE001
                print(f"  Warning: failed to parse array losses: {exc}")

        # Arrays
        self.arrays = self.parse_arrays_from_text(blocks, interactive=interactive)

        # Inverter types
        self.inverter_types = self._collect_inverter_types()

        # Monthly production + inverter capacities
        self.calculate_monthly_production(blocks)

        # Write outputs
        text_path = out_dir / f"{pdf_name}_analysis_v3.txt"
        json_path = out_dir / f"{pdf_name}_structured_v3.json"

        self.generate_text_report(str(text_path))
        self.generate_json_output(str(json_path))

        print("\nParsing complete!")
        print(f"  Text report: {text_path}")
        print(f"  JSON output: {json_path}")

        return self.to_dict()


def main() -> None:
    parser = argparse.ArgumentParser(description="Parse PVsyst PDF reports (V3)")
    parser.add_argument("pdf_file", help="Path to PVsyst PDF")
    parser.add_argument(
        "--output-dir", default=None, help="Output directory (default: PDF directory)"
    )
    parser.add_argument(
        "--interactive",
        action="store_true",
        help="Prompt to override array inverter/MPPT parsing",
    )
    parser.add_argument(
        "--powertrack-patch",
        action="store_true",
        help="Write PowerTrack patch JSON per inverter (keys PV0, PV1, ...)",
    )
    parser.add_argument(
        "--powertrack-patch-path",
        default=None,
        help="Output path for PowerTrack patch JSON (default: <output-dir>/<pdf>_powertrack_patch.json)",
    )
    args = parser.parse_args()

    pdf_path = args.pdf_file
    if not Path(pdf_path).exists():
        print(f"Error: PDF file not found: {pdf_path}")
        sys.exit(1)

    p = PVsystParser()
    p.parse_pdf(pdf_path, args.output_dir, interactive=args.interactive)

    if args.powertrack_patch:
        out_dir = Path(args.output_dir) if args.output_dir else Path(pdf_path).parent
        out_dir.mkdir(exist_ok=True)
        pdf_name = Path(pdf_path).stem
        patch_path = (
            Path(args.powertrack_patch_path)
            if args.powertrack_patch_path
            else (out_dir / f"{pdf_name}_powertrack_patch.json")
        )
        patches = p.to_powertrack_patches_by_inverter(omit_nulls=True)
        with open(patch_path, "w", encoding="utf-8") as f:
            json.dump(patches, f, indent=2, ensure_ascii=False)
        print(f"  PowerTrack patch JSON: {patch_path}")


if __name__ == "__main__":
    main()
