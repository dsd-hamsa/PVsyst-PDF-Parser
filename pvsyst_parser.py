#!/usr/bin/env python3
"""
PVsyst PDF Parser V1.0

A comprehensive parser for PVsyst PDF reports that handles complex inverter/MPPT grouping
notation and converts them to structured, readable formats with monthly production data.

"""

import json
import re
import sys
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple, Set
import camelot
import pdfplumber
from collections import defaultdict


class PVsystParser:
    """Comprehensive parser for PVsyst PDF reports with monthly production calculation."""

    def __init__(self):
        """Initialize the parser."""
        self.sections = {}
        self.tables = {}
        self.arrays = {}
        self.expanded_arrays = []
        self.monthly_production = {}
        self.inverter_capacities = {}
        self.system_monthly_production = {}
        self.orientations = {}
        self.system_monthly_globhor = {}
        self.module_info = {}
        self.inverter_info = {}
        self.associations = {}
        self.inverter_summary = {}

    def clean_nom_power(self, power_str: str) -> Optional[float]:
        """
        Convert strings like '595Wp', '62.5kWac', '540 W', '1.2MWp', '700 wp'
        into a float in kW or W depending on PVsyst conventions.
        
        RULES:
        - PV module "Unit Nom. Power" is always W (or Wp)
        - Inverter "Unit Nom. Power" is always kWac or kW
        - Strip all letters, keep numeric + decimal
        - Convert MW → kW if desired (optional)
        """

        if not power_str:
            return None

        s = power_str.strip().lower()

        # Detect unit
        is_mw = "mw" in s       # megawatt
        is_kw = "kw" in s       # kilowatt
        # If neither MW nor kW is present, assume module W/Wp

        # Extract numeric portion
        m = re.search(r"([0-9]*\.?[0-9]+)", s)
        if not m:
            return None

        value = float(m.group(1))

        # normalize to kW for inverter, W for module
        if is_mw:
            return value * 1000.0   # MW → kW
        elif is_kw:
            return value            # already kW
        else:
            return value            # module W


    def extract_equipment_info(self, blocks: Dict[int, Dict]) -> None:
        """
        Extract PV module & inverter info from the 'PV Array Characteristics' section.

        Looks for a block like:

            PV Array Characteristics
            PV module Inverter
            Manufacturer Hanwha Q Cells Manufacturer SMA
            Model Q.Peak-Duo-XL-G11S.3 / BFG-595 Model Sunny Tripower_Core1 62-US-41
            (Original PVsyst database) (Custom parameters definition)
            Unit Nom. Power 595Wp Unit Nom. Power 62.5kWac
        """
        # Combine all text from all pages
        full_text = ""
        for _, page_data in blocks.items():
            full_text += (page_data.get("full_text") or "") + "\n"

        # Find the "PV module Inverter" header and grab a local window after it
        m = re.search(r"PV\s+module\s+Inverter(.{0,600})", full_text,
                      re.IGNORECASE | re.DOTALL)
        if not m:
            return  # nothing found, leave module_info/inverter_info empty

        block = "PV module Inverter" + m.group(1)
        lines = [ln.strip() for ln in block.splitlines() if ln.strip()]

        module_info = {}
        inverter_info = {}

        # Manufacturer line
        manu_line = next((ln for ln in lines if "Manufacturer" in ln), None)
        if manu_line:
            mm = re.search(
                r"Manufacturer\s+(.+?)\s+Manufacturer\s+(.+)",
                manu_line,
                re.IGNORECASE,
            )
            if mm:
                module_info["manufacturer"] = mm.group(1).strip()
                inverter_info["manufacturer"] = mm.group(2).strip()

        # Model line
        model_line = next((ln for ln in lines if re.search(r"\bModel\b", ln)), None)
        if model_line:
            mm = re.search(
                r"Model\s+(.+?)\s+Model\s+(.+)",
                model_line,
                re.IGNORECASE,
            )
            if mm:
                module_info["model"] = mm.group(1).strip()
                inverter_info["model"] = mm.group(2).strip()

        # Unit Nom. Power line
        power_line = next((ln for ln in lines if "Unit Nom. Power" in ln), None)
        if power_line:
            # Example: "Unit Nom. Power 595Wp Unit Nom. Power 62.5kWac"
            mm = re.search(
                r"Unit\s+Nom\.?\s*Power\s+([0-9.,]+\s*[kM]?[Ww][A-Za-z]*)\s+"
                r"Unit\s+Nom\.?\s*Power\s+([0-9.,]+\s*[kM]?[Ww][A-Za-z]*)",
                power_line,
                re.IGNORECASE,
            )
            if mm:
                module_info["unit_nom_power_raw"] = mm.group(1).strip()
                inverter_info["unit_nom_power_raw"] = mm.group(2).strip()

                module_power = self.clean_nom_power(mm.group(1))
                if module_power is not None:
                    module_info["unit_nom_power_w"] = int(module_power)
                inverter_power = self.clean_nom_power(mm.group(2))
                if inverter_power is not None:
                    inverter_info["unit_nom_power_kw"] = inverter_power

        self.module_info = module_info
        self.inverter_info = inverter_info


    @staticmethod
    def pvsyst_azimuth_to_compass(az_pvsyst: float) -> float:
        """
        Convert PVsyst azimuth (0° = South, +West, -East)
        to compass azimuth (0° = North, 90° = East, 180° = South, 270° = West).
        """
        return (180.0 + az_pvsyst) % 360.0

    def extract_orientations(
        self, blocks: Dict[int, Dict]
    ) -> Dict[str, Dict[str, Any]]:
        """Extract Orientation #n blocks and their Tilt/Azimuth."""
        print("  Extracting orientations...")

        # Combine text from all pages
        all_text = ""
        for page_num, page_data in blocks.items():
            all_text += (page_data.get("full_text") or "") + "\n"

        orientations: Dict[str, Dict[str, Any]] = {}

        # Find each Orientation #n occurrence in order
        for m in re.finditer(r"Orientation\s*#\s*(\d+)", all_text, re.IGNORECASE):
            ori_id = m.group(1)

            # If we've already captured this orientation, skip later duplicates
            if ori_id in orientations:
                continue

            # Look at a local window after this occurrence
            window = all_text[m.start() : m.start() + 800]

            # Tilt/Azimuth 9 / 0°
            tilt_m = re.search(
                r"Tilt/Azimuth\s*([-\d.]+)\s*/\s*([-\d.]+)°", window, re.IGNORECASE
            )

            # A short "description" right after Orientation #n
            desc_m = re.search(
                r"Orientation\s*#\s*"
                + re.escape(ori_id)
                + r"\s*(.*?)(?:\n|Tilt/Azimuth)",
                window,
                re.IGNORECASE,
            )

            ori_data: Dict[str, Any] = {}

            if tilt_m:
                tilt = float(tilt_m.group(1))
                az_pv = float(tilt_m.group(2))
                az_compass = self.pvsyst_azimuth_to_compass(az_pv)

                ori_data["tilt"] = tilt
                ori_data["azimuth_pvsyst_deg"] = az_pv
                ori_data["azimuth_deg"] = az_compass
                ori_data["azimuth_compass_deg"] = az_compass

            if desc_m:
                desc = desc_m.group(1).strip()
                if desc:
                    ori_data["description"] = desc

            # Keep a small snippet for debugging / inspection
            ori_data["raw_snippet"] = window[:200]

            orientations[ori_id] = ori_data

        print(f"    Found {len(orientations)} orientations")
        return orientations

    def parse_inverter_range(self, inv_text: str) -> List[str]:
        """
        Parse complex inverter notation into individual inverter names.

        Examples:
        - "INV01" -> ["INV01"]
        - "INV02-05" -> ["INV02", "INV03", "INV04", "INV05"]
        - "INV02-05, 7,8" -> ["INV02", "INV03", "INV04", "INV05", "INV07", "INV08"]
        - "INV 9-11,13" -> ["INV09", "INV10", "INV11", "INV13"]
        """
        inverters = []

        # Clean up the input
        inv_text = inv_text.strip()

        # Split by commas to handle multiple ranges
        parts = [part.strip() for part in inv_text.split(",")]

        for part in parts:
            if "-" in part:
                # Handle ranges like "INV02-05" or "INV 9-11"
                range_match = re.search(r"INV\s*(\d+)\s*-\s*(\d+)", part, re.IGNORECASE)
                if range_match:
                    start = int(range_match.group(1))
                    end = int(range_match.group(2))
                    for i in range(start, end + 1):
                        inverters.append(f"INV{i:02d}")
            else:
                # Handle single inverters like "7" or "8"
                single_match = re.search(r"INV\s*(\d+)", part, re.IGNORECASE)
                if single_match:
                    inv_num = int(single_match.group(1))
                    inverters.append(f"INV{inv_num:02d}")
                else:
                    # Handle bare numbers like "7,8"
                    num_match = re.search(r"(\d+)", part)
                    if num_match:
                        inv_num = int(num_match.group(1))
                        inverters.append(f"INV{inv_num:02d}")

        return inverters

    def parse_mppt_range(self, mppt_text: str) -> List[str]:
        """
        Parse MPPT notation into individual MPPT numbers.

        Examples:
        - "MPPT 1" -> ["MPPT 1"]
        - "MPPT 1-5" -> ["MPPT 1", "MPPT 2", "MPPT 3", "MPPT 4", "MPPT 5"]
        - "MPPT 1,2,4" -> ["MPPT 1", "MPPT 2", "MPPT 4"]
        """
        mppts = []

        # Clean up the input
        mppt_text = mppt_text.strip()

        # Remove "MPPT" prefix if present
        mppt_text = re.sub(r"^MPPT\s*", "", mppt_text, flags=re.IGNORECASE)

        # Split by commas to handle multiple ranges
        parts = [part.strip() for part in mppt_text.split(",")]

        for part in parts:
            if "-" in part:
                # Handle ranges like "1-5"
                range_match = re.search(r"(\d+)\s*-\s*(\d+)", part)
                if range_match:
                    start = int(range_match.group(1))
                    end = int(range_match.group(2))
                    for i in range(start, end + 1):
                        mppts.append(f"MPPT {i}")
            else:
                # Handle single MPPTs
                num_match = re.search(r"(\d+)", part)
                if num_match:
                    mppt_num = int(num_match.group(1))
                    mppts.append(f"MPPT {mppt_num}")

        return mppts

    def expand_array_notation(self, array_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Expand an array into per-inverter / per-MPPT combinations based on
        the info parsed in _parse_array_block.

        Uses:
        - array_data["inverter_ids"]  (required for any expansion)
        - array_data["mppt_ids"]      (from header: MPPT 1-2, MPPT#1-3, etc.)
        - array_data["mppt_count"]    (fallback when header has no MPPT list,
                                        but the 'Number of inverters X * MPPT Y%' line exists)

        If no MPPT info is available, we still create one combo per inverter with mppt=None.
        """

        array_id = array_data.get("array_id")
        inverter_ids = array_data.get("inverter_ids") or []
        mppt_ids = array_data.get("mppt_ids")

        # Fallback: synthesize MPPT labels from mppt_count if header didn't list them
        if not mppt_ids:
            mppt_count = array_data.get("mppt_count")
            if mppt_count:
                mppt_ids = [f"MPPT {i}" for i in range(1, mppt_count + 1)]

        combos: List[Dict[str, Any]] = []

        if not inverter_ids:
            # No inverter info -> nothing to expand
            return combos

        original_notation = array_data.get("original_notation", "")

        if mppt_ids:
            # Full matrix: every MPPT on every inverter
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
            # Only know inverters, no MPPT detail
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


    def extract_tables(self, pdf_path: str) -> Dict[int, List[Dict]]:
        """Extract tables from PDF using camelot."""
        tables_by_page = {}

        print("  Extracting tables with camelot...")

        # Try lattice first (ruled tables)
        try:
            latt = camelot.read_pdf(pdf_path, pages="all", flavor="lattice")
            for t in latt:
                page_num = t.page
                if page_num not in tables_by_page:
                    tables_by_page[page_num] = []

                # Convert DataFrame to structured format
                table_data = {
                    "method": "lattice",
                    "accuracy": t.accuracy,
                    "whitespace": t.whitespace,
                    "header": [str(col) for col in t.df.columns],
                    "rows": [
                        list(map(str, row)) for row in t.df.fillna("").values.tolist()
                    ],
                }
                tables_by_page[page_num].append(table_data)

        except Exception as e:
            print(f"    Lattice extraction failed: {e}")

        # Try stream as fallback
        try:
            stream = camelot.read_pdf(pdf_path, pages="all", flavor="stream")
            for t in stream:
                page_num = t.page
                if page_num not in tables_by_page:
                    tables_by_page[page_num] = []

                # Convert DataFrame to structured format
                table_data = {
                    "method": "stream",
                    "accuracy": t.accuracy,
                    "whitespace": t.whitespace,
                    "header": [str(col) for col in t.df.columns],
                    "rows": [
                        list(map(str, row)) for row in t.df.fillna("").values.tolist()
                    ],
                }
                tables_by_page[page_num].append(table_data)

        except Exception as e:
            print(f"    Stream extraction failed: {e}")

        return tables_by_page

    def extract_text_blocks(self, pdf_path: str) -> Dict[int, Dict]:
        """Extract text blocks and key-value pairs from PDF."""
        blocks = {}

        print("  Extracting text with pdfplumber...")

        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                txt = page.extract_text() or ""
                lines = [ln for ln in (txt.splitlines()) if ln.strip()]

                # Convert Key: Value lines into pairs
                kv_pairs = []
                others = []
                for ln in lines:
                    if ":" in ln and not ln.strip().startswith(":"):
                        k, v = ln.split(":", 1)
                        if k.strip():
                            kv_pairs.append({"key": k.strip(), "value": v.strip()})
                            continue
                    others.append(ln.strip())

                blocks[i] = {"kv": kv_pairs, "text_lines": others, "full_text": txt}

        return blocks

    def identify_sections(self, blocks: Dict[int, Dict]) -> Dict[str, Dict]:
        """Identify and organize sections from the PDF text."""
        sections = {}

        print("  Identifying sections...")

        # Combine all text from all pages
        all_text = ""
        for page_num, page_data in blocks.items():
            all_text += page_data.get("full_text", "") + "\n"

        # Define section patterns
        section_patterns = {
            "Project Summary": r"Project summary|System summary|Results summary",
            "PV Array Characteristics": r"PV Array Characteristics",
            "System Losses": r"System losses|Loss diagram",
            "Main Results": r"Main results",
        }

        # Find section boundaries
        for section_name, pattern in section_patterns.items():
            matches = list(re.finditer(pattern, all_text, re.IGNORECASE))
            if matches:
                sections[section_name] = {
                    "start_positions": [m.start() for m in matches],
                    "matches": [m.group() for m in matches],
                }

        return sections

    def parse_arrays_from_text(self, blocks: Dict[int, Dict]) -> Dict[str, Dict]:
        """
        Version-agnostic array parsing using only PVsyst-controlled text.

        We:
        - find pages that contain "PV Array Characteristics" or "Array #"
        - merge those pages into one big text blob
        - split into array blocks on "Array #n"
        - parse each block for:
            * array_id
            * inverter id (INVxx)
            * Number of PV modules
            * Nominal (STC) kWp
            * Modules <strings> string(s) x <modules_in_series>
        """
        print("  Parsing array data (generic)...")

        # 1) find pages that look like PV Array Characteristics / arrays
        pages_with_arrays = []
        for page_num, page_data in blocks.items():
            full = page_data.get("full_text", "") or ""
            if re.search(r"PV Array Characteristics", full, re.IGNORECASE) or re.search(
                r"Array\s*#\s*\d+", full, re.IGNORECASE
            ):
                pages_with_arrays.append(page_num)

        if not pages_with_arrays:
            print("    No PV Array Characteristics / Array pages found.")
            return {}

        pages_with_arrays.sort()
        start_page = pages_with_arrays[0]
        end_page = pages_with_arrays[-1]

        # 2) merge text for those pages into a single blob
        combined_text = "\n".join(
            blocks[p]["full_text"] or "" for p in range(start_page, end_page + 1)
        )

        # 3) split into array blocks
        array_pattern = re.compile(
            r"(Array\s*#\s*(\d+).*?)(?=Array\s*#\s*\d+|AC wiring losses|Page \d+/\d+|$)",
            re.DOTALL | re.IGNORECASE,
        )

        arrays: Dict[str, Dict] = {}
        seen_ids = set()

        for match in array_pattern.finditer(combined_text):
            block_text = match.group(1)
            array_id = match.group(2)

            # PVsyst sometimes repeats array headers in other sections;
            # skip duplicate/short ones.
            if array_id in seen_ids:
                continue

            # require that this block actually contains a Modules line
            if not re.search(r"Modules\s+\d+\s+string", block_text, re.IGNORECASE):
                continue

            array_data = self._parse_array_block(block_text, array_id)
            arrays[array_id] = array_data
            seen_ids.add(array_id)

            print(
                f"    Parsed Array #{array_id}: "
                f"{array_data.get('modules_config_text', 'no Modules line')}"
            )

        # Use the parsed inverter / MPPT info to build combinations
        for arr in arrays.values():
            arr["expanded_combinations"] = self.expand_array_notation(arr)
            self.expanded_arrays.extend(arr["expanded_combinations"])

        print(f"    Total arrays parsed: {len(arrays)}")
        # Check for arrays that did not get expanded combinations and assign sequential MPPTs
        unexpanded_arrays = [arr_id for arr_id, arr_data in arrays.items() if not arr_data.get("expanded_combinations")]
        if unexpanded_arrays:
            print(f"    {len(unexpanded_arrays)} arrays did not get expanded, assigning sequential MPPTs")
            # Group unexpanded arrays by inverter
            arrays_by_inv = defaultdict(list)
            for arr_id in unexpanded_arrays:
                data = arrays[arr_id]
                inv_ids = data.get("inverter_ids", [])
                for inv_id in inv_ids:
                    arrays_by_inv[inv_id].append((int(arr_id), data))
            # Assign sequential MPPTs per inverter
            for inv_id, arr_list in arrays_by_inv.items():
                arr_list.sort(key=lambda x: x[0])
                mppt_idx = 1
                for array_id_int, data in arr_list:
                    array_id = str(array_id_int)
                    mppt_count = data.get("mppt_count", 1)
                    combos = []
                    for _ in range(mppt_count):
                        combos.append({
                            "array_id": array_id,
                            "inverter": inv_id,
                            "mppt": f"MPPT {mppt_idx}",
                            "original_notation": data.get("original_block_text", "").splitlines()[0],
                        })
                        mppt_idx += 1
                    data["expanded_combinations"] = combos
                    self.expanded_arrays.extend(combos)

        # If no MPPT-style headers were found (legacy format), build combos per inverter
        if not self.expanded_arrays:
            print("    No INV...MPPT headers found; using legacy inverter/MPPT assignment")

            # Group arrays by inverter_id
            arrays_by_inv = defaultdict(list)
            for array_id, data in arrays.items():
                inv_id = data.get("inverter_id")
                if inv_id:
                    arrays_by_inv[inv_id].append((int(array_id), data))

            # For each inverter, walk arrays in order and assign MPPTs sequentially
            for inv_id, arr_list in arrays_by_inv.items():
                arr_list.sort(key=lambda x: x[0])  # sort by array_id (numeric)
                mppt_idx = 1

                for array_id_int, data in arr_list:
                    array_id = str(array_id_int)
                    mppt_count = data.get("mppt_count", 1)

                    combos = []
                    for _ in range(mppt_count):
                        combos.append(
                            {
                                "array_id": array_id,
                                "inverter": inv_id,
                                "mppt": f"MPPT {mppt_idx}",
                                "original_notation": data.get(
                                    "original_block_text", ""
                                ).splitlines()[0],
                            }
                        )
                        mppt_idx += 1

                    data["expanded_combinations"] = combos
                    self.expanded_arrays.extend(combos)

        # Reassign MPPT numbers sequentially per inverter to avoid duplicates
        by_inverter = defaultdict(list)
        for combo in self.expanded_arrays:
            by_inverter[combo["inverter"]].append(combo)
        for inverter, combos in by_inverter.items():
            combos.sort(key=lambda x: (int(x["array_id"]), x["mppt"] or ""))
            mppt_idx = 1
            for combo in combos:
                combo["mppt"] = f"MPPT {mppt_idx}"
                mppt_idx += 1
        # ---- BACK-FILL ORIENTATION WHEN ONLY ONE ORIENTATION EXISTS ----
        if self.orientations and len(self.orientations) == 1:
            ori_id_str, ori_data = next(iter(self.orientations.items()))
            try:
                ori_id_int = int(ori_id_str)
            except ValueError:
                ori_id_int = ori_id_str

            for arr in arrays.values():
                if "orientation_id" not in arr:
                    arr["orientation_id"] = ori_id_int
                    # copy tilt/az if present
                    if "tilt" in ori_data:
                        arr["tilt"] = ori_data["tilt"]
                    if "azimuth_pvsyst_deg" in ori_data:
                        arr["azimuth_pvsyst_deg"] = ori_data["azimuth_pvsyst_deg"]
                    if "azimuth_compass_deg" in ori_data:
                        arr["azimuth_deg"] = ori_data["azimuth_compass_deg"]
                        arr["azimuth_compass_deg"] = ori_data["azimuth_compass_deg"]
        return arrays

    def _parse_array_block(self, section_text: str, array_id: str) -> Dict:
        """
        Parse a single Array # block using only PVsyst-controlled phrases.
        Works across PVsyst versions as long as basic wording is stable.
        """

        array_data: Dict[str, Any] = {
            "array_id": array_id,
            "original_block_text": section_text,
            "original_notation": f"Array #{array_id}",
        }

        header_line = section_text.splitlines()[0]  # "Array #1 - INV R1 - 201deg - 17/String"

        # --- UNIVERSAL INVERTER PARSER (V7/V7.4/V8.x compatible) ---
        # Extract inverter pattern from first line only (user-editable but structured)
        m_inv_any = re.search(
            r"INV\s*([A-Za-z]*)(\d+)"
            r"(?:\s*-\s*([A-Za-z]*)(\d+)\b(?!\s*(Modules|Module|String|Mod)))?",
            header_line,
            re.IGNORECASE,
        )


        inverter_ids = []
        if m_inv_any:
            prefix1, start, prefix2, end = m_inv_any.group(1), m_inv_any.group(2), m_inv_any.group(3), m_inv_any.group(4)
            prefix2 = prefix2 or prefix1
            start_n = int(start)
            end_n = int(end) if end else start_n

            for i in range(start_n, end_n + 1):
                inverter_ids.append(f"INV{prefix1}{i:02d}")

        if inverter_ids:
            array_data["inverter_ids"] = inverter_ids
            array_data["inverter_id"] = inverter_ids[0]

        # --- UNIVERSAL MPPT IDENTIFICATION (flexible across PVsyst versions) ---
        m_mppt_header = re.search(
            r"MPPT[#\s]*([0-9,\-\s]+)", 
            header_line,
            re.IGNORECASE,
        )

        mppt_ids = None
        if m_mppt_header:
            mppt_ids = self.parse_mppt_range(m_mppt_header.group(1))

        if mppt_ids:
            array_data["mppt_ids"] = mppt_ids

        m_mppt = re.search(
            r"Number of inverters\s*(\d+)\s*\*\s*MPPT\s*([\d.]+)%\s*([\d.]+)\s*unit",
            section_text,
            re.IGNORECASE,
        )

        # Orientation #n inside the block
        m_ori = re.search(
            r"Orientation\s*#\s*(\d+)",
            section_text,
            re.IGNORECASE,
        )
        if m_ori:
            array_data["orientation_id"] = int(m_ori.group(1))


        # Number of PV modules
        m_mods = re.search(
            r"Number of PV modules\s*(\d+)units?",
            section_text,
            re.IGNORECASE,
        )
        if m_mods:
            array_data["number_of_modules"] = int(m_mods.group(1))

        unit_wp = self.module_info.get("unit_nom_power_w")
        if unit_wp and "number_of_modules" in array_data:
            # Compute array nominal power from module rating
            nominal_kwp_from_module = unit_wp * array_data["number_of_modules"] / 1000.0
            array_data["nominal_stc_kwp_from_module"] = round(nominal_kwp_from_module, 3)

        # Nominal (STC) kWp
        m_stc = re.search(
            r"Nominal\s*\(STC\)\s*([\d.]+)kWp",
            section_text,
            re.IGNORECASE,
        )
        if m_stc:
            array_data["nominal_stc_kwp"] = float(m_stc.group(1))

        # Modules configuration: "Modules 5 string x 18 In series"
        m_cfg = re.search(
            r"Modules\s*(\d+)\s*string[s]?\s*x\s*(\d+)",
            section_text,
            re.IGNORECASE,
        )
        if m_cfg:
            strings = int(m_cfg.group(1))
            series = int(m_cfg.group(2))
            array_data["strings"] = strings
            array_data["modules_in_series"] = series
            array_data["modules_config_text"] = f"Modules {strings} string x {series}"

        # Tilt/Azimuth (now handles negative azimuths)
        m_tilt_az = re.search(
            r"Tilt/Azimuth\s*([-\d.]+)\s*/\s*([-\d.]+)\s*°",
            section_text,
            re.IGNORECASE,
        )
        if m_tilt_az:
            tilt = float(m_tilt_az.group(1))
            az_pv = float(m_tilt_az.group(2))
            az_compass = self.pvsyst_azimuth_to_compass(az_pv)

            array_data["tilt"] = tilt

            # Keep both conventions
            array_data["azimuth_pvsyst_deg"] = az_pv
            array_data["azimuth_deg"] = az_compass  # compass (0=N, 90=E, 180=S, 270=W)
            array_data["azimuth_compass_deg"] = az_compass  # alias, explicit name

        # U mpp / I mpp
        m_umpp = re.search(r"U mpp\s*([\d.]+)V", section_text, re.IGNORECASE)
        if m_umpp:
            array_data["u_mpp_v"] = float(m_umpp.group(1))

        m_impp = re.search(r"I mpp\s*([\d.]+)A", section_text, re.IGNORECASE)
        if m_impp:
            array_data["i_mpp_a"] = float(m_impp.group(1))

        
        if m_mppt:
            array_data["mppt_count"] = int(m_mppt.group(1))         # e.g. 1, 2, 3
            array_data["mppt_share_percent"] = float(m_mppt.group(2))  # e.g. 35.0
            array_data["inverter_unit_fraction"] = float(m_mppt.group(3))  # e.g. 0.3


        return array_data

    def extract_monthly_production(self, blocks: Dict[int, Dict]) -> Dict[str, float]:
        """
        Extract monthly production data (E_Grid) and GlobHor from the
        'Balances and main results' table using a line-based approach.

        This is more robust than a single large regex and fixes cases
        where January is missed due to slightly different spacing/formatting.
        """
        print("  Extracting monthly production data...")

        self.system_monthly_globhor = {}
        monthly_data: Dict[str, float] = {}

        # Collect all lines from all pages in order
        all_lines: list[str] = []
        for _, page_data in blocks.items():
            txt = page_data.get("full_text", "") or ""
            all_lines.extend(txt.splitlines())

        # Full month names as PVsyst prints them
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

            # Split into columns by whitespace
            parts = line.split()
            # Expected layout:
            # Month GlobHor DiffHor T_Amb GlobInc GlobEff EArray E_Grid PR
            # e.g.:
            # January 96.1 32.59 11.85 114.8 107.1 35712 34807 0.839
            if len(parts) < 8:
                # Too short to be a data row; probably the header line
                continue

            # parts[1] should be GlobHor; if it's not numeric, skip (header row)
            if not re.match(r"[-\d.,]+$", parts[1]):
                continue

            def to_float(s: str) -> float:
                return float(s.replace(",", ""))

            try:
                globhor = to_float(parts[1])    # first numeric column after Month
                e_grid = to_float(parts[-2])    # second-to-last column = E_Grid
            except ValueError:
                # Some strange non-numeric token; skip this line
                continue

            self.system_monthly_globhor[month] = globhor
            monthly_data[month] = e_grid

        total_annual = sum(monthly_data.values())
        print(f"    Found {len(monthly_data)} months, total annual: {total_annual:,.0f} kWh")

        self.system_monthly_production = monthly_data
        return monthly_data



    def extract_total_modules(self, blocks: Dict[int, Dict]) -> int:
        """Extract total system module count from PDF text blocks."""
        # Combine all text from all pages
        all_text = ""
        for page_num, page_data in blocks.items():
            all_text += page_data.get("full_text", "") + "\n"

        # Look for "Nb. of modules 1530units" pattern
        module_pattern = r"Nb\.\s*of\s*modules\s*(\d+)units?"
        match = re.search(module_pattern, all_text)

        if match:
            total_modules = int(match.group(1))
            print(f"    Total system modules: {total_modules}")
            return total_modules
        else:
            # Fallback: sum modules from array configurations
            total_modules = sum(
                array_data.get("number_of_modules", 0)
                for array_data in self.arrays.values()
            )
            print(f"    Total system modules (calculated): {total_modules}")
            return total_modules

    def calculate_inverter_capacities_and_modules(
        self,
    ) -> Tuple[Dict[str, float], Dict[str, int]]:
        """Calculate capacity and module count for each inverter based on their array usage."""
        # Group all combinations by inverter
        by_inverter = defaultdict(list)
        for combo in self.expanded_arrays:
            by_inverter[combo["inverter"]].append(combo)

        # Count how many inverters use each array
        array_usage_count = {}
        for inverter, combinations in by_inverter.items():
            for combo in combinations:
                array_id = combo["array_id"]
                if array_id not in array_usage_count:
                    array_usage_count[array_id] = set()
                array_usage_count[array_id].add(inverter)

        # Calculate capacities and modules for each inverter
        inverter_capacities = {}
        inverter_modules = {}

        print("  Calculating inverter capacities and module counts...")

        for inverter, combinations in by_inverter.items():
            total_capacity = 0.0
            total_modules = 0

            # Group by array for this inverter
            by_array = defaultdict(list)
            for combo in combinations:
                by_array[combo["array_id"]].append(combo)

            # Sum up capacity and modules from all arrays this inverter uses
            for array_id, array_combos in by_array.items():
                if array_id in self.arrays:
                    array_data = self.arrays[array_id]
                    array_capacity = array_data.get("nominal_stc_kwp", 0)
                    array_modules = array_data.get("number_of_modules", 0)

                    # Count how many inverters use this array
                    num_inverters_using_array = len(
                        array_usage_count.get(array_id, set())
                    )

                    # Count MPPTs per inverter for this array
                    mppts_per_inverter = len(array_combos)

                    # Calculate capacity and modules per MPPT for this array
                    total_mppts = num_inverters_using_array * mppts_per_inverter
                    capacity_per_mppt = (
                        array_capacity / total_mppts if total_mppts > 0 else 0
                    )
                    modules_per_mppt = (
                        array_modules / total_mppts if total_mppts > 0 else 0
                    )

                    # Add this array's contribution to the inverter
                    total_capacity += capacity_per_mppt * mppts_per_inverter
                    total_modules += int(modules_per_mppt * mppts_per_inverter)

            inverter_capacities[inverter] = round(total_capacity, 1)
            inverter_modules[inverter] = total_modules

        print(f"    Calculated capacities for {len(inverter_capacities)} inverters")
        return inverter_capacities, inverter_modules

    def calculate_monthly_production(
        self, blocks: Dict[int, Dict]
    ) -> Dict[str, Dict[str, float]]:
        """Calculate monthly production for each inverter based on module count."""

        # Extract monthly production data
        monthly_data = self.extract_monthly_production(blocks)

        # Get total system modules
        total_system_modules = self.extract_total_modules(blocks)

        # Calculate inverter capacities and modules
        inverter_capacities, inverter_modules = (
            self.calculate_inverter_capacities_and_modules()
        )

        # Store for later use
        self.inverter_capacities = inverter_capacities

        # if we have no inverter/module mapping, bail out cleanly
        if not inverter_modules:
            print("    No inverter/module mapping found (no MPPT notation).")
            print(
                "    Skipping per-inverter monthly allocation; system-level monthly only."
            )
            self.monthly_production = {}
            return {}

        # Calculate monthly production for each inverter
        inverter_monthly = {}

        print("  Calculating monthly production allocation...")

        for inverter, module_count in inverter_modules.items():
            inverter_monthly[inverter] = {}

            # Calculate this inverter's share of total production based on module count
            module_share = module_count / total_system_modules

            for month, system_production in monthly_data.items():
                inverter_production = system_production * module_share
                inverter_monthly[inverter][month] = round(inverter_production, 0)

        # Store for later use
        self.monthly_production = inverter_monthly

        print(
            f"    Calculated monthly production for {len(inverter_monthly)} inverters"
        )
        return inverter_monthly

    def generate_text_report(self, output_path: str):
        """Generate a comprehensive text report."""
        print(f"  Generating text report: {output_path}")

        with open(output_path, "w", encoding="utf-8") as f:
            f.write("PVsyst PDF Analysis Report (V2 with Monthly Production)\n")
            f.write("=" * 60 + "\n\n")

            # Summary
            f.write("SUMMARY\n")
            f.write("-" * 20 + "\n")
            f.write(f"Total Arrays Found: {len(self.arrays)}\n")
            f.write(f"Total Expanded Combinations: {len(self.expanded_arrays)}\n")
            f.write(f"Total Inverters: {len(self.inverter_capacities)}\n")
            f.write(f"Sections Identified: {len(self.sections)}\n")

            # Monthly Production Summary
            if self.monthly_production:
                f.write("MONTHLY PRODUCTION SUMMARY\n")
                f.write("-" * 35 + "\n")
                total_annual = 0
                total_capacity = 0

                for inverter in sorted(self.monthly_production.keys()):
                    capacity = self.inverter_capacities.get(inverter, 0)
                    annual = sum(self.monthly_production[inverter].values())
                    specific = annual / capacity if capacity > 0 else 0
                    total_annual += annual
                    total_capacity += capacity

                    f.write(
                        f"{inverter}: {capacity:.1f} kWp, {annual:,.0f} kWh/year ({specific:.0f} kWh/kWp)\n"
                    )

                f.write(
                    f"\nTotal System: {total_capacity:.1f} kWp, {total_annual:,.0f} kWh/year\n\n"
                )

            # Group by Inverter
            f.write("INVERTER CONFIGURATION\n")
            f.write("-" * 30 + "\n")
            f.write(
                "This section groups arrays by inverter, showing MPPT configurations:\n\n"
            )

            # Group all combinations by inverter
            by_inverter = defaultdict(list)
            for combo in self.expanded_arrays:
                by_inverter[combo["inverter"]].append(combo)

            for inverter in sorted(by_inverter.keys()):
                combinations = by_inverter[inverter]
                capacity = self.inverter_capacities.get(inverter, 0)
                f.write(f"{inverter} ({capacity:.1f} kWp)\n")
                f.write("-" * (len(inverter) + 10) + "\n")

                # Group by array for this inverter
                by_array = defaultdict(list)
                for combo in combinations:
                    by_array[combo["array_id"]].append(combo)

                for array_id in sorted(by_array.keys()):
                    array_combos = by_array[array_id]
                    array_data = self.arrays.get(array_id, {})

                    f.write(
                        f"  Array #{array_id} - {array_data.get('original_notation', 'Unknown')}\n"
                    )

                    # List MPPTs for this array
                    mppts = [
                        combo["mppt"]
                        for combo in sorted(array_combos, key=lambda x: x["mppt"])
                    ]
                    f.write(f"    MPPTs: {', '.join(mppts)}\n")

                    # Show array configuration details
                    if array_data:
                        f.write("    Configuration:\n")
                        for key, value in array_data.items():
                            if key not in [
                                "expanded_combinations",
                                "original_notation",
                            ]:
                                f.write(
                                    f"      {key.replace('_', ' ').title()}: {value}\n"
                                )
                    f.write("\n")

            # Individual Array Details
            f.write("INDIVIDUAL ARRAY DETAILS\n")
            f.write("-" * 30 + "\n")
            for array_id, array_data in self.arrays.items():
                f.write(f"Array #{array_id}\n")
                f.write(f"  Original Notation: {array_data['original_notation']}\n")
                f.write(
                    f"  Expanded to {len(array_data['expanded_combinations'])} combinations:\n"
                )

                # Group by inverter for this array
                by_inverter = defaultdict(list)
                for combo in array_data["expanded_combinations"]:
                    by_inverter[combo["inverter"]].append(combo["mppt"])

                for inverter, mppts in sorted(by_inverter.items()):
                    f.write(f"    {inverter}: {', '.join(sorted(mppts))}\n")

                f.write("  Configuration Details:\n")
                for key, value in array_data.items():
                    if key not in ["expanded_combinations", "original_notation"]:
                        f.write(f"    {key.replace('_', ' ').title()}: {value}\n")
                f.write("\n")

            # Sections
            f.write("IDENTIFIED SECTIONS\n")
            f.write("-" * 25 + "\n")
            for section_name, section_data in self.sections.items():
                f.write(f"{section_name}\n")
                f.write(
                    f"  Found at {len(section_data['start_positions'])} location(s)\n"
                )
                f.write(f"  Matches: {', '.join(section_data['matches'])}\n\n")

            # Tables Summary
            f.write("TABLES SUMMARY\n")
            f.write("-" * 20 + "\n")
            for page_num, tables in self.tables.items():
                f.write(f"Page {page_num}\n")
                for i, table in enumerate(tables):
                    f.write(
                        f"  Table {i + 1} ({table['method']}): {len(table['rows'])} rows, {len(table['header'])} columns\n"
                    )
                f.write("\n")

    def generate_json_output(self, output_path: str):
        """Generate structured JSON output with separated configurations, associations, and monthly production."""
        print(f"  Generating JSON output: {output_path}")

        # Create array configurations (clean technical specs)
        array_configurations = {}
        for array_id, array_data in self.arrays.items():
            # Clean configuration without expanded combinations
            config = {
                k: v
                for k, v in array_data.items()
                if k not in ["expanded_combinations", "original_notation"]
            }
            array_configurations[array_id] = config

        # === PER-MPPT STRING/MODULE/KWP ALLOCATION FOR EACH ARRAY ===
            mppt_allocation = {}

            combos_by_array = defaultdict(list)
            for combo in self.expanded_arrays:
                combos_by_array[combo["array_id"]].append((combo["inverter"], combo["mppt"]))

            for array_id, pairs in combos_by_array.items():

                unique_mppts = sorted(set(pairs))
                n_mppts = len(unique_mppts)

                arr = self.arrays.get(array_id, {})

                # SAFE STRING EXTRACTION
                strings = arr.get("strings")
                if not isinstance(strings, int):
                    strings = 0

                series = arr.get("modules_in_series")
                if not isinstance(series, int):
                    series = 0

                stc_kwp = (
                    arr.get("nominal_stc_kwp_from_module")
                    or arr.get("nominal_stc_kwp")
                )
                if not isinstance(stc_kwp, (int, float)):
                    stc_kwp = None

                # distribute strings safely
                if n_mppts > 0:
                    base = strings // n_mppts
                    remainder = strings % n_mppts
                else:
                    base = 0
                    remainder = 0

                for idx, (inv, mppt) in enumerate(unique_mppts):
                    extra = 1 if idx < remainder else 0
                    strings_here = base + extra
                    modules_here = strings_here * series

                    if stc_kwp:
                        dc_here = round(stc_kwp * (modules_here / (strings * series)), 3) if (strings * series) else None
                    else:
                        dc_here = None

                    mppt_allocation[(inv, mppt, array_id)] = {
                        "strings": strings_here,
                        "modules": modules_here,
                        "dc_kwp": dc_here
                    }



        # Create associations (inverter -> MPPT -> array config ID)
        associations = {}
        for combo in self.expanded_arrays:
            inv = combo["inverter"]
            mppt = combo["mppt"]
            array_id = combo["array_id"]

            if inv not in associations:
                associations[inv] = {}

            alloc = mppt_allocation.get((inv, mppt, array_id), {})

            associations[inv][mppt] = {
                "array_id": array_id,
                **alloc
            }
        
        self.associations = associations

        # Create inverter summary with monthly production
        inverter_summary = {}
        for inverter in associations.keys():
            capacity = self.inverter_capacities.get(inverter, 0)
            monthly_data = self.monthly_production.get(inverter, {})
            annual_total = sum(monthly_data.values())
            specific_production = annual_total / capacity if capacity > 0 else 0

            inverter_summary[inverter] = {
                "capacity_kwp": capacity,
                "annual_production_kwh": annual_total,
                "specific_production_kwh_per_kwp": round(specific_production, 0),
                "monthly_production": monthly_data,
            }

        self.inverter_summary = inverter_summary

        # Build a set of unique inverter IDs from arrays (for fallback)
        unique_inverters: Set[str] = set()
        for array_data in self.arrays.values():
            inv_ids = array_data.get("inverter_ids")
            if isinstance(inv_ids, list):
                unique_inverters.update(inv_ids)
            elif "inverter_id" in array_data:
                unique_inverters.add(array_data["inverter_id"])

        # Capacity fallback: use inverter capacities if present, otherwise sum arrays
        if self.inverter_capacities:
            total_capacity_kwp = sum(self.inverter_capacities.values())
        else:
            total_capacity_kwp = sum(
                array_data.get("nominal_stc_kwp", 0.0)
                for array_data in self.arrays.values()
            )

        # Annual production fallback:
        if self.monthly_production:
            total_annual_kwh = sum(
                sum(monthly.values()) for monthly in self.monthly_production.values()
            )
        elif self.system_monthly_production:
            total_annual_kwh = sum(self.system_monthly_production.values())
        else:
            total_annual_kwh = 0.0

        output_data = {
            "metadata": {
                "total_arrays": len(self.arrays),
                "total_expanded_combinations": len(self.expanded_arrays),
                "total_inverters": (
                    len(associations) if associations else len(unique_inverters)
                ),
                "total_system_capacity_kwp": total_capacity_kwp,
                "total_annual_production_kwh": total_annual_kwh,
            },
            "pv_module": self.module_info,
            "inverter": self.inverter_info,
            "array_configurations": array_configurations,
            "associations": associations,
            "inverter_summary": inverter_summary,
            "system_monthly_production": self.system_monthly_production,
            "system_monthly_globhor": self.system_monthly_globhor,
            "orientations": self.orientations,
        }

        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(output_data, f, indent=2, ensure_ascii=False)

    def to_dict(self) -> dict:
        """Return structured data as a Python dict (same structure as JSON output)."""
        # Create array configurations (clean technical specs)
        array_configurations = {}
        for array_id, array_data in self.arrays.items():
            config = {
                k: v
                for k, v in array_data.items()
                if k not in ["expanded_combinations", "original_notation"]
            }
            array_configurations[array_id] = config

        # Build a set of unique inverter IDs from arrays (for fallback)
        unique_inverters = set()
        for array_id, array_data in self.arrays.items():
            if "inverter_ids" in array_data:
                unique_inverters.update(array_data["inverter_ids"])
            else:
                notation = array_data.get("original_notation", "")
                for m in re.finditer(r"INV\s*(\d+)", notation, re.IGNORECASE):
                    inv_num = int(m.group(1))
                    unique_inverters.add(f"INV{inv_num:02d}")

        # Capacity fallback
        if self.inverter_capacities:
            total_capacity_kwp = sum(self.inverter_capacities.values())
        else:
            total_capacity_kwp = sum(
                array_data.get("nominal_stc_kwp", 0.0)
                for array_data in self.arrays.values()
            )

        # Annual production fallback
        if self.monthly_production:
            total_annual_kwh = sum(
                sum(monthly.values()) for monthly in self.monthly_production.values()
            )
        elif self.system_monthly_production:
            total_annual_kwh = sum(self.system_monthly_production.values())
        else:
            total_annual_kwh = 0.0

        # Inverter summary (same as generate_json_output)
        inverter_summary = {}
        for inverter in self.inverter_capacities.keys():
            capacity = self.inverter_capacities.get(inverter, 0)
            monthly_data = self.monthly_production.get(inverter, {})
            annual_total = sum(monthly_data.values())
            specific_production = annual_total / capacity if capacity > 0 else 0

            inverter_summary[inverter] = {
                "capacity_kwp": capacity,
                "annual_production_kwh": annual_total,
                "specific_production_kwh_per_kwp": round(specific_production, 0),
                "monthly_production": monthly_data,
            }

        # === PER-MPPT STRING/MODULE/KWP ALLOCATION FOR EACH ARRAY ===
        mppt_allocation = {}

        combos_by_array = defaultdict(list)
        for combo in self.expanded_arrays:
            combos_by_array[combo["array_id"]].append((combo["inverter"], combo["mppt"]))

        for array_id, pairs in combos_by_array.items():

            unique_mppts = sorted(set(pairs))
            n_mppts = len(unique_mppts)

            arr = self.arrays.get(array_id, {})

            # SAFE STRING EXTRACTION
            strings = arr.get("strings")
            if not isinstance(strings, int):
                strings = 0

            series = arr.get("modules_in_series")
            if not isinstance(series, int):
                series = 0

            stc_kwp = (
                arr.get("nominal_stc_kwp_from_module")
                or arr.get("nominal_stc_kwp")
            )
            if not isinstance(stc_kwp, (int, float)):
                stc_kwp = None

            # distribute strings safely
            if n_mppts > 0:
                base = strings // n_mppts
                remainder = strings % n_mppts
            else:
                base = 0
                remainder = 0

            for idx, (inv, mppt) in enumerate(unique_mppts):
                extra = 1 if idx < remainder else 0
                strings_here = base + extra
                modules_here = strings_here * series

                if stc_kwp:
                    dc_here = round(stc_kwp * (modules_here / (strings * series)), 3) if (strings * series) else None
                else:
                    dc_here = None

                mppt_allocation[(inv, mppt, array_id)] = {
                    "strings": strings_here,
                    "modules": modules_here,
                    "dc_kwp": dc_here
                }


        # Create associations (inverter -> MPPT -> array config ID)
        associations = {}
        for combo in self.expanded_arrays:
            inv = combo["inverter"]
            mppt = combo["mppt"]
            array_id = combo["array_id"]

            if inv not in associations:
                associations[inv] = {}

            alloc = mppt_allocation.get((inv, mppt, array_id), {})

            associations[inv][mppt] = {
                "array_id": array_id,
                **alloc
            }

        self.associations = associations

        return {
            "metadata": {
                "total_arrays": len(self.arrays),
                "total_expanded_combinations": len(self.expanded_arrays),
                "total_inverters": len(self.inverter_capacities) or len(unique_inverters),
                "total_system_capacity_kwp": total_capacity_kwp,
                "total_annual_production_kwh": total_annual_kwh,
            },
            "pv_module": self.module_info,
            "inverter": self.inverter_info,
            "array_configurations": array_configurations,
            "associations": associations,
            "inverter_summary": inverter_summary,
            "system_monthly_production": self.system_monthly_production,
            "system_monthly_globhor": self.system_monthly_globhor,
            "orientations": self.orientations,
        }


    def parse_pdf(self, pdf_path: str, output_dir: Optional[str] = None, generate_outputs: bool = True) -> Dict[str, Any]:
        """Main parsing function."""
        output_dir_path = None
        if generate_outputs:
            if output_dir is None:
                output_dir_path = Path(pdf_path).parent
            else:
                output_dir_path = Path(output_dir)
            output_dir_path.mkdir(exist_ok=True)

        pdf_name = Path(pdf_path).stem

        print(f"Parsing PVsyst PDF (V2 with Monthly): {pdf_path}")
        if generate_outputs and output_dir_path:
            print(f"Output directory: {output_dir_path}")

        # Extract data
        self.tables = self.extract_tables(pdf_path)
        blocks = self.extract_text_blocks(pdf_path)
        self.sections = self.identify_sections(blocks)
        self.extract_equipment_info(blocks)
        self.orientations = self.extract_orientations(blocks)
        self.arrays = self.parse_arrays_from_text(blocks)

        #        # Flatten expanded combinations
        #        self.expanded_arrays = []
        #        for array_data in self.arrays.values():
        #            self.expanded_arrays.extend(array_data['expanded_combinations'])

        # Calculate monthly production
        self.calculate_monthly_production(blocks)

        if generate_outputs and output_dir_path:
            # Generate outputs
            text_path = str(output_dir_path / f"{pdf_name}.txt")
            json_path = str(output_dir_path / f"{pdf_name}.json")

            self.generate_text_report(text_path)
            self.generate_json_output(json_path)

            print(f"\nParsing complete!")
            print(f"  Text report: {text_path}")
            print(f"  JSON output: {json_path}")
        else:
            print(f"\nParsing complete!")

        print(f"  Arrays found: {len(self.arrays)}")
        print(f"  Expanded combinations: {len(self.expanded_arrays)}")
        print(f"  Inverters: {len(self.inverter_capacities)}")
        print(f"  Total capacity: {sum(self.inverter_capacities.values()):.1f} kWp")
        print(
            f"  Total annual production: {sum(sum(monthly.values()) for monthly in self.monthly_production.values()):,.0f} kWh"
        )

        return self.to_dict()


def main():
    """Main entry point."""
    if len(sys.argv) < 2:
        print("Usage: python pvsyst_parser.py <pdf_file> [output_dir]")
        print("\nExample:")
        print(
            "  python pvsyst_parser.py 'AEP_FUSD - Farmersville HS_VC8_CPY_IFP_ALT_20240404.pdf'"
        )
        sys.exit(1)

    pdf_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None

    if not Path(pdf_path).exists():
        print(f"Error: PDF file {pdf_path} not found")
        sys.exit(1)

    # Parse the PDF
    parser = PVsystParser()
    result = parser.parse_pdf(pdf_path, output_dir)

    # Print summary
    print(f"\n=== PARSING COMPLETE ===")


if __name__ == "__main__":
    main()
