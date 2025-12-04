"""SOA (weekly) manager: allocate trips per SOA with local numbering (1-20 per week).

Usage examples:
  from soa_manager import init_soa_db, allocate_trip_to_soa
  init_soa_db()  # initialize records
  trip_no = allocate_trip_to_soa(97)   # returns 1 (first trip for SOA 97)
  trip_no = allocate_trip_to_soa(97)   # returns 2 (second trip for SOA 97)

This module keeps one JSON file under `Database/`:
- `soa_records.json` maps SOA numbers to lists of trip objects (with local numbering 1-20).

The source list of SOA entries is read from `SOA.json` if present.
Max trips per SOA is 20; if more trips are needed, allocate to a different SOA.
"""

from pathlib import Path
import json
from datetime import datetime
from typing import Optional, List
import shutil
import os
from datetime import date

ROOT = Path(__file__).parent
SOA_JSON = ROOT / "SOA.json"
RECORDS_JSON = ROOT / "soa_records.json"
MAX_TRIPS_PER_SOA = 20


def load_soa_list() -> List[dict]:
    """Load the SOA list from `SOA.json` if present, else return empty list."""
    if not SOA_JSON.exists():
        return []
    with open(SOA_JSON, "r", encoding="utf-8") as f:
        return json.load(f)


def init_soa_db() -> None:
	"""Initialize records file if missing."""
	if not RECORDS_JSON.exists():
		with open(RECORDS_JSON, "w", encoding="utf-8") as f:
			json.dump({}, f, indent=2)


def _read_records() -> dict:
	if not RECORDS_JSON.exists():
		return {}
	with open(RECORDS_JSON, "r", encoding="utf-8") as f:
		return json.load(f)


def _write_records(records: dict) -> None:
	with open(RECORDS_JSON, "w", encoding="utf-8") as f:
		json.dump(records, f, indent=2)


def allocate_trip_to_soa(soa_number: int, trip_metadata: Optional[dict] = None) -> int:
	"""Allocate a trip to the given SOA with local numbering (1-20).

	Args:
		soa_number: SOA week number
		trip_metadata: optional dict with metadata like date, plate, origin, ticket, blocks, waybill

	Returns the allocated trip number (local to this SOA, 1-20).
	Raises ValueError if SOA already has 20 trips.
	"""
	records = _read_records()
	soa_key = str(int(soa_number))
	
	arr = records.get(soa_key, [])
	
	# Check if SOA is full (20 trips max)
	if len(arr) >= MAX_TRIPS_PER_SOA:
		raise ValueError(f"SOA {soa_number} already has {len(arr)} trips (max {MAX_TRIPS_PER_SOA}). Use a different SOA.")
	
	# Local trip number is 1-based index
	local_trip_no = len(arr) + 1
	
	# Store trip object with metadata
	if trip_metadata:
		obj = {"trip": local_trip_no}
		obj.update(trip_metadata)
	else:
		obj = {"trip": local_trip_no}
	
	arr.append(obj)
	records[soa_key] = arr
	_write_records(records)
	return local_trip_no


def get_soa_trips(soa_number: int) -> List[dict]:
	"""Return list of trip objects (with metadata) for an SOA."""
	records = _read_records()
	return records.get(str(int(soa_number)), [])
def get_soa_summary(soa_number: int) -> dict:
	trips = get_soa_trips(soa_number)
	return {
		"soa": int(soa_number),
		"num_trips": len(trips),
		"trips": trips,
		"max_trips": MAX_TRIPS_PER_SOA,
	}


def get_next_unclaimed_soa() -> Optional[int]:
	"""Read SOA.json and return the first SOA number with timestamp=null, else None."""
	soa_list = load_soa_list()
	for entry in soa_list:
		if entry.get("timestamp") is None:
			return int(entry.get("SOA"))
	return None
def allocate_multiple_to_soa(soa_number: int, count: int, trip_metadata_list: Optional[List[dict]] = None) -> List[int]:
	"""Allocate multiple trips to SOA with optional metadata per trip.
	
	Raises ValueError if allocating would exceed 20 trips per SOA.
	"""
	allocated = []
	for i in range(count):
		meta = trip_metadata_list[i] if trip_metadata_list and i < len(trip_metadata_list) else None
		allocated.append(allocate_trip_to_soa(soa_number, trip_metadata=meta))
	return allocated
def export_soa_to_excel(soa_number: int, template_path: Optional[Path] = None, out_dir: Optional[Path] = None,
                        start_row: int = 11, col: str = "A") -> Optional[Path]:
    """Export the SOA trips into a dated copy of an Excel template.

    - `template_path`: Path to the template file. Defaults to `reference/print_soa_weekly.xlsx` relative to repo root.
    - `out_dir`: directory where the dated copy will be placed (defaults to `output/` folder).
    - `start_row` and `col` control where the trip data will be written (default column A, row 11).

    Returns the Path to the exported file or None on error.
    """
    try:
        import openpyxl
    except Exception:
        raise Exception("openpyxl required: python -m pip install openpyxl")

    # default template path (project root reference folder)
    repo_root = ROOT.parent
    if template_path is None:
        template_path = repo_root / "reference" / "print_soa_weekly.xlsx"
    if out_dir is None:
        out_dir = repo_root / "output"

    # ensure output directory exists
    out_dir.mkdir(parents=True, exist_ok=True)

    if not template_path.exists():
        print(f"Template not found: {template_path}")
        return None

    today = date.today()
    stamp = today.strftime("%b-%d-%Y").upper()
    # Filename format: SEP-28-2025(LTS97).xlsx
    out_name = f"{stamp}(LTS{soa_number}).xlsx"
    out_path = out_dir / out_name

    # Copy template to out_path
    shutil.copy2(template_path, out_path)

    # Load workbook and write trips
    wb = openpyxl.load_workbook(out_path)
    ws = wb.active

    trips = get_soa_trips(soa_number)
    # Header
    ws["A1"] = f"SOA {soa_number} - exported {today}"

    # Clear A11:H24 first
    for row in range(11, 25):
        for col_ord in range(ord('A'), ord('H') + 1):
            cell = f"{chr(col_ord)}{row}"
            ws[cell] = None

    # Populate rows 11+ with trip metadata
    # Columns: A=trip, B=date, C=origin, D=to (manual), E=plate, F=ticket, G=blocks, H=waybill
    r = start_row  # start_row = 11 by default
    for trip_obj in trips:
        # trip_obj is like {"trip": 1, "date": "...", "plate": "...", ...} (local numbering 1-20)
        ws[f"A{r}"] = trip_obj.get("trip")
        ws[f"B{r}"] = trip_obj.get("date")
        ws[f"C{r}"] = trip_obj.get("origin")
        ws[f"D{r}"] = None  # "to" is manual entry
        ws[f"E{r}"] = trip_obj.get("plate")
        ws[f"F{r}"] = trip_obj.get("ticket")
        ws[f"G{r}"] = trip_obj.get("blocks")
        ws[f"H{r}"] = trip_obj.get("waybill")
        r += 1

    wb.save(out_path)

    # Open the file on Windows
    try:
        os.startfile(str(out_path))
    except Exception:
        # not fatal; just return the path
        pass

    return out_path


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Manage SOA weekly trip allocations")
    parser.add_argument("action", choices=["init", "alloc", "alloc-n", "summary"], help="action")
    parser.add_argument("soa", type=int, nargs="?", help="SOA number (e.g. 97)")
    parser.add_argument("n", type=int, nargs="?", help="count for alloc-n")
    args = parser.parse_args()

    if args.action == "init":
        init_soa_db()
        print("Initialized SOA DB.")
    elif args.action == "alloc":
        if not args.soa:
            parser.error("alloc requires soa number")
        init_soa_db()
        trip = allocate_trip_to_soa(args.soa)
        print(f"Allocated trip {trip} to SOA {args.soa}")
    elif args.action == "alloc-n":
        if not args.soa or not args.n:
            parser.error("alloc-n requires soa and n")
        init_soa_db()
        trips = allocate_multiple_to_soa(args.soa, args.n)
        print(f"Allocated trips to SOA {args.soa}: {trips}")
    elif args.action == "summary":
        if not args.soa:
            parser.error("summary requires soa")
        print(get_soa_summary(args.soa))
