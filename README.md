# Billing Automation Pipeline

A Python-based GUI application for automating two-trip waybill entry, Excel record management, and weekly SOA (Statement of Account) reporting.

## Overview

This project streamlines the billing and documentation workflow for trip management by:
- **Two-Trip Entry**: Collect origin, destination, vehicle, and trip details for paired trips
- **Waybill Tracking**: Automatically manage waybill numbering and mark used waybills in a centralized WAYBILL RECORD
- **Weekly SOA Export**: Generate weekly SOA reports with automatic trip allocation and metadata tracking
- **Excel Integration**: Lock-safe Excel operations with retry logic and background threading

## Features

### 1. Two-Trip Waybill GUI
- **Trip Entry Form**: Input vehicle plate, waybill numbers, origin, destination, blocks, and ticket numbers
- **Auto-Increment Waybill**: Second trip waybill automatically increments from the first trip
- **Excel Export**: Fills the two-trip print template (`print_1st2ndtrip.xlsx`)
- **Waybill Record Integration**: Writes trip data to `WAYBILL RECORD.xlsx` with automatic sequencing

### 2. Waybill Management
- **Centralized Tracking**: All trips recorded in `WAYBILL RECORD.xlsx`
- **Lock-Safe Operations**: Gracefully handles file locks with retry logic and background threading
- **Status Marking**: Marks waybills as "TRANSFER" in Column C after confirmation

### 3. Weekly SOA Reporting
- **Automatic Allocation**: Each trip automatically allocated to the current SOA (Statement of Account)
- **Local Trip Numbering**: Trips numbered 1-20 per SOA week
- **Metadata Storage**: Full trip details stored (date, plate, origin, ticket, blocks, waybill)
- **Excel Export**: Generates dated SOA files (`MMM-DD-YYYY(LTS{soa}).xlsx`) with populated trip data
- **Max Capacity Protection**: Enforces 20-trip limit per SOA; prompts for new SOA if exceeded

### 4. SOA Selection on Startup
- **Dialog Prompt**: On app launch, shows next unclaimed SOA from `SOA.json`
- **Confirmation**: User can accept or decline the suggested SOA
- **Auto-Initialization**: SOA database automatically initialized on selection

## Project Structure

```
MY PROJECT/
├── GUI_userinput.py              # Main Tkinter application
├── README.md                     # This file
├── Database/
│   ├── read_waybills.py          # Waybill tracking & marking
│   ├── soa_manager.py            # SOA allocation & export
│   ├── SOA.json                  # Master list of SOA weeks (175 entries)
│   └── soa_records.json          # Trip records per SOA (auto-generated)
├── reference/
│   ├── print_1st2ndtrip.xlsx     # Template for two-trip printing
│   └── print_soa_weekly.xlsx     # Template for SOA weekly export
└── output/
    └── [dated SOA files]         # Generated SOA exports (e.g., DEC-04-2025(LTS97).xlsx)
```

## Installation

### Prerequisites
- **Python 3.8+**
- **Windows OS** (for Excel integration via `os.startfile()`)

### Dependencies
Install required packages:

```bash
pip install openpyxl
```

The following are part of Python's standard library:
- `tkinter` (GUI framework)
- `threading` (background operations)
- `json`, `pathlib`, `datetime`, `shutil`, `os`

## Usage

### Starting the Application

```bash
python GUI_userinput.py
```

**On First Launch:**
1. Dialog appears: "Create new SOA {next_soa} for today ({date})? Yes/No"
2. Click "Yes" to accept or "No" to exit
3. App opens with the selected SOA displayed

### Workflow: Single Day Entry

1. **Enter Trip 1 Data**
   - Vehicle Plate
   - Waybill Number(s) (comma-separated)
   - Origin & Destination
   - Driver Name, Helpers
   - Total Blocks, Trip Ticket Number

2. **Confirm Trip 1**
   - Data written to `print_1st2ndtrip.xlsx`
   - Data written to `WAYBILL RECORD.xlsx`
   - Trip allocated to current SOA (numbered 1, 2, 3... per SOA)
   - GUI advances to Trip 2 input
   - Waybill auto-increments for Trip 2

3. **Enter Trip 2 Data**
   - Fill same fields (waybill pre-filled with +1)

4. **Confirm Trip 2**
   - Data written to print template & waybill record
   - Trip allocated to current SOA
   - Both waybills marked as "TRANSFER" in background
   - Final summary displayed

5. **Export**
   - Click "Export to PDF" button
   - Opens `print_1st2ndtrip.xlsx` for printing
   - Automatically exports/populates SOA weekly file (`output/MMM-DD-YYYY(LTS{soa}).xlsx`)
   - GUI clears for next entry

### Error Handling

**Excel File Locked:**
- If `WAYBILL RECORD.xlsx` is open in Excel, the app gracefully skips the Excel write
- Trip data is still saved to SOA database
- Warning logged; user can try again or close Excel

**SOA Full (20 Trips):**
- If current SOA reaches 20 trips, allocation fails
- Error dialog: "SOA {number} is full (max 20 trips). Please use a different SOA."
- Manual action required: restart app and select next SOA

**Validation Errors:**
- Missing required fields show error dialog
- Log displays validation details
- User can correct and re-confirm

## Configuration Files

### SOA.json
Master list of Statement of Account weeks. Example:
```json
[
  {"SOA": 97, "timestamp": null},
  {"SOA": 98, "timestamp": null},
  {"SOA": 99, "timestamp": "2025-12-04T14:30:00"}
]
```
- `SOA`: Week number (97-271 provided)
- `timestamp`: Null = unclaimed; ISO timestamp = claimed

### soa_records.json
Auto-generated trip allocation record. Example:
```json
{
  "97": [
    {
      "trip": 1,
      "date": "2025-12-04",
      "plate": "ABC123",
      "origin": "Manila",
      "ticket": "T001",
      "blocks": 5,
      "waybill": "W001"
    },
    {
      "trip": 2,
      "date": "2025-12-04",
      "plate": "XYZ789",
      "origin": "Quezon City",
      "ticket": "T002",
      "blocks": 3,
      "waybill": "W002"
    }
  ]
}
```

## Excel Templates

### print_1st2ndtrip.xlsx
Two-trip print template with cells pre-mapped for:
- Trip 1: Plate (B3), Origin (C4), Destination (D4), Waybill (E3), etc.
- Trip 2: Plate (B13), Origin (C14), Destination (D14), Waybill (E13), etc.

**Auto-cleared after export** (not in the template; only in memory for next use).

### print_soa_weekly.xlsx
Weekly SOA template with:
- Header row (A1: "SOA {number} - exported {date}")
- Data rows 11-30 for up to 20 trips
- Columns: A=Trip#, B=Date, C=Origin, D=To (manual), E=Plate, F=Ticket, G=Blocks, H=Waybill

## Key Architecture Decisions

### 1. Local Trip Numbering
- Each SOA week has trips numbered **1-20** locally (not globally sequential)
- Simplifies manual record-keeping and matches warehouse document flow
- If > 20 trips in a week, allocate to next SOA number manually

### 2. Lock-Safe Excel Writes
- **Retry Logic**: 5 attempts with 1-second delays
- **Background Threading**: Waybill marking runs asynchronously to prevent UI freeze
- **Graceful Degradation**: If Excel file is locked, JSON is still updated (trip is claimed)

### 3. Metadata-Rich Trip Objects
- Each trip stored as a dictionary with full context (date, plate, origin, etc.)
- Enables one-step export to SOA weekly file without additional queries

### 4. Startup Dialog for SOA Selection
- Reads `SOA.json` for first unclaimed (`timestamp=null`) entry
- User confirms or declines
- Prevents accidental allocation to wrong week

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `openpyxl required` error | Run `pip install openpyxl` |
| App hangs when writing to Excel | Close `WAYBILL RECORD.xlsx` in Excel; app will retry automatically |
| Waybill not incrementing for Trip 2 | Ensure Trip 1 waybill is numeric (e.g., `1001` not `AB-1001`) |
| "SOA full" error on Trip 2 | Restart app, select next SOA number from dialog |
| Excel file not opening after export | Check `output/` folder; manually open with Excel |
| Missing template files | Ensure `reference/` folder exists with `print_1st2ndtrip.xlsx` and `print_soa_weekly.xlsx` |

## Development Notes

### Adding New Fields
To add a new trip field:
1. Add input field to `TwoTripApp.create_trip_frame()`
2. Update `validate_trip_data()` to include new field
3. Update `trip_metadata` dictionary in `confirm_trip1()` and `confirm_trip2()`
4. Update SOA export columns in `soa_manager.export_soa_to_excel()`

### Customizing SOA Numbering
To change max trips per SOA:
1. Edit `MAX_TRIPS_PER_SOA` in `Database/soa_manager.py`
2. Update Excel template row ranges (currently 11-30 for 20 trips)
3. Update GUI error message in `confirm_trip1()` and `confirm_trip2()`

### Excel Template Customization
Edit `reference/print_1st2ndtrip.xlsx` or `reference/print_soa_weekly.xlsx` directly:
- Column widths, fonts, borders are preserved
- Only data cells (e.g., B3, C4) are overwritten by the app

## Future Enhancements

- [ ] SOA timestamp auto-update (mark SOA as claimed after export)
- [ ] Batch SOA export (export multiple weeks at once)
- [ ] Trip edit/delete functionality
- [ ] Real-time trip counter display on main form
- [ ] Waybill batch import (from CSV or Excel file)
- [ ] Email notification on SOA export
- [ ] Database backup/restore functionality

## License

This project is proprietary and confidential. For internal use only.

## Support & Contact

For bugs, feature requests, or issues, please contact the development team or open an issue in the repository.

---

**Last Updated:** December 4, 2025  
**Version:** 1.0.0  
**Status:** Production Ready
