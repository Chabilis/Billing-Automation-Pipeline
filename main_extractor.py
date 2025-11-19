# main_extractor.py
# Lorkwen Trucking — Final OCR + INFO Lookup Engine
# by: Marlou V. Bation
# Day 2 Final — 19 Nov 2025

# Result after testing: Not so much but Ill build Streamlit GUI later

import pandas as pd
from pathlib import Path
from ocr_extractor import pdfText, extractdata

templates = Path("templates")
infoFile = templates / "reference_data.xlsx"
dfInfo = pd.read_excel(infoFile, sheet_name="INFO")
dfInfo = dfInfo.dropna(how="all").reset_index(drop=True)

# look up columns
dfInfo["search_keywords"] = (
    dfInfo["KEYWORDS"].fillna("") + " " +
    dfInfo["ORIGINS"].fillna("") + " " +
    dfInfo["FULL"].fillna("") + " " +
    dfInfo["FROM"].fillna("")
).str.upper().str.strip()

print(f"INFO database loaded: {len(dfInfo)} records")
print(dfInfo[["TRUCK","DRIVER","HELPER 1", "ORIGINS","search_keywords"]].dropna(how="all"))

def enhanceInfo(extracted_data):
    origin_keyword = extracted_data["origin"].upper()
    
    # Search in the super keyword column
    matched = dfInfo[dfInfo["search_keywords"].str.contains(origin_keyword, na=False)]
    
    if not matched.empty:
        row = matched.iloc[0]
        print(f"MATCH FOUND → '{origin_keyword}' → {row['FULL'] or row['ORIGINS']}")
        
        def safe_str(val):
            return "" if pd.isna(val) else str(val).strip()
        
        return {
            "plate_no": safe_str(row["TRUCK"]),
            "driver": safe_str(row["DRIVER"]),
            "helper1": safe_str(row["HELPER 1"]),
            "helper2": safe_str(row["HELPER 2"]),
            "shipper_full": safe_str(row["FULL"]),
            "from_location": safe_str(row["FROM"]),
            "to_location": safe_str(row["TO"]),
        }
    else:
        print(f"No match found for: {origin_keyword}")
        return None

# test run
if __name__ == "__main__":
    samplePDF = "uploads/sampledoc2.pdf"

    if not Path(samplePDF).exists():
        print("Put pdf in uploads and try again")
    else:
        print("proccessing with INFO sheet search...")
        raw_text = pdfText(samplePDF)
        base_data = extractdata(raw_text)
        info_data = enhanceInfo(base_data)          # def enhanceinfo()

    print("\n" + " FINAL RECEIPT DATA ".center(70, "="))
    print(f"{'Trip Ticket':<20}: {base_data['trip_ticket']}")
    print(f"{'Delivery Date':<20}: {base_data.get('delivery_date', 'Not extracted yet')}")
    print(f"{'Origin Keyword':<20}: {base_data['origin']}")
    print(f"{'Total Blocks':<20}: {base_data['total_blocks']}")
    print(f"{'Reference Nos':<20}: {base_data['ref_nos']}")
    print(f"{'Seal Nos':<20}: {', '.join(base_data['seal_nos']) if base_data['seal_nos'] else 'None'}")
    print(f"{'Plate No':<20}: {info_data['plate_no'] if info_data else 'Not found'}")
    print(f"{'Driver':<20}: {info_data['driver'] if info_data else 'Not found'}")
    print(f"{'Helper 1':<20}: {info_data['helper1'] if info_data else 'Not found'}")
    print(f"{'Helper 2':<20}: {info_data['helper2'] if info_data else 'Not found'}")
    print(f"{'Shipper Full':<20}: {info_data['shipper_full'] if info_data else 'Not found'}")
    print(f"{'From → To':<20}: {info_data['from_location'] if info_data else ''} → {info_data['to_location'] if info_data else ''}")
    print("=" * 70)