# ocr_extractor.py
# Lorkwen Trucking Billing Automation Pipeline - OCR Brain
# Made by: Marlou Bation

# NOTE: change test files at line 88
#       Plate no: fix at line 66
#       Origin: line 72

import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import re
import os
from datetime import datetime

def pdfText(pdf_path):
    # Pdf pages to img
    pages = convert_from_path(pdf_path, dpi=300) 
    full_text = ""

    for i, page in enumerate(pages):
        print(f" -- Processing page {i+1}/{len(pages)}")
        text = pytesseract.image_to_string(page, lang='eng') # ocr each page
        full_text += text + "\n"
    
    print("ocr complete - real text done")
    return full_text

def extractdata(raw_text):
    print("\nGetting some data...")

    data = {
       "delivery_date": "",
        "origin": "",
        "plate_no": "",
        "trip_ticket": "",
        "total_blocks": 0,
        "ref_nos": [],
        "seal_nos": [],
        "driver": "",
        "helper1": "",
    }

    # Trip ticket
    tt_match = re.search(r"TRIP.*CKET.*?(\d{5,6})", raw_text, re.IGNORECASE | re.DOTALL)
    if tt_match:
        data["trip_ticket"] = tt_match.group(1).strip()
    
    # Seal number
    sn_match = re.search(r"SEAL.*?(\d{5,6})", raw_text, re.IGNORECASE | re.DOTALL)
    if sn_match:
        data["seal_nos"].append(sn_match.group(1).strip())

    # blks
    blks = re.findall(r"\b(1\d{2}|7\d{2})\b", raw_text) # 100 - 799 range -----------------------------
    if blks:
        data["total_blocks"] = sum(int(x) for x in blks)
    
    # ref. numbers
    ref_num = re.findall(r"\b(V[A-Z]?\d+[A-Z]?)\b", raw_text)
    if ref_num:
        if len(ref_num) > 3:
            data["ref_nos"] = "/".join(ref_num[:3]) + "/..."    # if 4 or more ref nos.
        else:
            data["ref_nos"] = "/".join(ref_num)

    # 6. Plate No — look for JAR, etc. (you can expand later) --------------------------------------
    plate_match = re.search(r"([A-Z]{3}[- ]?\d{3,4})", raw_text)
    if plate_match:
        data["plate_no"] = plate_match.group(1).replace(" ", "-")
    
    # 7. Origin — keyword matching (we’ll make this smart with INFO sheet later) ---------------------------
    origins = ["JSI LILOAN", "5G", "BB5", "JENTEC", "FAST"]
    for origin in origins:
        if origin.upper() in raw_text.upper():
            data["origin"] = origin
            break
    
    return data

# Tesseract-OCR
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

print("OCR Engine - Ready")
print("="*60)

# for test run
if __name__ == "__main__":
    sample_pdf = "uploads/sampledoc2.pdf"    # < --- Test file here
    
    if not os.path.exists(sample_pdf):
        print(f"Put your PDF in: {sample_pdf}")
    else:
        raw_text = pdfText(sample_pdf)
        
        with open("output/raw_text.txt", "w", encoding="utf-8") as f:
            f.write(raw_text)
        
        extracted = extractdata(raw_text)
        
        print("\n" + "="*60)
        print("EXTRACTED DATA — SUCCESS")
        print("="*60)
        print(f"{'Delivery Date':15}: {extracted['delivery_date']}")
        print(f"{'Origin':15}: {extracted['origin']}")
        print(f"{'Plate No':15}: {extracted['plate_no']}")
        print(f"{'Trip Ticket':15}: {extracted['trip_ticket']}")
        print(f"{'Total Blocks':15}: {extracted['total_blocks']}")
        print(f"{'Reference Nos':15}: {extracted['ref_nos']}")
        print(f"{'Seal Nos':15}: {', '.join(extracted['seal_nos'])}")
        print(f"{'Driver':15}: {extracted['driver'] or 'Not found'}")
        print(f"{'Helper':15}: {extracted['helper1'] or 'Not found'}")
        print("="*60)

