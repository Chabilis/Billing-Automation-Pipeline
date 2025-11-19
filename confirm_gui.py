# confirm_gui.py
# Lorkwen Trucking — Streamlit Confirmation GUI
# Lead Engineer: Marlou V. Bation
# Day 3 — 20 Nov 2025

import streamlit as st
import pandas as pd
from pathlib import Path
import os
from datetime import datetime

# Import your OCR brain
from main_extractor import pdfText, extractdata, enhanceInfo

st.set_page_config(page_title="Lorkwen Trucking", layout="centered")
st.title("Lorkwen Trucking Billing Automation")
st.markdown("### Trip Ticket Confirmation — Day 3")

# File uploader
uploaded_file = st.file_uploader("Drop your scanned trip ticket PDF here", type="pdf")

if uploaded_file:
    # Save temporarily
    temp_path = Path("uploads/temp_upload.pdf")
    temp_path.write_bytes(uploaded_file.getvalue())
    
    with st.spinner("Running OCR + INFO lookup..."):
        raw_text = pdfText(str(temp_path))
        ocr_data = extractdata(raw_text)
        info_data = enhanceInfo(ocr_data) or {}
    
    st.success("OCR + Database lookup complete!")
    
    # === TWO COLUMNS — LEFT: Original scan   RIGHT: Editable form ===
    col1, col2 = st.columns([1, 1])
    
    with col1:
        # Safe image display — works even if PDF has weird pages
        try:
            from pdf2image import convert_from_path
            pages = convert_from_path(str(temp_path), first_page=1, last_page=1, dpi=150)
            st.image(pages[0], caption="Original Scanned Receipt (Page 1)", width=350)
        except Exception as e:
            st.warning(f"Could not preview image: {e}")
            st.info("But OCR already worked — data is safe!")
    
    with col2:
        st.subheader("Confirm & Edit Data")
        
        with st.form("confirm_form"):
            col_a, col_b = st.columns(2)
            
            with col_a:
                trip_ticket = st.text_input("Trip Ticket", value=ocr_data.get("trip_ticket", ""))
                delivery_date = st.text_input("Delivery Date (mm/dd/yyyy)", value=ocr_data.get("delivery_date", ""))
                origin = st.text_input("Origin Keyword", value=ocr_data.get("origin", ""))
                total_blocks = st.number_input("Total Blocks", value=ocr_data.get("total_blocks", 0), step=1)
                ref_nos = st.text_input("Reference Nos", value=ocr_data.get("ref_nos", ""))
                
            with col_b:
                plate_no = st.text_input("Plate No", value=info_data.get("plate_no", ""))
                driver = st.text_input("Driver", value=info_data.get("driver", ""))
                helper1 = st.text_input("Helper 1", value=info_data.get("helper1", ""))
                helper2 = st.text_input("Helper 2", value=info_data.get("helper2", ""))
                seal_nos = st.text_input("Seal Nos (comma separated)", value=", ".join(ocr_data.get("seal_nos", [])))
            
            shipper = st.text_input("Shipper Full Name", value=info_data.get("shipper_full", ""))
            route = st.text_input("Route", value=f"{info_data.get('from_location','')} → {info_data.get('to_location','')}")
            
            submitted = st.form_submit_button("Generate 1ST&2NDTRIP File")
            
            if submitted:
                # Here tomorrow we add the Excel filling code
                st.success("Data confirmed! Ready to generate Excel file.")
                st.balloons()