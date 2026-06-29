# Lab PDF → Excel Automation Tool

A Python desktop application that automatically extracts patient demographics and tumour marker results from lab report PDFs and appends them directly to an Excel database.

## Features
- Extracts Tube ID, Birth Date, Sex, Collection Date, Physician Name
- Automatically parses 6 tumour markers: AFP, CA125, CA15-3, CA19-9, CEA, CYFRA21-1
- User-friendly GUI with real-time field validation
- Batch processing for multiple PDFs (up to 50 at once)
- Standalone Windows executable available

## Technologies Used
- **Python** (pdfplumber, openpyxl, tkinter, regex)
- **PyInstaller** (for standalone .exe packaging)

## How to Use
1. Choose or create an Excel file with the required column layout
2. Load a lab PDF – the form fills automatically
3. Review and complete any missing fields
4. Click "Add record" – the data is appended to your Excel file
5. Batch process multiple PDFs for large workloads

## Impact
- **Time saved:** 5 minutes → 10 seconds per record (~2 hours/week)
- **Error reduction:** Eliminated transcription errors in biomarker data
- **Adoption:** Deployed to team; 0 support tickets in first week

## Installation
```bash
pip install pdfplumber openpyxl
python MYTRIAL.py
