# excel_converter.py

import os
import win32com.client as win32
import logging

def convert_to_pdf(excel_path, pdf_path):
    """
    Uses Microsoft Excel to convert a .xlsx file to a .pdf file
    with specific print settings for each sheet.

    Returns:
        bool: True on success, False on failure.
    """
    excel = None
    workbook = None
    try:
        # Get the absolute paths to be safe
        excel_path = os.path.abspath(excel_path)
        pdf_path = os.path.abspath(pdf_path)

        # Start an invisible Excel instance
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        # Open the workbook
        workbook = excel.Workbooks.Open(excel_path)

        # --- Define Print Constants ---
        xlLandscape = 2
        xlPortrait = 1
        xlFitToPagesWide = 1
        xlFitToPagesTall = 1

        # --- Apply Settings to each sheet ---
        # Sheet 1: Dashboard
        ws1 = workbook.Sheets("Dashboard")
        ws1.PageSetup.Orientation = xlLandscape
        ws1.PageSetup.Zoom = False
        ws1.PageSetup.FitToPagesWide = 1
        ws1.PageSetup.FitToPagesTall = 1

        # Sheet 2: User File Data
        ws2 = workbook.Sheets("User File Data")
        ws2.PageSetup.Orientation = xlPortrait
        ws2.PageSetup.Zoom = False
        ws2.PageSetup.FitToPagesWide = 1
        ws2.PageSetup.FitToPagesTall = False # Allow multiple pages tall

        # Sheet 3: Analysis Report
        ws3 = workbook.Sheets("Analysis Report")
        ws3.PageSetup.Orientation = xlLandscape
        ws3.PageSetup.Zoom = False
        ws3.PageSetup.FitToPagesWide = 1
        ws3.PageSetup.FitToPagesTall = False # Allow multiple pages tall
        
        # --- START OF FIX ---
        # Instead of passing a list of objects, we pass a list of sheet NAMES.
        # This is a much more reliable way to select multiple sheets via COM.
        sheet_names_to_export = ["Dashboard", "User File Data", "Analysis Report"]
        
        workbook.Sheets(sheet_names_to_export).Select()
        # --- END OF FIX ---

        # Export the current selection (which is our 3 sheets) to a single PDF file
        workbook.ActiveSheet.ExportAsFixedFormat(0, pdf_path)

        return True

    except Exception as e:
        logging.error(f"Excel to PDF conversion failed: {e}", exc_info=True)
        return False
    finally:
        # Close and clean up COM objects to prevent zombie Excel processes
        if workbook:
            workbook.Close(SaveChanges=False)
        if excel:
            excel.Quit()
        workbook = None
        excel = None