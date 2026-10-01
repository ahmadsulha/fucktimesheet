from openpyxl.styles import Alignment, Font
import copy
import sys
import glob
import os
import pandas as pd
import subprocess
import openpyxl

COORD_2SEN = "3.1356035766711665, 101.69252036618404"
COORD_WFH = "3.195078, 101.758454"


def process_and_convert(oil_days, input_xlsx="timesheet.xlsx", output_dir=r"D:\timesheet", output_pdf="processed_timesheet.pdf"):
    
    matching_files = glob.glob(r"D:\timesheet\*.xls")

    if not matching_files:
        raise FileNotFoundError("No timesheet file found in D:\\timesheet\\")

    latest_file = max(matching_files, key=os.path.getmtime)
    print(f"Processing file: {latest_file}")
    
    input_file = ensure_xlsx(latest_file)
    
    wb = openpyxl.load_workbook(input_file)
    ws = wb.active

    reason_cell = None
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and "Reason" in str(cell.value):
                reason_cell = cell
                break
        if reason_cell:
            break

    if reason_cell:
        print(f"Found 'Reason' at cell coordinate: {reason_cell.coordinate}")
    else:
        print("Cell containing 'Reason' was not found.")

    create_location_cell(reason_cell, ws)
    populate_reason_col(ws, oil_days)
    populate_location_col(ws)
    
    
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    processed_xlsx = os.path.join(output_dir, "processed_timesheet_test.xlsx")
    wb.save(processed_xlsx)
    print(f"Successfully generate timesheet at {processed_xlsx}!")

    
    cmd = ["soffice", "--headless", "--convert-to", "pdf", "--outdir", output_dir, processed_xlsx]
    subprocess.run(cmd, check=True)
    expected_pdf = os.path.join(output_dir, "processed_timesheet_test.pdf")
    print(f"Successfully converted {processed_xlsx} to {expected_pdf}!")

def populate_xlsx():
    pass

def ensure_xlsx(file_path: str, output_dir: str = r"D:\timesheet") -> str:
    """Converts .xls to .xlsx natively using LibreOffice, preserving all layout & formatting."""
    if file_path.lower().endswith('.xls'):
        print(f"Converting {file_path} to .xlsx via LibreOffice...")
        
        cmd = ["soffice", "--headless", "--convert-to", "xlsx", "--outdir", output_dir, file_path]
        subprocess.run(cmd, check=True)
        
        # Constructs path: D:\timesheet\filename.xlsx
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        xlsx_path = os.path.join(output_dir, f"{base_name}.xlsx")
        
        return xlsx_path
        
    return file_path

def create_location_cell(cell_to_copy, ws):
    target_cell = ws["BU9"]
    target_cell.value = "Location"

    target_cell.font = copy.copy(cell_to_copy.font)
    target_cell.fill = copy.copy(cell_to_copy.fill)
    target_cell.border = copy.copy(cell_to_copy.border)
    target_cell.alignment = copy.copy(cell_to_copy.alignment)

def populate_location_col(ws):
    font_size_8 = Font(size=8, name="ARIAL")

    for row_num in range(13, 44):
        reason_cell = ws.cell(row=row_num, column=69)
        location_cell = ws[f"BS{row_num}"]

        if reason_cell.value and "wio-2sen" in str(reason_cell.value).lower():
            location_cell.value = COORD_2SEN
            location_cell.font = font_size_8
        elif reason_cell.value and "wfh" in str(reason_cell.value).lower():
            location_cell.value = COORD_WFH
            location_cell.font = font_size_8

def populate_reason_col(ws, oil_days):
    font_size_8 = Font(size=8, name="ARIAL")
    center_align = Alignment(horizontal="center")

    ws.column_dimensions["BR"].width = 25

    oil_days = oil_days.split(",")

    for row_num in range(13, 44):
        day_cell = ws.cell(row=row_num, column=7)
        date_cell = ws.cell(row=row_num, column=1)
        time_in_cell = ws.cell(row=row_num, column=9)
        time_out_cell = ws.cell(row=row_num, column=13)

        day_val = str(day_cell.value or "").lower()
        ordinal_date = str(date_cell.value or "").split('-')[0]
        
        reason_cell = ws[f"BQ{row_num}"]

        ws.merge_cells(f"BQ{row_num}:BR{row_num}")

        if ordinal_date in oil_days:
            reason_cell.value = "OIL"
            reason_cell.font = font_size_8
            reason_cell.alignment = center_align
        elif day_val in ("wed", "thu") and not reason_cell.value:
            reason_cell.value = "WIO-2SEN"
            reason_cell.font = font_size_8
            reason_cell.alignment = center_align
        elif day_val in ("mon", "tue", "fri") and not reason_cell.value:
            reason_cell.value = "WFH"
            reason_cell.font = font_size_8
            reason_cell.alignment = center_align
        elif reason_cell.value and reason_cell.value.lower() == "abs":
            if day_val in ("wed", "thu") and time_in_cell.value is not None and time_in_cell.value == "0":
                reason_cell.value = "WIO-2SEN-Forgot clockin"
                reason_cell.font = font_size_8
                reason_cell.alignment = center_align
            elif day_val in ("wed", "thu") and time_out_cell.value is not None and time_out_cell.value == 0:
                reason_cell.value = "WIO-2SEN-Forgot clockout"
                reason_cell.font = font_size_8
                reason_cell.alignment = center_align
            elif day_val in ("mon", "tue", "fri") and time_in_cell.value is not None and time_in_cell.value == "0":
                reason_cell.value = "WFH-Forgot clockin"
                reason_cell.font = font_size_8
                reason_cell.alignment = center_align
            elif day_val in ("mon", "tue", "fri") and time_out_cell.value is not None and time_out_cell.value == "0":
                reason_cell.value = "WFH-Forgot clockout"
                reason_cell.font = font_size_8
                reason_cell.alignment = center_align

def main():
    input_xlsx = sys.argv[1] if len(sys.argv) > 1 else None
    oil_days = sys.argv[2] if len(sys.argv) > 2 else None 
    process_and_convert(input_xlsx=input_xlsx, oil_days=oil_days)

if __name__ == "__main__":
    process_and_convert()