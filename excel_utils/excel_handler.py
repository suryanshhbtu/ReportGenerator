import openpyxl

EXCEL_FILE = "templates/SampleData.xlsx"
MODIFIED_FILE = "templates/modified.xlsx"  # New file (downloadable)

def read_excel(file_path=EXCEL_FILE):
    """Reads an Excel file with multiple sheets and returns a dictionary."""
    try:
        wb = openpyxl.load_workbook(file_path)
        data = {}

        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            sheet_data = {}

            # Read non-empty cells
            for row in sheet.iter_rows(values_only=True):
                if any(row):  # Ignore empty rows
                    sheet_data[f"Row {sheet.max_row}"] = row

            data[sheet_name] = sheet_data

        return data
    except Exception as e:
        return {"error": str(e)}

def write_excel():
    """Writes dummy data to a new Excel file without overwriting the original."""
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)  # Load original file
        new_wb = openpyxl.Workbook()  # Create a new workbook

        for sheet_name in wb.sheetnames:
            original_sheet = wb[sheet_name]
            new_sheet = new_wb.create_sheet(title=sheet_name)

            # Copy content from the original sheet
            for row in original_sheet.iter_rows():
                for cell in row:
                    new_sheet[cell.coordinate] = cell.value

            # Writing dummy data
            new_sheet["A1"] = "Dummy Data 1"
            new_sheet["B2"] = "Dummy Data 2"
            new_sheet["C3"] = "Dummy Data 3"

        # Remove default empty sheet
        if "Sheet" in new_wb.sheetnames:
            std = new_wb["Sheet"]
            new_wb.remove(std)

        new_wb.save(MODIFIED_FILE)
        return {"message": "Data written successfully", "file": MODIFIED_FILE}
    except Exception as e:
        return {"error": str(e)}
