import openpyxl
from datetime import datetime

excel_report = r"C:\Users\domenico.munno\OneDrive - Kuehne+Nagel\Desktop\Python projects\ETA_Updater\ETA_Web-scraping_solution_complete.xlsx"
save_path = r"C:\Users\domenico.munno\OneDrive - Kuehne+Nagel\Desktop\Python projects\ETA_Updater\ETA_Web-scraping_solution_complete_finalized.xlsx"

def delete_rows_with_equal_dates(file_path):
    # Load the Excel file
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    # Iterate over the rows in reverse order
    for row in reversed(list(ws.iter_rows(min_row=2))):
        # Get the values in columns 3 and 5
        col3_value = row[2].value
        col5_value = row[4].value
        # Convert the values to datetime objects if they're strings
        if isinstance(col3_value, str):
            col3_value = datetime.strptime(col3_value, '%Y-%m-%d')
        if isinstance(col5_value, str):
            col5_value = datetime.strptime(col5_value, '%Y-%m-%d')
        # Compare only the date components of the values
        if col3_value.date() == col5_value.date():
            # Delete the row
            ws.delete_rows(row[0].row)
    # Save the changes
    wb.save(save_path)

delete_rows_with_equal_dates(excel_report)