# region imports
import win32com.client
import openpyxl
from openpyxl.utils import coordinate_from_string, column_index_from_string
from openpyxl.worksheet import cell_range
import re
import tkinter as tk
# endregion

# region class imports
import Scraper
import Interface
from Interface import interface
# endregion

def main_program():
    root = tk.Tk()
    app = interface(root)

    # TODO: Set up file dialog to handle selecting filepath
    msg = Scraper.open_email(
        r'C:\Users\Couch\Desktop\TimesheetReader\test.msg')

    # Load Excel workbook
    path = app.browse_file_dialog()
    wb = openpyxl.load_workbook(path)
    sheet = wb.active

    # Amount of cells (Start - Break - Finish = 3) for each day (7); 3*7days = 21
    MAX_CELL_COUNT = len(sheet['D5':'F11']*3)

    # Get list of times from email
    # TODO: Fix disgusting regex
    regex = r'\d?\d?\:?\d?\d?\s\w\.\w\.|-'

    times = Scraper.scrape_msg(msg, regex)

    # Create new list to copy times to
    # Append all elements as 0 to prefill data in Excel
    days = []
    for i in range(0, MAX_CELL_COUNT):
        days.append(0)

    times_index = 0
    for i in range(0, MAX_CELL_COUNT):
        if times_index < len(times):
            days[times_index] = str(times[times_index])
            times_index += 1

    # Format times
    days = Scraper.format_times(days)

    Interface.print_status(
        'Copying times to spreadsheet: {0} at path: {1}'.format(str(sheet), path))

    # write days data to cells
    i = 0
    for rowOfCells in sheet['D5':'F11']:
        for cell in rowOfCells:
            cell.value = days[i]
            i += 1
        print('\tRow: {0} copied!'.format(str(rowOfCells)))

    wb.save(path)

    Interface.print_status("Completed\n{0}".format('='*100))

    root.mainloop()

if __name__ == '__main__':
    main_program()