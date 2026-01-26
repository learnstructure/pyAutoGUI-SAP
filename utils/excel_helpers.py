from pathlib import Path
import win32com.client

def save_excel(file_name: str, save_as_name: str, location=None):
    """
    Attach to a running Excel instance and save a workbook.

    Parameters
    ----------
    file_name : str
        Name of the open (unsaved) workbook, e.g. "Book1"
    save_as_name : str
        Filename to save as, e.g. "Book1.xlsx"
    location : str or Path, optional
        Directory where the file should be saved.
        Defaults to the directory where this script lives.
    """

    # Connect to running Excel
    try:
        xl_app = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        raise RuntimeError("Excel is not running.")

    # Find the workbook by name
    xl_workbook = None
    for wb in xl_app.Workbooks:
        if wb.Name == file_name:
            xl_workbook = wb
            break

    if xl_workbook is None:
        raise RuntimeError(f'Workbook "{file_name}" is not open.')

    # Determine save directory
    if location is None:
        save_dir = Path(__file__).resolve().parent
    else:
        save_dir = Path(location).resolve()

    # Full save path
    excel_path = save_dir / save_as_name

    # Save the workbook
    xl_workbook.SaveAs(str(excel_path))

    print(f'Workbook "{file_name}" saved as "{excel_path}".')
    xl_workbook.Close(SaveChanges=False)

def open_excel(file_path: str):
    """
    Open an Excel workbook.

    Parameters
    ----------
    file_path : str
        Full path to the Excel file to open.
    """

    # Connect to running Excel or start a new instance
    try:
        xl_app = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        xl_app = win32com.client.Dispatch("Excel.Application")
        xl_app.Visible = True

    # Open the workbook
    xl_workbook = xl_app.Workbooks.Open(str(file_path))
    print(f'Workbook "{file_path}" opened.')

    #run a macro named 'formatPushover' if it exists
    # try:
    #     xl_app.Application.Run("formatPushover")
    #     print('Macro "formatPushover" executed.')
    # except Exception:
    #     print('Macro "formatPushover" not found or could not be executed.')

if __name__ == "__main__":
    # Example usage
    # save_excel("Book1", "Book1.xlsx", r"C:\Users\abinashm\OneDrive - University of Nevada, Reno\phd research\codes\pyAutoGUI-SAP\trash")
    open_excel(r"C:\Users\abinashm\OneDrive - University of Nevada, Reno\phd research\codes\pyAutoGUI-SAP\trash\Book1.xlsx")