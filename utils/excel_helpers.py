from pathlib import Path
import win32com.client
import win32gui
import win32con
import time
import pandas as pd

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

def get_cell_value(cell: str) -> float:
    """
    Get the value of a specified cell from the only open Excel workbook.
    """
    # Connect to running Excel
    xl_app = win32com.client.GetActiveObject("Excel.Application")
    
    # Since only one workbook is open
    wb = xl_app.Workbooks(1)
    ws = wb.Sheets(1)  # or wb.ActiveSheet

    value = ws.Range(str(cell)).Value
    return value

def open_excel(file_path: str):
    """
    Open an Excel workbook and bring it to the front.

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
    xl_workbook.Activate()
    
    # Bring Excel window to the front
    hwnd = xl_app.Hwnd  # Get Excel window handle
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)  # Restore if minimized
    win32gui.SetForegroundWindow(hwnd)  # Bring to front

    # Small delay to ensure window gets focus
    time.sleep(0.1)
    
    print(f'Workbook "{file_path}" opened and brought to front.')

def insert_empty_column_before(file_name: str, column: str):
    """
    Insert an empty column before the specified column in an already opened Excel workbook.

    Parameters
    ----------
    file_name : str
        Name of the workbook (with extension), e.g., "myfile.xlsx"
    column : str
        Column letter before which to insert an empty column, e.g., "D"
    """
    # Connect to running Excel
    xl_app = win32com.client.GetActiveObject("Excel.Application")
    
    # Find the workbook by name
    wb = None
    for book in xl_app.Workbooks:
        if book.Name.lower() == file_name.lower():
            wb = book
            break
    
    if wb is None:
        print(f'Workbook "{file_name}" is not open.')
        return
    
    # Use the active sheet (or specify a sheet)
    ws = wb.ActiveSheet
    
    # Insert empty column before the specified column
    ws.Columns(f"{column}:{column}").Insert()
    
    print(f'Empty column inserted before column {column} in "{file_name}".')


def read_sap_clipboard_and_trim(L):
    """
    Reads SAP2000 table data from the clipboard, formats it,
    and trims rows after the first occurrence of 'C to D' == 1.

    Returns
    -------
    df_trimmed : pandas.DataFrame
        Processed and trimmed DataFrame
    """

    # Read clipboard and select required columns
    df = pd.read_clipboard(header=None).iloc[:, [1, 2, 3, 4, 5, 6]]

    # Reset index and assign column names
    df.reset_index(drop=True, inplace=True)
    df.columns = ["Step", "Displacement", "Base shear", "A to B", "B to C", "C to D"]

    # Trim at first C → D transition (if it exists)
    mask = df["C to D"] == 1
    if mask.any():
        idx = df.index[mask][0]
        df = df.loc[:idx]

    disp_idx = df.columns.get_loc("Displacement") + 1

    df.insert(
        disp_idx,
        "Drift (%)",
        df["Displacement"] * 100 / L
    )

    return df


if __name__ == "__main__":
    # Example usage
    # save_excel("Book1", "Book1.xlsx", r"C:\Users\abinashm\OneDrive - University of Nevada, Reno\phd research\codes\pyAutoGUI-SAP\trash")
    open_excel(r"C:\Users\abinashm\OneDrive - University of Nevada, Reno\phd research\codes\pyAutoGUI-SAP\trash\Book1.xlsx")

