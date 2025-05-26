"""
This module provides utility functions for common tasks such as date manipulation,
directory creation, Excel file handling (saving with autofit columns),
and finding/preparing specific Excel files for processing.
"""

import os
from datetime import datetime
import shutil

import openpyxl
import pandas as pd
from dateutil.relativedelta import relativedelta
from openpyxl.utils import get_column_letter

# Constants for utils.py
EXCEL_EXTENSIONS = (
    ".xlsx",
    ".xls",
)  # Tuple of supported Excel file extensions for input files.


def get_prev_month_yyyymm() -> str:
    """
    Calculates the previous month from the current date and returns it as a string in 'YYYYMM' format.

    Returns:
        str: The previous month in 'YYYYMM' format.
    """
    today = datetime.today()
    prev_month_date = today - relativedelta(months=1)
    return prev_month_date.strftime("%Y%m")


def make_save_dir(base_save_dir: str) -> str:
    """
    Creates a subdirectory within the `base_save_dir` named after the previous month (YYYYMM).
    If the subdirectory already exists, it does nothing.

    Args:
        base_save_dir (str): The base directory where the new subdirectory will be created.
                             This path should be an existing directory.

    Returns:
        str: The full path to the created or existing subdirectory for the previous month.
    """
    prev_month_str = get_prev_month_yyyymm()
    save_dir = os.path.join(base_save_dir, prev_month_str)

    # Check if the directory exists, if not, create it.
    if not os.path.exists(save_dir):
        os.makedirs(save_dir, exist_ok=True)
        print(f"Created directory: {save_dir}")

    return save_dir


def save_excel_with_autofit(df: pd.DataFrame, path: str):
    """
    Saves a Pandas DataFrame to an Excel file and autofits the column widths.

    Args:
        df (pd.DataFrame): The DataFrame to save.
        path (str): The full path (including filename) where the Excel file will be saved.
    """
    df.to_excel(path, index=False)
    # Load the workbook and select the active sheet to adjust column widths.
    wb = openpyxl.load_workbook(path)
    ws = wb.active  # Get active worksheet

    # Iterate over columns to calculate max length for autofit.
    for idx, column_cells in enumerate(ws.columns):  # type: ignore # openpyxl.worksheet.worksheet.Worksheet.columns is a generator
        max_length = 0
        column_letter = get_column_letter(idx + 1)

        for cell in column_cells:
            try:
                if cell.value is not None:
                    # Calculate length of cell value string representation.
                    cell_value_str = str(cell.value)
                    max_length = max(max_length, len(cell_value_str))
            except Exception:
                # Skip if cell value cannot be processed.
                pass

        # Set a default minimum width if no content or very short content.
        adjusted_width = max_length + 2 if max_length > 0 else 10
        ws.column_dimensions[column_letter].width = adjusted_width  # type: ignore # openpyxl.worksheet.dimensions.ColumnDimension.width expects float

    wb.save(path)
    wb.close()


def find_and_prepare_excel_file(
    download_dir: str,
    file_prefix: str,
    save_dir: str,
    output_file_basename: str,
    prev_month: str,
) -> tuple[pd.DataFrame | None, str | None]:
    """
    Finds the first Excel file in the `download_dir` that starts with `file_prefix`,
    copies it to a structured path within `save_dir` using `output_file_basename` and `prev_month`,
    and then reads this copied file into a Pandas DataFrame.

    The copied file will always have an '.xlsx' extension, regardless of the original file's extension ('.xls' or '.xlsx').

    Args:
        download_dir (str): The directory to search for the source Excel file.
        file_prefix (str): The prefix the source Excel file is expected to have (e.g., "LoginHistory_").
        save_dir (str): The base directory where the found Excel file will be copied.
                        A subdirectory for `prev_month` might be implicitly handled by `make_save_dir`
                        if `save_dir` is a result of that, or this function will create `save_dir` if it doesn't exist.
        output_file_basename (str): The base name for the copied Excel file (e.g., "LoginReport").
                                    The final name will be like "LoginReport_YYYYMM.xlsx".
        prev_month (str): The previous month in 'YYYYMM' format, used for naming the copied file.

    Returns:
        tuple[pd.DataFrame | None, str | None]: A tuple containing:
            - The DataFrame read from the copied Excel file. None if no file was found or an error occurred.
            - The full path to the saved (copied) Excel file. None if no file was found.

    Raises:
        EnvironmentError: If `download_dir` is not specified (empty or None).
        RuntimeError: If an error occurs while reading the Excel file.
    """
    if not download_dir:
        raise EnvironmentError("Download directory ('download_dir') is not specified.")

    # Search for Excel files (both .xlsx and .xls) starting with the given prefix.
    excel_files = [
        f
        for f in os.listdir(download_dir)
        if f.startswith(file_prefix) and f.lower().endswith(EXCEL_EXTENSIONS)
    ]

    if not excel_files:
        print(
            f"Warning: No Excel file starting with '{file_prefix}' found in '{download_dir}'."
        )
        return None, None

    # Select the first found file.
    source_file_path = os.path.join(download_dir, excel_files[0])

    # Define the path for the copied file, standardizing to .xlsx extension.
    # Example: /path/to/save_dir/OutputBaseName_YYYYMM.xlsx
    destination_save_path = os.path.join(
        save_dir, f"{output_file_basename}_{prev_month}.xlsx"
    )

    # Ensure the save directory exists before copying.
    os.makedirs(save_dir, exist_ok=True)
    shutil.copy2(source_file_path, destination_save_path)
    print(f"Copied '{source_file_path}' to '{destination_save_path}'")

    try:
        # Read the *copied* file into a DataFrame.
        df = pd.read_excel(destination_save_path)
    except Exception as e:
        # It's generally better to catch specific exceptions, but for simplicity:
        raise RuntimeError(
            f"Error reading the Excel file '{destination_save_path}': {e}"
        )

    return df, destination_save_path
