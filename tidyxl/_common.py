"""
Shared helpers used by the public tidyxl functions
"""

from typing import List, Optional, Union

from openpyxl.workbook.workbook import Workbook


def validate_filetype(path: str) -> None:
    """
    Raise ValueError if the path does not point to an xlsx/xlsm file.

    Parameters
    ----------
    path : str
        Path to the Excel file to check
    """
    if not path.lower().endswith(('.xlsx', '.xlsm')):
        raise ValueError("File must be .xlsx or .xlsm format")


def resolve_sheet_names(
    wb: Workbook,
    sheets: Optional[Union[str, List[str]]]
) -> List[str]:
    """
    Normalize the ``sheets`` argument into a validated list of sheet names.

    Parameters
    ----------
    wb : openpyxl.workbook.workbook.Workbook
        The loaded workbook
    sheets : str, list of str, or None
        Worksheet names to read. If None, all sheets are selected.

    Returns
    -------
    List[str]
        The selected sheet names

    Raises
    ------
    ValueError
        If any requested sheet does not exist in the workbook
    """
    available_sheets = wb.sheetnames

    if sheets is None:
        return available_sheets
    sheet_names = [sheets] if isinstance(sheets, str) else list(sheets)

    for sheet_name in sheet_names:
        if sheet_name not in available_sheets:
            raise ValueError(
                f"Sheet '{sheet_name}' not found. Available sheets: {available_sheets}"
            )

    return sheet_names
