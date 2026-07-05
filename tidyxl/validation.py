"""
Data validation extraction functionality
"""

from typing import List, Optional, Union

import pandas as pd
from openpyxl import load_workbook

from ._common import resolve_sheet_names, validate_filetype

# Column order of the DataFrame returned by xlsx_validation
_VALIDATION_COLUMNS = [
    'sheet', 'ref', 'type', 'operator', 'formula1', 'formula2',
    'allow_blank', 'show_input_message', 'show_error_message',
    'prompt_title', 'prompt', 'error_title', 'error', 'error_style'
]


def xlsx_validation(
    path: str,
    sheets: Optional[Union[str, List[str]]] = None,
    check_filetype: bool = True
) -> pd.DataFrame:
    """
    Import data validation rules of cells in xlsx (Excel) files.

    Extracts data validation rules from Excel cells, including numeric ranges,
    date constraints, list restrictions, and custom formula-driven rules.

    Parameters
    ----------
    path : str
        Path to the Excel file (.xlsx or .xlsm)
    sheets : str, list of str, or None
        Worksheet names to read. If None, reads all sheets.
    check_filetype : bool, default True
        Whether to check that the file is a valid xlsx/xlsm file

    Returns
    -------
    pd.DataFrame
        A DataFrame with columns:
        - sheet: Worksheet with validation rule
        - ref: Cell addresses with rules (e.g., 'A1:A10')
        - type: Data validation type (whole, decimal, list, date, time, textLength, custom)
        - operator: Comparison operator (between, equal, notEqual, greaterThan, etc.)
        - formula1: First validation criterion
        - formula2: Second validation criterion (for between/notBetween)
        - allow_blank: Whether blank cells are allowed
        - show_input_message: Whether to show input message
        - show_error_message: Whether to show error message
        - prompt_title: Input message title
        - prompt: Input message text
        - error_title: Error message title
        - error: Error message text
        - error_style: Error style (stop, warning, information)
    """

    if check_filetype:
        validate_filetype(path)

    wb = load_workbook(filename=path, data_only=False)

    validation_list = []

    try:
        for sheet_name in resolve_sheet_names(wb, sheets):
            ws = wb[sheet_name]

            if not hasattr(ws, 'data_validations'):
                continue

            for dv in ws.data_validations.dataValidation:
                validation_list.append({
                    'sheet': sheet_name,
                    'ref': str(dv.sqref) if dv.sqref else None,
                    'type': dv.type,
                    'operator': dv.operator,
                    'formula1': dv.formula1,
                    'formula2': dv.formula2,
                    'allow_blank': dv.allowBlank,
                    'show_input_message': dv.showInputMessage,
                    'show_error_message': dv.showErrorMessage,
                    'prompt_title': dv.promptTitle,
                    'prompt': dv.prompt,
                    'error_title': dv.errorTitle,
                    'error': dv.error,
                    'error_style': dv.errorStyle
                })

    finally:
        wb.close()

    df = pd.DataFrame(validation_list, columns=_VALIDATION_COLUMNS)

    # Sort by sheet, then by ref
    if not df.empty:
        df = df.sort_values(['sheet', 'ref']).reset_index(drop=True)

    return df
