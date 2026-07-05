"""
Formatting information extraction functionality
"""

from typing import Any, Dict, List

from openpyxl import load_workbook


def xlsx_formats(path: str) -> Dict[str, Any]:
    """
    Import xlsx (Excel) formatting information.

    This function extracts formatting information from Excel files,
    providing details about fonts, fills, borders, and number formats.

    Parameters
    ----------
    path : str
        Path to the Excel file (.xlsx or .xlsm)

    Returns
    -------
    dict
        Dictionary containing formatting information with keys:
        - fonts: font formatting details
        - fills: fill/background formatting details
        - borders: border formatting details
        - number_formats: number format details
    """

    wb = load_workbook(filename=path, data_only=False)

    try:
        formats: Dict[str, List[Any]] = {
            'fonts': [_font_info(font) for font in getattr(wb, '_fonts', [])],
            'fills': [_fill_info(fill) for fill in getattr(wb, '_fills', [])],
            'borders': [_border_info(border) for border in getattr(wb, '_borders', [])],
            'number_formats': [
                {'format_code': str(format_code) if format_code else None, 'format_id': i}
                for i, format_code in enumerate(getattr(wb, '_number_formats', []))
            ]
        }
    finally:
        wb.close()

    return formats


def _font_info(font) -> Dict[str, Any]:
    return {
        'name': font.name,
        'size': font.size,
        'bold': font.bold,
        'italic': font.italic,
        'underline': font.underline,
        'color': str(font.color.rgb) if font.color and hasattr(font.color, 'rgb') else None
    }


def _fill_info(fill) -> Dict[str, Any]:
    return {
        'fill_type': fill.fill_type,
        'start_color': str(fill.start_color.rgb) if hasattr(fill.start_color, 'rgb') else None,
        'end_color': str(fill.end_color.rgb) if hasattr(fill.end_color, 'rgb') else None
    }


def _border_info(border) -> Dict[str, Any]:
    return {
        'left': str(border.left.style) if border.left else None,
        'right': str(border.right.style) if border.right else None,
        'top': str(border.top.style) if border.top else None,
        'bottom': str(border.bottom.style) if border.bottom else None
    }
