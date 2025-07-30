API Reference
=============

This page provides detailed documentation for all tidyxl functions.

Core Functions
--------------

.. currentmodule:: tidyxl

.. autosummary::
   :toctree: generated
   :nosignatures:

   xlsx_cells
   xlsx_sheet_names
   xlsx_names
   xlsx_validation
   xlsx_formats

Cell Data Extraction
--------------------

.. autofunction:: xlsx_cells

The main function for extracting Excel cell data into tidy format. Returns a pandas DataFrame where each row represents a single cell with comprehensive metadata.

**Output Columns:**

The function returns a DataFrame with 23 columns matching the R tidyxl package:

.. list-table::
   :header-rows: 1
   :widths: 20 15 65

   * - Column
     - Type
     - Description
   * - sheet
     - str
     - Worksheet name
   * - address
     - str
     - Cell address (A1 notation)
   * - row
     - int
     - Row number
   * - col
     - int
     - Column number
   * - is_blank
     - bool
     - Whether cell has a value
   * - content
     - str
     - Raw cell value before type conversion
   * - data_type
     - str
     - Cell type (character, numeric, logical, date, error, blank)
   * - error
     - str
     - Error value if cell contains error
   * - logical
     - bool
     - Boolean value if cell contains TRUE/FALSE
   * - numeric
     - float
     - Numeric value if cell contains number
   * - date
     - datetime
     - Date value if cell contains date
   * - character
     - str
     - String value if cell contains text
   * - formula
     - str
     - Formula if cell contains formula
   * - is_array
     - bool
     - Whether formula is array formula
   * - formula_ref
     - str
     - Range address for array/shared formulas
   * - formula_group
     - int
     - Formula group identifier
   * - comment
     - str
     - Cell comment text
   * - height
     - float
     - Row height in Excel units
   * - width
     - float
     - Column width in Excel units
   * - row_outline_level
     - int
     - Row outline/grouping level
   * - col_outline_level
     - int
     - Column outline/grouping level
   * - style_format
     - str
     - Style format identifier
   * - local_format_id
     - int
     - Local formatting identifier

Workbook Metadata
-----------------

.. autofunction:: xlsx_sheet_names

Returns the names of all worksheets in the Excel file as a list of strings, in the order they appear when opening the spreadsheet.

.. autofunction:: xlsx_names

Extracts named ranges and named formulas (defined names) from Excel files, including both global and sheet-specific named ranges.

**Output Columns:**

.. list-table::
   :header-rows: 1
   :widths: 20 15 65

   * - Column
     - Type
     - Description
   * - sheet
     - str
     - Sheet name (None if globally defined)
   * - name
     - str
     - Name of the formula/range
   * - formula
     - str
     - Cell range or formula definition
   * - comment
     - str
     - Description by spreadsheet author
   * - hidden
     - bool
     - Visibility status
   * - is_range
     - bool
     - Whether formula represents a cell range

Data Validation
---------------

.. autofunction:: xlsx_validation

Extracts data validation rules from Excel cells, including numeric ranges, date constraints, list restrictions, and custom formula-driven rules.

**Output Columns:**

.. list-table::
   :header-rows: 1
   :widths: 25 15 60

   * - Column
     - Type
     - Description
   * - sheet
     - str
     - Worksheet with validation rule
   * - ref
     - str
     - Cell addresses with rules (e.g., 'A1:A10')
   * - type
     - str
     - Data validation type (whole, decimal, list, date, time, textLength, custom)
   * - operator
     - str
     - Comparison operator (between, equal, notEqual, greaterThan, etc.)
   * - formula1
     - str
     - First validation criterion
   * - formula2
     - str
     - Second validation criterion (for between/notBetween)
   * - allow_blank
     - bool
     - Whether blank cells are allowed
   * - show_input_message
     - bool
     - Whether to show input message
   * - show_error_message
     - bool
     - Whether to show error message
   * - prompt_title
     - str
     - Input message title
   * - prompt
     - str
     - Input message text
   * - error_title
     - str
     - Error message title
   * - error
     - str
     - Error message text
   * - error_style
     - str
     - Error style (stop, warning, information)

Formatting Information
----------------------

.. autofunction:: xlsx_formats

Extracts formatting information from Excel files, providing details about fonts, fills, borders, and number formats.

**Output Structure:**

Returns a dictionary with the following keys:

.. list-table::
   :header-rows: 1
   :widths: 20 80

   * - Key
     - Description
   * - fonts
     - List of font formatting details (name, size, bold, italic, underline, color)
   * - fills
     - List of fill/background formatting details (fill_type, start_color, end_color)
   * - borders
     - List of border formatting details (left, right, top, bottom styles)
   * - number_formats
     - List of number format details (format_code, format_id)

Module Structure
----------------

The tidyxl package is organized into the following modules:

.. toctree::
   :maxdepth: 1

   api/cells
   api/workbook
   api/validation
   api/formats