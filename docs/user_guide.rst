User Guide
==========

This comprehensive guide covers all aspects of using tidyxl for Excel data analysis.

Understanding the Tidy Format
-----------------------------

The core concept of tidyxl is the "tidy" data format, where each Excel cell becomes a single row in a pandas DataFrame. This approach provides several advantages:

* **Granular Analysis**: Examine every cell individually with complete metadata
* **Flexible Filtering**: Easily filter by data type, location, or content
* **Preserve Structure**: Maintain Excel's original structure and formatting
* **Handle Complexity**: Work with messy, non-tabular spreadsheets

The 23-Column Structure
~~~~~~~~~~~~~~~~~~~~~~~

Each row in the tidyxl output contains 23 columns with comprehensive cell information:

**Position & Identity:**
- ``sheet``: Worksheet name
- ``address``: Cell address in A1 notation (e.g., "B5")
- ``row``, ``col``: Numeric position (1-based indexing)

**Content & Type:**
- ``is_blank``: Boolean indicating if cell has content
- ``content``: Raw string representation of cell value
- ``data_type``: One of: character, numeric, logical, date, error, blank

**Typed Values:**
- ``character``: String content (if text)
- ``numeric``: Numeric value (if number)
- ``logical``: Boolean value (if TRUE/FALSE)
- ``date``: Datetime object (if date)
- ``error``: Error string (if Excel error)

**Formulas:**
- ``formula``: Excel formula (if present)
- ``is_array``: Whether it's an array formula
- ``formula_ref``: Range reference for shared formulas  
- ``formula_group``: Group ID for related formulas

**Metadata:**
- ``comment``: Cell comment text
- ``height``, ``width``: Row/column dimensions
- ``row_outline_level``, ``col_outline_level``: Grouping levels
- ``style_format``, ``local_format_id``: Formatting identifiers

Working with Different Data Types
----------------------------------

Understanding Excel Data Types
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Excel stores different types of data, and tidyxl preserves these distinctions:

.. code-block:: python

   from tidyxl import xlsx_cells

   cells = xlsx_cells("sample.xlsx")
   
   # See all data types present
   print("Data types in file:")
   print(cells['data_type'].value_counts())

   # character: Text strings, including numbers stored as text
   # numeric: Numbers, including integers and floats
   # logical: TRUE/FALSE values
   # date: Date and datetime values
   # error: Excel errors like #DIV/0!, #N/A, etc.
   # blank: Empty cells (if include_blank_cells=True)

Accessing Typed Values
~~~~~~~~~~~~~~~~~~~~~~

Use the appropriate typed column for analysis:

.. code-block:: python

   # Working with numeric data
   numeric_cells = cells[cells['data_type'] == 'numeric']
   numbers = numeric_cells['numeric'].dropna()
   print(f"Sum: {numbers.sum()}")
   print(f"Average: {numbers.mean()}")
   print(f"Range: {numbers.min()} to {numbers.max()}")

   # Working with text data
   text_cells = cells[cells['data_type'] == 'character']
   texts = text_cells['character'].dropna()
   print(f"Unique text values: {texts.unique()}")

   # Working with dates
   date_cells = cells[cells['data_type'] == 'date']
   dates = date_cells['date'].dropna()
   print(f"Date range: {dates.min()} to {dates.max()}")

   # Working with boolean data
   bool_cells = cells[cells['data_type'] == 'logical']
   booleans = bool_cells['logical'].dropna()
   print(f"TRUE count: {booleans.sum()}")
   print(f"FALSE count: {(~booleans).sum()}")

Advanced Filtering Techniques
-----------------------------

Location-Based Filtering
~~~~~~~~~~~~~~~~~~~~~~~~

Filter cells by their position in the spreadsheet:

.. code-block:: python

   # Get header row (row 1)
   headers = cells[cells['row'] == 1]

   # Get specific column (column A = 1)
   column_a = cells[cells['col'] == 1]

   # Get a range of cells (A1:C10)
   cell_range = cells[
       (cells['col'] >= 1) & (cells['col'] <= 3) &
       (cells['row'] >= 1) & (cells['row'] <= 10)
   ]

   # Get specific sheet
   sales_sheet = cells[cells['sheet'] == 'Sales']

Content-Based Filtering
~~~~~~~~~~~~~~~~~~~~~~~

Filter based on cell content:

.. code-block:: python

   # Find cells containing specific text
   search_term = "Total"
   matching_cells = cells[
       cells['character'].str.contains(search_term, na=False)
   ]

   # Find cells with numeric values above threshold
   high_values = cells[
       (cells['data_type'] == 'numeric') & 
       (cells['numeric'] > 1000)
   ]

   # Find cells with formulas
   formula_cells = cells[cells['formula'].notna()]

   # Find cells with comments
   commented_cells = cells[cells['comment'].notna()]

   # Find error cells
   error_cells = cells[cells['data_type'] == 'error']

Working with Formulas
---------------------

Formula Analysis
~~~~~~~~~~~~~~~~

tidyxl preserves Excel formulas, enabling sophisticated analysis:

.. code-block:: python

   # Get all formulas
   formulas = cells[cells['formula'].notna()]
   print(f"Found {len(formulas)} formulas")

   # Analyze formula types
   formula_types = {}
   for _, cell in formulas.iterrows():
       formula = cell['formula']
       if formula.startswith('=SUM'):
           formula_types['SUM'] = formula_types.get('SUM', 0) + 1
       elif formula.startswith('=AVERAGE'):
           formula_types['AVERAGE'] = formula_types.get('AVERAGE', 0) + 1
       elif formula.startswith('=COUNT'):
           formula_types['COUNT'] = formula_types.get('COUNT', 0) + 1
       else:
           formula_types['OTHER'] = formula_types.get('OTHER', 0) + 1

   print("Formula distribution:", formula_types)

Cross-Sheet References
~~~~~~~~~~~~~~~~~~~~~~

Find formulas that reference other sheets:

.. code-block:: python

   # Find cross-sheet references
   cross_sheet = formulas[
       formulas['formula'].str.contains('!', na=False)
   ]

   print("Cross-sheet formulas:")
   for _, cell in cross_sheet.iterrows():
       print(f"  {cell['sheet']}.{cell['address']}: {cell['formula']}")

Working with Named Ranges
--------------------------

Named ranges are Excel's way of giving meaningful names to cell ranges or formulas:

.. code-block:: python

   from tidyxl import xlsx_names

   # Get all named ranges
   names = xlsx_names("data.xlsx")
   print(f"Found {len(names)} named ranges")

   # Show named ranges
   print(names[['name', 'formula', 'sheet', 'is_range']])

   # Separate ranges from formulas
   ranges = names[names['is_range'] == True]
   formulas = names[names['is_range'] == False]

   print(f"Cell ranges: {len(ranges)}")
   print(f"Named formulas: {len(formulas)}")

Using Named Range Information
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Named ranges can help you understand spreadsheet structure:

.. code-block:: python

   # Find cells that might be referenced by named ranges
   for _, named_range in ranges.iterrows():
       range_formula = named_range['formula']
       range_name = named_range['name']
       
       # Simple parsing for demonstration
       if '.' in range_formula and ':' in range_formula:
           sheet_range = range_formula.split('.')
           if len(sheet_range) == 2:
               sheet_name = sheet_range[0]
               cell_range = sheet_range[1]
               print(f"Named range '{range_name}' covers {cell_range} in {sheet_name}")

Data Validation Analysis
------------------------

Understanding validation rules helps you understand data constraints:

.. code-block:: python

   from tidyxl import xlsx_validation

   # Get validation rules
   validation = xlsx_validation("data.xlsx")
   
   if len(validation) > 0:
       print(f"Found {len(validation)} validation rules")
       
       # Group by validation type
       by_type = validation.groupby('type').size()
       print("Validation types:")
       for vtype, count in by_type.items():
           print(f"  {vtype}: {count} rules")

       # Show list validations (dropdowns)
       lists = validation[validation['type'] == 'list']
       for _, rule in lists.iterrows():
           print(f"Dropdown in {rule['ref']}: {rule['formula1']}")

Working with Formatting
-----------------------

Extract and analyze Excel formatting:

.. code-block:: python

   from tidyxl import xlsx_formats

   # Get formatting information
   formats = xlsx_formats("data.xlsx")
   
   # Analyze fonts
   if formats['fonts']:
       print("Font analysis:")
       for i, font in enumerate(formats['fonts']):
           print(f"  Font {i}: {font['name']}, size {font['size']}")
           if font['bold']:
               print(f"    - Bold")
           if font['italic']:
               print(f"    - Italic")

   # Analyze fills (background colors)
   if formats['fills']:
       print(f"Found {len(formats['fills'])} fill patterns")

   # Analyze borders
   if formats['borders']:
       print(f"Found {len(formats['borders'])} border styles")

Complex Analysis Examples
-------------------------

Data Quality Assessment
~~~~~~~~~~~~~~~~~~~~~~~

Use tidyxl to assess Excel data quality:

.. code-block:: python

   def assess_data_quality(cells):
       """Comprehensive data quality assessment"""
       quality_report = {}
       
       # Overall statistics
       total_cells = len(cells)
       content_cells = len(cells[~cells['is_blank']])
       quality_report['coverage'] = content_cells / total_cells
       
       # Data type distribution
       type_dist = cells['data_type'].value_counts(normalize=True)
       quality_report['type_distribution'] = type_dist.to_dict()
       
       # Error detection
       error_cells = cells[cells['data_type'] == 'error']
       quality_report['error_count'] = len(error_cells)
       
       # Mixed type columns (potential issues)
       mixed_cols = []
       for col in cells['col'].unique():
           col_data = cells[
               (cells['col'] == col) & 
               (~cells['is_blank'])
           ]
           types = col_data['data_type'].nunique()
           if types > 1:
               mixed_cols.append({
                   'column': col,
                   'types': col_data['data_type'].unique().tolist()
               })
       
       quality_report['mixed_type_columns'] = mixed_cols
       
       return quality_report

   # Run assessment
   quality = assess_data_quality(cells)
   print("Data Quality Report:")
   print(f"  Coverage: {quality['coverage']:.1%}")
   print(f"  Errors: {quality['error_count']}")
   print(f"  Mixed-type columns: {len(quality['mixed_type_columns'])}")

Spreadsheet Structure Analysis
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Understand the structure of complex spreadsheets:

.. code-block:: python

   def analyze_structure(cells):
       """Analyze spreadsheet structure"""
       structure = {}
       
       # Sheet analysis
       for sheet in cells['sheet'].unique():
           sheet_data = cells[cells['sheet'] == sheet]
           content_data = sheet_data[~sheet_data['is_blank']]
           
           # Find data boundaries
           if len(content_data) > 0:
               min_row, max_row = content_data['row'].min(), content_data['row'].max()
               min_col, max_col = content_data['col'].min(), content_data['col'].max()
               
               # Estimate table structure
               row_1_data = content_data[content_data['row'] == 1]
               likely_headers = len(row_1_data[
                   row_1_data['data_type'] == 'character'
               ]) / len(row_1_data) if len(row_1_data) > 0 else 0
               
               structure[sheet] = {
                   'data_range': f"{min_row}:{max_row}, {min_col}:{max_col}",
                   'total_cells': len(content_data),
                   'likely_headers': likely_headers > 0.7,
                   'data_types': content_data['data_type'].value_counts().to_dict()
               }
       
       return structure

   # Analyze structure
   structure = analyze_structure(cells)
   for sheet, info in structure.items():
       print(f"\nSheet '{sheet}':")
       print(f"  Data range: {info['data_range']}")
       print(f"  Has headers: {info['likely_headers']}")
       print(f"  Cell count: {info['total_cells']}")

Converting Back to Tables
-------------------------

Sometimes you need to convert tidy data back to traditional tabular format:

.. code-block:: python

   def reconstruct_table(cells, sheet_name, has_headers=True):
       """Convert tidy format back to tabular format"""
       # Filter to specific sheet and non-blank cells
       sheet_cells = cells[
           (cells['sheet'] == sheet_name) & 
           (~cells['is_blank'])
       ]
       
       if len(sheet_cells) == 0:
           return None
       
       # Create pivot table
       table = sheet_cells.pivot_table(
           index='row',
           columns='col',
           values='content',
           aggfunc='first'
       ).fillna('')
       
       # Convert to regular DataFrame with proper column names
       if has_headers and len(table) > 0:
           # Use first row as column names
           headers = table.iloc[0].tolist()
           data_rows = table.iloc[1:]
           result = pd.DataFrame(data_rows.values, columns=headers)
           result.index = range(len(result))
           return result
       else:
           return table

   # Reconstruct a table
   sales_table = reconstruct_table(cells, 'Sales')
   if sales_table is not None:
       print("Reconstructed Sales table:")
       print(sales_table)

Performance Considerations
--------------------------

For Large Files
~~~~~~~~~~~~~~~

When working with large Excel files:

.. code-block:: python

   # Read only what you need
   specific_sheets = xlsx_cells("large_file.xlsx", sheets=["Summary", "Data"])
   
   # Exclude blank cells if not needed
   content_only = xlsx_cells("file.xlsx", include_blank_cells=False)
   
   # Process sheets individually
   for sheet_name in xlsx_sheet_names("large_file.xlsx"):
       sheet_data = xlsx_cells("large_file.xlsx", sheets=sheet_name)
       # Process each sheet separately to manage memory
       process_sheet_data(sheet_data)

Memory Management
~~~~~~~~~~~~~~~~~

For memory-efficient processing:

.. code-block:: python

   # Filter early to reduce memory usage
   cells = xlsx_cells("data.xlsx")
   important_cells = cells[
       (cells['sheet'].isin(['Sales', 'Products'])) &
       (~cells['is_blank'])
   ]
   
   # Clean up unneeded columns if memory is tight
   essential_columns = ['sheet', 'address', 'row', 'col', 'data_type', 'content']
   lean_data = cells[essential_columns].copy()

Best Practices
--------------

1. **Always check sheet names first** using ``xlsx_sheet_names()``
2. **Filter for content cells** early with ``~cells['is_blank']``
3. **Use appropriate typed columns** (numeric, character, etc.) for analysis
4. **Check for errors** by filtering ``data_type == 'error'``
5. **Understand your data structure** before analysis
6. **Use validation rules** to understand data constraints
7. **Consider memory usage** with large files

Troubleshooting
---------------

Common Issues
~~~~~~~~~~~~~

**Empty DataFrame returned:**
- Check that the file exists and is readable
- Verify sheet names with ``xlsx_sheet_names()``
- Try ``include_blank_cells=True`` to see all cells

**Unexpected data types:**
- Excel may store numbers as text - check the ``content`` column
- Dates might appear as numeric - check cell formatting
- Use ``data_type`` column to understand Excel's interpretation

**Memory issues:**
- Process sheets individually
- Filter for content cells only
- Use specific sheet selection

**Performance issues:**
- Avoid reading all sheets if only some are needed
- Consider excluding blank cells
- Process data in chunks for very large files

This comprehensive guide should help you master tidyxl for Excel data analysis. For more specific examples, see the :doc:`examples` section.