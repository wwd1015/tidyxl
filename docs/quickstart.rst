Quick Start Guide
=================

This guide will get you up and running with tidyxl in minutes.

Basic Usage
-----------

Let's start with a simple example. First, import the main functions:

.. code-block:: python

   from tidyxl import xlsx_cells, xlsx_sheet_names

Exploring an Excel File
~~~~~~~~~~~~~~~~~~~~~~~

Before reading cell data, it's often helpful to see what sheets are available:

.. code-block:: python

   # List all worksheets
   sheets = xlsx_sheet_names("data.xlsx")
   print(f"Available sheets: {sheets}")
   # Output: Available sheets: ['Sales', 'Products', 'Summary']

Reading Cell Data
~~~~~~~~~~~~~~~~~

The main function is ``xlsx_cells()``, which reads Excel files into a tidy format:

.. code-block:: python

   # Read all cells from all sheets
   cells = xlsx_cells("data.xlsx")
   print(f"Total cells: {len(cells)}")
   print(f"Columns: {list(cells.columns)}")

   # Look at first few cells
   print(cells.head())

The result is a pandas DataFrame where each row represents a single Excel cell with comprehensive metadata.

Filtering Data
~~~~~~~~~~~~~~

One of the most common operations is filtering for cells that contain actual data:

.. code-block:: python

   # Get only cells with content (not blank)
   content_cells = cells[~cells['is_blank']]
   print(f"Cells with content: {len(content_cells)}")

   # Show basic cell information
   print(content_cells[['sheet', 'address', 'data_type', 'content']].head(10))

Working with Specific Sheets
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

You can read specific sheets to focus your analysis:

.. code-block:: python

   # Read only the Sales sheet
   sales_cells = xlsx_cells("data.xlsx", sheets="Sales")

   # Read multiple specific sheets
   important_sheets = xlsx_cells("data.xlsx", sheets=["Sales", "Products"])

Understanding Data Types
~~~~~~~~~~~~~~~~~~~~~~~~

tidyxl preserves the original data types and provides typed columns:

.. code-block:: python

   # See what data types are present
   print("Data types found:")
   print(cells['data_type'].value_counts())

   # Access typed values
   numeric_cells = cells[cells['data_type'] == 'numeric']
   print(f"Numeric values: {numeric_cells['numeric'].dropna().tolist()}")

   text_cells = cells[cells['data_type'] == 'character']
   print(f"Text values: {text_cells['character'].dropna().tolist()}")

   # Boolean values
   boolean_cells = cells[cells['data_type'] == 'logical']
   print(f"Boolean values: {boolean_cells['logical'].dropna().tolist()}")

Working with Formulas
~~~~~~~~~~~~~~~~~~~~~

tidyxl preserves Excel formulas, making it easy to analyze spreadsheet logic:

.. code-block:: python

   # Find all cells with formulas
   formula_cells = cells[cells['formula'].notna()]
   print(f"Found {len(formula_cells)} formulas:")

   for _, cell in formula_cells.iterrows():
       print(f"  {cell['address']}: {cell['formula']}")

Advanced Functions
------------------

Named Ranges
~~~~~~~~~~~~

Extract information about Excel named ranges:

.. code-block:: python

   from tidyxl import xlsx_names

   names = xlsx_names("data.xlsx")
   print("Named ranges:")
   print(names[['name', 'formula', 'sheet', 'is_range']])

Data Validation
~~~~~~~~~~~~~~~

Discover data validation rules applied to cells:

.. code-block:: python

   from tidyxl import xlsx_validation

   validation = xlsx_validation("data.xlsx")
   if len(validation) > 0:
       print("Validation rules:")
       print(validation[['sheet', 'ref', 'type', 'formula1']])

Formatting Information
~~~~~~~~~~~~~~~~~~~~~~

Extract formatting details:

.. code-block:: python

   from tidyxl import xlsx_formats

   formats = xlsx_formats("data.xlsx")
   print("Formatting information:")
   for category, items in formats.items():
       print(f"  {category}: {len(items)} entries")

Common Patterns
---------------

Finding Headers
~~~~~~~~~~~~~~~

Headers are typically in the first row:

.. code-block:: python

   # Find headers (row 1)
   headers = cells[(cells['row'] == 1) & (~cells['is_blank'])]
   header_names = headers['character'].dropna().tolist()
   print(f"Headers: {header_names}")

Analyzing Data Structure
~~~~~~~~~~~~~~~~~~~~~~~~

Understand the structure of your data:

.. code-block:: python

   # Data summary by sheet
   summary = cells.groupby('sheet').agg({
       'address': 'count',  # Total cells
       'is_blank': lambda x: (~x).sum(),  # Cells with content
       'data_type': lambda x: x.value_counts().to_dict()  # Type distribution
   }).rename(columns={'address': 'total_cells', 'is_blank': 'content_cells'})

   print(summary)

Finding Comments
~~~~~~~~~~~~~~~~

Locate cells with comments:

.. code-block:: python

   # Find all cells with comments
   commented_cells = cells[cells['comment'].notna()]
   print(f"Found {len(commented_cells)} cells with comments:")

   for _, cell in commented_cells.iterrows():
       print(f"  {cell['address']}: {cell['comment']}")

Converting to Tabular Format
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sometimes you need to convert the tidy format back to a traditional table:

.. code-block:: python

   def tidy_to_table(cells_df, sheet_name):
       """Convert tidy format back to tabular format"""
       sheet_cells = cells_df[
           (cells_df['sheet'] == sheet_name) & 
           (~cells_df['is_blank'])
       ]
       
       return sheet_cells.pivot_table(
           index='row',
           columns='col',
           values='content',
           aggfunc='first'
       ).fillna('')

   # Convert Sales sheet back to table format
   sales_table = tidy_to_table(cells, 'Sales')
   print(sales_table)

Error Handling
--------------

Handle common issues gracefully:

.. code-block:: python

   try:
       cells = xlsx_cells("data.xlsx")
   except FileNotFoundError:
       print("Excel file not found!")
   except ValueError as e:
       print(f"Invalid file or parameter: {e}")

   # Check if specific sheet exists
   available_sheets = xlsx_sheet_names("data.xlsx")
   if "NonExistent" not in available_sheets:
       print(f"Sheet not found. Available: {available_sheets}")

Performance Tips
----------------

For large Excel files, consider these optimization strategies:

.. code-block:: python

   # Read only specific sheets to reduce memory usage
   important_data = xlsx_cells("large_file.xlsx", sheets=["Summary", "Results"])

   # Exclude blank cells if you don't need them
   content_only = xlsx_cells("data.xlsx", include_blank_cells=False)

   # Process data in chunks if working with very large datasets
   for sheet in xlsx_sheet_names("data.xlsx"):
       sheet_data = xlsx_cells("data.xlsx", sheets=sheet)
       # Process each sheet individually
       process_sheet(sheet_data)

Next Steps
----------

Now that you understand the basics, explore:

* :doc:`user_guide` - Comprehensive guide with detailed examples
* :doc:`examples` - Real-world use cases and advanced techniques
* :doc:`api` - Complete API reference
* The ``examples/complete_demo.py`` file in the package for a comprehensive demonstration