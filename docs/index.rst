tidyxl: Tidy Excel Data Extraction for Python
============================================

.. image:: https://img.shields.io/pypi/v/tidyxl.svg
   :target: https://pypi.org/project/tidyxl/
   :alt: PyPI version

.. image:: https://img.shields.io/pypi/pyversions/tidyxl.svg
   :target: https://pypi.org/project/tidyxl/
   :alt: Python versions

.. image:: https://img.shields.io/badge/License-MIT-yellow.svg
   :target: https://opensource.org/licenses/MIT
   :alt: License: MIT

A Python package that imports Excel files (.xlsx, .xlsm) into a tidy format where each cell is represented as a single row with detailed metadata. This package is inspired by and closely mirrors the functionality of the `R tidyxl package <https://nacnudus.github.io/tidyxl/>`_.

Key Features
------------

* **Tidy Format**: Each Excel cell becomes one row with comprehensive metadata
* **Complete Cell Information**: Extract values, formulas, formatting, comments, and more
* **Multiple Worksheets**: Process all sheets or specify particular ones
* **Named Ranges**: Extract and analyze Excel named ranges and formulas
* **Data Validation**: Discover data validation rules applied to cells
* **Type Safety**: Separate columns for different data types (numeric, character, logical, date, error)
* **R tidyxl Compatible**: Identical API and output structure to the R package

Quick Start
-----------

Installation
~~~~~~~~~~~~

.. code-block:: bash

   pip install tidyxl

Basic Usage
~~~~~~~~~~~

.. code-block:: python

   from tidyxl import xlsx_cells, xlsx_sheet_names

   # Get all sheet names
   sheets = xlsx_sheet_names("data.xlsx")
   print(f"Found sheets: {sheets}")

   # Read all cells from all sheets
   cells = xlsx_cells("data.xlsx")
   print(f"Total cells: {len(cells)}")

   # Filter for cells with content
   content_cells = cells[~cells['is_blank']]
   print(content_cells[['sheet', 'address', 'data_type', 'character', 'numeric']].head())

Contents
--------

.. toctree::
   :maxdepth: 2
   :caption: User Guide

   installation
   quickstart
   user_guide
   examples

.. toctree::
   :maxdepth: 2
   :caption: API Reference

   api

.. toctree::
   :maxdepth: 1
   :caption: Development

   changelog
   contributing

Comparison with Other Libraries
-------------------------------

.. list-table::
   :header-rows: 1
   :widths: 20 15 20 15

   * - Feature
     - tidyxl
     - pandas.read_excel()
     - openpyxl
   * - Tidy format
     - ✅
     - ❌
     - ❌
   * - Preserve formulas
     - ✅
     - ❌
     - ✅
   * - Cell-level metadata
     - ✅
     - ❌
     - ✅
   * - Handle messy layouts
     - ✅
     - ❌
     - ✅
   * - Named ranges
     - ✅
     - ❌
     - ✅
   * - Data validation
     - ✅
     - ❌
     - ✅
   * - Formatting details
     - ✅
     - ❌
     - ✅
   * - Ease of use
     - ✅
     - ✅
     - ❌

Acknowledgments
---------------

* Inspired by the excellent `R tidyxl package <https://nacnudus.github.io/tidyxl/>`_ by Duncan Garmonsway
* Built on top of `openpyxl <https://openpyxl.readthedocs.io/>`_ for Excel file processing
* Uses `pandas <https://pandas.pydata.org/>`_ for data structure management

License
-------

This project is licensed under the MIT License - see the :doc:`License <license>` for details.

Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`