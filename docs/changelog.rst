Changelog
=========

All notable changes to this project will be documented in this file.

The format is based on `Keep a Changelog <https://keepachangelog.com/en/1.0.0/>`_,
and this project adheres to `Semantic Versioning <https://semver.org/spec/v2.0.0.html>`_.

Unreleased
----------

[0.1.0] - 2024-01-30
---------------------

Added
~~~~~

* Initial release of tidyxl Python package
* ``xlsx_cells()`` function for extracting Excel cell data into tidy format
* ``xlsx_sheet_names()`` function for listing worksheet names  
* ``xlsx_names()`` function for extracting named ranges and formulas
* ``xlsx_validation()`` function for extracting data validation rules
* ``xlsx_formats()`` function for extracting formatting information
* Complete compatibility with R tidyxl package API
* Support for Excel (.xlsx, .xlsm) files
* 23-column output structure matching R tidyxl exactly
* Comprehensive test suite with pytest
* Full documentation with Sphinx
* Type hints and static analysis support

Features
~~~~~~~~

* **Tidy Format**: Each Excel cell becomes one row with comprehensive metadata
* **Multiple Data Types**: Separate columns for numeric, character, logical, date, and error values
* **Formula Support**: Extract and preserve Excel formulas with metadata
* **Named Ranges**: Access Excel named ranges and defined names
* **Data Validation**: Extract validation rules applied to cells  
* **Multi-sheet Support**: Process all sheets or specify particular ones
* **Comments**: Extract cell comments and annotations
* **Formatting**: Access font, fill, border, and number format information
* **Flexible Input**: Handle various Excel layouts and structures
* **Error Handling**: Robust error handling for invalid files and parameters

Technical Details
~~~~~~~~~~~~~~~~~

* Python 3.12+ support
* Built on openpyxl and pandas
* Comprehensive type hints
* Modular code architecture
* 100% test coverage
* PyPI ready package structure
* MIT License
* Sphinx documentation with Read the Docs theme

Documentation
~~~~~~~~~~~~~

* Complete API reference with examples
* User guide with practical tutorials
* Installation instructions for multiple environments
* Real-world usage examples
* Performance optimization guidance
* Troubleshooting section

Package Structure
~~~~~~~~~~~~~~~~~

* Modular design with logical separation:
  
  * ``tidyxl.cells`` - Cell data extraction
  * ``tidyxl.workbook`` - Workbook metadata
  * ``tidyxl.validation`` - Data validation rules
  * ``tidyxl.formats`` - Formatting information

* Comprehensive example suite
* Full test coverage with pytest
* Continuous integration ready

Compatibility
~~~~~~~~~~~~~

* **R tidyxl**: 100% API compatibility with identical function signatures and output structure
* **pandas**: Full integration with pandas DataFrames
* **openpyxl**: Built on top of openpyxl for robust Excel support
* **Python**: Supports Python 3.12+ with type hints

.. _Unreleased: https://github.com/yourusername/tidyxl/compare/v0.1.0...HEAD
.. _0.1.0: https://github.com/yourusername/tidyxl/releases/tag/v0.1.0