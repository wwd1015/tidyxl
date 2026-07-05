# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.9] - 2026-07-05

### Changed
- Refactored internals for structure and efficiency: removed the dead
  `tidyxl.core` module and extracted shared file-type and sheet-name
  validation helpers
- `xlsx_cells()` is ~30% faster on large workbooks: row/column dimension
  lookups are cached, date-format detection is memoized, and results are
  emitted pre-sorted instead of sorted after the fact
- `xlsx_cells()` only preserves VBA archives for `.xlsm` files, reducing
  memory use for plain `.xlsx` files
- `local_format_id` is now a stable per-workbook index of the cell's
  number format instead of a run-dependent identifier
- `__version__` now reflects the actual package version

### Fixed
- Workbooks are closed reliably in all readers, including on error paths,
  eliminating leaked file handles
- Fixed a crash in `examples/new_functions_demo.py` caused by invalid
  pandas Series negation

## [0.1.0] - 2025-07-30

### Added
- Initial release of tidyxl Python package
- `xlsx_cells()` function for extracting Excel cell data into tidy format
- `xlsx_sheet_names()` function for listing worksheet names
- `xlsx_names()` function for extracting named ranges and formulas
- `xlsx_validation()` function for extracting data validation rules
- `xlsx_formats()` function for extracting formatting information
- Complete compatibility with R tidyxl package API
- Support for Excel (.xlsx, .xlsm) files
- 23-column output structure matching R tidyxl exactly
- Comprehensive test suite with pytest
- Full documentation and examples
- Type hints and static analysis support

### Features
- **Tidy Format**: Each Excel cell becomes one row with comprehensive metadata
- **Multiple Data Types**: Separate columns for numeric, character, logical, date, and error values
- **Formula Support**: Extract and preserve Excel formulas
- **Named Ranges**: Access Excel named ranges and defined names
- **Data Validation**: Extract validation rules applied to cells
- **Multi-sheet Support**: Process all sheets or specify particular ones
- **Comments**: Extract cell comments and annotations
- **Formatting**: Access font, fill, border, and number format information
- **Flexible Input**: Handle various Excel layouts and structures
- **Error Handling**: Robust error handling for invalid files and parameters

### Technical Details
- Python 3.12+ support
- Built on openpyxl and pandas
- Comprehensive type hints
- 100% test coverage
- PyPI ready package structure
- MIT License

[Unreleased]: https://github.com/wwd1015/tidyxl/compare/v0.1.9...HEAD
[0.1.9]: https://github.com/wwd1015/tidyxl/compare/v0.1.8...v0.1.9
[0.1.0]: https://github.com/wwd1015/tidyxl/releases/tag/v0.1.0