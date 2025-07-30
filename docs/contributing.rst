Contributing
============

We welcome contributions to tidyxl! This document provides guidelines for contributing to the project.

Getting Started
---------------

Development Setup
~~~~~~~~~~~~~~~~~

1. Fork the repository on GitHub
2. Clone your fork locally:

   .. code-block:: bash

      git clone https://github.com/yourusername/tidyxl.git
      cd tidyxl

3. Create a virtual environment and install development dependencies:

   .. code-block:: bash

      python -m venv venv
      source venv/bin/activate  # On Windows: venv\Scripts\activate
      pip install -e ".[dev,test,docs]"

4. Create a branch for your changes:

   .. code-block:: bash

      git checkout -b feature/your-feature-name

Development Workflow
~~~~~~~~~~~~~~~~~~~~

1. Make your changes
2. Run tests to ensure everything works:

   .. code-block:: bash

      pytest

3. Run code quality checks:

   .. code-block:: bash

      black .
      ruff check .
      mypy tidyxl

4. Add or update tests as needed
5. Update documentation if necessary
6. Commit your changes and push to your fork
7. Create a pull request

Code Style
----------

We use several tools to maintain code quality:

Formatting
~~~~~~~~~~

* **Black** for code formatting
* **Ruff** for linting and import sorting
* **MyPy** for type checking

Run these before committing:

.. code-block:: bash

   black tidyxl tests examples
   ruff check tidyxl tests examples --fix
   mypy tidyxl

Type Hints
~~~~~~~~~~

All new code should include type hints:

.. code-block:: python

   def xlsx_cells(
       path: str,
       sheets: Optional[Union[str, List[str]]] = None,
       check_filetype: bool = True,
       include_blank_cells: bool = True
   ) -> pd.DataFrame:
       """Function with proper type hints."""
       pass

Documentation
~~~~~~~~~~~~~

* Use Google-style docstrings
* Include examples in docstrings where helpful
* Update API documentation when adding new functions

.. code-block:: python

   def new_function(param1: str, param2: int = 10) -> Dict[str, Any]:
       """
       Brief description of the function.
       
       Longer description with more details about what the function does,
       when to use it, and any important considerations.
       
       Parameters
       ----------
       param1 : str
           Description of the first parameter
       param2 : int, optional
           Description of the second parameter, by default 10
           
       Returns
       -------
       Dict[str, Any]
           Description of what is returned
           
       Examples
       --------
       >>> result = new_function("hello", 20)
       >>> print(result)
       {'message': 'hello', 'count': 20}
       """
       pass

Testing
-------

Test Requirements
~~~~~~~~~~~~~~~~~

* All new functionality must have tests
* Tests should cover both success and error cases
* Use meaningful test names that describe what is being tested
* Aim for high test coverage

Test Structure
~~~~~~~~~~~~~~

Tests are organized in the ``tests/`` directory:

* ``tests/conftest.py`` - Shared fixtures
* ``tests/test_xlsx_cells.py`` - Tests for cell extraction
* ``tests/test_other_functions.py`` - Tests for other functions

Writing Tests
~~~~~~~~~~~~~

Use pytest fixtures and meaningful test names:

.. code-block:: python

   def test_xlsx_cells_basic_functionality(sample_excel_file):
       """Test that xlsx_cells returns correct data structure."""
       cells = xlsx_cells(sample_excel_file)
       
       assert isinstance(cells, pd.DataFrame)
       assert len(cells) > 0
       assert 'sheet' in cells.columns
       assert 'address' in cells.columns

   def test_xlsx_cells_handles_missing_file():
       """Test that xlsx_cells raises appropriate error for missing file."""
       with pytest.raises(FileNotFoundError):
           xlsx_cells("nonexistent_file.xlsx")

Running Tests
~~~~~~~~~~~~~

Run the full test suite:

.. code-block:: bash

   pytest

Run specific tests:

.. code-block:: bash

   pytest tests/test_xlsx_cells.py::TestXlsxCells::test_basic_functionality

Run with coverage:

.. code-block:: bash

   pytest --cov=tidyxl --cov-report=html

Documentation
-------------

Building Documentation
~~~~~~~~~~~~~~~~~~~~~~

The documentation is built with Sphinx:

.. code-block:: bash

   cd docs
   make html

The built documentation will be in ``docs/_build/html/``.

Documentation Standards
~~~~~~~~~~~~~~~~~~~~~~~

* Use reStructuredText format
* Include practical examples
* Cross-reference related functions
* Keep language clear and concise

Adding Examples
~~~~~~~~~~~~~~~

When adding new examples:

1. Add them to the appropriate ``.rst`` file in ``docs/``
2. Test the examples to ensure they work
3. Include expected output where helpful

Pull Request Process
--------------------

PR Requirements
~~~~~~~~~~~~~~~

Before submitting a pull request:

1. Ensure all tests pass
2. Run code quality checks
3. Update documentation if needed
4. Add tests for new functionality
5. Update CHANGELOG.md

PR Description
~~~~~~~~~~~~~~

Include in your PR description:

* What changes were made
* Why the changes were necessary
* Any breaking changes
* How to test the changes

Example PR description:

.. code-block:: text

   ## Summary
   
   Add support for reading Excel files with password protection.
   
   ## Changes Made
   
   - Added `password` parameter to `xlsx_cells()` function
   - Updated openpyxl integration to handle password-protected files
   - Added tests for password-protected files
   - Updated documentation with password examples
   
   ## Testing
   
   - All existing tests pass
   - Added new tests in `test_password_protection.py`
   - Tested with various password-protected Excel files

Review Process
~~~~~~~~~~~~~~

1. Maintainers will review your PR
2. Address any feedback or requested changes
3. Once approved, your PR will be merged

Types of Contributions
----------------------

Bug Reports
~~~~~~~~~~~

When reporting bugs:

1. Use the GitHub issue tracker
2. Include a clear title and description
3. Provide steps to reproduce
4. Include relevant code and error messages
5. Specify your environment (Python version, OS, etc.)

Feature Requests
~~~~~~~~~~~~~~~~

For new features:

1. Create an issue to discuss the feature first
2. Explain the use case and benefits
3. Consider if it fits with the project's goals
4. Be open to alternative approaches

Code Contributions
~~~~~~~~~~~~~~~~~~

Areas where contributions are welcome:

* Bug fixes
* Performance improvements
* New functionality (with prior discussion)
* Documentation improvements
* Test coverage improvements
* Example additions

Documentation Contributions
~~~~~~~~~~~~~~~~~~~~~~~~~~~

Help improve documentation by:

* Fixing typos or unclear explanations
* Adding examples
* Improving API documentation
* Creating tutorials

Community Guidelines
--------------------

Code of Conduct
~~~~~~~~~~~~~~~

* Be respectful and inclusive
* Welcome newcomers and help them get started
* Focus on constructive feedback
* Respect different viewpoints and experiences

Communication
~~~~~~~~~~~~~

* Use GitHub issues for bug reports and feature requests
* Use pull request discussions for code-related questions
* Be patient - maintainers volunteer their time

Recognition
~~~~~~~~~~~

Contributors are recognized in:

* The ``CONTRIBUTORS.md`` file
* Release notes for significant contributions
* Git commit history

Getting Help
------------

If you need help:

1. Check the documentation first
2. Search existing GitHub issues
3. Create a new issue with your question
4. Be specific about what you're trying to do

Thank you for contributing to tidyxl!