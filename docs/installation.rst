Installation
============

Requirements
------------

tidyxl requires Python 3.12 or later and has the following dependencies:

* `pandas <https://pandas.pydata.org/>`_ >= 2.0.0 - For data structure management
* `openpyxl <https://openpyxl.readthedocs.io/>`_ >= 3.1.0 - For Excel file processing

Installing from PyPI
---------------------

The simplest way to install tidyxl is from PyPI using pip:

.. code-block:: bash

   pip install tidyxl

This will install tidyxl and all its dependencies.

Installing from Source
----------------------

If you want to install the latest development version or contribute to the project, you can install from source:

.. code-block:: bash

   git clone https://github.com/yourusername/tidyxl.git
   cd tidyxl
   pip install -e .

Development Installation
------------------------

For development, install with additional development dependencies:

.. code-block:: bash

   git clone https://github.com/yourusername/tidyxl.git
   cd tidyxl
   pip install -e ".[dev,test,docs]"

This installs tidyxl in editable mode along with:

* **dev**: Development tools (black, ruff, mypy, build, twine)
* **test**: Testing tools (pytest, pytest-cov)
* **docs**: Documentation tools (sphinx, sphinx-rtd-theme, myst-parser)

Verifying Installation
----------------------

To verify that tidyxl is installed correctly, run:

.. code-block:: python

   import tidyxl
   print(tidyxl.__version__)

You should see the version number printed without any errors.

Conda Installation
------------------

tidyxl is not currently available on conda-forge, but you can install it in a conda environment using pip:

.. code-block:: bash

   conda create -n tidyxl-env python=3.12
   conda activate tidyxl-env
   pip install tidyxl

Troubleshooting
---------------

Common Installation Issues
~~~~~~~~~~~~~~~~~~~~~~~~~~

**ImportError: No module named 'openpyxl'**

This usually means openpyxl wasn't installed correctly. Try:

.. code-block:: bash

   pip install --upgrade openpyxl

**ModuleNotFoundError: No module named 'pandas'**

This means pandas is missing. Install it with:

.. code-block:: bash

   pip install --upgrade pandas

**Permission Errors on macOS/Linux**

If you get permission errors, try installing in user space:

.. code-block:: bash

   pip install --user tidyxl

Or use a virtual environment (recommended):

.. code-block:: bash

   python -m venv tidyxl-env
   source tidyxl-env/bin/activate  # On Windows: tidyxl-env\Scripts\activate
   pip install tidyxl

**Python Version Issues**

tidyxl requires Python 3.12+. Check your Python version:

.. code-block:: bash

   python --version

If you have an older version, consider using `pyenv <https://github.com/pyenv/pyenv>`_ to manage multiple Python versions.

Virtual Environments
---------------------

We strongly recommend using virtual environments to avoid dependency conflicts:

Using venv (built-in)
~~~~~~~~~~~~~~~~~~~~~~

.. code-block:: bash

   python -m venv tidyxl-env
   source tidyxl-env/bin/activate  # On Windows: tidyxl-env\Scripts\activate
   pip install tidyxl

Using conda
~~~~~~~~~~~

.. code-block:: bash

   conda create -n tidyxl-env python=3.12
   conda activate tidyxl-env
   pip install tidyxl

Using pipenv
~~~~~~~~~~~~

.. code-block:: bash

   pipenv install tidyxl
   pipenv shell

Next Steps
----------

Once tidyxl is installed, check out the :doc:`quickstart` guide to begin using the package.