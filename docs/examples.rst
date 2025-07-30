Examples
========

This section provides practical examples of using tidyxl for real-world Excel analysis tasks.

Basic Examples
--------------

Simple Data Extraction
~~~~~~~~~~~~~~~~~~~~~~

Start with a basic Excel file containing sales data:

.. code-block:: python

   from tidyxl import xlsx_cells, xlsx_sheet_names
   import pandas as pd

   # First, explore the file structure
   sheets = xlsx_sheet_names("sales_data.xlsx")
   print(f"Available sheets: {sheets}")

   # Read all cell data
   cells = xlsx_cells("sales_data.xlsx")
   print(f"Total cells: {len(cells)}")

   # Filter for actual content
   content = cells[~cells['is_blank']]
   print(f"Cells with data: {len(content)}")

   # Show the first few cells
   print(content[['sheet', 'address', 'data_type', 'content']].head())

Finding and Analyzing Formulas
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Discover how calculations are structured in a spreadsheet:

.. code-block:: python

   # Find all formulas
   formulas = cells[cells['formula'].notna()]
   
   print(f"Found {len(formulas)} formulas:")
   for _, cell in formulas.iterrows():
       print(f"  {cell['address']}: {cell['formula']}")

   # Analyze formula patterns
   sum_formulas = formulas[formulas['formula'].str.startswith('=SUM', na=False)]
   avg_formulas = formulas[formulas['formula'].str.startswith('=AVERAGE', na=False)]
   
   print(f"SUM formulas: {len(sum_formulas)}")
   print(f"AVERAGE formulas: {len(avg_formulas)}")

Data Quality Assessment
-----------------------

Identifying Data Issues
~~~~~~~~~~~~~~~~~~~~~~

Use tidyxl to find potential data quality problems:

.. code-block:: python

   def find_data_issues(cells):
       """Identify common data quality issues"""
       issues = []
       
       # Find error cells
       errors = cells[cells['data_type'] == 'error']
       if len(errors) > 0:
           issues.append(f"Found {len(errors)} error cells")
           for _, cell in errors.iterrows():
               issues.append(f"  Error in {cell['address']}: {cell['error']}")
       
       # Find mixed data types in columns
       for col in cells['col'].unique():
           col_data = cells[
               (cells['col'] == col) & 
               (~cells['is_blank']) & 
               (cells['row'] > 1)  # Skip header row
           ]
           
           if len(col_data) > 1:
               types = col_data['data_type'].unique()
               if len(types) > 1:
                   issues.append(f"Column {col} has mixed types: {types}")
       
       # Find suspiciously long text (potential data entry errors)
       long_text = cells[
           (cells['data_type'] == 'character') & 
           (cells['character'].str.len() > 100)
       ]
       if len(long_text) > 0:
           issues.append(f"Found {len(long_text)} cells with very long text")
       
       return issues

   # Check for issues
   issues = find_data_issues(cells)
   if issues:
       print("Data quality issues found:")
       for issue in issues:
           print(f"  - {issue}")
   else:
       print("No obvious data quality issues detected")

Detecting Inconsistent Formatting
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Find cells that might have formatting inconsistencies:

.. code-block:: python

   # Find numeric values stored as text
   numeric_as_text = cells[
       (cells['data_type'] == 'character') & 
       (cells['character'].str.match(r'^-?\d+\.?\d*$', na=False))
   ]
   
   if len(numeric_as_text) > 0:
       print(f"Found {len(numeric_as_text)} numeric values stored as text:")
       for _, cell in numeric_as_text.head().iterrows():
           print(f"  {cell['address']}: '{cell['character']}'")

   # Find potential date values stored as text
   potential_dates = cells[
       (cells['data_type'] == 'character') & 
       (cells['character'].str.match(r'\d{1,2}/\d{1,2}/\d{4}', na=False))
   ]
   
   if len(potential_dates) > 0:
       print(f"Found {len(potential_dates)} potential dates stored as text:")
       for _, cell in potential_dates.head().iterrows():
           print(f"  {cell['address']}: '{cell['character']}'")

Complex Analysis Examples
-------------------------

Financial Statement Analysis
~~~~~~~~~~~~~~~~~~~~~~~~~~~

Analyze a financial statement with multiple sections:

.. code-block:: python

   def analyze_financial_statement(cells):
       """Analyze financial statement structure"""
       
       # Find section headers (typically bold or in specific positions)
       content_cells = cells[~cells['is_blank']]
       
       # Look for common financial statement terms
       financial_terms = [
           'Revenue', 'Sales', 'Income', 'Expenses', 'Assets', 
           'Liabilities', 'Equity', 'Cash Flow'
       ]
       
       sections = {}
       for term in financial_terms:
           matches = content_cells[
               content_cells['character'].str.contains(term, case=False, na=False)
           ]
           if len(matches) > 0:
               sections[term] = matches[['address', 'character']].to_dict('records')
       
       # Find numeric values (amounts)
       amounts = content_cells[
           (content_cells['data_type'] == 'numeric') & 
           (content_cells['numeric'].abs() > 1000)  # Significant amounts
       ]
       
       return {
           'sections': sections,
           'total_amounts': len(amounts),
           'amount_range': {
               'min': amounts['numeric'].min() if len(amounts) > 0 else None,
               'max': amounts['numeric'].max() if len(amounts) > 0 else None
           }
       }

   # Analyze financial data
   fin_analysis = analyze_financial_statement(cells)
   print("Financial Statement Analysis:")
   print(f"Sections found: {list(fin_analysis['sections'].keys())}")
   print(f"Numeric amounts: {fin_analysis['total_amounts']}")

Survey Data Processing
~~~~~~~~~~~~~~~~~~~~~

Process survey data with multiple question types:

.. code-block:: python

   def process_survey_data(cells, sheet_name):
       """Process survey responses from Excel"""
       
       survey_data = cells[cells['sheet'] == sheet_name]
       content = survey_data[~survey_data['is_blank']]
       
       # Assume first row contains questions
       questions = content[content['row'] == 1]
       question_map = {
           q['col']: q['character'] 
           for _, q in questions.iterrows() 
           if q['data_type'] == 'character'
       }
       
       # Get response data (rows 2+)
       responses = content[content['row'] > 1]
       
       # Analyze response patterns
       analysis = {}
       for col, question in question_map.items():
           col_responses = responses[responses['col'] == col]
           
           if len(col_responses) > 0:
               # Determine question type based on responses
               response_types = col_responses['data_type'].unique()
               
               if 'numeric' in response_types:
                   # Numeric question (rating scale, etc.)
                   numeric_responses = col_responses[
                       col_responses['data_type'] == 'numeric'
                   ]['numeric']
                   analysis[question] = {
                       'type': 'numeric',
                       'responses': len(numeric_responses),
                       'mean': numeric_responses.mean(),
                       'range': [numeric_responses.min(), numeric_responses.max()]
                   }
               else:
                   # Categorical question
                   text_responses = col_responses[
                       col_responses['data_type'] == 'character'
                   ]['character']
                   analysis[question] = {
                       'type': 'categorical',
                       'responses': len(text_responses),
                       'categories': text_responses.value_counts().to_dict()
                   }
       
       return analysis

   # Process survey data
   if 'Survey' in cells['sheet'].unique():
       survey_analysis = process_survey_data(cells, 'Survey')
       print("Survey Analysis:")
       for question, stats in survey_analysis.items():
           print(f"\nQ: {question}")
           print(f"   Type: {stats['type']}")
           print(f"   Responses: {stats['responses']}")

Named Ranges and Validation
---------------------------

Working with Complex Spreadsheet Models
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Analyze spreadsheets with named ranges and data validation:

.. code-block:: python

   from tidyxl import xlsx_names, xlsx_validation

   def analyze_spreadsheet_model(file_path):
       """Comprehensive analysis of Excel model"""
       
       # Get all components
       cells = xlsx_cells(file_path)
       names = xlsx_names(file_path)
       validation = xlsx_validation(file_path)
       
       analysis = {
           'sheets': xlsx_sheet_names(file_path),
           'total_cells': len(cells),
           'content_cells': len(cells[~cells['is_blank']]),
           'formulas': len(cells[cells['formula'].notna()]),
           'named_ranges': len(names),
           'validation_rules': len(validation)
       }
       
       # Analyze named ranges
       if len(names) > 0:
           ranges = names[names['is_range'] == True]
           formulas = names[names['is_range'] == False]
           analysis['named_ranges_breakdown'] = {
               'cell_ranges': len(ranges),
               'named_formulas': len(formulas)
           }
       
       # Analyze validation rules
       if len(validation) > 0:
           val_types = validation['type'].value_counts().to_dict()
           analysis['validation_types'] = val_types
       
       # Find input vs calculation cells
       input_cells = cells[
           (~cells['is_blank']) & 
           (cells['formula'].isna()) &
           (cells['data_type'].isin(['numeric', 'character']))
       ]
       calc_cells = cells[cells['formula'].notna()]
       
       analysis['cell_classification'] = {
           'input_cells': len(input_cells),
           'calculation_cells': len(calc_cells),
           'ratio': len(calc_cells) / len(input_cells) if len(input_cells) > 0 else 0
       }
       
       return analysis

   # Analyze a complex model
   model_analysis = analyze_spreadsheet_model("financial_model.xlsx")
   print("Spreadsheet Model Analysis:")
   for key, value in model_analysis.items():
       print(f"  {key}: {value}")

Data Validation Analysis
~~~~~~~~~~~~~~~~~~~~~~~

Understand data entry constraints:

.. code-block:: python

   def analyze_data_validation(validation_df):
       """Analyze data validation rules"""
       
       if len(validation_df) == 0:
           return "No validation rules found"
       
       analysis = {}
       
       # Group by validation type
       for val_type in validation_df['type'].unique():
           type_rules = validation_df[validation_df['type'] == val_type]
           
           analysis[val_type] = {
               'count': len(type_rules),
               'sheets': type_rules['sheet'].unique().tolist(),
               'examples': []
           }
           
           # Add examples based on type
           for _, rule in type_rules.head(3).iterrows():
               example = {
                   'ref': rule['ref'],
                   'criteria': rule['formula1']
               }
               if rule['formula2']:
                   example['criteria'] += f" to {rule['formula2']}"
               analysis[val_type]['examples'].append(example)
       
       return analysis

   # Analyze validation rules
   if len(validation) > 0:
       val_analysis = analyze_data_validation(validation)
       print("Data Validation Analysis:")
       for val_type, info in val_analysis.items():
           print(f"\n{val_type.upper()} validation:")
           print(f"  Rules: {info['count']}")
           print(f"  Sheets: {', '.join(info['sheets'])}")
           for example in info['examples']:
               print(f"  Example: {example['ref']} - {example['criteria']}")

Advanced Techniques
-------------------

Building Data Lineage
~~~~~~~~~~~~~~~~~~~~~

Trace how data flows through a spreadsheet:

.. code-block:: python

   def build_data_lineage(cells):
       """Build a map of data dependencies"""
       
       formulas = cells[cells['formula'].notna()]
       lineage = {}
       
       for _, cell in formulas.iterrows():
           cell_ref = f"{cell['sheet']}.{cell['address']}"
           formula = cell['formula']
           
           # Simple dependency parsing (for demonstration)
           # In practice, you'd want more sophisticated parsing
           dependencies = []
           
           # Find cell references in formula
           import re
           cell_pattern = r'[A-Z]+\d+'
           references = re.findall(cell_pattern, formula)
           
           for ref in references:
               if ref != cell['address']:  # Don't include self-reference
                   dependencies.append(f"{cell['sheet']}.{ref}")
           
           # Find sheet references
           sheet_pattern = r'(\w+)\.([A-Z]+\d+)'
           sheet_refs = re.findall(sheet_pattern, formula)
           for sheet, ref in sheet_refs:
               dependencies.append(f"{sheet}.{ref}")
           
           lineage[cell_ref] = {
               'formula': formula,
               'depends_on': list(set(dependencies))
           }
       
       return lineage

   # Build lineage map
   lineage = build_data_lineage(cells)
   print("Data Lineage (sample):")
   for cell, info in list(lineage.items())[:5]:
       print(f"\n{cell}:")
       print(f"  Formula: {info['formula']}")
       print(f"  Depends on: {info['depends_on']}")

Automated Report Generation
~~~~~~~~~~~~~~~~~~~~~~~~~~

Generate summary reports from Excel data:

.. code-block:: python

   def generate_excel_report(file_path):
       """Generate comprehensive Excel file report"""
       
       # Gather all data
       sheets = xlsx_sheet_names(file_path)
       cells = xlsx_cells(file_path)
       names = xlsx_names(file_path)
       validation = xlsx_validation(file_path)
       
       report = []
       report.append(f"Excel File Analysis Report")
       report.append(f"=" * 40)
       report.append(f"File: {file_path}")
       report.append(f"Sheets: {len(sheets)} ({', '.join(sheets)})")
       report.append("")
       
       # Overall statistics
       total_cells = len(cells)
       content_cells = len(cells[~cells['is_blank']])
       report.append(f"Cell Statistics:")
       report.append(f"  Total cells: {total_cells:,}")
       report.append(f"  Cells with content: {content_cells:,}")
       report.append(f"  Coverage: {content_cells/total_cells:.1%}")
       report.append("")
       
       # Data type breakdown
       type_counts = cells['data_type'].value_counts()
       report.append(f"Data Types:")
       for dtype, count in type_counts.items():
           report.append(f"  {dtype}: {count:,} ({count/total_cells:.1%})")
       report.append("")
       
       # Formula analysis
       formulas = cells[cells['formula'].notna()]
       report.append(f"Formulas: {len(formulas)}")
       if len(formulas) > 0:
           # Count formula types
           formula_functions = {}
           for _, cell in formulas.iterrows():
               func_match = re.match(r'=([A-Z]+)', cell['formula'])
               if func_match:
                   func = func_match.group(1)
                   formula_functions[func] = formula_functions.get(func, 0) + 1
           
           report.append(f"  Top functions:")
           for func, count in sorted(formula_functions.items(), 
                                   key=lambda x: x[1], reverse=True)[:5]:
               report.append(f"    {func}: {count}")
       report.append("")
       
       # Named ranges
       if len(names) > 0:
           report.append(f"Named Ranges: {len(names)}")
           ranges = names[names['is_range'] == True]
           formulas_named = names[names['is_range'] == False]
           report.append(f"  Cell ranges: {len(ranges)}")
           report.append(f"  Named formulas: {len(formulas_named)}")
       report.append("")
       
       # Validation rules
       if len(validation) > 0:
           report.append(f"Data Validation Rules: {len(validation)}")
           val_types = validation['type'].value_counts()
           for vtype, count in val_types.items():
               report.append(f"  {vtype}: {count}")
       
       return "\n".join(report)

   # Generate report
   report = generate_excel_report("sample.xlsx")
   print(report)

These examples demonstrate the power and flexibility of tidyxl for Excel data analysis. The tidy format makes it easy to filter, analyze, and understand complex spreadsheet structures that would be difficult to work with using traditional tabular import methods.

For a complete interactive demonstration, run the ``examples/complete_demo.py`` script included with the package.