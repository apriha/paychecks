# paychecks
VBA code to generate an Excel workbook with dynamic pie charts for tracking paychecks

## Overview
**paychecks** is a VBA module that generates a workbook for tracking paychecks. Specifically, the user specifies the number of initial rows for paycheck data and the number of columns (i.e., "buckets") for earnings, before tax deductions, after tax deductions, and taxes. Running Main() generates the workbook. The workbook can then be used like any regular Excel workbook, and the paychecks macro is no longer needed. Additionally, *no macros are embedded in the generated workbook or are needed to use the generated workbook.*

The generated paychecks workbook contains a number of features to facilitate tracking paychecks:

* The requested number of columns for each group (e.g., earnings, taxes, etc.)
 * These represent the different "buckets" for pay source / pay destination
* Total column for each group
* Grand total row for each column
* Dynamic pie charts for "Pay Source" and "Pay Destination" that update in real-time to reflect the workbook's data and column titles

Hopefully this workbook will help the user track his or her paychecks and gain a better understanding of how the pies get sliced.

## Usage
1. Open the Visual Basic Editor in Excel and import / copy the code in paychecks.bas
2. Edit the constants in Main() for desired number of columns and rows
3. Run Main()
 3. The workbook should be created
4. Rename generic column titles
 4. These are the labels used in the dynamic pie charts
5. Save the workbook, add paycheck data, insert rows as necessary, and track paychecks.

## Example
Editing the constants as indicated in the code block below results in a workbook with 4 columns for earnings income, 1 column for before tax deductions, 3 columns for taxes, and 26 rows for paycheck data.

```
    ' column types and quantities
    Const EARNINGS_COLUMNS As Integer = 4    ' 1 or greater
    Const BEFORE_TAX_COLUMNS As Integer = 1  ' 0 or greater
    Const AFTER_TAX_COLUMNS As Integer = 0   ' 0 or greater
    Const TAX_COLUMNS As Integer = 3         ' 0 or greater
    
    ' initial number of paychecks / rows for paychecks
    Const PAYCHECK_ROWS As Integer = 26
```

A screenshot of the generated workbook:



This example workbook is also included in the *examples* directory.
