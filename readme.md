# Data Generator
Written By: Tyler Hooks 

## Description
This script reads a formatted Excel file and generates random test data in the form of SQL insert statements and writes the SQL to a file called **inserts.sql**.

## Usage
`python generator.py [-h] [-n rows]`

This script requires an Excel workbook called **tables.xlsx** to be in the same directory as the **generator.py** file.

The Excel workbook ***must*** have a sheet for each table, with the sheet having the same name as the table. For example, if you are writing data for tables, **CustomerTB** and **SalesTB**, then you must have two worksheets called **CustomerTB** and **SalesTB** in the workbook.

The columns must adhere to the following format:

| | | |
| - | - | - |
| Data Field | Pattern | Alias (Optional) |
| Data Field | Pattern | Alias (Optional) |
| Data Field | Pattern | Alias (Optional) |
| Data Field | Pattern | Alias (Optional) |
| Data Field | Pattern | Alias (Optional) |

**NO HEADERS ARE REQUIRED.**

### Columns

* **Data Field** - This is the name of the column, or field, in the table.
* **Format Pattern** - This is the format used to generate the data. Refer to the **Pattern** section for more information.
* **Alias** - This is optional. This is to be used when data integrity is required among multiple tables. For example, if **CustomerTB.CustomerID** and **SalesTB.CustID** should have the same values, then they should each have an alias cell populated with the same name (ex. **CID**).

## Patterns

Patterns are used to generate specific, yet random, data. Each pattern must be formatted identically to the examples below.

* **[a-z]\{n\}** - Generate *n* number of lowercase alphabetical characters.
* **[A-Z]\{n\}** - Generate *n* number of uppercase alphabetical characters.
* **[\{lower\}-\{upper\}]\{n\}** - Generate *n* number of integers or decimals between *\{lower\}* and *\{upper\}*.
* **{{alphanum}}\{n\}** - Generate *n* number of alphanumeric characters.
* **[a, b, c, 1, 2, 3]\{n\}** - Generate *n* number of items from the list.
* **{{MM-DD-YYYY}}** - Generate a date between the last 20 years and today. Supports zero-padded (MM) and non-zero-padded (M) values, as well as half-year (YY) and full-year (YYYY) formats. Supports both dash and forward slash separators in different formats (e.g., YYYY-MM-DD, DD/MM/YY, etc.). **Date formats must be capitalized (*MM*, not *mm*). 
* **{{first_name}}** - Generate a random first name.
* **{{last_name}}** - Generate a random last name.
* **{{email}}** - Generate a random email address.
* **{{primary_address}}** - Generate a random primary address.
* **{{secondary_address}}** - Generate a random secondary address.

### Example Pattern

Suppose you wanted to generate a value that followed the format **VX*cccnnnnn***, where *c* is an arbitrary uppercase alphabetical character and *n* is an integer between 1 and 5.
We can generate a value via the follwing pattern: **VX[A-Z]{3}[1-5]{5}**

Here are a few sample return values from the aforementioned pattern:
* VXFSA22112
* VXUBQ15345
* VXGXT11225