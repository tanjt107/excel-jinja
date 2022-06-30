# Intoduction
Excel file (.xlsx) is actually a zipped folder of XML files. When some functions (e.g. partial superscript, read only) are not supported by Python packages like openpyxl, we can consider changing the content of an Excel file direcly using [Jinja](https://jinja.palletsprojects.com/en/3.1.x/)

# How it works
1. Unzip the .xlsx file (or just replace the .xlsx to .zip)
2. Copy workbook.xml and all xml files in sheet folder to template folder
3. Replace all .xml in template folder to .j2
4. Change xml contents using Jinja syntax
5. Stream and dump contents to j2 templates (see `example.py`)

# See also
[XML Structure of an Excel file](https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-spreadsheetml-document?redirectedfrom=MSDN)
