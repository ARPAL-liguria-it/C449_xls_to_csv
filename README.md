# Conversion of predefined Excel reports to csv files for method C449

The application converts a portion of predefined Excel files 
to csv files.

The user must specify a folder in which the reports are stored 
and a folder where csv files should be saved.

The Excel files must have a sheet exactly named 
"_perCalcoliDiluizione_", and only the portion of that sheet
in columns A:B and rows 103:157 will be converted to csv.

The output csv file will have semicolon (;) as field separator, dot (.)
as decimal separator and strings, even the empty ones, quoted within double
marks (").