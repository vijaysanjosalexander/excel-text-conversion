# Excel to text conversion
Python utility to convert Excel files in .xlsx to plain text for Natural Language Processing. The utility will allow retaining the header information by adding the header to each and every cells in the excel file.


Excel 
--------------------------------------
|ID | Name | Address  | Age | Gender |
--------------------------------------
| 1 | VSA  | Bangalore | 30 | Male   |
 -------------------------------------
| 2 | ASA  | Bangalore | 28 | Male   |
--------------------------------------



Converted Text
---------------
SHEET NAME is Sheet 1 and ID is 1 and Name is VSA and Address is Bangalore and Age is 30 and Gender is Male and ID is 2 and Name is ASA and Address is Bangalore and Age is 28 and Gender is Male

The extraceted string is stored in UTF-8 format.

Date in Excel:
--------------
Excel stores date as float number and the xldate_as_datetime(xldate, datemode) function is used to convert float to date.
If the corresponding header has a date word in it, then only the value is passed to the conversion function.
