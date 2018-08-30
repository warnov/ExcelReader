# ExcelReader
This is a lightweight Excel Reader using DocumentFormat.OpenXml. It just does one thing:
Read the content of a cell and return it as a string.

The main method <code>ReadCell</code> just requires:
* The reference to the spreadsheet (either by the path of the Excel file or using directly a <code>SpreadSheet</code> from <code>DocumentFormat.OpenXml.Spreadsheet;</code>
* The number of the worksheet where the cell is
* The number of the row
* The name of the column
