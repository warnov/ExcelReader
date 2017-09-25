using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WarNov.ExcelReader
{
    public static class ExcelReader
    {
        public static string ReadCell(SpreadsheetDocument doc, int sheetNumber, int rowNumber, int cellNumber)
        {
            sheetNumber--;
            rowNumber--;
            cellNumber--;
            WorkbookPart wbPart = doc.WorkbookPart;
            Sheet mysheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(sheetNumber);
            Worksheet Worksheet = ((WorksheetPart)wbPart.GetPartById(mysheet.Id)).Worksheet;
            int wkschildno = 4;
            SheetData Rows = (SheetData)Worksheet.ChildElements.GetItem(wkschildno);
            Row currentrow = (Row)Rows.ChildElements.GetItem(rowNumber);
            var columnsForCurrentRow = currentrow.ChildElements.Count();
            string currentcellvalue = string.Empty;
            if (cellNumber < columnsForCurrentRow)
            {
                Cell currentcell = (Cell)currentrow.ChildElements.GetItem(cellNumber);

                if (currentcell.DataType != null)
                {
                    if (currentcell.DataType == CellValues.SharedString)
                    {
                        int id = -1;

                        if (Int32.TryParse(currentcell.InnerText, out id))
                        {
                            SharedStringItem item = GetSharedStringItemById(wbPart, id);

                            if (item.Text != null)
                            {
                                //code to take the string value  
                                currentcellvalue = item.Text.Text;
                            }
                            else if (item.InnerText != null)
                            {
                                currentcellvalue = item.InnerText;
                            }
                            else if (item.InnerXml != null)
                            {
                                currentcellvalue = item.InnerXml;
                            }
                        }
                    }
                }
                else
                {
                    currentcellvalue = currentcell.InnerText;
                }
            }
            return currentcellvalue.Replace('"', '\'');
        }

        public static string SheetName(SpreadsheetDocument doc, int sheetNumber)
        {
            sheetNumber--;
            WorkbookPart wbPart = doc.WorkbookPart;
            Sheet mysheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(sheetNumber);
            return mysheet.Name;

        }

        public static SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id)
        {
            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        }

        public static String CellError(int tab, int row, int column)
        {
            return $"Error reading the cell {row}, {column} in the tab {tab}";
        }

    }
}
