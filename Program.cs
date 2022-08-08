using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelGenerator
{
    class Program
    {
        static void Main()
        {
            DataTable SampleDataTable = Data.SampleData();
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(@"C:\Users\nilam\Desktop\Work\project\sampleData.xlsx", true);

            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;


            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            char[] refrence = "BCD".ToCharArray();

            foreach (DataRow item in SampleDataTable.Rows)
            {
                int skipRowIndex = 6;
                skipRowIndex += SampleDataTable.Rows.IndexOf(item);
                Row row = new Row();

                for (int i = 0; i < item.ItemArray.Length; i++)
                {
                    Cell cell = new Cell()
                    {
                        CellValue = new CellValue(item[i].ToString()),
                        CellReference = refrence[i].ToString() + skipRowIndex,
                        DataType = CellValues.String
                    };
                    row.Append(cell);
                }
                sheetData.Append(row);

            }

            workbookPart.Workbook.Save();
            spreadsheetDocument.Close();

        }
    }
}
