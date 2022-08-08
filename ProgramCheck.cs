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
    class ProgramCheck
    {
        private static void Main()
        {
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(@"C:\Users\nilam\Desktop\Work\project\GeneratedExcel1.xlsx", SpreadsheetDocumentType.Workbook);

            WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();


            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            
            Sheet sheet = new Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "TestSheet"
            };

            sheets.Append(sheet);
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            DataTable dt = Data.SampleData();

            foreach(DataRow item in dt.Rows)
            {
                Row row = new Row();

                for (int i = 0; i < item.ItemArray.Length; i++)
                {
                    Cell cell = new Cell()
                    {
                        CellValue = new CellValue(item[i].ToString()),
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
