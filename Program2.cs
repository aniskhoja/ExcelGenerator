using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Export_Excel;

namespace ExcelGenerator
{
    class Program2
    {
        static void Main()
        {
            Data SampleDataTable = Data.SampleData();
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(@"C:\Users\nilam\Desktop\Work\project\GeneratedExcel.xlsx", SpreadsheetDocumentType.Workbook);

            WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            workbookPart.Workbook.AppendChild(new Sheets());

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheet sheet = new Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "TestSheet"
            };

            spreadsheetDocument.WorkbookPart.Workbook.Sheets.Append(sheet);

            Row row = new Row();
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            sheetData.Append(row);


            Cell cell = new Cell
            {
                CellReference = "A1",
                CellValue = new CellValue("Hello World!"),
                DataType = CellValues.String
            };
            row.Append(cell);


            workbookPart.Workbook.Save();
            spreadsheetDocument.Close();

        }
    }
}
