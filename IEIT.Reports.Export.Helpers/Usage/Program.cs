using DocumentFormat.OpenXml.Packaging;
using IEIT.Reports.Export.Helpers.Spreadsheet;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.IO;

namespace Usage
{
    class Program
    {
        static void Main(string[] args)
        {
            var filepath = "Temp.xlsx";
            
            if (File.Exists(filepath))
            {
                File.Delete(filepath);
            }

            var doc = CreateSpreadsheetWorkbook(filepath);

            doc.GetWorksheet("list1").Write("Hello world!").To("B5");

            doc.WorkbookPart.Workbook.Save();
            doc.Close();

        }

        public static SpreadsheetDocument CreateSpreadsheetWorkbook(string filepath, string sheetName = "list1")
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                Create(filepath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = sheetName
            };
            sheets.Append(sheet);

            workbookpart.Workbook.Save();

            // Close the document.
            //spreadsheetDocument.Close();
            return spreadsheetDocument;
        }
    }
}
