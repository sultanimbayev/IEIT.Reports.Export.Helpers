using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using IEIT.Reports.Export.Helpers.Exceptions;
using IEIT.Reports.Export.Helpers.Spreadsheet.Intents;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class SpreadsheetHelper
    {

        /// <summary>
        /// Создает новый файл документа Excel с единственным листом который называется Sheet1
        /// Возвращает созданный документ
        /// </summary>
        /// <param name="filepath">Директория где будет создан файл</param>
        /// <returns>Возвращает созданный документ</returns>
        public static SpreadsheetDocument CreateBlank(string filepath)
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

            // Initializing Shared string table
            workbookpart.AddNewPart<SharedStringTablePart>();
            workbookpart.SharedStringTablePart.SharedStringTable = new SharedStringTable() { Count = 0, UniqueCount = 0 };

            //Initializing Stylesheet
            workbookpart.AddNewPart<WorkbookStylesPart>();
            var stylesheet = workbookpart.WorkbookStylesPart.Stylesheet = new Stylesheet();
            stylesheet.Fills = new Fills(
                new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }, // required, reserved by Excel
                new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } } // required, reserved by Excel
                ) { Count = 2 };
            stylesheet.Fonts = new Fonts(new Font()) { Count = 1 }; // blank font list
            stylesheet.Borders = new Borders(new Border()) { Count = 1 };
            stylesheet.CellFormats = new CellFormats(new CellFormat()) { Count = 1 }; // cell format list; empty one for index 0, seems to be required
            stylesheet.CellStyleFormats = new CellStyleFormats(new CellFormat()) { Count = 1 }; // blank cell format list


            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Sheet1"
            };
            sheets.Append(sheet);

            workbookpart.Workbook.Save();

            // Close the document.
            //spreadsheetDocument.Close();
            return spreadsheetDocument;
        }

        /// <summary>
        /// Сохранить изменения и закрыть документ
        /// </summary>
        /// <param name="document">Документ над которым производится операция</param>
        public static void SaveAndClose(this SpreadsheetDocument document)
        {
            document.Save();
            document.Close();
        }


    }
}
