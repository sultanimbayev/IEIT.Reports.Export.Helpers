using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class DocumentNewWorksheet
    {
        /// <summary>
        /// Создает новый лист в книге
        /// </summary>
        /// <param name="document">Документ таблиц OpenXml</param>
        /// <param name="sheetName">Наименование нового листа</param>
        /// <returns>Объект созданного листа Worksheet</returns>
        public static Worksheet NewWorksheet(this SpreadsheetDocument document, string sheetName)
        {
            if (document == null)
            {
                throw new ArgumentNullException("'Document' is null");
            }
            var workbookpart = document.WorkbookPart;
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            var sheets = document.WorkbookPart.Workbook.FirstDescendant<Sheets>();
            if(sheets == null) { sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets()); }

            //Getting max number of sheet id
            var maxSheetId = sheets.Descendants<Sheet>().Max(s => s.SheetId);

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = maxSheetId + 1,
                Name = sheetName
            };
            sheets.Append(sheet);

            return worksheetPart.Worksheet;
        }
    }
}
