using DocumentFormat.OpenXml.Packaging;
using IEIT.Reports.Export.Helpers.Exceptions;
using System;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class DocumentHasWorksheet
    {
        /// <summary>
        /// Получить информацию о существовании листа с указанным названием
        /// </summary>
        /// <param name="doc">Документ, из которого нужно получить информацию</param>
        /// <param name="sheetName">Название листа</param>
        /// <returns>true если лист с таким названием существует в книге, false в обратном случае</returns>
        public static bool HasWorksheet(this SpreadsheetDocument doc, string sheetName)
        {
            if (doc == null) { throw new ArgumentNullException("doc"); }
            if (doc.WorkbookPart == null || doc.WorkbookPart.Workbook == null) { throw new InvalidDocumentStructureException(); }
            return doc.WorkbookPart.Workbook.HasWorksheet(sheetName);
        }
    }
}
