using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Exceptions;
using System;
using System.Linq;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorkbookGetWorksheet
    {
        /// <summary>
        /// Получить лист по его названию. Возвращает null если такой лист не найден
        /// </summary>
        /// <param name="workbook">Рабочая книга документа</param>
        /// <param name="sheetName">Название листа</param>
        /// <returns>Рабочий лист с указанным названием или null если такой лист не найден</returns>
        public static Worksheet GetWorksheet(this Workbook workbook, string sheetName)
        {
            if (workbook == null) { throw new ArgumentNullException("workbook"); }
            if (workbook.WorkbookPart == null) { throw new InvalidDocumentStructureException(); }
            var rel = workbook.Descendants<Sheet>()
                .Where(s => s.Name.Value.Equals(sheetName))
                .FirstOrDefault();
            if (rel == null || rel.Id == null) { return null; }
            var wsPart = workbook.WorkbookPart.GetPartById(rel.Id) as WorksheetPart;
            if (wsPart == null) { return null; }
            return wsPart.Worksheet;
        }
    }
}
