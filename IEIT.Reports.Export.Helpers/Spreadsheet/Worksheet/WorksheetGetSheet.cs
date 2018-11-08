using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Exceptions;
using System;
using System.Linq;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetGetSheet
    {
        /// <summary>
        /// Получить свойства листа
        /// </summary>
        /// <param name="worksheet">Объект листа</param>
        /// <returns>Объект содержащий свойства листа <see cref="Sheet"/></returns>
        internal static Sheet GetSheet(this Worksheet worksheet)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            if (worksheet.WorksheetPart == null) { throw new InvalidDocumentStructureException(); }
            var wbPart = worksheet.GetWorkbookPart();
            if (wbPart == null) { return null; }
            var wsPartId = wbPart.GetIdOfPart(worksheet.WorksheetPart);
            var sheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Id != null && s.Id.Value == wsPartId).FirstOrDefault();
            return sheet;
        }
    }
}
