using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetGetWorkbookPart
    {
        /// <summary>
        /// Получить часть документа который содержит все рабочие листы.
        /// </summary>
        /// <param name="worksheet">Рабочий лист</param>
        /// <returns>Часть документа который содержит все рабочие листы</returns>
        public static WorkbookPart GetWorkbookPart(this Worksheet worksheet)
        {
            return worksheet.WorksheetPart.GetParentParts().FirstOrDefault(p => p.GetType() == typeof(WorkbookPart)) as WorkbookPart;
        }
    }
}
