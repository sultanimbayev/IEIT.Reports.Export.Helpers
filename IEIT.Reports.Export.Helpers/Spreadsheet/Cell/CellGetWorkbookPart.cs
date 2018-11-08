using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Exceptions;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class CellGetWorkbookPart
    {
        /// <summary>
        /// Получить рабочюю книгу документа в которой находится данная ячейка
        /// </summary>
        /// <param name="cell">Ячейка документа</param>
        /// <returns>Рабочая книга документа в которой находится данная ячейка</returns>
        public static WorkbookPart GetWorkbookPart(this Cell cell)
        {
            var ws = cell.ParentOfType<Worksheet>();
            if (ws == null) { throw new InvalidDocumentStructureException("Given cell is not part of worksheet!"); }
            return ws.GetWorkbookPart();
        }

    }
}
