using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Exceptions;
using System;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class CellGetSharedStringItem
    {
        /// <summary>
        /// Получить <see cref="SharedStringItem"/> объект относящийся к данной ячейке.
        /// Возвращяет null если у данной ячейки нет такого объекта.
        /// </summary>
        /// <param name="cell">Ячейка документа</param>
        /// <returns> <see cref="SharedStringItem"/> объект относящийся к данной ячейке</returns>
        public static SharedStringItem GetSharedStringItem(this Cell cell)
        {
            if (cell == null) { throw new ArgumentNullException("Argument 'cell' must not be null!"); }
            if (cell.CellValue == null || cell.CellValue.Text == null) { return null; }
            if (cell.DataType != CellValues.SharedString) { return null; }
            var wbPart = cell.GetWorkbookPart();
            if (wbPart == null) { throw new InvalidDocumentStructureException("Given worksheet of given cell is not part of workbook!"); }
            var itemId = int.Parse(cell.CellValue.Text);
            return wbPart.GetSharedStringItem(itemId);
        }

    }
}
