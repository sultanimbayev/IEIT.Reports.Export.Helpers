using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Exceptions;
using System;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class CellValueGetCell
    {
        /// <summary>
        /// Получить ячейку данного значения
        /// </summary>
        /// <param name="cellValue">Значение ячейки</param>
        /// <returns>Родительский элемент, ячейку в котором хранится данное значение</returns>
        public static Cell GetCell(this CellValue cellValue)
        {
            if (cellValue == null) { throw new ArgumentNullException("Given CellValue object is null"); }
            if (cellValue.Parent == null) { throw new InvalidDocumentStructureException("cellValue has no parent"); }
            if (cellValue.Parent == null || !(cellValue.Parent is Cell)) { throw new InvalidDocumentStructureException("CellValue object has no Cell parent!"); }
            return cellValue.Parent as Cell;
        }

    }
}
