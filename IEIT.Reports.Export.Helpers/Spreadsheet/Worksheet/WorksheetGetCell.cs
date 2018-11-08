using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Exceptions;
using System;
using System.Linq;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetGetCell
    {
        /// <summary>
        /// Создать ячейку. Если ячейка уже создана в указанном месте, 
        /// тогда данный метод будет идентичен методу <see cref="GetCell(Worksheet, string)"/>
        /// </summary>
        /// <param name="worksheet">Лист в котором нужно создать ячейку</param>
        /// <param name="cellAddress">Адрес новой ячейки</param>
        /// <returns>Созданную ячейку, если ячейка не существовала</returns>
        public static Cell GetCell(this Worksheet worksheet, string cellAddress)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            var rowNum = Utils.ToRowNum(cellAddress);
            var row = worksheet.GetRow(rowNum);
            if (row == null) { throw new IncompleteActionException($"Не удалось создать строку {rowNum}."); }
            return row.GetCell(cellAddress);
        }

        /// <summary>
        /// Создать ячейку. Если ячейка уже создана в указанном месте, 
        /// тогда данный метод будет идентичен методу <see cref="GetCell(Worksheet, string)"/>
        /// </summary>
        /// <param name="worksheet">Лист</param>
        /// <param name="columnNumber">Номер колонки ячейки</param>
        /// <param name="rowNumber">Номер строки ячейки</param>
        /// <returns>Созданную ячейку, если ячейка не существовала</returns>
        public static Cell GetCell(this Worksheet worksheet, int columnNumber, int rowNumber)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            var row = worksheet.GetRow(rowNumber);
            if (row == null) { throw new IncompleteActionException($"Не удалось создать строку {rowNumber}."); }
            return row.GetCell(columnNumber);
        }

        /// <summary>
        /// Создать ячейку. Если ячейка уже создана в указанном месте, 
        /// тогда данный метод будет идентичен методу <see cref="GetCell(Worksheet, string)"/>
        /// </summary>
        /// <param name="worksheet">Лист</param>
        /// <param name="columnNumber">Номер колонки ячейки</param>
        /// <param name="rowNumber">Номер строки ячейки</param>
        /// <returns>Созданную ячейку, если ячейка не существовала</returns>
        public static Cell GetCell(this Worksheet worksheet, uint columnNumber, uint rowNumber)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            var row = worksheet.GetRow(rowNumber);
            if (row == null) { throw new IncompleteActionException($"Не удалось создать строку {rowNumber}."); }
            return row.GetCell(columnNumber);
        }

    }
}
