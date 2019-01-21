using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    /// <summary>
    /// Позволяет получить объект ячейки по указанному адресу из данных листа (объекта SheetData).
    /// </summary>
    public static class SheetDataGetCell
    {
        /// <summary>
        /// Получить объект ячейки по указанному адресу. 
        /// </summary>
        /// <param name="sheetData">Данные листа из которого нужно получить ячейку</param>
        /// <param name="cellAddress">Адрес новой ячейки</param>
        /// <returns>Созданную ячейку, если ячейка не существовала</returns>
        public static Cell GetCell(this SheetData sheetData, string cellAddress)
        {
            if (sheetData == null) { throw new ArgumentNullException("sheetData", "SheetData object must not be null!"); }
            var rowNum = Utils.ToRowNum(cellAddress);
            var row = sheetData.GetRow(rowNum);
            if (row == null) { throw new IncompleteActionException($"Не удалось создать строку {rowNum}."); }
            return row.GetCell(cellAddress);
        }

        /// <summary>
        /// Получить объект ячейки по указанному адресу. 
        /// </summary>
        /// <param name="sheetData">Данные листа (SheetData)</param>
        /// <param name="columnNumber">Номер колонки ячейки</param>
        /// <param name="rowNumber">Номер строки ячейки</param>
        /// <returns>Созданную ячейку, если ячейка не существовала</returns>
        public static Cell GetCell(this SheetData sheetData, int columnNumber, int rowNumber)
        {
            if (sheetData == null) { throw new ArgumentNullException("sheetData", "SheetData object must not be null!"); }
            var row = sheetData.GetRow(rowNumber);
            if (row == null) { throw new IncompleteActionException($"Не удалось создать строку {rowNumber}."); }
            return row.GetCell(columnNumber);
        }

        /// <summary>
        /// Получить объект ячейки по указанному адресу. 
        /// </summary>
        /// <param name="sheetData">Данные листа (SheetData)</param>
        /// <param name="columnNumber">Номер колонки ячейки</param>
        /// <param name="rowNumber">Номер строки ячейки</param>
        /// <returns>Созданную ячейку, если ячейка не существовала</returns>
        public static Cell GetCell(this SheetData sheetData, uint columnNumber, uint rowNumber)
        {
            if (sheetData == null) { throw new ArgumentNullException("sheetData", "SheetData object must not be null!"); }
            var row = sheetData.GetRow(rowNumber);
            if (row == null) { throw new IncompleteActionException($"Не удалось создать строку {rowNumber}."); }
            return row.GetCell(columnNumber);
        }
    }
}
