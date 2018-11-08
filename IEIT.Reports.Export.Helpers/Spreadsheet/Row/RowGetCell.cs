using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    /// <summary>
    /// Получение и создание ячеек в строке
    /// </summary>
    public static class RowGetCell
    {
        /// <summary>
        /// Создать ячейку в строке и возвращает созданную ячейку. Если ячейка 
        /// уже существует, то возвращает существующюю ячейку. Не перезаписывает 
        /// старую ячейку.
        /// </summary>
        /// <param name="row">Объект строки в которой создается ячейка</param>
        /// <param name="columnName">Название столбца новой ячейки</param>
        /// <returns>
        /// Возвращает созданную ячейку. Если ячейка  уже существует, 
        /// то возвращает существующюю ячейку
        /// </returns>
        public static Cell GetCell(this Row row, string columnName)
        {
            var _colName = Utils.ToColumnName(columnName);
            var _colNum = Utils.ToColumNum(columnName);
            var cellAddress = _colName + row.GetRowNumber();
            var cell = row.Elements<Cell>()
                .Where(c => c.CellReference?.Value != null)
                .Where(c => Utils.ToColumNum(c.CellReference.Value) >= _colNum)
                .OrderBy(c => Utils.ToColumNum(c.CellReference.Value))
                .FirstOrDefault();

            if (cell != null && cell.CellReference.Value.Equals(cellAddress, StringComparison.OrdinalIgnoreCase)) { return cell; }

            var newCell = new Cell();
            newCell.CellReference = cellAddress;
            row.InsertBefore(newCell, cell);

            return newCell;
        }

        /// <summary>
        /// Создать ячейку в строке и возвращает созданную ячейку. Если ячейка 
        /// уже существует, то возвращает существующюю ячейку. Не перезаписывает 
        /// старую ячейку.
        /// </summary>
        /// <param name="row">Объект строки в которой создается ячейка</param>
        /// <param name="columnNumber">Номер столбца новой ячейки (начиная с 1-го)</param>
        /// <returns>
        /// Возвращает созданную ячейку. Если ячейка  уже существует, 
        /// то возвращает существующюю ячейку
        /// </returns>
        public static Cell GetCell(this Row row, int columnNumber)
        {
            var _colName = Utils.ToColumnName(columnNumber);
            return row.GetCell(_colName);
        }


        /// <summary>
        /// Создать ячейку в строке и возвращает созданную ячейку. Если ячейка 
        /// уже существует, то возвращает существующюю ячейку. Не перезаписывает 
        /// старую ячейку.
        /// </summary>
        /// <param name="row">Объект строки в которой создается ячейка</param>
        /// <param name="columnNumber">Номер столбца новой ячейки (начиная с 1-го)</param>
        /// <returns>
        /// Возвращает созданную ячейку. Если ячейка  уже существует, 
        /// то возвращает существующюю ячейку
        /// </returns>
        public static Cell GetCell(this Row row, uint columnNumber)
        {
            var _colName = Utils.ToColumnName(columnNumber);
            return row.GetCell(_colName);
        }

    }
}
