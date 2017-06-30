using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class RowHelper
    {

        /// <summary>
        /// Получить строку. 
        /// Вызывает ошибку если строка еще не существует. 
        /// Смотрите <seealso cref="MakeRow(Worksheet, uint)"/>
        /// </summary>
        /// <param name="worksheet">Лист в котором находится требуемая строка</param>
        /// <param name="rowNum">Номер запрашиваемой строки</param>
        /// <returns>Объект строки</returns>
        public static Row GetRow(this Worksheet worksheet, uint rowNum)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            var wsData = worksheet.GetFirstChild<SheetData>();
            if (wsData == null) { throw new InvalidDocumentStructureException(); }
            var row = wsData.Elements<Row>().FirstOrDefault(r => r.RowIndex.Value == rowNum);
            return row;
        }


        /// <summary>
        /// Получить строку. 
        /// Вызывает ошибку если строка еще не существует.
        /// Смотрите <seealso cref="MakeRow(Worksheet, uint)"/>
        /// </summary>
        /// <param name="worksheet">Лист в котором находится требуемая строка</param>
        /// <param name="rowNum">Номер запрашиваемой строки</param>
        /// <returns>Объект строки</returns>
        public static Row GetRow(this Worksheet worksheet, int rowNum)
        {
            return worksheet.GetRow((uint)rowNum);
        }


        /// <summary>
        /// Создать строку. Если строка уже существует, то данный метод будет
        /// идентичен методу <see cref="GetRow(Worksheet, uint)"/>
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowNum"></param>
        /// <returns></returns>
        public static Row MakeRow(this Worksheet worksheet, uint rowNum)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            var wsData = worksheet.GetFirstChild<SheetData>();
            if (wsData == null) { throw new InvalidDocumentStructureException(); }

            var row = wsData
                    .Elements<Row>()
                    .Where(r => r.RowIndex.Value >= rowNum)
                    .OrderBy(r => r.RowIndex.Value).FirstOrDefault();

            if (row != null && row.RowIndex == rowNum) { return row; }

            var newRow = new Row();
            newRow.RowIndex = rowNum;

            wsData.InsertBefore(newRow, row);

            return newRow;
        }


        /// <summary>
        /// Создать строку. Если строка уже существует, то данный метод будет
        /// идентичен методу <see cref="GetRow(Worksheet, uint)"/>
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowNum"></param>
        /// <returns></returns>
        public static Row MakeRow(this Worksheet worksheet, int rowNum)
        {
            return worksheet.MakeRow((uint)rowNum);
        }

        /// <summary>
        /// Получить все ячейки которые находятся в данной строчке
        /// </summary>
        /// <param name="row">Объект строки OpenXML</param>
        /// <returns>Список ячеек в этой строке</returns>
        public static IEnumerable<Cell> GetCells(this Row row)
        {
            return row.Descendants<Cell>();
        }

        /// <summary>
        /// Получить номер строки
        /// </summary>
        /// <param name="row">Объект строки OpenXML</param>
        /// <returns>Номер данной строки</returns>
        public static uint GetRowNumber(this Row row)
        {
            if(row.RowIndex != null) { return row.RowIndex.Value; } 
            var cell = row.Descendants<Cell>().FirstOrDefault(c => c.CellReference?.Value != null); 
            if(cell == null) { throw new InvalidDocumentStructureException("Не удается получить номер строки для объекта строки!"); } 
            return Utils.ToRowNum(cell.CellReference.Value); 
        }

        /// <summary>
        /// Получить ячейку на пересечении данной строки и указанной ячейки.
        /// Если ячейки не существует, то возвращает null.
        /// </summary>
        /// <param name="row">Объект строки в которой находится требуемая ячейка</param>
        /// <param name="columnNumber">Номер столбца запрашиваемой ячейки (начиная с 1-го)</param>
        /// <returns>
        /// Возвращает ячейку на пересечении данной строки и указанной ячейки.
        /// Если ячейки не существует, то возвращает null
        /// </returns>
        public static Cell GetCell(this Row row, int columnNumber)
        {
            var rowNum = row.GetRowNumber();
            var _colName = Utils.ToColumnName(columnNumber);
            var cellAddress = _colName + rowNum;
            return row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value == cellAddress);
        }

        /// <summary>
        /// Получить ячейку на пересечении данной строки и указанной ячейки.
        /// Если ячейки не существует, то возвращает null.
        /// </summary>
        /// <param name="row">Объект строки в которой находится требуемая ячейка</param>
        /// <param name="columnNumber">Номер столбца запрашиваемой ячейки (начиная с 1-го)</param>
        /// <returns>
        /// Возвращает ячейку на пересечении данной строки и указанной ячейки.
        /// Если ячейки не существует, то возвращает null
        /// </returns>
        public static Cell GetCell(this Row row, uint columnNumber)
        {
            return row.GetCell((int)columnNumber);
        }

        /// <summary>
        /// Получить ячейку на пересечении данной строки и указанной ячейки.
        /// Если ячейки не существует, то возвращает null.
        /// </summary>
        /// <param name="row">Объект строки в которой находится требуемая ячейка</param>
        /// <param name="columnName">Имя столбца запрашиваемой ячейки</param>
        /// <returns>
        /// Возвращает ячейку на пересечении данной строки и указанной ячейки.
        /// Если ячейки не существует, то возвращает null
        /// </returns>
        public static Cell GetCell(this Row row, string columnName)
        {
            var _colNum = Utils.ToColumNum(columnName);
            return row.GetCell((int)_colNum);
        }

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
        public static Cell MakeCell(this Row row, string columnName)
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
        public static Cell MakeCell(this Row row, int columnNumber)
        {
            var _colName = Utils.ToColumnName(columnNumber);
            return row.MakeCell(_colName);
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
        public static Cell MakeCell(this Row row, uint columnNumber)
        {
            var _colName = Utils.ToColumnName(columnNumber);
            return row.MakeCell(_colName);
        }

    }
}
