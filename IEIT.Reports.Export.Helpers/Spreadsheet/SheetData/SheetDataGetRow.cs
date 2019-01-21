using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    /// <summary>
    /// Позволяет получить объект строки по указанному адресу из данных листа (объекта SheetData).
    /// </summary>
    public static class SheetDataGetRow
    {
        /// <summary>
        /// Получить строку. 
        /// Создается новый элемент строки, если строка еще не существует.
        /// </summary>
        /// <param name="sheetData">Данные листа в котором находится требуемая строка</param>
        /// <param name="rowNum">Номер запрашиваемой строки</param>
        /// <returns>Объект строки</returns>
        public static Row GetRow(this SheetData sheetData, uint rowNum)
        {
            var row = sheetData
                    .Elements<Row>()
                    .Where(r => r.RowIndex.Value >= rowNum)
                    .OrderBy(r => r.RowIndex.Value).FirstOrDefault();

            if (row != null && row.RowIndex == rowNum) { return row; }

            var newRow = new Row();
            newRow.RowIndex = rowNum;

            sheetData.InsertBefore(newRow, row);

            return newRow;
        }

        /// <summary>
        /// Получить строку. 
        /// Создается новый элемент строки, если строка еще не существует.
        /// </summary>
        /// <param name="sheetData">Данные листа в котором находится требуемая строка</param>
        /// <param name="rowNum">Номер запрашиваемой строки</param>
        /// <returns>Объект строки</returns>
        public static Row GetRow(this SheetData sheetData, int rowNum)
        {
            if(rowNum <= 0) { throw new Exception("Row number must be greater than zero"); }
            return GetRow(sheetData, (uint)rowNum);
        }
    }
}
