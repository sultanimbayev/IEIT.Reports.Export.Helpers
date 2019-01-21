using DocumentFormat.OpenXml;
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
    /// Получение и создание строк в листе
    /// </summary>
    public static class _WorksheetGetRow
    {
        /// <summary>
        /// Получить строку. 
        /// Создается новый элемент строки, если строка еще не существует.
        /// </summary>
        /// <param name="worksheet">Лист в котором находится требуемая строка</param>
        /// <param name="rowNum">Номер запрашиваемой строки</param>
        /// <returns>Объект строки</returns>
        public static Row GetRow(this Worksheet worksheet, uint rowNum)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet is null"); }
            var wsData = worksheet.GetFirstChild<SheetData>();
            if (wsData == null) { throw new InvalidDocumentStructureException(); }
            return wsData.GetRow(rowNum);
        }


        /// <summary>
        /// Получить строку. Создается новый элемент строки, если строка еще не существует.
        /// </summary>
        /// <param name="worksheet">Лист в котором находится требуемая строка</param>
        /// <param name="rowNum">Номер запрашиваемой строки</param>
        /// <returns>Объект строки</returns>
        public static Row GetRow(this Worksheet worksheet, int rowNum)
        {
            return worksheet.GetRow((uint)rowNum);
        }

    }
}
