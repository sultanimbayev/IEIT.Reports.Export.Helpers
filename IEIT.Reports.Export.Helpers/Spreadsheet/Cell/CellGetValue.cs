using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class CellGetValue
    {
        /// <summary>
        /// Получить значение ячейки
        /// </summary>
        /// <param name="cell">Ячейка, значение которой нужно получить</param>
        /// <returns>Значение ячейки, если имеется. null, если значение не найдено</returns>
        public static string GetValue(this Cell cell)
        {
            if (cell == null) { throw new ArgumentNullException("Given Cell object is null"); }
            var item = cell.CellValue as OpenXmlElement;
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                item = cell.GetSharedStringItem();
            }
            if (cell.DataType != null && cell.DataType == CellValues.InlineString)
            {
                item = cell.InlineString;
            }
            return item?.InnerText;
        }
    }
}
