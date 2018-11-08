using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class RowGetCells
    {
        /// <summary>
        /// Получить все ячейки которые находятся в данной строчке
        /// </summary>
        /// <param name="row">Объект строки OpenXML</param>
        /// <returns>Список ячеек в этой строке</returns>
        public static IEnumerable<Cell> GetCells(this Row row)
        {
            return row.Descendants<Cell>();
        }
    }
}
