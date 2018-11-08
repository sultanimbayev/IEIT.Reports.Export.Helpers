using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Exceptions;
using System.Linq;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class RowGetNumber
    {
        /// <summary>
        /// Получить номер строки
        /// </summary>
        /// <param name="row">Объект строки OpenXML</param>
        /// <returns>Номер данной строки</returns>
        public static uint GetRowNumber(this Row row)
        {
            if (row.RowIndex != null) { return row.RowIndex.Value; }
            var cell = row.Descendants<Cell>().FirstOrDefault(c => c.CellReference?.Value != null);
            if (cell == null) { throw new InvalidDocumentStructureException("Не удается получить номер строки для объекта строки!"); }
            return Utils.ToRowNum(cell.CellReference.Value);
        }
    }
}
