using DocumentFormat.OpenXml.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class CellGetColumn
    {
        /// <summary>
        /// Получить объект <see cref="Models.Column"/>
        /// для работы со столбцом в которой находится ячейка.
        /// </summary>
        /// <param name="cell">Объект ячейки OpenXML</param>
        /// <returns>
        /// объект для работы со столбцом, в которой находится ячейка
        /// </returns>
        public static Models.Column GetColumn(this Cell cell)
        {
            var ws = cell.GetWorksheet();
            var _colName = Utils.ToColumnName(cell.CellReference.Value);
            return new Models.Column(ws, _colName);
        }
    }
}
