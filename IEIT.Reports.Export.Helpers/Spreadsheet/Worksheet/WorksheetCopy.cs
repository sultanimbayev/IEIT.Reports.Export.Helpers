using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Spreadsheet.Intents;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetCopy
    {

        /// <summary>
        /// Копировать ячейки
        /// </summary>
        /// <param name="worksheet">Лист из которого ячейки будут скопированы</param>
        /// <param name="cellsRange">Область копируемых ячеек, указывать в формате A1:B2. Можно указать адрес одной ячейки</param>
        /// <returns>"Намерение" <see cref="PasteIntent"/> для вставки ячеек</returns>
        public static PasteIntent Copy(this Worksheet worksheet, string cellsRange)
        {
            return new PasteIntent(worksheet, cellsRange);
        }
    }
}
