using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class DocumentGetStylesheet
    {
        /// <summary>
        /// Получить таблицу стилей
        /// </summary>
        /// <param name="document">Рабочий документ</param>
        /// <returns>Таблица стилей указанного документа</returns>
        public static Stylesheet GetStylesheet(this SpreadsheetDocument document)
        {
            return document.WorkbookPart.GetStylesheet();
        }
    }
}
