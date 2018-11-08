using DocumentFormat.OpenXml.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetGetName
    {
        /// <summary>
        /// Получить название листа
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public static string GetName(this Worksheet worksheet)
        {
            var sheet = worksheet.GetSheet();
            if (sheet == null) { return null; }
            return sheet.Name;
        }
    }
}
