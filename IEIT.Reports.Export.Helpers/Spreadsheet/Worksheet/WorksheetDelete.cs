using DocumentFormat.OpenXml.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetDelete
    {

        /// <summary>
        /// Удаляет лист
        /// </summary>
        /// <param name="worksheet">Рабочий лист</param>
        /// <returns>true при удачном удалении, false в обратном случае</returns>
        public static bool Delete(this Worksheet worksheet)
        {
            var WbPart = worksheet.GetWorkbookPart();
            var sheet = worksheet.GetSheet();

            if (sheet == null) { return false; }

            // Remove the sheet reference from the workbook.
            sheet.Remove();

            // Delete the worksheet part.
            return WbPart.DeletePart(worksheet.WorksheetPart);
        }
    }
}
