using DocumentFormat.OpenXml.Spreadsheet;
namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class CellGetWorksheet
    {
        /// <summary>
        /// Получить рабочий лист в которой находится ячейка
        /// </summary>
        /// <param name="cell">Ячейка документа</param>
        /// <returns>Рабочий лист в которой находится ячейка</returns>
        public static Worksheet GetWorksheet(this Cell cell)
        {
            return cell.ParentOfType<Worksheet>();
        }
    }
}
