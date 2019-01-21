using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetMergeCells
    {
        /// <summary>
        /// Объеденение ячеек по индексам.
        /// </summary>
        /// <param name="ws">Лист в котором требуется объединить ячейки</param>
        /// <param name="columnFrom">Номер начального столбца (начиная с 1)</param>
        /// <param name="rowFrom">Номер начальной строки (начиная с 1)</param>
        /// <param name="columnTo">Номер конечного столбца (начиная с 1)</param>
        /// <param name="rowTo">Номер конечной строки (начиная с 1)</param>
        public static void MergeCells(this Worksheet ws, int columnFrom, int rowFrom, int columnTo, int rowTo)
        {
            ws.MergeCells($"{Utils.ToColumnName(columnFrom)}{rowTo}:{Utils.ToColumnName(columnTo)}{rowTo}");
        }


        /// <summary>
        /// Объединить ячейки
        /// </summary>
        /// <param name="worksheet">Лист в котором требуется объединить ячейки</param>
        /// <param name="cellsRange">Область объединяемых ячеек</param>
        public static void MergeCells(this Worksheet worksheet, string cellsRange)
        {
            Regex rgxCellsRange = new Regex(Common.RGX_PAT_CA_RANGE);
            if (!rgxCellsRange.IsMatch(cellsRange)) { throw new Exception($"Не удалось распознать область '{cellsRange}' адресов ячеек. Проверьте формат."); }

            MergeCells mergeCells;
            if (worksheet.Elements<MergeCells>().Count() > 0)
            {
                mergeCells = worksheet.Elements<MergeCells>().First();
            }
            else
            {
                mergeCells = new MergeCells();


                // Insert a MergeCells object into the specified position.
                worksheet.Insert(mergeCells).AfterOneOf(
                    typeof(CustomSheetView)
                    , typeof(DataConsolidate)
                    , typeof(SortState)
                    , typeof(AutoFilter)
                    , typeof(Scenarios)
                    , typeof(ProtectedRanges)
                    , typeof(SheetProtection)
                    , typeof(SheetCalculationProperties)
                    , typeof(SheetData)
                    );
            }

            // Create the merged cell and append it to the MergeCells collection.
            MergeCell mergeCell = new MergeCell() { Reference = new StringValue(cellsRange) };
            mergeCells.Append(mergeCell);
        }
    }
}
