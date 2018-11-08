using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetAddFormattingRule
    {
        /// <summary>
        /// Добавить условное форматирование
        /// </summary>
        /// <param name="worksheet">Лист в который добавляется форматирование</param>
        /// <param name="formattingExpression">Выражение которое определяет ячейки для форматирования</param>
        /// <param name="style">Стиль условного форматирования</param>
        /// <param name="targetCellAddresses">Области ячеек для которых будет задействовано данное условие</param>
        public static void AddFormattingRule(this Worksheet worksheet, string formattingExpression, DifferentialFormat style, params string[] targetCellAddresses)
        {
            if (targetCellAddresses == null || targetCellAddresses.Count() == 0) { targetCellAddresses = new string[] { "1:1048576" }; }

            var styleSheet = worksheet.GetWorkbookPart().GetStylesheet();
            if (styleSheet.DifferentialFormats == null) { styleSheet.DifferentialFormats = new DifferentialFormats() { Count = 0 }; }
            var dfList = styleSheet.DifferentialFormats;

            dfList.AddDFormat(style);
            var formattingRule = Fabric.MakeFormattingRule(formattingExpression);
            formattingRule.FormatId = (uint)style.Index();

            IEnumerable<StringValue> stringValues = targetCellAddresses.Select(rng => new StringValue(rng));
            var sqref = new ListValue<StringValue>(stringValues);
            var condFormatting = new ConditionalFormatting(formattingRule) { SequenceOfReferences = sqref };

            worksheet.Insert(condFormatting).AfterOneOf(
                    typeof(MergeCells)
                    , typeof(CustomSheetView)
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
    }
}
