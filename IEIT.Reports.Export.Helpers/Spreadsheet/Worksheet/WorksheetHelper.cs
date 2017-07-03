using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Exceptions;
using IEIT.Reports.Export.Helpers.Spreadsheet.Intents;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetHelper
    {

        /// <summary>
        /// Получить рабочий лист
        /// </summary>
        /// <param name="doc">Документ, из которого нужно получить лист</param>
        /// <param name="sheetName">Название требуемого листа</param>
        /// <returns>Рабочий лист, название которого соответсвует указанному, или null если лист не найден.</returns>
        public static Worksheet GetWorksheet(this SpreadsheetDocument doc, string sheetName)
        {
            if (doc == null) { throw new ArgumentNullException("doc"); }
            if (doc.WorkbookPart == null || doc.WorkbookPart.Workbook == null) { throw new InvalidDocumentStructureException(); }
            return doc.WorkbookPart.Workbook.GetWorksheet(sheetName);
        }

        /// <summary>
        /// Получить информацию о существовании листа с указанным названием
        /// </summary>
        /// <param name="doc">Документ, из которого нужно получить информацию</param>
        /// <param name="sheetName">Название листа</param>
        /// <returns>true если лист с таким названием существует в книге, false в обратном случае</returns>
        public static bool HasWorksheet(this SpreadsheetDocument doc, string sheetName)
        {
            if (doc == null) { throw new ArgumentNullException("doc"); }
            if (doc.WorkbookPart == null || doc.WorkbookPart.Workbook == null) { throw new InvalidDocumentStructureException(); }
            return doc.WorkbookPart.Workbook.HasWorksheet(sheetName);
        }

        /// <summary>
        /// Получить лист по его названию. Возвращает null если такой лист не найден
        /// </summary>
        /// <param name="workbook">Рабочая книга документа</param>
        /// <param name="sheetName">Название листа</param>
        /// <returns>Рабочий лист с указанным названием или null если такой лист не найден</returns>
        public static Worksheet GetWorksheet(this Workbook workbook, string sheetName)
        {
            if (workbook == null) { throw new ArgumentNullException("workbook"); }
            if (workbook.WorkbookPart == null) { throw new InvalidDocumentStructureException(); }
            var rel = workbook.Descendants<Sheet>()
                .Where(s => s.Name.Value.Equals(sheetName))
                .FirstOrDefault();
            if (rel == null || rel.Id == null) { return null; }
            var wsPart = workbook.WorkbookPart.GetPartById(rel.Id) as WorksheetPart;
            if (wsPart == null) { return null; }
            return wsPart.Worksheet;
        }

        /// <summary>
        /// Получить информацию о существовании листа с указанным названием
        /// </summary>
        /// <param name="workbook">Рабочая книга документа</param>
        /// <param name="sheetName">Название листа</param>
        /// <returns>true если лист с таким названием существует в книге, false в обратном случае</returns>
        public static bool HasWorksheet(this Workbook workbook, string sheetName)
        {
            if (workbook == null) { throw new ArgumentNullException("workbook"); }
            if (workbook.WorkbookPart == null) { throw new InvalidDocumentStructureException(); }
            return workbook.Descendants<Sheet>()
                .Where(s => s.Name.Value.Equals(sheetName))
                .Count() > 0;
        }

        /// <summary>
        /// Получить свойства листа
        /// </summary>
        /// <param name="worksheet">Объект листа</param>
        /// <returns>Объект содержащий свойства листа <see cref="Sheet"/></returns>
        internal static Sheet GetSheet(this Worksheet worksheet)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            if (worksheet.WorksheetPart == null) { throw new InvalidDocumentStructureException(); }
            var wbPart = worksheet.GetWorkbookPart();
            if (wbPart == null) { return null; }
            var wsPartId = wbPart.GetIdOfPart(worksheet.WorksheetPart);
            var sheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Id != null && s.Id.Value == wsPartId).FirstOrDefault();
            return sheet;
        }

        /// <summary>
        /// Получить часть документа который содержит все рабочие листы.
        /// </summary>
        /// <param name="worksheet">Рабочий лист</param>
        /// <returns>Часть документа который содержит все рабочие листы</returns>
        public static WorkbookPart GetWorkbookPart(this Worksheet worksheet)
        {
            return worksheet.WorksheetPart.GetParentParts().FirstOrDefault() as WorkbookPart;
        }


        /// <summary>
        /// Переименовать лист
        /// </summary>
        /// <param name="worksheet">Лист который ты хочешь переименовать</param>
        /// <param name="newName">Новое название листа</param>
        /// <param name="updateReferences">Заменить все ссылки к данному листу?</param>
        /// <returns>true при удачном переименовывании, false в обратном случае</returns>
        public static bool Rename(this Worksheet worksheet, string newName, bool updateReferences = true)
        {
            var sheet = worksheet.GetSheet();
            if (sheet == null) { return false; }

            var wbPart = worksheet.GetWorkbookPart();
            if (wbPart == null) { return false; }

            if (updateReferences)
            {
                var pattern = $"'?{sheet.Name}'?!";
                var replacement = $"'{newName}'!";

                var wsParts = wbPart.GetPartsOfType<WorksheetPart>();
                foreach (var wsPart in wsParts)
                {
                    //TODO: Улучшить обновление ссылок на лист
                    wsPart.RootElement.RegexReplaceIn<OpenXmlLeafTextElement>(pattern, replacement);
                    foreach (var chartPart in wsPart.DrawingsPart.ChartParts)
                    {
                        chartPart.RootElement.RegexReplaceIn<OpenXmlLeafTextElement>(pattern, replacement);
                    }
                }
            }

            return (sheet.Name = newName).Equals(newName);
        }

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
            formattingRule.FormatId = (uint)style.GetIndex();

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

        /// <summary>
        /// Получить объект для работы со столбцом, с указанным адресом.
        /// </summary>
        /// <param name="worksheet">Объект рабочего листа</param>
        /// <param name="columnNumber">Номер запрашиваемого столбца (начиная с 1-го)</param>
        /// <returns>Объект для работы со столбцом</returns>
        public static Models.Column GetColumn(this Worksheet worksheet, int columnNumber)
        {
            return new Models.Column(worksheet, columnNumber);
        }

        /// <summary>
        /// Получить объект для работы со столбцом, с указанным адресом.
        /// </summary>
        /// <param name="worksheet">Объект рабочего листа</param>
        /// <param name="columnNumber">Номер запрашиваемого столбца (начиная с 1-го)</param>
        /// <returns>Объект для работы со столбцом</returns>
        public static Models.Column GetColumn(this Worksheet worksheet, uint columnNumber)
        {
            return new Models.Column(worksheet, columnNumber);
        }

        /// <summary>
        /// Получить объект для работы со столбцом, с указанным адресом.
        /// </summary>
        /// <param name="worksheet">Объект рабочего листа</param>
        /// <param name="columnName">Название запрашиваемого столбца, латинские буквы.</param>
        /// <returns>Объект для работы со столбцом</returns>
        public static Models.Column GetColumn(this Worksheet worksheet, string columnName)
        {
            return new Models.Column(worksheet, columnName);
        }

    }
}
