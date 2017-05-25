using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using IEIT.Reports.Export.Helpers.Exceptions;
using IEIT.Reports.Export.Helpers.Spreadsheet.Models;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class SpreadsheetHelper
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
            if (doc.WorkbookPart == null || doc.WorkbookPart.Workbook == null) { throw new IncorrectDocumentStructureException(); }
            return doc.WorkbookPart.Workbook.GetWorksheet(sheetName);
        }

        /// <summary>
        /// Получить часть документа с рабочим листом
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static Worksheet GetWorksheet(this Workbook workbook, string sheetName)
        {
            if (workbook == null) { throw new ArgumentNullException("workbook"); }
            if (workbook.WorkbookPart == null) { throw new IncorrectDocumentStructureException(); }
            var rel = workbook.Descendants<Sheet>()
                .Where(s => s.Name.Value.Equals(sheetName))
                .First();
            if(rel == null || rel.Id == null) { return null; }
            var wsPart = workbook.WorkbookPart.GetPartById(rel.Id) as WorksheetPart;
            if (wsPart == null) { return null; }
            return wsPart.Worksheet;
        }

        private static Sheet GetSheet(this Worksheet worksheet)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            if (worksheet.WorksheetPart == null) { throw new IncorrectDocumentStructureException(); }
            var wbPart = worksheet.GetWorkbookPart();
            if (wbPart == null) { return null; }
            var wsPartId = wbPart.GetIdOfPart(worksheet.WorksheetPart);
            var sheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Id != null && s.Id.Value == wsPartId).FirstOrDefault();
            return sheet;
        }

        private static WorkbookPart GetWorkbookPart(this Worksheet worksheet)
        {
            return worksheet.WorksheetPart.GetParentParts().FirstOrDefault() as WorkbookPart;
        }


        public static string GetName(this Worksheet worksheet)
        {
            var sheet = worksheet.GetSheet();
            if (sheet == null) { return null; }
            return sheet.Name;
        }

        public static Intent Write(this Worksheet ws, string value)
        {
            return makeIntentToWrite(ws, _writeAny, value);
        }

        public static Intent Write(this Worksheet ws, object value)
        {
            var _val = value != null ? value.ToString() : "-";
            return makeIntentToWrite(ws, _writeAny, _val);
        }

        public static Intent WriteText(this Worksheet ws, string text)
        {
            return makeIntentToWrite(ws, _writeText, text);
        }

        public static Intent WriteFormula(this Worksheet ws, string formula)
        {
            return makeIntentToWrite(ws, _writeFormula, formula);
        }

        public static Intent WriteNumber(this Worksheet ws, string number)
        {
            return makeIntentToWrite(ws, _writeNumber, number);
        }

        private static Intent makeIntentToWrite(Worksheet ws, Func<Worksheet, string, string, bool> writeDeleg, string text)
        {
            return new Intent(ws, writeDeleg, _setStyle).WithText(text);
        }

        public static Intent SetStyle(this Worksheet ws, UInt32Value styleIndex)
        {
            return new Intent(ws, _writeAny, _setStyle).WithStyle(styleIndex);
        }

        private static bool _writeAny(Worksheet worksheet, string cellAddress, string value)
        {
            if(value == null){ value = "-"; }
            if (value.StartsWith("=")) { return _writeFormula(worksheet, cellAddress, value); }
            if (value.IsNumber()) { return _writeNumber(worksheet, cellAddress, value); }
            return _writeText(worksheet, cellAddress, value);
        }

        private static bool _writeText(Worksheet worksheet, string cellAddress, string value)
        {
            if(worksheet == null) { throw new ArgumentNullException("worksheet"); }
            Cell cell = worksheet.MakeCell(cellAddress);
            if (cell == null) { throw new IncompleteActionException("Вставка ячейки."); }
            cell = cell.ReplaceBy(new Cell() { StyleIndex = cell.StyleIndex, CellReference = cell.CellReference });
            cell.CellValue = new CellValue(value);
            cell.DataType = CellValues.InlineString;
            cell.InlineString = new InlineString() { Text = new Text(value) };
            return true;
        }

        private static bool _writeNumber(Worksheet worksheet, string cellAddress, string value)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            Cell cell = worksheet.MakeCell(cellAddress);
            if (cell == null) { throw new IncompleteActionException("Вставка ячейки."); }
            cell = cell.ReplaceBy(new Cell() { StyleIndex = cell.StyleIndex, CellReference = cell.CellReference });
            if (!string.IsNullOrEmpty(value)) value = value.Replace(",", ".");
            cell.CellValue = new CellValue(value);
            //cell.DataType = CellValues.Number;
            //cell.DataType = new EnumValue<CellValues>(CellValues.Number);
            return true;
        }

        private static bool _writeFormula(Worksheet worksheet, string cellAddress, string formula)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            Cell cell = worksheet.MakeCell(cellAddress);
            if (cell == null) { throw new IncompleteActionException("Вставка ячейки."); }
            cell = cell.ReplaceBy(new Cell() { StyleIndex = cell.StyleIndex, CellReference = cell.CellReference });
            cell.CellFormula = new CellFormula(formula);
            cell.DataType = new EnumValue<CellValues>(CellValues.InlineString);
            return true;
        }

        private static bool _setStyle(Worksheet worksheet, string cellAddress, UInt32Value styleIndex)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            Cell cell = worksheet.MakeCell(cellAddress);
            if (cell == null) { throw new IncompleteActionException("Вставка ячейки."); }
            cell.StyleIndex = styleIndex;
            return true;
        }

        public static Cell GetCell(this Worksheet worksheet, string cellAddress)
        {
            if (worksheet == null){ throw new ArgumentNullException("worksheet"); }
            
            var rowNum = Utils.ToRowNum(cellAddress);
            var row = worksheet.GetRow(rowNum);

            var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value == cellAddress);
            return cell;
        }

        public static Cell MakeCell(this Worksheet worksheet, string cellAddress)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            var rowNum = Utils.ToRowNum(cellAddress);
            var row = worksheet.MakeRow(rowNum);

            if(row == null) { throw new IncompleteActionException("Создание строки."); }
            
            var cell = row.Elements<Cell>()
                .Where(c => Utils.ToColumNum(c.CellReference.Value) >= Utils.ToColumNum(cellAddress))
                .OrderBy(c => Utils.ToColumNum(c.CellReference.Value))
                .FirstOrDefault();

            if (cell != null && cell.CellReference.Value.Equals(cellAddress, StringComparison.OrdinalIgnoreCase)) { return cell; }
            
            var newCell = new Cell();
            newCell.CellReference = cellAddress;
            row.InsertBefore(newCell, cell);

            return newCell;

        }

        public static Row GetRow(this Worksheet worksheet, uint rowNum)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            var wsData = worksheet.GetFirstChild<SheetData>();
            if(wsData == null) { throw new IncorrectDocumentStructureException(); }
            var row = wsData.Elements<Row>().FirstOrDefault(r => r.RowIndex.Value == rowNum);
            return row;
        }

        public static Row MakeRow(this Worksheet worksheet, uint rowNum)
        {
            if(worksheet == null) { throw new ArgumentNullException("worksheet"); }
            var wsData = worksheet.GetFirstChild<SheetData>();
            if (wsData == null) { throw new IncorrectDocumentStructureException(); }
            
            var row = wsData
                    .Elements<Row>()
                    .Where(r => r.RowIndex.Value >= rowNum)
                    .OrderBy(r => r.RowIndex.Value).FirstOrDefault();

            if (row != null && row.RowIndex == rowNum) { return row; }

            var newRow = new Row();
            newRow.RowIndex = rowNum;

            wsData.InsertBefore(newRow, row);

            return newRow;
        }
        


        public static bool Rename(this Worksheet worksheet, string newName, bool updateReferences = true)
        {
            var sheet = worksheet.GetSheet();
            if (sheet == null) { return false; }

            var wbPart = worksheet.GetWorkbookPart();
            if(wbPart == null) { return false; }

            if (updateReferences)
            {
                var pattern = $"'?{sheet.Name}'?!";
                var replacement = $"'{newName}'!";
                
                var wsParts = wbPart.GetPartsOfType<WorksheetPart>();
                foreach(var wsPart in wsParts)
                {
                    wsPart.RootElement.RegexReplaceIn<OpenXmlLeafTextElement>(pattern, replacement);
                    foreach(var chartPart in wsPart.DrawingsPart.ChartParts)
                    {
                        chartPart.RootElement.RegexReplaceIn<OpenXmlLeafTextElement>(pattern, replacement);
                    }
                }
            }
            
            return (sheet.Name = newName).Equals(newName);
        }


    }
}
