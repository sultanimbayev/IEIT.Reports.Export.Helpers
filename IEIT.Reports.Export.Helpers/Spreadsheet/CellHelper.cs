using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Exceptions;
using System;
namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class CellHelper
    {
        /// <summary>
        /// Переместить значение хранящиеся в данном объекте в SharedString
        /// </summary>
        /// <param name="cellValue">Оъект со значениями ячейки</param>
        public static SharedStringItem MoveToSS(this CellValue cellValue)
        {
            var cell = cellValue.GetCell();
            if(cell.DataType != null && cell.DataType == CellValues.SharedString){ return cellValue.GetSharedStringItem(); }
            var text = cellValue.Text;
            var wbPart = cell.GetWorkbookPart();
            if(wbPart == null) { throw new InvalidDocumentStructureException("Given worksheet of given cell is not part of workbook!"); }
            if(wbPart.SharedStringTablePart == null) { wbPart.AddNewPart<SharedStringTablePart>(); }
            if(wbPart.SharedStringTablePart.SharedStringTable == null) { wbPart.SharedStringTablePart.SharedStringTable = new SharedStringTable().From("SST.Empty"); }
            var sst = wbPart.SharedStringTablePart.SharedStringTable;
            var itemIdx = sst.Count.Value;
            var newItem = sst.Add(text);
            cell.DataType = CellValues.SharedString;
            cellValue.Text = itemIdx.ToString();
            cell.InlineString = null;
            return newItem;
        }


        public static Cell GetCell(this CellValue cellValue)
        {
            if (cellValue == null) { throw new ArgumentNullException("Given CellValue object is null"); }
            if (cellValue.Parent == null) { throw new InvalidDocumentStructureException("cellValue has no parent"); }
            if (cellValue.Parent == null || !(cellValue.Parent is Cell)) { throw new InvalidDocumentStructureException("CellValue object has no Cell parent!"); }
            return cellValue.Parent as Cell;
        }

        public static Worksheet GetWorksheet(this Cell cell)
        {
            return cell.GetFirstParent<Worksheet>();
        }

        public static WorkbookPart GetWorkbookPart(this Cell cell)
        {
            var ws = cell.GetFirstParent<Worksheet>();
            if (ws == null) { throw new InvalidDocumentStructureException("Given cell is not part of worksheet!"); }
            return ws.GetWorkbookPart();
        }

        public static SharedStringItem GetSharedStringItem(this CellValue cellValue)
        {
            if(cellValue.Text == null) { return null; }
            var cell = cellValue.GetCell();
            if(cell.DataType != CellValues.SharedString) { return null; }
            var wbPart = cell.GetWorkbookPart();
            if (wbPart == null) { throw new InvalidDocumentStructureException("Given worksheet of given cell is not part of workbook!"); }
            var itemId = int.Parse(cellValue.Text);
            return wbPart.GetSharedStringItem(itemId);
        }

        /// <summary>
        /// Добавление текста в ячейку
        /// </summary>
        /// <param name="cell">Ячейка в которую ведется запись</param>
        /// <param name="text">Добавляемый текст</param>
        /// <param name="styles">Стиль добавляемого текта</param>
        /// <returns>Всегда true</returns>
        public static bool AppendText(this Cell cell, string text, RunProperties styles = null)
        {
            var item = cell.CellValue.MoveToSS();
            if (item == null) { return false; }
            item.Append(text, styles);
            return true;
        }

        /// <summary>
        /// Запись текста в ячейку
        /// </summary>
        /// <param name="cell">Ячейка в которую ведется запись</param>
        /// <param name="text">Записываемый текст</param>
        /// <returns>Всегда true</returns>
        public static bool WriteText(this Cell cell, string value)
        {
            cell = cell.ReplaceBy(new Cell() { StyleIndex = cell.StyleIndex, CellReference = cell.CellReference });
            cell.CellValue = new CellValue(value);
            cell.DataType = CellValues.InlineString;
            cell.InlineString = new InlineString() { Text = new Text(value) };
            return true;
        }

        /// <summary>
        /// Запись числа в ячейку
        /// </summary>
        /// <param name="cell">Ячейка в которую ведется запись</param>
        /// <param name="value">Записываемое число</param>
        /// <returns></returns>
        public static bool WriteNumber(this Cell cell, string value)
        {
            cell = cell.ReplaceBy(new Cell() { StyleIndex = cell.StyleIndex, CellReference = cell.CellReference });
            if (!string.IsNullOrEmpty(value)) value = value.Replace(",", ".");
            cell.CellValue = new CellValue(value);
            cell.DataType = CellValues.Number;
            return true;
        }

        /// <summary>
        /// Запись формулы в ячейку
        /// </summary>
        /// <param name="cell">Ячейка в которую ведется запись</param>
        /// <param name="formula">Записываемая формула</param>
        /// <returns>Всегда true</returns>
        public static bool WriteFormula(this Cell cell, string formula)
        {
            cell = cell.ReplaceBy(new Cell() { StyleIndex = cell.StyleIndex, CellReference = cell.CellReference });
            cell.CellFormula = new CellFormula(formula);
            cell.DataType = new EnumValue<CellValues>(CellValues.InlineString);
            return true;
        }

        /// <summary>
        /// Запись значения в ячейку
        /// <para>-Запись текста</para>
        /// <para>-Запись числа</para>
        /// <para>-Запись формулы</para>
        /// </summary>
        /// <param name="cell">Ячейка в которую ведется запись</param>
        /// <param name="value">Записываемое значение</param>
        /// <returns></returns>
        public static bool Write(this Cell cell, string value)
        {
            if (value == null) { value = "-"; }
            if (value.StartsWith("=")) { return cell.WriteFormula(value); }
            if (Utils.IsNumber(value)) { return cell.WriteNumber(value); }
            return cell.WriteText(value);
        }
    }
}
