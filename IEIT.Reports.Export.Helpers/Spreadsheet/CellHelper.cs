using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Exceptions;
using System;
using System.Linq;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class CellHelper
    {
        /// <summary>
        /// Переместить значение хранящиеся в данном объекте в SharedString.
        /// Не преобразует значения типа <see cref="CellValues.Boolean"/>
        /// <see cref="CellValues.Date"/> <see cref="CellValues.Error"/>
        /// <see cref="CellValues.Number"/> если не указан параметр <paramref name="force"/> как true
        /// </summary>
        /// <param name="cell">Ячейка, значение которой нужно сделать общим</param>
        /// <param name="force">Флаг "насильного" преобразования, если указан как true, то преобразует значение не смотря на его тип.
        /// А если указан false (по умолчанию), то преобразует только строковые значения.
        /// </param>
        /// <returns>Преобразованное значение <see cref="SharedStringItem"/> при удачном преобразовании, 
        /// либо null в обратном случае</returns>
        public static SharedStringItem MakeValueShared(this Cell cell, bool force=false)
        {

            if(cell.DataType != null && cell.DataType == CellValues.SharedString){ return cell.GetSharedStringItem(); }
            
            if(cell.DataType == null
                || cell.DataType == CellValues.Boolean 
                || cell.DataType == CellValues.Date
                || cell.DataType == CellValues.Error 
                || cell.DataType == CellValues.Number
                || force)
            {
                return null;
            }

            var wbPart = cell.GetWorkbookPart();
            if(wbPart == null) { throw new InvalidDocumentStructureException("Given worksheet of given cell is not part of workbook!"); }
            if(wbPart.SharedStringTablePart == null) { wbPart.AddNewPart<SharedStringTablePart>(); }
            if(wbPart.SharedStringTablePart.SharedStringTable == null) { wbPart.SharedStringTablePart.SharedStringTable = new SharedStringTable().From("SST.Empty"); }
            var sst = wbPart.SharedStringTablePart.SharedStringTable;
            if(cell.CellValue == null) { cell.CellValue = new CellValue(); }
            var itemIdx = sst.Elements().Count();
            SharedStringItem newItem;
            
            if (cell.DataType != null && cell.DataType == CellValues.InlineString)
            {
                var inStr = cell.InlineString;
                newItem = new SharedStringItem(inStr.Elements().Select(el => el.CloneNode(true)));
                sst.Append(newItem);
            }
            else
            {
                var text = cell.CellValue.Text;
                newItem = sst.Add(text);
            }

            cell.DataType = CellValues.SharedString;
            cell.CellValue.Text = itemIdx.ToString();
            cell.InlineString = null;
            return newItem;
        }


        /// <summary>
        /// Переместить значение хранящиеся в данном объекте в форматированную строку
        /// </summary>
        /// <param name="cellValue"></param>
        public static void ValueAsInlineString(this Cell cell)
        {
            if(cell.DataType != null && cell.DataType == CellValues.InlineString) { return; }
            if(cell.CellValue == null) { cell.CellValue = new CellValue(); }
            InlineString newInStr;
            if(cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                var ssItem = cell.GetSharedStringItem();
                newInStr = new InlineString(ssItem.Elements().Select(el => el.CloneNode(true)));
            }
            else
            {
                var text = cell.CellValue.Text;
                newInStr = new InlineString();
                newInStr.Text = new Text(text);
            }

            cell.CellValue = new CellValue();
            cell.InlineString = newInStr;
            cell.DataType = CellValues.InlineString;

        }

        /// <summary>
        /// Получить ячейку данного значения
        /// </summary>
        /// <param name="cellValue">Значение ячейки</param>
        /// <returns>Родительский элемент, ячейку в котором хранится данное значение</returns>
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

        public static SharedStringItem GetSharedStringItem(this Cell cell)
        {
            if(cell == null) { throw new ArgumentNullException("Argument 'cell' must not be null!"); }
            if(cell.CellValue == null || cell.CellValue.Text == null) { return null; }
            if(cell.DataType != CellValues.SharedString) { return null; }
            var wbPart = cell.GetWorkbookPart();
            if (wbPart == null) { throw new InvalidDocumentStructureException("Given worksheet of given cell is not part of workbook!"); }
            var itemId = int.Parse(cell.CellValue.Text);
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
            cell.ValueAsInlineString();
            if(cell.InlineString == null) { cell.InlineString = new InlineString(); }
            cell.InlineString.AppendText(text, styles);
            cell.CellValue = new CellValue();
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

        /// <summary>
        /// Получить значение ячейки
        /// </summary>
        /// <param name="cell">Ячейка, значение которой нужно получить</param>
        /// <returns>Значение ячейки, если имеется. null, если значение не найдено</returns>
        public static string GetValue(this Cell cell)
        {
            if (cell == null) { throw new ArgumentNullException("Given Cell object is null"); }
            if (cell.CellValue == null) { return null; }
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                var item = cell.GetSharedStringItem();
                if (item == null) { return null; }
                return item.InnerText;
            }
            return cell.CellValue.InnerText;
        }

    }
}
