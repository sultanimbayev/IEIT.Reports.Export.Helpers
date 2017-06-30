using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class CellHelper
    {

        public enum MatchOption
        {
            Contains,
            Equals,
            StartsWith,
            EndsWith
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


        /// <summary>
        /// Получить рабочий лист в которой находится ячейка
        /// </summary>
        /// <param name="cell">Ячейка документа</param>
        /// <returns>Рабочий лист в которой находится ячейка</returns>
        public static Worksheet GetWorksheet(this Cell cell)
        {
            return cell.GetFirstParent<Worksheet>();
        }

        /// <summary>
        /// Получить рабочюю книгу документа в которой находится данная ячейка
        /// </summary>
        /// <param name="cell">Ячейка документа</param>
        /// <returns>Рабочая книга документа в которой находится данная ячейка</returns>
        public static WorkbookPart GetWorkbookPart(this Cell cell)
        {
            var ws = cell.GetFirstParent<Worksheet>();
            if (ws == null) { throw new InvalidDocumentStructureException("Given cell is not part of worksheet!"); }
            return ws.GetWorkbookPart();
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
            var item = cell.CellValue as OpenXmlElement;
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                item = cell.GetSharedStringItem(); 
            }
            if (cell.DataType != null && cell.DataType == CellValues.InlineString)
            {
                item = cell.InlineString;
            }
            return item?.InnerText;
        }

        /// <summary>
        /// Найти ячейки по его содержанию
        /// </summary>
        /// <param name="worksheet">Рабочий лист документа в котором ведется поиск</param>
        /// <param name="searchText">Значение которое должно содержать ячейка</param>
        /// <returns>Ячейки содержание которых совпадает с указанным значением</returns>
        public static IEnumerable<Cell> FindCells(this Worksheet worksheet, string searchText, MatchOption match = MatchOption.Contains)
        {
            Func<string, string, bool> matchDeleg;
            switch (match)
            {
                default:
                    throw new NotImplementedException();
                case MatchOption.Contains:
                    matchDeleg = (_cellText, _searchTxt) => { return _cellText.Contains(_searchTxt); };
                    break;
                case MatchOption.Equals:
                    matchDeleg = (_cellText, _searchTxt) => { return _cellText.Equals(_searchTxt); };
                    break;
                case MatchOption.StartsWith:
                    matchDeleg = (_cellText, _searchTxt) => { return _cellText.StartsWith(_searchTxt); };
                    break;
                case MatchOption.EndsWith:
                    matchDeleg = (_cellText, _searchTxt) => { return _cellText.EndsWith(_searchTxt); };
                    break;
            }
            return worksheet.Descendants<Cell>().Where(c => { var val = c.GetValue(); return val != null && matchDeleg(val, searchText); });
        }

        /// <summary>
        /// Найти ячейки по его содержанию
        /// </summary>
        /// <param name="worksheet">Рабочий лист документа в котором ведется поиск</param>
        /// <param name="searchRgx">Объект регулярного выражения для поиска</param>
        /// <returns>Ячейки содержание которых совпадает с данным выражением</returns>
        public static IEnumerable<Cell> FindCells(this Worksheet worksheet, Regex searchRgx)
        {
            return worksheet.Descendants<Cell>().Where(c => { var val = c.GetValue(); return val != null && searchRgx.IsMatch(val); });
        }
        

        /// <summary>
        /// Получить объект ячейки. 
        /// Вызывает ошибку если ячейка еще не существует.
        /// <seealso cref="MakeCell(Worksheet, string)"/>
        /// </summary>
        /// <param name="worksheet">Лист в котором находится ячейка</param>
        /// <param name="cellAddress">Адрес ячейки</param>
        /// <returns>Объект ячейки который находится в данном листе по указанному адресу</returns>
        public static Cell GetCell(this Worksheet worksheet, string cellAddress)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            var rowNum = Utils.ToRowNum(cellAddress);
            var row = worksheet.GetRow(rowNum);
            if (row == null) { return null; }
            var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value == cellAddress);
            return cell;
        }



        /// <summary>
        /// Создать ячейку. Если ячейка уже создана в указанном месте, 
        /// тогда данный метод будет идентичен методу <see cref="GetCell(Worksheet, string)"/>
        /// </summary>
        /// <param name="worksheet">Лист в котором нужно создать ячейку</param>
        /// <param name="cellAddress">Адрес новой ячейки</param>
        /// <returns>Созданную ячейку, если ячейка не существовала</returns>
        public static Cell MakeCell(this Worksheet worksheet, string cellAddress)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            var rowNum = Utils.ToRowNum(cellAddress);
            var row = worksheet.MakeRow(rowNum);
            if (row == null) { throw new IncompleteActionException("Создание строки."); } 
            return row.MakeCell(cellAddress);
        }

        /// <summary>
        /// Получить строку в которой находится данная ячейка
        /// </summary>
        /// <param name="cell">Объект ячейки OpenXML</param>
        /// <returns>Объект строки, в которой находится ячейка</returns>
        public static Row GetRow(this Cell cell)
        {
            return cell.GetFirstParent<Row>();
        }

        /// <summary>
        /// Получить объект <see cref="Models.Column"/>
        /// для работы со столбцом в которой находится ячейка.
        /// </summary>
        /// <param name="cell">Объект ячейки OpenXML</param>
        /// <returns>
        /// объект для работы со столбцом, в которой находится ячейка
        /// </returns>
        public static Models.Column GetColumn(this Cell cell)
        {
            var ws = cell.GetWorksheet();
            var _colName = Utils.ToColumnName(cell.CellReference.Value);
            return new Models.Column(ws, _colName);
        }

    }
}
