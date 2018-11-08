using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class CellWrite
    {
        /// <summary>
        /// Запись текста в ячейку
        /// </summary>
        /// <param name="cell">Ячейка в которую ведется запись</param>
        /// <param name="value">Записываемый текст</param>
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
