using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    internal static class Actions
    {
        public static bool _writeAny(Worksheet worksheet, string cellAddress, string value)
        {
            if (value == null) { value = "-"; }
            if (value.StartsWith("=")) { return _writeFormula(worksheet, cellAddress, value); }
            if (value.IsNumber()) { return _writeNumber(worksheet, cellAddress, value); }
            return _writeText(worksheet, cellAddress, value);
        }

        public static bool _writeText(Worksheet worksheet, string cellAddress, string value)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            Cell cell = worksheet.MakeCell(cellAddress);
            if (cell == null) { throw new IncompleteActionException("Вставка ячейки."); }
            cell = cell.ReplaceBy(new Cell() { StyleIndex = cell.StyleIndex, CellReference = cell.CellReference });
            cell.CellValue = new CellValue(value);
            cell.DataType = CellValues.InlineString;
            cell.InlineString = new InlineString() { Text = new Text(value) };
            return true;
        }

        public static bool _writeNumber(Worksheet worksheet, string cellAddress, string value)
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

        public static bool _writeFormula(Worksheet worksheet, string cellAddress, string formula)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            Cell cell = worksheet.MakeCell(cellAddress);
            if (cell == null) { throw new IncompleteActionException("Вставка ячейки."); }
            cell = cell.ReplaceBy(new Cell() { StyleIndex = cell.StyleIndex, CellReference = cell.CellReference });
            cell.CellFormula = new CellFormula(formula);
            cell.DataType = new EnumValue<CellValues>(CellValues.InlineString);
            return true;
        }

        public static bool _setStyle(Worksheet worksheet, string cellAddress, UInt32Value styleIndex)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            Cell cell = worksheet.MakeCell(cellAddress);
            if (cell == null) { throw new IncompleteActionException("Вставка ячейки."); }
            cell.StyleIndex = styleIndex;
            return true;
        }


    }
}
