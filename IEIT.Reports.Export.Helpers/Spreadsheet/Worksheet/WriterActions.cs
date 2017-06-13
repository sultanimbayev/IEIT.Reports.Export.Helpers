using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Exceptions;
using System;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    internal static class WriterActions
    {

        private static Cell _getCell(Worksheet worksheet, string cellAddress)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            Cell cell = worksheet.MakeCell(cellAddress);
            if (cell == null) { throw new IncompleteActionException("Не удалось вставить ячейку!"); }
            return cell;
        }

        internal static bool _writeAny(Worksheet worksheet, string cellAddress, string value)
        {
            return _getCell(worksheet, cellAddress).Write(value);
        }

        internal static bool _writeText(Worksheet worksheet, string cellAddress, string value)
        {
            return _getCell(worksheet, cellAddress).WriteText(value);
        }


        internal static bool _appendText(Worksheet worksheet, string cellAddress, string value, RunProperties textStyle = null)
        {
            return _getCell(worksheet, cellAddress).AppendText(value, textStyle);
        }

        internal static bool _writeNumber(Worksheet worksheet, string cellAddress, string value)
        {
            return _getCell(worksheet, cellAddress).WriteNumber(value);
        }

        internal static bool _writeFormula(Worksheet worksheet, string cellAddress, string formula)
        {
            return _getCell(worksheet, cellAddress).WriteFormula(formula);
        }

        internal static bool _setStyle(Worksheet worksheet, string cellAddress, UInt32Value styleIndex)
        {
            _getCell(worksheet, cellAddress).StyleIndex = styleIndex;
            return true;
        }


    }
}
