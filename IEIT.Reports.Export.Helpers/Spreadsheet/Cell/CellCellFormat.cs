using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class CellCellFormat
    {
        public static CellFormat CellFormat(this Cell cell)
        {
            if(cell.StyleIndex == null) { return null; }
            var wbPart = cell.GetWorkbookPart();
            var stylesheet = wbPart.GetStylesheet();
            return stylesheet.CellFormat(cell.StyleIndex);
        }
        public static void CellFormat(this Cell cell, CellFormat cellFormat)
        {
            var wbPart = cell.GetWorkbookPart();
            var stylesheet = wbPart.GetStylesheet();
            cell.StyleIndex = stylesheet.CellFormat(cellFormat);
        }
    }
}
