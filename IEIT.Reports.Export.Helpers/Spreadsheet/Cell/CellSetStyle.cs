using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Styling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class CellSetStyle
    {
        public static void SetCellStyle(this Cell cell, xlCellStyle cellStyle)
        {
            if(cellStyle == null)
            {
                cell.StyleIndex = 0;
                return;
            }
            var wbPart = cell.GetWorkbookPart();
            cell.StyleIndex = cellStyle.GetStyleIndexFor(wbPart);
        }
    }
}
