using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Styling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetGetCellStyle
    {
        public static xlCellStyle GetCellStyle(this Worksheet ws, string cellAddress)
        {
            return ws.GetCell(cellAddress).GetStyle();
        }
    }
}
