using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Styling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class CellGetStyle
    {
        public static xlCellStyle GetStyle(this Cell cell)
        {
            var cellFormat = cell.CellFormat();
            if(cellFormat == null)
            {
                return new xlCellStyle();
            }
            return new xlCellStyle(cellFormat);
        }
    }
}
