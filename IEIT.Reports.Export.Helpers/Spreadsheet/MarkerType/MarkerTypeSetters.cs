using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class MarkerTypeSetters
    {
        public static MarkerType SetRowNum(this MarkerType marker, int rowNum)
        {
            if(marker == null) { return null; }
            marker.RowId = new RowId((rowNum - 1).ToString());
            return marker;
        }
        public static MarkerType SetColumnNum(this MarkerType marker, int columnNum)
        {
            if (marker == null) { return null; }
            marker.ColumnId = new ColumnId((columnNum - 1).ToString());
            return marker;
        }
        public static MarkerType SetTopOffset(this MarkerType marker, double topOffset = 0)
        {
            if (marker == null) { return null; }
            marker.RowOffset = new RowOffset(((int)topOffset).ToString());
            return marker;
        }
        public static MarkerType SetLeftOffset(this MarkerType marker, double leftOffset = 0)
        {
            if (marker == null) { return null; }
            marker.ColumnOffset = new ColumnOffset(((int)leftOffset).ToString());
            return marker;
        }


    }
}
