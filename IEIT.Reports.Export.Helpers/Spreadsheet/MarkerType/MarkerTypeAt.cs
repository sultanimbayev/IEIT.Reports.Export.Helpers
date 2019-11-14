using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class MarkerTypeAt
    {
        public static xdr.MarkerType At(this xdr.MarkerType marker, int rowNum, int columnNum)
        {
            if (marker == null) { return null; }
            marker.RowId = new xdr.RowId(rowNum.ToString());
            marker.ColumnId = new xdr.ColumnId(columnNum.ToString());
            marker.RowOffset = new xdr.RowOffset();
            marker.ColumnOffset = new xdr.ColumnOffset();
            return marker;
        }

        public static xdr.MarkerType At(this xdr.MarkerType marker, int rowNum, double topOffset, int columnNum, double leftOffset)
        {
            if (marker == null) { return null; }
            marker.RowId = new xdr.RowId(rowNum.ToString());
            marker.ColumnId = new xdr.ColumnId(columnNum.ToString());
            marker.RowOffset = new xdr.RowOffset(topOffset.ToString());
            marker.ColumnOffset = new xdr.ColumnOffset(leftOffset.ToString());
            return marker;
        }

        public static xdr.MarkerType At(this xdr.MarkerType marker, string cellAddress, double topOffset = 0, double leftOffset = 0)
        {
            if (marker == null) { return null; }
            var rowNum = Utils.ToRowNum(cellAddress);
            var columnNum = Utils.ToColumNum(cellAddress);
            marker.RowId = new xdr.RowId(rowNum.ToString());
            marker.ColumnId = new xdr.ColumnId(columnNum.ToString());
            marker.RowOffset = new xdr.RowOffset(topOffset.ToString());
            marker.ColumnOffset = new xdr.ColumnOffset(leftOffset.ToString());
            return marker;
        }
    }
}
