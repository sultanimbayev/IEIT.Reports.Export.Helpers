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
        public static T At<T>(this T marker, int rowNum, int columnNum) where T : xdr.MarkerType
        {
            if (marker == null) { return null; }
            if(rowNum < 1) { rowNum = 1; }
            if(columnNum < 1) { columnNum = 1; }
            marker.RowId = new xdr.RowId((rowNum - 1).ToString());
            marker.ColumnId = new xdr.ColumnId((columnNum - 1).ToString());
            marker.RowOffset = new xdr.RowOffset();
            marker.ColumnOffset = new xdr.ColumnOffset();
            return marker;
        }

        public static T At<T>(this T marker, int rowNum, double topOffset, int columnNum, double leftOffset) where T : xdr.MarkerType
        {
            if (marker == null) { return null; }
            if (rowNum < 1) { rowNum = 1; }
            if (columnNum < 1) { columnNum = 1; }
            marker.RowId = new xdr.RowId((rowNum - 1).ToString());
            marker.ColumnId = new xdr.ColumnId((columnNum - 1).ToString());
            if(topOffset < 0) { topOffset = 0; }
            if(leftOffset < 0) { leftOffset = 0; }
            marker.RowOffset = new xdr.RowOffset(topOffset.ToString());
            marker.ColumnOffset = new xdr.ColumnOffset(leftOffset.ToString());
            return marker;
        }

        public static T At<T>(this T marker, string cellAddress, double topOffset = 0, double leftOffset = 0) where T : xdr.MarkerType
        {
            if (marker == null) { return null; }
            var rowNum = Utils.ToRowNum(cellAddress);
            var columnNum = Utils.ToColumnNum(cellAddress);
            if (rowNum < 1) { rowNum = 1; }
            if (columnNum < 1) { columnNum = 1; }
            marker.RowId = new xdr.RowId((rowNum - 1).ToString());
            marker.ColumnId = new xdr.ColumnId((columnNum - 1).ToString());
            if (topOffset < 0) { topOffset = 0; }
            if (leftOffset < 0) { leftOffset = 0; }
            marker.RowOffset = new xdr.RowOffset(topOffset.ToString());
            marker.ColumnOffset = new xdr.ColumnOffset(leftOffset.ToString());
            return marker;
        }
    }
}
