using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class TwoCellAnchorGetWidth
    {
        public static double GetWidthInPixels(this xdr.TwoCellAnchor twoCellAnchor, double dpi = 96)
        {
            if (twoCellAnchor == null)
            {
                throw new ArgumentNullException("Cannot get anchor width, null given.");
            }
            var fromMarker = twoCellAnchor.FromMarker;
            if(fromMarker == null)
            {
                fromMarker = new xdr.FromMarker().At("A1");
                twoCellAnchor.FromMarker = fromMarker;
            }
            if (!int.TryParse(fromMarker?.ColumnId?.Text, out var startColumnId))
            {
                throw new Exception($"Cannot get top left columm number of given shape. Found \"{fromMarker?.ColumnId?.Text}\"");
            }
            if (!int.TryParse(fromMarker?.ColumnOffset?.Text, out var startColumnOffset))
            {
                throw new Exception($"Cannot get top left column offset of given shape. Found \"{fromMarker?.ColumnOffset?.Text}\"");
            }
            
            var toMarker = twoCellAnchor.ToMarker;
            if (toMarker == null)
            {
                toMarker = new xdr.ToMarker().At("C4");
                twoCellAnchor.ToMarker = toMarker;
            }
            if (!int.TryParse(toMarker?.ColumnId?.Text, out var endColumnId))
            {
                throw new Exception($"Cannot get top left columm number of given shape. Found \"{fromMarker?.ColumnId?.Text}\"");
            }
            if (!int.TryParse(toMarker?.ColumnOffset?.Text, out var endColumnOffset))
            {
                throw new Exception($"Cannot get top left column offset of given shape. Found \"{fromMarker?.ColumnOffset?.Text}\"");
            }

            var startColumnNum = startColumnId + 1;
            var endColumnNum = endColumnId + 1;
            var wdr = twoCellAnchor.ParentOfType<xdr.WorksheetDrawing>();
            var wsPart = wdr.DrawingsPart.ParentPartOfType<WorksheetPart>();
            var ws = wsPart.Worksheet;

            var columnNum = startColumnNum;
            var totalWidth = 0d;
            while (columnNum <= endColumnNum)
            {
                var column = ws.GetColumn(columnNum);
                var columnWidth = column.GetWidthInPixels(dpi);
                totalWidth += columnWidth;
                columnNum++;
            }
            var startColumnOffsetInPixels = Utils.ConvertEmuToPixels(startColumnOffset, dpi);
            var endColumnOffsetInPixels = Utils.ConvertEmuToPixels(endColumnOffset, dpi);
            totalWidth += endColumnOffsetInPixels - startColumnOffsetInPixels;
            if(totalWidth < 0) { totalWidth = 0; }
            return totalWidth;
        }
    }
}
