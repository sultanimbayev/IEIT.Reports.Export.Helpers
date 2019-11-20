using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class TwoCellAnchorGetHeight
    {
        public static double GetHeightInPixels(this xdr.TwoCellAnchor twoCellAnchor, double dpi = 96)
        {
            if (twoCellAnchor == null)
            {
                throw new ArgumentNullException("Cannot get anchor width, null given.");
            }
            var fromMarker = twoCellAnchor.FromMarker;
            if (fromMarker == null)
            {
                fromMarker = new xdr.FromMarker().At("A1");
                twoCellAnchor.FromMarker = fromMarker;
            }
            if (!int.TryParse(fromMarker?.RowId?.Text, out var startRowId))
            {
                throw new Exception($"Cannot get top left columm number of given shape. Found \"{fromMarker?.ColumnId?.Text}\"");
            }
            if (!int.TryParse(fromMarker?.RowOffset?.Text, out var startRowOffset))
            {
                throw new Exception($"Cannot get top left column offset of given shape. Found \"{fromMarker?.ColumnOffset?.Text}\"");
            }

            var toMarker = twoCellAnchor.ToMarker;
            if (toMarker == null)
            {
                toMarker = new xdr.ToMarker().At("C4");
                twoCellAnchor.ToMarker = toMarker;
            }
            if (!int.TryParse(toMarker?.RowId?.Text, out var endRowId))
            {
                throw new Exception($"Cannot get top left columm number of given shape. Found \"{fromMarker?.ColumnId?.Text}\"");
            }
            if (!int.TryParse(toMarker?.RowOffset?.Text, out var endRowOffset))
            {
                throw new Exception($"Cannot get top left column offset of given shape. Found \"{fromMarker?.ColumnOffset?.Text}\"");
            }

            var startRowNum = startRowId + 1;
            var endRowNum = endRowId + 1;
            var wdr = twoCellAnchor.ParentOfType<xdr.WorksheetDrawing>();
            var wsPart = wdr.DrawingsPart.ParentPartOfType<WorksheetPart>();
            var ws = wsPart.Worksheet;

            var rowNum = startRowNum;
            var totalHeight = 0d;
            while (rowNum <= endRowNum)
            {
                var row = ws.GetRow(rowNum);
                var rowHeight = row.GetHeightInPixels(dpi);
                totalHeight += rowHeight;
                rowNum++;
            }
            var startRowOffsetInPixels = Utils.ConvertEmuToPixels(startRowOffset, dpi);
            var endRowffsetInPixels = Utils.ConvertEmuToPixels(endRowOffset, dpi);
            totalHeight += endRowffsetInPixels - startRowOffsetInPixels;
            if(totalHeight < 0) { totalHeight = 0; }
            return totalHeight;
        }
    }
}
