using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class TwoCellAnchorSetWidth
    {
        public static xdr.TwoCellAnchor SetWidthInPixels(this xdr.TwoCellAnchor twoCellAnchor, double widthInPixels, double dpi = 96)
        {
            if (twoCellAnchor == null)
            {
                return null;
            }
            var fromMareker = twoCellAnchor.FromMarker;
            if (!int.TryParse(fromMareker?.ColumnId?.Text, out var startColumnId))
            {
                throw new Exception($"Cannot get top left columm number of given shape. Found \"{fromMareker?.ColumnId?.Text}\"");
            }
            if (!int.TryParse(fromMareker?.ColumnOffset?.Text, out var startColumnOffset))
            {
                throw new Exception($"Cannot get top left column offset of given shape. Found \"{fromMareker?.ColumnOffset?.Text}\"");
            }
            var startColumnNum = startColumnId + 1;
            var _newNormalizedWidth = widthInPixels;
            var wdr = twoCellAnchor.ParentOfType<xdr.WorksheetDrawing>();
            var wsPart = wdr.DrawingsPart.ParentPartOfType<WorksheetPart>();
            var ws = wsPart.Worksheet;

            var endColumnNum = startColumnNum;
            var endColumn = ws.GetColumn(endColumnNum);
            var startColumnOffsetInPixels = Utils.ConvertEmuToPixels(startColumnOffset, dpi);
            var offsetInPixels = _newNormalizedWidth - startColumnOffsetInPixels;
            var columnWidth = endColumn.GetWidthInPixels(dpi) - startColumnOffsetInPixels;
            while (columnWidth < offsetInPixels)
            {
                offsetInPixels -= columnWidth;
                endColumnNum++;
                endColumn = ws.GetColumn(endColumnNum);
                columnWidth = endColumn.GetWidthInPixels(dpi);
            }
            var toMarker = twoCellAnchor.ToMarker;
            if (toMarker == null)
            {
                toMarker = new xdr.ToMarker();
                twoCellAnchor.ToMarker = toMarker;
            }
            toMarker.SetColumnNum(endColumnNum);
            if(offsetInPixels < 0) { offsetInPixels = 0; }
            toMarker.SetLeftOffset(Utils.ConvertPixelsToEmu(offsetInPixels, dpi));
            return twoCellAnchor;
        }
    }
}
