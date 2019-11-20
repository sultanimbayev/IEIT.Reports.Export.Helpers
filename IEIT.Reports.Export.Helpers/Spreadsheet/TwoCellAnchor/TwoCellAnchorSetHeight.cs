using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class TwoCellAnchorSetHeight
    {
        public static xdr.TwoCellAnchor SetHeightInPixels(this xdr.TwoCellAnchor twoCellAnchor, double heightInPixels, double dpi = 96)
        {
            if (twoCellAnchor == null)
            {
                return null;
            }
            var fromMareker = twoCellAnchor.FromMarker;
            if (!int.TryParse(fromMareker?.RowId?.Text, out var startRowId))
            {
                throw new Exception($"Cannot get top left columm number of given shape. Found \"{fromMareker?.ColumnId?.Text}\"");
            }
            if (!int.TryParse(fromMareker?.RowOffset?.Text, out var startRowOffset))
            {
                throw new Exception($"Cannot get top left column offset of given shape. Found \"{fromMareker?.ColumnOffset?.Text}\"");
            }
            var startRowNum = startRowId + 1;
            var _newNormalizedHeight = heightInPixels;
            var wdr = twoCellAnchor.ParentOfType<xdr.WorksheetDrawing>();
            var wsPart = wdr.DrawingsPart.ParentPartOfType<WorksheetPart>();
            var ws = wsPart.Worksheet;

            var endRowNum = startRowNum;
            var endRow = ws.GetRow(endRowNum);
            var startRowOffsetInPixels = Utils.ConvertEmuToPixels(startRowOffset, dpi);
            var offsetInPixels = _newNormalizedHeight - startRowOffsetInPixels;
            var rowHeight = endRow.GetHeightInPixels(dpi) - startRowOffsetInPixels;
            while (rowHeight <= offsetInPixels)
            {
                offsetInPixels -= rowHeight;
                endRowNum++;
                endRow = ws.GetRow(endRowNum);
                rowHeight = endRow.GetHeightInPixels(dpi);
            }
            var toMarker = twoCellAnchor.ToMarker;
            if (toMarker == null)
            {
                toMarker = new xdr.ToMarker();
                twoCellAnchor.ToMarker = toMarker;
            }
            toMarker.SetRowNum(endRowNum);
            if (offsetInPixels < 0) { offsetInPixels = 0; }
            toMarker.SetTopOffset(Utils.ConvertPixelsToEmu(offsetInPixels, dpi));
            return twoCellAnchor;
        }
    }
}
