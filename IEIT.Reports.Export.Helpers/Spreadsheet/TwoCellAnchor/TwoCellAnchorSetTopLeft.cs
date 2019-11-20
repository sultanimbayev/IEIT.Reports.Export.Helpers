using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class TwoCellAnchorSetTopLeft
    {
        private const double __DPI = 96;
        public static xdr.TwoCellAnchor SetTopLeft(this xdr.TwoCellAnchor twoCellAnchor, string cellAddress, double topOffset = 0, double leftOffset = 0)
        {
            var initialWidth = twoCellAnchor.GetWidthInPixels(__DPI);
            var initalHeight = twoCellAnchor.GetHeightInPixels(__DPI);
            twoCellAnchor.FromMarker = new xdr.FromMarker().At(cellAddress, topOffset, leftOffset);
            twoCellAnchor.SetWidthInPixels(initialWidth, __DPI);
            twoCellAnchor.SetHeightInPixels(initalHeight, __DPI);
            return twoCellAnchor;
        }

        public static xdr.TwoCellAnchor SetTopLeft(this xdr.TwoCellAnchor twoCellAnchor, int rowNum, int columnNum)
        {
            var initialWidth = twoCellAnchor.GetWidthInPixels(__DPI);
            var initalHeight = twoCellAnchor.GetHeightInPixels(__DPI);
            twoCellAnchor.FromMarker = new xdr.FromMarker().At(rowNum, columnNum);
            twoCellAnchor.SetWidthInPixels(initialWidth, __DPI);
            twoCellAnchor.SetHeightInPixels(initalHeight, __DPI);
            return twoCellAnchor;
        }


        public static xdr.TwoCellAnchor SetTopLeft(this xdr.TwoCellAnchor twoCellAnchor, int rowNum, double topOffset, int columnNum, double leftOffset)
        {
            var initialWidth = twoCellAnchor.GetWidthInPixels(__DPI);
            var initalHeight = twoCellAnchor.GetHeightInPixels(__DPI);
            twoCellAnchor.FromMarker = new xdr.FromMarker().At(rowNum, topOffset, columnNum, leftOffset);
            twoCellAnchor.SetWidthInPixels(initialWidth, __DPI);
            twoCellAnchor.SetHeightInPixels(initalHeight, __DPI);
            return twoCellAnchor;
        }
    }
}
