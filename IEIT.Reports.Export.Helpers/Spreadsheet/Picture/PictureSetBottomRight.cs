using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class PictureSetBottomRight
    {
        public static xdr.Picture SetBottomRight(this xdr.Picture shape, string cellAddress, double topOffset = 0, double leftOffset = 0)
        {
            var twoCellAnchor = shape.Parent as xdr.TwoCellAnchor;
            if (twoCellAnchor != null)
            {
                twoCellAnchor.ToMarker = new xdr.ToMarker().At(cellAddress, topOffset, leftOffset);
                return shape;
            }
            return shape;
        }

        public static xdr.Picture SetBottomRight(this xdr.Picture shape, int rowNum, int columnNum)
        {
            var twoCellAnchor = shape.Parent as xdr.TwoCellAnchor;
            if (twoCellAnchor != null)
            {
                twoCellAnchor.ToMarker = new xdr.ToMarker().At(rowNum, columnNum);
                return shape;
            }
            return shape;
        }

        public static xdr.Picture SetBottomRight(this xdr.Picture shape, int rowNum, double topOffset, int columnNum, double leftOffset)
        {
            var twoCellAnchor = shape.Parent as xdr.TwoCellAnchor;
            if (twoCellAnchor != null)
            {
                twoCellAnchor.ToMarker = new xdr.ToMarker().At(rowNum, topOffset, columnNum, leftOffset);
                return shape;
            }
            return shape;
        }

    }
}
