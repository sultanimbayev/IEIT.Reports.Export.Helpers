using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ShapeSetBottomRight
    {
        public static xdr.Shape SetBottomRight(this xdr.Shape shape, string cellAddress, double topOffset = 0, double leftOffset = 0)
        {
            var twoCellAnchor = shape.Parent as xdr.TwoCellAnchor;
            if (twoCellAnchor != null)
            {
                twoCellAnchor.ToMarker = new xdr.ToMarker().At(cellAddress, topOffset, leftOffset);
                return shape;
            }
            return shape;
        }

        public static xdr.Shape SetBottomRight(this xdr.Shape shape, int rowNum, int columnNum)
        {
            var twoCellAnchor = shape.Parent as xdr.TwoCellAnchor;
            if (twoCellAnchor != null)
            {
                twoCellAnchor.ToMarker = new xdr.ToMarker().At(rowNum, columnNum);
                return shape;
            }
            return shape;
        }

        public static xdr.Shape SetBottomRight(this xdr.Shape shape, int rowNum, double topOffset, int columnNum, double leftOffset)
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
