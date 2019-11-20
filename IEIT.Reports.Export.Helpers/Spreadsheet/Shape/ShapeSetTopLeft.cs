using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ShapeSetTopLeft
    {

        public static xdr.Shape SetTopLeft(this xdr.Shape shape, string cellAddress, double topOffset = 0, double leftOffset = 0)
        {
            var twoCellAnchor = shape.Parent as xdr.TwoCellAnchor;
            if (twoCellAnchor != null)
            {
                twoCellAnchor.SetTopLeft(cellAddress, topOffset, leftOffset);
                return shape;
            }

            var oneCellAnchor = shape.Parent as xdr.OneCellAnchor;
            if (oneCellAnchor != null)
            {
                oneCellAnchor.FromMarker = new xdr.FromMarker().At(cellAddress, topOffset, leftOffset);
                return shape;
            }

            return shape;
        }

        public static xdr.Shape SetTopLeft(this xdr.Shape shape, int rowNum, int columnNum)
        {
            var twoCellAnchor = shape.Parent as xdr.TwoCellAnchor;
            if (twoCellAnchor != null)
            {
                twoCellAnchor.SetTopLeft(rowNum, columnNum);
                return shape;
            }

            var oneCellAnchor = shape.Parent as xdr.OneCellAnchor;
            if (oneCellAnchor != null)
            {
                oneCellAnchor.FromMarker = new xdr.FromMarker().At(rowNum, columnNum);
                return shape;
            }

            return shape;
        }


        public static xdr.Shape SetTopLeft(this xdr.Shape shape, int rowNum, double topOffset, int columnNum, double leftOffset)
        {
            var twoCellAnchor = shape.Parent as xdr.TwoCellAnchor;
            if(twoCellAnchor != null)
            {
                twoCellAnchor.SetTopLeft(rowNum, topOffset, columnNum, leftOffset);
                return shape;
            }

            var oneCellAnchor = shape.Parent as xdr.OneCellAnchor;
            if (oneCellAnchor != null)
            {
                oneCellAnchor.FromMarker = new xdr.FromMarker().At(rowNum, topOffset, columnNum, leftOffset);
                return shape;
            }

            return shape;
        }
    }
}
