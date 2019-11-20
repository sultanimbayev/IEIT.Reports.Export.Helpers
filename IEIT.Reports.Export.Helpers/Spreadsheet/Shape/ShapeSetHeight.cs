using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ShapeSetHeight
    {
        public static xdr.Shape SetHeightInPixels(this xdr.Shape shape, double heightInPixels, double dpi = 96)
        {
            if (shape.Parent == null) { return shape; }
            var twoCellAnchor = shape.Parent as xdr.TwoCellAnchor;
            if (twoCellAnchor == null)
            {
                var parentType = shape.Parent.GetType();
                throw new Exception($"Parent element of shape must be TwoCellAnchor. Found \"{parentType.Name}\"");
            }
            twoCellAnchor.SetHeightInPixels(heightInPixels, dpi);
            return shape;
        }
    }
}
