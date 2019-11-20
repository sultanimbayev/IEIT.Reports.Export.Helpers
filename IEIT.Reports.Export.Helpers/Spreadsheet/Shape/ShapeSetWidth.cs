using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ShapeSetWidth
    {
        public static xdr.Shape SetWidthInPixels(xdr.Shape shape, double newWidthInPixels, double dpi = 96)
        {
            if(shape.Parent == null) { return shape; }
            var twoCellAnchor = shape.Parent as xdr.TwoCellAnchor;
            if(twoCellAnchor == null) {
                var parentType = shape.Parent.GetType();
                throw new Exception($"Parent element of shape must be TwoCellAnchor. Found \"{parentType.Name}\"");
            }
            twoCellAnchor.SetWidthInPixels(newWidthInPixels, dpi);
            return shape;
        }
    }
}
