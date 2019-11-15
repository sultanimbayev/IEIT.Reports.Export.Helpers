using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ShapeSetOutlineWidth
    {
        public static xdr.Shape SetOutlineWidthInPixels(this xdr.Shape shape, float width)
        {
            if (shape == null) { return null; }
            if (shape.ShapeProperties == null)
            {
                shape.ShapeProperties = new xdr.ShapeProperties().Init(a.ShapeTypeValues.Rectangle);
            }
            var outline = shape.ShapeProperties.GetOutline();
            outline.SetWidthInPixels(width);
            return shape;
        }
    }
}
