using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ShapeSetOutlineColor
    {
        public static xdr.Shape SetOutlineColor(this xdr.Shape shape, Color color, float alpha = 1f)
        {
            if (shape == null) { return null; }
            if (shape.ShapeProperties == null)
            {
                shape.ShapeProperties = new xdr.ShapeProperties().Init(a.ShapeTypeValues.Rectangle);
            }
            var outline = shape.ShapeProperties.GetOutline();
            outline.SetSolidFill(color, 1);
            return shape;
        }
    }
}
