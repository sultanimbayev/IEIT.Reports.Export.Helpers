using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet.Shape
{
    public static class ShapeGetOutline
    {
        public static a.Outline GetOutline(this xdr.Shape shape)
        {
            if (shape == null) { return null; }
            if (shape.ShapeProperties == null)
            {
                shape.ShapeProperties = new xdr.ShapeProperties().Init(a.ShapeTypeValues.Rectangle);
            }
            return shape.ShapeProperties.GetOutline();
        }
    }
}
