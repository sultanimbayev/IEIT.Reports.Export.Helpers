using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using sysDr = System.Drawing;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ShapeRemoveFill
    {
        public static xdr.Shape RemoveFill(this xdr.Shape shape)
        {
            if (shape == null) { return null; }
            if (shape.ShapeProperties == null)
            {
                shape.ShapeProperties = new xdr.ShapeProperties().Init(a.ShapeTypeValues.Rectangle);
            }
            shape.ShapeProperties.RemoveFill();
            return shape;
        }
    }
}
