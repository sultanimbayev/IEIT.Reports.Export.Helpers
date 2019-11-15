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
    public static class ShapeSetSolidFill
    {
        public static xdr.Shape SetSolidFill(this xdr.Shape shape, sysDr.Color fillColor, float alpha = 1f)
        {
            if (shape == null) { return null; }
            if (shape.ShapeProperties == null)
            {
                shape.ShapeProperties = new xdr.ShapeProperties().Init(a.ShapeTypeValues.Rectangle);
            }
            shape.ShapeProperties.SetSolidFill(fillColor, alpha);
            return shape;
        }
    }
}
