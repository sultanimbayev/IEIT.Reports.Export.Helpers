using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ShapeSetFont
    {
        public static xdr.Shape SetFont(this xdr.Shape shape, Font font, Color? fontColor = null)
        {
            if(shape == null) { return null; }
            if(shape.TextBody == null)
            {
                shape.TextBody = new xdr.TextBody().InitDefault();
            }
            shape.TextBody.SetFont(font, fontColor);
            return shape;
        }
    }
}
