using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ShapeSetText
    {
        public static xdr.Shape SetText(this xdr.Shape shape, string text = "", Font font = null, Color? fontColor = null)
        {
            if(shape == null) { return null; }
            if(shape.TextBody == null)
            {
                shape.TextBody = new xdr.TextBody().InitDefault();
            }
            shape.TextBody.SetText(text, font, fontColor);
            return shape;
        }
    }
}
