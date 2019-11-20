using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using a = DocumentFormat.OpenXml.Drawing;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ShapeSetLineSpace
    {
        public static xdr.Shape SetLineSpace(this xdr.Shape shape, float heightMultiplier)
        {
            if (shape == null) { return null; }
            if (shape.TextBody == null)
            {
                shape.TextBody = new xdr.TextBody().InitDefault();
            }
            shape.TextBody.SetLineSpace(heightMultiplier);
            return shape;
        }
    }
}
