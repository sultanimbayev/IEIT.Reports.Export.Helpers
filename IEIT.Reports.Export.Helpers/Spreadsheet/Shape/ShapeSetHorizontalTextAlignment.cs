using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using a = DocumentFormat.OpenXml.Drawing;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ShapeSetHorizontalTextAlignment
    {
        public static xdr.Shape SetHorizontalTextAlignment(this xdr.Shape shape, a.TextAlignmentTypeValues textAlignment)
        {
            if(shape == null) { return null; }
            if(shape.TextBody == null)
            {
                shape.TextBody = new xdr.TextBody().InitDefault();
            }
            shape.TextBody.SetHorizontalAlignment(textAlignment);
            return shape;
        }
    }
}
