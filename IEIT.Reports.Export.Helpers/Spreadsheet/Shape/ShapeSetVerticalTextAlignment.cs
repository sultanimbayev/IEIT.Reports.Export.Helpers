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
    public static class ShapeSetVerticalTextAlignment
    {
        public static xdr.Shape SetVerticalTextAlignment(this xdr.Shape shape, a.TextAnchoringTypeValues alignment)
        {
            if(shape == null) { return null; }
            if(shape.TextBody == null)
            {
                shape.TextBody = new xdr.TextBody().InitDefault();
            }
            shape.TextBody.SetVerticalAlignment(alignment);
            return shape;
        }
    }
}
