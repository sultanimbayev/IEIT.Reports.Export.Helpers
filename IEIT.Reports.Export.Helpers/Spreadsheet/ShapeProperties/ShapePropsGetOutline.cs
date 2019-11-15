using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ShapePropsGetOutline
    {
        public static a.Outline GetOutline(this xdr.ShapeProperties shapeProperties)
        {
            var outline = shapeProperties.GetFirstChild<a.Outline>();
            if(outline == null)
            {
                outline = new a.Outline();
                shapeProperties.SetOutline(outline);
            }
            return outline;
        }
    }
}
