using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ShapePropsSetOutline
    {
        public static xdr.ShapeProperties SetOutline(this xdr.ShapeProperties shapeProperties, a.Outline outline)
        {
            if(shapeProperties == null) { return null; }
            shapeProperties.RemoveAllChildren<a.Outline>();
            if(outline != null)
            {
                shapeProperties.Append(outline);
            }
            return shapeProperties;
        }
    }
}
