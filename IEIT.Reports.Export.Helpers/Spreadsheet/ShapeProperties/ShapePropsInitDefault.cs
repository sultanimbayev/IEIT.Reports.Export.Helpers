using System;
using System.Collections.Generic;
using sysDr = System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ShapePropsGetDefault
    {
        public static xdr.ShapeProperties InitPictureDefault(this xdr.ShapeProperties shapeProperties)
        {
            if(shapeProperties == null) { return null; }
            var presetGeometry = new a.PresetGeometry() { Preset = a.ShapeTypeValues.Rectangle };
            shapeProperties.Append(presetGeometry);
            presetGeometry.Append(new a.AdjustValueList());
            return shapeProperties;
        }
    }
}
