using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using sysDr = System.Drawing;
using a = DocumentFormat.OpenXml.Drawing;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ShapePropsInit
    {
        public static xdr.ShapeProperties Init(this xdr.ShapeProperties shapeProperties, a.ShapeTypeValues shapeType)
        {

            if (shapeProperties == null) { return null; }

            var presetGeometry = new a.PresetGeometry() { Preset = shapeType }; //ShapeType - def
            shapeProperties.Append(presetGeometry);
            presetGeometry.Append(new a.AdjustValueList());

            shapeProperties.SetSolidFill(sysDr.Color.White);
            var outline = new a.Outline().InitDefault(); // Outline - def
            shapeProperties.Append(outline);
            return shapeProperties;
        }
    }
}
