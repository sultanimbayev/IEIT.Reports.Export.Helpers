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
        public static xdr.ShapeProperties InitDefault(this xdr.ShapeProperties shapeProperties)
        {
            if(shapeProperties == null) { return null; }

            var presetGeometry = new a.PresetGeometry() { Preset = a.ShapeTypeValues.Rectangle }; //ShapeType - def
            shapeProperties.Append(presetGeometry);
            presetGeometry.Append(new a.AdjustValueList());

            shapeProperties.SetSolidFill(sysDr.Color.Red, 0.27f);

            //var fill = new a.SolidFill(); // Fill - def on FillColor
            //shapeProperties.Append(fill);

            //var fillColor = new a.RgbColorModelHex();
            //fillColor.Val = sysDr.Color.Red.ToHex(); // FillColor - def
            //fill.Append(fillColor);

            //var colorAlpha = new a.Alpha() { Val = 27 * 1000 }; // FillAlpha - def = val / 1000
            //fillColor.Append(colorAlpha);

            var outline = new a.Outline().InitDefault(); // Outline - def
            shapeProperties.Append(outline);
            //outline.Width = (int)(1.5 * 12700); // OutlineWidth- def
            //var outlineFill = new a.SolidFill();
            //outline.Append(outlineFill);

            //var outlineFillColor = sysDr.Color.Red; // OutlineFillColor - def
            //outlineFill.Append(new a.RgbColorModelHex() { Val = outlineFillColor.ToHex() });

            return shapeProperties;
        }
    }
}
