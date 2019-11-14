using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using a = DocumentFormat.OpenXml.Drawing;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ShapeStyleInitDefault
    {
        public static xdr.ShapeStyle InitDefault(this xdr.ShapeStyle shapeStyle)
        {
            var lineReference1 = new a.LineReference() { Index = 0U };
            shapeStyle.Append(lineReference1);
            
            var rgbColorModelPercentage1 = new a.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 };
            lineReference1.Append(rgbColorModelPercentage1);

            var fillReference1 = new a.FillReference() { Index = 0U };
            shapeStyle.Append(fillReference1);

            var rgbColorModelPercentage2 = new a.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 };
            fillReference1.Append(rgbColorModelPercentage2);

            var effectReference1 = new a.EffectReference() { Index = 0U };
            shapeStyle.Append(effectReference1);

            var rgbColorModelPercentage3 = new a.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 };
            effectReference1.Append(rgbColorModelPercentage3);

            var fontReference1 = new a.FontReference() { Index = a.FontCollectionIndexValues.Major };
            shapeStyle.Append(fontReference1);

            return shapeStyle;
        }
    }
}
