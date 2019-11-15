using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;
using sysDr = System.Drawing;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class OutlineSetWidth
    {
        public static a.Outline SetWidthInPixels(this a.Outline outline, float width)
        {
            outline.Width = (int)(1.5 * 12700); // OutlineWidth- def

            outline.RemoveAllChildren<a.NoFill>();
            var solidFill = outline.GetFirstChild<a.SolidFill>();
            if(solidFill != null) { return outline; }
            var gradientFill = outline.GetFirstChild<a.GradientFill>();
            if(gradientFill != null) { return outline; }
            var patternFill = outline.GetFirstChild<a.PatternFill>();
            if(patternFill != null) { return outline; }

            solidFill = new a.SolidFill();
            outline.Append(solidFill);

            var fillColor = sysDr.Color.Black; // OutlineFillColor - def
            solidFill.Append(new a.RgbColorModelHex() { Val = fillColor.ToHex() });
            return outline;
        }
    }
}
