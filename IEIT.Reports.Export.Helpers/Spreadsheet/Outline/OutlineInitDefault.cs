using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;
using sysDr = System.Drawing;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class OutlineInitDefault
    {
        public static a.Outline InitDefault(this a.Outline outline)
        {
            outline.Width = (int)(12700); // OutlineWidth- def
            var outlineFill = new a.SolidFill();
            outline.Append(outlineFill);

            var outlineFillColor = sysDr.Color.Black; // OutlineFillColor - def
            outlineFill.Append(new a.RgbColorModelHex() { Val = outlineFillColor.ToHex() });

            return outline;
        }
    }
}
