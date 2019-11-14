using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;
using sysDr = System.Drawing;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class OutlineSetSolidFill
    {
        public static a.Outline SetSolidFill(this a.Outline outline, sysDr.Color? fillColor = null)
        {
            var solidFill = outline.GetFirstChild<a.SolidFill>();
            if (!fillColor.HasValue && solidFill != null) { return outline; }
            if(solidFill == null)
            {
                solidFill = new a.SolidFill();
                outline.Append(solidFill);
            }
            solidFill.RemoveAllChildren();
            var _c = fillColor ?? sysDr.Color.Black;
            solidFill.Append(new a.RgbColorModelHex() { Val = _c.ToHex() });
            return outline;
        }
    }
}
