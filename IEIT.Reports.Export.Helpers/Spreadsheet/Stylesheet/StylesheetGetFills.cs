using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class StylesheetGetFills
    {
        internal static Fills GetFills(this Stylesheet stylesheet)
        {
            if (stylesheet.Fills == null)
            {
                stylesheet.Fills = new Fills() { Count = 2 };
                stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
                stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
            }
            return stylesheet.Fills;
        }
    }
}
