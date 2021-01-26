using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class StylesheetBorder
    {
        internal static Borders GetBordersOf(Stylesheet stylesheet)
        {
            if (stylesheet.Borders == null) { stylesheet.Borders = new Borders(new Border()) { Count = 1 }; } // blank border list, if not exists
            return stylesheet.Borders;
        }

        public static Border Border(this Stylesheet stylesheet, int index)
        {
            return GetBordersOf(stylesheet).Elements<Border>().ElementAt(index);
        }
        public static Border Border(this Stylesheet stylesheet, uint index)
        {
            return GetBordersOf(stylesheet).Elements<Border>().ElementAt((int)index);
        }

        public static uint Border(this Stylesheet stylesheet, Border border)
        {
            var bordersList = GetBordersOf(stylesheet);
            var newBorderIndex = bordersList.MakeSame(border);
            bordersList.Count = (uint)bordersList.Elements().Count();
            return newBorderIndex;
        }
    }
}
