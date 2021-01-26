using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class StylesheetDifferentialFormat
    {
        public static DifferentialFormat DifferentialFormat(this Stylesheet stylesheet, int index)
        {
            var dformats = stylesheet.DifferentialFormats ?? (stylesheet.DifferentialFormats = new DifferentialFormats()
            {
                Count = 0
            });
            return dformats.Elements<DifferentialFormat>().ElementAt(index);
        }

        public static DifferentialFormat DifferentialFormat(this Stylesheet stylesheet, uint index)
        {
            return DifferentialFormat(stylesheet, (int)index);
        }

        public static uint DifferentialFormat(this Stylesheet stylesheet, DifferentialFormat dformat)
        {
            var dformats = stylesheet.DifferentialFormats ?? (stylesheet.DifferentialFormats = new DifferentialFormats()
            {
                Count = 0
            });
            var dformatIndex = dformats.MakeSame(dformat);
            dformats.Count = (uint)(dformats.Elements().Count());
            return dformatIndex;
        }
    }
}
