using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class TwoCellAnchorInitDefault
    {
        public static xdr.TwoCellAnchor InitDefault (this xdr.TwoCellAnchor twoCellAnchor)
        {
            twoCellAnchor.FromMarker = new xdr.FromMarker();
            twoCellAnchor.FromMarker.At("A1");
            twoCellAnchor.ToMarker = new xdr.ToMarker();
            twoCellAnchor.ToMarker.At("F6");
            return twoCellAnchor;
        }
    }
}
