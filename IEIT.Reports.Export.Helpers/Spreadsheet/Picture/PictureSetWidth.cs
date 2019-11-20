using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class PictureSetWidth
    {
        public static xdr.Picture SetWidthInPixels(this xdr.Picture picture, double widthInPixels, double dpi = 96)
        {
            if (picture.Parent == null) { return picture; }
            var twoCellAnchor = picture.Parent as xdr.TwoCellAnchor;
            if (twoCellAnchor == null)
            {
                var parentType = picture.Parent.GetType();
                throw new Exception($"Parent element of shape must be TwoCellAnchor. Found \"{parentType.Name}\"");
            }
            twoCellAnchor.SetWidthInPixels(widthInPixels, dpi);
            return picture;
        }
    }
}
