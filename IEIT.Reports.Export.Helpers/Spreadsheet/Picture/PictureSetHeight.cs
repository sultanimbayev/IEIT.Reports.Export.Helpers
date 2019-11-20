using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class PictureSetHeight
    {
        public static xdr.Picture SetHeightInPixels(this xdr.Picture picture, double heightInPixels, double dpi = 96)
        {
            if (picture.Parent == null) { return picture; }
            var twoCellAnchor = picture.Parent as xdr.TwoCellAnchor;
            if (twoCellAnchor == null)
            {
                var parentType = picture.Parent.GetType();
                throw new Exception($"Parent element of shape must be TwoCellAnchor. Found \"{parentType.Name}\"");
            }
            twoCellAnchor.SetHeightInPixels(heightInPixels, dpi);
            return picture;
        }
    }
}
