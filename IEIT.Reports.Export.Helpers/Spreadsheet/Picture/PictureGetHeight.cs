using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;


namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class PictureGetHeight
    {
        public static double GetHeightInPixels(this xdr.Shape shape, double dpi = 96d)
        {
            if (shape == null) { return -1d; }
            if (shape.Parent == null) { return -1d; }
            var twoCellAnchor = shape.Parent as xdr.TwoCellAnchor;
            if (twoCellAnchor == null)
            {
                throw new Exception($"Works only with TwoCellAnchor type. Found \"{shape.Parent.GetType().Name}\"");
            }
            return twoCellAnchor.GetHeightInPixels(dpi);
        }
    }
}
