using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class TextBodySetHorizontalAlignment
    {
        public static xdr.TextBody SetHorizontalAlignment(this xdr.TextBody textBody, a.TextAlignmentTypeValues textAlignment)
        {
            var paragraphs = textBody.Elements<a.Paragraph>();
            foreach(var p in paragraphs)
            {
                p.SetTextAlignment(textAlignment);
            }
            return textBody;
        }
    }
}
