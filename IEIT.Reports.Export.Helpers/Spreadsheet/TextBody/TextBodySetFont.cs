using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

using a = DocumentFormat.OpenXml.Drawing;
using System.Drawing;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class TextBodySetFont
    {
        public static xdr.TextBody SetFont(this xdr.TextBody textBody, Font font, Color? color = null)
        {
            var paragraphs = textBody.Elements<a.Paragraph>();
            foreach(var p in paragraphs)
            {
                p.SetFont(font, color);
            }
            return textBody;
        }
    }
}
