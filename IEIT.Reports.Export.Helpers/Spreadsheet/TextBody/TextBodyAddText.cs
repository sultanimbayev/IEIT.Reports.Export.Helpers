using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;

namespace IEIT.Reports.Export.Helpers.Spreadsheet.TextBody
{
    public static class TextBodyAddText
    {
        public static a.TextBody AddText(this a.TextBody textBody, string text, Font font = null, Color? fontColor = null)
        {
            var p = textBody.Elements<a.Paragraph>().LastOrDefault();
            if(p == null)
            {
                p = new a.Paragraph();
                textBody.Append(p);
            }
            p.AddText(text, font, fontColor);
            return textBody;
        }
    }
}
