using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class TextBodySetText
    {
        public static a.TextBody SetText(this a.TextBody textBody, string text, Font font = null, Color? fontColor = null)
        {
            textBody.RemoveAllChildren<a.Paragraph>();
            var p = new a.Paragraph();
            textBody.Append(p);
            p.AddText(text, font, fontColor);
            return textBody;
        }
    }
}
